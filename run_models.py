#!/usr/bin/env python3
"""オーストラリア天気予測 - 高速実行版"""

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import json
import warnings
from pathlib import Path

import torch
import torch.nn as nn
import torch.optim as optim
from torch.utils.data import DataLoader, TensorDataset

from sklearn.preprocessing import StandardScaler, LabelEncoder
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import (
    accuracy_score, precision_score, recall_score, f1_score,
    roc_auc_score, average_precision_score, confusion_matrix,
    roc_curve, precision_recall_curve
)

import optuna
from optuna.pruners import MedianPruner

warnings.filterwarnings('ignore')
optuna.logging.set_verbosity(optuna.logging.WARNING)

SEED = 42
np.random.seed(SEED)
torch.manual_seed(SEED)

device = torch.device('cuda' if torch.cuda.is_available() else 'cpu')
print(f'Using device: {device}')

DATA_PATH = Path('data/weatherAUS.csv')
FIGURES_PATH = Path('figures')
EXPERIMENTS_PATH = Path('experiments')

# ========== データ読み込み ==========
print('\n=== データ読み込み ===')
df = pd.read_csv(DATA_PATH)
print(f'データサイズ: {df.shape}')

# ========== EDA可視化 ==========
print('\n=== EDA可視化 ===')

# 欠損値ヒートマップ
fig, ax = plt.subplots(figsize=(12, 8))
missing_matrix = df.isnull().astype(int)
sns.heatmap(missing_matrix.T, cbar=True, yticklabels=df.columns, cmap='YlOrRd', ax=ax)
ax.set_title('Missing Values Heatmap', fontsize=14)
plt.tight_layout()
plt.savefig(FIGURES_PATH / 'missing_values.png', dpi=150, bbox_inches='tight')
plt.close()

# ターゲット分布
fig, axes = plt.subplots(1, 2, figsize=(12, 5))
target_counts = df['RainTomorrow'].value_counts()
colors = ['#3498db', '#e74c3c']
axes[0].bar(target_counts.index, target_counts.values, color=colors)
axes[0].set_title('RainTomorrow Distribution (Count)')
for i, v in enumerate(target_counts.values):
    axes[0].text(i, v + 1000, f'{v:,}', ha='center')
target_pct = df['RainTomorrow'].value_counts(normalize=True) * 100
axes[1].pie(target_pct.values, labels=target_pct.index, autopct='%1.1f%%', colors=colors)
axes[1].set_title('RainTomorrow Distribution (%)')
plt.tight_layout()
plt.savefig(FIGURES_PATH / 'target_distribution.png', dpi=150, bbox_inches='tight')
plt.close()

# 相関行列
numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
corr_matrix = df[numeric_cols].corr()
fig, ax = plt.subplots(figsize=(14, 12))
mask = np.triu(np.ones_like(corr_matrix, dtype=bool))
sns.heatmap(corr_matrix, mask=mask, annot=True, fmt='.2f', cmap='RdBu_r', center=0, ax=ax, annot_kws={'size': 8})
ax.set_title('Correlation Matrix')
plt.tight_layout()
plt.savefig(FIGURES_PATH / 'correlation_matrix.png', dpi=150, bbox_inches='tight')
plt.close()

print('EDA可視化完了')

# ========== 前処理 ==========
print('\n=== 前処理 ===')
df_clean = df.dropna(subset=['RainTomorrow']).copy()
df_clean['Date'] = pd.to_datetime(df_clean['Date'])
df_clean = df_clean.sort_values('Date').reset_index(drop=True)

train_end = '2015-06-30'
val_end = '2016-06-30'
train_mask = df_clean['Date'] <= train_end
val_mask = (df_clean['Date'] > train_end) & (df_clean['Date'] <= val_end)
test_mask = df_clean['Date'] > val_end

print(f'Train: {train_mask.sum():,} | Val: {val_mask.sum():,} | Test: {test_mask.sum():,}')

cat_cols = ['Location', 'WindGustDir', 'WindDir9am', 'WindDir3pm', 'RainToday']
num_cols = [c for c in df_clean.columns if c not in cat_cols + ['Date', 'RainTomorrow']]

# 欠損値補完
df_imputed = df_clean.copy()
train_data = df_imputed[train_mask]
for col in num_cols:
    median_val = train_data[col].median()
    df_imputed[col] = df_imputed[col].fillna(median_val)
for col in cat_cols:
    mode_val = train_data[col].mode().iloc[0] if len(train_data[col].mode()) > 0 else 'Unknown'
    df_imputed[col] = df_imputed[col].fillna(mode_val)

# 風向エンコーディング
wind_directions = ['N', 'NNE', 'NE', 'ENE', 'E', 'ESE', 'SE', 'SSE', 'S', 'SSW', 'SW', 'WSW', 'W', 'WNW', 'NW', 'NNW']
wind_to_angle = {d: i * (360 / 16) for i, d in enumerate(wind_directions)}
for col in ['WindGustDir', 'WindDir9am', 'WindDir3pm']:
    angles = df_imputed[col].map(wind_to_angle).fillna(0)
    df_imputed[f'{col}_sin'] = np.sin(np.deg2rad(angles))
    df_imputed[f'{col}_cos'] = np.cos(np.deg2rad(angles))

df_imputed['RainToday'] = (df_imputed['RainToday'] == 'Yes').astype(int)
df_imputed['RainTomorrow'] = (df_imputed['RainTomorrow'] == 'Yes').astype(int)
le = LabelEncoder()
df_imputed['Location_encoded'] = le.fit_transform(df_imputed['Location'])

feature_cols = num_cols + ['RainToday', 'Location_encoded'] + \
    [f'{c}_sin' for c in ['WindGustDir', 'WindDir9am', 'WindDir3pm']] + \
    [f'{c}_cos' for c in ['WindGustDir', 'WindDir9am', 'WindDir3pm']]

X_train = df_imputed.loc[train_mask, feature_cols].values
y_train = df_imputed.loc[train_mask, 'RainTomorrow'].values
X_val = df_imputed.loc[val_mask, feature_cols].values
y_val = df_imputed.loc[val_mask, 'RainTomorrow'].values
X_test = df_imputed.loc[test_mask, feature_cols].values
y_test = df_imputed.loc[test_mask, 'RainTomorrow'].values

scaler = StandardScaler()
X_train_scaled = scaler.fit_transform(X_train)
X_val_scaled = scaler.transform(X_val)
X_test_scaled = scaler.transform(X_test)

X_train_t = torch.FloatTensor(X_train_scaled).to(device)
y_train_t = torch.FloatTensor(y_train).to(device)
X_val_t = torch.FloatTensor(X_val_scaled).to(device)
y_val_t = torch.FloatTensor(y_val).to(device)
X_test_t = torch.FloatTensor(X_test_scaled).to(device)
y_test_t = torch.FloatTensor(y_test).to(device)

pos_weight = (y_train == 0).sum() / (y_train == 1).sum()
print(f'正例の重み: {pos_weight:.2f}')

# ========== モデル定義 ==========
class LogisticRegressionModel(nn.Module):
    def __init__(self, input_dim):
        super().__init__()
        self.linear = nn.Linear(input_dim, 1)
    def forward(self, x):
        return self.linear(x)

class MLPModel(nn.Module):
    def __init__(self, input_dim, hidden_dims, dropout=0.3):
        super().__init__()
        layers = []
        prev = input_dim
        for h in hidden_dims:
            layers.extend([nn.Linear(prev, h), nn.BatchNorm1d(h), nn.ReLU(), nn.Dropout(dropout)])
            prev = h
        layers.append(nn.Linear(prev, 1))
        self.net = nn.Sequential(*layers)
    def forward(self, x):
        return self.net(x)

def evaluate(y_true, y_proba):
    y_pred = (y_proba >= 0.5).astype(int)
    return {
        'accuracy': accuracy_score(y_true, y_pred),
        'precision': precision_score(y_true, y_pred, zero_division=0),
        'recall': recall_score(y_true, y_pred, zero_division=0),
        'f1': f1_score(y_true, y_pred, zero_division=0),
        'auc_roc': roc_auc_score(y_true, y_proba),
        'auc_pr': average_precision_score(y_true, y_proba)
    }

# ========== ロジスティック回帰 ==========
print('\n=== ロジスティック回帰のチューニング ===')

def train_lr(trial):
    lr = trial.suggest_float('lr', 1e-4, 1e-1, log=True)
    wd = trial.suggest_float('weight_decay', 1e-6, 1e-2, log=True)
    bs = trial.suggest_categorical('batch_size', [128, 256, 512])

    model = LogisticRegressionModel(X_train_t.shape[1]).to(device)
    criterion = nn.BCEWithLogitsLoss(pos_weight=torch.tensor([pos_weight]).to(device))
    optimizer = optim.Adam(model.parameters(), lr=lr, weight_decay=wd)
    loader = DataLoader(TensorDataset(X_train_t, y_train_t.unsqueeze(1)), batch_size=bs, shuffle=True)

    for epoch in range(50):
        model.train()
        for bx, by in loader:
            optimizer.zero_grad()
            loss = criterion(model(bx), by)
            loss.backward()
            optimizer.step()

    model.eval()
    with torch.no_grad():
        proba = torch.sigmoid(model(X_val_t)).cpu().numpy().flatten()
    return roc_auc_score(y_val, proba)

study_lr = optuna.create_study(direction='maximize')
study_lr.optimize(train_lr, n_trials=10, show_progress_bar=True)
print(f'Best Val AUC-ROC: {study_lr.best_value:.4f}')

# 最良パラメータで再学習
bp = study_lr.best_params
model_lr = LogisticRegressionModel(X_train_t.shape[1]).to(device)
criterion = nn.BCEWithLogitsLoss(pos_weight=torch.tensor([pos_weight]).to(device))
optimizer = optim.Adam(model_lr.parameters(), lr=bp['lr'], weight_decay=bp['weight_decay'])
loader = DataLoader(TensorDataset(X_train_t, y_train_t.unsqueeze(1)), batch_size=bp['batch_size'], shuffle=True)

lr_history = {'train_loss': [], 'val_loss': [], 'val_auc': []}
for epoch in range(100):
    model_lr.train()
    for bx, by in loader:
        optimizer.zero_grad()
        loss = criterion(model_lr(bx), by)
        loss.backward()
        optimizer.step()

    model_lr.eval()
    with torch.no_grad():
        tl = criterion(model_lr(X_train_t), y_train_t.unsqueeze(1)).item()
        vl = criterion(model_lr(X_val_t), y_val_t.unsqueeze(1)).item()
        vp = torch.sigmoid(model_lr(X_val_t)).cpu().numpy().flatten()
        va = roc_auc_score(y_val, vp)
    lr_history['train_loss'].append(tl)
    lr_history['val_loss'].append(vl)
    lr_history['val_auc'].append(va)

# 学習曲線
fig, axes = plt.subplots(1, 2, figsize=(12, 4))
axes[0].plot(lr_history['train_loss'], label='Train')
axes[0].plot(lr_history['val_loss'], label='Val')
axes[0].set_xlabel('Epoch'); axes[0].set_ylabel('Loss'); axes[0].legend()
axes[0].set_title('Logistic Regression - Loss')
axes[1].plot(lr_history['val_auc'], color='green')
axes[1].set_xlabel('Epoch'); axes[1].set_ylabel('AUC-ROC')
axes[1].set_title('Logistic Regression - AUC')
plt.tight_layout()
plt.savefig(FIGURES_PATH / 'learning_curves/logistic_regression.png', dpi=150)
plt.close()

model_lr.eval()
with torch.no_grad():
    y_proba_lr = torch.sigmoid(model_lr(X_test_t)).cpu().numpy().flatten()
metrics_lr = evaluate(y_test, y_proba_lr)
print(f'Test AUC-ROC: {metrics_lr["auc_roc"]:.4f}')

# ========== MLP ==========
print('\n=== MLPのチューニング ===')

def train_mlp(trial):
    n_layers = trial.suggest_int('n_layers', 1, 3)
    hidden = [trial.suggest_int(f'h{i}', 64, 256, step=64) for i in range(n_layers)]
    dropout = trial.suggest_float('dropout', 0.1, 0.4)
    lr = trial.suggest_float('lr', 1e-4, 1e-2, log=True)
    bs = trial.suggest_categorical('batch_size', [128, 256])

    model = MLPModel(X_train_t.shape[1], hidden, dropout).to(device)
    criterion = nn.BCEWithLogitsLoss(pos_weight=torch.tensor([pos_weight]).to(device))
    optimizer = optim.Adam(model.parameters(), lr=lr)
    loader = DataLoader(TensorDataset(X_train_t, y_train_t.unsqueeze(1)), batch_size=bs, shuffle=True)

    for epoch in range(50):
        model.train()
        for bx, by in loader:
            optimizer.zero_grad()
            loss = criterion(model(bx), by)
            loss.backward()
            optimizer.step()

    model.eval()
    with torch.no_grad():
        proba = torch.sigmoid(model(X_val_t)).cpu().numpy().flatten()
    return roc_auc_score(y_val, proba)

study_mlp = optuna.create_study(direction='maximize')
study_mlp.optimize(train_mlp, n_trials=10, show_progress_bar=True)
print(f'Best Val AUC-ROC: {study_mlp.best_value:.4f}')

bp = study_mlp.best_params
hidden = [bp[f'h{i}'] for i in range(bp['n_layers'])]
model_mlp = MLPModel(X_train_t.shape[1], hidden, bp['dropout']).to(device)
criterion = nn.BCEWithLogitsLoss(pos_weight=torch.tensor([pos_weight]).to(device))
optimizer = optim.Adam(model_mlp.parameters(), lr=bp['lr'])
loader = DataLoader(TensorDataset(X_train_t, y_train_t.unsqueeze(1)), batch_size=bp['batch_size'], shuffle=True)

mlp_history = {'train_loss': [], 'val_loss': [], 'val_auc': []}
best_auc = 0
best_state = None
for epoch in range(100):
    model_mlp.train()
    for bx, by in loader:
        optimizer.zero_grad()
        loss = criterion(model_mlp(bx), by)
        loss.backward()
        optimizer.step()

    model_mlp.eval()
    with torch.no_grad():
        tl = criterion(model_mlp(X_train_t), y_train_t.unsqueeze(1)).item()
        vl = criterion(model_mlp(X_val_t), y_val_t.unsqueeze(1)).item()
        vp = torch.sigmoid(model_mlp(X_val_t)).cpu().numpy().flatten()
        va = roc_auc_score(y_val, vp)
    mlp_history['train_loss'].append(tl)
    mlp_history['val_loss'].append(vl)
    mlp_history['val_auc'].append(va)
    if va > best_auc:
        best_auc = va
        best_state = {k: v.clone() for k, v in model_mlp.state_dict().items()}

model_mlp.load_state_dict(best_state)

fig, axes = plt.subplots(1, 2, figsize=(12, 4))
axes[0].plot(mlp_history['train_loss'], label='Train')
axes[0].plot(mlp_history['val_loss'], label='Val')
axes[0].set_xlabel('Epoch'); axes[0].set_ylabel('Loss'); axes[0].legend()
axes[0].set_title('MLP - Loss')
axes[1].plot(mlp_history['val_auc'], color='green')
axes[1].set_xlabel('Epoch'); axes[1].set_ylabel('AUC-ROC')
axes[1].set_title('MLP - AUC')
plt.tight_layout()
plt.savefig(FIGURES_PATH / 'learning_curves/mlp.png', dpi=150)
plt.close()

model_mlp.eval()
with torch.no_grad():
    y_proba_mlp = torch.sigmoid(model_mlp(X_test_t)).cpu().numpy().flatten()
metrics_mlp = evaluate(y_test, y_proba_mlp)
print(f'Test AUC-ROC: {metrics_mlp["auc_roc"]:.4f}')

# ========== Random Forest ==========
print('\n=== Random Forestのチューニング ===')

def train_rf(trial):
    n_est = trial.suggest_int('n_estimators', 100, 300, step=50)
    depth = trial.suggest_int('max_depth', 10, 30)
    model = RandomForestClassifier(n_estimators=n_est, max_depth=depth, class_weight='balanced', n_jobs=-1, random_state=SEED)
    model.fit(X_train_scaled, y_train)
    return roc_auc_score(y_val, model.predict_proba(X_val_scaled)[:, 1])

study_rf = optuna.create_study(direction='maximize')
study_rf.optimize(train_rf, n_trials=10, show_progress_bar=True)
print(f'Best Val AUC-ROC: {study_rf.best_value:.4f}')

bp = study_rf.best_params
model_rf = RandomForestClassifier(n_estimators=bp['n_estimators'], max_depth=bp['max_depth'], class_weight='balanced', n_jobs=-1, random_state=SEED)
model_rf.fit(X_train_scaled, y_train)
y_proba_rf = model_rf.predict_proba(X_test_scaled)[:, 1]
metrics_rf = evaluate(y_test, y_proba_rf)
print(f'Test AUC-ROC: {metrics_rf["auc_roc"]:.4f}')

# 特徴量重要度
feat_imp = pd.DataFrame({'feature': feature_cols, 'importance': model_rf.feature_importances_}).sort_values('importance', ascending=False)
fig, ax = plt.subplots(figsize=(10, 8))
sns.barplot(data=feat_imp.head(15), x='importance', y='feature', ax=ax, palette='viridis')
ax.set_title('Random Forest - Feature Importance')
plt.tight_layout()
plt.savefig(FIGURES_PATH / 'feature_importance/random_forest.png', dpi=150)
plt.close()

# ========== 結果比較 ==========
print('\n=== 最終結果 ===')
results = {'Logistic Regression': metrics_lr, 'MLP': metrics_mlp, 'Random Forest': metrics_rf}
results_df = pd.DataFrame(results).T.round(4)
print(results_df)

# 比較グラフ
fig, axes = plt.subplots(2, 3, figsize=(15, 10))
metrics_list = ['accuracy', 'precision', 'recall', 'f1', 'auc_roc', 'auc_pr']
titles = ['Accuracy', 'Precision', 'Recall', 'F1-Score', 'AUC-ROC', 'AUC-PR']
colors = ['#3498db', '#e74c3c', '#2ecc71']
for idx, (m, t) in enumerate(zip(metrics_list, titles)):
    ax = axes[idx//3, idx%3]
    vals = [results[k][m] for k in results]
    bars = ax.bar(results.keys(), vals, color=colors)
    ax.set_title(t); ax.set_ylim(0, 1)
    ax.tick_params(axis='x', rotation=45)
    for bar, v in zip(bars, vals):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.02, f'{v:.3f}', ha='center')
plt.tight_layout()
plt.savefig(FIGURES_PATH / 'model_comparison.png', dpi=150)
plt.close()

# ROCカーブ
fig, ax = plt.subplots(figsize=(8, 8))
for name, proba in [('Logistic Regression', y_proba_lr), ('MLP', y_proba_mlp), ('Random Forest', y_proba_rf)]:
    fpr, tpr, _ = roc_curve(y_test, proba)
    ax.plot(fpr, tpr, label=f'{name} (AUC={roc_auc_score(y_test, proba):.4f})', linewidth=2)
ax.plot([0, 1], [0, 1], 'k--')
ax.set_xlabel('FPR'); ax.set_ylabel('TPR'); ax.set_title('ROC Curve Comparison')
ax.legend(); ax.grid(alpha=0.3)
plt.tight_layout()
plt.savefig(FIGURES_PATH / 'roc_curves/comparison.png', dpi=150)
plt.close()

# 混同行列
fig, axes = plt.subplots(1, 3, figsize=(15, 5))
for idx, (name, proba) in enumerate([('Logistic Regression', y_proba_lr), ('MLP', y_proba_mlp), ('Random Forest', y_proba_rf)]):
    cm = confusion_matrix(y_test, (proba >= 0.5).astype(int))
    sns.heatmap(cm, annot=True, fmt='d', cmap='Blues', ax=axes[idx], xticklabels=['No', 'Yes'], yticklabels=['No', 'Yes'])
    axes[idx].set_title(name); axes[idx].set_xlabel('Predicted'); axes[idx].set_ylabel('Actual')
plt.tight_layout()
plt.savefig(FIGURES_PATH / 'confusion_matrices/comparison.png', dpi=150)
plt.close()

# JSON保存
tuning_results = {
    'logistic_regression': {'best_params': study_lr.best_params, 'best_val_auc': study_lr.best_value, 'test_metrics': {k: float(v) for k, v in metrics_lr.items()}},
    'mlp': {'best_params': study_mlp.best_params, 'hidden_dims': hidden, 'best_val_auc': study_mlp.best_value, 'test_metrics': {k: float(v) for k, v in metrics_mlp.items()}},
    'random_forest': {'best_params': study_rf.best_params, 'best_val_auc': study_rf.best_value, 'test_metrics': {k: float(v) for k, v in metrics_rf.items()}}
}
with open(EXPERIMENTS_PATH / 'tuning_results.json', 'w') as f:
    json.dump(tuning_results, f, indent=2)

print('\n完了！')
print(f'ベストモデル: {results_df["auc_roc"].idxmax()} (AUC-ROC: {results_df["auc_roc"].max():.4f})')
