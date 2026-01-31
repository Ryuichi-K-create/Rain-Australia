# オーストラリア天気予測プロジェクト

## 概要
オーストラリアの天気データ（weatherAUS.csv）を使用して、「明日が雨になるか否か」の2値分類を行う。

## データセット
- **ソース**: [Kaggle - Rain in Australia](https://www.kaggle.com/datasets/jsphyg/weather-dataset-rattle-package/data)
- **期間**: 2007年11月〜2017年6月（約10年分）
- **観測地点**: 49箇所
- **レコード数**: 145,460件
- **特徴量**: 23カラム
- **ターゲット**: RainTomorrow（Yes/No）
- **クラス分布**: No(78%) / Yes(22%) - 不均衡データ

## 技術スタック
- **言語**: Python
- **フレームワーク**: PyTorch
- **ハイパーパラメータ最適化**: Optuna
- **可視化**: matplotlib, seaborn
- **言語**: 日本語（スライド・コメント）

## 使用モデル
1. **ロジスティック回帰** - 線形モデル（PyTorch実装）
2. **MLP (Multi-Layer Perceptron)** - ニューラルネット（PyTorch実装）

## プロジェクト構造
```
RainAustralia/
├── CLAUDE.md                          # このファイル
├── data/
│   └── weatherAUS.csv                 # 入力データ
├── rain_prediction_australia.ipynb    # メインNotebook
├── presentation.tex                   # LaTeXスライド
├── figures/                           # 生成図表
│   ├── target_distribution.png
│   ├── missing_values.png
│   ├── learning_curves/
│   ├── confusion_matrices/
│   ├── roc_curves/
│   └── feature_importance/
└── experiments/
    └── tuning_results.json            # チューニング記録
```

## 重要な設計判断
1. **時系列分割**: ランダム分割ではなく時間順序を保持
   - Train: 〜2015/06 (80%)
   - Val: 2015/07〜2016/06 (10%)
   - Test: 2016/07〜 (10%)

2. **欠損値処理**: Location別中央値/最頻値補完

3. **クラス不均衡対策**: Weighted BCE Loss または Focal Loss

4. **風向エンコーディング**: サイクリカルエンコーディング（sin/cos）

## 評価指標
- Accuracy
- Precision
- Recall
- F1-Score
- AUC-ROC（主指標）
- AUC-PR

## 成果物
1. Jupyter Notebook（全コード・結果を含む）
2. LaTeX Beamerスライド（22スライド構成）
3. 可視化図表一式
4. パラメータチューニング記録
