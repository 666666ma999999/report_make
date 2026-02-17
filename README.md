# 月次報告書生成システム

上場会社の事業責任者が社外役員・監査役向けに月次報告を生成するためのツール。

## ディレクトリ構成

```
report/
├── boradmtg/
│   ├── csv/                          # 元データ（gitignore対象）
│   │   └── デジコン収益分析_YYYYMM.xlsx
│   ├── analyze_202601.py             # 全体変動分析スクリプト
│   ├── segment_analysis.py           # セグメント別分析スクリプト
│   ├── generate_board_report.py      # Excel報告書生成スクリプト
│   ├── analysis_YYYYMM.csv          # 分析結果（gitignore対象）
│   ├── segment_analysis_YYYYMM.csv  # セグメント分析結果（gitignore対象）
│   └── 月次報告_YYYYMM.xlsx         # 生成レポート（gitignore対象）
├── .gitignore
└── README.md
```

## 使い方

### 1. データ配置
`boradmtg/csv/` に対象月のxlsxファイルを配置

### 2. 分析実行
```bash
python3 boradmtg/analyze_202601.py
python3 boradmtg/segment_analysis.py
```

### 3. レポート生成
```bash
python3 boradmtg/generate_board_report.py
```

## 機密データについて
- `csv/` 配下の元データxlsx
- 生成された分析CSV
- 生成されたレポートxlsx

上記は `.gitignore` で除外されています。リポジトリにpushしないでください。
