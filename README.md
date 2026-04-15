# TA Support Tool

大学講義のTA（ティーチングアシスタント）業務を効率化するWebツールです。

## 機能

✨ **座席配置生成**
- グリッド配置（カスタマイズ可能）
- 混合配置（3人席＋2人席）
- シャッフル機能付き

📝 **課題採点シート作成**
- 課題数をカスタマイズ可能
- 自動集計式搭載
- Excel形式で出力

📋 **出席票生成**
- 講義日程の自動計算
- 出席/欠席チェック欄
- Excel形式で出力

## 使い方

### 1. ブラウザでアクセス

[https://yourusername.github.io/ta-support-tool/](https://yourusername.github.io/ta-support-tool/)

### 2. 名簿をアップロード

CSVファイルをアップロードしてください。必要な列：
- `学籍番号`：学生の学籍番号
- `学生氏名`：学生の名前

**CSVのサンプル形式：**
```
(6行のヘッダー行)
学籍番号,学生氏名
S001,山田太郎
S002,佐藤花子
```

### 3. 座席配置を設定

- **グリッド配置**：列数と行数を指定
- **混合配置**：3人席と2人席の配置を指定

### 4. 出力ファイルを選択

- 座席表（推奨常時）
- 課題採点シート（必要に応じて）
- 出席票（必要に応じて）

### 5. ファイルを生成

ボタンをクリックして、Excelファイルをダウンロードしてください。

## 技術スタック

- **フロントエンド**：HTML5, CSS3, Vanilla JavaScript
- **ライブラリ**：XLSX (Excel操作)
- **ホスティング**：GitHub Pages
- **ブラウザ互換性**：Chrome, Firefox, Safari, Edge

## ファイル構成

```
ta-support-tool/
├── index.html          # メインページ
├── css/
│   └── style.css       # スタイルシート
├── js/
│   ├── app.js          # メインアプリケーション
│   ├── utils.js        # ユーティリティ関数
│   └── generator.js    # Excel生成ロジック
├── README.md           # このファイル
├── LICENSE             # ライセンス
└── .gitignore          # Git除外設定
```

## ローカル実行

1. リポジトリをクローン
```bash
git clone https://github.com/yourusername/ta-support-tool.git
cd ta-support-tool
```

2. ローカルサーバーで実行
```bash
# Python 3
python3 -m http.server 8000

# Node.js
npx http-server
```

3. ブラウザでアクセス：`http://localhost:8000`

## 制限事項

- ⚠️ Excel出力は5列以下の座席配置に最適化しています
- ⚠️ 大人数（1000名以上）の名簿は動作が遅くなる可能性があります
- ⚠️ Internet Explorer には対応していません

## 今後の改善予定

- [ ] 座席配置のプレビュー機能
- [ ] 複数名簿の同時処理
- [ ] Google Sheetsとの連携
- [ ] 成績分布の可視化
- [ ] ダークモード対応

## ライセンス

MIT License - 自由に使用・改変・配布OK

## 貢献

バグ報告や機能リクエストは、Issues でお願いします。

---

作成者：教育支援チーム
