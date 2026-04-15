# GitHub へのプッシュ手順

GitHub CLIがインストールされていないため、以下のいずれかの方法を選択してください。

## 方法 1: GitHub Web UI を使用（簡単）

1. GitHub にログイン：https://github.com/new
2. リポジトリ名を入力：`ta-support-tool`
3. 説明を入力：`TA Support Tool - 座席配置・課題採点シート・出席票生成ツール`
4. 「Public」を選択
5. 「Create repository」をクリック

その後、以下をコマンドラインで実行：

```bash
cd /Users/x21095xx/個人事業/ta-support-tool

# リモートリポジトリを追加（YOUR_USERNAME を自分のGitHubユーザー名に置き換え）
git remote add origin https://github.com/YOUR_USERNAME/ta-support-tool.git
git branch -M main
git push -u origin main
```

## 方法 2: GitHub CLI をインストール (推奨)

```bash
# Homebrew でインストール（macOS）
brew install gh

# GitHub に認証
gh auth login

# リポジトリを作成してプッシュ
gh repo create ta-support-tool --public --source=. --remote=origin --push
```

## 方法 3: SSH キーを使用（既に設定済みの場合）

```bash
cd /Users/x21095xx/個人事業/ta-support-tool

# SSH でリモート追加
git remote add origin git@github.com:YOUR_USERNAME/ta-support-tool.git
git branch -M main
git push -u origin main
```

---

### GitHub Pages の設定

リポジトリをプッシュ後、以下の手順で GitHub Pages を有効にしてください：

1. GitHub でリポジトリを開く
2. Settings → Pages
3. Branch：`main`、Folder：`/ (root)` を選択
4. Save

その後、サイトは以下の URL で利用可能になります：
`https://YOUR_USERNAME.github.io/ta-support-tool/`
