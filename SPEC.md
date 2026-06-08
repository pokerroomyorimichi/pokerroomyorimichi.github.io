# YORIMICHI アンケートシステム 仕様書

**最終更新:** 2026年6月  
**作成:** BeeBloom合同会社 システム・IT部  
**担当窓口:** 太田 清人（ota@beebloom.fun）

---

## 1. 概要

ポーカールーム「YORIMICHI」来店時にQRコードを読み取ったお客様に対し、PPPアカウント登録前に流入経路・来店動機・ポーカー経験を収集するアンケートシステム。

### 目的
- マーケティング施策の効果測定
- 流入経路の把握（SNS / 口コミ / 検索など）
- お客様属性の把握（ポーカー経験）

---

## 2. システム構成

```
お客様がQRをスキャン
       ↓
アンケートページ（GitHub Pages）
https://pokerroomyorimichi.github.io/survey/
       ↓ 回答送信（バックグラウンド）
GAS Web App API
https://script.google.com/macros/s/AKfycbwwH3q2ynIl_.../exec
       ↓ 記録
Googleスプレッドシート
https://docs.google.com/spreadsheets/d/17iU4i2FQjdMJ-_ccSJ-fYJDZ5kUkM7fOasAlQNbtMyM/edit
       ↓ 3秒後に自動転送
PPPアカウント登録ページ
https://ppp.arkadia.bet/login
```

### 技術構成

| 役割 | 技術 | 備考 |
|---|---|---|
| フロントエンド（画面） | HTML/CSS/JS | GitHub Pages でホスト |
| バックエンドAPI | Google Apps Script | 回答の受け取り・記録 |
| データベース | Google スプレッドシート | 回答を自動蓄積 |
| ホスティング | GitHub Pages | 無料・警告バナーなし |

---

## 3. ユーザーフロー

```
[QRスキャン]
    ↓
[Q1] YORIMICHIをどこで知りましたか？（複数選択）
    ↓ 1つ以上選択で「次へ」が有効化
[Q2] 今回ご来店の決め手は何ですか？（複数選択）
    ↓ 1つ以上選択で「次へ」が有効化
[Q3] ポーカーのご経験を教えてください（単一選択）
    ↓ 1つ選択で「登録へ進む」が有効化
[送信 → 完了画面（3秒）]
    ↓ 自動転送
[PPPアカウント登録]
```

**設計ポイント：**
- 何も選択しないと次の設問に進めない（バリデーション済み）
- GAS APIへの送信が失敗してもPPPページへは必ず転送する（登録フローをブロックしない）
- 送信はバックグラウンドで行うため、お客様の操作を妨げない

---

## 4. 設問・選択肢

### Q1：YORIMICHIをどこで知りましたか？（複数選択可）
1. Instagramの投稿
2. Instagramの広告
3. X（Twitter）
4. 友人・知人の紹介
5. Google検索
6. Googleマップ・口コミ
7. 看板・チラシ
8. その他

### Q2：今回ご来店の決め手は何ですか？（複数選択可）
1. Instagramの投稿
2. Instagramの広告
3. X（Twitter）
4. 友人・知人の紹介
5. Google検索
6. Googleマップ・口コミ
7. 看板・チラシ
8. その他

### Q3：ポーカーのご経験を教えてください（単一選択）
1. 未経験（初めて）
2. 初心者（数回程度）
3. 経験者（定期的にプレイ）
4. 上級者（競技経験あり）

---

## 5. ファイル構成

### GitHub リポジトリ
**URL:** https://github.com/pokerroomyorimichi/pokerroomyorimichi.github.io  
**アカウント:** pokerroomyorimichi  
**ローカルパス:** `~/Desktop/BeeBloom/pokerroomyorimichi.github.io/`

```
pokerroomyorimichi.github.io/
├── survey/
│   ├── index.html   ← アンケート画面（選択肢はここを編集）
│   └── logo.png     ← YORIMICHIロゴ（完了画面に表示）
└── SPEC.md          ← 本仕様書
```

### GAS プロジェクト
**プロジェクト名:** YORIMICHI アンケートAPI  
**管理アカウント:** ota@yamaguchi-gp.com  
**ローカルバックアップ:** `~/.beebloom/07_Development/survey/Code.gs`

### スプレッドシート
**ファイル名:** YORIMICHI アンケート回答  
**URL:** https://docs.google.com/spreadsheets/d/17iU4i2FQjdMJ-_ccSJ-fYJDZ5kUkM7fOasAlQNbtMyM/edit  
**格納フォルダ:** https://drive.google.com/drive/folders/186JV0YdnCTqtu0JW_YAJeNFiwspm3XOr

---

## 6. スプレッドシート構成

### 「アンケート回答」シート

| 列 | 項目名 | 内容 |
|---|---|---|
| A | タイムスタンプ | 回答日時（例: 2026/06/08 19:30:00） |
| B | Q1_どこで知った | 複数選択の場合「、」区切り |
| C | Q2_来店の決め手 | 複数選択の場合「、」区切り |
| D | Q3_ポーカー経験 | 単一値 |

### 「流入経路分析」シート

月ごとの流入経路カウントを自動集計（数式ベース・手動操作不要）。  
回答が追加されると自動で反映される。

| 列 | 内容 |
|---|---|
| A | 年月（例: 2026/06） |
| B〜G | 各流入経路のカウント数 |
| H | 合計（人） |

---

## 7. 選択肢の変更手順

選択肢はGitHubのHTMLファイルを編集するだけで変更できる。  
**GASの再デプロイ・QRコードの変更は不要。**

### 手順

1. `~/Desktop/BeeBloom/pokerroomyorimichi.github.io/survey/index.html` をテキストエディタで開く

2. 以下の箇所を編集する（ファイル上部のJavaScript）：

```javascript
// Q1の選択肢
const REFERRAL_OPTIONS = [
  'Instagramの投稿',
  '友人・知人の紹介',
  // ...追加・削除・変更
];

// Q2の選択肢
const MOTIVE_OPTIONS = [
  'Instagramの投稿',
  // ...
];

// Q3の選択肢（4択を維持推奨）
const EXPERIENCE_OPTIONS = [
  '未経験（初めて）',
  // ...
];
```

3. ターミナルで以下を実行：

```bash
cd ~/Desktop/BeeBloom/pokerroomyorimichi.github.io
git add survey/index.html
git commit -m "Update survey options"
git push
```

4. 1〜2分後に反映される

---

## 8. GASの再デプロイ手順

Code.gsを変更した場合のみ必要。  
（選択肢の変更だけであれば不要）

1. https://script.google.com を開く
2. 「YORIMICHI アンケートAPI」プロジェクトを開く
3. 「デプロイ」→「デプロイを管理」
4. 鉛筆アイコン（編集）→「バージョン」を「新しいバージョン」に変更
5. 「デプロイ」をクリック

**デプロイ設定（変更しないこと）：**
- 次のユーザーとして実行: **自分**
- アクセスできるユーザー: **全員（匿名ユーザーを含む）**

---

## 9. QRコード

**QRに設定するURL:**
```
https://pokerroomyorimichi.github.io/survey/
```

選択肢を変更してもURLは変わらないため、QRコードの作り直しは不要。  
ただし、URLそのものが変わる場合（ドメイン変更など）は作り直しが必要。

---

## 10. トラブルシューティング

| 症状 | 原因 | 対処 |
|---|---|---|
| 選択肢がクリックできない | index.htmlの記述エラー | GitHubの変更履歴で直前のバージョンに戻す |
| スプレッドシートに記録されない | GASのデプロイ設定が「全員」になっていない | §8の手順で再デプロイ（アクセス設定を確認） |
| PPPページに飛ばない | JavaScriptのエラー | ブラウザのコンソールでエラー内容を確認 |
| ページが表示されない | GitHub Pagesが無効 | GitHubリポジトリの Settings → Pages を確認 |

---

## 11. 関係者・アクセス権限

| リソース | 管理者アカウント |
|---|---|
| GitHub（pokerroomyorimichi） | ota@beebloom.fun |
| GAS（アンケートAPI） | ota@yamaguchi-gp.com |
| スプレッドシート | ota@yamaguchi-gp.com |
| ローカルファイル | `/Users/kiyotoota/Desktop/BeeBloom/` |
