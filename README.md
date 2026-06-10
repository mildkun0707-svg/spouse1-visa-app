# 配偶者ビザ 質問書 自動作成システム（先生送信版）

クライアントがフォームに入力 → 先生（行政書士）にメール送信 → 先生がExcelを作成

## しくみ

```
クライアント画面（パスワード不要）
  └─ フォーム入力
  └─ 「📨 先生に送信する」ボタン
       └─ 先生のメールにJSONが届く

先生画面（サイドバーで管理者ログイン）
  └─ 受け取ったJSONを読み込む
  └─ 内容を確認・修正
  └─ 「✨ 質問書Excelを作成する」→ ダウンロード
```

> **クライアントの画面にはExcel作成・ダウンロードボタンは表示されません。**

---

## ファイル構成

| ファイル | 役割 |
|---|---|
| `streamlit_app.py` | アプリ本体（クライアント＋先生モード統合） |
| `fill_questionnaire.py` | 転記エンジン（JSON → Excel） |
| `master_schema.json` | 全219項目のマスタ定義 |
| `template.xlsx` | 入管公式書式テンプレート（転記先） |
| `requirements.txt` | 必要ライブラリ |
| `.streamlit/secrets.toml.example` | 設定テンプレート（コピーして使う） |

---

## セットアップ（ローカル）

```bash
# 1. ライブラリをインストール
pip install -r requirements.txt

# 2. シークレット設定
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
# → secrets.toml を開いてパスワード・メール設定を書き換える

# 3. 起動
streamlit run streamlit_app.py
```

---

## Streamlit Cloud への公開

1. このリポジトリをGitHubに push
2. [share.streamlit.io](https://share.streamlit.io) でアプリを作成
3. アプリ管理画面 → **Settings → Secrets** に `secrets.toml` の内容を貼り付け

> ⚠️ `secrets.toml` は `.gitignore` に含まれており、GitHubには **アップロードされません**。
> Streamlit Cloud の Secrets 画面にのみ設定してください。

---

## Gmailのアプリパスワード取得方法

1. Googleアカウントで2段階認証を有効にする
2. [myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords) にアクセス
3. 発行された16桁を `secrets.toml` の `password` に貼り付ける

---

## 注意

- このシステムは業務効率化ツールです。申請の許可を保証するものではありません。
- 出力後は必ず内容を目視確認し、ご署名のうえ提出してください。
- 入力データ（JSON）にはクライアントの個人情報が含まれます。取り扱いにご注意ください。
