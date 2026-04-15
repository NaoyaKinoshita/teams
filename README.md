# Teams Recording Bot

Teams 会議に自動参加し、録画の開始・終了を会議チャットに通知する Bot。

## アーキテクチャ概要

```
recording-bot@... を会議に招待
        ↓（カレンダー変更通知）
  Graph API Webhook
        ↓
  bot が会議に自動参加
  （Graph Communications API）
        ↓
  録画開始を検知 → チャットに通知
        ↓
  会議終了を検知 → チャットに通知
        ↓
  OneDrive 保存後に録画 URL を通知
```

### 設計方針

- `recording-bot@PersonalDev189.onmicrosoft.com` を会議に招待することで、録画を Azure に連携する意図を表明する
- Bot はカレンダー通知を受け取り、会議開始時刻に自動参加する
- 録画開始・終了・OneDrive 保存完了をチャットにテキスト通知する（Adaptive Card での意思確認は不要）

---

## 構成

| 項目 | 値 |
|------|----|
| 言語 | Python |
| テナント | PersonalDev189.onmicrosoft.com |
| Bot 名 | teams-recording-bot |
| リソースグループ | teams-recording-rg |
| Bot ユーザー | recording-bot@PersonalDev189.onmicrosoft.com |

---

## セットアップ手順

### 1. Microsoft 365 Business の契約

- 組織テナント（`xxxx.onmicrosoft.com`）が必要
- Microsoft 365 Personal では以下が使えないため不可
  - Teams 管理センター
  - カスタムアプリのデプロイ
  - Graph API の管理者同意

### 2. Azure サブスクリプションの作成

M365 Business には Azure サブスクリプションが含まれないため別途作成が必要。

1. [azure.microsoft.com/free](https://azure.microsoft.com/ja-jp/free/) にアクセス
2. 組織アカウント（`xxxx@xxxx.onmicrosoft.com`）でサインイン
3. 無料アカウントを作成（$200 クレジット / 30日間）

### 3. Azure Bot Service の作成

1. [portal.azure.com](https://portal.azure.com) にアクセス
2. 「Azure Bot」を検索 → 「+ 作成」
3. 以下の設定で作成：

| 項目 | 値 |
|------|----|
| ボット ハンドル | `teams-recording-bot` |
| リソース グループ | `teams-recording-rg` |
| 価格レベル | F0（無料） |
| Microsoft App ID | 新しい Microsoft アプリ ID を作成する |

### 4. Entra ID アプリ登録の設定

Azure Bot 作成後、Entra ID に自動作成されるアプリ登録を設定する。

1. [entra.microsoft.com](https://entra.microsoft.com) にアクセス
2. 「アプリの登録」→ `teams-recording-bot` を開く
3. 以下の値をメモ：
   - アプリケーション（クライアント）ID
   - ディレクトリ（テナント）ID

#### マルチテナント設定（必須）

Bot Framework の認証要件により、マルチテナントアプリとして設定する必要がある。

1. 「認証 (Preview)」を開く
2. 「サポートされているアカウントの種類」→「任意の組織ディレクトリ内のアカウント（マルチテナント）」に変更
3. 保存

#### クライアントシークレットの作成

1. 「証明書とシークレット」→「+ 新しいクライアントシークレット」
2. 説明: `teams-bot-secret`、有効期限: 24ヶ月
3. 作成直後に表示される「値」をコピー（後から確認不可）

#### API アクセス許可の設定

「API のアクセス許可」→「+ アクセス許可の追加」→「Microsoft Graph」→「アプリケーションの許可」

| 権限 | 用途 |
|------|------|
| `Calendars.Read` | recording-bot のカレンダー変更通知の受信 |
| `CallRecords.Read.All` | 通話記録の取得（録画 URL 検索） |
| `Calls.AccessMedia.All` | 会議メディアへのアクセス |
| `Calls.InitiateGroupCalls.All` | 通話の開始 |
| `Calls.JoinGroupCalls.All` | 会議への参加 |
| `Chat.ReadWrite.All` | 会議チャットへのメッセージ送信 |
| `OnlineMeetings.Read.All` | 会議情報の取得 |

追加後、**「管理者の同意を与えます」** をクリックして全権限を承認。

### 5. recording-bot ユーザーの作成

会議に招待するための専用ユーザーを M365 管理センターで作成する。

1. [admin.microsoft.com](https://admin.microsoft.com) にアクセス
2. 「ユーザー」→「アクティブなユーザー」→「ユーザーの追加」
3. 以下の設定で作成：
   - 表示名: `recording-bot`
   - ユーザー名: `recording-bot@<テナント>.onmicrosoft.com`
   - ライセンス: Microsoft 365 Business Basic（Teams 参加に必要、約 $6/月）
4. 作成後、Entra ID でオブジェクト ID をメモする（`RECORDING_BOT_USER_ID` として使用）

### 6. Teams 管理センターの設定

#### カスタムアプリのアップロード許可

1. [admin.teams.microsoft.com](https://admin.teams.microsoft.com) にアクセス
2. 「Teams のアプリ」→「セットアップ ポリシー」
3. 「グローバル（組織全体の既定値）」を開く
4. 「カスタム アプリのアップロード」を **オン** に変更して保存

#### Teams アプリのアップロード

1. `bot/manifest/` 配下の `manifest.json`・`color.png`・`outline.png` を ZIP に圧縮
2. Teams クライアント → 「アプリ」→「アプリを管理」→「アプリをアップロード」
3. ZIP ファイルをアップロードして Bot をインストール

### 7. CsApplicationAccessPolicy の設定（Bot が会議参加するために必須）

PowerShell から Teams モジュールを使って Bot アプリに会議アクセスを許可する。

```powershell
# Teams モジュールのインストール（未インストールの場合）
Install-Module -Name MicrosoftTeams -Force

# 接続
Connect-MicrosoftTeams

# ポリシー作成
New-CsApplicationAccessPolicy `
  -Identity "AllowBotJoinMeetingPolicy" `
  -AppIds "<APP_ID>" `
  -Description "Allow recording bot to join Teams meetings"

# テナント全体に適用
Grant-CsApplicationAccessPolicy -PolicyName "AllowBotJoinMeetingPolicy" -Global
```

> **注意**: `-Global` 適用の反映には最大 1〜2 時間かかる場合がある。

### 8. ngrok のセットアップ（ローカル開発用）

Graph API の Webhook は HTTPS の公開 URL が必要なため、ngrok でローカルポートを公開する。

```bash
# インストール（macOS）
brew install ngrok

# 起動
ngrok http 3978
```

表示された `https://xxxx.ngrok-free.app` を `.env` の `NOTIFICATION_URL` に設定する。

---

## 環境変数

```env
MICROSOFT_APP_ID=<アプリケーション（クライアント）ID>
MICROSOFT_APP_SECRET=<クライアントシークレット>
TENANT_ID=<ディレクトリ（テナント）ID>
NOTIFICATION_URL=https://<ngrok のホスト名>
RECORDING_BOT_USER_ID=<recording-bot ユーザーのオブジェクト ID>
PORT=3978
```

---

## ローカル起動

```bash
cd bot
uv run python app.py
```

起動後、以下のサブスクリプションが自動登録される：
- `callRecords` 通知（録画 URL 取得用フォールバック）
- recording-bot のカレンダー変更通知（会議参加スケジュール用）

---

## ディレクトリ構成

```
teams/
├── README.md
└── bot/
    ├── app.py            # aiohttp サーバー（3 エンドポイント）
    ├── bot.py            # Teams Bot ロジック（会議イベント処理）
    ├── config.py         # 環境変数の読み込み
    ├── graph_client.py   # Graph API クライアント（参加・通知処理）
    ├── adaptive_card.py  # Adaptive Card 定義
    ├── pyproject.toml
    ├── .env              # 環境変数（git 管理外）
    └── manifest/
        ├── manifest.json
        ├── color.png
        └── outline.png
```

### エンドポイント

| パス | 用途 |
|------|------|
| `POST /api/messages` | Bot Framework メッセージ受信 |
| `POST /api/notifications` | Graph callRecords・カレンダー変更通知 |
| `POST /api/calls` | Graph Communications API コールバック |

---

## 通知フロー詳細

```
recording-bot を会議に招待
        ↓
カレンダー変更通知 → /api/notifications
        ↓
会議開始時刻まで待機（asyncio.sleep）
        ↓
POST /communications/calls（会議参加リクエスト）
        ↓
/api/calls コールバック
  ├── state: establishing  → 待機
  ├── state: established   → _active_calls に登録、_call_threads に保存
  ├── state: terminated (subcode 2203)
  │     → 会議未開始のため 2 分後にリトライ（最大 10 回）
  └── state: terminated (正常終了)
        → 「会議が終了しました。」をチャットに送信
        → callRecord API を最大 5 回ポーリングして録画 URL を通知

録画開始時:
  recordingStatus: recording → 「録画が開始されました。」をチャットに送信

callRecords 通知 → /api/notifications
  → sessions > segments > recordings から contentUrl を取得して通知
  → 既に通知済みの場合はスキップ（_notified_recordings で管理）
```

---

## 使い方

1. Teams または Outlook でオンライン会議を作成する
2. 出席者に `recording-bot@PersonalDev189.onmicrosoft.com` を追加する
3. Bot が自動的に会議に参加し、録画の状態を会議チャットに通知する

---

## TODO

- [x] Python Bot コードの実装
- [x] Graph API Webhook によるカレンダー変更通知の受信
- [x] Graph Communications API による会議自動参加
- [x] 会議未開始時のリトライ処理（2 分間隔、最大 10 回）
- [x] キャンセル済み会議のスキップ処理
- [x] 録画開始・終了のチャット通知
- [x] callRecord API による録画 URL 取得・通知
- [x] CsApplicationAccessPolicy の設定
- [ ] 録画 URL のポーリング動作確認（Teams 録画では contentUrl が返らない可能性あり）
- [ ] OneDrive 変更通知による確実な録画 URL 通知（`Files.Read.All` 権限が必要）
- [ ] Azure Container Apps へのデプロイ
