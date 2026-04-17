# Meeting Recording Monitor

Teams アプリとして会議に自動追加され、録画の開始・停止を検知して Azure に通知する Bot。

- 専用ユーザーアカウント不要（Bot アプリ ID のみで動作）
- Teams Premium 不要（E3 ライセンスで動作）
- 管理センターの App Setup Policy でセキュリティグループ単位に配布可能

---

## 動作フロー

```
管理者が App Setup Policy で対象グループにアプリを配布
    ↓
対象ユーザーが会議を作成
    ↓
/chats 購読 でバックエンドが会議チャットを自動検知
    ↓
会議チャットへアプリをインストール + チャットメッセージを購読
    ↓
録画が開始（callRecordingStatus = initial）
    ↓
会議サイドパネルのタブがポーリングで録画開始を検知
    ↓
タブ内に「Azure に連携する」「スキップ」ボタンを表示
    ├── 「Azure に連携する」クリック → 同意を POST /api/consent に記録
    └── 「スキップ」クリック → 不同意を記録
    ↓
録画が停止（callRecordingStatus = chunkFinished）
    └── 同意済みの場合のみ Azure Webhook（AZURE_WEBHOOK_URL）に通知
```

---

## ディレクトリ構成

```
meeting-app/
├── app.py              # aiohttp サーバー・ルーティング
├── bot.py              # Bot Framework ハンドラー
├── graph_client.py     # Graph API 操作・録画状態管理
├── adaptive_card.py    # Adaptive Card 定義（将来拡張用）
├── config.py           # 環境変数読み込み
├── manifest/
│   ├── manifest.json   # Teams アプリマニフェスト
│   ├── color.png       # カラーアイコン（192×192 px）
│   └── outline.png     # アウトラインアイコン（32×32 px）
└── .env.example        # 環境変数テンプレート
```

---

## API エンドポイント

| メソッド | パス | 説明 |
|---|---|---|
| POST | `/api/messages` | Bot Framework からのメッセージ受信 |
| POST | `/api/notifications` | Graph API 変更通知受信 |
| GET | `/tab` | 会議サイドパネルタブの HTML |
| POST | `/api/tab-context` | タブからの threadId 受け取り |
| GET | `/api/recording-status` | 録画状態のポーリング（タブ用） |
| POST | `/api/consent` | Azure 連携の同意/スキップ受け取り |

---

## セットアップ手順

### 1. Entra ID アプリ登録の確認

既存の Bot アプリ登録（Azure Portal → Entra ID → アプリの登録）に以下の API 権限が付与されていること（管理者同意済み）：

| 権限 | 種別 | 用途 |
|------|------|------|
| `Calendars.Read` | Application | メールボックスのカレンダー読み取り |
| `CallRecords.Read.All` | Application | 通話記録の読み取り |
| `Calls.InitiateGroupCall.All` | Application | グループ通話の発信 |
| `Calls.JoinGroupCall.All` | Application | グループ通話・会議への参加 |
| `Chat.ReadWrite.All` | Application | チャットメッセージの読み書き |
| `ChatMessage.Read.All` | Application | チャットメッセージの購読・取得 |
| `OnlineMeetings.Read.All` | Application | オンライン会議の詳細読み取り |
| `TeamsAppInstallation.ReadWriteForChat.All` | Application | 会議チャットへのアプリインストール |

> **注意**: `Teamwork.Migrate.All` はテナント管理者の承認が難しい高権限のため使用しない。  
> チャットへのメッセージ送信はタブ UI で代替。

### 2. 環境変数の設定

```bash
cd meeting-app
cp .env.example .env
```

`.env` を編集：

```env
MICROSOFT_APP_ID=<アプリケーション（クライアント）ID>
MICROSOFT_APP_SECRET=<クライアントシークレット>
TENANT_ID=<ディレクトリ（テナント）ID>
NOTIFICATION_URL=https://<ngrok または本番 URL>
AZURE_WEBHOOK_URL=<録画停止通知先の Azure Webhook URL>
PORT=3980
TEAMS_APP_ID=<Teams 管理センターのアプリカタログ ID>
```

> `TEAMS_APP_ID` は管理センター「Teams アプリ」→「アプリの管理」でアプリを選択したときの **アプリ ID**（manifest.json の `id` とは別）。

### 3. アプリパッケージの作成

```bash
cd meeting-app/manifest
zip ../recording-monitor.zip manifest.json color.png outline.png
```

### 4. Teams 管理センターへのアップロード

1. [admin.teams.microsoft.com](https://admin.teams.microsoft.com) にアクセス
2. 「Teams アプリ」→「アプリの管理」→「アップロード」
3. `recording-monitor.zip` をアップロード
4. アップロード後に表示される **アプリ ID** を `.env` の `TEAMS_APP_ID` に設定

### 5. App Setup Policy の設定

1. 「Teams アプリ」→「セットアップ ポリシー」→「追加」
2. ポリシー名を入力（例: `RecordingMonitorPolicy`）
3. 「インストール済みアプリ」→「アプリの追加」→「Recording Monitor」を追加
4. 「会議拡張機能」→「アプリの追加」→「Recording Monitor」を追加
5. 保存後、「グループ ポリシーの割り当て」→「グループの追加」で対象グループを指定

> ポリシーの反映には最大 24 時間かかる場合があります。

### 6. ローカル起動

```bash
# ngrok を起動（別ターミナル）
ngrok http 3980

# .env の NOTIFICATION_URL を ngrok URL に更新してから起動
cd meeting-app
uv run python app.py
```

起動すると `/chats` 購読が自動作成され、以降の会議チャット作成を自動検知します。

---

## マニフェストの注意点

[manifest/manifest.json](manifest/manifest.json) の URL は ngrok ドメインに合わせて更新が必要：

```json
"contentUrl": "https://<NGROK_DOMAIN>/tab",
"websiteUrl": "https://<NGROK_DOMAIN>/tab",
"validDomains": ["<NGROK_DOMAIN>"]
```

`id` フィールド（`3f8a7c2d-...`）は Teams アプリの一意な ID。本番利用時は新しい GUID に変更：

```python
import uuid; print(uuid.uuid4())
```

---

## CsApplicationAccessPolicy について

**現在の実装では不要です。**

Bot が Graph Communications API 経由で会議に通話として参加する方式（旧実装）で必要だったコマンドです。現在の実装は会議に参加せず Graph 変更通知を受け取るだけなので、このポリシー設定は不要です。

参考として以下に記載します：

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

---

## 設計上の制約と判断

| 制約 | 対応 |
|---|---|
| アプリ権限でのチャットメッセージ送信には `Teamwork.Migrate.All` が必要 | チャット送信を廃止し、タブ内 UI で代替 |
| `/chats` 購読はアプリあたり 1 件まで | 起動時に既存の `/chats` 購読を削除してから再作成 |
| Bot の `installationUpdate` は App Setup Policy では発火しない | `/chats` 購読 + タブの `getContext()` で threadId を取得 |
