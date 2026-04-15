# Compliance Recording Bot

コンプライアンス録画を利用して Teams 会議に自動参加し、録画の開始・停止を検知して Azure に通知する Bot。

## bot/ との違い

| | bot/ | compliance-bot/ |
|---|---|---|
| 前提ライセンス | Microsoft 365 Business Basic | **Microsoft 365 E3 / E5** |
| 会議参加の仕組み | recording-bot ユーザーを招待 → カレンダー通知 | ポリシーで自動参加 |
| recording-bot ユーザー（$6/月） | 必要 | **不要** |
| カレンダー購読・スケジュール処理 | 必要 | **不要** |
| チャットメッセージ購読 | 必要 | **不要** |
| `recordingStatus` 通知 | 来ない | **call 通知で直接届く** |
| Teams アプリのアップロード | 必要 | **不要** |

---

## 前提条件

- Microsoft 365 E3 / E5（会議参加者全員分のライセンスが必要）
- Azure Bot Service（既存の `teams-recording-bot` を流用可）
- Entra ID アプリ登録（既存の `c12e5c31-...` を流用可）

---

## セットアップ手順

### 1. API アクセス許可の確認

Entra ID のアプリ登録に以下の権限が付与されていること（管理者同意済み）：

| 権限 | 用途 |
|------|------|
| `CallRecords.Read.All` | 通話記録の取得（録画 URL 検索） |
| `Calls.AccessMedia.All` | 会議メディアへのアクセス |
| `Calls.JoinGroupCalls.All` | 会議への参加 |
| `Chat.ReadWrite.All` | 会議チャットへのメッセージ送信 |

### 2. Azure Bot Service に Calling Webhook を設定

[portal.azure.com](https://portal.azure.com) → Azure Bot（`teams-recording-bot`）→「構成」→「通話 Webhook」

```
https://<ngrok または本番 URL>/api/calls
```

### 3. アプリケーションインスタンスの作成

コンプライアンス録画ポリシーは Entra ID の App ID ではなく、Teams 上の**アプリケーションインスタンスの ObjectId** を要求する。

> **アプリケーションインスタンスとは**
> Teams のポリシーシステムがアプリを「参加者」として扱うための Teams 専用オブジェクト。
> 通常ユーザーに似た UPN を持つが、ライセンス不要でログイン不可の特殊なアカウント種別。

```powershell
Connect-MicrosoftTeams

# アプリケーションインスタンスを作成
$appInstance = New-CsOnlineApplicationInstance `
  -UserPrincipalName "recording-bot-app@<テナント>.onmicrosoft.com" `
  -DisplayName "Recording Bot App" `
  -ApplicationId "c12e5c31-696d-454f-b246-287b96b06632"

# 確認（ObjectId をメモ）
$appInstance

# 同期
Sync-CsOnlineApplicationInstance `
  -ObjectId $appInstance.ObjectId `
  -ApplicationId "c12e5c31-696d-454f-b246-287b96b06632"
```

### 4. コンプライアンス録画ポリシーの設定

```powershell
# ポリシーを作成
New-CsTeamsComplianceRecordingPolicy `
  -Identity "ComplianceRecordingPolicy" `
  -Enabled $true

# ポリシーにアプリを追加
Set-CsTeamsComplianceRecordingPolicy `
  -Identity "ComplianceRecordingPolicy" `
  -ComplianceRecordingApplications @(
    New-CsTeamsComplianceRecordingApplication `
      -Id $appInstance.ObjectId `
      -Parent "ComplianceRecordingPolicy"
  )

# テナント全体に適用
Grant-CsTeamsComplianceRecordingPolicy `
  -PolicyName "ComplianceRecordingPolicy" `
  -Global
```

> **注意**: ポリシーの反映には 1〜2 時間かかる場合がある。

### 5. 動作確認

```powershell
# 作成済みインスタンスの確認
Get-CsOnlineApplicationInstance

# ポリシーの確認
Get-CsTeamsComplianceRecordingPolicy -Identity "ComplianceRecordingPolicy"
```

---

## 環境変数

```env
MICROSOFT_APP_ID=<アプリケーション（クライアント）ID>
MICROSOFT_APP_SECRET=<クライアントシークレット>
TENANT_ID=<ディレクトリ（テナント）ID>
NOTIFICATION_URL=https://<ngrok または本番 URL>
AZURE_WEBHOOK_URL=<録画停止通知先の Azure Webhook URL>
PORT=3979
```

---

## ローカル起動

```bash
cd compliance-bot
uv run python app.py
```

起動後、`callRecords` サブスクリプションが自動登録される。

---

## 通知フロー

```
対象ユーザーが会議に参加
    ↓（CsTeamsComplianceRecordingPolicy により自動）
Teams が Bot を会議に招待
    ↓
/api/calls コールバック
    ├── state: established   → 会議参加確立
    ├── recordingStatus: recording    → 「録画が開始されました。」をチャットに送信
    ├── recordingStatus: notRecording → 「録画が停止されました。」+ Azure Webhook 通知
    └── state: terminated    → 「会議が終了しました。」+ 録画 URL ポーリング開始

callRecords 通知 → /api/notifications
    → 録画 URL が取得できたらチャットに通知
```

---

## bot/ との使い分け

| 環境 | 推奨方式 |
|------|---------|
| 技術検証・小規模 | `bot/`（Business Basic + recording-bot ユーザー） |
| 本番・全社展開（E3/E5 環境） | `compliance-bot/`（ポリシーで自動参加） |
