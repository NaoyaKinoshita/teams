# recording-functions

meeting-app から録画停止通知を受け取り、SharePoint の録画ファイルを Azure Blob Storage に保存する FastAPI サーバー。

## 役割

meeting-app が録画の OneDrive 保存完了（`callRecordingStatus: chunkFinished`）を検知した際に Webhook で通知を受け取り、以下を実行する：

1. Graph API で SharePoint 共有 URL から `@microsoft.graph.downloadUrl`（事前認証済み URL）を取得
2. Azure Blob Storage へ直接ストリームアップロード（`upload_blob_from_url`）

Blob パスは `{sanitized_threadId}/{filename}` 形式で保存される。

---

## エンドポイント

| メソッド | パス | 説明 |
|---|---|---|
| POST | `/api/recording-stopped` | 録画停止通知の受け取り・Blob 保存トリガー |

### リクエスト payload

```json
{
  "event": "recording_saved",
  "threadId": "19:meeting_xxx@thread.v2",
  "recordingUrl": "https://xxx.sharepoint.com/...",
  "timestamp": "2026-04-16T00:00:00+00:00"
}
```

---

## 環境変数

| 変数名 | 説明 |
|---|---|
| `MICROSOFT_APP_ID` | Entra ID アプリ ID（Graph API トークン取得用） |
| `MICROSOFT_APP_SECRET` | クライアントシークレット |
| `TENANT_ID` | ディレクトリ（テナント）ID |
| `AZURE_STORAGE_CONNECTION_STRING` | Blob Storage の接続文字列 |
| `STORAGE_CONTAINER_NAME` | 保存先コンテナ名（デフォルト: `recordings`） |

> `MICROSOFT_APP_ID` / `MICROSOFT_APP_SECRET` / `TENANT_ID` は meeting-app と同じアプリ登録を使用。
> そのアプリに `Files.Read.All`（Application）権限が必要。

---

## デプロイ

Azure Container Apps（`recording-functions`）にデプロイ済み。

```bash
# イメージのビルド＆プッシュ
docker buildx build --platform linux/amd64 \
  -t teamsrecodingbot.azurecr.io/recording-functions:latest \
  --push .

# Container App の更新
az containerapp update \
  --name recording-functions \
  --resource-group teams-recording-rg \
  --image teamsrecodingbot.azurecr.io/recording-functions:latest
```

---

## ローカル起動

```bash
uv run uvicorn main:app --reload --port 8000
```
