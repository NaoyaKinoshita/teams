# recording-functions

meeting-app から録画 OneDrive 保存完了通知を受け取る FastAPI サーバー。

## 役割

meeting-app が録画の OneDrive 保存完了（`callRecordingStatus: success`）を検知した際に、
Azure 連携に同意済みのチャットに対して POST 通知を受け取る。

現時点では受け口のみ実装済み。録画 URL を使った後続処理（文字起こし等）はここに追加する。

## エンドポイント

| メソッド | パス | 説明 |
|---|---|---|
| POST | `/api/recording-stopped` | 録画保存完了通知の受け取り |

### リクエスト payload

```json
{
  "event": "recording_saved",
  "threadId": "19:meeting_xxx@thread.v2",
  "recordingUrl": "https://xxx.sharepoint.com/...",
  "timestamp": "2026-04-16T00:00:00+00:00"
}
```

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

## ローカル起動

```bash
uv run uvicorn main:app --reload --port 8000
```
