import logging

from fastapi import FastAPI, Request, Response

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI()


@app.post("/api/recording-stopped")
async def recording_stopped(req: Request):
    """録画の OneDrive 保存完了通知を受け取るエンドポイント

    meeting-app から以下の payload が POST される:
    {
        "event": "recording_saved",
        "threadId": "19:meeting_xxx@thread.v2",
        "recordingUrl": "https://xxx.sharepoint.com/...",
        "timestamp": "2026-04-16T00:00:00+00:00"
    }
    """
    try:
        body = await req.json()
    except Exception:
        return Response(content='{"error": "Invalid JSON"}', media_type="application/json", status_code=400)

    thread_id = body.get("threadId", "")
    recording_url = body.get("recordingUrl", "")
    timestamp = body.get("timestamp", "")

    logger.info(f"[recording-stopped] threadId={thread_id} recordingUrl={recording_url} timestamp={timestamp}")

    # TODO: 録画 URL を使った後続処理をここに実装
    # 例1: Azure AI Speech Service で録画ファイルを文字起こし
    # 例2: Azure Storage Queue にメッセージを追加して非同期処理
    # 例3: Azure Service Bus にイベントを送信

    return {"status": "received", "threadId": thread_id}
