import json
import logging

import azure.functions as func

app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)


@app.route(route="recording-stopped", methods=["POST"])
async def recording_stopped(req: func.HttpRequest) -> func.HttpResponse:
    """録画停止通知を受け取るエンドポイント

    Bot から以下の payload が POST される:
    {
        "event": "recording_stopped",
        "callId": "xxx",
        "threadId": "19:xxx",
        "timestamp": "2026-04-15T00:36:03Z"
    }
    """
    logging.info("[recording-stopped] 通知受信")

    try:
        body = req.get_json()
    except ValueError:
        return func.HttpResponse(
            json.dumps({"error": "Invalid JSON"}),
            mimetype="application/json",
            status_code=400,
        )

    call_id = body.get("callId", "")
    thread_id = body.get("threadId", "")
    timestamp = body.get("timestamp", "")

    logging.info(f"[recording-stopped] callId={call_id} threadId={thread_id} timestamp={timestamp}")

    # TODO: 文字起こし処理などの後続処理をここに実装
    # 例1: Azure Storage Queue にメッセージを追加して非同期処理
    # 例2: Azure AI Speech Service で録画ファイルを文字起こし
    # 例3: Azure Service Bus にイベントを送信

    return func.HttpResponse(
        json.dumps({"status": "received", "callId": call_id}),
        mimetype="application/json",
        status_code=200,
    )
