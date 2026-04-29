import asyncio
import sys
import traceback
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI, Request, Response
from fastapi.responses import HTMLResponse, JSONResponse, PlainTextResponse
from botbuilder.core import (
    BotFrameworkAdapter,
    BotFrameworkAdapterSettings,
    TurnContext,
)
from botbuilder.schema import Activity

from bot import MeetingRecordingBot
from config import DefaultConfig
from graph_client import (
    consent_azure_integration,
    create_chats_subscription,
    get_recording_status,
    handle_recording_notification,
    setup_meeting_chat,
    subscription_renewal_loop,
)

TEMPLATES_DIR = Path(__file__).parent / "templates"

CONFIG = DefaultConfig()
SETTINGS = BotFrameworkAdapterSettings(
    app_id=CONFIG.APP_ID,
    app_password=CONFIG.APP_SECRET,
    channel_auth_tenant=CONFIG.TENANT_ID,
)
ADAPTER = BotFrameworkAdapter(SETTINGS)
BOT = MeetingRecordingBot()


async def on_error(context: TurnContext, error: Exception):
    print(f"\n[on_turn_error] 予期しないエラー: {error}", file=sys.stderr)
    traceback.print_exc()


ADAPTER.on_turn_error = on_error


async def _setup_subscriptions():
    await asyncio.sleep(2)
    try:
        result = await create_chats_subscription()
        print(f"[startup] /chats 購読: {result.get('id', result)}")
    except Exception as e:
        print(f"[startup] /chats 購読失敗: {e}")


@asynccontextmanager
async def lifespan(app: FastAPI):
    if CONFIG.NOTIFICATION_URL:
        asyncio.create_task(_setup_subscriptions())
        asyncio.create_task(subscription_renewal_loop())
    else:
        print("[startup] NOTIFICATION_URL 未設定のためスキップ")
    yield


router = FastAPI(lifespan=lifespan)


@router.post("/api/messages")
async def messages(req: Request):
    """Bot Framework からのメッセージを受け取るエンドポイント"""
    if "application/json" not in req.headers.get("content-type", ""):
        return Response(status_code=415)

    body = await req.json()
    activity = Activity().deserialize(body)
    auth_header = req.headers.get("Authorization", "")

    print(
        f"[messages] type={activity.type} action={getattr(activity, 'action', None)} "
        f"convType={activity.conversation.conversation_type if activity.conversation else None} "
        f"convId={activity.conversation.id[:40] if activity.conversation else None}"
    )

    invoke_response = await ADAPTER.process_activity(activity, auth_header, BOT.on_turn)
    if invoke_response:
        return JSONResponse(
            content=invoke_response.body, status_code=invoke_response.status
        )
    return Response(status_code=200)


@router.post("/api/notifications")
async def notifications(req: Request):
    """Graph API からの変更通知を受け取るエンドポイント"""
    print(f"[notifications] POST {req.url}")
    try:
        validation_token = req.query_params.get("validationToken")
        if validation_token:
            print(f"[notifications] 検証トークンを返します: {validation_token[:20]}...")
            return PlainTextResponse(content=validation_token, status_code=200)

        body = await req.json()
        asyncio.create_task(handle_recording_notification(body))
        return Response(status_code=202)
    except Exception as e:
        print(f"[notifications] エラー: {e}", file=sys.stderr)
        traceback.print_exc()
        return Response(status_code=500)




@router.get("/tab", response_class=HTMLResponse)
async def tab():
    """会議サイドパネルタブのコンテンツ"""
    return (TEMPLATES_DIR / "tab.html").read_text(encoding="utf-8")


@router.get("/notification", response_class=HTMLResponse)
async def notification():
    """コンテンツバブル（会議バナー）から開かれる通知 UI"""
    return (TEMPLATES_DIR / "notification.html").read_text(encoding="utf-8")


@router.post("/api/tab-context")
async def tab_context(req: Request):
    """タブから threadId を受け取りチャット購読を作成する"""
    try:
        body = await req.json()
        thread_id = body.get("threadId", "")
        if thread_id:
            print(f"[Tab] threadId 受信: {thread_id}")
            asyncio.create_task(setup_meeting_chat(thread_id))
        return Response(status_code=200)
    except Exception as e:
        print(f"[Tab] エラー: {e}")
        return Response(status_code=500)


@router.get("/api/recording-status")
async def recording_status(threadId: str = ""):
    """タブ向けに録画状態・同意状態を返す"""
    if not threadId:
        return JSONResponse({"error": "threadId required"}, status_code=400)
    return JSONResponse(get_recording_status(threadId))


@router.post("/api/consent")
async def consent(req: Request):
    """タブからの Azure 連携同意/スキップを受け取る"""
    try:
        body = await req.json()
        thread_id = body.get("threadId", "")
        agreed = body.get("agreed", False)
        if not thread_id:
            return JSONResponse({"error": "threadId required"}, status_code=400)
        if agreed:
            consent_azure_integration(thread_id)
            print(f"[Consent] Azure 連携に同意: {thread_id}")
        else:
            print(f"[Consent] Azure 連携をスキップ: {thread_id}")
        return Response(status_code=200)
    except Exception as e:
        print(f"[Consent] エラー: {e}")
        return Response(status_code=500)


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("app:router", host="0.0.0.0", port=CONFIG.PORT, reload=False)
