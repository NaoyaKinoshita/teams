import asyncio
import sys
import traceback
from http import HTTPStatus

from aiohttp import web
from aiohttp.web import Request, Response, json_response
from botbuilder.core import BotFrameworkAdapter, BotFrameworkAdapterSettings, TurnContext
from botbuilder.schema import Activity

from bot import TeamsRecordingBot
from config import DefaultConfig
from graph_client import (
    create_call_record_subscription,
    create_calendar_subscription,
    handle_recording_notification,
    handle_call_notification,
)

CONFIG = DefaultConfig()
SETTINGS = BotFrameworkAdapterSettings(
    app_id=CONFIG.APP_ID,
    app_password=CONFIG.APP_SECRET,
    channel_auth_tenant=CONFIG.TENANT_ID,
)
ADAPTER = BotFrameworkAdapter(SETTINGS)
BOT = TeamsRecordingBot(CONFIG.APP_ID)


async def on_error(context: TurnContext, error: Exception):
    print(f"\n[on_turn_error] 予期しないエラー: {error}", file=sys.stderr)
    traceback.print_exc()
    await context.send_activity("エラーが発生しました。管理者に連絡してください。")


ADAPTER.on_turn_error = on_error


async def messages(req: Request) -> Response:
    """Bot Framework からのメッセージを受け取るエンドポイント"""
    if "application/json" not in req.headers.get("Content-Type", ""):
        return Response(status=HTTPStatus.UNSUPPORTED_MEDIA_TYPE)

    body = await req.json()
    activity = Activity().deserialize(body)
    auth_header = req.headers.get("Authorization", "")

    invoke_response = await ADAPTER.process_activity(activity, auth_header, BOT.on_turn)
    if invoke_response:
        return json_response(data=invoke_response.body, status=invoke_response.status)
    return Response(status=HTTPStatus.OK)


async def notifications(req: Request) -> Response:
    """Graph API からの変更通知を受け取るエンドポイント（callRecords・カレンダー）"""
    print(f"[notifications] {req.method} {req.rel_url}")
    try:
        validation_token = req.rel_url.query.get("validationToken")
        if validation_token:
            print(f"[notifications] 検証トークンを返します: {validation_token[:20]}...")
            return Response(text=validation_token, content_type="text/plain", status=HTTPStatus.OK)

        body = await req.json()
        await handle_recording_notification(body, BOT, ADAPTER)
        return Response(status=HTTPStatus.ACCEPTED)
    except Exception as e:
        print(f"[notifications] エラー: {e}", file=sys.stderr)
        traceback.print_exc()
        return Response(status=HTTPStatus.INTERNAL_SERVER_ERROR)


async def calls(req: Request) -> Response:
    """Graph Communications API からの通話状態変更通知を受け取るエンドポイント"""
    print(f"[calls] {req.method} {req.rel_url}")
    try:
        validation_token = req.rel_url.query.get("validationToken")
        if validation_token:
            return Response(text=validation_token, content_type="text/plain", status=HTTPStatus.OK)

        body = await req.json()
        print(f"[calls] 受信: {body}")
        await handle_call_notification(body)
        return Response(status=HTTPStatus.ACCEPTED)
    except Exception as e:
        print(f"[calls] エラー: {e}", file=sys.stderr)
        traceback.print_exc()
        return Response(status=HTTPStatus.INTERNAL_SERVER_ERROR)


async def _setup_subscriptions():
    """サーバー起動後にサブスクリプションを登録する"""
    await asyncio.sleep(2)
    try:
        result = await create_call_record_subscription()
        print(f"[startup] callRecords サブスクリプション: {result.get('id')}")
    except Exception as e:
        print(f"[startup] callRecords サブスクリプション失敗: {e}")

    if CONFIG.RECORDING_BOT_USER_ID:
        try:
            result = await create_calendar_subscription()
            print(f"[startup] カレンダーサブスクリプション: {result.get('id')}")
        except Exception as e:
            print(f"[startup] カレンダーサブスクリプション失敗: {e}")


async def on_startup(app: web.Application):
    if not CONFIG.NOTIFICATION_URL:
        print("[startup] NOTIFICATION_URL が未設定のため、サブスクリプションをスキップします。")
        return
    asyncio.create_task(_setup_subscriptions())


APP = web.Application()
APP.router.add_post("/api/messages", messages)
APP.router.add_post("/api/notifications", notifications)
APP.router.add_post("/api/calls", calls)
APP.on_startup.append(on_startup)

if __name__ == "__main__":
    web.run_app(APP, host="localhost", port=CONFIG.PORT)
