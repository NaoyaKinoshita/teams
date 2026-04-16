import asyncio
import sys
import traceback
from contextlib import asynccontextmanager

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
    handle_app_installed,
)

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


TAB_HTML = """<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Recording Monitor</title>
  <script src="https://res.cdn.office.net/teams-js/2.22.0/js/MicrosoftTeams.min.js"></script>
  <style>
    * { box-sizing: border-box; }
    body { font-family: "Segoe UI", sans-serif; margin: 0; background: #f3f2f1;
           display: flex; align-items: flex-start; justify-content: center; padding: 24px 16px; }
    .container { width: 100%; max-width: 360px; text-align: center; color: #323130; }
    .icon { font-size: 40px; margin-bottom: 12px; }
    h1 { font-size: 18px; margin: 0 0 6px; }
    .desc { font-size: 13px; color: #605e5c; margin: 0 0 20px; }
    .status-badge { display: inline-block; padding: 4px 12px; border-radius: 12px;
                    font-size: 12px; font-weight: 600; margin-bottom: 20px; }
    .status-idle    { background: #edebe9; color: #605e5c; }
    .status-rec     { background: #fde7e9; color: #a4262c; }
    .status-ok      { background: #dff6dd; color: #107c10; }
    .card { background: #fff; border-radius: 8px; padding: 16px;
            box-shadow: 0 1px 4px rgba(0,0,0,.1); margin-bottom: 16px; text-align: left; }
    .card-title { font-size: 14px; font-weight: 600; margin: 0 0 8px; }
    .card-body  { font-size: 13px; color: #605e5c; margin: 0 0 14px; }
    .btn { display: block; width: 100%; padding: 10px; border: none; border-radius: 4px;
           font-size: 14px; font-weight: 600; cursor: pointer; margin-bottom: 8px; }
    .btn-primary { background: #0078d4; color: #fff; }
    .btn-primary:hover { background: #106ebe; }
    .btn-secondary { background: #edebe9; color: #323130; }
    .btn-secondary:hover { background: #e1dfdd; }
    .url-box { font-size: 11px; word-break: break-all; background: #f3f2f1;
               padding: 8px; border-radius: 4px; color: #0078d4; }
    #footer { font-size: 11px; color: #a19f9d; margin-top: 8px; }
  </style>
</head>
<body>
  <div class="container">
    <div class="icon">&#128280;</div>
    <h1>Recording Monitor</h1>
    <p class="desc">録画の開始・停止を検知して Azure に通知します。</p>
    <div id="badge" class="status-badge status-idle">待機中</div>
    <div id="card-area"></div>
    <div id="footer">初期化中...</div>
  </div>

  <script>
    let threadId = "";
    let lastRecording = null;
    let consented = false;
    let consentSent = false;

    function setFooter(msg) {
      document.getElementById("footer").textContent = msg;
    }

    function setBadge(state) {
      const el = document.getElementById("badge");
      el.className = "status-badge";
      if (state === "recording") {
        el.classList.add("status-rec");
        el.textContent = "録画中";
      } else if (state === "ok") {
        el.classList.add("status-ok");
        el.textContent = "Azure 連携済み";
      } else {
        el.classList.add("status-idle");
        el.textContent = "待機中";
      }
    }

    function renderCard(data) {
      const area = document.getElementById("card-area");

      if (!data.recording && !data.recordingUrl) {
        area.innerHTML = "";
        setBadge("idle");
        return;
      }

      if (data.recording && !consentSent) {
        setBadge("recording");
        area.innerHTML = `
          <div class="card">
            <div class="card-title">&#128250; 録画が開始されました</div>
            <div class="card-body">録画データを Azure に連携しますか？<br>OneDrive 保存後に Webhook へ通知します。</div>
            <button class="btn btn-primary" id="btn-ok">Azure に連携する</button>
            <button class="btn btn-secondary" id="btn-skip">スキップ</button>
          </div>`;
        document.getElementById("btn-ok").onclick = () => sendConsent(true);
        document.getElementById("btn-skip").onclick = () => sendConsent(false);
        return;
      }

      if (data.recording && consentSent) {
        setBadge(consented ? "ok" : "recording");
        area.innerHTML = `<div class="card">
          <div class="card-body">${consented
            ? "&#10003; Azure 連携に同意済みです。OneDrive 保存後に通知します。"
            : "スキップを選択しました。Azure には通知しません。"}</div>
        </div>`;
        return;
      }

      if (data.recordingUrl) {
        setBadge("idle");
        area.innerHTML = `<div class="card">
          <div class="card-title">&#9989; 録画が OneDrive に保存されました</div>
          <div class="url-box"><a href="${data.recordingUrl}" target="_blank">${data.recordingUrl}</a></div>
        </div>`;
        return;
      }
    }

    function sendConsent(agreed) {
      consentSent = true;
      consented = agreed;
      fetch("/api/consent", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ threadId, agreed })
      });
      renderCard({ recording: true, recordingUrl: "" });
    }

    function poll() {
      if (!threadId) return;
      fetch("/api/recording-status?threadId=" + encodeURIComponent(threadId))
        .then(r => r.json())
        .then(data => {
          if (data.recording && !lastRecording) {
            consentSent = false;
            consented = false;
          }
          lastRecording = data.recording;
          renderCard(data);
        })
        .catch(() => {});
    }

    microsoftTeams.app.initialize().then(() => {
      microsoftTeams.app.getContext().then((context) => {
        threadId = context.chat?.id || context.meeting?.id || "";
        setFooter(threadId ? "監視中" : "threadId 取得失敗");
        if (threadId) {
          fetch("/api/tab-context", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ threadId })
          });
          setInterval(poll, 2000);
          poll();
        }
      });
    });
  </script>
</body>
</html>"""


@router.get("/tab", response_class=HTMLResponse)
async def tab():
    """会議サイドパネルタブのコンテンツ"""
    return TAB_HTML


@router.post("/api/tab-context")
async def tab_context(req: Request):
    """タブから threadId を受け取りチャット購読を作成する"""
    try:
        body = await req.json()
        thread_id = body.get("threadId", "")
        if thread_id:
            print(f"[Tab] threadId 受信: {thread_id}")
            asyncio.create_task(handle_app_installed(thread_id))
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
