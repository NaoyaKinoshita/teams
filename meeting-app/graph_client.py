import asyncio
import re
from datetime import datetime, timedelta, timezone
from urllib.parse import quote

import aiohttp

from config import DefaultConfig

CONFIG = DefaultConfig()

# 会議チャットのサブスクリプション管理 {thread_id: subscription_id}
_chat_subscriptions: dict = {}

# /chats 購読の ID
_chats_subscription_id: str = ""

# 録画中の thread_id
_recording_active: set = set()

# Azure 連携に同意済みの thread_id
_integrate_consented: set = set()

# 録画 URL 通知済みの thread_id
_notified_recordings: set = set()

# 録画 URL 保存 {thread_id: url}
_recording_urls: dict = {}

# Bot Framework のコンテンツバブル送信用コンテキスト {thread_id: {"meeting_id": ..., "service_url": ...}}
_meeting_contexts: dict = {}

# コンテンツバブル送信済みの thread_id（録画毎に1回だけ送信するため）
_bubble_sent: set = set()


def get_recording_status(thread_id: str) -> dict:
    return {
        "recording": thread_id in _recording_active,
        "consented": thread_id in _integrate_consented,
        "recordingUrl": _recording_urls.get(thread_id, ""),
    }


def consent_azure_integration(thread_id: str):
    _integrate_consented.add(thread_id)
    print(f"[Graph] Azure 連携同意: {thread_id}")


async def get_access_token() -> str:
    url = f"https://login.microsoftonline.com/{CONFIG.TENANT_ID}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": CONFIG.APP_ID,
        "client_secret": CONFIG.APP_SECRET,
        "scope": "https://graph.microsoft.com/.default",
    }
    async with aiohttp.ClientSession() as session:
        async with session.post(url, data=data) as resp:
            result = await resp.json()
            if "access_token" not in result:
                raise RuntimeError(f"トークン取得失敗: {result}")
            return result["access_token"]


async def get_bot_framework_token() -> str:
    url = f"https://login.microsoftonline.com/{CONFIG.TENANT_ID}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": CONFIG.APP_ID,
        "client_secret": CONFIG.APP_SECRET,
        "scope": "https://api.botframework.com/.default",
    }
    async with aiohttp.ClientSession() as session:
        async with session.post(url, data=data) as resp:
            result = await resp.json()
            if "access_token" not in result:
                raise RuntimeError(f"Bot Framework トークン取得失敗: {result}")
            return result["access_token"]


async def _install_app_in_chat(thread_id: str):
    print(f"[Graph] アプリインストール試行: threadId={thread_id} appId={CONFIG.TEAMS_APP_ID}")
    token = await get_access_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    body = {
        "teamsApp@odata.bind": (
            f"https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/{CONFIG.TEAMS_APP_ID}"
        )
    }
    async with aiohttp.ClientSession() as session:
        async with session.post(
            f"https://graph.microsoft.com/v1.0/chats/{thread_id}/installedApps",
            headers=headers,
            json=body,
        ) as resp:
            if resp.status in (200, 201):
                print(f"[Graph] アプリインストール完了: {thread_id}")
            elif resp.status == 409:
                print(f"[Graph] アプリは既にインストール済み（問題なし）: {thread_id}")
            else:
                result = await resp.json()
                print(f"[Graph] アプリインストール失敗 status={resp.status}: {result}")


async def setup_meeting_chat(thread_id: str):
    """会議チャットへのアプリインストールとメッセージ購読を行う"""
    if thread_id in _chat_subscriptions:
        print(f"[Graph] 既に購読中のためスキップ: {thread_id}")
        return

    await _install_app_in_chat(thread_id)

    try:
        sub = await _create_chat_message_subscription(thread_id)
        sub_id = sub.get("id")
        if sub_id:
            _chat_subscriptions[thread_id] = sub_id
            print(f"[Graph] チャット購読作成: threadId={thread_id} subId={sub_id}")
        else:
            print(f"[Graph] チャット購読作成失敗: {sub}")
    except Exception as e:
        print(f"[Graph] チャット購読作成エラー: {e}")


async def cleanup_old_chats_subscriptions():
    """既存の /chats 購読を全て削除する（起動時の上限リセット用）"""
    token = await get_access_token()
    headers = {"Authorization": f"Bearer {token}"}
    async with aiohttp.ClientSession() as session:
        async with session.get(
            "https://graph.microsoft.com/v1.0/subscriptions",
            headers=headers,
        ) as resp:
            result = await resp.json()
            for sub in result.get("value", []):
                if sub.get("resource") == "/chats":
                    await _delete_subscription(sub["id"])
                    print(f"[Graph] 古い /chats 購読を削除: {sub['id']}")


async def create_chats_subscription() -> dict:
    """全チャットの作成通知を購読する（会議チャット自動検知用）"""
    global _chats_subscription_id
    await cleanup_old_chats_subscriptions()
    token = await get_access_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    expiration = datetime.now(timezone.utc) + timedelta(hours=1)
    body = {
        "changeType": "created",
        "notificationUrl": f"{CONFIG.NOTIFICATION_URL}/api/notifications",
        "resource": "/chats",
        "expirationDateTime": expiration.isoformat(),
        "clientState": "teams-meeting-app-chats",
    }
    async with aiohttp.ClientSession() as session:
        async with session.post(
            "https://graph.microsoft.com/v1.0/subscriptions",
            headers=headers,
            json=body,
        ) as resp:
            result = await resp.json()
            _chats_subscription_id = result.get("id", "")
            print(f"[Graph] /chats 購読: {_chats_subscription_id}")
            return result


async def _renew_subscription(sub_id: str) -> bool:
    """購読の有効期限を1時間延長する"""
    try:
        token = await get_access_token()
        headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
        expiration = datetime.now(timezone.utc) + timedelta(hours=1)
        body = {"expirationDateTime": expiration.isoformat()}
        async with aiohttp.ClientSession() as session:
            async with session.patch(
                f"https://graph.microsoft.com/v1.0/subscriptions/{sub_id}",
                headers=headers,
                json=body,
            ) as resp:
                if resp.status == 200:
                    print(f"[Graph] 購読更新: subId={sub_id}")
                    return True
                else:
                    result = await resp.json()
                    print(f"[Graph] 購読更新失敗 status={resp.status}: {result}")
                    return False
    except Exception as e:
        print(f"[Graph] 購読更新エラー: {e}")
        return False


async def subscription_renewal_loop():
    """50分ごとに全購読を更新するバックグラウンドタスク"""
    while True:
        await asyncio.sleep(50 * 60)
        print("[Graph] 購読更新タスク起動")

        if _chats_subscription_id:
            ok = await _renew_subscription(_chats_subscription_id)
            if not ok:
                print("[Graph] /chats 購読を再作成します")
                try:
                    await create_chats_subscription()
                except Exception as e:
                    print(f"[Graph] /chats 購読再作成エラー: {e}")

        for thread_id, sub_id in list(_chat_subscriptions.items()):
            ok = await _renew_subscription(sub_id)
            if not ok:
                print(f"[Graph] チャット購読を再作成します: {thread_id}")
                _chat_subscriptions.pop(thread_id, None)
                try:
                    await setup_meeting_chat(thread_id)
                except Exception as e:
                    print(f"[Graph] チャット購読再作成エラー: {e}")


def store_meeting_context(thread_id: str, meeting_id: str, service_url: str):
    """Bot Framework activity から取得した会議コンテキストを保存する"""
    _meeting_contexts[thread_id] = {
        "meeting_id": meeting_id,
        "service_url": service_url,
    }
    print(f"[Bot] 会議コンテキスト保存: threadId={thread_id} meetingId={meeting_id[:30]}... serviceUrl={service_url}")


async def get_organizer_ids(thread_id: str) -> list:
    """会議主催者の AAD Object ID を取得する（onlineMeetingInfo.organizer を使用）"""
    token = await get_access_token()
    headers = {"Authorization": f"Bearer {token}"}
    async with aiohttp.ClientSession() as session:
        async with session.get(
            f"https://graph.microsoft.com/v1.0/chats/{thread_id}?$select=onlineMeetingInfo",
            headers=headers,
        ) as resp:
            result = await resp.json()
            organizer_id = (
                result.get("onlineMeetingInfo", {})
                .get("organizer", {})
                .get("id")
            )
            if organizer_id:
                print(f"[Bubble] 主催者: {organizer_id}")
                return [organizer_id]
            print(f"[Bubble] 主催者取得失敗、フォールバック: {result}")
            return []


async def send_content_bubble(thread_id: str):
    """録画開始時に会議のコンテンツバブル（バナー）を送信する"""
    if thread_id in _bubble_sent:
        print(f"[Bubble] 既に送信済みのためスキップ: {thread_id}")
        return

    context = _meeting_contexts.get(thread_id)
    if not context:
        print(f"[Bubble] 会議コンテキスト未取得のためスキップ: {thread_id}")
        return

    meeting_id = context["meeting_id"]
    service_url = context["service_url"]

    try:
        organizer_aad_ids = await get_organizer_ids(thread_id)
        print(f"[Bubble] 取得した organizer_aad_ids: {organizer_aad_ids}")
    except Exception as e:
        print(f"[Bubble] メンバー ID 取得エラー: {e}")
        return

    if not organizer_aad_ids:
        print(f"[Bubble] 受信者なしのためスキップ: {thread_id}")
        return

    try:
        token = await get_bot_framework_token()
    except Exception as e:
        print(f"[Bubble] Bot Framework トークン取得エラー: {e}")
        return

    # AAD Object ID から Teams User MRI (29:xxxx) への変換
    recipients = []
    members_url = f"{service_url.rstrip('/')}/v3/conversations/{thread_id}/members"
    headers = {"Authorization": f"Bearer {token}"}
    try:
        async with aiohttp.ClientSession() as session:
            async with session.get(members_url, headers=headers) as resp:
                if resp.status == 200:
                    members = await resp.json()
                    print(f"[Bubble] members 取得成功: {len(members)}名")
                    for member in members:
                        aad_id = member.get("aadObjectId") or member.get("objectId")
                        if aad_id in organizer_aad_ids:
                            mri_id = member.get("id")
                            if mri_id:
                                recipients.append(mri_id)
                                email = member.get("email") or member.get("userPrincipalName") or "不明"
                                print(f"[Bubble] 主催者のメールアドレス: {email}")
                    print(f"[Bubble] MRI 変換結果: {recipients}")
                else:
                    text = await resp.text()
                    print(f"[Bubble] Members API エラー: status={resp.status} response={text}")
    except Exception as e:
        print(f"[Bubble] Members API 呼び出しエラー: {e}")

    if not recipients:
        print(f"[Bubble] MRI変換失敗のため送信スキップ: {thread_id}")
        return

    notification_url = f"{CONFIG.NOTIFICATION_URL.rstrip('/')}/notification?threadId={quote(thread_id)}"
    body = {
        "type": "targetedMeetingNotification",
        "value": {
            "recipients": recipients,
            "surfaces": [
                {
                    "surface": "meetingStage",
                    "contentType": "task",
                    "content": {
                        "value": {
                            "height": 260,
                            "width": 400,
                            "title": "録画が開始されました",
                            "url": notification_url,
                        }
                    },
                }
            ],
        },
    }
    url = f"{service_url.rstrip('/')}/v1/meetings/{meeting_id}/notification"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    async with aiohttp.ClientSession() as session:
        async with session.post(url, headers=headers, json=body) as resp:
            text = await resp.text()
            print(f"[Bubble] コンテンツバブル送信: status={resp.status} url={url}")
            if resp.status not in (200, 201, 202):
                print(f"[Bubble] レスポンス: {text}")
            else:
                _bubble_sent.add(thread_id)


async def get_chat(thread_id: str) -> dict:
    """チャットの詳細を取得する"""
    token = await get_access_token()
    headers = {"Authorization": f"Bearer {token}"}
    async with aiohttp.ClientSession() as session:
        async with session.get(
            f"https://graph.microsoft.com/v1.0/chats/{thread_id}",
            headers=headers,
        ) as resp:
            return await resp.json()


async def handle_chats_notification(notification: dict):
    """新しいチャット作成通知を処理して会議チャットなら購読する"""
    resource_url = notification.get("resource", "")
    m = re.search(r"chats\('([^']+)'\)", resource_url)
    if not m:
        parts = resource_url.strip("/").split("/")
        thread_id = parts[-1] if len(parts) >= 2 else ""
    else:
        thread_id = m.group(1)

    if not thread_id:
        print(f"[Chats] threadId 取得失敗: {resource_url}")
        return

    if not thread_id.startswith("19:meeting_"):
        try:
            chat = await get_chat(thread_id)
            if chat.get("chatType") != "meeting":
                print(f"[Chats] 会議チャットではないためスキップ: {thread_id}")
                return
        except Exception as e:
            print(f"[Chats] チャット詳細取得エラー: {e}")
            return

    print(f"[Chats] 会議チャット検知: {thread_id}")
    asyncio.create_task(setup_meeting_chat(thread_id))


async def handle_app_removed(thread_id: str):
    """会議チャットからのアンインストール時に購読を削除する"""
    sub_id = _chat_subscriptions.pop(thread_id, None)
    if sub_id:
        await _delete_subscription(sub_id)
    _recording_active.discard(thread_id)
    _integrate_consented.discard(thread_id)


async def _create_chat_message_subscription(thread_id: str) -> dict:
    token = await get_access_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    expiration = datetime.now(timezone.utc) + timedelta(hours=1)
    body = {
        "changeType": "created",
        "notificationUrl": f"{CONFIG.NOTIFICATION_URL}/api/notifications",
        "resource": f"chats/{thread_id}/messages",
        "expirationDateTime": expiration.isoformat(),
        "clientState": "teams-meeting-app",
    }
    async with aiohttp.ClientSession() as session:
        async with session.post(
            "https://graph.microsoft.com/v1.0/subscriptions",
            headers=headers,
            json=body,
        ) as resp:
            return await resp.json()


async def _delete_subscription(sub_id: str):
    try:
        token = await get_access_token()
        headers = {"Authorization": f"Bearer {token}"}
        async with aiohttp.ClientSession() as session:
            async with session.delete(
                f"https://graph.microsoft.com/v1.0/subscriptions/{sub_id}",
                headers=headers,
            ) as resp:
                print(f"[Graph] 購読削除: subId={sub_id} status={resp.status}")
    except Exception as e:
        print(f"[Graph] 購読削除エラー: {e}")


async def notify_azure_recording_stopped(thread_id: str):
    """OneDrive 保存完了を Azure Webhook に通知する"""
    if not CONFIG.AZURE_WEBHOOK_URL:
        print(f"[Azure] AZURE_WEBHOOK_URL 未設定のためスキップ: {thread_id}")
        return
    payload = {
        "event": "recording_saved",
        "threadId": thread_id,
        "recordingUrl": _recording_urls.get(thread_id, ""),
        "timestamp": datetime.now(timezone.utc).isoformat(),
    }
    try:
        async with aiohttp.ClientSession() as session:
            async with session.post(
                CONFIG.AZURE_WEBHOOK_URL,
                json=payload,
                headers={"Content-Type": "application/json"},
                timeout=aiohttp.ClientTimeout(total=30),
            ) as resp:
                print(f"[Azure] Webhook 通知: status={resp.status} threadId={thread_id}")
    except Exception as e:
        print(f"[Azure] Webhook 通知エラー: {e}")


async def get_chat_message(chat_id: str, message_id: str) -> dict:
    token = await get_access_token()
    headers = {"Authorization": f"Bearer {token}"}
    url = f"https://graph.microsoft.com/v1.0/chats/{chat_id}/messages/{message_id}"
    async with aiohttp.ClientSession() as session:
        async with session.get(url, headers=headers) as resp:
            return await resp.json()


async def handle_recording_notification(body: dict):
    """チャットメッセージ通知から録画開始/停止を検知する"""
    notifications = body.get("value", [])
    for notification in notifications:
        client_state = notification.get("clientState", "")

        if client_state == "teams-meeting-app-chats":
            await handle_chats_notification(notification)
            continue

        if client_state != "teams-meeting-app":
            print(f"[Notification] 不明な clientState: {notification.get('clientState')}")
            continue

        resource_url = notification.get("resource", "")
        thread_id = ""
        message_id = ""
        if "chats('" in resource_url:
            m = re.search(r"chats\('([^']+)'\)/messages\('([^']+)'\)", resource_url)
            if m:
                thread_id = m.group(1)
                message_id = m.group(2)
        else:
            parts = resource_url.split("/")
            if len(parts) >= 4 and parts[0] == "chats":
                thread_id = parts[1]
                message_id = parts[3] if len(parts) > 3 else ""

        if not thread_id or not message_id:
            print(f"[Notification] パース失敗: {resource_url}")
            continue

        try:
            message = await get_chat_message(thread_id, message_id)
        except Exception as e:
            print(f"[Chat] メッセージ取得エラー: {e}")
            continue

        message_type = message.get("messageType", "")
        event_detail = message.get("eventDetail") or {}
        odata_type = event_detail.get("@odata.type", "")

        if message_type not in ("systemEventMessage", "unknownFutureValue"):
            continue

        if "callRecordingEventMessageDetail" not in odata_type:
            continue

        recording_status = (
            event_detail.get("recordingStatus")
            or event_detail.get("callRecordingStatus")
            or ""
        )
        print(f"[Chat] 録画イベント検知: status={recording_status} threadId={thread_id}")

        if recording_status == "initial" and thread_id not in _recording_active:
            _recording_active.add(thread_id)
            _bubble_sent.discard(thread_id)
            print(f"[Chat] 録画開始: {thread_id}")
            asyncio.create_task(send_content_bubble(thread_id))

        elif recording_status == "chunkFinished" and thread_id in _recording_active:
            _recording_active.discard(thread_id)
            print(f"[Chat] 録画停止: {thread_id}")

        elif recording_status == "success":
            recording_url = event_detail.get("callRecordingUrl", "")
            if recording_url and thread_id not in _notified_recordings:
                _notified_recordings.add(thread_id)
                _recording_urls[thread_id] = recording_url
                print(f"[Chat] 録画 URL 保存: {thread_id} url={recording_url}")
                if thread_id in _integrate_consented:
                    asyncio.create_task(notify_azure_recording_stopped(thread_id))

        elif recording_status == "failure":
            print(f"[Chat] 録画失敗: {thread_id}")
