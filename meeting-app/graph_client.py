import asyncio
import json
from datetime import datetime, timedelta, timezone

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

# 会議主催者情報 {thread_id: {"id": ..., "displayName": ..., "email": ...}}
_meeting_organizers: dict = {}


def get_recording_status(thread_id: str) -> dict:
    """タブ向けに録画状態・同意状態を返す"""
    return {
        "recording": thread_id in _recording_active,
        "consented": thread_id in _integrate_consented,
        "recordingUrl": _recording_urls.get(thread_id, ""),
    }


def consent_azure_integration(thread_id: str):
    """Azure 連携への同意を記録する（Adaptive Card の「連携する」押下時に呼ぶ）"""
    _integrate_consented.add(thread_id)
    print(f"[Graph] Azure 連携同意: {thread_id}")


async def get_access_token() -> str:
    """クライアントクレデンシャルフローでアクセストークンを取得"""
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


async def _install_app_in_chat(thread_id: str):
    """会議チャットにアプリをインストールして Bot をメンバーにする"""
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
    """会議チャットのアプリインストールとメッセージ購読を行う"""
    if thread_id in _chat_subscriptions:
        print(f"[Graph] 既に購読中のためスキップ: {thread_id}")
        return

    # Bot を会議チャットのメンバーにする（メッセージ送信に必要）
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


# 後方互換性のためエイリアスを残す
handle_app_installed = setup_meeting_chat


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
        await asyncio.sleep(50 * 60)  # 50分待機（1時間期限の10分前）
        print("[Graph] 購読更新タスク起動")

        # /chats 購読の更新
        if _chats_subscription_id:
            ok = await _renew_subscription(_chats_subscription_id)
            if not ok:
                # 更新失敗時は再作成
                print("[Graph] /chats 購読を再作成します")
                try:
                    await create_chats_subscription()
                except Exception as e:
                    print(f"[Graph] /chats 購読再作成エラー: {e}")

        # チャットメッセージ購読の更新
        for thread_id, sub_id in list(_chat_subscriptions.items()):
            ok = await _renew_subscription(sub_id)
            if not ok:
                # 更新失敗時は再作成
                print(f"[Graph] チャット購読を再作成します: {thread_id}")
                _chat_subscriptions.pop(thread_id, None)
                try:
                    await setup_meeting_chat(thread_id)
                except Exception as e:
                    print(f"[Graph] チャット購読再作成エラー: {e}")


def get_meeting_organizer(thread_id: str) -> dict:
    """会議の主催者情報を返す（未取得の場合は空 dict）"""
    return _meeting_organizers.get(thread_id, {})


async def get_chat_organizer(thread_id: str) -> dict:
    """チャットメンバーから主催者（roles に owner を持つメンバー）を取得して返す"""
    token = await get_access_token()
    headers = {"Authorization": f"Bearer {token}"}
    async with aiohttp.ClientSession() as session:
        async with session.get(
            f"https://graph.microsoft.com/v1.0/chats/{thread_id}/members",
            headers=headers,
        ) as resp:
            result = await resp.json()
            for member in result.get("value", []):
                if "owner" in member.get("roles", []):
                    return {
                        "id": member.get("userId", ""),
                        "displayName": member.get("displayName", ""),
                        "email": member.get("email", ""),
                    }
    return {}


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
    print(f"[Chats] 会議参加通知 全情報:\n{json.dumps(notification, ensure_ascii=False, indent=2)}")
    resource_url = notification.get("resource", "")
    # resource: /chats('19:meeting_...')
    import re
    m = re.search(r"chats\('([^']+)'\)", resource_url)
    if not m:
        # スラッシュ区切り形式: /chats/{id}
        parts = resource_url.strip("/").split("/")
        thread_id = parts[-1] if len(parts) >= 2 else ""
    else:
        thread_id = m.group(1)

    if not thread_id:
        print(f"[Chats] threadId 取得失敗: {resource_url}")
        return

    # 会議チャットかどうか確認
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

    try:
        organizer = await get_chat_organizer(thread_id)
        if organizer:
            _meeting_organizers[thread_id] = organizer
            print(f"[Chats] 主催者情報: {json.dumps(organizer, ensure_ascii=False)}")
        else:
            print(f"[Chats] 主催者情報取得失敗（owner ロールなし）: {thread_id}")
    except Exception as e:
        print(f"[Chats] 主催者情報取得エラー: {e}")

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


async def send_text_to_chat(thread_id: str, text: str):
    """会議チャットにテキストメッセージを送信する"""
    token = await get_access_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    body = {"body": {"contentType": "text", "content": text}}
    async with aiohttp.ClientSession() as session:
        async with session.post(
            f"https://graph.microsoft.com/v1.0/chats/{thread_id}/messages",
            headers=headers,
            json=body,
        ) as resp:
            result = await resp.json()
            print(f"[Graph] チャットへメッセージ送信: {result.get('id', result)}")
            return result


async def send_adaptive_card_to_chat(thread_id: str, card: dict):
    """会議チャットに Adaptive Card を送信する"""
    token = await get_access_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    body = {
        "body": {
            "contentType": "html",
            "content": '<attachment id="card"></attachment>',
        },
        "attachments": [
            {
                "id": "card",
                "contentType": "application/vnd.microsoft.card.adaptive",
                "content": json.dumps(card),
            }
        ],
    }
    async with aiohttp.ClientSession() as session:
        async with session.post(
            f"https://graph.microsoft.com/v1.0/chats/{thread_id}/messages",
            headers=headers,
            json=body,
        ) as resp:
            result = await resp.json()
            print(f"[Graph] チャットへ Adaptive Card 送信: {result.get('id', result)}")
            return result


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
    print(f"[Notification] 受信ボディ全体:\n{json.dumps(body, ensure_ascii=False, indent=2)}")
    notifications = body.get("value", [])
    for notification in notifications:
        client_state = notification.get("clientState", "")

        # /chats 購読からの通知（会議チャット自動検知）
        if client_state == "teams-meeting-app-chats":
            await handle_chats_notification(notification)
            continue

        if client_state != "teams-meeting-app":
            print(f"[Notification] 不明な clientState: {notification.get('clientState')}")
            continue

        resource_url = notification.get("resource", "")
        # resource 形式:
        #   chats/{threadId}/messages/{messageId}  (スラッシュ区切り)
        #   chats('{threadId}')/messages('{messageId}')  (OData 形式)
        thread_id = ""
        message_id = ""
        if "chats('" in resource_url:
            # OData 形式をパース
            import re
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

        print(f"[Notification] resource={resource_url}")
        print(f"[Notification] messageType={message_type} eventType={odata_type or '(なし)'}")

        # systemEventMessage または unknownFutureValue（新形式）を処理対象とする
        if message_type not in ("systemEventMessage", "unknownFutureValue"):
            body_content = (message.get("body") or {}).get("content", "")[:80]
            print(f"[Notification] 通常メッセージ: {body_content}")
            continue

        if "callRecordingEventMessageDetail" not in odata_type:
            print(f"[Notification] 録画以外のイベント: {odata_type}")
            continue

        print(f"[Chat] event_detail 全体: {event_detail}")
        recording_status = (
            event_detail.get("recordingStatus")
            or event_detail.get("callRecordingStatus")
            or ""
        )
        print(f"[Chat] 録画イベント検知: status={recording_status} threadId={thread_id}")

        if recording_status == "initial" and thread_id not in _recording_active:
            # 録画開始 → 状態を記録（タブ側がポーリングで検知してコンセント UI を表示）
            _recording_active.add(thread_id)
            print(f"[Chat] 録画開始を記録: {thread_id}")

        elif recording_status == "chunkFinished" and thread_id in _recording_active:
            # 録画停止 → 状態を更新（Azure 通知は OneDrive 保存完了後に行う）
            _recording_active.discard(thread_id)
            print(f"[Chat] 録画停止: {thread_id}")

        elif recording_status == "success":
            # OneDrive 保存完了 → URL 確定後に Azure 通知
            recording_url = event_detail.get("callRecordingUrl", "")
            if recording_url and thread_id not in _notified_recordings:
                _notified_recordings.add(thread_id)
                _recording_urls[thread_id] = recording_url
                print(f"[Chat] 録画 URL 保存: {thread_id} url={recording_url}")
                if thread_id in _integrate_consented:
                    asyncio.create_task(notify_azure_recording_stopped(thread_id))

        elif recording_status == "failure":
            print(f"[Chat] 録画失敗: {thread_id}")
