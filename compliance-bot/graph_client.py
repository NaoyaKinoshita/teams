import asyncio
import json
from datetime import datetime, timedelta, timezone

import aiohttp

from config import DefaultConfig

CONFIG = DefaultConfig()

# 参加確立済みの会議 {call_id: {"thread_id": str}}
_active_calls: dict = {}

# 通話とチャットスレッドの対応（callRecords 通知で利用。終了後も保持）
_call_threads: dict = {}

# 録画 URL 通知済みの call_id（二重通知防止）
_notified_recordings: set = set()

# 録画中の call_id（二重通知防止）
_recording_active: set = set()

# Azure 連携に同意済みの call_id
_integrate_consented: set = set()


def consent_azure_integration(call_id: str):
    """Azure 連携への同意を記録する（Adaptive Card の「連携する」押下時に呼ぶ）"""
    _integrate_consented.add(call_id)
    print(f"[Graph] Azure 連携同意: {call_id}")


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


async def create_call_record_subscription() -> dict:
    """callRecords サブスクリプションを登録する（録画 URL 取得用）"""
    token = await get_access_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    expiration = datetime.now(timezone.utc) + timedelta(hours=1)
    body = {
        "changeType": "created",
        "notificationUrl": f"{CONFIG.NOTIFICATION_URL}/api/notifications",
        "resource": "/communications/callRecords",
        "expirationDateTime": expiration.isoformat(),
        "clientState": "teams-compliance-bot",
    }
    async with aiohttp.ClientSession() as session:
        async with session.post(
            "https://graph.microsoft.com/v1.0/subscriptions", headers=headers, json=body
        ) as resp:
            result = await resp.json()
            print(f"[Graph] callRecords サブスクリプション: {result.get('id', result)}")
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


async def notify_azure_recording_stopped(call_id: str, thread_id: str):
    """録画停止を Azure Webhook に通知する"""
    if not CONFIG.AZURE_WEBHOOK_URL:
        print(f"[Azure] AZURE_WEBHOOK_URL 未設定のためスキップ: {call_id}")
        return

    payload = {
        "event": "recording_stopped",
        "callId": call_id,
        "threadId": thread_id,
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
                print(f"[Azure] Webhook 通知: status={resp.status} callId={call_id}")
    except Exception as e:
        print(f"[Azure] Webhook 通知エラー: {e}")


async def get_call_record(call_id: str) -> dict:
    token = await get_access_token()
    headers = {"Authorization": f"Bearer {token}"}
    url = (
        f"https://graph.microsoft.com/v1.0/communications/callRecords/{call_id}"
        "?$expand=sessions($expand=segments)"
    )
    async with aiohttp.ClientSession() as session:
        async with session.get(url, headers=headers) as resp:
            return await resp.json()


async def _notify_recording_url(call_id: str, thread_id: str):
    """通話記録から録画 URL を取得してチャットに通知する（会議終了後にポーリング）"""
    wait_minutes = [2, 3, 5, 5, 10]
    for i, wait in enumerate(wait_minutes):
        await asyncio.sleep(wait * 60)
        try:
            record = await get_call_record(call_id)
            sessions = record.get("sessions", [])
            for session in sessions:
                for segment in session.get("segments", []):
                    for rec in segment.get("recordings", []):
                        content_url = rec.get("contentUrl", "")
                        if content_url:
                            print(f"[Graph] 録画 URL 取得成功: {content_url}")
                            _notified_recordings.add(call_id)
                            await send_text_to_chat(
                                thread_id,
                                f"録画が OneDrive に保存されました。\n{content_url}",
                            )
                            return
            print(f"[Graph] 録画 URL 未取得 (試行 {i+1}/{len(wait_minutes)}): {call_id}")
        except Exception as e:
            print(f"[Graph] 録画 URL 取得エラー (試行 {i+1}): {e}")
    print(f"[Graph] 録画 URL を取得できませんでした: {call_id}")


async def handle_call_notification(body: dict):
    """通話状態変更通知を処理する（コンプライアンス録画方式）"""
    notifications = body.get("value", [])
    for notification in notifications:
        resource_data = notification.get("resourceData", {})
        resource_url = notification.get("resource", "")

        # participants 更新など resourceData がリストの通知はスキップ
        if isinstance(resource_data, list):
            print(f"[Call] 参加者更新通知をスキップ: {resource_url}")
            continue

        url_parts = resource_url.split("/")
        if "calls" in url_parts:
            idx = url_parts.index("calls")
            call_id = url_parts[idx + 1] if idx + 1 < len(url_parts) else resource_data.get("id")
        else:
            call_id = resource_data.get("id", "")

        call_state = resource_data.get("state", "")
        # コンプライアンス録画では recordingStatus が正しく届く
        recording_status = resource_data.get("recordingStatus", "")

        print(f"[Call] callId={call_id} state={call_state} recordingStatus={recording_status}")

        if call_state == "established":
            chat_info = resource_data.get("chatInfo") or {}
            thread_id = chat_info.get("threadId", "")
            if thread_id:
                _active_calls[call_id] = {"thread_id": thread_id}
                _call_threads[call_id] = thread_id
            print(f"[Call] 会議参加確立: {call_id} threadId={thread_id}")

        elif call_state == "terminated":
            active = _active_calls.pop(call_id, None)
            if active:
                thread_id = active["thread_id"]
                print(f"[Call] 会議終了: {call_id}")
                await send_text_to_chat(thread_id, "会議が終了しました。")
                if call_id in _recording_active:
                    _recording_active.discard(call_id)
                    if call_id in _integrate_consented:
                        asyncio.create_task(notify_azure_recording_stopped(call_id, thread_id))
                if call_id in _integrate_consented and call_id not in _notified_recordings:
                    asyncio.create_task(_notify_recording_url(call_id, thread_id))

        # コンプライアンス録画では recordingStatus が call 通知で届く
        if recording_status == "recording" and call_id not in _recording_active:
            _recording_active.add(call_id)
            active = _active_calls.get(call_id)
            thread_id = (active or {}).get("thread_id", "") or _call_threads.get(call_id, "")
            if thread_id:
                print(f"[Call] 録画開始を検知: {call_id}")
                from adaptive_card import create_azure_confirm_card
                card = create_azure_confirm_card(call_id)
                await send_adaptive_card_to_chat(thread_id, card)

        elif recording_status == "notRecording" and call_id in _recording_active:
            _recording_active.discard(call_id)
            active = _active_calls.get(call_id)
            thread_id = (active or {}).get("thread_id", "") or _call_threads.get(call_id, "")
            if thread_id:
                print(f"[Call] 録画停止を検知: {call_id}")
                if call_id in _integrate_consented:
                    # 連携同意済み → Azure 通知 + URL 取得
                    await send_text_to_chat(thread_id, "録画が停止されました。OneDrive への保存完了後に Azure への連携を開始します。")
                    asyncio.create_task(notify_azure_recording_stopped(call_id, thread_id))
                    if call_id not in _notified_recordings:
                        asyncio.create_task(_notify_recording_url(call_id, thread_id))
                else:
                    # スキップ済み → 通知のみ
                    await send_text_to_chat(thread_id, "録画が停止されました。")


async def handle_recording_notification(body: dict):
    """callRecords 通知から録画 URL を通知する"""
    notifications = body.get("value", [])
    for notification in notifications:
        if notification.get("clientState") != "teams-compliance-bot":
            continue

        resource_data = notification.get("resourceData", {})
        call_id = resource_data.get("id")
        if not call_id or call_id in _notified_recordings:
            continue

        record = await get_call_record(call_id)
        thread_id = _call_threads.get(call_id, "")
        if not thread_id:
            print(f"[Graph] callRecords 受信したがスレッド ID 不明: {call_id}")
            continue

        print(f"[Graph] callRecords 通知受信: {call_id}")
        sessions = record.get("sessions", [])
        for session in sessions:
            for segment in session.get("segments", []):
                for rec in segment.get("recordings", []):
                    content_url = rec.get("contentUrl", "")
                    if content_url and call_id not in _notified_recordings:
                        _notified_recordings.add(call_id)
                        await send_text_to_chat(
                            thread_id,
                            f"録画が OneDrive に保存されました。\n{content_url}",
                        )
