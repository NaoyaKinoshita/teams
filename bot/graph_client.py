import asyncio
import json
import urllib.parse
from datetime import datetime, timedelta, timezone

import aiohttp

from config import DefaultConfig

CONFIG = DefaultConfig()

# 会話参照を保持するストア（本番では Redis 等を使用）
_conversation_store: dict = {}

# 参加確立済みの会議 {call_id: {"thread_id": str, "join_url": str}}
_active_calls: dict = {}

# 参加試行済みの join URL（重複排除用）
_scheduled_joins: set = set()

# 参加リクエスト中のコール {call_id: {join_url, attempt}}
_pending_joins: dict = {}

# 通話とチャットスレッドの対応（callRecords 通知で利用。終了後も保持）
_call_threads: dict = {}

# 録画 URL 通知済みの call_id（二重通知防止）
_notified_recordings: set = set()

# 録画中の call_id（録画停止検知用）
_recording_active: set = set()


def save_conversation_reference(key: str, conversation_reference: dict):
    _conversation_store[key] = conversation_reference


def get_all_conversation_references() -> list[dict]:
    return list(_conversation_store.values())


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
    """通話記録サブスクリプションを登録する（フォールバック用）"""
    token = await get_access_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    expiration = datetime.now(timezone.utc) + timedelta(hours=1)
    body = {
        "changeType": "created",
        "notificationUrl": f"{CONFIG.NOTIFICATION_URL}/api/notifications",
        "resource": "/communications/callRecords",
        "expirationDateTime": expiration.isoformat(),
        "clientState": "teams-recording-bot",
    }
    async with aiohttp.ClientSession() as session:
        async with session.post(
            "https://graph.microsoft.com/v1.0/subscriptions", headers=headers, json=body
        ) as resp:
            result = await resp.json()
            print(f"[Graph] callRecords サブスクリプション: {result.get('id', result)}")
            return result


async def create_calendar_subscription() -> dict:
    """recording-bot のカレンダー変更通知サブスクリプションを登録する"""
    token = await get_access_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    expiration = datetime.now(timezone.utc) + timedelta(hours=1)
    body = {
        "changeType": "created,updated",
        "notificationUrl": f"{CONFIG.NOTIFICATION_URL}/api/notifications",
        "resource": f"/users/{CONFIG.RECORDING_BOT_USER_ID}/events",
        "expirationDateTime": expiration.isoformat(),
        "clientState": "teams-recording-bot-calendar",
    }
    async with aiohttp.ClientSession() as session:
        async with session.post(
            "https://graph.microsoft.com/v1.0/subscriptions", headers=headers, json=body
        ) as resp:
            result = await resp.json()
            print(f"[Graph] カレンダーサブスクリプション: {result.get('id', result)}")
            return result


async def get_calendar_event(event_id: str) -> dict:
    """recording-bot のカレンダーイベント詳細を取得する"""
    token = await get_access_token()
    headers = {"Authorization": f"Bearer {token}"}
    url = f"https://graph.microsoft.com/v1.0/users/{CONFIG.RECORDING_BOT_USER_ID}/events/{event_id}"
    async with aiohttp.ClientSession() as session:
        async with session.get(url, headers=headers) as resp:
            return await resp.json()


def parse_teams_join_url(join_url: str) -> dict:
    """Teams 会議 URL から organizer ID と tenant ID を抽出する"""
    try:
        parsed = urllib.parse.urlparse(join_url)
        query = urllib.parse.parse_qs(parsed.query)
        context_str = query.get("context", ["{}"])[0]
        context = json.loads(context_str)
        return {
            "organizer_id": context.get("Oid", ""),
            "tenant_id": context.get("Tid", CONFIG.TENANT_ID),
        }
    except Exception as e:
        print(f"[Graph] URL 解析エラー: {e}")
        return {}


async def get_online_meeting_by_join_url(organizer_id: str, join_url: str) -> dict | None:
    """onlineMeeting を joinUrl から検索して joinMeetingId を取得する"""
    token = await get_access_token()
    headers = {"Authorization": f"Bearer {token}"}
    url = f"https://graph.microsoft.com/v1.0/users/{organizer_id}/onlineMeetings"
    params = {"$filter": f"joinWebUrl eq '{join_url}'"}
    async with aiohttp.ClientSession() as session:
        async with session.get(url, headers=headers, params=params) as resp:
            result = await resp.json()
            meetings = result.get("value", [])
            if meetings:
                return meetings[0]
            print(f"[Graph] onlineMeeting が見つかりません: {result}")
            return None


async def join_meeting(join_url: str) -> dict:
    """Graph Communications API で会議に参加する"""
    info = parse_teams_join_url(join_url)
    if not info.get("organizer_id"):
        print(f"[Graph] オーガナイザー ID を取得できませんでした: {join_url}")
        return {}

    organizer_id = info["organizer_id"]
    tenant_id = info["tenant_id"]

    # joinMeetingId を使った参加（organizerMeetingInfo より汎用的で確実）
    meeting_info = None
    online_meeting = await get_online_meeting_by_join_url(organizer_id, join_url)
    if online_meeting:
        jm_settings = online_meeting.get("joinMeetingIdSettings") or {}
        join_meeting_id = jm_settings.get("joinMeetingId") or online_meeting.get("joinMeetingId")
        passcode = jm_settings.get("passcode")
        if join_meeting_id:
            meeting_info = {
                "@odata.type": "#microsoft.graph.joinMeetingIdMeetingInfo",
                "joinMeetingId": join_meeting_id,
                "passcode": passcode,
            }
            print(f"[Graph] joinMeetingId 使用: {join_meeting_id}")

    # フォールバック: organizerMeetingInfo
    if not meeting_info:
        print(f"[Graph] joinMeetingId 取得失敗。organizerMeetingInfo にフォールバック")
        meeting_info = {
            "@odata.type": "#microsoft.graph.organizerMeetingInfo",
            "organizer": {
                "@odata.type": "#microsoft.graph.identitySet",
                "user": {
                    "@odata.type": "#microsoft.graph.identity",
                    "id": organizer_id,
                    "tenantId": tenant_id,
                },
            },
        }

    token = await get_access_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    body = {
        "callbackUri": f"{CONFIG.NOTIFICATION_URL}/api/calls",
        "requestedModalities": ["audio"],
        "meetingInfo": meeting_info,
        "tenantId": CONFIG.TENANT_ID,
        "mediaConfig": {
            "@odata.type": "#microsoft.graph.serviceHostedMediaConfig",
        },
    }
    async with aiohttp.ClientSession() as session:
        async with session.post(
            "https://graph.microsoft.com/v1.0/communications/calls",
            headers=headers,
            json=body,
        ) as resp:
            result = await resp.json()
            call_id = result.get("id")
            print(f"[Graph] 会議参加リクエスト: {call_id or result}")
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


async def schedule_join(join_url: str, start_time: datetime):
    """会議開始時刻に合わせて参加をスケジュールする"""
    now = datetime.now(timezone.utc)
    delay = (start_time - now).total_seconds()
    if delay > 0:
        print(f"[Graph] {delay:.0f}秒後に会議参加を試みます")
        await asyncio.sleep(delay)
    await _attempt_join(join_url, attempt=1)


async def _attempt_join(join_url: str, attempt: int):
    """会議参加を試みる。結果は /api/calls コールバックで判定する"""
    if attempt > 10:
        print(f"[Graph] 最大試行回数に達しました。登録を解除します")
        _scheduled_joins.discard(join_url)
        return

    result = await join_meeting(join_url)
    call_id = result.get("id")

    if call_id:
        # establishing 状態 → /api/calls のコールバックで established/terminated を待つ
        _pending_joins[call_id] = {"join_url": join_url, "attempt": attempt}
        print(f"[Graph] 参加リクエスト送信 (試行 {attempt}/10): callId={call_id}")
    else:
        error = result.get("error", {})
        print(f"[Graph] 参加リクエスト自体が失敗: {error}")
        _scheduled_joins.discard(join_url)


async def handle_calendar_notification(body: dict):
    """カレンダー変更通知を処理して会議参加をスケジュールする"""
    notifications = body.get("value", [])
    for notification in notifications:
        if notification.get("clientState") != "teams-recording-bot-calendar":
            continue

        resource_data = notification.get("resourceData", {})
        event_id = resource_data.get("id")
        if not event_id:
            continue

        event = await get_calendar_event(event_id)

        # キャンセル済みの会議はスキップ
        if event.get("isCancelled"):
            subject = event.get("subject", "")
            print(f"[Graph] キャンセル済みのためスキップ: {subject}")
            # スケジュール済みの場合は登録を削除
            online_meeting = event.get("onlineMeeting") or {}
            join_url = online_meeting.get("joinUrl", "")
            _scheduled_joins.discard(join_url)
            continue

        online_meeting = event.get("onlineMeeting")
        if not online_meeting:
            continue

        join_url = online_meeting.get("joinUrl")
        if not join_url:
            continue

        # 同じ会議への重複スケジュールを防ぐ
        if join_url in _scheduled_joins:
            print(f"[Graph] 既にスケジュール済みのためスキップ: {event.get('subject')}")
            continue
        _scheduled_joins.add(join_url)

        start_str = event.get("start", {}).get("dateTime")
        time_zone = event.get("start", {}).get("timeZone", "UTC")
        if not start_str:
            continue

        # ISO形式の日時をパース
        start_time = datetime.fromisoformat(start_str.rstrip("Z")).replace(tzinfo=timezone.utc)
        print(f"[Graph] 会議を検出: {event.get('subject')} / 開始: {start_time}")

        asyncio.create_task(schedule_join(join_url, start_time))


async def handle_call_notification(body: dict):
    """通話状態変更通知を処理する"""
    notifications = body.get("value", [])
    for notification in notifications:
        resource_data = notification.get("resourceData", {})
        resource_url = notification.get("resource", "")

        # participants 更新など resourceData がリストの通知はスキップ
        if isinstance(resource_data, list):
            print(f"[Call] 参加者更新通知をスキップ: {resource_url}")
            continue

        # call_id: URL が /calls/{id}/xxx の形式にも対応
        url_parts = resource_url.split("/")
        if "calls" in url_parts:
            idx = url_parts.index("calls")
            call_id = url_parts[idx + 1] if idx + 1 < len(url_parts) else resource_data.get("id")
        else:
            call_id = resource_data.get("id", "")

        call_state = resource_data.get("state", "")
        recording_status = resource_data.get("recordingStatus", "")

        print(f"[Call] callId={call_id} state={call_state} recordingStatus={recording_status}")

        if call_state == "established":
            # 参加確立 → アクティブコールとして登録
            pending = _pending_joins.pop(call_id, {})
            chat_info = resource_data.get("chatInfo") or {}
            thread_id = chat_info.get("threadId", "")
            if thread_id:
                _active_calls[call_id] = {
                    "thread_id": thread_id,
                    "join_url": pending.get("join_url", ""),
                }
                _call_threads[call_id] = thread_id  # callRecords 通知用に永続保持
            print(f"[Call] 会議参加確立: {call_id} threadId={thread_id}")

        elif call_state == "terminated":
            pending = _pending_joins.pop(call_id, {})
            active = _active_calls.pop(call_id, None)
            result_info = resource_data.get("resultInfo") or {}
            subcode = result_info.get("subcode", 0)

            if subcode == 2203 and pending:
                # 会議未開始（誰もいない）→ 2分後にリトライ
                join_url = pending["join_url"]
                attempt = pending["attempt"]
                print(f"[Call] 会議未開始のため2分後にリトライ (試行 {attempt}/10)")
                asyncio.create_task(_retry_join_after_delay(join_url, attempt + 1))
            else:
                # その他の終了（正常終了など）
                join_url = pending.get("join_url", "") or (active or {}).get("join_url", "")
                if join_url:
                    _scheduled_joins.discard(join_url)
                # 会議に参加済みだった場合は終了通知
                if active:
                    thread_id = active["thread_id"]
                    print(f"[Call] 会議終了を検知: {call_id}")
                    await send_text_to_chat(thread_id, "会議が終了しました。")
                    # 録画停止通知がまだ届いていない場合（録画中のまま会議終了）は Azure 通知とポーリングを起動
                    if call_id in _recording_active:
                        _recording_active.discard(call_id)
                        asyncio.create_task(notify_azure_recording_stopped(call_id, thread_id))
                    if call_id not in _notified_recordings:
                        asyncio.create_task(_notify_recording_url(call_id, thread_id))

        # 録画状態の変化を検知
        if recording_status == "recording" and call_id not in _recording_active:
            _recording_active.add(call_id)
            active = _active_calls.get(call_id)
            thread_id = active["thread_id"] if active else ""
            if thread_id:
                print(f"[Call] 録画開始を検知: {call_id}")
                await send_text_to_chat(thread_id, "録画が開始されました。")
            else:
                print(f"[Call] 録画開始を検知したがスレッド ID 不明: {call_id}")

        elif recording_status == "notRecording" and call_id in _recording_active:
            _recording_active.discard(call_id)
            active = _active_calls.get(call_id)
            thread_id = (active["thread_id"] if active else "") or _call_threads.get(call_id, "")
            if thread_id:
                print(f"[Call] 録画停止を検知: {call_id}")
                await send_text_to_chat(thread_id, "録画が停止されました。OneDrive への保存処理が開始されます。")
                asyncio.create_task(notify_azure_recording_stopped(call_id, thread_id))
                asyncio.create_task(_notify_recording_url(call_id, thread_id))
            else:
                print(f"[Call] 録画停止を検知したがスレッド ID 不明: {call_id}")


async def _retry_join_after_delay(join_url: str, attempt: int):
    """指定時間待ってから再参加を試みる"""
    await asyncio.sleep(120)
    await _attempt_join(join_url, attempt)


async def _notify_recording_url(call_id: str, thread_id: str):
    """通話記録から録画 URL を取得してチャットに通知する（会議終了後にポーリング）"""
    wait_minutes = [2, 3, 5, 5, 10]  # 合計 25 分以内に 5 回試行
    for i, wait in enumerate(wait_minutes):
        await asyncio.sleep(wait * 60)
        try:
            record = await get_call_record(call_id)
            # callRecord の sessions > segments から録画情報を探す
            sessions = record.get("sessions", [])
            for session in sessions:
                for segment in session.get("segments", []):
                    recordings = segment.get("recordings", [])
                    for rec in recordings:
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


async def handle_recording_notification(body: dict, bot, adapter):
    """callRecords 通知を処理する"""
    notifications = body.get("value", [])
    for notification in notifications:
        client_state = notification.get("clientState", "")

        if client_state == "teams-recording-bot-calendar":
            await handle_calendar_notification({"value": [notification]})
            continue

        if client_state != "teams-recording-bot":
            continue

        resource_data = notification.get("resourceData", {})
        call_id = resource_data.get("id")
        if not call_id:
            continue

        # 既に録画 URL を通知済みであれば二重通知をスキップ
        if call_id in _notified_recordings:
            print(f"[Graph] 既に録画 URL 通知済みのためスキップ: {call_id}")
            continue

        record = await get_call_record(call_id)
        modalities = record.get("modalities", [])
        if "audio" not in modalities:
            continue

        print(f"[Graph] callRecords 通知受信: {call_id}")

        # スレッド ID が判明している場合は録画 URL を探して通知
        thread_id = _call_threads.get(call_id, "")
        if thread_id:
            sessions = record.get("sessions", [])
            for session in sessions:
                for segment in session.get("segments", []):
                    for rec in segment.get("recordings", []):
                        content_url = rec.get("contentUrl", "")
                        if content_url and call_id not in _notified_recordings:
                            print(f"[Graph] callRecords から録画 URL 取得: {content_url}")
                            _notified_recordings.add(call_id)
                            await send_text_to_chat(
                                thread_id,
                                f"録画が OneDrive に保存されました。\n{content_url}",
                            )
        else:
            print(f"[Graph] callRecords 受信したがスレッド ID 不明: {call_id}")
