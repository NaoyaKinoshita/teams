from botbuilder.core import TurnContext
from botbuilder.core.teams import TeamsActivityHandler

from graph_client import handle_app_removed, setup_meeting_chat, store_meeting_context


class MeetingRecordingBot(TeamsActivityHandler):

    def _capture_meeting_context(self, turn_context: TurnContext):
        conversation = turn_context.activity.conversation
        channel_data = turn_context.activity.channel_data or {}
        service_url = turn_context.activity.service_url or ""

        if not conversation or not conversation.id.startswith("19:meeting_"):
            return

        meeting = channel_data.get("meeting") if isinstance(channel_data, dict) else None
        meeting_id = (meeting or {}).get("id", "") if isinstance(meeting, dict) else ""
        if meeting_id and service_url:
            store_meeting_context(conversation.id, meeting_id, service_url)

    async def on_turn(self, turn_context: TurnContext):
        try:
            self._capture_meeting_context(turn_context)
        except Exception as e:
            print(f"[Bot] 会議コンテキスト取得エラー: {e}")
        await super().on_turn(turn_context)

    async def on_installation_update_activity(self, turn_context: TurnContext):
        action = turn_context.activity.action
        conversation = turn_context.activity.conversation
        channel_data = turn_context.activity.channel_data or {}

        meeting = channel_data.get("meeting") if isinstance(channel_data, dict) else None
        is_meeting_chat = bool(meeting) or conversation.id.startswith("19:meeting_")

        if not is_meeting_chat:
            return

        thread_id = conversation.id
        if action == "add":
            print(f"[Bot] 会議チャットにインストール: {thread_id}")
            await setup_meeting_chat(thread_id)
        elif action == "remove":
            print(f"[Bot] 会議チャットからアンインストール: {thread_id}")
            await handle_app_removed(thread_id)
