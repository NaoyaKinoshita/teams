from botbuilder.core import TurnContext
from botbuilder.core.teams import TeamsActivityHandler

from graph_client import consent_azure_integration, handle_app_installed, handle_app_removed


class MeetingRecordingBot(TeamsActivityHandler):

    async def on_installation_update_activity(self, turn_context: TurnContext):
        """アプリのインストール/アンインストール時に呼ばれる"""
        action = turn_context.activity.action  # "add" or "remove"
        conversation = turn_context.activity.conversation
        channel_data = turn_context.activity.channel_data or {}

        # 会議チャットへのインストールかどうかを判定
        # - channelData に meeting が含まれる場合
        # - conversation.id が "19:meeting_" で始まる場合
        meeting = channel_data.get("meeting") if isinstance(channel_data, dict) else None
        is_meeting_chat = bool(meeting) or conversation.id.startswith("19:meeting_")

        if not is_meeting_chat:
            return

        thread_id = conversation.id
        if action == "add":
            print(f"[Bot] 会議チャットにインストール: {thread_id}")
            await handle_app_installed(thread_id)
        elif action == "remove":
            print(f"[Bot] 会議チャットからアンインストール: {thread_id}")
            await handle_app_removed(thread_id)

    async def on_message_activity(self, turn_context: TurnContext):
        """Adaptive Card のボタン押下を処理する"""
        if turn_context.activity.value:
            await self._handle_card_submit(turn_context)

    async def _handle_card_submit(self, turn_context: TurnContext):
        value = turn_context.activity.value
        action = value.get("action")
        thread_id = value.get("callId", "")  # callId フィールドに threadId を格納している

        if action == "integrate":
            consent_azure_integration(thread_id)
            await turn_context.send_activity("Azure 連携を受け付けました。録画停止後に処理を開始します。")
        elif action == "skip":
            await turn_context.send_activity("Azure 連携をスキップしました。")
