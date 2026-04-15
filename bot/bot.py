from botbuilder.core import CardFactory, TurnContext
from botbuilder.core.teams import TeamsActivityHandler
from botbuilder.schema import Activity, ActivityTypes

from adaptive_card import create_azure_confirm_card
from graph_client import save_conversation_reference


class TeamsRecordingBot(TeamsActivityHandler):

    def __init__(self, app_id: str):
        self._app_id = app_id

    async def on_teams_members_added_async(self, members_added, team_info, turn_context: TurnContext):
        """Bot が会話に追加されたときの処理（500エラーの原因となるデフォルト動作を上書き）"""
        from botbuilder.core import TurnContext as TC
        conv_ref = TC.get_conversation_reference(turn_context.activity)
        save_conversation_reference(turn_context.activity.conversation.id, conv_ref.serialize())
        await turn_context.send_activity("Teams Recording Bot が追加されました。録画の開始を監視しています。")

    async def on_message_activity(self, turn_context: TurnContext):
        """メッセージアクティビティを処理する（Adaptive Card の応答含む）"""
        from botbuilder.core import TurnContext as TC
        conv_ref = TC.get_conversation_reference(turn_context.activity)
        save_conversation_reference(turn_context.activity.conversation.id, conv_ref.serialize())

        if turn_context.activity.value:
            await self._handle_card_submit(turn_context)
        else:
            await turn_context.send_activity("Teams Recording Bot が起動しています。録画の開始を監視中です。")

    async def _handle_card_submit(self, turn_context: TurnContext):
        """Adaptive Card のボタン押下を処理する"""
        value = turn_context.activity.value
        action = value.get("action")
        meeting_id = value.get("meetingId", "")

        if action == "integrate":
            await turn_context.send_activity(f"Azure への連携を開始します。\n通話 ID: {meeting_id}")
            # TODO: Azure 連携の実装（例: Azure Storage への録画保存）
        elif action == "skip":
            await turn_context.send_activity("Azure 連携をスキップしました。")

    async def on_event_activity(self, turn_context: TurnContext):
        """会議内イベント（録画開始・停止など）を処理する"""
        event_name = turn_context.activity.name or ""
        print(f"[event] {event_name}")

        if "recording" in event_name.lower():
            value = turn_context.activity.value or {}
            status = value.get("status", "")
            print(f"[event] 録画ステータス変更: {status}")

            if status == "recordingStarted":
                # 会話参照を保存して Adaptive Card を送信
                from botbuilder.core import TurnContext as TC
                conv_ref = TC.get_conversation_reference(turn_context.activity)
                save_conversation_reference(turn_context.activity.conversation.id, conv_ref.serialize())
                await self._send_recording_card(turn_context, turn_context.activity.conversation.id)

    async def _send_recording_card(self, turn_context: TurnContext, meeting_id: str):
        """録画開始の Adaptive Card を送信する"""
        card = create_azure_confirm_card(meeting_id)
        await turn_context.send_activity(
            Activity(
                type=ActivityTypes.message,
                attachments=[CardFactory.adaptive_card(card)],
            )
        )

    async def on_teams_meeting_start_async(self, meeting, turn_context: TurnContext):
        """会議開始イベントを処理する"""
        from botbuilder.core import TurnContext as TC
        conv_ref = TC.get_conversation_reference(turn_context.activity)
        save_conversation_reference(turn_context.activity.conversation.id, conv_ref.serialize())
        await turn_context.send_activity("会議が開始されました。録画の開始を監視しています。")

    async def on_teams_meeting_end_async(self, meeting, turn_context: TurnContext):
        """会議終了イベントを処理する"""
        await turn_context.send_activity("会議が終了しました。")

    async def send_recording_started_card(self, conversation_reference: dict, adapter, call_id: str = ""):
        """Graph API 通知経由で Adaptive Card を送信する（callRecords フォールバック用）"""
        from botbuilder.schema import ConversationReference
        conv_ref = ConversationReference().deserialize(conversation_reference)

        async def _send_card(turn_context: TurnContext):
            card = create_azure_confirm_card(call_id)
            await turn_context.send_activity(
                Activity(
                    type=ActivityTypes.message,
                    attachments=[CardFactory.adaptive_card(card)],
                )
            )

        await adapter.continue_conversation(conv_ref, _send_card, self._app_id)
