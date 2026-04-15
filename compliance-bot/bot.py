from botbuilder.core import TurnContext
from botbuilder.core.teams import TeamsActivityHandler


class ComplianceRecordingBot(TeamsActivityHandler):

    def __init__(self, app_id: str):
        self._app_id = app_id

    async def on_teams_members_added_async(self, members_added, team_info, turn_context: TurnContext):
        """Bot が会話に追加されたときの処理"""
        from botbuilder.core import TurnContext as TC
        conv_ref = TC.get_conversation_reference(turn_context.activity)
        print(f"[Bot] 会話追加: {conv_ref.conversation.id}")

    async def on_message_activity(self, turn_context: TurnContext):
        """メッセージ・Adaptive Card のボタン押下を処理する"""
        if turn_context.activity.value:
            await self._handle_card_submit(turn_context)
        else:
            await turn_context.send_activity("Compliance Recording Bot が起動しています。")

    async def _handle_card_submit(self, turn_context: TurnContext):
        """Adaptive Card のボタン押下を処理する"""
        from graph_client import consent_azure_integration

        value = turn_context.activity.value
        action = value.get("action")
        call_id = value.get("callId", "")

        if action == "integrate":
            consent_azure_integration(call_id)
            await turn_context.send_activity("Azure 連携を受け付けました。録画停止後に処理を開始します。")
        elif action == "skip":
            await turn_context.send_activity("Azure 連携をスキップしました。")
