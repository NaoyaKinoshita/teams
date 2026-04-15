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
        await turn_context.send_activity("Compliance Recording Bot が起動しています。")
