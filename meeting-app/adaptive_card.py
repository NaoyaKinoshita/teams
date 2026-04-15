def create_azure_confirm_card(thread_id: str = "") -> dict:
    """録画開始時に Azure 連携の意思確認を行う Adaptive Card を生成する"""
    return {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.4",
        "body": [
            {
                "type": "TextBlock",
                "size": "Medium",
                "weight": "Bolder",
                "text": "録画が開始されました",
            },
            {
                "type": "TextBlock",
                "text": "この録画を Azure に連携しますか？連携すると録画データを Azure Storage に自動保存できます。",
                "wrap": True,
            },
        ],
        "actions": [
            {
                "type": "Action.Submit",
                "title": "Azure に連携する",
                "data": {"action": "integrate", "callId": thread_id},
                "style": "positive",
            },
            {
                "type": "Action.Submit",
                "title": "スキップ",
                "data": {"action": "skip", "callId": thread_id},
            },
        ],
    }
