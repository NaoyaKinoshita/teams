import os
from dotenv import load_dotenv

load_dotenv()


class DefaultConfig:
    APP_ID = os.environ.get("MICROSOFT_APP_ID", "")
    APP_SECRET = os.environ.get("MICROSOFT_APP_SECRET", "")
    TENANT_ID = os.environ.get("TENANT_ID", "")
    NOTIFICATION_URL = os.environ.get("NOTIFICATION_URL", "")
    AZURE_WEBHOOK_URL = os.environ.get("AZURE_WEBHOOK_URL", "")
    PORT = int(os.environ.get("PORT", 3980))
    # Teams アプリカタログ上の ID（管理センターのアプリ管理画面で確認できる値）
    TEAMS_APP_ID = os.environ.get("TEAMS_APP_ID", "e007c6dd-0c55-4f39-8c07-bb325b6e6a6d")
