import os
from dotenv import load_dotenv

load_dotenv()


class DefaultConfig:
    APP_ID = os.environ.get("MICROSOFT_APP_ID", "")
    APP_SECRET = os.environ.get("MICROSOFT_APP_SECRET", "")
    TENANT_ID = os.environ.get("TENANT_ID", "")
    NOTIFICATION_URL = os.environ.get("NOTIFICATION_URL", "")
    RECORDING_BOT_USER_ID = os.environ.get("RECORDING_BOT_USER_ID", "")
    AZURE_WEBHOOK_URL = os.environ.get("AZURE_WEBHOOK_URL", "")
    PORT = int(os.environ.get("PORT", 3978))
