import base64
import logging
import os
import re

import aiohttp
from azure.storage.blob.aio import BlobServiceClient
from fastapi import BackgroundTasks, FastAPI, Request, Response

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

APP_ID = os.getenv("MICROSOFT_APP_ID", "")
APP_SECRET = os.getenv("MICROSOFT_APP_SECRET", "")
TENANT_ID = os.getenv("TENANT_ID", "")
STORAGE_CONNECTION_STRING = os.getenv("AZURE_STORAGE_CONNECTION_STRING", "")
STORAGE_CONTAINER = os.getenv("STORAGE_CONTAINER_NAME", "recordings")

app = FastAPI()


async def get_graph_token() -> str:
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": APP_ID,
        "client_secret": APP_SECRET,
        "scope": "https://graph.microsoft.com/.default",
    }
    async with aiohttp.ClientSession() as session:
        async with session.post(url, data=data) as resp:
            result = await resp.json()
            if "access_token" not in result:
                raise RuntimeError(f"トークン取得失敗: {result}")
            return result["access_token"]


async def get_sharepoint_item(sharing_url: str, token: str) -> dict:
    """SharePoint 共有 URL から driveItem を取得する"""
    encoded = base64.urlsafe_b64encode(sharing_url.encode()).decode().rstrip("=")
    url = f"https://graph.microsoft.com/v1.0/shares/u!{encoded}/driveItem"
    headers = {"Authorization": f"Bearer {token}"}
    async with aiohttp.ClientSession() as session:
        async with session.get(url, headers=headers) as resp:
            result = await resp.json()
            if resp.status != 200:
                raise RuntimeError(f"driveItem 取得失敗 status={resp.status}: {result}")
            return result


async def download_to_blob(sharing_url: str, thread_id: str):
    """SharePoint から動画をストリーム取得して Blob Storage に保存する"""
    if not STORAGE_CONNECTION_STRING:
        logger.warning("[Blob] AZURE_STORAGE_CONNECTION_STRING 未設定のためスキップ")
        return

    try:
        token = await get_graph_token()
        item = await get_sharepoint_item(sharing_url, token)

        download_url = item.get("@microsoft.graph.downloadUrl", "")
        filename = item.get("name", "recording.mp4")

        if not download_url:
            logger.error("[Blob] downloadUrl が取得できませんでした")
            return

        # blob パス: {sanitized_thread_id}/{filename}
        thread_safe = re.sub(r"[^a-zA-Z0-9_-]", "-", thread_id)
        blob_path = f"{thread_safe}/{filename}"

        logger.info(f"[Blob] アップロード開始: {blob_path}")

        # upload_blob_from_url で Azure が直接 SharePoint からストリーム取得
        async with BlobServiceClient.from_connection_string(STORAGE_CONNECTION_STRING) as client:
            blob = client.get_blob_client(container=STORAGE_CONTAINER, blob=blob_path)
            await blob.upload_blob_from_url(source_url=download_url, overwrite=True)

        logger.info(f"[Blob] アップロード完了: {blob_path}")

    except Exception as e:
        logger.error(f"[Blob] エラー: {e}")


@app.post("/api/recording-stopped")
async def recording_stopped(req: Request, background_tasks: BackgroundTasks):
    try:
        body = await req.json()
    except Exception:
        return Response(content='{"error": "Invalid JSON"}', media_type="application/json", status_code=400)

    thread_id = body.get("threadId", "")
    recording_url = body.get("recordingUrl", "")
    timestamp = body.get("timestamp", "")

    logger.info(f"[recording-stopped] threadId={thread_id} url={recording_url} timestamp={timestamp}")

    if recording_url:
        background_tasks.add_task(download_to_blob, recording_url, thread_id)

    return {"status": "received", "threadId": thread_id}
