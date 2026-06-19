"""
Shared blob utilities — used by submission_prepare, submission_receiver, and download_generator.

Local dev: STORAGE_ACCOUNT_KEY is set → account-key SAS (no RBAC needed).
Azure:     STORAGE_ACCOUNT_KEY is absent → User Delegation SAS via Managed Identity.
"""
import json
import logging
import os
from datetime import UTC, datetime, timedelta

from azure.storage.blob import (
    BlobSasPermissions,
    BlobServiceClient,
    ContentSettings,
    generate_blob_sas,
)

log = logging.getLogger(__name__)

_STORAGE_ACCOUNT_NAME = os.getenv("STORAGE_ACCOUNT_NAME", "")
_STORAGE_ACCOUNT_KEY = os.getenv("STORAGE_ACCOUNT_KEY", "")
_BLOB_CONTAINER = os.getenv("BLOB_CONTAINER_NAME", "submissions")


def _blob_service(credential) -> BlobServiceClient:
    cred = _STORAGE_ACCOUNT_KEY if _STORAGE_ACCOUNT_KEY else credential
    return BlobServiceClient(
        account_url=f"https://{_STORAGE_ACCOUNT_NAME}.blob.core.windows.net",
        credential=cred,
    )


def _sas_url(blob_name: str, credential, permission: BlobSasPermissions, expiry: datetime) -> str:
    if _STORAGE_ACCOUNT_KEY:
        sas = generate_blob_sas(
            account_name=_STORAGE_ACCOUNT_NAME,
            container_name=_BLOB_CONTAINER,
            blob_name=blob_name,
            account_key=_STORAGE_ACCOUNT_KEY,
            permission=permission,
            expiry=expiry,
        )
    else:
        service = BlobServiceClient(
            account_url=f"https://{_STORAGE_ACCOUNT_NAME}.blob.core.windows.net",
            credential=credential,
        )
        delegation_key = service.get_user_delegation_key(
            key_start_time=datetime.now(UTC),
            key_expiry_time=expiry,
        )
        sas = generate_blob_sas(
            account_name=_STORAGE_ACCOUNT_NAME,
            container_name=_BLOB_CONTAINER,
            blob_name=blob_name,
            user_delegation_key=delegation_key,
            permission=permission,
            expiry=expiry,
        )
    return f"https://{_STORAGE_ACCOUNT_NAME}.blob.core.windows.net/{_BLOB_CONTAINER}/{blob_name}?{sas}"


def generate_write_sas(blob_name: str, credential, *, expiry_minutes: int = 15) -> str:
    """Return a short-lived write SAS URL for direct client-side PUT upload."""
    expiry = datetime.now(UTC) + timedelta(minutes=expiry_minutes)
    url = _sas_url(blob_name, credential, BlobSasPermissions(create=True, write=True), expiry)
    log.info("generate_write_sas: blob_name=%s expiry_minutes=%d", blob_name, expiry_minutes)
    return url


def upload_json(data: dict, blob_name: str, credential) -> str:
    """
    Serialize data as JSON and upload to blob storage.
    Returns blob_path (stable reference, no expiry).
    """
    service = _blob_service(credential)
    blob_client = service.get_blob_client(container=_BLOB_CONTAINER, blob=blob_name)
    content = json.dumps(data, indent=2, ensure_ascii=False).encode("utf-8")
    blob_client.upload_blob(
        content,
        overwrite=True,
        content_settings=ContentSettings(content_type="application/json; charset=utf-8"),
    )
    blob_path = f"{_BLOB_CONTAINER}/{blob_name}"
    log.info("upload_json: blob_name=%s", blob_name)
    return blob_path


def upload_zip(
    content: bytes,
    blob_name: str,
    filename: str,
    credential,
    *,
    sas_expiry_hours: int = 1,
) -> tuple[str, str]:
    """
    Upload a ZIP to blob storage and return (blob_path, sas_url).

    blob_path: stable container/blob path reference (no expiry).
    sas_url:   short-lived signed read URL (valid for sas_expiry_hours).
    """
    service = _blob_service(credential)
    blob_client = service.get_blob_client(container=_BLOB_CONTAINER, blob=blob_name)
    blob_client.upload_blob(
        content,
        overwrite=True,
        content_settings=ContentSettings(
            content_type="application/zip",
            content_disposition=f'attachment; filename="{filename}"',
        ),
    )
    log.info("upload_zip: blob_name=%s container=%s", blob_name, _BLOB_CONTAINER)

    expiry = datetime.now(UTC) + timedelta(hours=sas_expiry_hours)
    sas_url = _sas_url(blob_name, credential, BlobSasPermissions(read=True), expiry)
    blob_path = f"{_BLOB_CONTAINER}/{blob_name}"
    return blob_path, sas_url
