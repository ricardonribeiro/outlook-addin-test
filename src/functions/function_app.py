"""
Azure Functions Python v2 model entry point.

FRAMEWORK EXCEPTION: this file must live at the project root (src/functions/),
not in a subfolder. The Functions v2 host discovers functions by importing
function_app.py from the working directory. Everything else lives in shared/.

Functions registered:
  - submission_prepare:  POST /api/submissions/prepare  (issues write SAS URLs)
  - submission_receiver: POST /api/submissions           (writes payload.json, enqueues)
  - download_generator:  POST /api/downloads             (uploads ZIP, returns read SAS URL)
  - health:              GET  /api/health                (anonymous liveness check)
"""

import base64
import json
import logging
import os
import re
import uuid
from datetime import UTC, datetime

import azure.functions as func
from azure.identity import DefaultAzureCredential

from shared.auth import TokenValidationError, validate_token
from shared.blob import generate_write_sas, upload_zip
from shared.models import (
    AttachmentUploadSlot,
    PrepareResponse,
    SubmissionQueueMessage,
    SubmissionResponse,
    validate_payload,
)
from shared.service_bus import send_to_queue

logging.basicConfig(
    format="%(levelname)-8s | %(filename)s:%(lineno)d | %(name)s | %(message)s",
    level=logging.INFO,
    force=True,
)

app = func.FunctionApp(http_auth_level=func.AuthLevel.ANONYMOUS)

_ALLOWED_ORIGIN = os.getenv("ALLOWED_CORS_ORIGIN", "*")
_azure_credential = DefaultAzureCredential()


def _cors(extra: dict | None = None) -> dict:
    headers = {
        "Access-Control-Allow-Origin": _ALLOWED_ORIGIN,
        "Access-Control-Allow-Methods": "POST, GET, OPTIONS",
        "Access-Control-Allow-Headers": "Content-Type, Authorization",
    }
    if extra:
        headers.update(extra)
    return headers


# ── Health ────────────────────────────────────────────────────────────────────


@app.route(route="health", methods=["GET"], auth_level=func.AuthLevel.ANONYMOUS)
def health(req: func.HttpRequest) -> func.HttpResponse:
    return func.HttpResponse(
        json.dumps({"status": "ok"}),
        status_code=200,
        mimetype="application/json",
        headers=_cors(),
    )


# ── submission-prepare ────────────────────────────────────────────────────────
# Validates the token, reserves a submissionId, and returns a short-lived write
# SAS URL per attachment so the add-in can PUT each file directly to blob storage
# without routing the binary through this function.


@app.route(route="submissions/prepare", methods=["POST", "OPTIONS"], auth_level=func.AuthLevel.ANONYMOUS)
def submission_prepare(req: func.HttpRequest) -> func.HttpResponse:
    log = logging.getLogger("submission_prepare")

    if req.method == "OPTIONS":
        return func.HttpResponse(status_code=204, headers=_cors())

    log.info("POST /api/submissions/prepare received")

    auth_header = req.headers.get("Authorization", "")
    try:
        claims = validate_token(auth_header)
    except TokenValidationError as exc:
        log.warning("Auth failed: %s", exc)
        return func.HttpResponse(
            json.dumps({"error": "Unauthorized", "detail": str(exc)}),
            status_code=401,
            mimetype="application/json",
            headers=_cors(),
        )

    upn = claims.get("preferred_username") or claims.get("upn") or claims.get("sub", "unknown")

    try:
        body: dict = req.get_json()
    except ValueError:
        return func.HttpResponse(
            json.dumps({"error": "Request body must be valid JSON"}),
            status_code=400,
            mimetype="application/json",
            headers=_cors(),
        )

    attachments: list[dict] = body.get("attachments", [])
    submission_id = f"SUB-{str(uuid.uuid4())[:8].upper()}"

    slots: list[AttachmentUploadSlot] = []
    for att in attachments:
        name: str = att.get("name", "file")
        mime: str = att.get("mimeType", "application/octet-stream")
        size: int = att.get("sizeBytes", 0)
        ext = os.path.splitext(name)[1] or ""
        blob_name = f"attachments/{uuid.uuid4()}{ext}"
        upload_url = generate_write_sas(blob_name, _azure_credential)
        slots.append({
            "name": name,
            "mimeType": mime,
            "sizeBytes": size,
            "blobPath": blob_name,
            "uploadUrl": upload_url,
        })

    log.info("submission_prepare: id=%s attachments=%d upn=%s", submission_id, len(slots), upn)

    response: PrepareResponse = {"submissionId": submission_id, "attachments": slots}
    return func.HttpResponse(
        json.dumps(response),
        status_code=200,
        mimetype="application/json",
        headers=_cors(),
    )


# ── submission-receiver ───────────────────────────────────────────────────────
# Called after all attachments are uploaded directly to blob. Validates the token
# and enqueues the full submission payload — consumer receives everything inline.


@app.route(route="submissions", methods=["POST", "OPTIONS"], auth_level=func.AuthLevel.ANONYMOUS)
def submission_receiver(req: func.HttpRequest) -> func.HttpResponse:
    log = logging.getLogger("submission_receiver")

    if req.method == "OPTIONS":
        return func.HttpResponse(status_code=204, headers=_cors())

    log.info("POST /api/submissions received")

    auth_header = req.headers.get("Authorization", "")
    try:
        claims = validate_token(auth_header)
    except TokenValidationError as exc:
        log.warning("Auth failed: %s", exc)
        return func.HttpResponse(
            json.dumps({"error": "Unauthorized", "detail": str(exc)}),
            status_code=401,
            mimetype="application/json",
            headers=_cors(),
        )

    authenticated_upn = (
        claims.get("preferred_username") or claims.get("upn") or claims.get("sub", "unknown")
    )
    log.info("Token valid for upn=%s", authenticated_upn)

    try:
        body: dict = req.get_json()
    except ValueError:
        return func.HttpResponse(
            json.dumps({"error": "Request body must be valid JSON"}),
            status_code=400,
            mimetype="application/json",
            headers=_cors(),
        )

    missing = validate_payload(body)
    if missing:
        log.warning("Payload invalid — missing fields: %s", missing)
        return func.HttpResponse(
            json.dumps({"error": "Missing required fields", "fields": missing}),
            status_code=400,
            mimetype="application/json",
            headers=_cors(),
        )

    submission_id: str = body["submissionId"]
    correlation_id: str = body["correlationId"]
    message_id: str = body["messageId"]
    received_at = datetime.now(UTC).isoformat()

    log.info(
        "Payload valid — id=%s subject=%r sender=%s attachments=%d corr=%s",
        submission_id,
        body.get("subject"),
        body.get("sender"),
        len(body.get("attachments", [])),
        correlation_id,
    )

    # Enqueue the full payload — consumer receives everything inline, no blob fetch needed.
    queue_message: SubmissionQueueMessage = {
        "submissionId": submission_id,
        "messageId": message_id,
        "correlationId": correlation_id,
        "sender": body.get("sender", ""),
        "recipients": body.get("recipients", []),
        "subject": body.get("subject", ""),
        "timestamp": body.get("timestamp", ""),
        "bodyText": body.get("bodyText", ""),
        "attachments": body.get("attachments", []),
        "receivedAt": received_at,
        "authenticatedUpn": authenticated_upn,
    }
    try:
        send_to_queue(queue_message, message_id=message_id)
    except Exception as exc:
        log.error("Enqueue failed for %s corr=%s: %s", submission_id, correlation_id, exc)
        return func.HttpResponse(
            json.dumps({"error": "Failed to enqueue — see function logs"}),
            status_code=500,
            mimetype="application/json",
            headers=_cors(),
        )

    log.info("Submission enqueued id=%s corr=%s upn=%s", submission_id, correlation_id, authenticated_upn)

    response: SubmissionResponse = {
        "submissionId": submission_id,
        "receivedAt": received_at,
        "correlationId": correlation_id,
    }
    return func.HttpResponse(
        json.dumps(response),
        status_code=202,
        mimetype="application/json",
        headers=_cors(),
    )


# ── download-generator ────────────────────────────────────────────────────────
# Accepts a base64-encoded ZIP, uploads it to blob storage under downloads/,
# and returns a 1-hour SAS URL. The add-in opens the URL in the system browser.


@app.route(route="downloads", methods=["POST", "OPTIONS"], auth_level=func.AuthLevel.ANONYMOUS)
def download_generator(req: func.HttpRequest) -> func.HttpResponse:
    log = logging.getLogger("download_generator")

    if req.method == "OPTIONS":
        return func.HttpResponse(status_code=204, headers=_cors())

    log.info("POST /api/downloads received")

    auth_header = req.headers.get("Authorization", "")
    try:
        claims = validate_token(auth_header)
    except TokenValidationError as exc:
        log.warning("Auth failed: %s", exc)
        return func.HttpResponse(
            json.dumps({"error": "Unauthorized", "detail": str(exc)}),
            status_code=401,
            mimetype="application/json",
            headers=_cors(),
        )

    upn = claims.get("preferred_username") or claims.get("upn") or claims.get("sub", "unknown")
    log.info("Token valid for upn=%s", upn)

    try:
        body: dict = req.get_json()
    except ValueError:
        return func.HttpResponse(
            json.dumps({"error": "Request body must be valid JSON"}),
            status_code=400,
            mimetype="application/json",
            headers=_cors(),
        )

    content_b64: str = body.get("contentBase64", "")
    filename: str = body.get("filename", "download.zip")

    if not content_b64:
        return func.HttpResponse(
            json.dumps({"error": "Missing contentBase64"}),
            status_code=400,
            mimetype="application/json",
            headers=_cors(),
        )

    try:
        content = base64.b64decode(content_b64)
    except Exception:
        return func.HttpResponse(
            json.dumps({"error": "contentBase64 is not valid base64"}),
            status_code=400,
            mimetype="application/json",
            headers=_cors(),
        )

    safe_name = re.sub(r'[^A-Za-z0-9._-]', '-', filename)[:80] or 'download.zip'
    blob_name = f"downloads/{safe_name}"
    try:
        _, sas_url = upload_zip(content, blob_name, filename, _azure_credential)
    except Exception as exc:
        log.error("Blob upload failed: %s", exc)
        return func.HttpResponse(
            json.dumps({"error": "Failed to upload — see function logs"}),
            status_code=500,
            mimetype="application/json",
            headers=_cors(),
        )

    log.info("Download blob uploaded name=%s upn=%s", blob_name, upn)

    return func.HttpResponse(
        json.dumps({"downloadUrl": sas_url}),
        status_code=200,
        mimetype="application/json",
        headers=_cors(),
    )


# ── submission-processor (Service Bus queue trigger) ─────────────────────────
#
# %SERVICEBUS_QUEUE_NAME% expands the app setting at runtime.
# ServiceBusConnection__fullyQualifiedNamespace tells the Functions host to
# authenticate with the function app's Managed Identity (no connection string).

# @app.service_bus_queue_trigger(
#     arg_name="msg",
#     queue_name="%SERVICEBUS_QUEUE_NAME%",
#     connection="ServiceBusConnection",
# )
# def submission_processor(msg: func.ServiceBusMessage) -> None:
#     log = logging.getLogger("submission_processor")
#
#     raw = msg.get_body().decode("utf-8")
#     log.info("Message received from queue delivery_count=%d", msg.delivery_count)
#
#     try:
#         payload: dict = json.loads(raw)
#     except json.JSONDecodeError as exc:
#         log.error("Failed to parse message body: %s — preview: %s", exc, raw[:200])
#         raise
#
#     log.info(
#         "Processing submission id=%s corr=%s subject=%r sender=%s upn=%s payload=%s",
#         payload.get("submissionId"),
#         payload.get("correlationId"),
#         payload.get("subject"),
#         payload.get("sender"),
#         payload.get("authenticatedUpn"),
#         payload.get("payloadPath"),
#     )
#
#     # Consumer would: fetch payloadPath from blob, read payload.json,
#     # then process attachments at their blobPaths.
#
#     log.info("Submission processed id=%s corr=%s", payload.get("submissionId"), payload.get("correlationId"))
#     # Returning normally = message acknowledged (removed from queue).
#     # Raising an exception = message returns to queue for retry / dead-letter.
