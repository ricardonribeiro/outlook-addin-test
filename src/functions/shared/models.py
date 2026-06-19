"""
Payload contract shared between function_app.py and the TypeScript add-in (types.ts).
Wire format uses camelCase throughout.
"""
from typing import TypedDict


# ── Attachment shapes ──────────────────────────────────────────────────────────

class AttachmentInfo(TypedDict):
    """Attachment metadata only — no binary content."""
    name: str
    mimeType: str
    sizeBytes: int


class AttachmentRef(AttachmentInfo, total=False):
    """Attachment metadata + stable blob path after direct upload."""
    blobPath: str


class AttachmentUploadSlot(TypedDict):
    """One entry returned by /api/submissions/prepare per attachment."""
    name: str
    mimeType: str
    sizeBytes: int
    blobPath: str    # e.g. "attachments/550e8400-e29b.pdf" — stable, stored in payload.json
    uploadUrl: str   # write SAS URL valid for 15 minutes — used for the direct PUT


# ── /api/submissions/prepare ──────────────────────────────────────────────────

class PrepareResponse(TypedDict):
    submissionId: str
    attachments: list[AttachmentUploadSlot]


# ── /api/submissions ──────────────────────────────────────────────────────────

class SubmissionPayload(TypedDict):
    """What the add-in POSTs to /api/submissions after all attachments are uploaded."""
    submissionId: str         # echoed back from /prepare
    messageId: str            # Outlook itemId — used as Service Bus message ID for dedup
    correlationId: str        # UUID generated client-side per click
    sender: str
    recipients: list[str]     # to + cc addresses
    subject: str
    timestamp: str            # ISO 8601 email sent date
    bodyText: str
    attachments: list[AttachmentRef]  # metadata + blobPath for each


class SubmissionResponse(TypedDict):
    """What /api/submissions returns."""
    submissionId: str
    payloadPath: str    # stable blob path to submission/SUB-XXXXX/payload.json
    receivedAt: str
    correlationId: str


class SubmissionQueueMessage(TypedDict, total=False):
    """Lightweight message written to Service Bus — consumer fetches payloadPath for full data."""
    submissionId: str
    payloadPath: str        # consumer reads this blob to get the full submission
    correlationId: str
    receivedAt: str
    authenticatedUpn: str
    sender: str             # denormalised for cheap filtering without fetching the blob
    subject: str


REQUIRED_PAYLOAD_FIELDS = (
    "submissionId",
    "messageId",
    "correlationId",
    "sender",
    "recipients",
    "subject",
    "timestamp",
    "bodyText",
    "attachments",
)


def validate_payload(body: dict) -> list[str]:
    """Return a list of missing required field names, empty if valid."""
    return [f for f in REQUIRED_PAYLOAD_FIELDS if f not in body]
