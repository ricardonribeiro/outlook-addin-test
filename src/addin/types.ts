// Wire format — camelCase on the wire.
// Keep in sync with src/functions/shared/models.py.

export interface AttachmentInfo {
  name: string;
  mimeType: string;
  sizeBytes: number;
}

export interface AttachmentRef extends AttachmentInfo {
  blobPath: string;  // stable blob path after direct upload
}

export interface AttachmentWithContent extends AttachmentInfo {
  contentBase64: string;  // used locally when building download ZIPs
}

// Returned by /api/submissions/prepare — one entry per attachment
export interface AttachmentUploadSlot extends AttachmentInfo {
  blobPath: string;    // stable path stored in payload.json
  uploadUrl: string;   // write SAS URL valid for 15 min — use for direct PUT
}

// POST /api/submissions/prepare response
export interface PrepareResponse {
  submissionId: string;
  attachments: AttachmentUploadSlot[];
}

// POST /api/submissions request body
export interface SubmissionPayload {
  submissionId: string;       // echoed from PrepareResponse
  messageId: string;          // Outlook itemId — Service Bus dedup key
  correlationId: string;      // crypto.randomUUID() generated at click time
  sender: string;
  recipients: string[];
  subject: string;
  timestamp: string;          // ISO 8601 email sent date
  bodyText: string;
  attachments: AttachmentRef[];
}

// POST /api/submissions response
export interface SubmissionResponse {
  submissionId: string;
  payloadPath: string;   // stable blob reference to submission/{SUB-XXXXX}/payload.json
  receivedAt: string;
  correlationId: string;
}

// POST /api/downloads request
export interface DownloadRequest {
  contentBase64: string;
  filename: string;
}

// POST /api/downloads response
export interface DownloadResponse {
  downloadUrl: string;
}
