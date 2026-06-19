# Attachment Upload Plan

## Problem

The Service Bus Standard SKU has a **256 KB message limit**. Base64-encoding attachment
content inside the JSON payload adds ~33% overhead, so anything beyond a small file will
breach this limit. The current `syncToTestApi` flow intentionally omits attachment content
for this reason.

## Solution: SAS-based direct upload to Blob Storage

The add-in uploads each attachment directly to Azure Blob Storage using a short-lived SAS
URL issued by the Function App. The enquiry payload then carries a `blobUrl` reference
instead of raw bytes.

## Flow

```
Add-in
  1. For each attachment, call:
       POST /api/upload-sas?filename=invoice.pdf&correlationId=<uuid>
     Function returns a short-lived SAS URL (5 min TTL, write-only, scoped to one blob).

  2. PUT <sas-url>  (raw bytes, direct to Blob Storage — bypasses the Function App entirely)

  3. POST /api/enquiries  with payload:
       attachments: [{ name, mimeType, sizeBytes, blobUrl }]
       (blobUrl is the blob path without the SAS token — Function App reads via Managed Identity)

Function: enquiry_receiver
  4. Validates the Bearer token.
  5. Builds the Service Bus message with blobUrl references (no bytes on the wire).

Service Bus message
  6. Stays well under 256 KB regardless of attachment size.

Function: enquiry_processor
  7. Reads attachment bytes from Blob Storage via Managed Identity when needed.
```

## What needs to be built

### Infra (`infra/main.tf`)
- Add a `attachments` container to the existing `azurerm_storage_account.main`.
- Grant the Function App's Managed Identity `Storage Blob Data Contributor` on the container.
- Grant developers' accounts `Storage Blob Data Reader` for local testing.

### Function App — new endpoint (`src/functions/function_app.py`)
```python
# GET /api/upload-sas?filename=<name>&correlationId=<uuid>
# Returns: { "uploadUrl": "<sas-url>", "blobUrl": "<canonical-path>" }
```
- Generates a SAS URL using `generate_blob_sas` (azure-storage-blob SDK).
- SAS permissions: write-only, 5-minute TTL.
- Blob path: `attachments/{correlationId}/{filename}` (namespaced by correlationId to avoid collisions).
- Requires the storage account connection string or Managed Identity credential.

### Add-in (`src/addin/commands.ts`)
- In `syncToTestApi`, after acquiring the token, fetch a SAS URL per attachment:
  ```typescript
  const { uploadUrl, blobUrl } = await getSasUrl(token, apiBase, correlationId, att.name);
  await uploadAttachment(uploadUrl, attachmentContent);
  ```
- Map attachments to `{ name, mimeType, sizeBytes, blobUrl }` in the payload.

### Types (`src/addin/types.ts` + `src/functions/shared/models.py`)
- Add `blobUrl?: string` to `AttachmentInfo` / the Python equivalent.
- Both sides must stay in lockstep (as noted in the existing type comments).

## Type changes

```typescript
// types.ts
export interface AttachmentInfo {
  name: string;
  mimeType: string;
  sizeBytes: number;
  blobUrl?: string;   // set when attachment was uploaded; absent for metadata-only payloads
}
```

## Security notes

- The SAS URL is write-only and expires in 5 minutes — the add-in cannot read other blobs.
- The canonical `blobUrl` stored in the payload has no SAS token; only the Function App
  (via Managed Identity) or users with explicit role assignments can read it.
- Never store the SAS token itself in the Service Bus message.

## Defer until

This is only needed when moving to a real deployment or when testing with emails that have
large attachments. For local dev, the "Download Payload" button already covers the
"inspect the full payload including attachment content" use case.
