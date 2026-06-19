import {
  createNestablePublicClientApplication,
  type IPublicClientApplication,
  InteractionRequiredAuthError,
} from '@azure/msal-browser';
import JSZip from 'jszip';
import type {
  AttachmentUploadSlot,
  AttachmentWithContent,
  DownloadResponse,
  PrepareResponse,
  SubmissionPayload,
  SubmissionResponse,
} from './types';

const NOTIFICATION_KEY = 'submission-sync';
const TIMEOUT_MS = 30_000;

// bind preserves the original call site — DevTools shows the real commands.ts:LINE
const log = {
  info:  console.info.bind(console,  '[Submission]'),
  warn:  console.warn.bind(console,  '[Submission]'),
  error: console.error.bind(console, '[Submission]'),
};

const CLIENT_ID = import.meta.env.VITE_CLIENT_ID as string;
const TENANT_ID = import.meta.env.VITE_TENANT_ID as string;
const API_SCOPE  = import.meta.env.VITE_API_SCOPE as string;

// createNestablePublicClientApplication auto-detects Office NAA context and
// uses the Office host as a token broker (NestedAppAuthController) when available,
// falling back to StandardController otherwise.
let _msalApp: IPublicClientApplication | null = null;

async function getMsalApp(): Promise<IPublicClientApplication> {
  if (_msalApp) return _msalApp;
  _msalApp = await createNestablePublicClientApplication({
    auth: {
      clientId: CLIENT_ID,
      authority: `https://login.microsoftonline.com/${TENANT_ID}`,
      // NAA broker for OWA requires this exact scheme (brk-{broker-client-id}://auth).
      // Without it, MSAL defaults to brk-multihub://localhost:3000 which Azure AD rejects.
      redirectUri: 'brk-9199bf20-a13f-4107-85dc-02114787ef48://auth',
    },
  });
  return _msalApp;
}

Office.onReady(() => {
  Office.actions.associate('syncToTestApi', syncToTestApi);
  Office.actions.associate('downloadPayload', downloadPayload);
});

// ── syncToTestApi ─────────────────────────────────────────────────────────────
// Three-step flow:
//   1. POST /api/submissions/prepare  → submissionId + per-attachment write SAS URLs
//   2. PUT each attachment directly to blob (no function hop, raw binary)
//   3. POST /api/submissions with metadata + blobPaths

async function syncToTestApi(event: Office.AddinCommands.Event): Promise<void> {
  const item = Office.context.mailbox.item as Office.MessageRead | null;
  if (!item) { event.completed(); return; }

  log.info('syncToTestApi: started', { subject: item.subject });
  await notify(item, 'info', 'Submitting to enquiry system…');

  try {
    let token: string;
    try {
      log.info('syncToTestApi: acquiring NAA token');
      const msalApp   = await getMsalApp();
      const loginHint = Office.context.mailbox?.userProfile?.emailAddress;
      const result    = await Promise.race([
        msalApp.ssoSilent({ scopes: [API_SCOPE], loginHint }),
        new Promise<never>((_, reject) =>
          setTimeout(() => reject(new Error('NAA token timed out (10s). Check App Registration SPA platform and access_as_user scope.')), 10_000)
        ),
      ]);
      token = result.accessToken;
      log.info('syncToTestApi: token acquired', { scopes: result.scopes });
    } catch (err: unknown) {
      const isInteraction = err instanceof InteractionRequiredAuthError;
      log.warn('syncToTestApi: token acquisition failed', { interactionRequired: isInteraction, error: String((err as Error)?.message ?? err) });
      await notify(item, 'error', isInteraction
        ? 'Sign-in required — open the add-in task pane once to complete consent, then retry.'
        : `Auth failed: ${String((err as Error)?.message ?? err).slice(0, 100)}`);
      return;
    }

    const apiBase = (import.meta.env.VITE_API_BASE_URL as string) ?? '';
    const isLocalhost = apiBase.startsWith('http://localhost') || apiBase.startsWith('http://127.0.0.1');
    if (!apiBase.startsWith('https://') && !isLocalhost) {
      throw new Error('VITE_API_BASE_URL must use HTTPS in non-localhost environments.');
    }

    // Step 1: Read metadata (no binary content yet) and call /prepare
    log.info('syncToTestApi: reading item metadata');
    const [bodyText, toList, ccList] = await Promise.all([
      getBody(item),
      getRecipients(item.to),
      getRecipients(item.cc),
    ]);
    const officeAttachments = item.attachments ?? [];

    log.info('syncToTestApi: calling /prepare', { attachmentCount: officeAttachments.length });
    const prepareRes = await fetch(`${apiBase}/api/submissions/prepare`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
      body: JSON.stringify({
        attachments: officeAttachments.map((a) => ({
          name: a.name,
          mimeType: a.contentType,
          sizeBytes: a.size,
        })),
      }),
    });

    if (!prepareRes.ok) {
      const detail = await prepareRes.text().catch(() => '');
      log.error('syncToTestApi: /prepare failed', { status: prepareRes.status, detail: detail.slice(0, 200) });
      await notify(item, 'error', `Prepare failed ${prepareRes.status}${detail ? ': ' + detail.slice(0, 80) : '.'}`);
      return;
    }

    const { submissionId, attachments: uploadSlots }: PrepareResponse = await prepareRes.json();
    log.info('syncToTestApi: prepared', { submissionId, slots: uploadSlots.length });

    // Step 2: Fetch each attachment binary and PUT directly to blob
    if (uploadSlots.length > 0) {
      log.info('syncToTestApi: uploading attachments directly to blob');
      await Promise.all(
        officeAttachments.map(async (a, i) => {
          const slot = uploadSlots[i];
          const content = await getAttachmentContent(item, a.id);
          await uploadAttachmentToBlob(slot.uploadUrl, content.content, a.contentType, a.name);
          log.info('syncToTestApi: attachment uploaded', { name: a.name, blobPath: slot.blobPath });
        })
      );
    }

    // Step 3: POST metadata + blobPaths to /api/submissions
    const payload: SubmissionPayload = {
      submissionId,
      messageId: item.itemId,
      correlationId: crypto.randomUUID(),
      sender: item.from?.emailAddress ?? '',
      recipients: [...toList, ...ccList],
      subject: item.subject ?? '',
      timestamp: item.dateTimeCreated?.toISOString() ?? new Date().toISOString(),
      bodyText,
      attachments: uploadSlots.map((slot: AttachmentUploadSlot) => ({
        name: slot.name,
        mimeType: slot.mimeType,
        sizeBytes: slot.sizeBytes,
        blobPath: slot.blobPath,
      })),
    };

    log.info('syncToTestApi: posting metadata to /api/submissions', {
      correlationId: payload.correlationId,
      sender: payload.sender,
      recipientCount: payload.recipients.length,
      attachmentCount: payload.attachments.length,
      bodyLength: payload.bodyText.length,
    });

    const controller = new AbortController();
    const timerId = setTimeout(() => controller.abort(), TIMEOUT_MS);
    let response: Response;
    try {
      response = await fetch(`${apiBase}/api/submissions`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
        body: JSON.stringify(payload),
        signal: controller.signal,
      });
    } finally {
      clearTimeout(timerId);
    }

    if (!response.ok) {
      const detail = await response.text().catch(() => '');
      log.error('syncToTestApi: server error', { status: response.status, detail: detail.slice(0, 200) });
      await notify(item, 'error', response.status === 401
        ? 'Server rejected token (401). Check App Registration audience config.'
        : `Server error ${response.status}${detail ? ': ' + detail.slice(0, 80) : '.'}`);
      return;
    }

    const result: SubmissionResponse = await response.json();
    log.info('syncToTestApi: success', { submissionId: result.submissionId, correlationId: result.correlationId });
    await notify(item, 'info', `Submitted — reference: ${result.submissionId}`);
  } catch (err: unknown) {
    const isTimeout = (err as { name?: string })?.name === 'AbortError';
    log.error('syncToTestApi: unexpected error', { timeout: isTimeout, error: String((err as Error)?.message ?? err) });
    await notify(item, 'error', isTimeout
      ? 'Request timed out. Check network or function URL.'
      : `Unexpected error: ${String((err as Error)?.message ?? err).slice(0, 100)}`);
  } finally {
    event.completed();
  }
}

// ── downloadPayload ───────────────────────────────────────────────────────────
// Builds a ZIP client-side from fresh Outlook attachment content, uploads to
// blob under downloads/, and opens the 1-hour SAS URL in the system browser.
// openBrowserWindow is the right API for Win32 and New Outlook (hidden webview can't
// do window.open for downloads). In some OWA environments it isn't exposed in the
// FunctionFile iframe, so fall back to window.open — OWA is a real browser context.

async function downloadPayload(event: Office.AddinCommands.Event): Promise<void> {
  const item = Office.context.mailbox.item as Office.MessageRead | null;
  if (!item) { event.completed(); return; }

  log.info('downloadPayload: started', { subject: item.subject });
  await notify(item, 'info', 'Preparing download…');

  try {
    let token: string;
    try {
      log.info('downloadPayload: acquiring NAA token');
      const msalApp   = await getMsalApp();
      const loginHint = Office.context.mailbox?.userProfile?.emailAddress;
      const result    = await Promise.race([
        msalApp.ssoSilent({ scopes: [API_SCOPE], loginHint }),
        new Promise<never>((_, reject) =>
          setTimeout(() => reject(new Error('NAA token timed out (10s).')), 10_000)
        ),
      ]);
      token = result.accessToken;
    } catch (err: unknown) {
      const isInteraction = err instanceof InteractionRequiredAuthError;
      await notify(item, 'error', isInteraction
        ? 'Sign-in required — open the add-in task pane once to complete consent, then retry.'
        : `Auth failed: ${String((err as Error)?.message ?? err).slice(0, 100)}`);
      return;
    }

    const apiBase = (import.meta.env.VITE_API_BASE_URL as string) ?? '';

    log.info('downloadPayload: building ZIP');
    const { zipBlob } = await buildZip(item);
    const contentBase64 = await blobToBase64(zipBlob);
    log.info('downloadPayload: ZIP ready', { sizeBytes: zipBlob.size });

    const controller = new AbortController();
    const timerId = setTimeout(() => controller.abort(), TIMEOUT_MS);
    let response: Response;
    try {
      response = await fetch(`${apiBase}/api/downloads`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
        body: JSON.stringify({ contentBase64, filename: 'submission.zip' }),
        signal: controller.signal,
      });
    } finally {
      clearTimeout(timerId);
    }

    if (!response.ok) {
      const detail = await response.text().catch(() => '');
      log.error('downloadPayload: upload failed', { status: response.status, detail: detail.slice(0, 200) });
      await notify(item, 'error', `Upload failed ${response.status}${detail ? ': ' + detail.slice(0, 80) : '.'}`);
      return;
    }

    const { downloadUrl } = await response.json() as DownloadResponse;
    log.info('downloadPayload: got SAS URL, opening browser');

    if (typeof (Office.context.ui as { openBrowserWindow?: unknown }).openBrowserWindow === 'function') {
      Office.context.ui.openBrowserWindow(downloadUrl);
    } else {
      window.open(downloadUrl, '_blank');
    }
    await notify(item, 'info', 'Download ready — opening in browser.');
  } catch (err: unknown) {
    const isTimeout = (err as { name?: string })?.name === 'AbortError';
    log.error('downloadPayload: failed', { timeout: isTimeout, error: String((err as Error)?.message ?? err) });
    await notify(item, 'error', isTimeout
      ? 'Request timed out (10s). Check network or function URL.'
      : `Download failed: ${String((err as Error)?.message ?? err).slice(0, 100)}`);
  } finally {
    event.completed();
  }
}

// ── Helpers ───────────────────────────────────────────────────────────────────

async function uploadAttachmentToBlob(
  uploadUrl: string,
  contentBase64: string,
  mimeType: string,
  filename: string,
): Promise<void> {
  const binary = atob(contentBase64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) {
    bytes[i] = binary.charCodeAt(i);
  }
  const response = await fetch(uploadUrl, {
    method: 'PUT',
    headers: {
      'x-ms-blob-type': 'BlockBlob',
      'Content-Type': mimeType,
      'Content-Disposition': `attachment; filename="${filename}"`,
    },
    body: bytes.buffer,
  });
  if (!response.ok) {
    const detail = await response.text().catch(() => '');
    throw new Error(`Blob PUT failed ${response.status}: ${detail.slice(0, 200)}`);
  }
}

interface ZipResult {
  zipBlob: Blob;
}

async function buildZip(item: Office.MessageRead): Promise<ZipResult> {
  const [bodyText, toList, ccList] = await Promise.all([
    getBody(item),
    getRecipients(item.to),
    getRecipients(item.cc),
  ]);

  log.info('buildZip: fetching attachment contents', { count: (item.attachments ?? []).length });
  const attachments: AttachmentWithContent[] = await Promise.all(
    (item.attachments ?? []).map(async (a) => {
      try {
        const content = await getAttachmentContent(item, a.id);
        log.info('buildZip: attachment fetched', { name: a.name, sizeBytes: a.size });
        return { name: a.name, mimeType: a.contentType, sizeBytes: a.size, contentBase64: content.content };
      } catch (err) {
        log.warn('buildZip: attachment fetch failed', { name: a.name, error: String((err as Error)?.message ?? err) });
        return { name: a.name, mimeType: a.contentType, sizeBytes: a.size, contentBase64: '' };
      }
    })
  );

  const zip = new JSZip();
  zip.file('metadata.json', JSON.stringify({
    sender: item.from?.emailAddress ?? '',
    recipients: [...toList, ...ccList],
    subject: item.subject ?? '',
    timestamp: item.dateTimeCreated?.toISOString() ?? new Date().toISOString(),
    bodyText,
    attachments: attachments.map(({ name, mimeType, sizeBytes }) => ({ name, mimeType, sizeBytes })),
  }, null, 2));

  const folder = zip.folder('attachments')!;
  for (const att of attachments) {
    if (att.contentBase64) {
      folder.file(att.name, att.contentBase64, { base64: true });
    }
  }

  const zipBlob = await zip.generateAsync({ type: 'blob' });
  log.info('buildZip: ZIP ready', { sizeBytes: zipBlob.size });
  return { zipBlob };
}

function notify(
  item: Office.MessageRead,
  type: 'info' | 'error',
  message: string
): Promise<void> {
  return new Promise((resolve) => {
    const isFinal = type === 'info' && (message.startsWith('Submitted') || message.startsWith('Download ready'));
    const data: Office.NotificationMessageDetails =
      type === 'error'
        ? {
            type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
            message: message.slice(0, 150),
          }
        : {
            type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
            message: message.slice(0, 150),
            icon: 'Icon.16x16',
            persistent: isFinal,
          };
    item.notificationMessages.replaceAsync(NOTIFICATION_KEY, data, () => resolve());
  });
}

function getBody(item: Office.MessageRead): Promise<string> {
  return new Promise((resolve) => {
    item.body.getAsync(Office.CoercionType.Text, {}, (res) => {
      resolve(
        res.status === Office.AsyncResultStatus.Succeeded ? (res.value ?? '') : ''
      );
    });
  });
}

function getRecipients(recipients: Office.EmailAddressDetails[] | undefined): Promise<string[]> {
  return Promise.resolve((recipients ?? []).map((r) => r.emailAddress));
}

function getAttachmentContent(item: Office.MessageRead, attachmentId: string): Promise<Office.AttachmentContent> {
  return new Promise((resolve, reject) => {
    item.getAttachmentContentAsync(attachmentId, (res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        resolve(res.value);
      } else {
        reject(new Error(res.error?.message ?? 'Failed to read attachment'));
      }
    });
  });
}

function blobToBase64(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve((reader.result as string).split(',')[1]);
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(blob);
  });
}
