import { createNestablePublicClientApplication, InteractionRequiredAuthError, } from '@azure/msal-browser';
const NOTIFICATION_KEY = 'enquiry-sync';
const TIMEOUT_MS = 10000;
const CLIENT_ID = import.meta.env.VITE_CLIENT_ID;
const TENANT_ID = import.meta.env.VITE_TENANT_ID;
const API_SCOPE = `api://localhost:3000/${CLIENT_ID}/access_as_user`;
// createNestablePublicClientApplication auto-detects Office NAA context and
// uses the Office host as a token broker (NestedAppAuthController) when available,
// falling back to StandardController otherwise.x
let _msalApp = null;
async function getMsalApp() {
    if (_msalApp)
        return _msalApp;
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
});
async function syncToTestApi(event) {
    const item = Office.context.mailbox.item;
    if (!item) {
        event.completed();
        return;
    }
    await notify(item, 'info', 'Syncing to enquiry system…');
    try {
        // 1. Acquire token via NAA — ssoSilent uses the Office host broker; loginHint
        //    from the mailbox profile lets MSAL skip account selection silently.
        //    10s timeout guards against the broker never responding.
        let token;
        try {
            const msalApp = await getMsalApp();
            const loginHint = Office.context.mailbox?.userProfile?.emailAddress;
            const result = await Promise.race([
                msalApp.ssoSilent({ scopes: [API_SCOPE], loginHint }),
                new Promise((_, reject) => setTimeout(() => reject(new Error('NAA token timed out (10s). Check App Registration SPA platform and access_as_user scope.')), 10000)),
            ]);
            token = result.accessToken;
        }
        catch (err) {
            const msg = err instanceof InteractionRequiredAuthError
                ? 'Sign-in required — open the add-in task pane once to complete consent, then retry.'
                : `Auth failed: ${String(err?.message ?? err).slice(0, 100)}`;
            await notify(item, 'error', msg);
            return;
        }
        const apiBase = import.meta.env.VITE_API_BASE_URL ?? '';
        const isLocalhost = apiBase.startsWith('http://localhost') || apiBase.startsWith('http://127.0.0.1');
        if (!apiBase.startsWith('https://') && !isLocalhost) {
            throw new Error('VITE_API_BASE_URL must use HTTPS in non-localhost environments.');
        }
        // 2. Read item data concurrently
        const [bodyText, toList, ccList] = await Promise.all([
            getBody(item),
            getRecipients(item.to),
            getRecipients(item.cc),
        ]);
        const payload = {
            messageId: item.itemId,
            correlationId: crypto.randomUUID(),
            sender: item.from?.emailAddress ?? '',
            recipients: [...toList, ...ccList],
            subject: item.subject ?? '',
            timestamp: item.dateTimeCreated?.toISOString() ?? new Date().toISOString(),
            bodyText,
            attachments: (item.attachments ?? []).map((a) => ({
                name: a.name,
                mimeType: a.contentType,
                sizeBytes: a.size,
            })),
        };
        // 3. POST with 10s hard timeout
        const controller = new AbortController();
        const timerId = setTimeout(() => controller.abort(), TIMEOUT_MS);
        let response;
        try {
            response = await fetch(`${apiBase}/api/enquiries`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    Authorization: `Bearer ${token}`,
                },
                body: JSON.stringify(payload),
                signal: controller.signal,
            });
        }
        finally {
            clearTimeout(timerId);
        }
        if (!response.ok) {
            const detail = await response.text().catch(() => '');
            const msg = response.status === 401
                ? 'Server rejected token (401). Check App Registration audience config.'
                : `Server error ${response.status}${detail ? ': ' + detail.slice(0, 80) : '.'}`;
            await notify(item, 'error', msg);
            return;
        }
        const result = await response.json();
        await notify(item, 'info', `Submitted — reference: ${result.enquiryRefId}`);
    }
    catch (err) {
        const isTimeout = err?.name === 'AbortError';
        await notify(item, 'error', isTimeout
            ? 'Request timed out (10s). Check network or function URL.'
            : `Unexpected error: ${String(err?.message ?? err).slice(0, 100)}`);
    }
    finally {
        event.completed();
    }
}
// ── Helpers ───────────────────────────────────────────────────────────────────
function notify(item, type, message) {
    return new Promise((resolve) => {
        const isFinal = type === 'info' && message.startsWith('Submitted');
        const data = type === 'error'
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
function getBody(item) {
    return new Promise((resolve) => {
        item.body.getAsync(Office.CoercionType.Text, {}, (res) => {
            resolve(res.status === Office.AsyncResultStatus.Succeeded ? (res.value ?? '') : '');
        });
    });
}
function getRecipients(recipients) {
    return new Promise((resolve) => {
        if (!recipients) {
            resolve([]);
            return;
        }
        recipients.getAsync({}, (res) => {
            resolve(res.status === Office.AsyncResultStatus.Succeeded
                ? (res.value ?? []).map((r) => r.emailAddress)
                : []);
        });
    });
}
