# Claude Code Build Prompt — Outlook Add-In Test Harness (Event-Driven + Terraform)
---

## 0. What you are building (read this first)

A **minimal but real Outlook add-in** that proves out an **event-driven enquiry pattern** end to end, using Azure Service Bus to mock the agentic/event layer.

> **Terminology used in this doc (kept deliberately plain):**
> - **enquiry** = the unit of work created from a broker email. An underwriter opens an email pack from a broker and turns it into an enquiry — the record the business then works on. In this test harness the email is captured and sent on to become an enquiry.
> - **enquiry-receiver** = the HTTP-triggered function the add-in calls. It receives the email payload, authenticates it, and puts it on the queue.
> - **enquiry-queue** = the Azure Service Bus queue the message sits on, waiting to be processed. This stands in for the real event-driven agentic layer.
> - **enquiry-processor** = the queue-triggered function that picks the message off the queue and processes it (here: just logs it; in production an agent would run).
> - **enquiry reference ID** (`enquiryRefId`) = a fake ID the receiver generates so the underwriter has something to confirm the email was accepted.

The flow:

1. An underwriter opens or selects an email in Outlook.
2. A custom button (**"Sync to Test API"**) appears in the Outlook **ribbon / menu bar**. There is **NO task pane** — clicking the button runs a function directly (an ExecuteFunction command), and feedback is shown via an Outlook **notification message** on the email.
3. Clicking it acquires an **Entra ID identity token** for the signed-in user (SSO).
4. The add-in extracts the email's metadata, body, and attachment info via **Office.js**.
5. It POSTs that payload — with the token as a Bearer header — to the **enquiry-receiver**, an HTTP-triggered Azure Function I host in **my own Azure subscription**.
6. The receiver **validates the token**, enforces the payload contract, stamps a server-side receipt, and **writes the message to the enquiry-queue** (Azure Service Bus). It returns a fake **enquiry reference ID** to the add-in immediately (fire-and-forget from the client's perspective).
7. The **enquiry-processor**, a second Service Bus-triggered Azure Function, consumes the queue message and **logs it** — this proves the event-driven loop end to end.
8. The add-in shows the underwriter a success confirmation with that reference ID, or a meaningful error, as an Outlook **notification message** on the email (no task pane).

This is a **throwaway test harness** to validate deployment, auth, event-driven, and infra requirements — not production code. Favour a working end-to-end skeleton over completeness.

### Architecture decision already made (do NOT change)
The add-in **POSTs to the Function App, which writes to Service Bus.** The add-in does **NOT** write to Service Bus directly. Rationale: the add-in is an untrusted client; writing directly would require embedding a SAS token/connection string (credential leak) or per-user Service Bus RBAC (no validation boundary, tight coupling). The Function is the trust boundary: it authenticates, validates the contract, stamps server-side fields, and is the only thing holding Service Bus credentials (via Managed Identity — see section 6).

---

## 1. Architecture

```
+---------------------+   1. SSO token       +----------------------+
|  Outlook client     | -------------------> |  Entra ID (my tenant)|
|  (Office.js add-in) | <------------------- |                      |
+---------+-----------+   identity token     +----------------------+
          |
          | 2. POST /api/enquiries  (Bearer token + EnquiryPayload)
          v
+----------------------------------+
|  enquiry-receiver             |   <- TRUST BOUNDARY
|  (HTTP-triggered function)       |   - validates JWT (issuer/audience/signature)
|                                  |   - enforces payload contract
|                                  |   - stamps receivedAt + enquiryRefId
|                                  |   - returns { enquiryRefId } immediately (202)
|        |  3. send message        |
|        v                         |
|  +----------------------------+  |
|  | enquiry-queue           |  |   <- mocks the event-driven agentic layer
|  | (Azure Service Bus queue)  |  |
|  +-----------+----------------+  |
+--------------+-------------------+
               | 4. queue trigger
               v
+----------------------------------+
|  enquiry-processor            |
|  (Service Bus-triggered function)|   - reads message off the queue
|                                  |   - logs payload + correlationId
|                                  |   - (stub) downstream agent would run here
+----------------------------------+
```

Connection between Function and Service Bus uses **Managed Identity**, not connection strings.

---

## 2. Tech stack (use these unless I say otherwise)

- **Add-in**: Office.js, plain TypeScript + Vite. Keep it lean.
- **Functions**: **Azure Functions, Python, v2 programming model (decorator-based), managed with `uv`.** Python 3.11. Two functions in one Function App:
  - `enquiry-receiver` — HTTP trigger, `POST /api/enquiries`.
  - `enquiry-processor` — Service Bus queue trigger on `enquiry-queue`.
  - Use `uv` for dependency management (`uv init`, `uv add`, `uv run`). A `pyproject.toml` (not `requirements.txt`) is the source of truth, BUT also generate a `requirements.txt` via `uv export` because the Azure Functions remote build (Oryx) expects one at publish time — document this in the README.
- **Auth**: Office.js SSO via `Office.auth.getAccessToken()`. The enquiry-receiver validates the token via **PyJWT** (`jwt.PyJWKClient` to fetch signing keys) against my tenant's JWKS endpoint.
- **Messaging**: Azure Service Bus (queue). Function -> Service Bus via the **`azure-servicebus`** SDK using **`DefaultAzureCredential`** from **`azure-identity`** (Managed Identity in Azure, az-CLI creds locally).
- **Infra**: **Terraform** (see section 9). All Azure resources as code.
- **Local dev**: end to end locally over HTTPS. Azure Functions Core Tools (`uv run func start`) for the Python functions, Vite dev server for the add-in, `office-addin-dev-certs` for localhost HTTPS. For local Service Bus, default to a real Service Bus namespace in Azure (cheapest reliable path) — `[DECISION]` ask me if I'd prefer the local emulator, otherwise use a real dev namespace.

---

## 3. Repo structure to create

```
/
|-- README.md                  <- run locally + deploy (you write this)
|-- SETUP.md                   <- I provide this; reference it, don't overwrite
|-- .gitignore
|-- .env.example               <- all config keys, no secrets
|
|-- /addin
|   |-- manifest.xml           <- ExecuteFunction button, NO taskpane
|   |-- package.json
|   |-- vite.config.ts
|   |-- tsconfig.json
|   \-- /src
|       |-- commands.html      <- loads Office.js + commands.ts in the function runtime
|       \-- commands.ts        <- command handler: token + extract + POST + notify
|
|-- /functions
|   |-- pyproject.toml          <- uv project + deps (source of truth)
|   |-- uv.lock                 <- generated by uv
|   |-- requirements.txt        <- generated via `uv export`, needed by Azure remote build
|   |-- host.json
|   |-- local.settings.json.example
|   |-- function_app.py         <- v2 model entry point: registers both functions
|   \-- /shared
|       |-- auth.py             <- JWT validation (PyJWT)
|       |-- service_bus.py      <- sender helper (azure-servicebus + Managed Identity)
|       \-- models.py           <- payload contract (Python dataclasses / TypedDict)
|
\-- /infra
    |-- main.tf
    |-- variables.tf
    |-- outputs.tf
    |-- providers.tf
    |-- terraform.tfvars.example
    \-- README.md              <- terraform init/plan/apply order + post steps
```

---

## 4. The payload contract (single source of truth)

The wire format is **JSON**. Because the add-in is TypeScript and the functions are Python, the contract is expressed in BOTH languages and they must stay in lockstep. The TypeScript interface below is the canonical description; mirror it exactly in Python.

**Add-in side** — define in `/addin/src/types.ts`:

```typescript
interface EnquiryPayload {
  messageId: string;        // Office.context.mailbox.item.itemId — idempotency key
  correlationId: string;    // UUID generated client-side at click time
  sender: string;
  recipients: string[];     // to + cc
  subject: string;
  timestamp: string;        // ISO 8601, email sent date
  bodyText: string;
  attachments: {
    name: string;
    mimeType: string;
    sizeBytes: number;
    // METADATA ONLY for the test harness. Comment where base64 content would go.
  }[];
}

interface EnquiryResponse {
  enquiryRefId: string;     // fake e.g. "ENQ-" + short uuid
  receivedAt: string;       // ISO 8601, stamped server-side
  correlationId: string;    // echoed back
}

// What actually goes on the queue (server-enriched by the receiver)
interface EnquiryQueueMessage extends EnquiryPayload {
  enquiryRefId: string;
  receivedAt: string;
  authenticatedUpn: string; // extracted from the validated token
}
```

**Functions side** — mirror in `/functions/shared/models.py` (use `dataclass` or `TypedDict`; field names must match the JSON keys above exactly, i.e. `messageId`, `correlationId`, etc. — keep camelCase on the wire even though Python convention is snake_case, OR serialise explicitly. Pick camelCase-on-the-wire and document it). Validate incoming JSON against this shape and reject malformed payloads with 400.

Keep attachments to **metadata only**. Leave a clearly commented stub for content extraction.

---

## 5. Behavioural requirements (mirror the NorthStandard principles)

1. **UW-initiated trigger** — button only, reads only the selected/open item.
2. **Auth before send** — never POST before a valid token; never enqueue before token validated server-side.
3. **Progress feedback** — since there is no task pane, show an "in progress" Outlook notification message immediately on click, then replace it with the final result. (No button-disable/spinner — that's a task-pane pattern.)
4. **Idempotency-aware** — `correlationId` per click; `messageId` as idempotency key; both logged by both functions. Service Bus supports duplicate detection — enable it on the queue via Terraform and set the message's `messageId` property to the email `messageId` so the platform dedupes. Comment where an app-level dedupe store would otherwise go.
5. **Fire and forget (client side)** — the enquiry-receiver returns as soon as the message is enqueued; it does NOT wait for the enquiry-processor. The client awaits only the enqueue ack.
6. **Meaningful errors** — distinguish (a) auth failure, (b) network/timeout, (c) enquiry-receiver rejected request (non-2xx). Different messages each.
7. **Timeout** — client aborts after 10s.
8. **Success confirmation** — show returned `enquiryRefId` in a success notification message.
9. **HTTPS only** — refuse non-HTTPS API base URL in config.
10. **Correlation logging** — both functions log `correlationId`, `messageId`, outcome, timestamp as structured JSON to stdout (captured by Application Insights).

---

## 6. Function specifics

Both functions live in `function_app.py` (v2 model: decorate handlers with `@app.route(...)` and `@app.service_bus_queue_trigger(...)`). Shared logic in `/functions/shared/`.

### enquiry-receiver (HTTP trigger)
- Route: `POST /api/enquiries`. Also `GET /api/health` (anonymous, `auth_level=ANONYMOUS`) returning `{ status: "ok" }`.
- **JWT validation** (`shared/auth.py`, using **PyJWT**): use `jwt.PyJWKClient("https://login.microsoftonline.com/{TENANT_ID}/discovery/v2.0/keys")` to fetch the signing key, then `jwt.decode(...)` validating signature, issuer (handle both v1 `https://sts.windows.net/{tid}/` and v2 `https://login.microsoftonline.com/{tid}/v2.0`), and audience (the App Registration client ID / App ID URI). Return HTTP 401 on any failure. Extract `preferred_username`/`upn` claim for `authenticatedUpn`.
- On success: validate payload shape, build the queue message (stamp `receivedAt`, generate `enquiryRefId` with `uuid.uuid4()`, copy `authenticatedUpn` from token claims), send to the enquiry-queue, return the `EnquiryResponse` JSON with 202 Accepted.
- **CORS**: allow the add-in origin from env.

### enquiry-processor (Service Bus queue trigger)
- Binds to `enquiry-queue`. On message: parse, log the full enriched message + correlation/idempotency keys, then `# STUB: downstream agent (the NorthStandard "Submission Output Agent") would run here`.
- Let exceptions raise so the platform's retry/dead-letter behaviour is exercised (proves event-driven semantics).

### Service Bus access — Managed Identity, not connection strings
- `shared/service_bus.py` uses `ServiceBusClient(fully_qualified_namespace, credential=DefaultAzureCredential())` from `azure-servicebus` + `azure-identity`.
- The queue trigger binding authenticates with identity too: configure the binding's connection setting as an identity-based connection (e.g. app setting `ServiceBusConnection__fullyQualifiedNamespace`) rather than a connection string.
- In Azure: the Function App's system-assigned Managed Identity is granted `Azure Service Bus Data Sender` (used by enquiry-receiver) and `Azure Service Bus Data Receiver` (used by enquiry-processor) via Terraform.
- Locally: `DefaultAzureCredential` falls back to your `az login` credentials — document that the developer needs the Sender/Receiver role on the dev namespace too.

### Config (env / app settings)
`TENANT_ID`, `API_AUDIENCE`, `ALLOWED_CORS_ORIGIN`, `SERVICEBUS_FQDN` (e.g. `myns.servicebus.windows.net`), `SERVICEBUS_QUEUE_NAME`. No connection strings anywhere.

---

## 7. Add-in specifics

This add-in is a **function command (UI-less button)** — NOT a task pane. The button sits in the Outlook ribbon/menu bar and runs a JavaScript function directly when clicked. Feedback is shown via Outlook notification messages, not a panel.

- **Manifest**: XML format. `<Requirements>` -> `Mailbox` `MinVersion="1.5"`. `MessageRead` surface so the button shows when reading email.
  - Define the button as a control of `xsi:type="Button"` with an `<Action xsi:type="ExecuteFunction">` and a `<FunctionName>` (e.g. `syncToTestApi`) — NOT `ShowTaskpane`.
  - Add a `<Runtimes>` / function-file reference (`<FunctionFile>` in the VersionOverrides) pointing at `commands.html`, which loads `commands.ts`.
  - Place the button on the `MessageReadCommandSurface`.
  - Include `<WebApplicationInfo>` with my App Registration client ID + App ID URI scope for SSO.
- **commands.html**: a minimal HTML file whose only job is to load Office.js and the compiled `commands.ts`. It is never shown to the user — it hosts the function runtime.
- **commands.ts flow**:
  1. `Office.onReady()` then register the command with `Office.actions.associate("syncToTestApi", handler)`.
  2. Handler receives an `event` argument (the `Office.AddinCommands.Event`).
  3. Show an "in progress" notification via `Office.context.mailbox.item.notificationMessages.replaceAsync(...)` with an `InformationalMessage` (this replaces the disable-button + spinner pattern a task pane would use).
  4. `Office.auth.getAccessToken({ allowSignInPrompt: true })` — handle 13001/13002/13003 etc. **Note the caveat below about prompts in a UI-less command.**
  5. Read item fields (subject, from, to, cc, `body.getAsync`, `item.attachments`), build payload, `crypto.randomUUID()` correlationId.
  6. `fetch` POST with `AbortController` 10s timeout + Bearer header.
  7. Replace the notification with success (showing the `enquiryRefId`) or a specific error message.
  8. **Always call `event.completed()`** at the very end (success or failure) — the function runtime requires it, or the command hangs.
- **Feedback = notification messages only.** Use a stable notification key so the "in progress" message is replaced by the final result rather than stacking. Success = `InformationalMessage`; error = `ErrorMessage` (or informational with an error string — Outlook's persistent error type is limited, document the choice).

### SSO caveat for UI-less commands (call this out in the README)
`getAccessToken` can normally trigger an interactive sign-in/consent dialog. In a function command there is **no task pane to host that dialog**, so an interactive prompt may not surface cleanly. For this test that's fine because admin consent is granted up front (SETUP.md), making token acquisition **silent**. If a prompt is ever required, the documented fallback is to surface a notification telling the user to open the add-in's (optional) consent page once. Keep the happy path silent; just note the limitation.

---

## 8. README contents
- Prerequisites (Node for the add-in, **`uv` + Python 3.11 for the functions**, az CLI, Functions Core Tools v4, Terraform, dev certs).
- Run locally end to end (terraform apply for the dev namespace -> `uv run func start` in `/functions` -> Vite dev server in `/addin` -> sideload manifest).
- Point add-in at deployed function URL.
- Clear note: throwaway test harness, function-command (menu-bar button) model, feedback via notifications.
- The SSO caveat for UI-less commands (silent token works because admin consent is pre-granted).

---

## 9. Terraform (/infra) — you write this, I run it

Define ALL Azure resources as code. Use the `azurerm` provider. Resources:

- Resource group (tagged per the MANDATORY tagging section below — applies to the RG and every persistent resource).
- **Service Bus**: namespace (Standard SKU — needed for duplicate detection), queue `enquiry-queue` with **duplicate detection enabled**, sensible lock duration + max delivery count (so dead-lettering is observable).
- **Storage account** (required by the Function App).
- **Log Analytics workspace + Application Insights** (so function logs/traces are queryable — important for proving the loop).
- **Function App** (Linux, **Python 3.11, Functions v4 runtime**) on a **Consumption (Y1) plan** `[DECISION]` — ask if I want Premium; default Consumption for cost. Set the runtime stack to Python (in `azurerm_linux_function_app`, `site_config.application_stack.python_version = "3.11"`).
- **System-assigned Managed Identity** on the Function App.
- **Role assignments**: grant the Function App's identity `Azure Service Bus Data Sender` AND `Azure Service Bus Data Receiver` on the namespace (one app hosts both functions).
- **Function App settings**: wire `TENANT_ID`, `API_AUDIENCE`, `ALLOWED_CORS_ORIGIN`, `SERVICEBUS_FQDN`, `SERVICEBUS_QUEUE_NAME`, `APPLICATIONINSIGHTS_CONNECTION_STRING`. No Service Bus connection string — identity only. Also set `SCM_DO_BUILD_DURING_DEPLOYMENT=true` (so Oryx installs the Python deps from `requirements.txt` on publish) and the identity-based Service Bus connection setting for the queue trigger binding: `ServiceBusConnection__fullyQualifiedNamespace = <namespace>.servicebus.windows.net`.
- CORS on the Function App set to the add-in origin.

Terraform requirements:
- `variables.tf` with everything I must supply (tenant id, client id / audience, add-in origin, region, name prefix) — and `terraform.tfvars.example`.
- `outputs.tf` emitting: Function App default hostname (the `/api/enquiries` base URL), Service Bus namespace FQDN, App Insights name.
- `providers.tf` pinning `azurerm` to a recent 3.x/4.x version and requiring `features {}`.
- Function CODE is deployed SEPARATELY AFTER `terraform apply` — Terraform provisions infra, not code. For Python: `cd functions && uv export --no-hashes -o requirements.txt && func azure functionapp publish <name>` (the remote Oryx build installs from `requirements.txt`). Document clearly in `/infra/README.md`.
- Do NOT include the App Registration in Terraform — I create it by hand (SETUP.md) and pass its IDs as variables. Rationale: SSO app registrations need fiddly manual portal steps (authorising Office client IDs) better done once by hand for a test.

### MANDATORY tagging (Indicium tenant policy — do NOT omit)
Every persistent resource AND the resource group MUST carry the tags below. Each tag name starts with an underscore. Implement this so the tags can't be forgotten on any one resource:

- Define a single `tags` map (build it in `locals` from the tag variables below) and apply `tags = local.common_tags` to the resource group and EVERY persistent resource (Service Bus namespace, storage account, Log Analytics workspace, Application Insights, Function App, App Service Plan). Prefer also setting the `azurerm` provider's `default_tags` if the pinned provider version supports it, as a belt-and-braces second layer — but still set `tags` explicitly per resource so it works regardless of provider version.

The required tags and the values fixed for THIS deployment:

| Tag | Value for this deployment | How to wire it |
| --- | --- | --- |
| `_purpose` | `Testing` | Hardcode in `locals` — this harness is a test by definition. |
| `_business_criticality` | `Low` | Hardcode in `locals`. The `Testing` purpose REQUIRES `Low`, and it's accurate ("nobody would notice if it was gone"). |
| `_end_date` | `170826` (ddmmyy) | Variable `end_date`, defaulted to `170826`. The `Testing` purpose REQUIRES an end date no more than 2 months out; today is 17 Jun 2026 so 17 Aug 2026 is the latest valid value. Use the ddmmyy format — it's mandatory for Azure. Validate in the variable that it is not "None" (Testing forbids None).
| `_owner_email` | `[firstname.lastname@indicium.ai]` | Variable `owner_email`, no default — I MUST supply it. Add a validation that it ends in `@indicium.ai`. |
| `_project` | e.g. `outlook-addin-test-harness` | Variable `project`, defaulted to `outlook-addin-test-harness`. |
| `_description` | optional | Variable `description`, defaulted to a one-line explanation; optional per policy. |

Encode the policy interdependencies as Terraform variable `validation` blocks so a non-compliant `terraform plan` fails fast rather than deploying untagged or wrongly-tagged resources:
- `_purpose = "Testing"` => `_business_criticality` must be `Low` AND `_end_date` must not be `None`.
- `_end_date` (when not None) must match the `ddmmyy` regex `^[0-3][0-9][0-1][0-9][0-9]{2}$`.
- `owner_email` must match `^[^@]+@indicium\.ai$`.

Put `owner_email`, `end_date`, `project`, and `description` in `variables.tf` and list them in `terraform.tfvars.example`. `_purpose` and `_business_criticality` are fixed in `locals` (not user-editable) because the policy pins them for a Testing deployment.

`/infra/README.md` must give the exact order:
1. Create App Registration by hand (SETUP.md step 2).
2. `terraform init` -> `plan` -> `apply` with my tfvars.
3. Read outputs.
4. `cd functions && uv export --no-hashes -o requirements.txt && func azure functionapp publish <name>` to deploy function code.
5. Paste the function URL + client ID into the add-in config + manifest.
6. Grant my own user the Service Bus Data Sender/Receiver roles for local dev (az CLI command provided).

---

## 10. Rules of engagement for you (Claude Code)

- Build the whole skeleton so it runs end to end locally before polishing.
- Do NOT create Azure resources, App Registrations, or run `az`/`terraform`/`func` commands — those are my manual steps.
- Do NOT invent my tenant ID, client IDs, URLs, or resource names — use variables/placeholders and list every value I must supply in `.env.example`, `terraform.tfvars.example`, and `local.settings.json.example`.
- Where a `[DECISION]` tag appears and I haven't answered, use the stated default and note it in the README.
- No secrets in the repo. No Service Bus connection strings anywhere — Managed Identity only.
- Prefer a small number of well-commented files.
- When done, print a checklist of what I must do manually before it runs, cross-referencing SETUP.md and /infra/README.md.
