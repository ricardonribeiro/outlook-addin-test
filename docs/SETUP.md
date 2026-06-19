# SETUP.md — Manual Setup (the non-code steps you do yourself)

Claude Code builds the application **and the Terraform**. **You** do everything in this file — the Entra ID setup that can't/shouldn't be automated, plus running Terraform and deploying the function code. Work through it in order.

> Assumption: you have your own Azure subscription and an M365 tenant you control (or a free Microsoft 365 Developer tenant with test mailboxes).

---

## Step 0 — Prerequisites (one-time, on your machine)

- **Azure CLI** installed and logged in: `az login`
- **Terraform** (v1.6+)
- **Azure Functions Core Tools** v4: `npm i -g azure-functions-core-tools@4`
- **Node.js** LTS (v20) and npm — for the **add-in** (Office.js / TypeScript / Vite).
- **uv** + **Python 3.11** — for the **functions** (Azure Functions, Python v2 model). Install uv: `curl -LsSf https://astral.sh/uv/install.sh | sh`. uv will manage the Python version and the virtual environment.
- **A Microsoft 365 tenant you control.** No dev tenant? Sign up for the free Microsoft 365 Developer Program for a sandbox tenant with mailboxes.
- **Office add-in dev certs**: `npx office-addin-dev-certs install` (Office add-ins require HTTPS even locally).

---

## Step 1 — Create the App Registration (Entra ID) — THE IMPORTANT ONE

This is the identity for both the add-in and the function. It is the single most important manual step and the one that most closely mirrors the real NorthStandard requirement. It is **not** in Terraform on purpose — the SSO config below is fiddly in the portal and best done once by hand for a test.

**Azure Portal -> Microsoft Entra ID -> App registrations -> New registration:**

1. **Name**: `outlook-addin-test`
2. **Supported account types**: "Accounts in this organizational directory only" (single tenant).
3. **Redirect URI**: leave blank for now.
4. **Register.** Copy the **Application (client) ID** and **Directory (tenant) ID**.

### 1a. Expose an API (what SSO needs)
- **Expose an API** -> set the **Application ID URI**. For SSO, Microsoft's pattern is `api://<host>/<client-id>`, where **`<host>` is the domain your add-in files are served from** — the same URL you put in the manifest's `<SourceLocation>` / `AppDomains`. It is not arbitrary: Microsoft requires the App ID URI host to match the add-in's hosting domain.
  - **Local dev**: the add-in runs on the Vite dev server at `https://localhost:3000`, so `<host>` is `localhost:3000` -> use `api://localhost:3000/<client-id>`.
  - **Deployed**: once the add-in is hosted (same Function App, or a static host like Azure Static Web Apps), `<host>` is that domain, e.g. `api://<your-addin-host>.azurestaticapps.net/<client-id>`.
  - Because the host differs between local and deployed, add the deployed host as a **second** Application ID URI when you deploy (the field accepts multiple), or re-edit this value then.
- Add a scope:
  - **Scope name**: `access_as_user`
  - **Who can consent**: Admins and users
  - **Display name / description**: "Access the test add-in as the signed-in user"
  - **State**: Enabled
- Save.

### 1b. Authorise the Office host applications (required for Office SSO)
Under **Expose an API -> Authorized client applications**, add each of these and tick your `access_as_user` scope. These are the Office hosts allowed to request a token silently — without them `getAccessToken()` fails:

```
d3590ed6-52b3-4102-aeff-aad2292ab01c   (Microsoft Office desktop)
ea5a67f6-b6f3-4338-b240-c655ddc3cc8e   (Microsoft Office, alt)
57fb890c-0dab-4253-a5e0-7188c88b2bb4   (Office on the web)
08e18876-6177-487e-b8b5-cf950c1e598c   (Office on the web, SharePoint)
bc59ab01-8403-45c6-8796-ac3ef710b3e3   (Outlook on the web)
93d53678-613d-4013-afc1-62e9e444a0a5   (Office on the web, other)
```

### 1c. API permissions
- **API permissions**: leave the default `Microsoft Graph -> User.Read` (delegated).
- Add **My APIs -> outlook-addin-test -> access_as_user** (delegated).
- **Grant admin consent for {your tenant}** (you can, it's your tenant).

### 1d. SPA redirect (for the Office.js SSO / NAA flow)
- **Authentication -> Add a platform -> Single-page application**.
- Add `https://localhost:3000/commands.html` and later your deployed add-in URL. (This is the function-command host file, not a task pane.)

**Record these — they become Terraform variables and add-in config:**
- `TENANT_ID` = Directory (tenant) ID
- `CLIENT_ID` = Application (client) ID
- `API_AUDIENCE` = the App ID URI you set above (e.g. `api://localhost:3000/<client-id>` for local dev). The token validation accepts either the full App ID URI or the bare `<client-id>` as the audience — match whichever you configure.

---

## Step 2 — Fill in Terraform variables

In `/infra`, copy `terraform.tfvars.example` to `terraform.tfvars` and fill in:
- `tenant_id`, `api_audience` (App ID URI), `client_id`
- `add_in_origin` (e.g. `https://localhost:3000` for first run)
- `location` (e.g. `uksouth`), `name_prefix` (e.g. `oaddintest`)
- **Mandatory tagging variables** (Indicium tenant policy — see below):
  - `owner_email` — your `firstname.lastname@indicium.ai`. **No default; you must supply it.** Terraform will reject anything not ending in `@indicium.ai`.
  - `end_date` — defaults to `170826` (17 Aug 2026, ddmmyy). This is the 2-month cap the policy requires for a `Testing` deployment. Bring it nearer if you'll finish sooner; do not push it past the cap.
  - `project` — defaults to `outlook-addin-test-harness`. Change if you prefer.
  - `description` — optional one-liner.

> **About the tags (Indicium mandatory tagging policy).** Every persistent resource and the resource group must carry six underscore-prefixed tags. Two are fixed in the Terraform and you can't set them, because the policy pins them for this kind of deployment:
> - `_purpose = Testing` (this is a test harness).
> - `_business_criticality = Low` (required by the Testing purpose; also accurate).
>
> The other four are the variables above (`_owner_email`, `_end_date`, `_project`, `_description`). The Terraform has validation rules baked in, so a `terraform plan` will fail fast if, for example, `end_date` is `None` or `owner_email` isn't an Indicium address — that's intentional, it stops you deploying non-compliant resources. If you ever change `_purpose` away from `Testing`, re-check the policy: `Client Development` allows up to 6 months, `Internal Development` up to 3, `Client Service` requires criticality at or above Medium, etc.

> `terraform.tfvars.example` is one of the files **Claude Code generates** when it builds the repo (it's in the `/infra` folder of the repo structure defined in the build prompt). If you haven't run Claude Code against the build prompt yet, do that first — this file won't exist until then.

`terraform.tfvars` is gitignored — never commit it.

---

## Step 3 — Provision infrastructure with Terraform

```
cd infra
terraform init
terraform plan -out tfplan
terraform apply tfplan
```

This creates: resource group, Service Bus namespace + `enquiry-queue` (with duplicate detection), storage account, Log Analytics + Application Insights, the Function App with system-assigned Managed Identity, and the Service Bus Data Sender/Receiver role assignments. Every persistent resource and the resource group is tagged per the Indicium mandatory tagging policy (see Step 2).

Then read the outputs:
```
terraform output
```
Note the **Function App hostname** (your `/api/enquiries` base URL) and the **Service Bus FQDN**.

---

## Step 4 — Deploy the function code

Terraform provisions infra only; function code is deployed separately. The functions are Python managed by uv, and the Azure remote build expects a `requirements.txt`, so export it from uv before publishing:

```
cd ../functions
uv sync                                   # create the venv + install deps locally
uv export --no-hashes -o requirements.txt # Azure's Oryx build installs from this
func azure functionapp publish <function-app-name-from-tf-output>
```

---

## Step 5 — Grant YOUR user Service Bus roles (for local dev)

`DefaultAzureCredential` uses your `az login` identity when running locally, so your own account needs the data roles on the namespace:

```
az role assignment create \
  --assignee <your-user-object-id-or-upn> \
  --role "Azure Service Bus Data Sender" \
  --scope <service-bus-namespace-resource-id>

az role assignment create \
  --assignee <your-user-object-id-or-upn> \
  --role "Azure Service Bus Data Receiver" \
  --scope <service-bus-namespace-resource-id>
```

(The namespace resource ID is a Terraform output; or `az servicebus namespace show`.)

---

## Step 6 — Point the add-in at the deployed function

1. Put the Function App HTTPS URL into the add-in config (Claude Code documents exactly where).
2. Edit `manifest.xml`:
   - `<SourceLocation>` and `AppDomains` -> your add-in host URL.
   - `<WebApplicationInfo>` `<Id>` -> your `CLIENT_ID`, `<Resource>` -> your App ID URI.
3. Rebuild/redeploy the add-in (or keep on localhost for the first test).

---

## Step 7 — Sideload the add-in into Outlook (no admin portal needed)

For a personal test you do NOT need the M365 Integrated Apps portal. Sideload directly:

- **Outlook on the web / new Outlook on Windows**: Settings -> search "add-in" -> My add-ins -> Add a custom add-in -> Add from file -> upload `manifest.xml`.
- Open or select an email -> your **"Sync to Test API"** button appears in the ribbon / menu bar. Clicking it runs immediately (no panel opens).

> The centralised path (M365 Admin Centre -> Integrated Apps -> Upload custom apps) is the real NorthStandard route — not needed for the basic test.

---

## Step 8 — Test the end-to-end event-driven flow

1. Open an email in your test mailbox.
2. Click **Sync to Test API**. Consent on first run (or silent if admin consent granted).
3. An Outlook **notification message** appears on the email showing success with a fake `ENQ-xxxx` reference ID — this confirms the HTTP function authenticated and enqueued. (There is no task pane; the button runs the function directly and reports via notifications.)
4. Confirm the enquiry-processor picked the message off the queue. In Application Insights (Logs / transaction search) or via live logs:
   ```
   func azure functionapp logstream <function-app-name>
   ```
   You should see BOTH functions log the same `correlationId` — the enquiry-receiver on enqueue, the enquiry-processor on dequeue. That round trip is the proof the event-driven loop works.
5. Optional: check the Service Bus queue metrics in the portal to see messages in/out.

---

## Step 9 — Tear down when done (avoid charges)

```
cd infra
terraform destroy
```

Then delete the `outlook-addin-test` App Registration manually in Entra ID (it isn't managed by Terraform).

> Reminder: the `_end_date` tag (default `170826`) is the date you committed to under the tagging policy. Either tear down by then, or bump `end_date` in `terraform.tfvars` and re-apply to extend it (staying within the 2-month Testing cap). Don't just let it lapse — the tag is what lets Indicium audit and reclaim stale resources.

---

## Summary — the non-code setup checklist

| # | Task | Where | One-time? |
|---|---|---|---|
| 0 | Install tooling + dev certs + get M365 dev tenant | Local machine | Yes |
| 1 | **Create + configure App Registration (SSO, scope, authorised Office clients, consent)** | Entra ID portal | Yes |
| 2 | Fill in terraform.tfvars | Local edit | Yes |
| 3 | `terraform init/plan/apply` | Terraform | Per infra change |
| 4 | `func azure functionapp publish` | Functions Core Tools | Per code change |
| 5 | Grant your user Service Bus data roles | az CLI | Yes |
| 6 | Put function URL + client ID into add-in config + manifest | Local edit | Per deploy |
| 7 | Sideload manifest into Outlook | Outlook UI | Per manifest change |
| 8 | Test end to end (watch both functions log same correlationId) | Outlook + App Insights | — |
| 9 | Tear down | terraform destroy + Entra ID | At end |

**The two things that matter most:** Step 1 (the App Registration — exactly what you'll ask NorthStandard's infra team for) and Step 8 (seeing both functions log the same correlationId, which proves the add-in -> enquiry-receiver -> Service Bus -> enquiry-processor event-driven loop end to end).
