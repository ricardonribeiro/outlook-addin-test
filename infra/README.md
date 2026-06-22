# Infra — Terraform

All Azure resources for the Outlook add-in test harness. Terraform provisions infrastructure only; function code is deployed separately.

## What this creates

| Resource | Purpose |
|---|---|
| Resource group | Container for all resources |
| Service Bus namespace (Standard) | Needed for duplicate detection |
| Service Bus queue `submission-queue` | Duplicate detection enabled; dead-lettering after 5 failures |
| Service Bus authorization rule (`func-rule`) | Send+listen connection string for the Function App (no Managed Identity RBAC) |
| Storage account | Functions host storage + blob storage for submissions/downloads |
| Blob container `submissions` | Submission ZIPs and download artifacts (30/7-day lifecycle expiry) |
| Blob container `deploymentpackage` | Flex Consumption deployment artifact |
| Log Analytics workspace | Backend for Application Insights |
| Application Insights | Function logs and correlationId traces |
| App Service Plan (Flex Consumption `FC1`) | Hosts the Function App |
| Function App (Flex Consumption, Python 3.13, v4) | Hosts both functions |
| Static Web App (Free) | Hosts the add-in (dev environment) |

---

## Order of operations

### Step 1 — Create the App Registration (before terraform apply)

Do this first. Follow `SETUP.md` Step 1. Record:
- `tenant_id` (Directory ID)
- `client_id` (Application ID)
- `api_audience` (the App ID URI you set, e.g. `api://localhost:3000/<client-id>`)

Terraform does **not** manage the App Registration — it requires manual portal steps (authorising Office client IDs for SSO) that cannot be automated.

### Step 2 — Fill in variables

Two environments are available. They differ **only in where the add-in and the Function App run** (which sets `api_audience`, the JWT audience the function validates):

| Env | Add-in runs on | Function App runs on |
|---|---|---|
| `local` | Vite dev server (`https://localhost:3000`) | local Functions host (`func start`) |
| `dev` | Azure Static Web App | deployed on Azure |

In **both** environments the Service Bus namespace and the Storage Account are always provisioned in Azure — there are no local or emulated equivalents. In `local` mode the locally-running Functions host connects to those same cloud resources. Both envs also share one `name_prefix`, so they manage the **same** Azure resources and state file; switching env only changes the live `API_AUDIENCE` and CORS values.

Use the environment-specific file that matches what you're deploying:

```bash
cd infra

# Local (add-in on localhost:3000, functions via func start):
cp envs/local.tfvars.example envs/local.tfvars
# Edit envs/local.tfvars — fill in subscription_id, tenant_id, client_id, owner_email

# Dev (add-in on SWA, functions on Azure) — see the two-step note in the file header:
cp envs/dev.tfvars.example envs/dev.tfvars
# Edit envs/dev.tfvars — SWA hostname placeholders are filled in after first apply
```

Both files are covered by the root `.gitignore` (`*.tfvars` pattern) — never commit them.

### Step 3 — Init, plan, apply

```bash
terraform init

# Pass the env-specific var file:
terraform plan -var-file=envs/local.tfvars -out tfplan    # or dev.tfvars
terraform apply tfplan
```

### Step 4 — Read outputs

```bash
terraform output
```

Note these values — you'll use them in the next steps:
- `function_app_hostname` → `VITE_API_BASE_URL` in add-in config
- `function_app_name` → used in `func azure functionapp publish`
- `service_bus_fqdn` → `SERVICEBUS_FQDN` in `local.settings.json`
- `service_bus_namespace_id` → `--scope` for granting your user Service Bus roles

### Step 5 — Deploy function code

Terraform provisions infra only. Deploy the Python code separately:

```bash
cd ../src/functions
uv sync                                      # ensure venv is up to date
uv export --no-hashes -o requirements.txt   # Oryx installs from this on remote build
func azure functionapp publish <function-app-name-from-step-4>
```

The remote Oryx build reads `requirements.txt` and installs Python dependencies.
`pyproject.toml` is the source of truth; regenerate `requirements.txt` from it whenever you add a dependency.

### Step 6 — Wire up the add-in

1. Put `function_app_hostname` into `src/addin/.env.dev` as `VITE_API_BASE_URL`.
2. Edit `src/addin/manifest.xml`:
   - `<SourceLocation>` and `<AppDomains>` → your deployed add-in host URL.
   - `<WebApplicationInfo><Id>` → your `client_id`.
   - `<WebApplicationInfo><Resource>` → your `api_audience`.
3. Build and redeploy the add-in; re-sideload the updated manifest (see `SETUP.md` Step 6–7).

### Step 7 — Grant your user Service Bus roles (local dev)

`DefaultAzureCredential` uses your `az login` identity locally. Your account needs the data roles:

```bash
az role assignment create \
  --assignee <your-upn-or-object-id> \
  --role "Azure Service Bus Data Sender" \
  --scope <service_bus_namespace_id from terraform output>

az role assignment create \
  --assignee <your-upn-or-object-id> \
  --role "Azure Service Bus Data Receiver" \
  --scope <service_bus_namespace_id from terraform output>
```

---

## Tagging policy

Every resource and the resource group carries six `_`-prefixed tags required by Indicium's mandatory tagging policy.

- `_purpose = Testing` and `_business_criticality = Low` are **hardcoded** in `locals` — the policy pins these for a test deployment.
- `_owner_email`, `_end_date`, `_project`, `_description` come from `terraform.tfvars`.
- `terraform plan` will fail fast with a clear error if `owner_email` is not `@indicium.ai` or `end_date` is not in `ddmmyy` format or is `None`.
- The default `end_date` of `170826` is 17 Aug 2026 (the 2-month Testing cap from the 17 Jun 2026 creation date). Bring it nearer if you'll finish sooner.

---

## Tear down

```bash
cd infra
terraform destroy
```

Then delete the `outlook-addin-test` App Registration manually in Entra ID (it isn't managed by Terraform).
