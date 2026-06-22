# ── Azure subscription ─────────────────────────────────────────────────────────

variable "subscription_id" {
  description = "Azure subscription ID."
  type        = string
}

# ── Identity (from your App Registration — created manually per SETUP.md) ──────

variable "tenant_id" {
  description = "Entra ID Directory (tenant) ID."
  type        = string
}

variable "client_id" {
  description = "App Registration Application (client) ID."
  type        = string
}

variable "api_audience" {
  description = <<-EOT
    App ID URI set on the App Registration (e.g. api://localhost:3000/<client-id>
    for local dev, or api://<your-addin-host>/<client-id> for the dev deployment).
    Used as the JWT audience the enquiry-receiver validates against.
  EOT
  type        = string
}

# ── Deployment configuration ───────────────────────────────────────────────────

variable "location" {
  description = "Azure region (e.g. uksouth, eastus)."
  type        = string
}

variable "name_prefix" {
  description = <<-EOT
    Short prefix for all resource names (lowercase alphanumeric + hyphens, ≤12 chars).
    Keep it short — it is used as a prefix for the storage account name (24-char Azure limit).
    Example: oaddintest
  EOT
  type        = string

  validation {
    condition     = length(replace(var.name_prefix, "-", "")) <= 12
    error_message = "name_prefix (without hyphens) must be ≤12 characters to keep the storage account name within Azure's 24-char limit."
  }
}

# ── Mandatory Indicium tagging policy ─────────────────────────────────────────
# _purpose and _business_criticality are fixed in locals (not user-editable) because
# the policy pins them for a Testing deployment.

variable "owner_email" {
  description = "Owner email address. Must end in @mesh-ai.com (Azure subscription policy — no default; you must supply this)."
  type        = string

  validation {
    condition     = can(regex("^[^@]+@mesh-ai\\.com$", var.owner_email))
    error_message = "owner_email must end in @mesh-ai.com — required by the subscription's Required Tags policy."
  }
}

variable "end_date" {
  description = <<-EOT
    Deployment end date in ddmmyy format (e.g. 170826 = 17 Aug 2026).
    The Testing purpose requires an end date no more than 2 months out from creation.
    Must not be 'None' (forbidden for Testing purpose).
    Default is 170826 — the 2-month cap from the 17 Jun 2026 creation date.
  EOT
  type        = string
  default     = "170826"

  validation {
    condition     = var.end_date != "None" && can(regex("^[0-3][0-9][0-1][0-9][0-9]{2}$", var.end_date))
    error_message = "end_date must be in ddmmyy format (e.g. 170826) and must not be 'None' — Testing purpose requires an end date."
  }
}

variable "project" {
  description = "Project name tag (_project)."
  type        = string
  default     = "outlook-addin-test-harness"
}

variable "description" {
  description = "Optional one-line description tag (_description)."
  type        = string
  default     = "Minimal Outlook add-in test harness for event-driven enquiry pattern."
}

variable "swa_location" {
  description = <<-EOT
    Azure region for the Static Web App. SWA has a separate set of valid regions
    from general Azure resources — must be one of:
    centralus, eastus2, eastasia, westeurope, westus2.
    Does not need to match var.location.
  EOT
  type        = string
  default     = "westeurope"

  validation {
    condition     = contains(["centralus", "eastus2", "eastasia", "westeurope", "westus2"], var.swa_location)
    error_message = "swa_location must be one of: centralus, eastus2, eastasia, westeurope, westus2."
  }
}
