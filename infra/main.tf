# ── Mandatory tags (applied to every resource via tags = local.common_tags) ────
# _purpose and _business_criticality are hardcoded — policy pins them for Testing.
locals {
  common_tags = {
    _purpose              = "Testing"
    _business_criticality = "Low"
    _end_date             = var.end_date
    _owner_email          = var.owner_email
    _project              = var.project
    _description          = var.description
  }
}

# ── Resource Group ─────────────────────────────────────────────────────────────

resource "azurerm_resource_group" "main" {
  name     = "${var.name_prefix}-rg"
  location = var.location
  tags     = local.common_tags
}

# ── Service Bus ────────────────────────────────────────────────────────────────
# Standard SKU required for duplicate detection.

resource "azurerm_servicebus_namespace" "main" {
  name                = "${var.name_prefix}-sbns"
  location            = azurerm_resource_group.main.location
  resource_group_name = azurerm_resource_group.main.name
  sku                 = "Standard"
  tags                = local.common_tags
}

resource "azurerm_servicebus_queue" "submission" {
  name         = "submission-queue"
  namespace_id = azurerm_servicebus_namespace.main.id

  # Duplicate detection: the submission-receiver sets the Service Bus messageId to the
  # email itemId so the platform deduplicates re-sends of the same email.
  requires_duplicate_detection          = true
  duplicate_detection_history_time_window = "PT10S"

  lock_duration              = "PT1M"   # how long a consumer has before message re-appears
  max_delivery_count         = 5        # after 5 failures the message goes to dead-letter
  dead_lettering_on_message_expiration = true
  default_message_ttl        = "P14D"   # 14 days
}

# Authorization rule for the Function App — send+listen, no manage.
# Avoids requiring Managed Identity RBAC (which needs Owner on the subscription).
resource "azurerm_servicebus_namespace_authorization_rule" "func" {
  name         = "func-rule"
  namespace_id = azurerm_servicebus_namespace.main.id
  listen       = true
  send         = true
  manage       = false
}

# ── Storage Account (required by the Functions host) ──────────────────────────
# Name rules: 3–24 chars, lowercase alphanumeric only, globally unique.
# Hyphens are stripped; keep name_prefix short (≤12 chars after stripping hyphens).

resource "azurerm_storage_account" "main" {
  name                     = lower(replace("${var.name_prefix}sa", "-", ""))
  resource_group_name      = azurerm_resource_group.main.name
  location                 = azurerm_resource_group.main.location
  account_tier             = "Standard"
  account_replication_type = "LRS"
  tags                     = local.common_tags

  blob_properties {
    cors_rule {
      allowed_headers    = ["*"]
      allowed_methods    = ["DELETE", "GET", "HEAD", "MERGE", "POST", "OPTIONS", "PUT", "PATCH"]
      allowed_origins    = [
        "https://localhost:3000",
        "https://${azurerm_static_web_app.addin.default_host_name}",
      ]
      exposed_headers    = ["*"]
      max_age_in_seconds = 3600
    }
  }
}

# ── Blob container (submissions + downloads) ──────────────────────────────────
# Single container for all blobs: submission ZIPs (stored permanently until
# lifecycle policy expires them) and download ZIPs (accessed via SAS URL).
# Blobs are private; access is via User Delegation SAS URLs only.
#
# Flat layout:
#   submissions/SUB-XXXXXXXX/<filename>  ← submission_receiver (permanent until 30d policy)
#   submissions/<uuid>/<filename>        ← download_generator   (1-hour SAS, 30d max)

resource "azurerm_storage_container" "submissions" {
  name                  = "submissions"
  storage_account_id    = azurerm_storage_account.main.id
  container_access_type = "private"
}

resource "azurerm_storage_management_policy" "submissions_lifecycle" {
  storage_account_id = azurerm_storage_account.main.id

  rule {
    name    = "delete-submissions-after-30-days"
    enabled = true
    filters {
      blob_types   = ["blockBlob"]
      prefix_match = ["submissions/submission/", "submissions/attachments/"]
    }
    actions {
      base_blob {
        delete_after_days_since_modification_greater_than = 30
      }
    }
  }

  rule {
    name    = "delete-downloads-after-7-days"
    enabled = true
    filters {
      blob_types   = ["blockBlob"]
      prefix_match = ["submissions/downloads/"]
    }
    actions {
      base_blob {
        delete_after_days_since_modification_greater_than = 7
      }
    }
  }
}

# ── Observability ──────────────────────────────────────────────────────────────
# Log Analytics workspace + workspace-based Application Insights.
# Function logs and the correlationId round-trip are queryable here.

resource "azurerm_log_analytics_workspace" "main" {
  name                = "${var.name_prefix}-law"
  location            = azurerm_resource_group.main.location
  resource_group_name = azurerm_resource_group.main.name
  sku                 = "PerGB2018"
  retention_in_days   = 30
  tags                = local.common_tags
}

resource "azurerm_application_insights" "main" {
  name                = "${var.name_prefix}-ai"
  location            = azurerm_resource_group.main.location
  resource_group_name = azurerm_resource_group.main.name
  workspace_id        = azurerm_log_analytics_workspace.main.id
  application_type    = "web"
  tags                = local.common_tags
}

# ── Function App (Consumption Y1 — change sku_name to EP1 for Premium) ────────

resource "azurerm_service_plan" "main" {
  name                = "${var.name_prefix}-asp"
  resource_group_name = azurerm_resource_group.main.name
  location            = azurerm_resource_group.main.location
  os_type             = "Linux"
  sku_name            = "Y1" # Consumption plan; ask before changing to Premium
  tags                = local.common_tags
}

resource "azurerm_linux_function_app" "main" {
  name                       = "${var.name_prefix}-func"
  resource_group_name        = azurerm_resource_group.main.name
  location                   = azurerm_resource_group.main.location
  storage_account_name       = azurerm_storage_account.main.name
  storage_account_access_key = azurerm_storage_account.main.primary_access_key
  service_plan_id            = azurerm_service_plan.main.id
  functions_extension_version = "~4"

  # System-assigned Managed Identity — grants Service Bus access without connection strings.
  identity {
    type = "SystemAssigned"
  }

  site_config {
    application_stack {
      python_version = "3.13"
    }

    # CORS: allow the add-in origin. Code-level CORS in function_app.py handles
    # local dev; this handles Azure-deployed traffic.
    cors {
      allowed_origins = [
        "https://${azurerm_static_web_app.addin.default_host_name}",
        "https://localhost:3000",
      ]
      support_credentials = false
    }
  }

  app_settings = {
    "FUNCTIONS_WORKER_RUNTIME" = "python"

    # Application Insights
    "APPLICATIONINSIGHTS_CONNECTION_STRING" = azurerm_application_insights.main.connection_string

    # JWT validation
    "TENANT_ID"    = var.tenant_id
    "API_AUDIENCE" = var.api_audience

    # CORS (used by function code for local-dev parity)
    "ALLOWED_CORS_ORIGIN" = "https://${azurerm_static_web_app.addin.default_host_name}"

    # Blob storage — used by submission_receiver and download_generator
    "STORAGE_ACCOUNT_NAME"  = azurerm_storage_account.main.name
    "STORAGE_ACCOUNT_KEY"   = azurerm_storage_account.main.primary_access_key
    "BLOB_CONTAINER_NAME"   = azurerm_storage_container.submissions.name

    # Service Bus — connection string auth (no Managed Identity RBAC needed)
    "SERVICEBUS_FQDN"            = "${azurerm_servicebus_namespace.main.name}.servicebus.windows.net"
    "SERVICEBUS_QUEUE_NAME"      = azurerm_servicebus_queue.submission.name
    "SERVICEBUS_CONNECTION_STRING" = azurerm_servicebus_namespace_authorization_rule.func.primary_connection_string

    # Connection string for the Service Bus queue trigger binding.
    # The binding's `connection` parameter is "ServiceBusConnection"; setting it
    # as a plain connection string bypasses Managed Identity RBAC requirements.
    "ServiceBusConnection" = azurerm_servicebus_namespace_authorization_rule.func.primary_connection_string

    # Tell Oryx (Azure remote build) to install Python deps from requirements.txt on publish.
    "SCM_DO_BUILD_DURING_DEPLOYMENT" = "true"
  }

  tags = local.common_tags
}

# ── Static Web App (add-in hosting) ───────────────────────────────────────────
# Free tier: static HTML/JS, global CDN, valid HTTPS cert included.
# The deployment token (api_key) is output as addin_deploy_token.
#
# SWA regions differ from general Azure regions — valid values include:
#   centralus, eastus2, eastasia, westeurope, westus2
# Set swa_location in terraform.tfvars (defaults to westeurope).
#
# The SWA hostname has a random slug assigned by Azure (e.g. proud-pond-0123456789).
# It is not predictable before apply — use `terraform output addin_url` after apply.

resource "azurerm_static_web_app" "addin" {
  name                = "${var.name_prefix}-swa"
  resource_group_name = azurerm_resource_group.main.name
  location            = var.swa_location
  sku_tier            = "Free"
  sku_size            = "Free"
  tags                = local.common_tags
}

