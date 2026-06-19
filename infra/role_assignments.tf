# ── Service Bus RBAC (Managed Identity, no connection strings) ─────────────────
#
# REQUIRES Owner or User Access Administrator on the subscription/resource group.
# If your account only has Contributor, apply this file separately once someone
# with the right permissions is available:
#
#   terraform apply -target=azurerm_role_assignment.func_sb_sender \
#                   -target=azurerm_role_assignment.func_sb_receiver
#
# Or have an Owner run the equivalent az CLI commands (see infra/README.md).
#
# Without these assignments the Function App's Managed Identity cannot send to
# or receive from Service Bus, and both functions will get 403 at runtime.

# submission-receiver writes messages to the queue
resource "azurerm_role_assignment" "func_sb_sender" {
  scope                = azurerm_servicebus_namespace.main.id
  role_definition_name = "Azure Service Bus Data Sender"
  principal_id         = azurerm_linux_function_app.main.identity[0].principal_id
}

# submission-processor reads messages from the queue (currently disabled)
resource "azurerm_role_assignment" "func_sb_receiver" {
  scope                = azurerm_servicebus_namespace.main.id
  role_definition_name = "Azure Service Bus Data Receiver"
  principal_id         = azurerm_linux_function_app.main.identity[0].principal_id
}

# ── Blob storage RBAC (submission_receiver + download_generator) ───────────────
#
# Storage Blob Delegator: required to call get_user_delegation_key, which is
# needed to sign User Delegation SAS tokens (avoids storing account keys).
# Must be scoped to the storage account (cannot be scoped to a container).
#
# Storage Blob Data Contributor: allows the function to write blobs into the
# blobs container. Scoped to the storage account for simplicity; tighten
# to the container scope if the policy requires least-privilege per resource.

resource "azurerm_role_assignment" "func_storage_blob_delegator" {
  scope                = azurerm_storage_account.main.id
  role_definition_name = "Storage Blob Delegator"
  principal_id         = azurerm_linux_function_app.main.identity[0].principal_id
}

resource "azurerm_role_assignment" "func_storage_blob_contributor" {
  scope                = azurerm_storage_account.main.id
  role_definition_name = "Storage Blob Data Contributor"
  principal_id         = azurerm_linux_function_app.main.identity[0].principal_id
}
