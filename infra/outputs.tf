output "function_app_hostname" {
  description = "Base URL of the Function App — use as VITE_API_BASE_URL in src/addin/.env.* and as the API endpoint in the add-in manifest."
  value       = "https://${azurerm_linux_function_app.main.default_hostname}"
}

output "function_app_name" {
  description = "Function App resource name — used in: func azure functionapp publish <name>"
  value       = azurerm_linux_function_app.main.name
}

output "service_bus_fqdn" {
  description = "Service Bus namespace FQDN — use as SERVICEBUS_FQDN and ServiceBusConnection__fullyQualifiedNamespace in local.settings.json."
  value       = "${azurerm_servicebus_namespace.main.name}.servicebus.windows.net"
}

output "service_bus_namespace_id" {
  description = "Service Bus namespace resource ID — pass as --scope when granting your user the Data Sender/Receiver roles for local dev (SETUP.md Step 5)."
  value       = azurerm_servicebus_namespace.main.id
}

output "application_insights_name" {
  description = "Application Insights resource name — search here for correlationId round-trip proof after the end-to-end test."
  value       = azurerm_application_insights.main.name
}

output "resource_group_name" {
  description = "Resource group name."
  value       = azurerm_resource_group.main.name
}

output "addin_url" {
  description = "Add-in hosting URL (SWA) — use as the base for SourceLocation in manifest.xml and as the host in the App ID URI (api://<host>/<client-id>)."
  value       = "https://${azurerm_static_web_app.addin.default_host_name}"
}

output "addin_deploy_token" {
  description = "SWA deployment token — pass to `swa deploy --deployment-token`. Treat as a secret."
  value       = azurerm_static_web_app.addin.api_key
  sensitive   = true
}

output "storage_account_key" {
  description = "Storage account primary access key — populate STORAGE_ACCOUNT_KEY and AzureWebJobsStorage in dev.settings.json."
  value       = azurerm_storage_account.main.primary_access_key
  sensitive   = true
}

output "service_bus_connection_string" {
  description = "Service Bus func-rule connection string — populate SERVICEBUS_CONNECTION_STRING and ServiceBusConnection in dev.settings.json."
  value       = azurerm_servicebus_namespace_authorization_rule.func.primary_connection_string
  sensitive   = true
}
