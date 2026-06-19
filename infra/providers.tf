terraform {
  required_version = ">= 1.6"
  required_providers {
    azurerm = {
      source  = "hashicorp/azurerm"
      version = "~> 4.77"
    }
  }
}

provider "azurerm" {
  subscription_id = var.subscription_id

  features {}
}
