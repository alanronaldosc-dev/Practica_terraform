output "resource_group_name" {
  description = "Nombre del Resource Group creado"
  value       = azurerm_resource_group.main.name
}

output "resource_group_location" {
  description = "Ubicación del Resource Group"
  value       = azurerm_resource_group.main.location
}

output "virtual_network_name" {
  description = "Nombre de la Virtual Network creada"
  value       = azurerm_virtual_network.main.name
}

output "virtual_network_address_space" {
  description = "Espacio de direcciones de la Virtual Network"
  value       = azurerm_virtual_network.main.address_space
}

output "common_tags" {
  description = "Etiquetas comunes aplicadas a los recursos"
  value       = local.common_tags
}
