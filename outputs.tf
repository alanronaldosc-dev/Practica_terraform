output "application_name" {
  description = "Nombre de la aplicación"
  value       = local.application_name
}

output "unique_name" {
  description = "Nombre único generado con sufijo aleatorio"
  value       = local.unique_name
}

output "random_suffix" {
  description = "Sufijo aleatorio generado"
  value       = random_string.suffix.result
}

output "string_length" {
  description = "Longitud configurada para el string aleatorio"
  value       = var.length
}

output "environment" {
  description = "Entorno de despliegue"
  value       = var.environment
}

output "enable_morning" {
  description = "Flag de activación matutina"
  value       = var.enable_morning
}

output "regions" {
  description = "Lista de regiones configuradas"
  value       = var.regions
}

output "environment_tags" {
  description = "Mapa de etiquetas por entorno"
  value       = var.enviroment_tags
}

output "application_config" {
  description = "Configuración completa de la aplicación"
  value       = var.aplication_config
}

output "allowed_networks" {
  description = "Redes permitidas"
  value       = var.allowed_networks
}
