variable "project_name" {
  description = "Nombre corto del proyecto o aplicación."
  type        = string

  validation {
    condition     = length(var.project_name) >= 3 && length(var.project_name) <= 20
    error_message = "project_name debe contener entre 3 y 20 caracteres."
  }
}

variable "environment" {
  description = "Ambiente donde se desplegarán los recursos."
  type        = string

  validation {
    condition = contains(
      ["dev", "test", "qa", "stg", "prod"],
      var.environment
    )
    error_message = "environment debe ser dev, test, qa, stg o prod."
  }
}

variable "location" {
  description = "Región de Azure donde se desplegarán los recursos."
  type        = string
  default     = "mexicocentral"
}

variable "vnet_address_space" {
  description = "Espacio de direcciones CIDR asignado a la Virtual Network."
  type        = list(string)
  default     = ["10.0.0.0/16"]
}

variable "tags" {
  description = "Etiquetas comunes para los recursos de Azure."
  type        = map(string)
  default = {
    managed_by = "terraform"
  }
}

variable "subscription_id" {
  description = "ID de la suscripción de Azure donde se desplegarán los recursos."
  type        = string
  default     = "947b767e-d2d0-4136-938f-918c6fbc48c0"
  sensitive   = true
}
