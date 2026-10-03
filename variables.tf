variable "length" {
  description = "Length of the random string"
  type        = number
  default     = 20
}

variable "application_name" {
  description = "Name of the application"
  type        = string
  default     = "myapp"
}

variable "environment" {
  description = "Deployment environment (e.g. dev, staging, prod)"
  type        = string
  default     = "dev"
}

variable "enable_morning" {
    description = "value"
    type = bool
    default = true
}

variable "regions" {
    description = "lista de regiones"
    type = list(string)
    default = [ "us-east-1", "us-west-2" ]
}

variable "enviroment_tags" {
    description = "value"
    type = map(string)
    default = {
      dev = "desarrollo"
      prod = "produccion"
    }
}
variable "aplication_config" {
  description = "Configuracion de sistema"
  type = object({
    version = string
    maintainer = string
    dependences = list(string)
  })
  default = {
    version = "1.0.0"
    maintainer = "Jhon"
    dependences = ["dependecy1", "dependecy2"]
  }
}

variable "allowed_networks" {
  description = "Lista de redes permitidas"
  type = set(string)
  default = ["10.0.0.0/16","10.1.0.0/16"]
}
