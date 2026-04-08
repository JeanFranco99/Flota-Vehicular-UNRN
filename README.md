# Flota Vehicular UNRN

Sistema de gestión de flota vehicular para la Universidad Nacional de Río Negro (UNRN), desarrollado con Google Apps Script.

## 📋 Descripción

Aplicación web para administrar el uso de vehículos institucionales, permitiendo:
- Solicitar vehículos para viajes oficiales
- Asignar y gestionar vehículos disponibles
- Registrar verificaciones de estado (entrada/salida)
- Controlar mantenimientos
- Gestionar usuarios con roles diferenciados

## 🚀 Características

- **Autenticación**: Integración con Google Workspace
- **Roles de usuario**:
  - **Encargado**: Panel de gestión completo (aprobación, asignación, verificaciones, mantenimientos)
  - **Solicitante**: Solicitud de vehículos y seguimiento de solicitudes
- **Gestión de solicitudes**: Creación, aprobación, rechazo y asignación de vehículos
- **Verificaciones**: Checklist completo de verificación de estado del vehículo (entrada/salida)
- **Mantenimientos**: Registro y seguimiento de mantenimientos por vehículo
- **Configuración flexible**: Hoja de configuración para gestionar usuarios y roles
- **Comprobantes**: Generación de comprobantes imprimibles

## 📁 Estructura del Proyecto

```
Flota_Vehicular/
├── Código.js           # Backend principal (Google Apps Script)
├── Encargado.html      # Interfaz para gestores
├── Solicitante.html    # Interfaz para solicitantes
├── comprobante.html    # Template para comprobantes
├── appsscript.json     # Configuración del proyecto
├── .clasp.json         # Configuración CLI clasp
└── README.md           # Este archivo
```

## 🔧 Requisitos Previos

- Cuenta Google Workspace (@gmail.com)
- Acceso a Google Sheets
- Google Apps Script (incluido en Google Drive)
- [clasp](https://github.com/google/clasp) (opcional, para desarrollo local)

## 📦 Instalación y Despliegue

### Opción 1: Importar proyecto existente

1. Ir a [Google Apps Script](https://script.google.com/)
2. Crear nuevo proyecto en blanco
3. Copiar el contenido de los archivos `.js` y `.html` al editor

### Opción 2: Usando clasp (desarrollo local)

```bash
# 1. Instalar clasp globalmente
npm install -g @google/clasp

# 2. Autenticarse
clasp login

# 3. Clonar el proyecto (si ya está desplegado)
clasp clone <SCRIPT_ID>

# 4. Push de archivos locales
clasp push

# 5. Abrir en editor
clasp open
```

## ⚙️ Configuración

### Spreadsheet de Base de Datos

El sistema usa una planilla de cálculo como base de datos. ID actual:

### Hojas requeridas

| Hoja | Descripción |
|------|-------------|
| **Solicitudes** | Registro de solicitudes de vehículos |
| **Vehículos** | Catálogo de vehículos disponibles |
| **Mantenimientos** | Registro de mantenimientos |
| **Configuración** | Usuarios y roles (Encargado/Solicitante) |
| **Verificaciones** | Checklist de verificación entrada/salida |


### Agregar usuarios

1. Abrir la hoja de **Configuración**
2. Agregar email del usuario en la columna correspondiente:
   - Encargados: rol "Encargado"
   - Otros usuarios automáticamente son "Solicitante"

## 📖 Uso

### Flujo Solicitante

1. Acceder a la aplicación web
2. Completar formulario de solicitud:
   - Fechas de salida y regreso
   - Itinerario/trajectory
   - Tareas a realizar
   - Participantes
   - Conductores
3. Enviar solicitud
4. Recibir email de confirmación
5. Consultar estado de la solicitud

### Flujo Encargado

1. Ver lista de solicitudes pendientes
2. Aprobar o rechazar solicitudes
3. Asignar vehículo disponible
4. Registrar verificaciones (km, estado)
5. Gestionar mantenimientos
6. Generar comprobantes

## 🔍 Validaciones

- **Patentes**: Formato argentino (`AAA 123` o `AA 123 AA`)
- **Fechas**: Formato `dd/MM/yyyy`
- **Dominio email**: Restringido a `@unrn.edu.ar`

## 📊 Tecnologías

- **Runtime**: Google Apps Script (V8)
- **Frontend**: HTML Service + CSS
- **Backend**: Google Apps Script (Server-side)
- **Database**: Google Sheets
- **Auth**: Google Workspace (Session.getActiveUser())

## 📝 Notas de Desarrollo

- Todas las operaciones de escritura usan `LockService` para evitar conflictos de concurrencia
- Validación de formato de patentes con regex: `^(?:[A-Z]{3}\s\d{3}|[A-Z]{2}\s\d{3}\s[A-Z]{2})$`

## 📄 Licencia

Uso interno - Universidad Nacional de Río Negro

---
