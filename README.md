# 🏢 SindicalizApp - Portal de Gestión Sindical

## 📋 Descripción

**SindicalizApp** es una aplicación web desarrollada para la gestión integral de un sindicato, construida con **Google Apps Script**, **Google Sheets** y **Google Web App**.

## 🎯 Características Principales

### 👥 Roles de Usuario

- **SOCIO**: Usuario base con acceso a módulos personales
- **DIRIGENTE**: Gestión de solicitudes de socios
- **ADMIN**: Acceso completo + panel administrativo

### 📦 Módulos Disponibles

#### 1. **Mis Datos** 👤

- Visualización de información personal
- Edición de región, correo y teléfono
- Estado de vinculación (ACTIVO/DESVINCULADO)

#### 2. **Préstamos** 💰

- Préstamos de Emergencia (hasta 10 cuotas)
- Préstamos de Vacaciones (hasta 8 cuotas)
- Historial con filtros y ordenamiento
- Edición/eliminación de solicitudes pendientes

#### 3. **Justificaciones** 📄

- Tipos: Turno ISS, Certificado Vacaciones, Licencia Médica, Fuerza Mayor
- Adjuntar comprobantes (JPG, PNG, PDF)
- Switch de habilitación por fecha límite (ADMIN)
- Historial con estados

#### 4. **Apelaciones** ⚖️

- Apelación de multas por mes
- Adjuntar comprobante + liquidación de sueldo (obligatorio)
- Validación de fechas (desde Marzo 2025)
- Historial completo

#### 5. **Permiso Médico** 🏥

- Tipos de permiso configurables
- Notificación automática a empresa
- Adjuntar documento posterior
- Advertencia para "Mutuo Acuerdo"

#### 6. **Registro de Asistencia** 📊

- Historial de participación en asambleas
- Vista de solo lectura

#### 7. **Sistema QR** 📱

- Verificación de vinculación QR
- Acceso para todos los usuarios

#### 8. **Gestión de Socios** (DIRIGENTE/ADMIN) 👨‍💼

- Tabs por módulo: Préstamos, Justificaciones, Apelaciones, Permisos
- Gestión de solicitudes de terceros

#### 9. **Panel Admin** (ADMIN) ⚙️

- Generación de informes
- Herramientas administrativas

## 🎨 Tecnologías Utilizadas

### Frontend

- **HTML5** + **CSS3**
- **TailwindCSS** (via CDN)
- **JavaScript** (Vanilla)
- **Google Fonts**: Inter
- **Material Icons Round**

### Backend

- **Google Apps Script**
- **Google Sheets** (Base de datos)
- **Google Drive** (Almacenamiento de archivos)

### Diseño

- **Glassmorphism** effects
- **Gradientes dinámicos**
- **Animaciones suaves**
- **Responsive design** (Mobile-first)
- **Dark mode background**

## 📂 Estructura del Proyecto

```
SINDICALIZAPP/
├── Index.html          # Frontend completo (HTML + CSS + JS)
├── Code.gs            # Backend Google Apps Script (por crear)
├── README.md          # Este archivo
└── .gitignore         # Archivos ignorados
```

## 🚀 Características Técnicas

### Validaciones

- ✅ Formato RUT chileno automático
- ✅ Validación de correos electrónicos
- ✅ Validación de teléfonos (+569)
- ✅ Límites de caracteres en campos de texto
- ✅ Validación de archivos (tamaño máx 5MB)
- ✅ Validación de fechas (rangos específicos)

### Seguridad

- 🔒 Sistema de roles y permisos
- 🔒 Validación de estado de usuario (DESVINCULADO)
- 🔒 Sesión con localStorage (opcional)
- 🔒 Recuperación de contraseña por correo

### UX/UI

- 🎨 Diseño moderno con glassmorphism
- 🎨 Animaciones suaves (fade-in, bounce, scale)
- 🎨 Loaders y spinners durante operaciones
- 🎨 Modales de confirmación
- 🎨 Mensajes de éxito/error
- 🎨 Responsive para móviles y desktop

## 📱 Vistas Principales

1. **Login** - Autenticación con RUT y contraseña
2. **Dashboard** - Menú principal con tarjetas de módulos
3. **Perfil** - Datos personales editables
4. **Préstamos** - Menú → Formulario → Historial
5. **Justificaciones** - Tabs (Nueva/Historial)
6. **Apelaciones** - Tabs (Nueva/Historial)
7. **Permisos Médicos** - Tabs (Nueva/Historial)
8. **Asistencia** - Solo historial
9. **Gestión** - Tabs por módulo (DIRIGENTE)
10. **Admin** - Herramientas administrativas

## 🔧 Funciones JavaScript Principales

### Core

- `handleLogin()` - Autenticación
- `handleLogout()` - Cierre de sesión
- `switchView()` - Cambio de vistas
- `navAction()` - Navegación por módulos
- `applyRolePermissions()` - Aplicar permisos según rol

### Helpers

- `formatRutDisplay()` - Formatear RUT
- `formatRutInput()` - Formatear input RUT en tiempo real
- `parsePhoneToSuffix()` - Parsear teléfono
- `showWarning()` - Mostrar advertencia
- `showSuccess()` - Mostrar éxito
- `showConstruction()` - Módulo en construcción
- `parseCustomDate()` - Parsear fechas personalizadas

### Módulos

- `submitLoanRequest()` - Enviar préstamo
- `submitJustification()` - Enviar justificación
- `submitAppeal()` - Enviar apelación
- `submitPermission()` - Enviar permiso médico
- `loadAttendanceRecords()` - Cargar asistencia
- `loadManagementView()` - Cargar gestión

## 🎯 Próximas Funcionalidades

- [ ] Informe de Apelaciones
- [ ] Informe de Asistencia
- [ ] Cambio de Rol
- [ ] Notificaciones push
- [ ] Exportación de datos
- [ ] Gráficos estadísticos

## 👨‍💻 Autor

Desarrollado para **Sindicato SLIM n°3**

## 📄 Licencia

Proyecto privado - Todos los derechos reservados

---

**Versión**: 1.0.0  
**Última actualización**: Diciembre 2024
