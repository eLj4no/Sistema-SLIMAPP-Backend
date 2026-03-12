# 🏢 SLIMAPP — Sistema de Gestión Sindical
### Sindicato SLIM N°3

---

## 📋 Descripción

**SLIMAPP** es una aplicación web desarrollada para la gestión integral del **Sindicato SLIM N°3**, construida sobre **Google Apps Script**, **Google Sheets** como base de datos y **Google Drive** para almacenamiento de archivos. Se sirve como **Google Web App** y es accesible desde cualquier dispositivo con navegador.

---

## 📁 Estructura del Proyecto

```
SLIMAPP/
├── Code.gs          # Backend completo (Google Apps Script)
├── Index.html       # Frontend principal (HTML + CSS + JS en un solo archivo)
├── QR_Access.html   # Página independiente para el sistema de control QR
└── README.md        # Este archivo
```

---

## 🗄️ Arquitectura de Base de Datos

El sistema utiliza **6 Google Spreadsheets independientes**, cada uno dedicado a un módulo.

| Clave CONFIG | Nombre hoja principal | Descripción |
|---|---|---|
| `USUARIOS` | `BD_SLIMAPP` | Registro maestro de socios |
| `JUSTIFICACIONES` | `BD_JUSTIFICACIONES` | Justificaciones de inasistencia |
| `APELACIONES` | `BD_APELACIONES` | Apelaciones de multas |
| `PRESTAMOS` | `BD_PRESTAMOS` | Solicitudes de préstamos |
| `PERMISOS_MEDICOS` | `BD_Permisos medicos` | Permisos médicos |
| `CREDENCIALES` | `IMPRESION` | Estado de credenciales sindicales |

### Hojas auxiliares
- `CONFIG_JUSTIFICACIONES` — Configuración del switch de habilitación y fecha de evento
- `Validación-Prestamos` — Hoja de validación manual de préstamos por directiva
- `Registros-eliminados` — Respaldo de préstamos eliminados
- `HISTORIAL_CREDENCIALES` — Log de cambios de estado de credenciales

### Carpetas Google Drive

| Módulo | Carpeta |
|---|---|
| Justificaciones | `1UD9hQz1FuacSb3QYrahRl7IfvlpKn8v6` |
| Apelaciones — Comprobantes | `15BmK5pf5Txrxdzdrny23S5q35NDxLy4P` |
| Apelaciones — Liquidaciones | `1dR7fM6TW99tunNaMZliyvXc-L23nHKVY` |
| Apelaciones — Devoluciones | `1LGLKA3fiCJXf2ouIqlxq3jk_ZSxI3IyM` |
| Permisos Médicos | `1nCYxD5sJLszBBA6s2DquGW8vlKGZp4ty` |

---

## 👥 Roles de Usuario

| Rol | Acceso |
|---|---|
| `SOCIO` | Acceso a sus propios módulos personales |
| `DIRIGENTE` | Puede gestionar trámites a nombre de socios + módulo Gestión |
| `ADMIN` | Acceso completo + Panel Admin + configuración de switches |

---

## 📦 Módulos del Sistema

### 1. 👤 Mis Datos
- Visualización de datos personales del socio
- Edición de: región, correo, teléfono, banco, tipo de cuenta, número de cuenta
- Visualización del estado de vinculación (`ACTIVO` / `DESVINCULADO`)
- Visualización del estado de credencial sindical
- Visualización del estado en negociación colectiva

### 2. 💰 Préstamos
- **Tipos disponibles:** Préstamo de Emergencia (hasta 10 cuotas, $200.000) y Préstamo de Vacaciones (hasta 8 cuotas, $150.000)
- Validación de préstamo activo por tipo (no se puede tener dos del mismo tipo simultáneamente)
- Cálculo automático de fecha de término (lógica contable: si se solicita después del día 24, el descuento empieza el mes siguiente)
- Medios de pago configurables
- **Estados:** `Solicitado` → `Enviado` → `Vigente` → `Finalizado`
- Edición y eliminación de solicitudes en estado `Solicitado`
- Historial completo con filtros
- **Trigger diario** (`procesarValidacionPrestamos`): sincroniza hoja de validación con BD principal y envía notificaciones al cambiar estado
- Generación de informe Excel de préstamos solicitados (envío por correo al ADMIN)

### 3. 📄 Justificaciones
- **Tipos de motivo:** Turno ISS, Certificado Vacaciones, Licencia Médica, Fuerza Mayor
- Adjuntar comprobante de respaldo (JPG, PNG, PDF — máximo 15MB)
- **Estados:** `Enviado` → `Aceptado` / `Aceptado/Obs` / `Rechazado`
- **Switch de habilitación** controlado por ADMIN (columna HABILITADO en CONFIG_JUSTIFICACIONES)
- **Modo A — Por evento:** Si el ADMIN configura una `Fecha_Evento`, el sistema genera automáticamente un código de asamblea en formato `YYYY_MM_DD` y valida que el socio no haya justificado ya ese evento específico
- **Modo B — Por mes:** Si no hay evento activo, valida que el socio no tenga justificación enviada o aceptada en el mes calendario actual
- **Trigger cada 8 horas** (`verificarCambiosJustificaciones`): detecta cambios de estado y notifica al socio por correo
- **Notificaciones por correo:**
  - Al socio cuando gestiona por sí mismo: "Justificación Ingresada" (naranja `#ea580c`)
  - Al dirigente cuando gestiona a nombre de un socio: "Respaldo Gestión" (gris `#475569`)
  - Copia al socio cuando el dirigente gestiona: mensaje personalizado indicando que fue ingresada por un dirigente

### 4. ⚖️ Apelaciones de Multas
- Apelación de descuentos por inasistencia, por mes calendario
- **Liquidación de sueldo obligatoria** + comprobante opcional
- Validación de mes (solo desde Marzo 2025, sin meses futuros)
- **Estados:** `Enviado` → `Aceptado` / `Aceptado-Obs` / `Rechazado`
- Comprobante de devolución adjuntable cuando la apelación es aceptada
- **Trigger cada 8 horas** (`verificarCambiosApelaciones`): notifica al socio cambios de estado
- **Trigger cada 1 hora** (`procesarPermisosComprobantesDevolucion`): otorga permisos de acceso a comprobantes de devolución cuando se suben
- **Notificaciones por correo:**
  - Al socio cuando gestiona solo: "Comprobante de Apelación" (azul `#1d4ed8`)
  - Al dirigente cuando gestiona a nombre de un socio: "Gestión Realizada" (gris `#475569`)
  - Copia al socio cuando el dirigente gestiona: mensaje con detalle completo

### 5. 🏥 Permisos Médicos
- Tipos de permiso configurables
- Rango de fecha de inicio: máximo 7 días antes o después de la fecha actual
- Adjuntar documento de respaldo **posterior** a la atención médica (desde historial)
- Notificación automática al representante legal de la empresa al solicitar y al adjuntar documento
- Doble validación anti-race-condition para prevenir duplicados por concurrencia
- **Advertencia especial** para permisos de tipo "Mutuo Acuerdo"
- **Estados:** `Solicitado` → `Documento Adjuntado` → `Anulado`
- Anulación solo en estado `Solicitado`, con notificación al representante legal
- **Notificaciones por correo:**
  - Al socio cuando gestiona solo: aviso de solicitud registrada con recordatorio de adjuntar documento (verde `#10b981`)
  - Al dirigente cuando gestiona a nombre de un socio: "Permiso Médico Ingresado" (gris `#475569`)
  - Copia al socio cuando el dirigente gestiona: mensaje con aviso y recordatorio

### 6. 📊 Registro de Asistencia
- Historial de participación en asambleas del socio
- Vista de solo lectura
- Datos registrados mediante el Sistema QR

### 7. 🎫 Estado de Credencial
- Visualización del estado actual de la credencial sindical del socio
- Estados posibles: `SOLICITADO`, `DISPONIBLE`, `ENTREGADO`, `NO VIGENTE`, `DATOS INCORRECTOS`, `REIMPRIMIR`
- Colores e iconos diferenciados por estado
- **Trigger semanal** (`verificarCambiosCredenciales`, domingos 8 AM): detecta cambios de estado en la hoja IMPRESION y notifica al socio por correo
- Primera ejecución inicializa el estado sin enviar notificaciones

### 8. 👨‍💼 Gestión de Socios (DIRIGENTE / ADMIN)
- Panel para gestionar trámites a nombre de socios
- **Tabs disponibles:** Préstamos, Justificaciones, Apelaciones, Permisos Médicos
- Búsqueda de socios por RUT con validación de rol (DIRIGENTE solo puede gestionar SOCIOs)
- **ConsultaID:** consulta del ID de credencial y código QR de un socio
  - Muestra: RUT, nombre, cargo, estado, rol, ID credencial, imagen QR
  - Restricción: DIRIGENTE no puede consultar datos de otros DIRIGENTES ni ADMINs

### 9. ⚙️ Panel Admin (ADMIN)
- Switch de habilitación del módulo de justificaciones
- Configuración de `Fecha_Evento` para asambleas (genera código ASAMBLEA automático en formato `YYYY_MM_DD`)
- Generación de informe de préstamos solicitados (Excel adjunto por correo)
- Configuración de triggers (`configurarTriggers`)
- Herramientas de mantenimiento

### 10. 📱 Sistema QR (`QR_Access.html`)
- Página independiente servida desde la misma Web App
- Validación de vinculación del socio por RUT
- Registro de asistencia a asambleas con control de duplicados
- Parámetros GET: `action`, `rut`, `asamblea`

---

## 🔐 Sistema de Permisos de Archivos (Google Drive)

Toda subida de archivos pasa por la función centralizada `subirArchivoConPermisos`, que:

1. Valida tamaño (máximo 15MB)
2. Crea el archivo en la carpeta correspondiente con nombre estandarizado (`PREFIJO-UUID-RUT`)
3. Configura el archivo como privado (`PRIVATE`)
4. Otorga acceso de lectura **silencioso** (sin email de notificación de Drive) a los correos involucrados mediante **3 intentos con fallback**:
   - Intento 1: Drive API Avanzada (`Drive.Permissions.insert` con `sendNotificationEmails: false`)
   - Intento 2: `file.addViewer()` con `Utilities.sleep(1000)`
   - Intento 3: Reintento final con `Utilities.sleep(2000)`
5. Retorna URL de acceso y reporte de permisos otorgados/fallidos

La función `validarCorreosParaPermisos` prepara la lista de correos a los que se otorga acceso y genera alertas si algún correo no es válido.

---

## 📬 Sistema de Notificaciones por Correo

Todos los correos se envían con la función `enviarCorreoEstilizado`, que genera HTML responsivo con el diseño institucional del sindicato.

### Colores por módulo
| Módulo | Color | Hex |
|---|---|---|
| Justificaciones | Naranja | `#ea580c` |
| Apelaciones (socio) | Azul institucional | `#1d4ed8` |
| Permisos Médicos | Verde | `#10b981` |
| Gestión de Dirigente | Gris/Azul | `#475569` |
| Préstamos | Azul | `#2563eb` |
| Credenciales | Variable por estado | — |

### Lógica de envío en gestión de dirigente

Cuando un **DIRIGENTE o ADMIN** realiza un trámite a nombre de un socio, el sistema envía:

1. **Correo al dirigente** — asunto "Respaldo Gestión [Módulo]", con todos los detalles del trámite y enlace al archivo
2. **Copia al socio** (si tiene correo válido registrado) — mismo contenido pero con mensaje personalizado indicando que fue ingresado por un dirigente

Si el socio **no tiene correo registrado**, solo se envía al dirigente y el sistema continúa sin error.

---

## ⚡ Triggers Automáticos

Configurados mediante `configurarTriggers()` (ejecutar manualmente una sola vez). **Advertencia:** esta función elimina todos los triggers existentes antes de recrearlos.

| Función | Frecuencia | Descripción |
|---|---|---|
| `verificarCambiosJustificaciones` | Cada 8 horas | Detecta cambios de estado y notifica al socio |
| `verificarCambiosApelaciones` | Cada 8 horas | Detecta cambios de estado y notifica al socio |
| `procesarValidacionPrestamos` | Diario 8 AM | Sincroniza validaciones y notifica cambios de estado |
| `verificarCambiosPrestamos` | Diario 8 AM | Verificación adicional de cambios en préstamos |
| `procesarPermisosComprobantesDevolucion` | Cada 1 hora | Otorga permisos a comprobantes de devolución subidos |
| `verificarCambiosCredenciales` | Domingos 8 AM | Detecta cambios de estado en credenciales y notifica |

---

## 🛠️ Funciones Principales del Backend (`Code.gs`)

### Autenticación y Usuarios
| Función | Descripción |
|---|---|
| `validarUsuario(rut, password)` | Login con RUT e ID credencial |
| `obtenerDatosUsuario(rut)` | Datos completos del perfil |
| `actualizarDatoUsuario(rut, campo, valor)` | Edición de campos editables |
| `obtenerUsuarioPorRut(rut)` | Búsqueda con caché de 10 minutos |
| `recuperarContrasena(rut)` | Consulta de correo registrado |
| `enviarContrasenaCorreo(rut)` | Envío de contraseña por correo |
| `verificarRolUsuario(rut, roles)` | Validación de permisos por rol |

### Justificaciones
| Función | Descripción |
|---|---|
| `enviarJustificacion(...)` | Crea registro, sube archivo, envía correos |
| `validarJustificacionMesActual(rut)` | Valida restricción por evento o mes |
| `obtenerHistorialJustificaciones(rut)` | Historial del socio |
| `obtenerEstadoSwitchJustificaciones()` | Lee CONFIG: habilitado + fecha evento |
| `toggleSwitchJustificaciones(estado, fecha)` | Actualiza switch (ADMIN) |
| `verificarDisponibilidadJustificaciones()` | Verifica si el módulo está activo |

### Apelaciones
| Función | Descripción |
|---|---|
| `enviarApelacion(...)` | Crea registro, sube archivos, envía correos |
| `verificarDisponibilidadApelaciones(mes)` | Valida mes (no futuro, mínimo Marzo 2025) |
| `obtenerHistorialApelaciones(rut)` | Historial del socio |
| `subirComprobanteDevolucion(id, archivo)` | Adjunta comprobante de devolución |

### Préstamos
| Función | Descripción |
|---|---|
| `crearSolicitudPrestamo(...)` | Crea solicitud con fecha de término calculada |
| `obtenerHistorialPrestamos(rut)` | Historial del socio |
| `modificarSolicitud(id, cuotas, medio)` | Editar solicitud en estado Solicitado |
| `eliminarSolicitud(id)` | Eliminar con respaldo en hoja Registros-eliminados |
| `procesarValidacionPrestamos()` | Trigger: sincroniza validaciones |

### Permisos Médicos
| Función | Descripción |
|---|---|
| `solicitarPermisoMedico(...)` | Crea solicitud con doble validación anti-race |
| `adjuntarDocumentoPermiso(id, archivo)` | Adjunta documento posterior |
| `obtenerHistorialPermisosMedicos(rut)` | Historial del socio |
| `eliminarPermisoMedico(id)` | Anula permiso en estado Solicitado |

### Credenciales
| Función | Descripción |
|---|---|
| `consultarIdCredencialBackend(rutConsultante, rutBuscado)` | Consulta con validación de rol |
| `obtenerEstadoCredencialPorRut(rut)` | Estado actual de credencial |
| `verificarCambiosCredenciales()` | Trigger: detecta cambios y notifica |
| `enviarNotificacionCredencial(correo, nombre, estado, rut)` | Envía email estilizado por estado |

### Sistema QR
| Función | Descripción |
|---|---|
| `validarUsuarioQR(rut)` | Verifica que el RUT existe y está ACTIVO |
| `checkinQR(rut, nombreAsamblea)` | Registra asistencia con control de duplicados |

### Utilidades y Herramientas
| Función | Descripción |
|---|---|
| `generarLinksRegistroYQR()` | Genera links de registro y QR para usuarios sin ellos |
| `configurarTriggers()` | Configura los 6 triggers del sistema |
| `enviarCorreoEstilizado(...)` | Función centralizada de envío con plantilla HTML |
| `subirArchivoConPermisos(...)` | Sube archivo a Drive con permisos silenciosos (3 intentos) |
| `validarCorreosParaPermisos(...)` | Prepara lista de correos para permisos de archivo |
| `formatRutServer(rut)` | Formatea RUT con puntos y guión |
| `cleanRut(rut)` | Limpia RUT a solo dígitos + DV |
| `esCorreoValido(correo)` | Valida formato de correo electrónico |
| `generarCodigoAsamblea(fecha)` | Genera código `YYYY_MM` desde fecha |
| `generarCodigoAsambleaEvento(fecha)` | Genera código `YYYY_MM_DD` desde fecha de evento |

---

## 🎨 Tecnologías

### Frontend
- **HTML5 + CSS3** en un único archivo `Index.html`
- **TailwindCSS** via CDN
- **JavaScript** Vanilla (sin frameworks)
- **SweetAlert2** para modales y alertas
- **Google Fonts:** Inter
- **Material Icons Round**
- **Diseño:** Glassmorphism, gradientes, dark header, paneles con efecto cristal

### Backend
- **Google Apps Script** (V8 runtime)
- **Google Sheets** como base de datos (6 spreadsheets)
- **Google Drive** para almacenamiento de archivos
- **Drive API Avanzada** habilitada (servicio Drive v2 para permisos silenciosos)
- **MailApp / GmailApp** para envío de correos HTML
- **CacheService** para caché de usuarios (10 minutos)
- **LockService** para control de concurrencia (30 segundos)

---

## ⚠️ Consideraciones Técnicas Importantes

### Zona horaria
Las fechas sin componente de hora (ej: `"2026-02-25"`) se interpretan como medianoche UTC, lo que puede generar desfase de 1 día en Chile (UTC-3/UTC-4). La solución establecida es anclar a mediodía: `new Date(fecha + 'T12:00:00')`.

### Coerción de tipos en Google Sheets
Google Sheets convierte automáticamente ciertas cadenas de texto a fechas (ej: `"2026-02"`). Para evitarlo se usa `setNumberFormat('@STRING@')` en columnas sensibles (como `MES_APELACION`) y `getDisplayValues()` en lugar de `getValues()` cuando se lee texto formateado.

### Template literals en Apps Script
Los backticks (`` ` ``) pueden generar errores de cadena no cerrada en ciertos contextos de deployment. Para HTML complejo se prefiere concatenación con `+`.

### Función `configurarTriggers()`
**Elimina todos los triggers existentes** antes de recrearlos. Nunca ejecutar sin revisar que la función refleje los schedules de producción correctos.

### Función `subirArchivoConPermisos()`
Es compartida por Justificaciones, Apelaciones y Permisos Médicos. Cualquier modificación afecta los tres módulos simultáneamente.

### `obtenerUsuarioPorRut()`
Utiliza `CacheService` con TTL de 10 minutos. Si se actualizan datos de un usuario, se debe invalidar el caché con `cache.remove('user_' + rutLimpio)`.

### `verificarCambiosCredenciales()` — Primera ejecución
En la primera ejecución solo inicializa el estado anterior (columna J) sin enviar notificaciones. Comportamiento esperado y documentado.

### `localStorage` y deployments múltiples
`localStorage` no se comparte entre distintos deployments de Apps Script. El módulo QR debe estar integrado en el mismo deployment que el sistema principal para funcionar correctamente.

---

## 🚀 Deployment

La aplicación se sirve mediante `doGet(e)`:
- Sin parámetros GET → sirve `Index.html` (aplicación principal)
- Con parámetros `action`, `rut` o `asamblea` → sirve `QR_Access.html` (control de asistencia QR)

**URL base del Web App:**
```
https://script.google.com/a/~/macros/s/AKfycbzrmy_GgdzMpOLfycvxxUPHU6iyuL9Jv6As_4kxG7mG8oQ4RbV-ALUZw0oeSJnqbvvc/exec
```

---

## 👨‍💻 Desarrollo

**Organización:** Sindicato SLIM N°3  
**Repositorio:** `Daroce12/SLIMAPP`  
**Rama principal:** `main`

---

*Proyecto privado — Todos los derechos reservados*
