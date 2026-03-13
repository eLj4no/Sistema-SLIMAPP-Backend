# 🏢 SLIMAPP — Sistema de Gestión Sindical
### Sindicato SLIM N°3

---

## 📋 Descripción

**SLIMAPP** es una aplicación web desarrollada para la gestión integral del **Sindicato SLIM N°3**, construida sobre **Google Apps Script**, **Google Sheets** como base de datos y **Google Drive** para almacenamiento de archivos. Se sirve como **Google Web App** y es accesible desde cualquier dispositivo con navegador.

---

## 📁 Estructura del Proyecto

```
SLIMAPP/
├── Code.gs              # Backend completo (Google Apps Script)
├── Index.html           # Frontend principal (HTML + CSS + JS en un solo archivo)
├── QR_Access.html       # Página de vinculación QR personal del socio
├── QR_Asistencia.html   # Página de registro QR en punto de control
└── README.md            # Este archivo
```

---

## 🗄️ Arquitectura de Base de Datos

El sistema utiliza **7 Google Spreadsheets independientes**, cada uno dedicado a un módulo.

| Clave CONFIG | ID Spreadsheet | Nombre hoja principal | Descripción |
|---|---|---|---|
| `USUARIOS` | `1m7KLd3b3BzKOAI10I5E32MVf_L34XWAGFonhTg37TVM` | `BD_SLIMAPP` | Registro maestro de socios |
| `JUSTIFICACIONES` | `1Hwbly__MXjl9uwJb-spXdah-R3v9SAMOCFHem92uOUg` | `BD_JUSTIFICACIONES` | Justificaciones de inasistencia |
| `APELACIONES` | `11nrvVsf84THWQ7j6NfAr_unyIcBV7aykxACS8R27PwE` | `BD_APELACIONES` | Apelaciones de multas |
| `PRESTAMOS` | `1h-_sJD4rOCuMjlfSouP7a6gfoodHyzI4MOBRUyOW5XU` | `BD_PRESTAMOS` | Solicitudes de préstamos |
| `PERMISOS_MEDICOS` | `1VYfm7cOgL3mVfVoI8DubIm8WG2srzQw9a6DtIEs3UMM` | `BD_Permisos medicos` | Permisos médicos |
| `CREDENCIALES` | `1HVyPxdYKuvIybeOCAPwAJaVHwlxEuOik4YW0XOXBE5o` | `IMPRESION` | Estado de credenciales sindicales |
| `ASISTENCIA` | `1SRQ8Mlc6bBdb0mitAfn4I-EUAS4BOrZRbqS9YAmg3Sk` | `BD_ASISTENCIA` | Registro de asistencia a asambleas |

### Hojas auxiliares
- `CONFIG_JUSTIFICACIONES` — Configuración del switch de habilitación y fecha de evento
- `Validación-Prestamos` — Hoja de validación manual de préstamos por directiva
- `Registros-eliminados` — Respaldo de préstamos eliminados
- `HISTORIAL_CREDENCIALES` — Log de cambios de estado de credenciales
- `PUNTOS_CONTROL` — Registro de puntos de control QR habilitados

### Estructura `BD_ASISTENCIA`

| Col | Campo | Descripción |
|---|---|---|
| A | `FECHA_HORA` | Fecha y hora del registro (`dd/MM/yyyy HH:mm:ss`) |
| B | `RUT` | RUT del socio |
| C | `NOMBRE` | Nombre completo del socio |
| D | `ASAMBLEA` | Nombre del punto de control / asamblea |
| E | `TIPO_ASISTENCIA` | `Asistencia QR`, `PRESENCIAL` o `VIRTUAL` |
| F | `GESTION` | `Sistema`, nombre del dirigente o admin |
| G | `CODIGO_TEMP` | Código temporal (uso interno) |
| H | `NOTIF_CORREO` | Estado de notificación: vacío → `ENVIADO` o `SIN CORREO` |

### Carpetas Google Drive

| Módulo | Carpeta ID |
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
- Visualización del estado de credencial sindical con badge de color
- Visualización del estado en negociación colectiva

### 2. 📋 Justificaciones
- Formulario de envío de justificación de inasistencia
- Adjunto de archivo (PDF, imagen) hasta 15 MB
- Validación de restricción por evento (`Fecha_Evento` en CONFIG) o por mes (fallback)
- Código de asamblea en formato `YYYY_MM_DD` (modo evento) o `YYYY_MM` (modo mes)
- Historial de justificaciones del socio con estado de revisión
- Switch de habilitación/deshabilitación por admin

### 3. ⚖️ Apelaciones de Multas
- Formulario de apelación con motivo, comprobante de pago y liquidación de sueldo
- Opción de adjuntar comprobante de devolución posterior
- Historial de apelaciones con estado y enlace a archivos
- Posibilidad de anular apelaciones en estado `PENDIENTE`

### 4. 💰 Préstamos
- Formulario de solicitud con monto (hasta UF 10), cuotas (hasta 10) y medio de pago
- Simulador de cuotas según monto y cantidad seleccionada
- Historial de préstamos con estado y detalle de cuotas
- Edición de préstamos en estado `PENDIENTE` (monto, cuotas, medio de pago)
- Eliminación de préstamos en estado `PENDIENTE`
- Notificaciones automáticas al cambiar el estado (trigger diario)

### 5. 🏥 Permisos Médicos
- Formulario de solicitud de permiso médico (con o sin goce de sueldo)
- Adjunto de certificado médico
- Historial de permisos con estado
- Posibilidad de anular permisos en estado `PENDIENTE`

### 6. 📱 Registro Asistencia (QR)
- Visualización del historial personal de asistencias a asambleas
- Muestra: fecha/hora, asamblea, tipo de asistencia, registrado por
- Compatible con registros de tipo QR, PRESENCIAL y VIRTUAL
- Notificación por correo enviada automáticamente a las 20:00 hrs del día del registro

### 7. 🪪 Estado Credencial
- Visualización del estado actual de la credencial sindical
- Badge con colores según estado: `DISPONIBLE`, `ENTREGADO`, `SOLICITADO`, `NO VIGENTE`, `DATOS INCORRECTOS`, `REIMPRIMIR`
- Información detallada del significado de cada estado
- Notificación automática semanal al detectar cambios (trigger dominical)

### 8. 👨‍💼 Gestión de Socios (DIRIGENTE / ADMIN)
- Panel para gestionar trámites a nombre de socios
- **Tabs disponibles:** Préstamos, Justificaciones, Apelaciones, Permisos Médicos
- Búsqueda de socios por RUT con validación de rol (DIRIGENTE solo puede gestionar SOCIOs)
- **ConsultaID:** consulta del ID de credencial y código QR de un socio
  - Muestra: RUT, nombre, cargo, estado, rol, ID credencial, imagen QR
  - Restricción: DIRIGENTE no puede consultar datos de otros DIRIGENTES ni ADMINs
- **Registro manual de asistencia:** ingreso de asistencia PRESENCIAL o VIRTUAL a nombre de un socio

### 9. ⚙️ Panel Admin (ADMIN)
- Switch de habilitación del módulo de justificaciones
- Configuración de `Fecha_Evento` para asambleas (genera código `YYYY_MM_DD` automático)
- Generación de informe de préstamos solicitados (Excel adjunto por correo)
- Configuración de triggers (`configurarTriggers`)
- Herramientas de mantenimiento y corrección de permisos Drive

### 10. 📲 Sistema QR — Vinculación (`QR_Access.html`)
- Página independiente para **vincular el dispositivo personal** del socio
- Parámetro `action=register` + `rut=...`: vincula el RUT al `localStorage` del dispositivo
- Vista por defecto (sin parámetros): muestra estado del dispositivo (vinculado / sin vincular)
- Parámetro `action=checkin` + `asamblea=...`: redirige al flujo de registro en punto de control (requiere vinculación previa)

### 11. 🏁 Sistema QR — Punto de Control (`QR_Asistencia.html`)
- Página independiente activada por parámetro `control=NOMBRE_PUNTO` en la URL
- Lee el RUT vinculado en el `localStorage` del dispositivo
- Llama a `registrarAsistencia(rut, control)` en el backend
- Valida que el dispositivo esté vinculado y que no exista registro duplicado
- Muestra resultado: éxito con datos del socio, advertencia si ya fue registrado, o error descriptivo
- Mensaje informativo sobre envío de confirmación por correo (delegado al trigger de las 20:00)

---

## 🚀 Routing `doGet(e)`

La función `doGet` enruta las peticiones GET a tres destinos distintos:

| Condición | Archivo servido | Uso |
|---|---|---|
| Parámetro `control` presente **o** `action=checkin` | `QR_Asistencia.html` | Registro en punto de control QR |
| Parámetro `action`, `rut` o `asamblea` presentes | `QR_Access.html` | Vinculación QR personal del socio |
| Sin parámetros GET | `Index.html` | Aplicación principal |

**URL base del Web App:**
```
https://script.google.com/a/~/macros/s/AKfycbzrmy_GgdzMpOLfycvxxUPHU6iyuL9Jv6As_4kxG7mG8oQ4RbV-ALUZw0oeSJnqbvvc/exec
```

---

## 🔐 Sistema de Permisos de Archivos (Google Drive)

Toda subida de archivos pasa por la función centralizada `subirArchivoConPermisos`, que:

1. Valida tamaño (máximo 15 MB)
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
| Asistencia QR | Verde | `#10b981` |

### Notificaciones de Asistencia

El envío de correo de confirmación de asistencia está **desacoplado del registro**. El flujo es:

1. `registrarAsistencia()` escribe la fila en `BD_ASISTENCIA` con columna H vacía y retorna inmediatamente
2. El trigger diario de las 20:00 ejecuta `verificarNotificacionesAsistencia()`
3. La función busca filas con columna H vacía, obtiene el correo del socio desde `BD_SLIMAPP`, envía el correo y marca la columna H como `ENVIADO` o `SIN CORREO`

### Lógica de envío en gestión de dirigente

Cuando un **DIRIGENTE o ADMIN** realiza un trámite a nombre de un socio, el sistema envía:

1. **Correo al dirigente** — asunto "Respaldo Gestión [Módulo]", con todos los detalles del trámite y enlace al archivo
2. **Copia al socio** (si tiene correo válido registrado) — mismo contenido con mensaje personalizado indicando que fue gestionado por un dirigente

Si el socio **no tiene correo registrado**, solo se envía al dirigente y el sistema continúa sin error.

---

## ⚡ Triggers Automáticos

Configurados mediante `configurarTriggers()` (ejecutar manualmente una sola vez).

> ⚠️ **Advertencia:** esta función **elimina todos los triggers existentes** antes de recrearlos. Verificar que los schedules del código coincidan con los de producción antes de ejecutar.

| Función | Frecuencia | Descripción |
|---|---|---|
| `verificarCambiosJustificaciones` | Cada 8 horas | Detecta cambios de estado y notifica al socio |
| `verificarCambiosApelaciones` | Cada 8 horas | Detecta cambios de estado y notifica al socio |
| `procesarValidacionPrestamos` | Diario 8 AM | Sincroniza validaciones y notifica cambios de estado |
| `verificarCambiosPrestamos` | Diario 8 AM | Verificación adicional de cambios en préstamos |
| `procesarPermisosComprobantesDevolucion` | Cada 1 hora | Otorga permisos a comprobantes de devolución subidos |
| `verificarCambiosCredenciales` | Domingos 8 AM | Detecta cambios de estado en credenciales y notifica |
| `verificarNotificacionesAsistencia` | Diario 20:00 | Envía correos de confirmación de asistencias pendientes |

---

## 🛠️ Funciones Principales del Backend (`Code.gs`)

### Autenticación y Usuarios
| Función | Descripción |
|---|---|
| `validarUsuario(rut, password)` | Login con RUT e ID credencial |
| `obtenerDatosUsuario(rut)` | Datos completos del perfil |
| `actualizarDatoUsuario(rut, campo, valor)` | Edición de campos editables |
| `obtenerUsuarioPorRut(rut)` | Búsqueda con caché CacheService (TTL 10 min) |
| `recuperarContrasena(rut)` | Consulta de correo registrado |
| `enviarContrasenaCorreo(rut)` | Envío de contraseña por correo |
| `verificarRolUsuario(rut, roles)` | Validación de permisos por rol |

### Justificaciones
| Función | Descripción |
|---|---|
| `enviarJustificacion(...)` | Crea registro, sube archivo, envía correos |
| `validarJustificacionMesActual(rut)` | Valida restricción por evento o por mes |
| `obtenerHistorialJustificaciones(rut)` | Historial del socio |
| `obtenerEstadoSwitchJustificaciones()` | Lee CONFIG: habilitado + fecha evento |
| `toggleSwitchJustificaciones(estado, fecha)` | Actualiza switch (ADMIN) |
| `verificarCambiosJustificaciones()` | Trigger: detecta cambios y envía notificaciones |

### Apelaciones
| Función | Descripción |
|---|---|
| `enviarApelacion(...)` | Crea registro con archivos adjuntos |
| `obtenerHistorialApelaciones(rut)` | Historial del socio |
| `anularApelacion(id, rut)` | Anula apelación en estado PENDIENTE |
| `verificarCambiosApelaciones()` | Trigger: detecta cambios y notifica |

### Préstamos
| Función | Descripción |
|---|---|
| `solicitarPrestamo(...)` | Crea solicitud de préstamo |
| `obtenerHistorialPrestamos(rut)` | Historial del socio |
| `editarPrestamo(...)` | Edita préstamo en estado PENDIENTE |
| `eliminarPrestamo(id, rut)` | Elimina préstamo en estado PENDIENTE |
| `procesarValidacionPrestamos()` | Trigger: sincroniza validaciones |
| `verificarCambiosPrestamos()` | Trigger: detecta cambios y notifica |

### Permisos Médicos
| Función | Descripción |
|---|---|
| `enviarPermisoMedico(...)` | Crea registro, sube certificado, envía correos |
| `obtenerHistorialPermisosMedicos(rut)` | Historial del socio |
| `anularPermisoMedico(id, rut)` | Anula permiso en estado PENDIENTE |

### Asistencia QR
| Función | Descripción |
|---|---|
| `registrarAsistencia(rutInput, nombreControl)` | Registro optimizado con caché + lock reducido |
| `obtenerHistorialAsistencia(rut)` | Historial de asistencias del socio |
| `verificarNotificacionesAsistencia()` | Trigger 20:00: envía correos a registros con columna H vacía |
| `validarDispositivoQR(rut)` | Valida que el RUT pertenece a un socio activo (vinculación) |

### Credenciales
| Función | Descripción |
|---|---|
| `obtenerEstadoCredencial(rut)` | Lee estado desde hoja IMPRESION |
| `verificarCambiosCredenciales()` | Trigger dominical: detecta cambios y notifica |

### Utilidades y Admin
| Función | Descripción |
|---|---|
| `generarLinksRegistroYQR()` | Genera links de registro y QR para usuarios sin ellos |
| `configurarTriggers()` | Configura los 7 triggers del sistema |
| `enviarCorreoEstilizado(...)` | Función centralizada de envío con plantilla HTML |
| `subirArchivoConPermisos(...)` | Sube archivo a Drive con permisos silenciosos (3 intentos) |
| `validarCorreosParaPermisos(...)` | Prepara lista de correos para permisos de archivo |
| `corregirPermisosJustificacionesExistentes()` | Herramienta de corrección masiva de permisos (manual) |
| `formatRutServer(rut)` | Formatea RUT con puntos y guión |
| `cleanRut(rut)` | Limpia RUT a solo dígitos + DV |
| `esCorreoValido(correo)` | Valida formato de correo electrónico |
| `generarCodigoAsamblea(fecha)` | Genera código `YYYY_MM` desde fecha |
| `generarCodigoAsambleaEvento(fecha)` | Genera código `YYYY_MM_DD` desde fecha de evento |

---

## ⚡ Optimización de Concurrencia — `registrarAsistencia()`

La función de registro QR fue optimizada para soportar alta concurrencia simultánea:

| Aspecto | Antes | Después |
|---|---|---|
| Búsqueda de usuario | Dentro del lock (lee BD_SLIMAPP en cada llamada) | Fuera del lock con `CacheService` (TTL 10 min) |
| Envío de correo | Dentro del lock (~7-10 seg bloqueado) | Delegado al trigger de las 20:00 |
| Tiempo de lock estimado | 7-10 segundos | 1.5-2 segundos |
| Capacidad concurrente estimada | 3-4 usuarios simultáneos | ~15-20 usuarios simultáneos |

---

## 🎨 Tecnologías

### Frontend
- **HTML5 + CSS3** en archivos únicos por vista
- **TailwindCSS** via CDN
- **JavaScript** Vanilla (sin frameworks)
- **SweetAlert2** para modales y alertas
- **Google Fonts:** Inter
- **Material Icons Round**
- **Diseño:** Glassmorphism, gradientes, dark header, paneles con efecto cristal

### Backend
- **Google Apps Script** (V8 runtime)
- **Google Sheets** como base de datos (7 spreadsheets)
- **Google Drive** para almacenamiento de archivos
- **Drive API Avanzada** habilitada (servicio Drive v2 para permisos silenciosos)
- **MailApp / GmailApp** para envío de correos HTML estilizados
- **CacheService** para caché de usuarios (TTL 10 minutos)
- **LockService** para control de concurrencia (timeout 30 segundos)
- **QuickChart API** (`quickchart.io/qr`) para generación de códigos QR

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
Utiliza `CacheService` con TTL de 10 minutos. Si se actualizan datos de un usuario en la hoja, el caché se invalida automáticamente en ese plazo. Para forzar invalidación inmediata: `cache.remove('user_' + rutLimpio)`.

### `verificarCambiosCredenciales()` — Primera ejecución
En la primera ejecución solo inicializa el estado anterior (columna J) sin enviar notificaciones. Comportamiento esperado y documentado.

### `localStorage` y deployments múltiples
`localStorage` no se comparte entre distintos deployments de Apps Script. Los archivos `QR_Access.html` y `QR_Asistencia.html` **deben estar integrados en el mismo deployment** que la aplicación principal para que la vinculación del dispositivo funcione correctamente.

### `appendRow` y orden de registros
`appendRow` siempre escribe después de la última fila con datos, independiente del tipo de registro. Esto es seguro para datos mixtos (QR, PRESENCIAL, VIRTUAL) en `BD_ASISTENCIA`.

### Fórmulas con `=IMAGE()` en Sheets
`getValues()` y `getDisplayValues()` retornan vacío para celdas con fórmulas `=IMAGE(...)`. Usar `getFormulas()` para leer el contenido de estas celdas.

---

## 👨‍💻 Desarrollo

**Organización:** Sindicato SLIM N°3
**Repositorio:** `Daroce12/SLIMAPP`
**Rama principal:** `main`

---

*Proyecto privado — Todos los derechos reservados*
