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

El sistema utiliza **8 Google Spreadsheets independientes**, cada uno dedicado a un módulo.

| Clave CONFIG | ID Spreadsheet | Nombre hoja principal | Descripción |
|---|---|---|---|
| `USUARIOS` | `1m7KLd3b3BzKOAI10I5E32MVf_L34XWAGFonhTg37TVM` | `BD_SLIMAPP` | Registro maestro de socios |
| `JUSTIFICACIONES` | `1Hwbly__MXjl9uwJb-spXdah-R3v9SAMOCFHem92uOUg` | `BD_JUSTIFICACIONES` | Justificaciones de inasistencia |
| `APELACIONES` | `11nrvVsf84THWQ7j6NfAr_unyIcBV7aykxACS8R27PwE` | `BD_APELACIONES` | Apelaciones de multas |
| `PRESTAMOS` | `1h-_sJD4rOCuMjlfSouP7a6gfoodHyzI4MOBRUyOW5XU` | `BD_PRESTAMOS` | Solicitudes de préstamos |
| `PERMISOS_MEDICOS` | `1VYfm7cOgL3mVfVoI8DubIm8WG2srzQw9a6DtIEs3UMM` | `BD_Permisos medicos` | Permisos médicos |
| `CREDENCIALES` | `1HVyPxdYKuvIybeOCAPwAJaVHwlxEuOik4YW0XOXBE5o` | `IMPRESION` | Estado de credenciales sindicales |
| `ASISTENCIA` | `1SRQ8Mlc6bBdb0mitAfn4I-EUAS4BOrZRbqS9YAmg3Sk` | `BD_ASISTENCIA` | Registro de asistencia a asambleas |
| `GAMIFICACION` | `1SHDIhGv6XOc30Epm4vdusp3QGVD-pWzhIwzeD6iqbXQ` | `BD_GAMIFICACION` | Progreso y logros SLIM Quest |

### Hojas auxiliares

| Hoja | Spreadsheet | Descripción |
|---|---|---|
| `CONFIG_JUSTIFICACIONES` | JUSTIFICACIONES | Switch de habilitación y fecha de evento/asamblea |
| `Validación-Prestamos` | PRESTAMOS | Validación manual de préstamos por la directiva |
| `Registros-eliminados` | PRESTAMOS | Respaldo de préstamos eliminados |
| `HISTORIAL_CREDENCIALES` | CREDENCIALES | Log de cambios de estado de credenciales |
| `PUNTOS_CONTROL` | ASISTENCIA | Registro de puntos de control QR habilitados |
| `BANCO_PREGUNTAS` | GAMIFICACION | Banco de preguntas del quiz SLIM Quest |

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

### Estructura `BD_GAMIFICACION`

| Col | Campo | Descripción |
|---|---|---|
| A | `RUT` | RUT limpio del socio |
| B | `NOMBRE` | Nombre del socio |
| C | `XP_TOTAL` | Puntos de experiencia acumulados |
| D | `GRADO` | Grado actual (Aspirante → Dirigente) |
| E | `LOGROS` | Array JSON de logros desbloqueados |
| F | `QUIZ_ULTIMO_DIA` | Fecha del último quiz completado |
| G | `RACHA_ACTUAL` | Días consecutivos completando el quiz |
| H | `ULTIMA_ACTIVIDAD` | Última fecha de actividad |
| I | `QUIZ_ULTIMO_DIA` | (alias interno de control de racha) |
| J | `RACHA_MAX` | Racha máxima histórica |
| K | `ESTADO` | `ACTIVO` o `DESVINCULADO` |
| L | `QUIZZES_COMPLETADOS` | Contador total de quizzes completados |

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
| `TESTING` | Acceso restringido: solo `Mis Datos` y `SLIM Quest` |

---

## 📦 Módulos del Sistema

### 1. 👤 Mis Datos
- Visualización y edición de datos personales del socio (correo, teléfono, banco, tipo de cuenta, número de cuenta)
- Badge de estado de credencial (VIGENTE / VENCIDA / PENDIENTE / REVOCADA)
- Validación de RUT en tiempo real mediante módulo 11
- Integración con `CacheService` (TTL 10 min) para optimizar consultas

### 2. 📝 Justificaciones de Inasistencia
- Envío de justificación de inasistencia a asamblea con adjunto de documento (JPG, PNG, PDF — máx. 15 MB)
- Lógica de habilitación dual: switch manual en hoja `CONFIG_JUSTIFICACIONES` + restricción por `Fecha_Evento` (genera código `YYYY_MM_DD` automático)
- Historial de justificaciones del socio con estados y enlace a documento
- Switch controlado por admin desde Panel Admin
- Notificaciones automáticas por correo (socio + representante legal) vía `verificarCambiosJustificaciones`

### 3. ⚖️ Apelaciones de Multas
- Envío de apelación de multa con comprobante de pago y liquidación de sueldo (JPG, PNG, PDF — máx. 15 MB)
- Seguimiento de estado: Pendiente → En revisión → Aprobada / Rechazada
- Comprobante de devolución generado por la directiva y subido a Drive (enlace entregado al socio)
- Procesamiento en background con `procesarPermisosComprobantesDevolucion` (guard de tiempo de ejecución, tracking por fila, distinción error permanente vs. transitorio)
- Historial de apelaciones con acceso a documentos adjuntos

### 4. 💰 Préstamos
- Solicitud de préstamo (educacional / emergencia / especial) con formulario completo
- Validación manual por directiva en hoja `Validación-Prestamos`
- Cambio automático a `Pagado` al cumplirse la fecha de término (`verificarCambiosPrestamos`)
- Generación y envío de informe Excel de préstamos solicitados por correo al admin
- Switch de habilitación controlado por admin
- Notificaciones automáticas de cambios de estado al socio

### 5. 🏥 Permisos Médicos
- Tres tipos de permiso con goce de remuneración:
  - Media jornada (examen médico)
  - 1 día completo (atención fuera de ciudad — requiere especificar ciudad y centro médico)
  - 3 días hábiles (accidente o enfermedad grave de familiar directo)
- Adjunto de documento médico opcional durante la solicitud (máx. 15 MB)
- Notificación al representante legal con columna `NOTIFICADO_REP_LEGAL` y retry cada 30 min (`reintentarNotificacionRepLegal`)
- Notificación al socio con columna `NOTIFICADO_SOCIO` y retry cada 30 min (`reintentarNotificacionSocio`)
- Switch de habilitación controlado por admin

### 6. 📋 Contrato Colectivo
- Visualización del Contrato Colectivo vigente (Sindicato SLIM N°3 — ISS, vigencia hasta diciembre 2029)
- Navegación mediante acordeón de 7 capítulos
- Control de scroll en apertura de sección
- Switch de habilitación controlado por admin

### 7. 🧮 Calculadora de Horas Extraordinarias
- Calculadora oficial basada en fórmula DT: `(Base ÷ 30 × 30) ÷ (h/sem × 4) × 1.5`
- Dos perfiles de cargo: Limpieza y Guardias de Seguridad
- Valores remuneracionales actualizados para 2026
- Switch de habilitación controlado por admin

### 8. 🎮 SLIM Quest _(beta)_
- Sistema de gamificación para socios con 6 grados de progresión:

| Grado | XP Requerido | Icono |
|---|---|---|
| Aspirante | 0 – 1.500 | 🌱 |
| Aprendiz | 1.501 – 4.500 | ⚙️ |
| Trabajador | 4.501 – 10.000 | 🔩 |
| Defensor | 10.001 – 18.000 | 🛡️ |
| Negociador | 18.001 – 30.000 | ⚖️ |
| Dirigente | 30.001+ | 🏆 |

- Quiz diario de preguntas sobre derechos laborales y contrato colectivo (3 niveles: BASICO / INTERMEDIO / AVANZADO)
- Sistema de rachas con bonos de XP en hitos: 3, 7, 14, 21, 30, 60, 100 días consecutivos
- Nivel Secreto (`DIRIGENTE`): acceso exclusivo para socios en grado máximo
- Leaderboard con top 10 de socios más activos
- Logros desbloqueables almacenados en JSON por socio
- Sincronización automática de socios desde `BD_SLIMAPP` vía trigger diario a la 1:00 AM
- Switch de habilitación controlado por admin (bypass para ADMIN)

### 9. ⚙️ Panel Admin _(solo ADMIN)_
- Switch de habilitación/deshabilitación de todos los módulos con switch:
  - Justificaciones, Préstamos, Permisos Médicos, Contrato Colectivo, Calculadora HE, SLIM Quest, Registro Asistencia
- Badges de estado en tiempo real para cada módulo (cargados vía `obtenerEstadosSwitchDashboard`)
- Configuración de `Fecha_Evento` para asambleas (genera código `YYYY_MM_DD`)
- Generación y envío de informe de préstamos solicitados (Excel adjunto por correo)
- Herramientas de mantenimiento y corrección de permisos en Drive
- Configuración de triggers (`configurarTriggers`)

### 10. 🪪 Consulta ID / Credencial _(ADMIN / DIRIGENTE)_
- Búsqueda de socio por RUT con visualización de estado de credencial
- Badge de estado: VIGENTE, VENCIDA, PENDIENTE o REVOCADA
- Historial de cambios de credencial
- Verificación semanal automática de cambios (`verificarCambiosCredenciales`, domingos 8 AM)

### 11. 👥 Gestión de Socios _(DIRIGENTE / ADMIN)_
- Registro de asistencia manual (PRESENCIAL / VIRTUAL) a nombre de socios
- Búsqueda de socio por RUT para trámites dirigidos
- Gestión integral de trámites a nombre de socios (justificaciones, apelaciones, permisos médicos)

### 12. 📲 Sistema QR — Vinculación Personal (`QR_Access.html`)
- Página independiente para vincular el dispositivo personal del socio
- Parámetro `action=register` + `rut=...`: vincula el RUT al `localStorage` del dispositivo
- Vista por defecto (sin parámetros): muestra estado del dispositivo (vinculado / sin vincular)
- Parámetro `action=checkin` + `asamblea=...`: redirige al flujo de registro en punto de control

### 13. 🏁 Sistema QR — Punto de Control (`QR_Asistencia.html`)
- Página independiente activada por parámetro `control=NOMBRE_PUNTO` en la URL
- Lee el RUT vinculado en el `localStorage` del dispositivo
- Llama a `registrarAsistencia(rut, control)` en el backend
- Valida vinculación previa y evita registros duplicados
- Muestra resultado: éxito con datos del socio, advertencia si ya fue registrado, o error descriptivo
- Notificación por correo delegada al trigger de las 20:00

---

## 🔀 Routing `doGet(e)`

La función `doGet` enruta las peticiones GET a tres destinos distintos:

| Condición | Archivo servido | Uso |
|---|---|---|
| Parámetro `control` presente **o** `action=checkin` | `QR_Asistencia.html` | Registro en punto de control QR |
| Parámetros `action`, `rut` o `asamblea` presentes | `QR_Access.html` | Vinculación QR personal del socio |
| Sin parámetros GET | `Index.html` | Aplicación principal |

**URL base del Web App:**
```
https://script.google.com/a/~/macros/s/AKfycbzrmy_GgdzMpOLfycvxxUPHU6iyuL9Jv6As_4kxG7mG8oQ4RbV-ALUZw0oeSJnqbvvc/exec
```

---

## 🔐 Sistema de Permisos de Archivos (Google Drive)

Toda subida de archivos pasa por la función centralizada `subirArchivoConPermisos`, que:

1. Valida tamaño (máximo **15 MB**)
2. Crea el archivo en la carpeta correspondiente con nombre estandarizado (`PREFIJO-UUID-RUT`)
3. Configura el archivo como privado (`PRIVATE`)
4. Otorga acceso de lectura **silencioso** (sin email de notificación de Drive) a los correos involucrados mediante **3 intentos con fallback**:
   - Intento 1: Drive API Avanzada (`Drive.Permissions.insert` con `sendNotificationEmails: false`)
   - Intento 2: `file.addViewer()` con `Utilities.sleep(1000)`
   - Intento 3: Reintento final con `Utilities.sleep(2000)`
5. Retorna URL de acceso y reporte de permisos otorgados/fallidos

La función `validarCorreosParaPermisos` prepara la lista de correos a los que se otorga acceso y genera alertas si algún correo no es válido.

---

## 🔔 Sistema de Notificaciones por Correo

Todos los correos de notificación utilizan `enviarCorreoEstilizado()` con diseño HTML responsivo, tema de color por módulo y tabla de detalles del trámite. Los correos se envían vía `MailApp` con `name: "Sindicato SLIM N°3"`.

### Flujo de notificación por módulo

| Módulo | Destinatario | Trigger |
|---|---|---|
| Justificaciones | Socio + Rep. Legal | `verificarCambiosJustificaciones` (cada 8 h) |
| Apelaciones | Socio + Rep. Legal | `verificarCambiosApelaciones` (cada 8 h) |
| Préstamos | Socio | `procesarValidacionPrestamos` / `verificarCambiosPrestamos` (diario 8 AM) |
| Permisos Médicos — Rep. Legal | Rep. Legal | Inmediato + retry cada 30 min (`reintentarNotificacionRepLegal`) |
| Permisos Médicos — Socio | Socio | Inmediato + retry cada 30 min (`reintentarNotificacionSocio`) |
| Credenciales | Socio | `verificarCambiosCredenciales` (domingos 8 AM) |
| Asistencia | Socio | `verificarNotificacionesAsistencia` (diario 20:00) |
| SLIM Quest (subida de grado) | Socio | Inmediato al completar quiz |

**Correo representante legal:** `juancarlos.pacheco@cl.issworld.com`

**Centinela `SIN_CORREO`:** registros sin correo válido reciben este valor permanente en su columna de notificación para evitar reintentos infinitos.

---

## ⚡ Triggers Automáticos

Configurados mediante `configurarTriggers()` (ejecutar manualmente una sola vez desde el editor de Apps Script).

> ⚠️ **Advertencia:** esta función **elimina todos los triggers existentes** antes de recrearlos. Verificar que los schedules del código coincidan con los de producción antes de ejecutar. Los triggers personalizados (como `sincronizarSociosGamificacion`) deben añadirse manualmente desde la interfaz de Apps Script o mediante `configurarTriggerGamificacion()`.

| Función | Frecuencia | Descripción |
|---|---|---|
| `verificarCambiosJustificaciones` | Cada 8 horas | Detecta cambios de estado y notifica al socio |
| `verificarCambiosApelaciones` | Cada 8 horas | Detecta cambios de estado y notifica al socio |
| `procesarValidacionPrestamos` | Diario 8 AM | Sincroniza validaciones y notifica cambios de estado |
| `verificarCambiosPrestamos` | Diario 8 AM | Verificación adicional de cambios en préstamos |
| `procesarPermisosComprobantesDevolucion` | Cada 1 hora | Otorga permisos a comprobantes de devolución subidos |
| `verificarCambiosCredenciales` | Domingos 8 AM | Detecta cambios de estado en credenciales y notifica |
| `verificarNotificacionesAsistencia` | Diario 20:00 | Envía correos de confirmación de asistencias pendientes |
| `sincronizarSociosGamificacion` | Diario 1:00 AM | Sincroniza socios nuevos/actualizados con SLIM Quest |
| `reintentarNotificacionRepLegal` | Cada 30 min | Reintenta notificaciones fallidas al Rep. Legal (Permisos Médicos) |
| `reintentarNotificacionSocio` | Cada 30 min | Reintenta notificaciones fallidas al socio (Permisos Médicos) |

---

## 🔧 Switches de Módulos (PropertiesService)

Los módulos habilitables/deshabilitables utilizan `PropertiesService.getScriptProperties()` para persistencia. El valor por defecto (clave ausente) es siempre **habilitado**.

| Clave PropertiesService | Módulo | Función getter | Función toggle |
|---|---|---|---|
| `prestamos_habilitado` | Préstamos | `obtenerEstadoSwitchPrestamos()` | `toggleSwitchPrestamos(estado)` |
| `contrato_colectivo_habilitado` | Contrato Colectivo | `obtenerEstadoSwitchContratoColectivo()` | `toggleSwitchContratoColectivo(estado)` |
| `slimquest_habilitado` | SLIM Quest | `obtenerEstadoSwitchSlimQuest()` | `toggleSwitchSlimQuest(estado)` |
| `calculadora_habilitada` | Calculadora HE | `obtenerEstadoSwitchCalculadora()` | `toggleSwitchCalculadora(estado)` |
| `permisos_medicos_habilitado` | Permisos Médicos | `obtenerEstadoSwitchPermisosMedicos()` | `toggleSwitchPermisosMedicos(estado)` |
| `asistencia_habilitada` | Registro Asistencia | `obtenerEstadoSwitchAsistencia()` | `toggleSwitchAsistencia(estado)` |
| _(hoja CONFIG_JUSTIFICACIONES)_ | Justificaciones | `obtenerEstadoSwitchJustificaciones()` | Desde Panel Admin |

La función `obtenerEstadosSwitchDashboard()` retorna el estado de todos los módulos en una sola llamada para mostrar badges en el dashboard del admin.

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
| `obtenerHistorialJustificaciones(rut)` | Historial del socio |
| `verificarCambiosJustificaciones()` | Trigger: detecta cambios y notifica |
| `obtenerEstadoSwitchJustificaciones()` | Lee estado de hoja `CONFIG_JUSTIFICACIONES` |

### Apelaciones

| Función | Descripción |
|---|---|
| `enviarApelacion(...)` | Crea registro, sube archivos, envía correos |
| `obtenerHistorialApelaciones(rut)` | Historial del socio |
| `verificarCambiosApelaciones()` | Trigger: detecta cambios y notifica |
| `procesarPermisosComprobantesDevolucion()` | Trigger: otorga permisos a comprobantes subidos |

### Préstamos

| Función | Descripción |
|---|---|
| `enviarSolicitudPrestamo(...)` | Crea solicitud en hoja de préstamos |
| `obtenerHistorialPrestamos(rut)` | Historial del socio |
| `procesarValidacionPrestamos()` | Trigger: sincroniza validaciones, notifica cambios |
| `verificarCambiosPrestamos()` | Trigger: marca préstamos vencidos como `Pagado` |
| `generarInformePrestamos(correo)` | Genera Excel y lo envía por correo |

### Permisos Médicos

| Función | Descripción |
|---|---|
| `enviarPermisoMedico(...)` | Crea solicitud con adjunto opcional |
| `obtenerHistorialPermisosMedicos(rut)` | Historial del socio |
| `reintentarNotificacionRepLegal()` | Trigger: reintenta notificaciones fallidas al Rep. Legal |
| `reintentarNotificacionSocio()` | Trigger: reintenta notificaciones fallidas al socio |

### Asistencia

| Función | Descripción |
|---|---|
| `registrarAsistencia(rut, control)` | Registra asistencia QR en `BD_ASISTENCIA` |
| `registrarAsistenciaManual(rut, tipo, asamblea, gestorRut)` | Registro PRESENCIAL o VIRTUAL por dirigente/admin |
| `obtenerHistorialAsistencia(rut)` | Historial del socio |
| `verificarNotificacionesAsistencia()` | Trigger: envía correos de confirmación pendientes |
| `obtenerPuntosControl()` | Lista de puntos de control habilitados |

### Credenciales

| Función | Descripción |
|---|---|
| `buscarSocioPorRut(rut)` | Búsqueda para Consulta ID |
| `verificarCambiosCredenciales()` | Trigger: detecta cambios de estado y notifica al socio |
| `enviarCorreoEstadoCredencial(correo, nombre, estadoNuevo)` | Correo HTML de cambio de credencial |

### SLIM Quest — Gamificación

| Función | Descripción |
|---|---|
| `getProgresoSocio(rut)` | Progreso completo del socio (XP, grado, logros, racha) |
| `guardarXP(rut, cantidad, motivo)` | Acumula XP al socio con LockService |
| `otorgarLogro(rut, codigo, nombre, icono)` | Desbloquea logro al socio |
| `completarQuiz(rut, xpGanado, correctas)` | Procesa resultado de quiz (racha, bonos, nivel) |
| `obtenerPreguntasQuiz(rut, cantidad)` | Obtiene preguntas ponderadas según grado |
| `obtenerPreguntasSecreto(rut, cantidad)` | Preguntas exclusivas para grado Dirigente |
| `getLeaderboard(rut)` | Top 10 socios por XP |
| `sincronizarSociosGamificacion()` | Sincroniza socios desde `BD_SLIMAPP` |
| `inicializarSocioGamificacion_(rut)` | Crea registro inicial de socio en gamificación |

### Auxiliares y Configuración

| Función | Descripción |
|---|---|
| `doGet(e)` | Router principal de la Web App |
| `cleanRut(rut)` | Limpia RUT (elimina puntos, guión, mayúsculas) |
| `esCorreoValido(correo)` | Valida formato de correo electrónico |
| `enviarCorreoEstilizado(...)` | Genera y envía correo HTML con diseño corporativo |
| `subirArchivoConPermisos(...)` | Sube archivo a Drive con permisos otorgados (3 intentos) |
| `validarCorreosParaPermisos(...)` | Prepara lista de correos para otorgar acceso |
| `getSheet(ssKey, sheetKey)` | Obtiene hoja de cálculo por clave de CONFIG |
| `getSpreadsheet(ssKey)` | Obtiene spreadsheet por clave de CONFIG |
| `configurarTriggers()` | Configura todos los triggers automáticos |
| `configurarTriggerGamificacion()` | Configura trigger diario de sincronización SLIM Quest |
| `obtenerEstadosSwitchDashboard()` | Estado de todos los módulos en una llamada |

---

## 🔑 Notas Técnicas Importantes

### Manejo de fechas
- Siempre anclar strings de fecha al mediodía (`T12:00:00`) para evitar desfases UTC → Santiago (UTC-3).
- Usar `getValues()` en lugar de `getDisplayValues()` para columnas de fecha (evita inversión día/mes en formato US).
- Usar `setNumberFormat('@STRING@')` para prevenir que Sheets convierta texto a Date automáticamente.

### Concurrencia
- `LockService` (nivel script) en operaciones críticas. Mantener secciones críticas cortas (~1.5–2 s).
- Operaciones no críticas (como envío de correos) se mueven fuera del lock.
- `CacheService` con TTL de 10 minutos para búsquedas de usuario (`user_RUT`).

### Template literals
- Evitar backticks `` ` `` en Apps Script — usar concatenación con `+` para prevenir errores de string no cerrado.

### Formularios y archivos
- Eliminar `capture="environment"` de inputs de archivo en producción (causaría que móviles abran cámara en lugar de selector de archivos).
- Fórmulas `=IMAGE()` en Sheets: usar `getFormulas()` en lugar de `getValues()` o `getDisplayValues()`.

### Web App URL
- La URL debe usar el formato `/a/~/macros/s/` para deployments de Google Workspace.
- `ScriptApp.getService().getUrl()` no retorna este formato confiablemente.

### Scripts embebidos en Sheets
- Scripts dentro de un archivo Sheets deben usar `getActiveSpreadsheet()` y literales de string, sin referencias al objeto `CONFIG` del proyecto backend.

### Sentinel `SIN_CORREO`
- Registros sin correo válido reciben este valor en columnas de notificación para evitar reintentos infinitos.

### `appendRow`
- Siempre escribe después de la última fila con datos, independiente del tipo de registro. Seguro para datos mixtos.

---

## 🔐 Seguridad y Control de Acceso

- Validación de rol en todas las funciones sensibles mediante `verificarRolUsuario(rut, rolesPermitidos)`.
- Socios desvinculados (`ESTADO = DESVINCULADO`) solo pueden acceder a **Mis Datos** (excepto ADMIN).
- Perfil `TESTING`: acceso exclusivo a **Mis Datos** y **SLIM Quest**.
- El acceso QR requiere vinculación previa del dispositivo vía `QR_Access.html`.
- Los archivos subidos a Drive se configuran como **privados** con permisos silenciosos solo para los involucrados.

---

## 🎨 Stack Frontend

| Tecnología | Uso |
|---|---|
| **TailwindCSS** (CDN) | Framework CSS para todo el diseño |
| **Material Icons Round** (CDN) | Iconografía del sistema |
| **SweetAlert** | Modales de alerta y confirmación |
| **QuickChart API** | Generación de QR (`https://quickchart.io/qr?size=300&text={url}`) |
| **SLIMSound Engine** | Motor de sonidos (Web Audio API) integrado en el frontend |
| **Google Apps Script `google.script.run`** | Comunicación asíncrona frontend → backend |

---

## 🏛️ Directiva Sindicato SLIM N°3

| Cargo | Nombre |
|---|---|
| Presidente | Carlos Orellana G. |
| Tesorero | Franco Collao V. |
| Secretario | Carlos Pacheco M. |
| Directora | Felicita Anartes C. |
| Directora | Sofía Leonardini C. |

**Oficina:** Ahumada 312, Of. 323, Santiago

---

## 👨‍💻 Desarrollo

**Organización:** Sindicato SLIM N°3
**Desarrollador:** Alejandro Peñailillo G. — DUOC UC, Técnico Analista Programador
**Repositorio:** `Daroce12/SLIMAPP`
**Rama principal:** `main`
**Plataforma:** Google Apps Script + Google Workspace

---

*Proyecto privado — Todos los derechos reservados*
