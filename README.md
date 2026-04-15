# 🏢 SLIMAPP — Sistema de Gestión Sindical
### Sindicato SLIM N°3

---

## 📋 Descripción

**SLIMAPP** es una aplicación web desarrollada para la gestión integral del **Sindicato SLIM N°3**, construida sobre **Google Apps Script**, **Google Sheets** como base de datos y **Google Drive** para almacenamiento de archivos. Se sirve como **Google Web App** y es accesible desde cualquier dispositivo con navegador.

Actualmente gestiona aproximadamente **2.230 socios activos** distribuidos en distintas regiones de Chile.

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

| Clave CONFIG | ID Spreadsheet | Hoja principal | Descripción |
|---|---|---|---|
| `USUARIOS` | `1m7KLd3b3BzKOAI10I5E32MVf_L34XWAGFonhTg37TVM` | `BD_SLIMAPP` | Registro maestro de socios |
| `JUSTIFICACIONES` | `1Hwbly__MXjl9uwJb-spXdah-R3v9SAMOCFHem92uOUg` | `BD_JUSTIFICACIONES` | Justificaciones de inasistencia |
| `APELACIONES` | `11nrvVsf84THWQ7j6NfAr_unyIcBV7aykxACS8R27PwE` | `BD_APELACIONES` | Apelaciones de multas |
| `PRESTAMOS` | `1h-_sJD4rOCuMjlfSouP7a6gfoodHyzI4MOBRUyOW5XU` | `BD_PRESTAMOS` | Solicitudes de préstamos |
| `PERMISOS_MEDICOS` | `1VYfm7cOgL3mVfVoI8DubIm8WG2srzQw9a6DtIEs3UMM` | `BD_Permisos medicos` | Permisos médicos |
| `CREDENCIALES` | `1HVyPxdYKuvIybeOCAPwAJaVHwlxEuOik4YW0XOXBE5o` | `IMPRESION` | Estado de credenciales sindicales |
| `ASISTENCIA` | `1SRQ8Mlc6bBdb0mitAfn4I-EUAS4BOrZRbqS9YAmg3Sk` | `BD_ASISTENCIA` | Registro de asistencia a asambleas |
| `GAMIFICACION` | `1SHDIhGv6XOc30Epm4vdusp3QGVD-pWzhIwzeD6iqbXQ` | `BD_GAMIFICACION` | Progreso y logros SLIM Quest |

---

## 📊 Estructura Detallada de cada Base de Datos

### 🗂️ USUARIOS — `BD_SLIMAPP`

Hoja principal: `BD_SLIMAPP` · Hojas auxiliares: `IDENTIFICADOR`, `ENVIADOS`, `PENDIENTES`

| # | Col | Campo | Descripción |
|---|---|---|---|
| 0 | A | `RUT` | RUT del socio (sin puntos ni guión) |
| 1 | B | `RUT VALIDADO` | Resultado de validación: `RUT VÁLIDO` / `RUT INVÁLIDO` |
| 2 | C | `FECHA DE INGRESO` | Fecha de ingreso a la empresa |
| 3 | D | `NOMBRE SISTEMA` | Nombre completo en mayúsculas |
| 4 | E | `CARGO` | Cargo del trabajador (ej. `GUARDIA DE SEGURIDAD`) |
| 5 | F | `CORREO` | Correo electrónico personal |
| 6 | G | `SITE` | Sitio o lugar de trabajo |
| 7 | H | `REGION` | Región del país (ej. `11. VIII Region del Biobío`) |
| 8 | I | `SEXO` | Sexo del socio |
| 9 | J | `ESTADO` | `ACTIVO` / `DESVINCULADO` |
| 10 | K | `DETALLE DESVINCULACION` | Motivo de desvinculación si aplica |
| 11 | L | `ID CREDENCIAL` | Código credencial sindical (ej. `S03678`) — usado como contraseña de login |
| 12 | M | `CORREO REGISTRADO` | Correo con el que el socio se registró en la app |
| 13 | N | `CONTACTO` | Número de teléfono celular (ej. `+56996114947`) |
| 14 | O | `ROL` | Rol en el sistema: `SOCIO` / `DIRIGENTE` / `ADMIN` / `TESTING` |
| 15 | P | `Link Registro` | URL de registro QR individual del socio |
| 16 | Q | `QR Registro` | Fórmula `=IMAGE(...)` con imagen QR del link de registro |
| 17 | R | `BANCO` | Banco del socio (ej. `BANCO ESTADO (Cuenta RUT)`) |
| 18 | S | `TIPO_CUENTA` | Tipo de cuenta bancaria (ej. `CUENTA VISTA`) |
| 19 | T | `NUEMRO_CUENTA` | Número de cuenta bancaria |
| 20 | U | `ESTADO_NEG_COLECT_2026` | Estado en negociación colectiva 2026 |
| 21 | V | `TALLA_POLERA` | Talla de polera para vestuario |
| 22 | W | `TALLA_POLAR` | Talla de polar para vestuario |
| 23 | X | `TALLA_PANTALON` | Talla de pantalón para vestuario |
| 24 | Y | `TALLA_CALZADO` | Número de calzado |
| 25 | Z | `CALZADO_ESPECIAL` | Indica si requiere calzado especial: `SI` / `NO` |
| 26 | AA | `URL_CERT_PIE_DIABETICO` | URL del certificado médico de pie diabético si aplica |

**Hoja `IDENTIFICADOR`** — Espejo de sesión activa del socio (17 columnas: RUT, NOMBRE, CORREO, CONTACTO, ID, REGION, ACTIVO, CODIGO_TEMP, ASAMBLEA, BANCO, TIPO_CUENTA, NUEMRO_CUENTA, CARGO, Talla POLERA, Talla POLAR, Talla PANTALON, TALLA CALZADO)

**Hoja `ENVIADOS`** — Registro de socios que recibieron correo de bienvenida (RUT, NOMBRE, CORREO)

**Hoja `PENDIENTES`** — Socios pendientes de registro en la app (RUT, NOMBRE, CORREO, CELULAR, ID, REGION)

---

### 🗂️ JUSTIFICACIONES — `BD_JUSTIFICACIONES`

Hojas: `BD_JUSTIFICACIONES`, `CONFIG_JUSTIFICACIONES`, `Resp_Justificaciones`

| # | Col | Campo | Descripción |
|---|---|---|---|
| 0 | A | `ID` | UUID único de la solicitud |
| 1 | B | `FECHA` | Fecha y hora de envío de la solicitud |
| 2 | C | `RUT` | RUT del socio |
| 3 | D | `NOMBRE` | Nombre completo del socio |
| 4 | E | `REGION` | Región del socio |
| 5 | F | `MOTIVO` | Tipo de motivo (ej. `Fuerza Mayor`, `Licencia Médica`) |
| 6 | G | `ARGUMENTO` | Texto libre con el argumento del socio |
| 7 | H | `RESPALDO` | URL del documento de respaldo en Google Drive |
| 8 | I | `ESTADO` | Estado actual: `Enviado` / `Aceptado` / `Rechazado` |
| 9 | J | `OBSERVACION` | Observación del dirigente al gestionar |
| 10 | K | `NOTIFICACION` | Estado notificación al socio: `Enviado` / `SIN_CORREO` |
| 11 | L | `ASAMBLEA` | Código de asamblea (formato `YYYY_MM_DD`) |
| 12 | M | `Gestion` | `Socio` / `Dirigente` |
| 13 | N | `Dirigente` | Nombre del dirigente gestor (si aplica) |
| 14 | O | `Correo Dirigente` | Correo del dirigente gestor (si aplica) |

**Hoja `CONFIG_JUSTIFICACIONES`** — Configuración del módulo:
- `Habilitado` (Boolean): Switch de activación del módulo
- `Fecha Límite` (ISO 8601): Fecha tope para enviar justificaciones
- `Fecha_Evento` (Date): Fecha de la asamblea/evento (genera código `YYYY_MM_DD`)

---

### 🗂️ APELACIONES — `BD_APELACIONES`

Hojas: `BD_APELACIONES`, `MASIVAS`, `Resp_Apelaciones`

| # | Col | Campo | Descripción |
|---|---|---|---|
| 0 | A | `ID` | UUID único de la apelación |
| 1 | B | `Fecha Solicitud` | Fecha y hora de envío |
| 2 | C | `Rut` | RUT del socio |
| 3 | D | `Nombre` | Nombre completo del socio |
| 4 | E | `Correo` | Correo del socio |
| 5 | F | `Mes Apelación` | Mes al que corresponde la multa apelada |
| 6 | G | `Tipo Motivo` | Tipo de motivo (ej. `Fuerza Mayor`) |
| 7 | H | `Detalle Motivo` | Descripción libre del motivo |
| 8 | I | `URL Comprobante` | URL del comprobante en Drive |
| 9 | J | `URL Liquidación` | URL de la liquidación de sueldo en Drive |
| 10 | K | `Estado` | Estado actual: `Solicitado` / `En Revisión` / `Aceptado` / `Rechazado` / `Aceptado-Obs` / `Pagado` |
| 11 | L | `Observación` | Observación pública del dirigente |
| 12 | M | `Notificado` | Estado de notificación al socio |
| 13 | N | `Gestión` | `Socio` / `Dirigente` |
| 14 | O | `Nombre Dirigente` | Nombre del dirigente gestor |
| 15 | P | `Correo Dirigente` | Correo del dirigente gestor |
| 16 | Q | `URL Comprobante Devolución` | URL del comprobante de devolución de multa |
| 17 | R | `PERMISO_DEVOLUCION` | Estado del permiso Drive para comprobante: `OK` / `REINTENTO_N` |
| 18 | S | `LOG_PERMISOS` | Log de intentos de otorgamiento de permisos Drive |
| 19 | T | `Observacion Interna` | Nota interna visible solo para dirigentes/admin |
| 20 | U | `SEMANA` | Semana de gestión asignada |

**Hoja `MASIVAS`** — Apelaciones masivas procesadas directamente (RUT, NOMBRE, VALOR, DESCRIPCION, LISTADO, ESTADO, COMPROBANTE)

**Hoja `Resp_Apelaciones`** — Formulario de respaldo heredado (10 columnas: Marca temporal, RUT, NOMBRE, REGIÓN, MES ASAMBLEA, DIRIGENTE, DESCARGOS, LIQUIDACIÓN DE SUELDO, RESPALDO, MOTIVO DE AYUDA)

---

### 🗂️ PRÉSTAMOS — `BD_PRESTAMOS`

Hojas: `BD_PRESTAMOS`, `Validación-Prestamos`, `Registros-eliminados`

| # | Col | Campo | Descripción |
|---|---|---|---|
| 0 | A | `ID` | UUID único del préstamo |
| 1 | B | `Fecha` | Fecha y hora de la solicitud |
| 2 | C | `RUT` | RUT del socio |
| 3 | D | `Nombre` | Nombre completo del socio |
| 4 | E | `Correo` | Correo del socio |
| 5 | F | `Tipo Préstamo` | Tipo: `Vacaciones` / `Emergencia` / `Escolaridad` / etc. |
| 6 | G | `Monto` | Monto solicitado (ej. `$150.000`) |
| 7 | H | `Cuotas` | Número de cuotas para devolución |
| 8 | I | `Medio Pago` | Medio de pago: `Cuenta RUT` / `Transferencia` / etc. |
| 9 | J | `Estado` | Estado: `Solicitado` / `Vigente` / `Rechazado` / `Finalizado` |
| 10 | K | `Fecha Termino` | Fecha estimada de término del préstamo |
| 11 | L | `Gestion` | `Socio` / `Dirigente` |
| 12 | M | `Nombre Dirigente` | Nombre del dirigente gestor |
| 13 | N | `Correo Dirigente` | Correo del dirigente gestor |
| 14 | O | `Informe` | Marca de procesamiento: `OK` cuando fue procesado |
| 15 | P | `Observacion` | Observación adicional del dirigente |

**Hoja `Validación-Prestamos`** — Validación manual por la directiva (ID, Fecha, RUT, Nombre, Validación [`ACEPTADO`/`RECHAZADO`], Observacion, Nombre Informe)

**Hoja `Registros-eliminados`** — Respaldo de préstamos eliminados (mismas 15 columnas de BD_PRESTAMOS sin `Observacion`)

---

### 🗂️ PERMISOS MÉDICOS — `BD_Permisos medicos`

Hoja única: `BD_Permisos medicos`

| # | Col | Campo | Descripción |
|---|---|---|---|
| 0 | A | `ID` | UUID único del permiso |
| 1 | B | `Fecha Solicitud` | Fecha y hora de la solicitud |
| 2 | C | `Rut` | RUT del socio |
| 3 | D | `Nombre` | Nombre completo del socio |
| 4 | E | `Correo` | Correo del socio |
| 5 | F | `Tipo Permiso` | Descripción completa del tipo de permiso seleccionado |
| 6 | G | `Fecha Inicio Permiso` | Fecha de inicio del permiso |
| 7 | H | `Motivo/Detalle` | Descripción libre del motivo |
| 8 | I | `URL Documento Respaldo` | URL del documento en Drive (o `Sin documento`) |
| 9 | J | `Estado` | Estado: `Solicitado` / `Solicitado con Documento` / `Documento Adjuntado` / `Aceptado` / `Rechazado` |
| 10 | K | `Fecha Subida Documento` | Fecha en que se adjuntó el documento (puede ser posterior) |
| 11 | L | `Notificado Rep. Legal` | Boolean — si se notificó al representante legal |
| 12 | M | `Gestión` | `Socio` / `Dirigente` |
| 13 | N | `Nombre Dirigente` | Nombre del dirigente gestor |
| 14 | O | `Correo Dirigente` | Correo del dirigente gestor |
| 15 | P | `NOTIFICADO_SOCIO` | Boolean — si se notificó al socio |

**Tipos de permiso disponibles:**
- Media jornada (Examen médico — CON GOCE DE REMUNERACIÓN)
- Jornada completa (Enfermedad grave hijo menor — CON GOCE)
- 1 día completo (Cláusula Trigésimo Segunda)
- Y otros definidos en el contrato colectivo

---

### 🗂️ ASISTENCIA — `BD_ASISTENCIA`

Hojas: `BD_ASISTENCIA`, `PUNTOS_CONTROL`

| # | Col | Campo | Descripción |
|---|---|---|---|
| 0 | A | `FECHA_HORA` | Fecha y hora del registro (`dd/MM/yyyy HH:mm:ss`) |
| 1 | B | `RUT` | RUT del socio |
| 2 | C | `NOMBRE` | Nombre completo del socio |
| 3 | D | `ASAMBLEA` | Nombre del punto de control / asamblea |
| 4 | E | `TIPO_ASISTENCIA` | `Asistencia QR` / `PRESENCIAL` / `VIRTUAL` |
| 5 | F | `GESTION` | `Sistema` / nombre del dirigente o admin |
| 6 | G | `CODIGO_TEMP` | Código temporal (uso interno para prevenir duplicados) |
| 7 | H | `NOTIF_CORREO` | Estado de notificación: vacío → `ENVIADO` o `SIN_CORREO` |

**Hoja `PUNTOS_CONTROL`** — Registro de puntos de control QR habilitados:

| # | Col | Campo | Descripción |
|---|---|---|---|
| 0 | A | `NOMBRE_PUNTO` | Nombre del punto de control (patrón: `TIPO_DD-MM-YYYY_CODIGO-REGION`) |
| 1 | B | `URL control` | URL de la página QR del punto de control |
| 2 | C | `QR_CODE` | Imagen QR generada |
| 3 | D | `(URL Web App)` | URL base del deployment |
| 4 | E | `HORA_APERTURA` | Hora de apertura del punto de control |
| 5 | F | `HORA_CIERRE` | Hora de cierre del punto de control |
| 6 | G | `TIPO` | Tipo de punto de control |

---

### 🗂️ CREDENCIALES — `IMPRESION`

Hojas: `IMPRESION`, `HISTORIAL_CREDENCIALES`

La hoja `IMPRESION` maneja el estado de credenciales sindicales físicas. Las columnas A–K incluyen:
- `RUT` (A), `ID CREDENCIAL` (B), `ESTADO` (C), `NOMBRE BD` (D), `NOMBRE FORMULARIO` (E)
- `CORREO` (F), `ESTADO CREDENCIAL` (G), `REGION` (H), `OBSERVACION` (I)
- `ESTADO ANTERIOR` (J) — usado por `verificarCambiosCredenciales` para detectar cambios
- `EMAIL ESTADO` (K) — registro de último email enviado

**Estados de credencial:** `VIGENTE` / `VENCIDA` / `PENDIENTE` / `REVOCADA` / `ENTREGADO` / `DISPONIBLE` / `SOLICITADO` / `NO VIGENTE` / `DATOS INCORRECTOS` / `REIMPRIMIR`

**Hoja `HISTORIAL_CREDENCIALES`** — Log de cambios de estado (FECHA, RUT, NOMBRE, CORREO, ESTADO ANTERIOR, ESTADO NUEVO, EMAIL ENVIADO)

---

### 🗂️ GAMIFICACIÓN — `BD_GAMIFICACION`

Hojas: `BD_GAMIFICACION`, `BANCO_PREGUNTAS`

| # | Col | Campo | Descripción |
|---|---|---|---|
| 0 | A | `RUT` | RUT limpio del socio |
| 1 | B | `NOMBRE` | Nombre del socio |
| 2 | C | `XP_TOTAL` | Puntos de experiencia acumulados |
| 3 | D | `GRADO` | Grado actual (Aspirante → Dirigente) |
| 4 | E | `LOGROS` | Array JSON de logros desbloqueados |
| 5 | F | `QUIZ_ULTIMO_DIA` | Fecha del último quiz completado |
| 6 | G | `RACHA_ACTUAL` | Días consecutivos completando el quiz |
| 7 | H | `ULTIMA_ACTIVIDAD` | Última fecha de actividad |
| 8 | I | `QUIZ_ULTIMO_DIA` | Alias interno de control de racha |
| 9 | J | `RACHA_MAX` | Racha máxima histórica |
| 10 | K | `ESTADO` | `ACTIVO` / `DESVINCULADO` |
| 11 | L | `QUIZZES_COMPLETADOS` | Contador total de quizzes completados |

---

### 🗂️ DESVINCULADOS_CENTRALIZADOS _(no integrado al código aún)_

Hojas: `Hoja 1`, `REACTIVADOS`

**Hoja `Hoja 1`** — Registro centralizado de socios desvinculados (RUT, VALIDACION DE RUT, NOMBRE, ESTADO, DESCRIPCION)

**Hoja `REACTIVADOS`** — Socios reactivados tras desvinculación (RUT, VALIDACION DE RUT, NOMBRE, ESTADO, DESCRIPCION, OBSERVACION)

---

## 📁 Carpetas Google Drive

| Módulo | Clave CONFIG | Carpeta ID |
|---|---|---|
| Justificaciones | `JUSTIFICACIONES` | `1UD9hQz1FuacSb3QYrahRl7IfvlpKn8v6` |
| Apelaciones — Comprobantes | `APELACIONES_COMPROBANTES` | `15BmK5pf5Txrxdzdrny23S5q35NDxLy4P` |
| Apelaciones — Liquidaciones | `APELACIONES_LIQUIDACIONES` | `1dR7fM6TW99tunNaMZliyvXc-L23nHKVY` |
| Apelaciones — Devoluciones | `APELACIONES_DEVOLUCIONES` | `1LGLKA3fiCJXf2ouIqlxq3jk_ZSxI3IyM` |
| Permisos Médicos | `PERMISOS_MEDICOS` | `1nCYxD5sJLszBBA6s2DquGW8vlKGZp4ty` |
| Vestuario — Documentos | `VESTUARIO_DOCS` | `1A4PVsIn8ndNMXdqnO9GZCovjtNdfr0BI` |

---

## 🔐 Seguridad y Control de Acceso

- Validación de rol en todas las funciones sensibles mediante `verificarRolUsuario(rut, rolesPermitidos)`.
- Socios desvinculados (`ESTADO = DESVINCULADO`) solo pueden acceder a **Mis Datos** (excepto ADMIN).
- Perfil `TESTING`: acceso exclusivo a **Mis Datos** y **SLIM Quest**.
- El acceso QR requiere vinculación previa del dispositivo vía `QR_Access.html`.
- Los archivos subidos a Drive se configuran como **privados** con permisos silenciosos solo para los involucrados.
- Login: RUT como usuario + ID Credencial como contraseña (validación Módulo 11).

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

Los módulos se cargan en el dashboard en este orden:

1. **👤 Mis Datos** — Datos personales, banco, credencial
2. **📜 Contrato Colectivo** — Acordeón de 7 capítulos
3. **📝 Justificaciones** — Justificaciones de inasistencia a asambleas
4. **⚖️ Apelaciones** — Apelaciones de multas por inasistencia
5. **💰 Préstamos** — Solicitud de préstamos sindicales
6. **🏥 Permiso Médico** — Permisos médicos laborales
7. **🧮 Calculadora HE** — Calculadora de horas extraordinarias
8. **📋 Registro Asistencia** — Registro QR y presencial/virtual
9. **🎮 SLIM Quest** — Gamificación sindical

### Módulos adicionales (sin switch en dashboard):
- **⚙️ Panel Admin** — Solo ADMIN: switches, triggers, informes
- **🪪 Consulta ID / Credencial** — ADMIN/DIRIGENTE: estado de credencial por RUT
- **👥 Gestión de Socios** — DIRIGENTE/ADMIN: trámites a nombre de socios
- **🔗 Dirigentes** — Links a secciones restringidas de Google Sites

---

### Detalle de módulos principales

#### 1. 👤 Mis Datos
- Edición de correo, teléfono, banco (wizard 3 pasos), tipo y número de cuenta
- Badge de estado de credencial (VIGENTE / VENCIDA / PENDIENTE / REVOCADA)
- Validación RUT en tiempo real (Módulo 11)
- CacheService TTL 10 min para consultas frecuentes

#### 2. 📜 Contrato Colectivo
- Visualización completa (Sindicato SLIM N°3 — ISS, vigencia hasta febrero 2026 prorrogado)
- Acordeón de 7 capítulos con scroll automático
- Cards de la directiva (Presidente en ancho completo)
- Switch de habilitación admin

#### 3. 📝 Justificaciones
- Tipos: Fuerza Mayor, Licencia Médica, y otros
- Adjunto de documento obligatorio para ciertos tipos (JPG, PNG, PDF — máx. 15 MB)
- Restricción por evento activo (`Fecha_Evento` → código `YYYY_MM_DD`)
- Notificación al representante legal con retry cada 30 min
- Switch + badge en dashboard

#### 4. ⚖️ Apelaciones
- Adjunto de comprobante y liquidación de sueldo
- Estados: Solicitado → En Revisión → Aceptado / Aceptado-Obs / Rechazado / Pagado
- Emails de estado con colores específicos por estado
- Retry de permisos Drive con contador `REINTENTO_N` (máx. 5 intentos)
- Columna `LOG_PERMISOS` (col S, índice 18) para auditoría
- Switch + badge en dashboard

#### 5. 💰 Préstamos
- Dos opciones (A/B) por tipo de préstamo
- Montos y cuotas según antigüedad y CC 2026
- Confirmación en modal HTML nativo (sin `Swal.fire().then()`)
- Validación de préstamos por directiva en hoja `Validación-Prestamos`
- Switch + badge en dashboard

#### 6. 🏥 Permiso Médico
- Adjunto opcional al momento de la solicitud
- Tipo especial: `1 día completo` (Cláusula Trigésimo Segunda)
- 3 eventos de notificación según estado del documento
- Retry para `NOTIFICADO_REP_LEGAL` y `NOTIFICADO_SOCIO`
- Switch + badge en dashboard

#### 7. 🧮 Calculadora HE
- Fórmula DT: `FHE = 1.5 / ((NHC/NDC) × 30)`
- Turnos 5×2 → factor `0.00795455`; Turnos 4×4 y 7×7 → factor `0.00833333`
- Perfiles: Limpieza y Guardias de Seguridad
- Nota de Asignación Familiar en sección guardias
- Switch de habilitación admin

#### 8. 📋 Registro Asistencia
- Panel virtual + flujo QR con puntos de control
- Patrón nombre punto: `TIPO_DD-MM-YYYY_CODIGO-REGION`
- Prevención de duplicados: localStorage + validación backend
- Trigger de notificación por correo delegado a las 20:00
- `LockService`: búsqueda de usuario fuera del lock
- Switch + badge en dashboard

#### 9. 🎮 SLIM Quest
- 6 grados de progresión:

| Grado | XP Requerido | Icono |
|---|---|---|
| Aspirante | 0 – 1.500 | 🌱 |
| Aprendiz | 1.501 – 4.500 | ⚙️ |
| Trabajador | 4.501 – 10.000 | 🔩 |
| Defensor | 10.001 – 18.000 | 🛡️ |
| Negociador | 18.001 – 30.000 | ⚖️ |
| Dirigente | 30.001+ | 🏆 |

- Quiz diario (3 niveles: BASICO / INTERMEDIO / AVANZADO)
- Sistema de rachas con bonos en hitos: 3, 7, 14, 21, 30, 60, 100 días consecutivos
- Nivel Secreto `DIRIGENTE`: preguntas exclusivas
- Leaderboard top 10 (formato nombre: `JUAN LOPEZ G.`)
- Logros en JSON por socio
- Sincronización automática desde `BD_SLIMAPP` — trigger diario 1:00 AM
- Motor de sonidos `SLIMSound` (Web Audio API, 17 tipos, mute persistido en localStorage)
- Switch + badge en dashboard (bypass para ADMIN)

---

## ⏰ Triggers Automáticos

| Función | Frecuencia | Descripción |
|---|---|---|
| `verificarCambiosCredenciales` | Diario 8:00 AM | Detecta cambios en `BD_CREDENCIALES` y notifica socios |
| `sincronizarSociosGamificacion` | Diario 1:00 AM | Sincroniza socios nuevos/desvinculados a BD_GAMIFICACION |
| `reintentarNotificacionRepLegal` | Cada 30 min | Retry de notificaciones pendientes a Rep. Legal |
| `reintentarNotificacionSocio` | Cada 30 min | Retry de notificaciones pendientes al socio |
| `procesarValidacionPrestamos` | Periódico | Lee `Validación-Prestamos` y notifica resultados |
| `enviarNotificacionAsistencia` | Diario 20:00 | Envía correos de confirmación de asistencia pendientes |

---

## 🔧 Switches de Módulos

Todos los switches se almacenan en `PropertiesService` con la convención de nombre `{modulo}_habilitado`.

| Módulo | Clave PropertiesService |
|---|---|
| Justificaciones | `justificaciones_habilitado` |
| Apelaciones | `apelaciones_habilitado` |
| Préstamos | `prestamos_habilitado` |
| Permisos Médicos | `permisos_medicos_habilitado` |
| Contrato Colectivo | `contrato_colectivo_habilitado` |
| Calculadora HE | `calculadora_he_habilitado` |
| Registro Asistencia | `registro_asistencia_habilitado` |
| SLIM Quest | `slimquest_habilitado` |

El estado de todos los switches se carga en una sola llamada mediante `obtenerEstadosSwitchDashboard()`.

---

## 🎨 Stack Frontend

| Tecnología | Uso |
|---|---|
| **TailwindCSS** (CDN) | Framework CSS para todo el diseño |
| **Material Icons Round** (CDN) | Iconografía del sistema |
| **SweetAlert** | Modales de alerta y confirmación |
| **QuickChart API** | Generación de QR (`https://quickchart.io/qr?size=300&text={url}`) |
| **SLIMSound Engine** | Motor de sonidos (Web Audio API) integrado en el frontend |
| **`google.script.run`** | Comunicación asíncrona frontend → backend |

---

## 🔑 Referencia de Funciones Backend (Code.gs)

### Autenticación y Usuario

| Función | Descripción |
|---|---|
| `doGet(e)` | Router principal de la Web App |
| `validarUsuario(rut, password)` | Login con RUT + ID Credencial |
| `obtenerDatosUsuario(rut)` | Datos completos del socio desde `BD_SLIMAPP` |
| `actualizarDatosContacto(rut, correo, telefono)` | Actualiza correo y teléfono |
| `actualizarDatosBancarios(rut, banco, tipoCuenta, numeroCuenta)` | Actualiza datos bancarios (atómico) |

### Justificaciones

| Función | Descripción |
|---|---|
| `obtenerConfigJustificaciones()` | Lee switch, fecha límite y código de asamblea |
| `enviarJustificacion(datos)` | Registra nueva justificación en `BD_JUSTIFICACIONES` |
| `obtenerJustificacionesSocio(rut)` | Historial de justificaciones del socio |
| `gestionarJustificacion(id, estado, obs, rutGestor)` | Acepta o rechaza justificación (DIRIGENTE/ADMIN) |
| `reintentarNotificacionRepLegal()` | Trigger: reintentos de notificación pendientes |

### Apelaciones

| Función | Descripción |
|---|---|
| `enviarApelacion(datos)` | Registra nueva apelación con archivos adjuntos |
| `obtenerApelacionesSocio(rut)` | Historial de apelaciones del socio |
| `gestionarApelacion(id, estado, obs, rutGestor)` | Cambia estado y notifica al socio |
| `adjuntarComprobanteDevolucion(id, archivoData)` | Adjunta comprobante de devolución de multa |

### Préstamos

| Función | Descripción |
|---|---|
| `obtenerOpcionesPrestamoSocio(rut)` | Calcula montos disponibles según antigüedad |
| `solicitarPrestamo(datos)` | Registra nueva solicitud de préstamo |
| `obtenerPrestamosSocio(rut)` | Historial de préstamos del socio |
| `procesarValidacionPrestamos()` | Trigger: procesa validaciones de la directiva |

### Permisos Médicos

| Función | Descripción |
|---|---|
| `solicitarPermisoMedico(datos)` | Registra nueva solicitud de permiso médico |
| `adjuntarDocumentoPermiso(id, archivoData)` | Adjunta documento posterior a la solicitud |
| `obtenerPermisosMedicosSocio(rut)` | Historial de permisos del socio |
| `gestionarPermisoMedico(id, estado, obs, rutGestor)` | Gestiona el permiso (DIRIGENTE/ADMIN) |

### Asistencia

| Función | Descripción |
|---|---|
| `registrarAsistenciaQR(rut, codigoTemporal)` | Registro de asistencia vía QR |
| `registrarAsistenciaManual(rut, tipo, asamblea, rutGestor)` | Registro presencial/virtual manual |
| `obtenerAsistenciasSocio(rut)` | Historial de asistencias del socio |
| `obtenerPuntosControl()` | Lista de puntos de control activos |
| `crearPuntoControl(nombre, horaApertura, horaCierre)` | Crea nuevo punto de control QR |

### Credenciales

| Función | Descripción |
|---|---|
| `obtenerEstadoCredencialPorRut(rut)` | Estado de credencial del socio |
| `verificarCambiosCredenciales()` | Trigger: detecta cambios y notifica |
| `consultarIDSocio(rut)` | Consulta completa de credencial (DIRIGENTE/ADMIN) |

### SLIM Quest

| Función | Descripción |
|---|---|
| `obtenerDatosGamificacion(rut)` | Progreso completo del socio |
| `desbloquearLogro(rut, nombre, descripcion, icono)` | Desbloquea logro al socio |
| `completarQuiz(rut, xpGanado, correctas)` | Procesa resultado del quiz (racha, bonos, nivel) |
| `obtenerPreguntasQuiz(rut, cantidad)` | Preguntas ponderadas según grado |
| `obtenerPreguntasSecreto(rut, cantidad)` | Preguntas exclusivas para grado Dirigente |
| `getLeaderboard(rut)` | Top 10 socios por XP |
| `sincronizarSociosGamificacion()` | Sincroniza socios desde `BD_SLIMAPP` |

### Auxiliares y Configuración

| Función | Descripción |
|---|---|
| `cleanRut(rut)` | Limpia RUT (elimina puntos, guión, mayúsculas) |
| `esCorreoValido(correo)` | Valida formato de correo |
| `enviarCorreoEstilizado(...)` | Envía correo HTML con diseño corporativo |
| `subirArchivoConPermisos(...)` | Sube archivo a Drive (3 intentos con retry) |
| `validarCorreosParaPermisos(...)` | Prepara lista de correos para acceso Drive |
| `getSheet(ssKey, sheetKey)` | Obtiene hoja por clave de CONFIG |
| `getSpreadsheet(ssKey)` | Obtiene spreadsheet por clave de CONFIG |
| `obtenerEstadosSwitchDashboard()` | Estado de todos los módulos en una sola llamada |
| `configurarTriggers()` | Configura todos los triggers automáticos |
| `generarLinksRegistroYQR()` | Genera links y QR de registro para todos los socios |
| `verificarRolUsuario(rut, rolesPermitidos)` | Valida permisos de rol en funciones sensibles |

---

## 🔑 Notas Técnicas Importantes

### Manejo de fechas
- Siempre anclar strings de fecha al mediodía (`T12:00:00`) para evitar desfases UTC → Santiago (UTC-3).
- Usar `getValues()` para columnas de fecha (evita inversión día/mes en formato US de `getDisplayValues()`).
- Usar `setNumberFormat('@STRING@')` para prevenir que Sheets convierta texto a Date automáticamente.
- Strings como `"2026-02-25"` se interpretan como medianoche UTC → usar `'T12:00:00'` al construir Date.

### Concurrencia y performance
- Usar `getUserLock()` para operaciones iniciadas por el usuario (evita conflicto con triggers que usan `getScriptLock()`).
- `getScriptLock()` en triggers puede retener el lock hasta 60 s, bloqueando envíos simultáneos.
- `CacheService` TTL 10 min para búsquedas de usuario (`user_RUT`).
- Búsqueda de usuario fuera del lock crítico en operaciones de asistencia.
- `SpreadsheetApp.flush()` antes de `getDataRange().getValues()` en funciones de validación para evitar race conditions.

### Restricciones de Google Apps Script
- Evitar backticks `` ` `` (template literals) — usar concatenación con `+` para prevenir errores de string no cerrado.
- `Swal.fire().then()` es poco confiable en iframes sandboxed de GAS — usar modales HTML nativos con handlers separados.
- `innerHTML +=` en loops destruye referencias DOM — pasar `this` en handlers `onclick`.
- `appendRow` es atómico y siempre escribe tras la última fila con datos.
- `ScriptApp.getService().getUrl()` no retorna el formato `/a/~/macros/s/` de forma confiable.

### Archivos y Drive
- Eliminar `capture="environment"` de inputs de archivo (fuerza cámara en móvil, impide seleccionar PDF).
- `Drive.Permissions.insert` falla intermitentemente — usar 3 intentos con `Utilities.sleep()`.
- Fórmulas `=IMAGE()` en Sheets: usar `getFormulas()` en lugar de `getValues()`.

### Columnas y estructura
- Los índices de columna se usan directamente en el código (no los nombres de cabecera).
- Renombrar cabeceras no rompe el código, pero insertar/eliminar columnas sí.
- Sentinel `SIN_CORREO`: registros sin correo válido reciben este valor en columnas de notificación.

### Patrones de navegación
- `navAction` usa cadenas `if/else if` (no `switch/case`) — importante al agregar nuevos módulos.
- Visibilidad de switches debe establecerse dentro del branch `navAction`, no en login/`setupDashboard`.

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

**Correo Representante Legal:** `juancarlos.pacheco@cl.issworld.com`

---

## 👨‍💻 Desarrollo

**Organización:** Sindicato SLIM N°3
**Desarrollador:** Alejandro Peñailillo G. — DUOC UC, Técnico Analista Programador
**Repositorio:** `Daroce12/SLIMAPP`
**Rama principal:** `main`
**Plataforma:** Google Apps Script + Google Workspace

---

*Proyecto privado — Todos los derechos reservados*
