/**
 * Servir HTML
 */
// REEMPLAZAR CON ESTO:
function doGet(e) {
  // CASO 1: Checkin de punto de control (QR de asamblea)
  // Detecta parámetro 'control' o acción 'checkin'
  if (e.parameter.control || e.parameter.action === 'checkin') {
    const template = HtmlService.createTemplateFromFile('QR_Asistencia');
    template.params = e.parameter;
    return template.evaluate()
        .setTitle('Registro Asistencia - Sindicato SLIM n°3')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
  }
  
  // CASO 2: Vinculación QR personal (action=register, rut=...)
  if (e.parameter.action || e.parameter.rut || e.parameter.asamblea) {
    const template = HtmlService.createTemplateFromFile('QR_Access');
    template.data = e.parameter;
    return template.evaluate()
        .setTitle('Control QR - Sindicato SLIM n°3')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
  }
  
  // CASO 3: Aplicación principal
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Sindicato SLIM n°3 - App Socios')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

// ==========================================
// CONFIGURACIÓN GLOBAL - IDs DE SPREADSHEETS Y CARPETAS
// ==========================================
const CONFIG = {
  SPREADSHEETS: {
    USUARIOS: "1m7KLd3b3BzKOAI10I5E32MVf_L34XWAGFonhTg37TVM",
    JUSTIFICACIONES: "1Hwbly__MXjl9uwJb-spXdah-R3v9SAMOCFHem92uOUg",
    APELACIONES: "11nrvVsf84THWQ7j6NfAr_unyIcBV7aykxACS8R27PwE",
    PRESTAMOS: "1h-_sJD4rOCuMjlfSouP7a6gfoodHyzI4MOBRUyOW5XU",
    PERMISOS_MEDICOS: "1VYfm7cOgL3mVfVoI8DubIm8WG2srzQw9a6DtIEs3UMM",
    CREDENCIALES: "1HVyPxdYKuvIybeOCAPwAJaVHwlxEuOik4YW0XOXBE5o",
    ASISTENCIA: "1SRQ8Mlc6bBdb0mitAfn4I-EUAS4BOrZRbqS9YAmg3Sk"
  },
  HOJAS: {
    USUARIOS: "BD_SLIMAPP",
    JUSTIFICACIONES: "BD_JUSTIFICACIONES",
    CONFIG_JUSTIFICACIONES: "CONFIG_JUSTIFICACIONES",
    APELACIONES: "BD_APELACIONES",
    PRESTAMOS: "BD_PRESTAMOS",
    VALIDACION_PRESTAMOS: "Validación-Prestamos",
    PERMISOS_MEDICOS: "BD_Permisos medicos",
    CREDENCIALES: "IMPRESION",
    HISTORIAL_CREDENCIALES: "HISTORIAL_CREDENCIALES",
    ASISTENCIA: "BD_ASISTENCIA",
    PUNTOS_CONTROL: "PUNTOS_CONTROL"
  },
  CARPETAS: {
    JUSTIFICACIONES: "1UD9hQz1FuacSb3QYrahRl7IfvlpKn8v6",
    APELACIONES_COMPROBANTES: "15BmK5pf5Txrxdzdrny23S5q35NDxLy4P",
    APELACIONES_LIQUIDACIONES: "1dR7fM6TW99tunNaMZliyvXc-L23nHKVY",
    APELACIONES_DEVOLUCIONES: "1LGLKA3fiCJXf2ouIqlxq3jk_ZSxI3IyM",
    PERMISOS_MEDICOS: "1nCYxD5sJLszBBA6s2DquGW8vlKGZp4ty"
  },
  CORREOS: {
    REPRESENTANTE_LEGAL: "juancarlos.pacheco@cl.issworld.com"
  },
  COLUMNAS: {
    USUARIOS: {
      RUT: 0,
      RUT_VALIDADO: 1,
      FECHA_INGRESO: 2,
      NOMBRE: 3,
      CARGO: 4,
      CORREO: 5,
      SITE: 6,
      REGION: 7,
      SEXO: 8,
      ESTADO: 9,
      DETALLE_DESVINCULACION: 10,
      ID_CREDENCIAL: 11,
      CORREO_REGISTRADO: 12,
      CONTACTO: 13,
      ROL: 14,
      LINK_REGISTRO: 15,
      QR_REGISTRO: 16,
      BANCO: 17,
      TIPO_CUENTA: 18,
      NUMERO_CUENTA: 19,
      ESTADO_NEG_COLECT: 20
    },
    JUSTIFICACIONES: {
      ID: 0,
      FECHA: 1,
      RUT: 2,
      NOMBRE: 3,
      REGION: 4,
      MOTIVO: 5,
      ARGUMENTO: 6,
      RESPALDO: 7,
      ESTADO: 8,
      OBSERVACION: 9,
      NOTIFICACION: 10,
      ASAMBLEA: 11,
      GESTION: 12,
      DIRIGENTE: 13,
      CORREO_DIRIGENTE: 14
    },
    APELACIONES: {
      ID: 0,
      FECHA_SOLICITUD: 1,
      RUT: 2,
      NOMBRE: 3,
      CORREO: 4,
      MES_APELACION: 5,
      TIPO_MOTIVO: 6,
      DETALLE_MOTIVO: 7,
      URL_COMPROBANTE: 8,
      URL_LIQUIDACION: 9,
      ESTADO: 10,
      OBSERVACION: 11,
      NOTIFICADO: 12,
      GESTION: 13,
      NOMBRE_DIRIGENTE: 14,
      CORREO_DIRIGENTE: 15,
      URL_COMPROBANTE_DEVOLUCION: 16
    },
    PRESTAMOS: {
      ID: 0,
      FECHA: 1,
      RUT: 2,
      NOMBRE: 3,
      CORREO: 4,
      TIPO: 5,
      MONTO: 6,
      CUOTAS: 7,
      MEDIO_PAGO: 8,
      ESTADO: 9,
      FECHA_TERMINO: 10,
      GESTION: 11,
      NOMBRE_DIRIGENTE: 12,
      CORREO_DIRIGENTE: 13,
      INFORME: 14,
      OBSERVACION: 15
    },
    PERMISOS_MEDICOS: {
      ID: 0,
      FECHA_SOLICITUD: 1,
      RUT: 2,
      NOMBRE: 3,
      CORREO: 4,
      TIPO_PERMISO: 5,
      FECHA_INICIO: 6,
      MOTIVO_DETALLE: 7,
      URL_DOCUMENTO: 8,
      ESTADO: 9,
      FECHA_SUBIDA: 10,
      NOTIFICADO_REP_LEGAL: 11,
      GESTION: 12,
      NOMBRE_DIRIGENTE: 13,
      CORREO_DIRIGENTE: 14
    }
  }
};

// ==========================================
// FUNCIÓN HELPER: FORMATEAR FECHA A dd/mm/yyyy - hh:mm
// ==========================================

/**
 * Formatea una fecha a formato dd/mm/yyyy - hh:mm
 * @param {Date|string} fecha - Fecha a formatear
 * @returns {string} Fecha formateada o string vacío si es inválida
 */
function formatearFechaConHora(fecha) {
  try {
    if (!fecha) return "";
    
    let fechaObj;
    
    // Si es string, convertir a Date
    if (typeof fecha === 'string') {
      fechaObj = new Date(fecha);
    } else if (fecha instanceof Date) {
      fechaObj = fecha;
    } else {
      return fecha.toString(); // Devolver como está si no se puede procesar
    }
    
    // Validar que la fecha es válida
    if (isNaN(fechaObj.getTime())) {
      return fecha.toString();
    }
    
    // Extraer componentes
    const dia = String(fechaObj.getDate()).padStart(2, '0');
    const mes = String(fechaObj.getMonth() + 1).padStart(2, '0');
    const anio = fechaObj.getFullYear();
    const hora = String(fechaObj.getHours()).padStart(2, '0');
    const minutos = String(fechaObj.getMinutes()).padStart(2, '0');
    
    // Formato: dd/mm/yyyy - hh:mm
    return `${dia}/${mes}/${anio} - ${hora}:${minutos}`;
    
  } catch (e) {
    Logger.log('Error formateando fecha: ' + e.toString());
    return fecha ? fecha.toString() : "";
  }
}

/**
 * Formatea una fecha a formato dd/mm/yyyy (sin hora)
 * @param {Date|string} fecha - Fecha a formatear
 * @returns {string} Fecha formateada o string vacío si es inválida
 */
function formatearFechaSinHora(fecha) {
  try {
    if (!fecha) return "";
    
    let fechaObj;
    
    // Si es string, convertir a Date
    if (typeof fecha === 'string') {
      fechaObj = new Date(fecha);
    } else if (fecha instanceof Date) {
      fechaObj = fecha;
    } else {
      return fecha.toString();
    }
    
    // Validar que la fecha es válida
    if (isNaN(fechaObj.getTime())) {
      return fecha.toString();
    }
    
    // Extraer componentes
    const dia = String(fechaObj.getDate()).padStart(2, '0');
    const mes = String(fechaObj.getMonth() + 1).padStart(2, '0');
    const anio = fechaObj.getFullYear();
    
    // Formato: dd/mm/yyyy
    return `${dia}/${mes}/${anio}`;
    
  } catch (e) {
    Logger.log('Error formateando fecha: ' + e.toString());
    return fecha ? fecha.toString() : "";
  }
}

/**
 * Función helper para obtener un spreadsheet específico
 * @param {string} spreadsheetKey - Clave del spreadsheet en CONFIG.SPREADSHEETS
 * @returns {Spreadsheet} - Objeto Spreadsheet
 */
function getSpreadsheet(spreadsheetKey) {
  const spreadsheetId = CONFIG.SPREADSHEETS[spreadsheetKey];
  if (!spreadsheetId) {
    throw new Error(`Spreadsheet key "${spreadsheetKey}" no encontrado en CONFIG`);
  }
  return SpreadsheetApp.openById(spreadsheetId);
}

// ==========================================
// SISTEMA CENTRALIZADO DE PERMISOS DE ARCHIVOS
// Agregar DESPUÉS de la sección CONFIG
// ==========================================

/**
 * Valida si un correo electrónico es válido para otorgar permisos
 * @param {string} correo - Correo a validar
 * @returns {boolean} true si es válido
 */
function esCorreoValido(correo) {
  if (!correo || typeof correo !== 'string') return false;
  const correoLimpio = correo.trim().toLowerCase();
  const regexCorreo = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return regexCorreo.test(correoLimpio);
}

/**
 * Valida los correos de los usuarios involucrados antes de procesar archivos
 * @param {Object} beneficiario - Objeto con datos del beneficiario {rut, nombre, correo}
 * @param {Object} gestor - Objeto con datos del gestor (puede ser null si es el mismo)
 * @param {boolean} esGestionDirigente - true si la gestión es realizada por un dirigente/admin
 * @returns {Object} {valido: boolean, alertas: [], correosParaPermisos: []}
 */
function validarCorreosParaPermisos(beneficiario, gestor, esGestionDirigente) {
  const resultado = {
    valido: true,
    alertas: [],
    correosParaPermisos: [],
    alertaBeneficiario: false,
    alertaGestor: false
  };
  
  // Validar correo del beneficiario
  const correoBeneficiarioValido = esCorreoValido(beneficiario.correo);
  
  if (correoBeneficiarioValido) {
    resultado.correosParaPermisos.push({
      correo: beneficiario.correo.trim().toLowerCase(),
      tipo: 'beneficiario',
      nombre: beneficiario.nombre
    });
  } else {
    resultado.alertaBeneficiario = true;
    if (esGestionDirigente) {
      resultado.alertas.push({
        tipo: 'warning',
        mensaje: `El socio ${beneficiario.nombre} no tiene un correo electrónico válido registrado. No podrá acceder al archivo adjunto. Infórmele que debe actualizar sus datos en "Mis Datos".`
      });
    } else {
      resultado.alertas.push({
        tipo: 'warning',
        mensaje: `No tienes un correo electrónico válido registrado. No podrás acceder al archivo adjunto desde tu correo. Por favor, actualiza tus datos en el módulo "Mis Datos".`
      });
    }
  }
  
  // Validar correo del gestor (solo si es gestión de dirigente y es diferente al beneficiario)
  if (esGestionDirigente && gestor) {
    const correoGestorValido = esCorreoValido(gestor.correo);
    
    if (correoGestorValido) {
      const yaExiste = resultado.correosParaPermisos.some(
        c => c.correo === gestor.correo.trim().toLowerCase()
      );
      
      if (!yaExiste) {
        resultado.correosParaPermisos.push({
          correo: gestor.correo.trim().toLowerCase(),
          tipo: 'gestor',
          nombre: gestor.nombre
        });
      }
    } else {
      resultado.alertaGestor = true;
      resultado.alertas.push({
        tipo: 'info',
        mensaje: `Tu correo electrónico no está registrado correctamente. El archivo se procesará, pero no recibirás acceso directo. Actualiza tus datos en "Mis Datos".`
      });
    }
  }
  
  return resultado;
}

/**
 * Sube un archivo a Google Drive y otorga permisos de lectura
 * @param {Object} archivoData - {base64, mimeType, fileName}
 * @param {string} carpetaId - ID de la carpeta de destino
 * @param {string} nombreArchivo - Nombre personalizado para el archivo
 * @param {Array} correosParaPermisos - [{correo, tipo, nombre}, ...]
 * @param {Array} correosAdicionales - Correos adicionales (ej: representante legal)
 * @returns {Object} {success, url, permisosOtorgados: [], permisosError: []}
 */
function subirArchivoConPermisos(archivoData, carpetaId, nombreArchivo, correosParaPermisos, correosAdicionales) {
  correosAdicionales = correosAdicionales || [];
  
  const resultado = {
    success: false,
    url: '',
    permisosOtorgados: [],
    permisosError: [],
    mensajeError: ''
  };
  
  try {
    // Validar tamaño del archivo
    const sizeInBytes = (archivoData.base64.length * 3) / 4;
    if (sizeInBytes > 15 * 1024 * 1024) {
      resultado.mensajeError = "El archivo es demasiado grande (máximo 15MB).";
      return resultado;
    }
    
    // Obtener carpeta
    const folder = DriveApp.getFolderById(carpetaId);
    
    // Crear blob
    const blob = Utilities.newBlob(
      Utilities.base64Decode(archivoData.base64),
      archivoData.mimeType,
      archivoData.fileName
    );
    
    // Obtener extensión original
    let extension = "";
    const nameParts = archivoData.fileName.split('.');
    if (nameParts.length > 1) {
      extension = "." + nameParts.pop();
    }
    
    // Establecer nombre del archivo
    blob.setName(nombreArchivo + extension);
    
    // Crear archivo
    const file = folder.createFile(blob);
    
    // Esperar a que Drive procese el archivo
    Utilities.sleep(1500);
    
    // Configurar como privado primero
    file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
    
    // Esperar un poco más
    Utilities.sleep(1000);
    
    // Combinar correos para permisos
    const todosLosCorreos = [...correosParaPermisos];
    
    // Agregar correos adicionales
    if (correosAdicionales && correosAdicionales.length > 0) {
      correosAdicionales.forEach(function(correo) {
        if (esCorreoValido(correo)) {
          const yaExiste = todosLosCorreos.some(function(c) {
            return c.correo === correo.trim().toLowerCase();
          });
          if (!yaExiste) {
            todosLosCorreos.push({
              correo: correo.trim().toLowerCase(),
              tipo: 'adicional',
              nombre: 'Usuario adicional'
            });
          }
        }
      });
    }
    
    // Otorgar permisos SILENCIOSOS a cada correo usando Drive API Avanzada
    var fileId = file.getId(); // Obtenemos el ID para usar la API avanzada

    todosLosCorreos.forEach(function(item) {
      
      // ── INTENTO 1: API Avanzada (sin email de notificación) ──
      try {
        var recursoPermiso = {
          'role': 'reader',
          'type': 'user',
          'value': item.correo
        };
        Drive.Permissions.insert(recursoPermiso, fileId, {
          sendNotificationEmails: false
        });
        resultado.permisosOtorgados.push({
          correo: item.correo,
          tipo: item.tipo,
          nombre: item.nombre
        });
        Logger.log("✅ Permiso silencioso otorgado a " + item.tipo + ": " + item.correo);

      } catch (permError) {
        Logger.log("⚠️ Fallo API Avanzada para " + item.correo + " - " + permError + " - Intentando método alternativo...");
        
        // Esperar antes del primer fallback para reducir conflictos de cuota
        Utilities.sleep(1000);
        
        // ── INTENTO 2: Método tradicional (addViewer) ──
        try {
          file.addViewer(item.correo);
          resultado.permisosOtorgados.push({
            correo: item.correo,
            tipo: item.tipo,
            nombre: item.nombre
          });
          Logger.log("✅ Permiso otorgado via addViewer a " + item.tipo + ": " + item.correo);

        } catch (fallbackError) {
          Logger.log("⚠️ Fallo addViewer para " + item.correo + " - " + fallbackError + " - Reintentando en 2s...");
          
          // Esperar más tiempo antes del último reintento
          Utilities.sleep(2000);
          
          // ── INTENTO 3 (ÚLTIMO): Reintento final con addViewer ──
          try {
            file.addViewer(item.correo);
            resultado.permisosOtorgados.push({
              correo: item.correo,
              tipo: item.tipo,
              nombre: item.nombre
            });
            Logger.log("✅ Permiso otorgado en reintento final para: " + item.correo);

          } catch (finalError) {
            resultado.permisosError.push({
              correo: item.correo,
              tipo: item.tipo,
              nombre: item.nombre,
              error: finalError.toString()
            });
            Logger.log("❌ Error fatal al otorgar permiso a " + item.tipo + " (" + item.correo + "): " + finalError);
          }
        }
      }
    });
    
    resultado.success = true;
    resultado.url = file.getUrl();
    
    Logger.log("📊 Archivo subido: " + nombreArchivo);
    Logger.log("   - URL: " + resultado.url);
    Logger.log("   - Permisos exitosos: " + resultado.permisosOtorgados.length);
    Logger.log("   - Permisos fallidos: " + resultado.permisosError.length);
    
    return resultado;
    
  } catch (error) {
    Logger.log("❌ Error al subir archivo: " + error.toString());
    resultado.mensajeError = "Error al subir el archivo: " + error.toString();
    return resultado;
  }
}

/**
 * Genera el mensaje de alerta para mostrar al usuario sobre permisos
 */
function generarAlertaPermisos(validacionCorreos, resultadoSubida) {
  const alerta = {
    mostrarAlerta: false,
    tipoAlerta: 'info',
    mensajeAlerta: '',
    detalles: []
  };
  
  // Agregar alertas de validación de correos
  if (validacionCorreos.alertas && validacionCorreos.alertas.length > 0) {
    alerta.mostrarAlerta = true;
    validacionCorreos.alertas.forEach(function(a) {
      alerta.detalles.push(a.mensaje);
    });
    
    if (validacionCorreos.alertaBeneficiario) {
      alerta.tipoAlerta = 'warning';
    }
  }
  
  // Agregar errores de permisos si los hay
  if (resultadoSubida && resultadoSubida.permisosError && resultadoSubida.permisosError.length > 0) {
    alerta.mostrarAlerta = true;
    alerta.tipoAlerta = 'warning';
    resultadoSubida.permisosError.forEach(function(err) {
      alerta.detalles.push("No se pudo otorgar acceso a " + err.nombre + " (" + err.correo + ")");
    });
  }
  
  // Construir mensaje final
  if (alerta.mostrarAlerta) {
    alerta.mensajeAlerta = alerta.detalles.join('\n\n');
  }
  
  return alerta;
}

/**
 * Obtener datos de usuario por RUT - Función auxiliar centralizada
 */
function obtenerUsuarioPorRut(rutInput) {
  var cache = CacheService.getScriptCache();
  var rutLimpio = cleanRut(rutInput);
  var cacheKey = 'user_' + rutLimpio;
  
  var cached = cache.get(cacheKey);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {
      Logger.log('Error parsing cache: ' + e);
    }
  }
  
  var sheet = getSheet('USUARIOS', 'USUARIOS');
  var COL = CONFIG.COLUMNAS.USUARIOS;
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { encontrado: false };
  
  var data = sheet.getRange(2, 1, lastRow - 1, COL.ESTADO_NEG_COLECT + 1).getDisplayValues();  // ← MODIFICADO
  
  for (var i = 0; i < data.length; i++) {
    if (cleanRut(data[i][COL.RUT]) === rutLimpio) {
      var usuario = {
        encontrado: true,
        rut: data[i][COL.RUT],
        nombre: data[i][COL.NOMBRE],
        correo: data[i][COL.CORREO],
        region: data[i][COL.REGION],
        cargo: data[i][COL.CARGO],
        site: data[i][COL.SITE],
        estado: data[i][COL.ESTADO],
        rol: data[i][COL.ROL],
        contacto: data[i][COL.CONTACTO],
        estadoNegColect: data[i][COL.ESTADO_NEG_COLECT] || "",
        banco: data[i][COL.BANCO] || "",
        tipoCuenta: data[i][COL.TIPO_CUENTA] || "",
        numeroCuenta: data[i][COL.NUMERO_CUENTA] || ""
      };
      
      try {
        cache.put(cacheKey, JSON.stringify(usuario), 600);
      } catch (e) {
        Logger.log('Error guardando en cache: ' + e);
      }
      
      return usuario;
    }
  }
  
  return { encontrado: false };
}

/**
 * Genera el código de asamblea en formato YYYY_MM
 * @param {Date} fecha - Fecha de la solicitud
 * @return {string} Código de asamblea (ejemplo: "2026_01")
 */
function generarCodigoAsamblea(fecha) {
  if (!fecha || !(fecha instanceof Date)) {
    fecha = new Date(); // Si no hay fecha, usa la actual
  }
  
  const year = fecha.getFullYear();
  const month = String(fecha.getMonth() + 1).padStart(2, '0'); // Mes con 2 dígitos
  
  return `${year}_${month}`;
}

/**
 * Genera el código de asamblea desde la fecha del evento configurado
 * Formato: YYYY_MM_DD (ej: "2026_02_27")
 * @param {string|Date} fechaEvento - Fecha del evento en formato "YYYY-MM-DD" o Date
 * @returns {string} Código de asamblea con día (ejemplo: "2026_02_27")
 */
function generarCodigoAsambleaEvento(fechaEvento) {
  try {
    var fecha;
    if (typeof fechaEvento === 'string') {
      var soloFecha = fechaEvento.split('T')[0];
      var partes = soloFecha.split('-');
      // Anclar a mediodía para evitar desfase de zona horaria
      fecha = new Date(parseInt(partes[0]), parseInt(partes[1]) - 1, parseInt(partes[2]), 12, 0, 0);
    } else if (fechaEvento instanceof Date) {
      fecha = fechaEvento;
    } else {
      return generarCodigoAsamblea(new Date());
    }
    
    const year = fecha.getFullYear();
    const month = String(fecha.getMonth() + 1).padStart(2, '0');
    const day = String(fecha.getDate()).padStart(2, '0');
    
    return `${year}_${month}_${day}`;
  } catch(e) {
    Logger.log('Error en generarCodigoAsambleaEvento: ' + e.toString());
    return generarCodigoAsamblea(new Date());
  }
}

/**
 * Función helper mejorada para obtener hoja específica con manejo de errores
 * @param {string} spreadsheetKey - Clave del spreadsheet en CONFIG.SPREADSHEETS
 * @param {string} sheetKey - Clave de la hoja en CONFIG.HOJAS
 * @param {boolean} createIfNotExists - Si true, crea la hoja si no existe (default: false)
 * @returns {Sheet|null} - Objeto Sheet o null si no existe
 */
function getSheet(spreadsheetKey, sheetKey, createIfNotExists = false) {
  try {
    const ss = getSpreadsheet(spreadsheetKey);
    const sheetName = CONFIG.HOJAS[sheetKey];
    
    if (!sheetName) {
      console.error(`❌ Clave de hoja "${sheetKey}" no encontrada en CONFIG.HOJAS`);
      return null;
    }
    
    let sheet = ss.getSheetByName(sheetName);
    
    // Si no existe y se solicita creación automática
    if (!sheet && createIfNotExists) {
      console.warn(`⚠️ Hoja "${sheetName}" no existe. Creándola...`);
      sheet = ss.insertSheet(sheetName);
      console.log(`✅ Hoja "${sheetName}" creada exitosamente`);
    }
    
    if (!sheet) {
      console.error(`❌ Hoja "${sheetName}" no encontrada en spreadsheet ${spreadsheetKey}`);
      return null;
    }
    
    return sheet;
    
  } catch (e) {
    console.error(`❌ Error obteniendo hoja ${sheetKey} de ${spreadsheetKey}: ${e.toString()}`);
    return null;
  }
}

/**
 * Validar Usuario (Login)
 */
function validarUsuario(rutInput, passwordInput) {
  try {
    const sheet = getSheet('USUARIOS', 'USUARIOS');
    const data = sheet.getDataRange().getDisplayValues();
    const rutLimpioInput = cleanRut(rutInput);
    const COL = CONFIG.COLUMNAS.USUARIOS;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      if (cleanRut(row[COL.RUT]) === rutLimpioInput) {      
        const passDb = String(row[COL.ID_CREDENCIAL]);
        const nombreUsuario = row[COL.NOMBRE];
        const rolUsuario = String(row[COL.ROL]).trim().toUpperCase(); // ✅ Normalizar
        const estadoUsuario = String(row[COL.ESTADO]).toUpperCase();
       
        if (String(passDb).toUpperCase() === String(passwordInput).toUpperCase()) {
          const resultado = {
            success: true,
            message: "Login exitoso",
            user: nombreUsuario || "Socio",
            role: rolUsuario || "SOCIO",
            state: estadoUsuario || "ACTIVO",
            estadoNegColect: String(row[COL.ESTADO_NEG_COLECT] || "").trim()
          };
          return resultado;
        } else {
          return { 
            success: false, 
            message: "Contraseña incorrecta",
            errorType: "password"  // ⭐ NUEVO: Identificador del tipo de error
          };
        }
      }
    }
    return { 
      success: false, 
      message: "RUT no encontrado",
      errorType: "rut"  // ⭐ NUEVO: Identificador del tipo de error
    };
  } catch (e) {
    Logger.log('ERROR en validarUsuario: ' + e.toString());
    return { success: false, message: "Error Servidor: " + e.toString() };
  }
}

/**
 * Obtener Datos Completos del Usuario
 */
function obtenerDatosUsuario(rutInput) {
  try {
    const sheet = getSheet('USUARIOS', 'USUARIOS');
    const data = sheet.getDataRange().getDisplayValues();
    const rutLimpioInput = cleanRut(rutInput);
    const COL = CONFIG.COLUMNAS.USUARIOS;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (cleanRut(row[COL.RUT]) === rutLimpioInput) {
        return {
          success: true,
          datos: {
            rut: row[COL.RUT] || "---",          
            nombre: row[COL.NOMBRE] || "Sin Nombre",
            cargo: row[COL.CARGO] || "---",        
            site: row[COL.SITE] || "---",          
            region: row[COL.REGION],                
            estado: String(row[COL.ESTADO]).toUpperCase(),
            correo: row[COL.CORREO],
            contacto: row[COL.CONTACTO],
            estadoNegColect: row[COL.ESTADO_NEG_COLECT] || "",
            banco: row[COL.BANCO] || "",
            tipoCuenta: row[COL.TIPO_CUENTA] || "",
            numeroCuenta: row[COL.NUMERO_CUENTA] || "",
            estadoCredencial: obtenerEstadoCredencialPorRut(row[COL.RUT])
          }
        };
      }
    }
    return { success: false, message: "Datos no encontrados." };
  } catch (e) {
    return { success: false, message: "Error Datos: " + e.toString() };
  }
}

/**
 * Actualizar Dato Usuario
 */
function actualizarDatoUsuario(rutInput, campo, valor) {
  var lock = LockService.getScriptLock();
  if (lock.tryLock(30000)) { // ✅ Aumentado a 30 segundos para alta concurrencia
    try {
      const sheet = getSheet('USUARIOS', 'USUARIOS');
      const data = sheet.getDataRange().getValues();
      const rutLimpioInput = cleanRut(rutInput);
      const COL = CONFIG.COLUMNAS.USUARIOS;
     
      let colIndex = -1;
      if (campo === 'region') colIndex = COL.REGION;        
      else if (campo === 'correo') colIndex = COL.CORREO;
      else if (campo === 'contacto') colIndex = COL.CONTACTO;
      else if (campo === 'banco') colIndex = COL.BANCO;
      else if (campo === 'tipoCuenta') colIndex = COL.TIPO_CUENTA;
      else if (campo === 'numeroCuenta') colIndex = COL.NUMERO_CUENTA;
     
      if (colIndex === -1) return { success: false, message: "Campo inválido" };

      for (let i = 1; i < data.length; i++) {
        if (cleanRut(String(data[i][COL.RUT])) === rutLimpioInput) {
          sheet.getRange(i + 1, colIndex + 1).setValue(valor);
          
          // ✅ NUEVO: Invalidar caché del usuario
          var cache = CacheService.getScriptCache();
          cache.remove('user_' + rutLimpioInput);
          
          return { success: true, message: "OK" };
        }
      }
      return { success: false, message: "Usuario no hallado para editar" };
    } catch (e) {
      return { success: false, message: "Error Update: " + e.toString() };
    } finally {
      lock.releaseLock();
    }
  } else {
    return { success: false, message: "Servidor ocupado." };
  }
}

/**
 * RECUPERAR CONTRASEÑA
 */
function recuperarContrasena(rutInput) {
  try {
    const sheet = getSheet('USUARIOS', 'USUARIOS');
    const data = sheet.getDataRange().getDisplayValues();
    const rutLimpio = cleanRut(rutInput);
    const COL = CONFIG.COLUMNAS.USUARIOS;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (cleanRut(row[COL.RUT]) === rutLimpio) {
        const correo = row[COL.CORREO];
        return { success: true, correo: correo || "No registrado" };
      }
    }
    
    return { success: false, message: "Usuario no encontrado." };
  } catch (e) {
    return { success: false, message: "Error: " + e.toString() };
  }
}

function enviarContrasenaCorreo(rutInput) {
  try {
    const sheet = getSheet('USUARIOS', 'USUARIOS');
    const data = sheet.getDataRange().getDisplayValues();
    const rutLimpio = cleanRut(rutInput);
    const COL = CONFIG.COLUMNAS.USUARIOS;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (cleanRut(row[COL.RUT]) === rutLimpio) {
        const nombre = row[COL.NOMBRE];
        const correo = row[COL.CORREO];
        const password = row[COL.ID_CREDENCIAL];
        
        if (!correo || !correo.includes("@")) {
          return { success: false, message: "No tienes un correo registrado. Contacta con la directiva." };
        }
        
        enviarCorreoEstilizado(
          correo,
          "Recuperación de Contraseña - Sindicato SLIM n°3",
          "Recuperación de Contraseña",
          `Hola ${nombre}, has solicitado recuperar tu contraseña de acceso al portal.`,
          {
            "Tu contraseña es": password,
            "RUT": row[COL.RUT]
          },
          "#3b82f6"
        );
        
        return { success: true, message: "Contraseña enviada exitosamente." };
      }
    }
    
    return { success: false, message: "Usuario no encontrado." };
  } catch (e) {
    return { success: false, message: "Error: " + e.toString() };
  }
}

// ==========================================
// LÓGICA DE PRÉSTAMOS
// ==========================================

function crearSolicitudPrestamo(rutGestor, tipo, cuotas, medioPago, rutBeneficiario) {
  var lock = LockService.getScriptLock();
  if (lock.tryLock(30000)) { // ✅ Aumentado a 30 segundos para alta concurrencia
    try {
      const sheetUsers = getSheet('USUARIOS', 'USUARIOS');
      const sheetPrestamos = getSheet('PRESTAMOS', 'PRESTAMOS');
      const COL_USER = CONFIG.COLUMNAS.USUARIOS;
      const COL_PRES = CONFIG.COLUMNAS.PRESTAMOS;

      const dataUsers = sheetUsers.getDataRange().getDisplayValues();
      
      // 1. Identificar al Gestor
      let gestor = null;
      const rutLimpioGestor = cleanRut(rutGestor);
      for (let i = 1; i < dataUsers.length; i++) {
        if (cleanRut(dataUsers[i][COL_USER.RUT]) === rutLimpioGestor) {
          gestor = {
             rut: dataUsers[i][COL_USER.RUT],
             nombre: dataUsers[i][COL_USER.NOMBRE],
             correo: dataUsers[i][COL_USER.CORREO]
          };
          break;
        }
      }
      if (!gestor) return { success: false, message: "Error de sesión." };

      // 2. Identificar al Beneficiario
      let rutTarget = rutBeneficiario ? cleanRut(rutBeneficiario) : rutLimpioGestor;
      let beneficiario = null;
      let esGestionDirigente = (rutTarget !== rutLimpioGestor);

      if (!esGestionDirigente) {
         beneficiario = gestor;
      } else {
         for (let i = 1; i < dataUsers.length; i++) {
            if (cleanRut(dataUsers[i][COL_USER.RUT]) === rutTarget) {
               beneficiario = {
                  rut: dataUsers[i][COL_USER.RUT],
                  nombre: dataUsers[i][COL_USER.NOMBRE],
                  correo: dataUsers[i][COL_USER.CORREO]
               };
               break;
            }
         }
         if (!beneficiario) return { success: false, message: "RUT del socio no encontrado." };
      }

      // 3. Validar préstamos activos (LÓGICA NUEVA: Por Tipo)
      const dataPrestamos = sheetPrestamos.getDataRange().getDisplayValues();
      for (let i = 1; i < dataPrestamos.length; i++) {
        const row = dataPrestamos[i];
        const rowRut = cleanRut(row[COL_PRES.RUT]);
        const rowEstado = row[COL_PRES.ESTADO];
        const rowTipo = row[COL_PRES.TIPO]; // Leemos el tipo de la fila
        
        const estadosActivos = ["Solicitado", "Enviado", "Vigente"];
        
        // CAMBIO AQUI: Validamos RUT + ESTADO + TIPO IGUAL
        // Usamos .includes() para ser flexibles (ej: "Préstamo de Emergencia" vs "Emergencia")
        if (rowRut === cleanRut(beneficiario.rut) && 
            estadosActivos.includes(rowEstado) && 
            rowTipo.includes(tipo)) {
              
          return { success: false, message: `El socio ${beneficiario.nombre} ya tiene un préstamo de tipo "${tipo}" en estado "${rowEstado}".` };
        }
      }

      // --- LOGICA DEL MONTO ---
      let montoTexto = "$0";
      if (tipo.includes('Emergencia')) {
        montoTexto = "$200.000";
      } else { 
        montoTexto = "$150.000";
      }

      // 4. Preparar Datos y CALCULAR FECHA TÉRMINO (Lógica Contable)
      const fechaSolicitud = new Date();
      const diaSolicitud = fechaSolicitud.getDate();
      const idUnico = Utilities.getUuid();
      
      let fechaInicioPago = new Date(fechaSolicitud);
      
      if (diaSolicitud > 24) {
        fechaInicioPago.setMonth(fechaInicioPago.getMonth() + 1);
      }
      
      let fechaTermino = new Date(fechaInicioPago);
      let numCuotas = parseInt(cuotas);
      
      if (!isNaN(numCuotas)) {
        fechaTermino.setMonth(fechaTermino.getMonth() + numCuotas);
        fechaTermino = new Date(fechaTermino.getFullYear(), fechaTermino.getMonth() + 1, 0);
      }

      let gestion = esGestionDirigente ? "Dirigente" : "Socio";
      let nomDirigente = esGestionDirigente ? gestor.nombre : "";
      let correoDirigente = esGestionDirigente ? gestor.correo : "";

      // 5. Guardar en Base de Datos
      const newRow = [];
      newRow[COL_PRES.ID] = idUnico;
      newRow[COL_PRES.FECHA] = fechaSolicitud;
      newRow[COL_PRES.RUT] = beneficiario.rut;
      newRow[COL_PRES.NOMBRE] = beneficiario.nombre;
      newRow[COL_PRES.CORREO] = beneficiario.correo;
      newRow[COL_PRES.TIPO] = tipo;
      newRow[COL_PRES.MONTO] = "'" + montoTexto; 
      newRow[COL_PRES.CUOTAS] = cuotas;
      newRow[COL_PRES.MEDIO_PAGO] = medioPago;
      newRow[COL_PRES.ESTADO] = "Solicitado";
      newRow[COL_PRES.FECHA_TERMINO] = fechaTermino;
      newRow[COL_PRES.GESTION] = gestion;
      newRow[COL_PRES.NOMBRE_DIRIGENTE] = nomDirigente;
      newRow[COL_PRES.CORREO_DIRIGENTE] = correoDirigente;
      newRow[COL_PRES.INFORME] = ""; 

      sheetPrestamos.appendRow(newRow);

      // 6. Enviar Correos
      if (esCorreoValido(beneficiario.correo)) {
        var datosCorreoSocio = {
            "FECHA SOLICITUD": Utilities.formatDate(fechaSolicitud, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
            "RUT": formatRutServer(beneficiario.rut),
            "NOMBRE": beneficiario.nombre,
            "TIPO PRÉSTAMO": tipo,
            "MONTO": montoTexto,
            "CUOTAS": cuotas,
            "MEDIO PAGO": medioPago,
            "FECHA TÉRMINO": Utilities.formatDate(fechaTermino, Session.getScriptTimeZone(), "dd/MM/yyyy"), 
            "GESTION": gestion,
            "NOMBRE DIRIGENTE": nomDirigente || ""
        };

        enviarCorreoEstilizado(
          beneficiario.correo,
          "Solicitud de Préstamo - Sindicato SLIM n°3",
          "Solicitud de Préstamo Ingresada",
          `Hola <strong>${beneficiario.nombre}</strong>, se ha ingresado exitosamente una solicitud de préstamo a tu nombre.`,
          datosCorreoSocio,
          "#2563eb"
        );
      }

      // Correo Dirigente
      if (esGestionDirigente && esCorreoValido(correoDirigente) && correoDirigente !== beneficiario.correo) {
        var datosCorreoDirigente = {
            "FECHA SOLICITUD": Utilities.formatDate(fechaSolicitud, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
            "RUT SOCIO": formatRutServer(beneficiario.rut),
            "NOMBRE SOCIO": beneficiario.nombre,
            "TIPO PRÉSTAMO": tipo,
            "MONTO": montoTexto,
            "CUOTAS": cuotas,
            "FECHA TÉRMINO": Utilities.formatDate(fechaTermino, Session.getScriptTimeZone(), "dd/MM/yyyy"),
            "GESTION": "Dirigente"
        };

        enviarCorreoEstilizado(
          gestor.correo,
          "Respaldo Gestión Préstamo - Sindicato SLIM n°3",
          "Solicitud de Préstamo Creada",
          `Has ingresado una solicitud de préstamo para el socio <strong>${beneficiario.nombre}</strong>.`,
          datosCorreoDirigente,
          "#475569"
        );
      }

      return { success: true, message: "Solicitud creada exitosamente." };
    } catch (e) {
      return { success: false, message: "Error al solicitar: " + e.toString() };
    } finally {
      lock.releaseLock();
    }
  } else {
    return { success: false, message: "Servidor ocupado." };
  }
}

// ==========================================
// SINCRONIZACIÓN AUTOMÁTICA (Validación -> BD -> Notificación)
// VERSIÓN ACTUALIZADA: Incluye MONTO en el correo + Corrección columna OK
// ==========================================

function procesarValidacionPrestamos() {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(60000)) {
    try {
      const ss = getSpreadsheet('PRESTAMOS');
      const sheetValidacion = ss.getSheetByName(CONFIG.HOJAS.VALIDACION_PRESTAMOS);
      const sheetBD = getSheet('PRESTAMOS', 'PRESTAMOS');
      
      if (!sheetValidacion) {
        console.warn("⚠️ La hoja 'Validación-Prestamos' no existe. Creándola...");
        const nuevaHoja = ss.insertSheet(CONFIG.HOJAS.VALIDACION_PRESTAMOS);
        nuevaHoja.appendRow(["ID", "Fecha", "RUT", "Nombre", "Validación", "Observación", "Nombre Informe"]);
        console.log("✅ Hoja 'Validación-Prestamos' creada exitosamente");
        return;
      }
      
      if (!sheetBD) {
        console.error("❌ No se encontró la hoja BD_PRESTAMOS.");
        return;
      }

      const dataValidacion = sheetValidacion.getDataRange().getValues();
      const dataBD = sheetBD.getDataRange().getValues();
      const COL_BD = CONFIG.COLUMNAS.PRESTAMOS;
      
      // MAPEO DE COLUMNAS HOJA VALIDACIÓN:
      // A(0): ID | B(1): Fecha | C(2): RUT | D(3): Nombre | E(4): Validación | F(5): Observación | G(6): Nombre Informe
      const VAL_COL = { ID: 0, VALIDACION: 4, OBS: 5 }; 
      
      // ⭐ CORRECCIÓN: Columna O (índice 14) para marcar "OK" - NO la columna N
      const COL_INFORME = 14; // ← CAMBIO AQUÍ (antes era 13)

      let procesadosCount = 0;

      // Recorrer hoja de Validación (saltando cabecera)
      for (let i = 1; i < dataValidacion.length; i++) {
        const idSolicitud = String(dataValidacion[i][VAL_COL.ID]).trim();
        const resultadoValidacion = String(dataValidacion[i][VAL_COL.VALIDACION]).toUpperCase().trim();
        const observacionAdmin = String(dataValidacion[i][VAL_COL.OBS]);

        // Solo procesamos si hay ID y una decisión clara (ACEPTADO o RECHAZADO)
        if (!idSolicitud || (resultadoValidacion !== "ACEPTADO" && resultadoValidacion !== "RECHAZADO")) {
          continue;
        }

        // Buscar coincidencia en BD_PRESTAMOS
        for (let j = 1; j < dataBD.length; j++) {
          const idBD = String(dataBD[j][COL_BD.ID]).trim();
          
          // ⭐ Verificar columna O (Informe) para ver si ya fue procesado antes
          const informeEnviado = String(dataBD[j][COL_INFORME]); 

          if (idBD === idSolicitud) {
            
            // SI YA DICE "OK", SALTAMOS (Ya fue procesado históricamente)
            if (informeEnviado === "OK") {
               console.log(`ℹ️ ID ${idSolicitud}: Ya procesado anteriormente (OK en columna O)`);
               continue; 
            }

            // Si llegamos aquí, es una solicitud nueva validada que requiere acción
            let nuevoEstado = "";
            let tituloCorreo = "";
            let colorCorreo = "";
            let mensajeIntro = "";

            if (resultadoValidacion === "ACEPTADO") {
              nuevoEstado = "Vigente"; 
              tituloCorreo = "Solicitud Aprobada";
              colorCorreo = "#15803d"; // Verde
              mensajeIntro = `Nos complace informarte que tu solicitud de préstamo ha sido <strong>APROBADA</strong> por la empresa.`;
            } else {
              nuevoEstado = "Rechazado";
              tituloCorreo = "Solicitud Rechazada";
              colorCorreo = "#b91c1c"; // Rojo
              mensajeIntro = `Te informamos que tu solicitud de préstamo ha sido <strong>RECHAZADA</strong> por la empresa.`;
            }

            // 1. Actualizar estado en BD_PRESTAMOS
            sheetBD.getRange(j + 1, COL_BD.ESTADO + 1).setValue(nuevoEstado);
            
            // 2. Enviar Correo con MONTO incluido
            const correoUsuario = dataBD[j][COL_BD.CORREO];
            const nombreUsuario = dataBD[j][COL_BD.NOMBRE];
            
            if (esCorreoValido(correoUsuario)) {
              // ⭐ PREPARAR FECHA TÉRMINO FORMATEADA
              let fechaTerminoStr = "S/D";
              const fechaTerminoRaw = dataBD[j][COL_BD.FECHA_TERMINO];
              
              if (fechaTerminoRaw) {
                try {
                  const fechaTermino = new Date(fechaTerminoRaw);
                  if (!isNaN(fechaTermino.getTime())) {
                    fechaTerminoStr = Utilities.formatDate(fechaTermino, Session.getScriptTimeZone(), "dd/MM/yyyy");
                  }
                } catch (e) {
                  console.warn(`⚠️ Error formateando fecha término para ID ${idSolicitud}: ${e}`);
                }
              }
              
              // ⭐ AGREGAR FECHA TÉRMINO AL CORREO
              const datosCorreo = {
                "FECHA SOLICITUD": Utilities.formatDate(new Date(dataBD[j][COL_BD.FECHA]), Session.getScriptTimeZone(), "dd/MM/yyyy"),
                "RUT": formatRutServer(dataBD[j][COL_BD.RUT]),
                "NOMBRE": nombreUsuario,
                "TIPO PRÉSTAMO": dataBD[j][COL_BD.TIPO],
                "MONTO": dataBD[j][COL_BD.MONTO] || "$0",
                "ESTADO": nuevoEstado.toUpperCase(),
                "FECHA TÉRMINO": fechaTerminoStr, // ← NUEVO CAMPO
                "OBSERVACIÓN": observacionAdmin || "Sin observaciones",
                "RESULTADO": resultadoValidacion
              };

              enviarCorreoEstilizado(
                correoUsuario,
                `Resultado Solicitud Préstamo - Sindicato SLIM n°3`,
                tituloCorreo,
                `Hola <strong>${nombreUsuario}</strong>, ${mensajeIntro}`,
                datosCorreo,
                colorCorreo
              );
              
              // ⭐ 3. MARCAR COMO PROCESADO en Columna O (índice 14)
              sheetBD.getRange(j + 1, COL_INFORME + 1).setValue("OK");
              console.log(`✅ ID ${idSolicitud}: Procesado como ${nuevoEstado}, notificado y marcado OK en columna O.`);
              procesadosCount++;
              
            } else {
              sheetBD.getRange(j + 1, COL_INFORME + 1).setValue("ERROR_NO_MAIL");
              console.warn(`⚠️ ID ${idSolicitud}: Procesado sin correo válido.`);
            }
            
            break; // Terminar búsqueda en BD para este ID específico
          }
        }
      }
      
      if (procesadosCount > 0) {
         console.log(`✅ Resumen final: ${procesadosCount} solicitudes nuevas procesadas correctamente.`);
      } else {
         console.log("ℹ️ No hay solicitudes nuevas para procesar en este momento.");
      }

    } catch (e) {
      console.error("❌ Error en sincronización de validación de préstamos: " + e.toString());
    } finally {
      lock.releaseLock();
    }
  } else {
    console.warn("⚠️ No se pudo obtener el lock del script. Servidor ocupado.");
  }
}

function obtenerHistorialPrestamos(rutInput) {
  try {
    var sheet = getSheet('PRESTAMOS', 'PRESTAMOS');
    var COL = CONFIG.COLUMNAS.PRESTAMOS;
    
    // ✅ VALIDACIÓN 1: Verificar que la hoja existe
    if (!sheet) {
      Logger.log('❌ Hoja PRESTAMOS no encontrada');
      return { success: false, message: "Hoja no encontrada" };
    }
    
    var lastRow = sheet.getLastRow();
    Logger.log('📊 Préstamos - Total de filas: ' + lastRow);
    
    if (lastRow < 2) return { success: true, registros: [] };
    
    // ✅ CORRECCIÓN 2: Leer TODAS las columnas que existen en la hoja
    var lastCol = sheet.getLastColumn();
    Logger.log('📊 Préstamos - Total de columnas: ' + lastCol);
    
    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
    Logger.log('📊 Datos leídos: ' + data.length + ' filas');
    
    var rutLimpio = cleanRut(rutInput);
    Logger.log('🔍 Buscando préstamos para RUT: ' + rutLimpio);
    
    var registros = [];

    // ✅ CORRECCIÓN 3: Empezar desde índice 0 (data ya no incluye header)
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      
      // ✅ VALIDACIÓN 4: Verificar que la fila tiene datos
      if (!row[COL.RUT]) {
        Logger.log('⚠️ Fila ' + (i + 2) + ' sin RUT, saltando...');
        continue;
      }
      
      const rutFila = cleanRut(row[COL.RUT]);
      
      // ✅ LOGGING para debugging
      if (i < 5) { // Solo loguear las primeras 5 para no saturar
        Logger.log('Fila ' + (i + 2) + ' - RUT: ' + rutFila + ' | Buscado: ' + rutLimpio + ' | Coincide: ' + (rutFila === rutLimpio));
      }
      
      if (rutFila !== rutLimpio) continue;
      
      // ✅ EXTRACCIÓN de datos con valores por defecto
      const tipo = row[COL.TIPO] || "Préstamo";
      const cuotas = row[COL.CUOTAS] || "S/D";
      const medio = row[COL.MEDIO_PAGO] || "S/D";
      const monto = row[COL.MONTO] || "$0";
      const observacion = row[COL.OBSERVACION] || "";
      
      // ✅ FORMATEO de fecha de término
      let fechaTerminoStr = "S/D";
      const ftRaw = row[COL.FECHA_TERMINO];
      if (ftRaw) {
         try {
           const d = new Date(ftRaw);
           if (!isNaN(d.getTime())) {
             fechaTerminoStr = Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy");
           } else {
             fechaTerminoStr = String(ftRaw).split(' ')[0]; 
           }
         } catch(e) { 
           fechaTerminoStr = String(ftRaw).split(' ')[0]; 
         }
      }

      registros.push({
      id: row[COL.ID] || "",
      fecha: formatearFechaConHora(row[COL.FECHA]) || "",
      tipo: tipo,
      monto: monto,
      cuotas: cuotas,
      medio: medio,
      estado: row[COL.ESTADO] || "Solicitado",
      observacion: observacion,
      fechaTermino: formatearFechaSinHora(row[COL.FECHA_TERMINO]),
      gestion: row[COL.GESTION] || "Socio",
      nomDirigente: row[COL.NOMBRE_DIRIGENTE] || ""
    });
      
      Logger.log('✅ Préstamo agregado: ' + row[COL.ID] + ' - ' + tipo);
    }
    
    Logger.log('📦 Total de préstamos encontrados: ' + registros.length);
    
    registros.reverse();
    return { success: true, registros: registros };

  } catch (e) {
    Logger.log('❌ ERROR en obtenerHistorialPrestamos: ' + e.toString());
    Logger.log('Stack: ' + e.stack);
    return { success: false, message: "Error: " + e.toString() };
  }
}

// ==========================================
// FUNCIÓN ELIMINAR PRÉSTAMO (Con Respaldo Histórico)
// ==========================================

function eliminarSolicitud(idSolicitud) {
  var lock = LockService.getScriptLock();
  if (lock.tryLock(30000)) { // ✅ Aumentado a 30 segundos para alta concurrencia
    try {
      const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEETS.PRESTAMOS);
      const sheet = ss.getSheetByName("BD_PRESTAMOS");
      const data = sheet.getDataRange().getValues();
      const COL = CONFIG.COLUMNAS.PRESTAMOS;
      
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][COL.ID]) === String(idSolicitud)) {
          
          // RESPALDO: Guardamos copia SIEMPRE antes de borrar
          const sheetEliminados = ss.getSheetByName("Registros-eliminados");
          if (sheetEliminados) {
            sheetEliminados.appendRow(data[i]);
          } else {
            return { success: false, message: "Error crítico: No existe la hoja de respaldo." };
          }
          
          // Eliminar fila
          sheet.deleteRow(i + 1);
          return { success: true, message: "Registro eliminado y respaldado correctamente." };
        }
      }
      return { success: false, message: "No encontrado." };
    } catch (e) { return { success: false, message: "Error: " + e.toString() }; }
    finally { lock.releaseLock(); }
  } else { return { success: false, message: "Servidor ocupado." }; }
}

function modificarSolicitud(idSolicitud, nuevasCuotas, nuevoMedio) {
  var lock = LockService.getScriptLock();
  if (lock.tryLock(30000)) { // ✅ Aumentado a 30 segundos para alta concurrencia
    try {
      const sheet = getSheet('PRESTAMOS', 'PRESTAMOS');
      const data = sheet.getDataRange().getValues();
      const COL = CONFIG.COLUMNAS.PRESTAMOS;
      
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][COL.ID]) === String(idSolicitud)) {
          const estado = String(data[i][COL.ESTADO]);
          if (estado !== "Solicitado") return { success: false, message: "No se puede editar. Estado: " + estado };
          
          // Recalcular Fecha Término con lógica financiera
          const fechaSolicitud = new Date(data[i][COL.FECHA]);
          const diaSolicitud = fechaSolicitud.getDate();
          
          let fechaInicioPago = new Date(fechaSolicitud);
          if (diaSolicitud > 24) {
             fechaInicioPago.setMonth(fechaInicioPago.getMonth() + 1);
          }
          
          let fechaTermino = new Date(fechaInicioPago);
          fechaTermino.setMonth(fechaTermino.getMonth() + parseInt(nuevasCuotas));
          
          // Ajustar al último día del mes
          fechaTermino = new Date(fechaTermino.getFullYear(), fechaTermino.getMonth() + 1, 0);
          
          sheet.getRange(i + 1, COL.FECHA_TERMINO + 1).setValue(fechaTermino);

          return { success: true, message: "Modificado correctamente." };
        }
      }
      return { success: false, message: "No encontrado." };
    } catch (e) { return { success: false, message: "Error: " + e.toString() }; }
    finally { lock.releaseLock(); }
  } else { return { success: false, message: "Servidor ocupado." }; }
}

/**
 * Verificar y actualizar préstamos que ya cumplieron su fecha de término
 * Cambia de "Vigente" → "Pagado" automáticamente
 */
function verificarCambiosPrestamos() {
  try {
    // ⭐ CORRECCIÓN: Usar getSheet() en lugar de getActiveSpreadsheet()
    const sheet = getSheet('PRESTAMOS', 'PRESTAMOS');
    
    if (!sheet) {
      console.error("❌ No se pudo acceder a la hoja BD_PRESTAMOS");
      return { success: false, error: "Hoja no encontrada" };
    }
    
    const data = sheet.getDataRange().getValues();
    const COL = CONFIG.COLUMNAS.PRESTAMOS;
    
    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);
    
    let prestamosActualizados = 0;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const estado = String(row[COL.ESTADO]).trim();
      const fechaTerminoRaw = row[COL.FECHA_TERMINO];
      
      if (estado !== "Vigente") continue;
      
      if (!fechaTerminoRaw) {
        console.warn(`⚠️ Fila ${i + 1}: Préstamo vigente sin fecha de término`);
        continue;
      }
      
      let fechaTermino;
      try {
        fechaTermino = new Date(fechaTerminoRaw);
        fechaTermino.setHours(0, 0, 0, 0);
      } catch (e) {
        console.error(`❌ Fila ${i + 1}: Error al convertir fecha: ${fechaTerminoRaw}`);
        continue;
      }
      
      if (isNaN(fechaTermino.getTime())) {
        console.error(`❌ Fila ${i + 1}: Fecha de término inválida: ${fechaTerminoRaw}`);
        continue;
      }
      
      if (hoy > fechaTermino) {
        sheet.getRange(i + 1, COL.ESTADO + 1).setValue("Pagado");
        
        const correo = row[COL.CORREO];
        const nombre = row[COL.NOMBRE];
        const tipo = row[COL.TIPO];
        const monto = row[COL.MONTO];
        const cuotas = row[COL.CUOTAS];
        const idPrestamo = row[COL.ID];
        
        console.log(`✅ Fila ${i + 1}: Préstamo ID ${idPrestamo} cambiado a "Pagado"`);
        
        if (esCorreoValido(correo)) {
          try {
            enviarCorreoEstilizado(
              correo,
              "Préstamo Completado - Sindicato SLIM n°3",
              "Préstamo Finalizado",
              `Hola <strong>${nombre}</strong>, tu préstamo ha sido completado exitosamente.`,
              { 
                "ID": idPrestamo,
                "TIPO PRÉSTAMO": tipo,
                "MONTO": monto,
                "CUOTAS": cuotas,
                "ESTADO": "PAGADO",
                "FECHA TÉRMINO": Utilities.formatDate(fechaTermino, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
                "FECHA FINALIZACIÓN": Utilities.formatDate(hoy, Session.getScriptTimeZone(), 'dd/MM/yyyy')
              },
              "#10b981"
            );
          } catch (mailError) {
            console.error(`⚠️ Error enviando correo: ${mailError}`);
          }
        }
        
        prestamosActualizados++;
      }
    }
    
    if (prestamosActualizados > 0) {
      console.log(`📊 RESUMEN: ${prestamosActualizados} préstamo(s) actualizado(s) a "Pagado"`);
    } else {
      console.log("ℹ️ No hay préstamos que actualizar.");
    }
    
    return { success: true, prestamosActualizados: prestamosActualizados };
    
  } catch (e) {
    console.error("❌ Error verificando préstamos: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

// ==========================================
// SWITCH MÓDULO PRÉSTAMOS
// ==========================================

/**
 * Obtener estado del switch de préstamos.
 * Por defecto habilitado (null = primera ejecución).
 */
function obtenerEstadoSwitchPrestamos() {
  try {
    var props = PropertiesService.getScriptProperties();
    var estado = props.getProperty('prestamos_habilitado');
    var habilitado = (estado === null || estado === 'true');
    return { success: true, habilitado: habilitado };
  } catch (e) {
    Logger.log('Error en obtenerEstadoSwitchPrestamos: ' + e.toString());
    return { success: true, habilitado: true };
  }
}

/**
 * Actualizar estado del switch de préstamos (solo ADMIN).
 */
function toggleSwitchPrestamos(estado) {
  try {
    var props = PropertiesService.getScriptProperties();
    props.setProperty('prestamos_habilitado', estado ? 'true' : 'false');
    return { success: true };
  } catch (e) {
    return { success: false, message: 'Error: ' + e.toString() };
  }
}

// ==========================================
// LÓGICA DE JUSTIFICACIONES (CON SWITCH)
// ==========================================

/**
 * Obtener estado del switch de justificaciones
 */
function obtenerEstadoSwitchJustificaciones() {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get('justif_switch_state');
    if (cached) {
      try {
        return JSON.parse(cached);
      } catch (e) {
        Logger.log('Error parsing switch cache: ' + e);
      }
    }

    const ss = getSpreadsheet('JUSTIFICACIONES');
    let sheetConfig = ss.getSheetByName(CONFIG.HOJAS.CONFIG_JUSTIFICACIONES);
    
    if (!sheetConfig) {
      sheetConfig = ss.insertSheet(CONFIG.HOJAS.CONFIG_JUSTIFICACIONES);
      sheetConfig.appendRow(["Habilitado", "Fecha Límite"]);
      sheetConfig.appendRow([false, ""]);
    }
    
    const data = sheetConfig.getRange(2, 1, 1, 3).getValues();
    const habilitado = data[0][0] === true || data[0][0] === "TRUE" || data[0][0] === "true";
    const fechaLimiteValue = data[0][1];
    const fechaEventoRaw = data[0][2];
    const fechaEvento = (fechaEventoRaw && String(fechaEventoRaw).trim() !== "") 
      ? (fechaEventoRaw instanceof Date 
          ? Utilities.formatDate(fechaEventoRaw, Session.getScriptTimeZone(), "yyyy-MM-dd")
          : String(fechaEventoRaw).trim())
      : null;
    
    if (habilitado && fechaLimiteValue) {
      // Convertir ambas fechas a UTC para comparación justa
      const ahora = new Date();
      const limite = new Date(fechaLimiteValue);
      
      // Log para debug (puedes comentar después)
      Logger.log("Fecha actual (UTC): " + ahora.toISOString());
      Logger.log("Fecha límite (UTC): " + limite.toISOString());
      Logger.log("Fecha actual > Fecha límite: " + (ahora > limite));
      
      if (ahora > limite) {
        // Ya pasó la fecha límite, deshabilitar automáticamente
        sheetConfig.getRange(2, 1).setValue(false);
        var resultado = { 
        habilitado: habilitado, 
        fechaLimite: fechaLimiteValue,
        fechaEvento: fechaEvento
      };
      
      // ✅ Guardar en caché por 5 minutos
      try {
        cache.put('justif_switch_state', JSON.stringify(resultado), 300);
      } catch (e) {
        Logger.log('Error caching switch state: ' + e);
      }
      
      return resultado;
      }
    }
    
    return { 
      habilitado: habilitado, 
      fechaLimite: fechaLimiteValue,
      fechaEvento: fechaEvento
    };
    
  } catch (e) {
    Logger.error("Error en obtenerEstadoSwitchJustificaciones: " + e.toString());
    return { 
      habilitado: false, 
      fechaLimite: "", 
      error: e.toString() 
    };
  }
}

/**
 * Actualizar estado del switch de justificaciones
 */
function actualizarSwitchJustificaciones(nuevoEstado, fechaLimite, fechaEvento) {
  var lock = LockService.getScriptLock();
  if (lock.tryLock(30000)) { // ✅ Aumentado a 30 segundos para alta concurrencia
    try {
      const ss = getSpreadsheet('JUSTIFICACIONES');
      let sheetConfig = ss.getSheetByName(CONFIG.HOJAS.CONFIG_JUSTIFICACIONES);
      
      if (!sheetConfig) {
        sheetConfig = ss.insertSheet(CONFIG.HOJAS.CONFIG_JUSTIFICACIONES);
        sheetConfig.appendRow(["Habilitado", "Fecha Límite", "Fecha_Evento"]);
        sheetConfig.appendRow([false, "", ""]);
      }
      
      sheetConfig.getRange(2, 1).setValue(nuevoEstado);
      if (fechaLimite) {
        sheetConfig.getRange(2, 2).setValue(fechaLimite);
      }
      const valorEvento = (fechaEvento && String(fechaEvento).trim() !== "") ? String(fechaEvento).trim() : "";
      sheetConfig.getRange(2, 3).setValue(valorEvento);
      
      return { success: true, message: "Estado actualizado correctamente." };
      
    } catch (e) {
      return { success: false, message: "Error: " + e.toString() };
    } finally {
      lock.releaseLock();
    }
  } else {
    return { success: false, message: "Servidor ocupado." };
  }
}

function verificarDisponibilidadJustificaciones() {
  const estadoSwitch = obtenerEstadoSwitchJustificaciones();
  
  if (!estadoSwitch.habilitado) {
    return { 
      habilitado: false, 
      mensaje: "Módulo de justificaciones temporalmente deshabilitado.\nConsulte con la directiva." 
    };
  }
  
  return { habilitado: true };
}

/**
 * Valida si el usuario puede enviar una justificación
 * Si hay evento configurado: valida por código de evento (ASAMBLEA)
 * Si no hay evento: valida por mes calendario (comportamiento anterior)
 * @param {string} rut - RUT del usuario
 * @returns {Object} {permitido: boolean, mensaje: string, justificacionExistente: Object|null}
 */
function validarJustificacionMesActual(rut) {
  try {
    SpreadsheetApp.flush(); // Fuerza sincronización antes de leer
    const sheet = getSheet('JUSTIFICACIONES', 'JUSTIFICACIONES');
    const data = sheet.getDataRange().getValues();
    const COL = CONFIG.COLUMNAS.JUSTIFICACIONES;
    const hoy = new Date();
    
    // Obtener configuración del evento activo
    const configSwitch = obtenerEstadoSwitchJustificaciones();
    const fechaEvento = configSwitch.fechaEvento || null;
    const codigoEventoActivo = fechaEvento ? generarCodigoAsambleaEvento(fechaEvento) : null;
    
    Logger.log(`✅ Validando justificación | Evento: ${codigoEventoActivo || 'Sin evento (modo mes)'}`);
    
    // ========================
    // MODO A: Validación por evento específico
    // ========================
    if (codigoEventoActivo) {
      const justificacionesDelEvento = [];
      
      for (let i = 1; i < data.length; i++) {
        const filaRut = data[i][COL.RUT];
        const filaAsamblea = String(data[i][COL.ASAMBLEA] || "").trim();
        const filaEstado = data[i][COL.ESTADO];
        const filaFecha = data[i][COL.FECHA];
        
        if (cleanRut(filaRut) === cleanRut(rut) && filaAsamblea === codigoEventoActivo) {
          justificacionesDelEvento.push({
            id: data[i][COL.ID],
            estado: filaEstado,
            tipo: data[i][COL.MOTIVO],
            fecha: filaFecha ? Utilities.formatDate(new Date(filaFecha), Session.getScriptTimeZone(), "dd/MM/yyyy") : "",
            asamblea: filaAsamblea
          });
        }
      }
      
      if (justificacionesDelEvento.length === 0) {
        return { permitido: true, mensaje: "Puede enviar la justificación", justificacionExistente: null };
      }
      
      const todasRechazadas = justificacionesDelEvento.every(j => j.estado === 'Rechazado');
      if (todasRechazadas) {
        return { permitido: true, mensaje: "Puede reintentar (anterior rechazada)", justificacionExistente: justificacionesDelEvento[0] };
      }
      
      const justEnviada = justificacionesDelEvento.find(j => j.estado === 'Enviado');
      const justAceptada = justificacionesDelEvento.find(j => j.estado === 'Aceptado' || j.estado === 'Aceptado/Obs');
      
      if (justEnviada) {
        return {
          permitido: false,
          mensaje: `Ya tienes una justificación pendiente para el evento ${codigoEventoActivo}`,
          justificacionExistente: justEnviada,
          tipoBloqueo: 'enviada',
          codigoEvento: codigoEventoActivo
        };
      }
      
      if (justAceptada) {
        return {
          permitido: false,
          mensaje: `Ya tienes una justificación aceptada para el evento ${codigoEventoActivo}`,
          justificacionExistente: justAceptada,
          tipoBloqueo: 'aceptada',
          codigoEvento: codigoEventoActivo
        };
      }
      
      // Caso por defecto
      return {
        permitido: false,
        mensaje: `Ya existe una justificación para el evento ${codigoEventoActivo}`,
        justificacionExistente: justificacionesDelEvento[0],
        tipoBloqueo: 'enviada',
        codigoEvento: codigoEventoActivo
      };
    }
    
    // ========================
    // MODO B: Fallback — validación por mes calendario (comportamiento anterior)
    // ========================
    const mesActual = hoy.getMonth();
    const yearActual = hoy.getFullYear();
    const justificacionesDelMes = [];
    
    for (let i = 1; i < data.length; i++) {
      const filaRut = data[i][COL.RUT];
      const filaFecha = new Date(data[i][COL.FECHA]);
      const filaEstado = data[i][COL.ESTADO];
      const filaAsamblea = data[i][COL.ASAMBLEA];
      
      if (cleanRut(filaRut) === cleanRut(rut) && 
          filaFecha.getMonth() === mesActual && 
          filaFecha.getFullYear() === yearActual) {
        justificacionesDelMes.push({
          id: data[i][COL.ID],
          estado: filaEstado,
          tipo: data[i][COL.MOTIVO],
          fecha: Utilities.formatDate(filaFecha, Session.getScriptTimeZone(), "dd/MM/yyyy"),
          asamblea: filaAsamblea || ""
        });
      }
    }
    
    if (justificacionesDelMes.length === 0) {
      return { permitido: true, mensaje: "Puede enviar la justificación", justificacionExistente: null };
    }
    
    const nombreMes = hoy.toLocaleString('es-CL', { month: 'long', year: 'numeric' });
    const todasRechazadas = justificacionesDelMes.every(j => j.estado === 'Rechazado');
    
    if (todasRechazadas) {
      return { permitido: true, mensaje: "Puede reintentar (anterior rechazada)", justificacionExistente: justificacionesDelMes[0] };
    }
    
    const hayEnviada = justificacionesDelMes.some(j => j.estado === 'Enviado');
    const hayAceptada = justificacionesDelMes.some(j => j.estado === 'Aceptado' || j.estado === 'Aceptado/Obs');
    
    if (hayEnviada) {
      const justificacion = justificacionesDelMes.find(j => j.estado === 'Enviado');
      return {
        permitido: false,
        mensaje: `Ya tienes una justificación pendiente para ${nombreMes}`,
        justificacionExistente: justificacion,
        tipoBloqueo: 'enviada'
      };
    }
    
    if (hayAceptada) {
      const justificacion = justificacionesDelMes.find(j => j.estado === 'Aceptado' || j.estado === 'Aceptado/Obs');
      return {
        permitido: false,
        mensaje: `Ya tienes una justificación aceptada para ${nombreMes}`,
        justificacionExistente: justificacion,
        tipoBloqueo: 'aceptada'
      };
    }
    
    return {
      permitido: false,
      mensaje: `Límite de justificaciones alcanzado para ${nombreMes}`,
      justificacionExistente: justificacionesDelMes[0],
      tipoBloqueo: 'enviada'
    };
    
  } catch (error) {
    Logger.log('Error en validarJustificacionMesActual: ' + error.toString());
    return {
      permitido: false,
      mensaje: "Error al validar: " + error.message,
      justificacionExistente: null
    };
  }
}

// ==========================================
// FUNCIÓN ENVIAR JUSTIFICACIÓN - VERSIÓN MEJORADA
// Reemplazar la función existente completamente
// ==========================================

function enviarJustificacion(rutGestor, tipo, motivo, archivoData, rutBeneficiario) {
  var CARPETA_ID = CONFIG.CARPETAS.JUSTIFICACIONES;
  
  // Verificar disponibilidad del módulo
  var disp = verificarDisponibilidadJustificaciones();
  if (!disp.habilitado) return { success: false, message: disp.mensaje };
  
  var lock = LockService.getScriptLock();
  if (lock.tryLock(30000)) {
    try {
      var sheetJustif = getSheet('JUSTIFICACIONES', 'JUSTIFICACIONES');
      var COL_JUST = CONFIG.COLUMNAS.JUSTIFICACIONES;
      
      // Obtener datos del gestor
      var gestor = obtenerUsuarioPorRut(rutGestor);
      if (!gestor.encontrado) return { success: false, message: "Error de sesión." };
      
      // Determinar beneficiario
      var beneficiario;
      var rutTarget = rutBeneficiario ? cleanRut(rutBeneficiario) : cleanRut(rutGestor);
      var esGestionDirigente = rutTarget !== cleanRut(rutGestor);
      
      if (!esGestionDirigente) {
        beneficiario = gestor;
      } else {
        beneficiario = obtenerUsuarioPorRut(rutBeneficiario);
        if (!beneficiario.encontrado) return { success: false, message: "RUT del socio no encontrado." };
      }
      
      // Validar justificación del mes
      var validacion = validarJustificacionMesActual(beneficiario.rut);
      if (!validacion.permitido) {
        return {
          success: false,
          message: validacion.mensaje,
          tipoError: 'restriccion_mes',
          justificacionExistente: validacion.justificacionExistente,
          tipoBloqueo: validacion.tipoBloqueo,
          codigoEvento: validacion.codigoEvento || null
        };
      }
      
      // ========== VALIDAR CORREOS ANTES DE SUBIR ARCHIVO ==========
      var validacionCorreos = validarCorreosParaPermisos(
        { rut: beneficiario.rut, nombre: beneficiario.nombre, correo: beneficiario.correo },
        esGestionDirigente ? { rut: gestor.rut, nombre: gestor.nombre, correo: gestor.correo } : null,
        esGestionDirigente
      );
      
      var idUnico = Utilities.getUuid();
      var fileUrl = "Sin archivo";
      var alertaPermisos = null;
      
      // ========== SUBIR ARCHIVO SI EXISTE ==========
      if (archivoData && archivoData.base64) {
        var nombreArchivo = "JUSTIF-" + idUnico + "-" + cleanRut(beneficiario.rut);
        
        var resultadoSubida = subirArchivoConPermisos(
          archivoData,
          CARPETA_ID,
          nombreArchivo,
          validacionCorreos.correosParaPermisos,
          [] // Sin correos adicionales para justificaciones
        );
        
        if (!resultadoSubida.success) {
          return { success: false, message: resultadoSubida.mensajeError };
        }
        
        fileUrl = resultadoSubida.url;
        
        // Generar alerta de permisos si hay problemas
        alertaPermisos = generarAlertaPermisos(validacionCorreos, resultadoSubida);
      } else {
        // Si no hay archivo, igual verificar si hay alertas de correo
        alertaPermisos = generarAlertaPermisos(validacionCorreos, null);
      }
      
      // ========== CREAR REGISTRO EN LA BASE DE DATOS ==========
      var fechaHoy = new Date();
      var estado = "Enviado";
      var configActiva = obtenerEstadoSwitchJustificaciones();
      var codigoAsamblea = configActiva.fechaEvento 
        ? generarCodigoAsambleaEvento(configActiva.fechaEvento) 
        : generarCodigoAsamblea(fechaHoy);
      
      var gestion = "Socio";
      var nomDirigente = "";
      var correoDirigente = "";
      
      if (esGestionDirigente) {
        gestion = "Dirigente";
        nomDirigente = gestor.nombre;
        correoDirigente = gestor.correo;
      }
      
      var newRow = [];
      newRow[COL_JUST.ID] = idUnico;
      newRow[COL_JUST.FECHA] = fechaHoy;
      newRow[COL_JUST.RUT] = beneficiario.rut;
      newRow[COL_JUST.NOMBRE] = beneficiario.nombre;
      newRow[COL_JUST.REGION] = beneficiario.region;
      newRow[COL_JUST.MOTIVO] = tipo;
      newRow[COL_JUST.ARGUMENTO] = motivo;
      newRow[COL_JUST.RESPALDO] = fileUrl;
      newRow[COL_JUST.ESTADO] = estado;
      newRow[COL_JUST.OBSERVACION] = "";
      newRow[COL_JUST.NOTIFICACION] = estado;
      newRow[COL_JUST.ASAMBLEA] = codigoAsamblea;
      newRow[COL_JUST.GESTION] = gestion;
      newRow[COL_JUST.DIRIGENTE] = nomDirigente;
      newRow[COL_JUST.CORREO_DIRIGENTE] = correoDirigente;
      
      sheetJustif.appendRow(newRow);
      
      // Agregar validación de datos
      var lastRow = sheetJustif.getLastRow();
      var cellEstado = sheetJustif.getRange(lastRow, COL_JUST.ESTADO + 1);
      var rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Enviado', 'Aceptado', 'Aceptado/Obs', 'Rechazado'], true)
        .setAllowInvalid(false)
        .build();
      cellEstado.setDataValidation(rule);
      
      // ========== ENVIAR CORREOS ==========
      // Correo al socio (solo cuando gestiona por sí mismo)
      if (!esGestionDirigente && esCorreoValido(beneficiario.correo)) {
        
        // Construimos el link del archivo o S/D si no hay URL válida
        let respaldoDisplay = "";
        if (fileUrl && fileUrl.includes("http")) {
           respaldoDisplay = `<a href="${fileUrl}" style="color: ${'#ea580c'}; text-decoration: none; font-weight: bold;">Ver Documento Adjunto</a>`;
        } else {
           respaldoDisplay = ""; // Se convertirá en S/D automáticamente
        }

        // Datos formateados exactamente como se solicitaron
        var datosCorreo = {
            "FECHA": Utilities.formatDate(fechaHoy, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
            "RUT": formatRutServer(beneficiario.rut),
            "NOMBRE": beneficiario.nombre,
            "REGION": beneficiario.region, // Asegúrate que este dato venga de 'obtenerUsuarioPorRut'
            "MOTIVO": tipo,
            "ARGUMENTO": motivo,
            "RESPALDO": respaldoDisplay,
            "OBSERVACION": "", // Observación inicial suele estar vacía
            "ASAMBLEA": codigoAsamblea,
            "GESTION": gestion,
            "DIRIGENTE": nomDirigente
        };

        enviarCorreoEstilizado(
          beneficiario.correo,
          "Justificación Ingresada - Sindicato SLIM n°3",
          "Comprobante de Justificación",
          `Hola <strong>${beneficiario.nombre}</strong>, tu justificación ha sido ingresada correctamente en el sistema. A continuación los detalles registrados:`,
          datosCorreo,
          "#ea580c" // Color naranja para justificaciones
        );
      }
      
      // Correo de respaldo al dirigente (si aplica)
      if (esGestionDirigente && esCorreoValido(correoDirigente) && correoDirigente !== beneficiario.correo) {
         
         // 1. Construimos el enlace específicamente para el diseño del dirigente (color #475569)
         let respaldoDisplayDirigente = "";
         if (fileUrl && fileUrl.includes("http")) {
            respaldoDisplayDirigente = `<a href="${fileUrl}" style="color: #475569; text-decoration: none; font-weight: bold;">Ver Documento Adjunto</a>`;
         } else {
            respaldoDisplayDirigente = ""; // Se convertirá en S/D automáticamente
         }

         // 2. Construimos el objeto de datos con el enlace incluido
         var datosCorreoDirigente = {
            "FECHA": Utilities.formatDate(fechaHoy, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
            "RUT": formatRutServer(beneficiario.rut),
            "NOMBRE": beneficiario.nombre,
            "REGION": beneficiario.region,
            "MOTIVO": tipo,
            "ARGUMENTO": motivo,
            "RESPALDO": respaldoDisplayDirigente, // <--- AQUI ESTA EL CAMBIO (Antes decía "Documento Cargado")
            "OBSERVACION": "",
            "ASAMBLEA": codigoAsamblea,
            "GESTION": gestion,
            "DIRIGENTE": nomDirigente
        };

        enviarCorreoEstilizado(
          correoDirigente,
          "Respaldo Gestión Justificación - Sindicato SLIM n°3",
          "Gestión Realizada",
          `Has ingresado exitosamente una justificación para el socio <strong>${beneficiario.nombre}</strong>.`,
          datosCorreoDirigente,
          "#475569" // Color gris/azul para administración
        );
      }

      // Copia al socio cuando el dirigente gestiona en su nombre
      if (esGestionDirigente && esCorreoValido(beneficiario.correo)) {
        var respaldoDisplaySocio = "";
        if (fileUrl && fileUrl.includes("http")) {
          respaldoDisplaySocio = "<a href=\"" + fileUrl + "\" style=\"color: #ea580c; text-decoration: none; font-weight: bold;\">Ver Documento Adjunto</a>";
        }

        var datosCorreoSocio = {
          "FECHA": Utilities.formatDate(fechaHoy, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
          "RUT": formatRutServer(beneficiario.rut),
          "NOMBRE": beneficiario.nombre,
          "REGION": beneficiario.region,
          "MOTIVO": tipo,
          "ARGUMENTO": motivo,
          "RESPALDO": respaldoDisplaySocio,
          "OBSERVACION": "",
          "ASAMBLEA": codigoAsamblea,
          "GESTION": gestion,
          "DIRIGENTE": nomDirigente
        };

        enviarCorreoEstilizado(
          beneficiario.correo,
          "Justificaci\u00f3n Ingresada - Sindicato SLIM n\u00b03",
          "Comprobante de Justificaci\u00f3n",
          "Hola <strong>" + beneficiario.nombre + "</strong>, un dirigente ha ingresado una justificaci\u00f3n a tu nombre. A continuaci\u00f3n los detalles registrados:",
          datosCorreoSocio,
          "#ea580c"
        );
      }
      
      // ========== PREPARAR RESPUESTA ==========
      var respuesta = {
        success: true,
        message: "Justificación enviada exitosamente."
      };
      
      // Agregar alerta si hay problemas con permisos
      if (alertaPermisos && alertaPermisos.mostrarAlerta) {
        respuesta.mostrarAlerta = true;
        respuesta.tipoAlerta = alertaPermisos.tipoAlerta;
        respuesta.mensajeAlerta = alertaPermisos.mensajeAlerta;
      }
      
      return respuesta;
      
    } catch (e) {
      Logger.log("Error en enviarJustificacion: " + e.toString());
      return { success: false, message: "Error: " + e.toString() };
    } finally {
      lock.releaseLock();
    }
  } else {
    return { success: false, message: "Servidor ocupado." };
  }
}

function eliminarJustificacion(idJustif) {
  var lock = LockService.getScriptLock();
  if (lock.tryLock(30000)) { // ✅ Aumentado a 30 segundos para alta concurrencia
    try {
      const sheet = getSheet('JUSTIFICACIONES', 'JUSTIFICACIONES');
      const data = sheet.getDataRange().getValues();
      const COL = CONFIG.COLUMNAS.JUSTIFICACIONES;
      
      const estadoSwitch = obtenerEstadoSwitchJustificaciones();
      
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][COL.ID]) === String(idJustif)) {
          const estado = String(data[i][COL.ESTADO]);
          
          if (!estadoSwitch.habilitado && estado === "Enviado") {
            return { 
              success: false, 
              message: "El plazo para agregar o modificar información ha vencido. Si al final del mes aparece con multa puede realizar la apelación." 
            };
          }
          
          if (estado !== "Enviado") {
            return { success: false, message: "No se puede eliminar." };
          }
          
          sheet.deleteRow(i + 1); 
          return { success: true, message: "Eliminado." };
        }
      }
      return { success: false, message: "No encontrado." };
    } catch (e) { 
      return { success: false, message: "Error: " + e.toString() }; 
    } finally { 
      lock.releaseLock(); 
    }
  } else { 
    return { success: false, message: "Ocupado." }; 
  }
} 

function obtenerHistorialJustificaciones(rutInput) {
  try {
    const sheet = getSheet('JUSTIFICACIONES', 'JUSTIFICACIONES');
    const COL = CONFIG.COLUMNAS.JUSTIFICACIONES;
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, registros: [] };
    
    // ⭐ CORRECCIÓN: Calcular correctamente el número de columnas
    var lastCol = sheet.getLastColumn();
    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
    
    const rutLimpio = cleanRut(rutInput);
    const registros = [];

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (cleanRut(row[COL.RUT]) === rutLimpio) {
        registros.push({
          id: row[COL.ID],
          fecha: formatearFechaConHora(row[COL.FECHA]),
          tipo: row[COL.MOTIVO],
          motivo: row[COL.ARGUMENTO],
          url: row[COL.RESPALDO],
          estado: row[COL.ESTADO],
          obs: row[COL.OBSERVACION],
          asamblea: row[COL.ASAMBLEA],
          gestion: row[COL.GESTION],    
          nomDirigente: row[COL.DIRIGENTE]
        });
      }
    }
    
    registros.reverse();
    return { success: true, registros: registros };
  } catch (e) { 
    Logger.log("❌ Error en obtenerHistorialJustificaciones: " + e.toString());
    return { success: false, message: "Error: " + e.toString() }; 
  }
}

function verificarCambiosJustificaciones() {
  try {
    // ⭐ VALIDACIÓN
    const sheet = getSheet('JUSTIFICACIONES', 'JUSTIFICACIONES');
    if (!sheet) {
      console.error("❌ No se pudo acceder a la hoja de justificaciones");
      return;
    }
    
    const data = sheet.getDataRange().getValues();
    const COL = CONFIG.COLUMNAS.JUSTIFICACIONES;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const idRegistro = String(row[COL.ID]);
      const estadoActual = String(row[COL.ESTADO]);
      const estadoNotif = String(row[COL.NOTIFICACION]);
      const nombre = row[COL.NOMBRE];
      const tipo = row[COL.MOTIVO];
      const obs = row[COL.OBSERVACION];
      const asamblea = row[COL.ASAMBLEA];
      
      // ⭐ Verificar y corregir código de asamblea si falta
      const fechaSolicitud = row[COL.FECHA];
      const asambleaActual = row[COL.ASAMBLEA];
      
      if (fechaSolicitud && !asambleaActual) {
        const codigoAsamblea = generarCodigoAsamblea(new Date(fechaSolicitud));
        sheet.getRange(i + 1, COL.ASAMBLEA + 1).setValue(codigoAsamblea);
        console.log(`✅ Código de asamblea generado para fila ${i + 1}: ${codigoAsamblea}`);
      }
      
      if (estadoActual !== estadoNotif) {
        const rutUsuario = row[COL.RUT];
        
        // ⭐ VALIDACIÓN: Verificar que existe la hoja de usuarios
        const sheetUsers = getSheet('USUARIOS', 'USUARIOS');
        if (!sheetUsers) {
          console.error("❌ No se pudo acceder a la hoja de usuarios");
          continue;
        }
        
        const dataUsers = sheetUsers.getDataRange().getDisplayValues();
        const COL_USER = CONFIG.COLUMNAS.USUARIOS;
        let correoUsuario = "";
        
        for (let j = 1; j < dataUsers.length; j++) {
          if (cleanRut(dataUsers[j][COL_USER.RUT]) === cleanRut(rutUsuario)) {
            correoUsuario = dataUsers[j][COL_USER.CORREO];
            break;
          }
        }
        
        if (correoUsuario && correoUsuario.includes("@")) {
          let color = "#ea580c";
          let titulo = "Actualización de Justificación";
          
          if (estadoActual.includes("Aceptado")) { 
            color = "#15803d"; 
            titulo = "Justificación Aceptada"; 
          } else if (estadoActual.includes("Rechazado")) { 
            color = "#b91c1c"; 
            titulo = "Justificación Rechazada"; 
          }
          
          enviarCorreoEstilizado(
            correoUsuario, 
            titulo + " - Sindicato SLIM n°3", 
            titulo, 
            `Hola ${nombre}, el estado de tu justificación ha cambiado.`, 
            { 
              "ID": idRegistro,
              "Tipo": tipo, 
              "Nuevo Estado": estadoActual, 
              "Observación": obs || "Sin observaciones",
              "Asamblea": asamblea || "Pendiente asignación"
            }, 
            color
          );
        }
        
        sheet.getRange(i + 1, COL.NOTIFICACION + 1).setValue(estadoActual);
      }
    }
  } catch (e) { 
    console.error("❌ Error verificando justificaciones: " + e.toString()); 
  }
}

// ==========================================
// LÓGICA DE APELACIONES
// ==========================================

function verificarDisponibilidadApelaciones(mesApelacion) {
  try {
    const hoy = new Date();
    const diaActual = hoy.getDate();
    
    const limiteInferior = new Date(2025, 2, 1);
    limiteInferior.setHours(0, 0, 0, 0);
    
    const partes = mesApelacion.split("-");
    const yearSel = parseInt(partes[0]);
    const monthSel = parseInt(partes[1]) - 1;
    const fechaSeleccionada = new Date(yearSel, monthSel, 1);
    fechaSeleccionada.setHours(0, 0, 0, 0);
    
    if (fechaSeleccionada < limiteInferior) {
      return { 
        habilitado: false, 
        mensaje: "No se pueden apelar meses anteriores a Marzo 2025." 
      };
    }
    
    const mesActual = hoy.getMonth();
    const yearActual = hoy.getFullYear();
    
    if (yearSel === yearActual && monthSel === mesActual) {
      if (diaActual < 25) {
        return {
          habilitado: false,
          mensaje: "Las apelaciones del mes en curso solo están disponibles a partir del día 25."
        };
      }
    }
    
    const fechaHoy = new Date(yearActual, mesActual, 1);
    fechaHoy.setHours(0, 0, 0, 0);
    
    if (fechaSeleccionada > fechaHoy) {
      return {
        habilitado: false,
        mensaje: "No se pueden apelar meses futuros."
      };
    }
    
    return { habilitado: true };
    
  } catch (e) {
    return { habilitado: false, mensaje: "Error validando disponibilidad: " + e.toString() };
  }
}

// ==========================================
// FUNCIÓN ENVIAR APELACIÓN - VERSIÓN ACTUALIZADA (DISEÑO + SILENCIO)
// Reemplazar la función existente completamente
// ==========================================

function enviarApelacion(rutGestor, mesApelacion, tipoMotivo, detalleMotivo, archivoComprobante, archivoLiquidacion, rutBeneficiario) {
  var CARPETA_COMPROBANTES_ID = CONFIG.CARPETAS.APELACIONES_COMPROBANTES;
  var CARPETA_LIQUIDACIONES_ID = CONFIG.CARPETAS.APELACIONES_LIQUIDACIONES;
  
  // Validar disponibilidad
  var validacion = verificarDisponibilidadApelaciones(mesApelacion);
  if (!validacion.habilitado) {
    return { success: false, message: validacion.mensaje };
  }
  
  var lock = LockService.getScriptLock();
  if (lock.tryLock(30000)) {
    try {
      var sheetApelaciones = getSheet('APELACIONES', 'APELACIONES');
      var COL_APEL = CONFIG.COLUMNAS.APELACIONES;
      
      // Obtener datos del gestor
      var gestor = obtenerUsuarioPorRut(rutGestor);
      if (!gestor.encontrado) return { success: false, message: "Error de sesión." };
      
      // Determinar beneficiario
      var beneficiario;
      var rutTarget = rutBeneficiario ? cleanRut(rutBeneficiario) : cleanRut(rutGestor);
      var esGestionDirigente = rutTarget !== cleanRut(rutGestor);
      
      if (!esGestionDirigente) {
        beneficiario = gestor;
      } else {
        beneficiario = obtenerUsuarioPorRut(rutBeneficiario);
        if (!beneficiario.encontrado) return { success: false, message: "RUT del socio no encontrado." };
      }
      
      // Verificar apelaciones existentes
      var dataApelaciones = sheetApelaciones.getDataRange().getDisplayValues();
      for (var i = 1; i < dataApelaciones.length; i++) {
        var row = dataApelaciones[i];
        var estadoActual = String(row[COL_APEL.ESTADO]);
        var estadosBloqueantes = ["Enviado", "Aceptado", "Aceptado-Obs"];
        
        if (cleanRut(row[COL_APEL.RUT]) === cleanRut(beneficiario.rut) &&
            row[COL_APEL.MES_APELACION] === mesApelacion &&
            estadosBloqueantes.indexOf(estadoActual) !== -1) {
          
          var mensajeError = "";
          if (estadoActual === "Enviado") {
            mensajeError = "Ya tienes una apelación pendiente para este mes. Verifica el estado en tu historial.";
          } else {
            mensajeError = "Este mes ya fue resuelto favorablemente. Verifica los detalles en tu historial.";
          }
          
          return { success: false, message: mensajeError };
        }
      }
      
      // Validar liquidación obligatoria
      if (!archivoLiquidacion || !archivoLiquidacion.base64) {
        return { success: false, message: "La liquidación de sueldo es obligatoria." };
      }
      
      // ========== VALIDAR CORREOS ANTES DE SUBIR ARCHIVOS ==========
      var validacionCorreos = validarCorreosParaPermisos(
        { rut: beneficiario.rut, nombre: beneficiario.nombre, correo: beneficiario.correo },
        esGestionDirigente ? { rut: gestor.rut, nombre: gestor.nombre, correo: gestor.correo } : null,
        esGestionDirigente
      );
      
      var idUnico = Utilities.getUuid();
      var urlComprobante = ""; // Se guardará vacío en BD si no hay
      var urlLiquidacion = "";
      var alertaPermisosGlobal = { mostrarAlerta: false, detalles: [] };
      
      // ========== SUBIR COMPROBANTE (Opcional) ==========
      if (archivoComprobante && archivoComprobante.base64) {
        var nombreArchivoComp = "APEL-COMP-" + idUnico + "-" + cleanRut(beneficiario.rut);
        
        var resultadoComp = subirArchivoConPermisos(
          archivoComprobante,
          CARPETA_COMPROBANTES_ID,
          nombreArchivoComp,
          validacionCorreos.correosParaPermisos, // Usa la función silenciosa automáticamente
          []
        );
        
        if (resultadoComp.success) {
          urlComprobante = resultadoComp.url;
          if (resultadoComp.permisosError && resultadoComp.permisosError.length > 0) {
            alertaPermisosGlobal.mostrarAlerta = true;
            resultadoComp.permisosError.forEach(function(err) {
              alertaPermisosGlobal.detalles.push("Comprobante: No se pudo dar acceso a " + err.nombre);
            });
          }
        } else {
          // Si falla la subida opcional, registramos error pero continuamos
          console.error("Error subiendo comprobante: " + resultadoComp.mensajeError);
        }
      }
      
      // ========== SUBIR LIQUIDACIÓN (Obligatoria) ==========
      var nombreArchivoLiq = "APEL-LIQ-" + idUnico + "-" + cleanRut(beneficiario.rut);
      
      var resultadoLiq = subirArchivoConPermisos(
        archivoLiquidacion,
        CARPETA_LIQUIDACIONES_ID,
        nombreArchivoLiq,
        validacionCorreos.correosParaPermisos, // Usa la función silenciosa automáticamente
        []
      );
      
      if (!resultadoLiq.success) {
        return { success: false, message: "Error al subir la liquidación: " + resultadoLiq.mensajeError };
      }
      
      urlLiquidacion = resultadoLiq.url;
      
      if (resultadoLiq.permisosError && resultadoLiq.permisosError.length > 0) {
        alertaPermisosGlobal.mostrarAlerta = true;
        resultadoLiq.permisosError.forEach(function(err) {
          alertaPermisosGlobal.detalles.push("Liquidación: No se pudo dar acceso a " + err.nombre);
        });
      }
      
      // Agregar alertas de validación de correos
      if (validacionCorreos.alertas && validacionCorreos.alertas.length > 0) {
        alertaPermisosGlobal.mostrarAlerta = true;
        validacionCorreos.alertas.forEach(function(a) {
          alertaPermisosGlobal.detalles.push(a.mensaje);
        });
      }
      
      // ========== CREAR REGISTRO EN LA BASE DE DATOS ==========
      var fechaHoy = new Date();
      var estado = "Enviado";
      
      var gestion = "Socio";
      var nomDirigente = "";
      var correoDirigente = "";
      
      if (esGestionDirigente) {
        gestion = "Dirigente";
        nomDirigente = gestor.nombre;
        correoDirigente = gestor.correo;
      }
      
      var newRow = [];
      newRow[COL_APEL.ID] = idUnico;
      newRow[COL_APEL.FECHA_SOLICITUD] = fechaHoy;
      newRow[COL_APEL.RUT] = beneficiario.rut;
      newRow[COL_APEL.NOMBRE] = beneficiario.nombre;
      newRow[COL_APEL.CORREO] = beneficiario.correo;
      newRow[COL_APEL.MES_APELACION] = mesApelacion;
      newRow[COL_APEL.TIPO_MOTIVO] = tipoMotivo;
      newRow[COL_APEL.DETALLE_MOTIVO] = detalleMotivo || "";
      newRow[COL_APEL.URL_COMPROBANTE] = urlComprobante;
      newRow[COL_APEL.URL_LIQUIDACION] = urlLiquidacion;
      newRow[COL_APEL.ESTADO] = estado;
      newRow[COL_APEL.OBSERVACION] = "";
      newRow[COL_APEL.NOTIFICADO] = estado;
      newRow[COL_APEL.GESTION] = gestion;
      newRow[COL_APEL.NOMBRE_DIRIGENTE] = nomDirigente;
      newRow[COL_APEL.CORREO_DIRIGENTE] = correoDirigente;
      newRow[COL_APEL.URL_COMPROBANTE_DEVOLUCION] = "";
      
      sheetApelaciones.appendRow(newRow);

      // Forzar formato texto en la celda MES_APELACION para evitar conversión automática a Date
      var lastRowApel = sheetApelaciones.getLastRow();
      sheetApelaciones.getRange(lastRowApel, COL_APEL.MES_APELACION + 1)
        .setNumberFormat('@STRING@')
        .setValue(mesApelacion);
      
      // Validación de celda (opcional, para integridad)
      var lastRow = sheetApelaciones.getLastRow();
      var cellEstado = sheetApelaciones.getRange(lastRow, COL_APEL.ESTADO + 1);
      var rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Enviado', 'Aceptado', 'Aceptado-Obs', 'Rechazado'], true)
        .setAllowInvalid(false)
        .build();
      cellEstado.setDataValidation(rule);
      
      // Formatear mes para visualización
      var fechaMes = new Date(mesApelacion + "-02");
      var nombreMes = fechaMes.toLocaleString('es-CL', { month: 'long', year: 'numeric' });
      nombreMes = nombreMes.charAt(0).toUpperCase() + nombreMes.slice(1);

      // ========== ENVIAR CORREOS ==========
      
      // 1. Correo al Beneficiario (solo cuando gestiona por sí mismo)
      if (!esGestionDirigente && esCorreoValido(beneficiario.correo)) {
        
        // Construir enlaces HTML con color del tema apelaciones (Rojo #dc2626)
        var linkComprobanteSocio = (urlComprobante && urlComprobante.includes("http")) 
            ? `<a href="${urlComprobante}" style="color: #dc2626; text-decoration: none; font-weight: bold;">Ver Comprobante</a>` 
            : "";
            
        var linkLiquidacionSocio = (urlLiquidacion && urlLiquidacion.includes("http")) 
            ? `<a href="${urlLiquidacion}" style="color: #dc2626; text-decoration: none; font-weight: bold;">Ver Liquidación</a>` 
            : "";

        var datosCorreoSocio = {
            "FECHA SOLICITUD": Utilities.formatDate(fechaHoy, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
            "RUT": formatRutServer(beneficiario.rut),
            "NOMBRE": beneficiario.nombre,
            "MES APELACION": nombreMes,
            "TIPO MOTIVO": tipoMotivo,
            "DETALLE MOTIVO": detalleMotivo || "", // Saldrá S/D si es vacío
            "URL COMPROBANTE": linkComprobanteSocio, // Saldrá S/D si es vacío
            "URL LIQUIDACIÓN": linkLiquidacionSocio, // Saldrá S/D si es vacío
            "OBSERVACIÓN": "",
            "GESTIÓN": gestion,
            "NOMBRE DIRIGENTE": nomDirigente || "" // Saldrá S/D si es vacío
        };

        enviarCorreoEstilizado(
          beneficiario.correo,
          "Apelación Ingresada - Sindicato SLIM n°3",
          "Comprobante de Apelación",
          `Hola <strong>${beneficiario.nombre}</strong>, hemos recibido correctamente tu apelación de multa. A continuación los detalles registrados:`,
          datosCorreoSocio,
          "#1d4ed8" // Color azul institucional para confirmación de ingreso
        );
      }
      
      // 2. Correo al Dirigente (si gestiona a tercero)
      if (esGestionDirigente && esCorreoValido(correoDirigente) && correoDirigente !== beneficiario.correo) {
        
        // Enlaces con color administrativo (#475569)
        var linkComprobanteDirigente = (urlComprobante && urlComprobante.includes("http")) 
            ? `<a href="${urlComprobante}" style="color: #475569; text-decoration: none; font-weight: bold;">Ver Comprobante</a>` 
            : "";
            
        var linkLiquidacionDirigente = (urlLiquidacion && urlLiquidacion.includes("http")) 
            ? `<a href="${urlLiquidacion}" style="color: #475569; text-decoration: none; font-weight: bold;">Ver Liquidación</a>` 
            : "";

        var datosCorreoDirigente = {
            "FECHA SOLICITUD": Utilities.formatDate(fechaHoy, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
            "RUT": formatRutServer(beneficiario.rut),
            "NOMBRE": beneficiario.nombre,
            "MES APELACION": nombreMes,
            "TIPO MOTIVO": tipoMotivo,
            "DETALLE MOTIVO": detalleMotivo || "",
            "URL COMPROBANTE": linkComprobanteDirigente,
            "URL LIQUIDACIÓN": linkLiquidacionDirigente,
            "OBSERVACIÓN": "",
            "GESTIÓN": gestion,
            "NOMBRE DIRIGENTE": nomDirigente
        };

        enviarCorreoEstilizado(
          correoDirigente,
          "Respaldo Gestión Apelación - Sindicato SLIM n°3",
          "Gestión Realizada",
          `Has ingresado exitosamente una apelación para el socio <strong>${beneficiario.nombre}</strong>.`,
          datosCorreoDirigente,
          "#475569" // Color gris/azul
        );
      }

      // ✅ NUEVO: Copia al socio cuando el dirigente gestiona en su nombre
      if (esGestionDirigente && esCorreoValido(beneficiario.correo)) {
        var linkComprobanteCopiaSocio = (urlComprobante && urlComprobante.includes("http"))
            ? "<a href=\"" + urlComprobante + "\" style=\"color: #dc2626; text-decoration: none; font-weight: bold;\">Ver Comprobante</a>"
            : "";
        var linkLiquidacionCopiaSocio = (urlLiquidacion && urlLiquidacion.includes("http"))
            ? "<a href=\"" + urlLiquidacion + "\" style=\"color: #dc2626; text-decoration: none; font-weight: bold;\">Ver Liquidación</a>"
            : "";

        var datosCorreoSocioCopia = {
            "FECHA SOLICITUD": Utilities.formatDate(fechaHoy, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
            "RUT": formatRutServer(beneficiario.rut),
            "NOMBRE": beneficiario.nombre,
            "MES APELACION": nombreMes,
            "TIPO MOTIVO": tipoMotivo,
            "DETALLE MOTIVO": detalleMotivo || "",
            "URL COMPROBANTE": linkComprobanteCopiaSocio,
            "URL LIQUIDACIÓN": linkLiquidacionCopiaSocio,
            "OBSERVACIÓN": "",
            "GESTIÓN": gestion,
            "NOMBRE DIRIGENTE": nomDirigente
        };

        enviarCorreoEstilizado(
          beneficiario.correo,
          "Apelaci\u00f3n Ingresada - Sindicato SLIM n\u00b03",
          "Comprobante de Apelaci\u00f3n",
          "Hola <strong>" + beneficiario.nombre + "</strong>, un dirigente ha ingresado una apelaci\u00f3n de multa a tu nombre. A continuaci\u00f3n los detalles registrados:",
          datosCorreoSocioCopia,
          "#1d4ed8"
        );
      }

      // ========== PREPARAR RESPUESTA ==========
      
      // ========== PREPARAR RESPUESTA ==========
      var respuesta = {
        success: true,
        message: "Apelación enviada exitosamente."
      };
      
      if (alertaPermisosGlobal.mostrarAlerta && alertaPermisosGlobal.detalles.length > 0) {
        respuesta.mostrarAlerta = true;
        respuesta.tipoAlerta = validacionCorreos.alertaBeneficiario ? 'warning' : 'info';
        respuesta.mensajeAlerta = alertaPermisosGlobal.detalles.join('\n\n');
      }
      
      return respuesta;
      
    } catch (e) {
      return { success: false, message: "Error al enviar apelación: " + e.toString() };
    } finally {
      lock.releaseLock();
    }
  } else {
    return { success: false, message: "Servidor ocupado." };
  }
}

function obtenerHistorialApelaciones(rutInput) {
  try {
    const sheet = getSheet('APELACIONES', 'APELACIONES');
    const COL = CONFIG.COLUMNAS.APELACIONES;
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, registros: [] };
    
    // ⭐ CORRECCIÓN: Calcular correctamente el número de columnas
    var lastCol = sheet.getLastColumn();
    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();
    
    const rutLimpio = cleanRut(rutInput);
    const registros = [];
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (cleanRut(row[COL.RUT]) === rutLimpio) {
        registros.push({
          id: row[COL.ID],
          fecha: formatearFechaConHora(row[COL.FECHA_SOLICITUD]),  // ✅ CORREGIDO
          mesApelacion: row[COL.MES_APELACION],
          tipoMotivo: row[COL.TIPO_MOTIVO],
          detalleMotivo: row[COL.DETALLE_MOTIVO],
          urlComprobante: row[COL.URL_COMPROBANTE],
          urlLiquidacion: row[COL.URL_LIQUIDACION],
          estado: row[COL.ESTADO],
          obs: row[COL.OBSERVACION],
          gestion: row[COL.GESTION],
          nomDirigente: row[COL.NOMBRE_DIRIGENTE],
          urlComprobanteDevolucion: row[COL.URL_COMPROBANTE_DEVOLUCION] || ""
        });
      }
    }
    
    registros.reverse();
    return { success: true, registros: registros };
    
  } catch (e) { 
    Logger.log("❌ Error en obtenerHistorialApelaciones: " + e.toString());
    return { success: false, message: "Error: " + e.toString() }; 
  }
}

function eliminarApelacion(idApelacion) {
  var lock = LockService.getScriptLock();
  if (lock.tryLock(30000)) { // ✅ Aumentado a 30 segundos para alta concurrencia
    try {
      const sheet = getSheet('APELACIONES', 'APELACIONES');
      const data = sheet.getDataRange().getValues();
      const COL = CONFIG.COLUMNAS.APELACIONES;
      
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][COL.ID]) === String(idApelacion)) {
          const estado = String(data[i][COL.ESTADO]);
          
          // CAMBIO AQUI: Permitimos eliminar si es "Enviado" O "Rechazado"
          if (estado !== "Enviado" && estado !== "Rechazado") {
            return { success: false, message: "Solo se pueden eliminar apelaciones en estado 'Enviado' o 'Rechazado'." };
          }
          
          sheet.deleteRow(i + 1);
          return { success: true, message: "Apelación eliminada correctamente." };
        }
      }
      
      return { success: false, message: "Apelación no encontrada." };
      
    } catch (e) { 
      return { success: false, message: "Error: " + e.toString() }; 
    } finally { 
      lock.releaseLock(); 
    }
  } else { 
    return { success: false, message: "Servidor ocupado." }; 
  }
}

function verificarCambiosApelaciones() {
  try {
    // ⭐ VALIDACIÓN
    const sheet = getSheet('APELACIONES', 'APELACIONES');
    if (!sheet) {
      console.error("❌ No se pudo acceder a la hoja de apelaciones");
      return;
    }
    
    const data = sheet.getDataRange().getDisplayValues();
    const COL = CONFIG.COLUMNAS.APELACIONES;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const idRegistro = String(row[COL.ID]);
      const estadoActual = String(row[COL.ESTADO]);
      const estadoNotif = String(row[COL.NOTIFICADO]);
      const correo = row[COL.CORREO];
      const nombre = row[COL.NOMBRE];
      // Manejar tanto Date object (celda formateada como fecha) como String "yyyy-MM"
      const mesApelRaw = row[COL.MES_APELACION];
      const mesApel = (mesApelRaw instanceof Date)
        ? Utilities.formatDate(mesApelRaw, "GMT", "yyyy-MM")  // GMT lee el valor UTC correcto
        : String(mesApelRaw);
      const tipoMotivo = row[COL.TIPO_MOTIVO];
      const obs = row[COL.OBSERVACION];
      const urlDevolucion = row[COL.URL_COMPROBANTE_DEVOLUCION];
      
      if (estadoActual !== estadoNotif) {
        if (correo && correo.includes("@")) {
          let color = "#ea580c";
          let titulo = "Actualización de Apelación";
          
          if (estadoActual.includes("Aceptado")) { 
            color = "#15803d"; 
            titulo = "Apelación Aceptada"; 
          } else if (estadoActual.includes("Rechazado")) { 
            color = "#b91c1c"; 
            titulo = "Apelación Rechazada"; 
          } else if (estadoActual === "Pagado") {
            color = "#065f46";
            titulo = "Devolución de Multa Procesada";
          }
          
          // Extraer año y mes desde el string "yyyy-MM" (seguro porque getDisplayValues devuelve texto)
          const partesMes = mesApel.split("-");
          const añoMes = parseInt(partesMes[0]);
          const numMes = parseInt(partesMes[1]) - 1; // 0-indexed para el constructor de Date
          const fechaMes = new Date(añoMes, numMes, 15, 12, 0, 0); // Anclado al mediodía local
          const nombreMes = fechaMes.toLocaleString('es-CL', { month: 'long', year: 'numeric' });
          
          // Construir enlace comprobante de devolución si existe y el estado es Pagado
          const linkDevolucion = (estadoActual === "Pagado" && urlDevolucion && String(urlDevolucion).includes("http"))
            ? `<a href="${urlDevolucion}" style="color: #065f46; text-decoration: none; font-weight: bold;">Ver Comprobante de Devolución</a>`
            : "";

          const datosCorreoApelacion = {
            "ID": idRegistro,
            "MES APELADO": nombreMes.toUpperCase(),
            "MOTIVO": tipoMotivo,
            "NUEVO ESTADO": estadoActual,
            "OBSERVACIÓN": obs || "Sin observaciones"
          };

          if (estadoActual === "Pagado" && linkDevolucion) {
            datosCorreoApelacion["COMPROBANTE DEVOLUCIÓN"] = linkDevolucion;
          }

          enviarCorreoEstilizado(
            correo, 
            titulo + " - Sindicato SLIM n°3", 
            titulo, 
            `Hola ${nombre}, el estado de tu apelación ha cambiado.`, 
            datosCorreoApelacion, 
            color
          );
        }
        
        sheet.getRange(i + 1, COL.NOTIFICADO + 1).setValue(estadoActual);
      }
    }
  } catch (e) { 
    console.error("❌ Error verificando apelaciones: " + e.toString()); 
  }
}

function procesarPermisosComprobantesDevolucion() {
  try {
    // ⭐ VALIDACIÓN
    const sheet = getSheet('APELACIONES', 'APELACIONES');
    if (!sheet) {
      console.error("❌ No se pudo acceder a la hoja de apelaciones");
      return;
    }
    
    const data = sheet.getDataRange().getValues();
    const COL = CONFIG.COLUMNAS.APELACIONES;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const urlComprobanteDevolucion = String(row[COL.URL_COMPROBANTE_DEVOLUCION]);
      const correoUsuario = row[COL.CORREO];
      
      if (urlComprobanteDevolucion && 
          urlComprobanteDevolucion.includes("drive.google.com") && 
          correoUsuario && 
          correoUsuario.includes("@")) {
        
        try {
          let fileId = "";
          if (urlComprobanteDevolucion.includes("/d/")) {
            fileId = urlComprobanteDevolucion.split("/d/")[1].split("/")[0];
          } else if (urlComprobanteDevolucion.includes("id=")) {
            fileId = urlComprobanteDevolucion.split("id=")[1].split("&")[0];
          }
          
          if (fileId) {
            const file = DriveApp.getFileById(fileId);
            const viewers = file.getViewers();
            const hasAccess = viewers.some(viewer => viewer.getEmail() === correoUsuario);
            
            if (!hasAccess) {

              // ── INTENTO 1: API Avanzada (sin email de notificación) ──
              let permisoOtorgado = false;
              try {
                const recursoPermiso = {
                  'role': 'reader',
                  'type': 'user',
                  'value': correoUsuario
                };
                Drive.Permissions.insert(recursoPermiso, fileId, {
                  sendNotificationEmails: false
                });
                permisoOtorgado = true;
                console.log(`✅ Permiso silencioso otorgado (API Avanzada) a ${correoUsuario} para archivo ${fileId}`);

              } catch (apiError) {
                console.warn(`⚠️ Fallo API Avanzada para ${correoUsuario} - ${apiError} - Intentando addViewer...`);
                Utilities.sleep(1000);

                // ── INTENTO 2: addViewer ──
                try {
                  file.addViewer(correoUsuario);
                  permisoOtorgado = true;
                  console.log(`✅ Permiso otorgado via addViewer a ${correoUsuario} para archivo ${fileId}`);

                } catch (fallbackError) {
                  console.warn(`⚠️ Fallo addViewer para ${correoUsuario} - ${fallbackError} - Reintentando en 2s...`);
                  Utilities.sleep(2000);

                  // ── INTENTO 3 (ÚLTIMO): Reintento final ──
                  try {
                    file.addViewer(correoUsuario);
                    permisoOtorgado = true;
                    console.log(`✅ Permiso otorgado en reintento final a ${correoUsuario} para archivo ${fileId}`);
                  } catch (finalError) {
                    console.error(`❌ Error fatal al otorgar permiso a ${correoUsuario} para archivo ${fileId}: ${finalError}`);
                  }
                }
              }
            }
          }
        } catch (fileErr) {
          console.error(`⚠️ Error procesando archivo para fila ${i + 1}: ${fileErr}`);
        }
      }
    }
  } catch (e) {
    console.error("❌ Error en procesarPermisosComprobantesDevolucion: " + e.toString());
  }
}

// ==========================================
// MÓDULO: PERMISO MÉDICO - VERSIÓN CORREGIDA V2
// Comparación directa de fechas sin conversiones problemáticas
// ==========================================

function solicitarPermisoMedico(rutGestor, tipoPermiso, fechaInicio, motivo, rutBeneficiario, archivoData) {
  const CORREO_REPRESENTANTE_LEGAL = CONFIG.CORREOS.REPRESENTANTE_LEGAL;
  const CARPETA_ID = CONFIG.CARPETAS.PERMISOS_MEDICOS;
  
  const lock = LockService.getScriptLock();
  if (lock.tryLock(30000)) {
    try {
      const sheetUsers = getSheet('USUARIOS', 'USUARIOS');
      const sheetPermisos = getSheet('PERMISOS_MEDICOS', 'PERMISOS_MEDICOS');
      const COL_USER = CONFIG.COLUMNAS.USUARIOS;
      const COL_PERM = CONFIG.COLUMNAS.PERMISOS_MEDICOS;
      
      const dataUsers = sheetUsers.getDataRange().getDisplayValues();
      
      // ========================================
      // VALIDACIÓN DE USUARIO GESTOR
      // ========================================
      let gestor = null;
      const rutLimpioGestor = cleanRut(rutGestor);
      for (let i = 1; i < dataUsers.length; i++) {
        if (cleanRut(dataUsers[i][COL_USER.RUT]) === rutLimpioGestor) {
          gestor = { 
            rut: dataUsers[i][COL_USER.RUT], 
            nombre: dataUsers[i][COL_USER.NOMBRE], 
            correo: dataUsers[i][COL_USER.CORREO] 
          };
          break;
        }
      }
      if (!gestor) return { success: false, message: "Error de sesión." };
      
      // ========================================
      // DETERMINAR BENEFICIARIO
      // ========================================
      let rutTarget = rutBeneficiario ? cleanRut(rutBeneficiario) : rutLimpioGestor;
      let beneficiario = null;
      
      if (rutTarget === rutLimpioGestor) {
        beneficiario = gestor;
      } else {
        for (let i = 1; i < dataUsers.length; i++) {
          if (cleanRut(dataUsers[i][COL_USER.RUT]) === rutTarget) {
            beneficiario = { 
              rut: dataUsers[i][COL_USER.RUT], 
              nombre: dataUsers[i][COL_USER.NOMBRE], 
              correo: dataUsers[i][COL_USER.CORREO] 
            };
            break;
          }
        }
        if (!beneficiario) return { success: false, message: "RUT del socio no encontrado." };
      }

      // ========================================
      // VALIDACIÓN DE FECHA DE INICIO
      // ========================================
      const fechaInicioObj = new Date(fechaInicio + 'T12:00:00');
      const hoy = new Date();
      hoy.setHours(0, 0, 0, 0);
      fechaInicioObj.setHours(0, 0, 0, 0);

      const diffDias = Math.floor((fechaInicioObj - hoy) / (1000 * 60 * 60 * 24));

      if (diffDias < -7 || diffDias > 7) {
        return {
          success: false,
          message: "La fecha de inicio debe estar dentro del rango de 7 días antes o después de hoy."
        };
      }

      // ========================================
      // ✅ VALIDACIÓN MEJORADA: UN PERMISO ACTIVO POR FECHA DE INICIO
      // Usando comparación directa de strings
      // ========================================
      
      // Normalizar la fecha de entrada a formato yyyy-MM-dd
      const fechaInicioNormalizada = fechaInicio.trim(); // "2026-02-12"
      
      Logger.log(`🔍 Validando para RUT: ${beneficiario.rut} | Fecha Inicio: ${fechaInicioNormalizada}`);

      // Obtener TODOS los datos de permisos
      const dataPermisos = sheetPermisos.getDataRange().getDisplayValues();
      let permisoConMismaFechaInicio = null;
      let permisoAnuladoMismaFecha = null;

      for (let i = 1; i < dataPermisos.length; i++) {
        const rowRut = cleanRut(dataPermisos[i][COL_PERM.RUT]);
        
        if (rowRut === cleanRut(beneficiario.rut)) {
          // Obtener fecha de inicio del registro (columna G, índice 6)
          let fechaInicioRegistro = dataPermisos[i][COL_PERM.FECHA_INICIO];
          
          // Normalizar a string yyyy-MM-dd
          let fechaInicioRegistroNormalizada = "";
          
          if (fechaInicioRegistro && fechaInicioRegistro.toString().trim() !== "") {
            // Si ya viene como "2026-02-12", usarlo directamente
            if (fechaInicioRegistro.toString().match(/^\d{4}-\d{2}-\d{2}$/)) {
              fechaInicioRegistroNormalizada = fechaInicioRegistro.toString().trim();
            } else {
              // Si viene en otro formato, intentar parsearlo
              try {
                const fechaObj = new Date(fechaInicioRegistro);
                if (!isNaN(fechaObj.getTime())) {
                  const year = fechaObj.getFullYear();
                  const month = String(fechaObj.getMonth() + 1).padStart(2, '0');
                  const day = String(fechaObj.getDate()).padStart(2, '0');
                  fechaInicioRegistroNormalizada = `${year}-${month}-${day}`;
                }
              } catch (e) {
                Logger.log(`⚠️ Error parseando fecha en fila ${i + 1}: ${e}`);
                continue;
              }
            }
          }
          
          Logger.log(`  Fila ${i + 1}: Fecha=${fechaInicioRegistroNormalizada} | Estado=${dataPermisos[i][COL_PERM.ESTADO]}`);
          
          // COMPARACIÓN DIRECTA DE STRINGS
          if (fechaInicioRegistroNormalizada === fechaInicioNormalizada) {
            const estadoPermiso = dataPermisos[i][COL_PERM.ESTADO];
            
            Logger.log(`  ⚠️ MATCH ENCONTRADO! Estado: ${estadoPermiso}`);
            
            // Distinguir entre anulados y activos
            if (estadoPermiso === "Anulado") {
              permisoAnuladoMismaFecha = {
                id: dataPermisos[i][COL_PERM.ID],
                tipo: dataPermisos[i][COL_PERM.TIPO_PERMISO],
                fechaSolicitud: dataPermisos[i][COL_PERM.FECHA_SOLICITUD]
              };
              // NO hacer break, seguir buscando por si hay uno activo
            } else {
              // Permiso ACTIVO con la misma fecha de inicio
              permisoConMismaFechaInicio = {
                id: dataPermisos[i][COL_PERM.ID],
                tipo: dataPermisos[i][COL_PERM.TIPO_PERMISO],
                estado: estadoPermiso,
                fechaSolicitud: dataPermisos[i][COL_PERM.FECHA_SOLICITUD],
                fechaInicio: fechaInicioRegistro
              };
              Logger.log(`  🚫 BLOQUEANDO: Permiso activo encontrado ID ${permisoConMismaFechaInicio.id}`);
              break; // Encontramos uno activo, detener búsqueda
            }
          }
        }
      }

      // Si existe un permiso ACTIVO con la misma fecha de inicio, rechazar
      if (permisoConMismaFechaInicio) {
        const fechaSolicitudStr = new Date(permisoConMismaFechaInicio.fechaSolicitud).toLocaleDateString('es-CL');
        
        const mensajeError = `❌ Ya existe un permiso médico ACTIVO con la misma fecha de inicio.\n\n` +
          `📋 DETALLES DEL PERMISO EXISTENTE:\n` +
          `• ID: ${permisoConMismaFechaInicio.id}\n` +
          `• Tipo: ${permisoConMismaFechaInicio.tipo}\n` +
          `• Estado: ${permisoConMismaFechaInicio.estado}\n` +
          `• Fecha de solicitud: ${fechaSolicitudStr}\n` +
          `• Fecha de inicio: ${fechaInicioNormalizada}\n\n` +
          `⚠️ RESTRICCIÓN: No se puede solicitar más de un permiso para la misma fecha de inicio.\n\n` +
          `Si cometió un error, puede anular el permiso existente desde el historial y crear uno nuevo, o contactar con la directiva para solicitar modificaciones.`;
        
        Logger.log(`🚫 SOLICITUD RECHAZADA: ${mensajeError}`);
        
        return { 
          success: false, 
          message: mensajeError
        };
      }

      // Informar si hubo uno anulado (solo para logs)
      if (permisoAnuladoMismaFecha) {
        Logger.log(`ℹ️ INFO: Usuario ya anuló un permiso para la fecha ${fechaInicioNormalizada} (ID: ${permisoAnuladoMismaFecha.id}). Permitiendo crear uno nuevo.`);
      } else {
        Logger.log(`✅ VALIDACIÓN PASADA: No hay permisos activos para la fecha ${fechaInicioNormalizada}`);
      }
      
      // ========================================
      // FIN VALIDACIÓN PRIMERA FASE
      // ========================================

      // ========================================
      // PREPARAR DATOS DEL REGISTRO
      // ========================================
      const fechaHoyCompleta = new Date();
      const idUnico = Utilities.getUuid();
      const estado = "Solicitado";
      
      let gestion = "Socio";
      let nomDirigente = "";
      let correoDirigente = "";
      
      if (rutTarget !== rutLimpioGestor) {
        gestion = "Dirigente";
        nomDirigente = gestor.nombre;
        correoDirigente = gestor.correo;
      }
      
      // Preparar datos según estructura
      const newRow = [];
      newRow[COL_PERM.ID] = idUnico;
      newRow[COL_PERM.FECHA_SOLICITUD] = fechaHoyCompleta;
      newRow[COL_PERM.RUT] = beneficiario.rut;
      newRow[COL_PERM.NOMBRE] = beneficiario.nombre;
      newRow[COL_PERM.CORREO] = beneficiario.correo;
      newRow[COL_PERM.TIPO_PERMISO] = tipoPermiso;
      newRow[COL_PERM.FECHA_INICIO] = fechaInicioNormalizada; // ✅ Guardar como string normalizado
      newRow[COL_PERM.MOTIVO_DETALLE] = motivo;
      // ===== SUBIR DOCUMENTO SI FUE ADJUNTADO DURANTE LA SOLICITUD =====
      let urlDocFinal = "Sin documento";
      let estadoFinal = estado; // "Solicitado" por defecto
      let fechaSubidaFinal = "";

      if (archivoData && archivoData.base64) {
        const correosParaDoc = [];
        if (esCorreoValido(beneficiario.correo)) {
          correosParaDoc.push({ correo: beneficiario.correo.trim().toLowerCase(), tipo: 'beneficiario', nombre: beneficiario.nombre });
        }
        if (rutTarget !== rutLimpioGestor && esCorreoValido(gestor.correo) && gestor.correo.trim().toLowerCase() !== beneficiario.correo.trim().toLowerCase()) {
          correosParaDoc.push({ correo: gestor.correo.trim().toLowerCase(), tipo: 'gestor', nombre: gestor.nombre });
        }

        const nombreArchivo = "PERMISO-" + idUnico + "-" + cleanRut(beneficiario.rut);
        const resultadoSubida = subirArchivoConPermisos(archivoData, CARPETA_ID, nombreArchivo, correosParaDoc, [CORREO_REPRESENTANTE_LEGAL]);

        if (!resultadoSubida.success) {
          return { success: false, message: "Error al subir el documento adjunto: " + resultadoSubida.mensajeError };
        }

        urlDocFinal = resultadoSubida.url;
        estadoFinal = "Documento Adjuntado";
        fechaSubidaFinal = fechaHoyCompleta;
      }
      // ===== FIN SUBIDA DOCUMENTO =====

      newRow[COL_PERM.URL_DOCUMENTO] = urlDocFinal;
      newRow[COL_PERM.ESTADO] = estadoFinal;
      newRow[COL_PERM.FECHA_SUBIDA] = fechaSubidaFinal;
      newRow[COL_PERM.NOTIFICADO_REP_LEGAL] = false;
      newRow[COL_PERM.GESTION] = gestion;
      newRow[COL_PERM.NOMBRE_DIRIGENTE] = nomDirigente;
      newRow[COL_PERM.CORREO_DIRIGENTE] = correoDirigente;
      
      // ========================================
      // 🔥 DOBLE VALIDACIÓN: Prevenir race conditions
      // Re-validar justo antes de escribir
      // ========================================
      Logger.log(`🔄 SEGUNDA VALIDACIÓN (anti-race-condition)...`);
      
      const dataPermisosPreEscritura = sheetPermisos.getDataRange().getDisplayValues();
      for (let i = 1; i < dataPermisosPreEscritura.length; i++) {
        const rowRut = cleanRut(dataPermisosPreEscritura[i][COL_PERM.RUT]);
        if (rowRut === cleanRut(beneficiario.rut)) {
          let fechaInicioRegistro = dataPermisosPreEscritura[i][COL_PERM.FECHA_INICIO];
          let fechaInicioRegistroNormalizada = "";
          
          if (fechaInicioRegistro && fechaInicioRegistro.toString().trim() !== "") {
            if (fechaInicioRegistro.toString().match(/^\d{4}-\d{2}-\d{2}$/)) {
              fechaInicioRegistroNormalizada = fechaInicioRegistro.toString().trim();
            } else {
              try {
                const fechaObj = new Date(fechaInicioRegistro);
                const year = fechaObj.getFullYear();
                const month = String(fechaObj.getMonth() + 1).padStart(2, '0');
                const day = String(fechaObj.getDate()).padStart(2, '0');
                fechaInicioRegistroNormalizada = `${year}-${month}-${day}`;
              } catch (e) {
                continue;
              }
            }
          }
          
          const estadoPermiso = dataPermisosPreEscritura[i][COL_PERM.ESTADO];
          
          if (fechaInicioRegistroNormalizada === fechaInicioNormalizada && estadoPermiso !== "Anulado") {
            Logger.log(`⚠️ RACE CONDITION DETECTADA Y BLOQUEADA para RUT ${beneficiario.rut} - Fecha inicio: ${fechaInicioNormalizada}`);
            return { 
              success: false, 
              message: "Se detectó otra solicitud en proceso con la misma fecha de inicio. Por favor, recarga la página y verifica tu historial." 
            };
          }
        }
      }
      
      Logger.log(`✅ Segunda validación pasada. Procediendo a escribir...`);
      
      // ========================================
      // ESCRIBIR REGISTRO (PROTEGIDO POR LOCK)
      // ========================================
      sheetPermisos.appendRow(newRow);
      var filaRepLegal = sheetPermisos.getLastRow();
      Logger.log(`✅ Permiso creado exitosamente. ID: ${idUnico}`);
      
      // ========================================
      // ENVIAR NOTIFICACIONES
      // ========================================
      // Correo al socio (solo cuando gestiona por sí mismo)
      if (gestion !== "Dirigente" && beneficiario.correo && beneficiario.correo.includes("@")) {
        const fechaInicioObjEmail = new Date(fechaInicioNormalizada);
        const fechaInicioEmailStr = fechaInicioObjEmail.toLocaleDateString('es-CL', { year: 'numeric', month: 'long', day: 'numeric' });
        
        const mensajeSocio = urlDocFinal !== "Sin documento"
          ? `Hola ${beneficiario.nombre}, se ha registrado tu solicitud de permiso médico y el documento de respaldo ha sido adjuntado exitosamente. No necesitas realizar ninguna acción adicional.`
          : `Hola ${beneficiario.nombre}, se ha registrado tu solicitud de permiso médico.\n<strong>IMPORTANTE:</strong> Debes adjuntar el documento de respaldo en el historial del módulo una vez realizada la atención médica.`;

        const datosSocio = { 
          "ID": idUnico,
          "Trabajador": beneficiario.nombre,
          "RUT": beneficiario.rut,
          "Tipo": tipoPermiso,
          "Fecha Inicio": new Date(fechaInicio + 'T12:00:00').toLocaleDateString('es-CL', { year: 'numeric', month: 'long', day: 'numeric' }),
          "Motivo": motivo,
          "Fecha Solicitud": fechaHoy.toLocaleDateString('es-CL', { year: 'numeric', month: 'long', day: 'numeric' }),
          "Documento": urlDocFinal !== "Sin documento"
            ? '<a href="' + urlDocFinal + '" style="color:#10b981;text-decoration:none;font-weight:600;">📎 Ver Documento Adjunto</a>'
            : "Pendiente de adjuntar desde el historial"
        };

        enviarCorreoEstilizado(
          beneficiario.correo,
          "Solicitud Permiso Médico - Sindicato SLIM n°3",
          "Permiso Médico Solicitado",
          mensajeSocio,
          datosSocio,
          "#10b981"
        );
      }
      
      try {
        enviarCorreoEstilizado(
          CORREO_REPRESENTANTE_LEGAL,
          "Notificacion Permiso Medico - Sindicato SLIM n3",
          "Nueva Solicitud de Permiso Medico",
          "Se ha registrado una solicitud de permiso medico para el trabajador <strong>" + beneficiario.nombre + "</strong>.",
          { 
            "ID": idUnico,
            "Trabajador": beneficiario.nombre,
            "RUT": beneficiario.rut,
            "Tipo": tipoPermiso,
            "Fecha Inicio": fechaInicioNormalizada,
            "Motivo": motivo,
            "Estado": estadoFinal,
            "Fecha Solicitud": fechaHoyCompleta.toLocaleDateString(),
            "Documento": urlDocFinal !== "Sin documento"
              ? '<a href="' + urlDocFinal + '" style="color:#10b981;text-decoration:none;font-weight:600;">Ver Documento Adjunto</a>'
              : "Pendiente"
          },
          "#10b981"
        );
        sheetPermisos.getRange(filaRepLegal, COL_PERM.NOTIFICADO_REP_LEGAL + 1).setValue(true);
      } catch (eRepLegal) {
        Logger.log("Advertencia solicitarPermisoMedico: Fallo envio Rep. Legal fila " + filaRepLegal + " - " + eRepLegal.toString());
        // Queda en false para ser reintentado por el trigger
      }
      
      if (gestion === "Dirigente" && correoDirigente && correoDirigente.includes("@") && correoDirigente !== beneficiario.correo) {
        enviarCorreoEstilizado(
          correoDirigente,
          "Respaldo Gestión Permiso Médico - Sindicato SLIM n°3",
          "Permiso Médico Ingresado",
          `Has ingresado exitosamente un permiso médico para el socio <strong>${beneficiario.nombre}</strong>.`,
          {
            "ID": idUnico,
            "Socio": beneficiario.nombre,
            "Tipo": tipoPermiso,
            "Estado": estadoFinal,
            "Documento": urlDocFinal !== "Sin documento"
              ? '<a href="' + urlDocFinal + '" style="color:#475569;text-decoration:none;font-weight:600;">📎 Ver Documento Adjunto</a>'
              : "Sin documento adjunto"
          },
          "#475569"
        );
      }

      // ✅ NUEVO: Copia al socio cuando el dirigente gestiona en su nombre
      if (gestion === "Dirigente" && beneficiario.correo && beneficiario.correo.includes("@")) {
        const mensajeSocioDirigente = urlDocFinal !== "Sin documento"
          ? `Hola <strong>${beneficiario.nombre}</strong>, un dirigente ha solicitado un permiso médico a tu nombre y el documento de respaldo ha sido adjuntado exitosamente. No necesitas realizar ninguna acción adicional.`
          : `Hola <strong>${beneficiario.nombre}</strong>, un dirigente ha solicitado un permiso médico a tu nombre.\n<strong>IMPORTANTE:</strong> Debes adjuntar el documento de respaldo en el historial del módulo una vez realizada la atención médica.`;

        enviarCorreoEstilizado(
          beneficiario.correo,
          "Permiso Médico Solicitado - Sindicato SLIM n°3",
          "Permiso Médico Ingresado",
          mensajeSocioDirigente,
          {
            "ID": idUnico,
            "Trabajador": beneficiario.nombre,
            "RUT": beneficiario.rut,
            "Tipo": tipoPermiso,
            "Fecha Inicio": new Date(fechaInicio + 'T12:00:00').toLocaleDateString('es-CL', { year: 'numeric', month: 'long', day: 'numeric' }),
            "Motivo": motivo,
            "Dirigente": nomDirigente,
            "Estado": estadoFinal,
            "Fecha Solicitud": fechaHoy.toLocaleDateString('es-CL', { year: 'numeric', month: 'long', day: 'numeric' }),
            "Documento": urlDocFinal !== "Sin documento"
              ? '<a href="' + urlDocFinal + '" style="color:#10b981;text-decoration:none;font-weight:600;">📎 Ver Documento Adjunto</a>'
              : "Pendiente de adjuntar desde el historial"
          },
          "#10b981"
        );
      }
      
      return {
        success: true,
        message: urlDocFinal !== "Sin documento"
          ? "Permiso médico registrado con documento adjunto exitosamente."
          : "Permiso médico solicitado. No olvides adjuntar el documento de respaldo desde el historial."
      };
      
    } catch (e) {
      Logger.log("❌ Error en solicitarPermisoMedico: " + e.toString());
      return { success: false, message: "Error: " + e.toString() };
    } finally {
      lock.releaseLock();
    }
  } else {
    return { success: false, message: "Servidor ocupado. Intenta nuevamente." };
  }
}

// ==========================================
// FUNCIÓN ADJUNTAR DOCUMENTO PERMISO - VERSIÓN MEJORADA
// Reemplazar la función existente completamente
// ==========================================

function adjuntarDocumentoPermiso(idPermiso, archivoData) {
  var CARPETA_ID = CONFIG.CARPETAS.PERMISOS_MEDICOS;
  var CORREO_REPRESENTANTE_LEGAL = CONFIG.CORREOS.REPRESENTANTE_LEGAL;
  
  var lock = LockService.getScriptLock();
  if (lock.tryLock(30000)) {
    try {
      var sheetPermisos = getSheet('PERMISOS_MEDICOS', 'PERMISOS_MEDICOS');
      var data = sheetPermisos.getDataRange().getValues();
      var COL = CONFIG.COLUMNAS.PERMISOS_MEDICOS;
      
      var rowIndex = -1;
      var beneficiario = null;
      var tipoPermiso = "";
      var gestionTipo = "";
      var correoGestor = "";
      
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][COL.ID]) === String(idPermiso)) {
          rowIndex = i + 1;
          beneficiario = {
            nombre: data[i][COL.NOMBRE],
            correo: data[i][COL.CORREO],
            rut: data[i][COL.RUT]
          };
          tipoPermiso = data[i][COL.TIPO_PERMISO];
          gestionTipo = data[i][COL.GESTION];
          correoGestor = data[i][COL.CORREO_DIRIGENTE];
          break;
        }
      }
      
      if (rowIndex === -1) return { success: false, message: "Permiso no encontrado." };
      
      // ========== VALIDAR CORREOS ==========
      var esGestionDirigente = gestionTipo === "Dirigente" && esCorreoValido(correoGestor);
      
      var correosParaPermisos = [];
      var alertas = [];
      
      // Correo del beneficiario
      if (esCorreoValido(beneficiario.correo)) {
        correosParaPermisos.push({
          correo: beneficiario.correo.trim().toLowerCase(),
          tipo: 'beneficiario',
          nombre: beneficiario.nombre
        });
      } else {
        alertas.push("El socio " + beneficiario.nombre + " no tiene correo válido. No podrá acceder al documento.");
      }
      
      // Correo del gestor
      if (esGestionDirigente && correoGestor !== beneficiario.correo) {
        correosParaPermisos.push({
          correo: correoGestor.trim().toLowerCase(),
          tipo: 'gestor',
          nombre: 'Dirigente'
        });
      }
      
      // ========== SUBIR ARCHIVO ==========
      var nombreArchivo = "PERMISO-" + idPermiso + "-" + cleanRut(beneficiario.rut);
      
      var resultadoSubida = subirArchivoConPermisos(
        archivoData,
        CARPETA_ID,
        nombreArchivo,
        correosParaPermisos,
        [CORREO_REPRESENTANTE_LEGAL]
      );
      
      if (!resultadoSubida.success) {
        return { success: false, message: resultadoSubida.mensajeError };
      }
      
      // ========== ACTUALIZAR REGISTRO ==========
      var fechaSubida = new Date();
      var nuevoEstado = "Documento Adjuntado";
      
      sheetPermisos.getRange(rowIndex, COL.URL_DOCUMENTO + 1).setValue(resultadoSubida.url);
      sheetPermisos.getRange(rowIndex, COL.ESTADO + 1).setValue(nuevoEstado);
      sheetPermisos.getRange(rowIndex, COL.FECHA_SUBIDA + 1).setValue(fechaSubida);
      sheetPermisos.getRange(rowIndex, COL.NOTIFICADO_REP_LEGAL + 1).setValue(false);
      
      // ========== ENVIAR CORREOS ==========
      if (esCorreoValido(beneficiario.correo)) {
        enviarCorreoEstilizado(
          beneficiario.correo,
          "Documento Adjuntado - Sindicato SLIM n°3",
          "Documento de Permiso Médico Adjuntado",
          "Hola " + beneficiario.nombre + ", tu documento de respaldo ha sido adjuntado exitosamente.",
          {
            "ID": idPermiso,
            "Tipo Permiso": tipoPermiso,
            "Estado": nuevoEstado,
            "Documento": '<a href="' + resultadoSubida.url + '" style="color: #10b981; text-decoration: none; font-weight: 600;">📎 Ver Documento</a>'  // ✅ CORRECCIÓN
          },
          "#10b981"
        );
      }
      
      try {
        enviarCorreoEstilizado(
          CORREO_REPRESENTANTE_LEGAL,
          "Documento Permiso Medico Adjuntado - Sindicato SLIM n3",
          "Documento de Permiso Medico Disponible",
          "El trabajador <strong>" + beneficiario.nombre + "</strong> ha adjuntado el documento de respaldo para su permiso medico.",
          {
            "ID": idPermiso,
            "Trabajador": beneficiario.nombre,
            "RUT": beneficiario.rut,
            "Tipo Permiso": tipoPermiso,
            "Documento": '<a href="' + resultadoSubida.url + '" style="color: #10b981; font-weight: bold;">Disponible para revision</a>',
            "Fecha Adjunto": fechaSubida.toLocaleDateString()
          },
          "#475569"
        );
        sheetPermisos.getRange(rowIndex, COL.NOTIFICADO_REP_LEGAL + 1).setValue(true);
      } catch (eRepLegal) {
        Logger.log("Advertencia adjuntarDocumentoPermiso: Fallo envio Rep. Legal fila " + rowIndex + " - " + eRepLegal.toString());
        // Queda en false para ser reintentado por el trigger
      }
      
      // ========== PREPARAR RESPUESTA ==========
      var respuesta = {
        success: true,
        message: "Documento adjuntado y notificaciones enviadas."
      };
      
      // Agregar alertas si hay problemas
      if (alertas.length > 0 || (resultadoSubida.permisosError && resultadoSubida.permisosError.length > 0)) {
        respuesta.mostrarAlerta = true;
        respuesta.tipoAlerta = 'warning';
        
        var todosDetalles = alertas.slice(); // Copia del array
        if (resultadoSubida.permisosError) {
          resultadoSubida.permisosError.forEach(function(err) {
            todosDetalles.push("No se pudo dar acceso a " + err.nombre);
          });
        }
        
        respuesta.mensajeAlerta = todosDetalles.join('\n\n');
      }
      
      return respuesta;
      
    } catch (e) {
      return { success: false, message: "Error: " + e.toString() };
    } finally {
      lock.releaseLock();
    }
  } else {
    return { success: false, message: "Servidor ocupado." };
  }
}

function obtenerHistorialPermisosMedicos(rutInput) {
  try {
    const sheet = getSheet('PERMISOS_MEDICOS', 'PERMISOS_MEDICOS');
    const COL = CONFIG.COLUMNAS.PERMISOS_MEDICOS;
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, registros: [] };

    // ⭐ CORRECCIÓN: Usar getValues() para obtener objetos Date reales
    var lastCol = sheet.getLastColumn();
    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    
    const rutLimpio = cleanRut(rutInput);
    const registros = [];
    
    for (let i = 0; i < data.length; i++) { // ⭐ CAMBIO: Empezar en 0
      const row = data[i];
      if (cleanRut(row[COL.RUT]) === rutLimpio) {
        registros.push({
          id: row[COL.ID],
          fecha: formatearFechaConHora(row[COL.FECHA_SOLICITUD]),
          tipoPermiso: row[COL.TIPO_PERMISO],
          fechaInicio: formatearFechaSinHora(row[COL.FECHA_INICIO]),
          motivo: row[COL.MOTIVO_DETALLE],
          urlDocumento: row[COL.URL_DOCUMENTO],
          estado: row[COL.ESTADO],
          gestion: row[COL.GESTION],
          nomDirigente: row[COL.NOMBRE_DIRIGENTE]
        });
      }
    }
    
    registros.reverse();
    return { success: true, registros: registros };
    
  } catch (e) {
    Logger.log("❌ Error en obtenerHistorialPermisosMedicos: " + e.toString());
    return { success: false, message: "Error: " + e.toString() };
  }
}

function eliminarPermisoMedico(idPermiso) {
  const CORREO_REPRESENTANTE_LEGAL = CONFIG.CORREOS.REPRESENTANTE_LEGAL;
  
  var lock = LockService.getScriptLock();
  if (lock.tryLock(30000)) { // ✅ Aumentado a 30 segundos para alta concurrencia
    try {
      const sheet = getSheet('PERMISOS_MEDICOS', 'PERMISOS_MEDICOS');
      const data = sheet.getDataRange().getDisplayValues();
      const COL = CONFIG.COLUMNAS.PERMISOS_MEDICOS;
      
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][COL.ID]) === String(idPermiso)) {
          const estado = String(data[i][COL.ESTADO]);
          
          if (estado !== "Solicitado") {
            return { success: false, message: "Solo se pueden anular permisos en estado 'Solicitado'." };
          }
          
          const beneficiario = {
            nombre: data[i][COL.NOMBRE],
            correo: data[i][COL.CORREO],
            rut: data[i][COL.RUT]
          };
          const tipoPermiso = data[i][COL.TIPO_PERMISO];
          const fechaInicio = data[i][COL.FECHA_INICIO];
          
          if (beneficiario.correo && beneficiario.correo.includes("@")) {
            enviarCorreoEstilizado(
              beneficiario.correo,
              "Permiso Médico Anulado - Sindicato SLIM n°3",
              "Solicitud de Permiso Anulada",
              `Hola ${beneficiario.nombre}, tu solicitud de permiso médico ha sido anulada. No se hará uso de este permiso.`,
              { 
                "ID": idPermiso,
                "Tipo Permiso": tipoPermiso,
                "Fecha Inicio": fechaInicio,
                "Estado": "Anulado",
                "Acción": "Solicitud eliminada del sistema"
              },
              "#ef4444"
            );
          }
          
          enviarCorreoEstilizado(
            CORREO_REPRESENTANTE_LEGAL,
            "Permiso Médico Anulado - Sindicato SLIM n°3",
            "Solicitud de Permiso Anulada",
            `La solicitud de permiso médico del trabajador <strong>${beneficiario.nombre}</strong> ha sido anulada. No se hará uso de este permiso.`,
            { 
              "ID": idPermiso,
              "Trabajador": beneficiario.nombre,
              "RUT": beneficiario.rut,
              "Tipo Permiso": tipoPermiso,
              "Fecha Inicio": fechaInicio,
              "Estado": "Anulado por el usuario"
            },
            "#475569"
          );
          
          sheet.deleteRow(i + 1);
          return { success: true, message: "Permiso anulado y notificaciones enviadas." };
        }
      }
      
      return { success: false, message: "Permiso no encontrado." };
      
    } catch (e) {
      return { success: false, message: "Error: " + e.toString() };
    } finally {
      lock.releaseLock();
    }
  } else {
    return { success: false, message: "Servidor ocupado." };
  }
}

// REEMPLAZAR CON ESTO:
// ==========================================
// MÓDULO: REGISTRO ASISTENCIA
// ==========================================

/**
 * Registra la asistencia de un socio vía QR de punto de control.
 * VERSIÓN OPTIMIZADA: Usuario con caché + lock reducido + correo delegado al trigger 20:00
 * BD_ASISTENCIA columnas: FECHA_HORA(A), RUT(B), NOMBRE(C), ASAMBLEA(D), TIPO_ASISTENCIA(E), GESTION(F), CODIGO_TEMP(G), NOTIF_CORREO(H)
 */
function registrarAsistencia(rutInput, nombreControl) {

  const rutLimpio = cleanRut(rutInput);
  if (!rutLimpio) return { success: false, message: "RUT inválido." };

  // ── OPTIMIZACIÓN 1: Búsqueda de usuario CON CACHÉ, FUERA del lock ──────────
  // obtenerUsuarioPorRut usa CacheService (TTL 10 min), evita leer BD_SLIMAPP
  // en cada llamada simultánea cuando hay mucha concurrencia.
  const usuario = obtenerUsuarioPorRut(rutInput);
  if (!usuario.encontrado) {
    return { success: false, message: "RUT no encontrado en el sistema." };
  }

  // ── OPTIMIZACIÓN 2: Sección crítica reducida al mínimo indispensable ────────
  // El lock protege SOLO la verificación de duplicado + escritura en BD_ASISTENCIA.
  // Sin correos adentro → tiempo de lock baja de ~7-10 seg a ~1.5-2 seg.
  // Capacidad estimada: de 3-4 usuarios simultáneos a ~15-20.
  var lock = LockService.getScriptLock();
  if (lock.tryLock(30000)) {
    try {
      const ssAsistencia = getSpreadsheet('ASISTENCIA');
      let sheetAsistencia = ssAsistencia.getSheetByName(CONFIG.HOJAS.ASISTENCIA);

      if (!sheetAsistencia) {
        sheetAsistencia = ssAsistencia.insertSheet(CONFIG.HOJAS.ASISTENCIA);
        sheetAsistencia.appendRow(["FECHA_HORA", "RUT", "NOMBRE", "ASAMBLEA", "TIPO_ASISTENCIA", "GESTION", "CODIGO_TEMP", "NOTIF_CORREO"]);
      }

      // Verificar duplicado en esta asamblea
      const dataAsistencia = sheetAsistencia.getDataRange().getDisplayValues();
      for (let i = 1; i < dataAsistencia.length; i++) {
        const row = dataAsistencia[i];
        if (cleanRut(row[1]) === rutLimpio && row[3] === nombreControl) {
          return {
            success: false,
            yaRegistrado: true,
            message: "Ya registraste tu asistencia en este punto de control."
          };
        }
      }

      // Registrar fila — columna H vacía intencionalmente
      const fechaStr = Utilities.formatDate(new Date(), 'America/Santiago', 'dd/MM/yyyy HH:mm:ss');
      sheetAsistencia.appendRow([
        fechaStr,          // A: FECHA_HORA
        usuario.rut,       // B: RUT
        usuario.nombre,    // C: NOMBRE
        nombreControl,     // D: ASAMBLEA
        "Asistencia QR",  // E: TIPO_ASISTENCIA
        "Sistema",         // F: GESTION
        "",                // G: CODIGO_TEMP
        ""                 // H: NOTIF_CORREO — trigger 20:00 lo completa
      ]);

      // ── OPTIMIZACIÓN 3: Correo delegado al trigger de las 20:00 ─────────────
      // verificarNotificacionesAsistencia() revisa filas con H="" y envía el correo.
      // Así el lock se libera inmediatamente sin esperar el envío del email.
      return {
        success: true,
        nombre: usuario.nombre,
        rut: usuario.rut,
        fecha: fechaStr,
        asamblea: nombreControl,
        correoEnviado: false,
        mensajeCorreo: usuario.correo && usuario.correo.includes("@")
          ? "Recibirás una confirmación en tu correo a más tardar esta noche."
          : "No tienes correo registrado. Puedes ver tu historial en el módulo 'Registro Asistencia'."
      };

    } catch (e) {
      Logger.log("❌ Error en registrarAsistencia: " + e.toString());
      return { success: false, message: "Error del servidor: " + e.toString() };
    } finally {
      lock.releaseLock();
    }
  } else {
    return { success: false, message: "Sistema ocupado, intenta nuevamente en unos segundos." };
  }
}

/**
 * Obtiene el historial de asistencias del usuario.
 * Llamada desde Index.html (módulo Registro Asistencia)
 * BD_ASISTENCIA columnas: FECHA_HORA(0), RUT(1), NOMBRE(2), ASAMBLEA(3), TIPO_ASISTENCIA(4), GESTION(5), CODIGO_TEMP(6)
 */
function obtenerHistorialAsistencia(rutInput) {
  try {
    const rutLimpio = cleanRut(rutInput);
    if (!rutLimpio) return { success: false, message: "RUT inválido." };

    const ssAsistencia = getSpreadsheet('ASISTENCIA');
    const sheet = ssAsistencia.getSheetByName(CONFIG.HOJAS.ASISTENCIA);

    if (!sheet) return { success: true, registros: [] };

    const data = sheet.getDataRange().getDisplayValues();
    const registros = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (cleanRut(row[1]) === rutLimpio) {  // Columna B = RUT (índice 1)
        registros.push({
          fecha: row[0] || "",               // A: FECHA_HORA
          asamblea: row[3] || "Asamblea",    // D: ASAMBLEA
          tipo: row[4] || "Asistencia QR",   // E: TIPO_ASISTENCIA
          gestion: row[5] || "Sistema",      // F: GESTION
          dirigente: ""
        });
      }
    }

    registros.reverse();
    return { success: true, registros: registros };

  } catch (e) {
    Logger.log("❌ Error en obtenerHistorialAsistencia: " + e.toString());
    return { success: false, message: "Error: " + e.toString() };
  }
}

/**
 * Trigger diario a las 20:00 hrs.
 * Verifica registros en BD_ASISTENCIA sin notificación enviada (columna H vacía),
 * busca el correo del socio en BD_SLIMAPP y envía la notificación correspondiente.
 */
function verificarNotificacionesAsistencia() {
  try {
    Logger.log('🔔 Iniciando verificación de notificaciones pendientes de asistencia...');

    const ssAsistencia = getSpreadsheet('ASISTENCIA');
    const sheet = ssAsistencia.getSheetByName(CONFIG.HOJAS.ASISTENCIA);

    if (!sheet) {
      Logger.log('⚠️ Hoja BD_ASISTENCIA no encontrada.');
      return;
    }

    const data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) {
      Logger.log('ℹ️ No hay registros en BD_ASISTENCIA.');
      return;
    }

    // Construir mapa RUT → correo desde BD_SLIMAPP
    const sheetUsuarios = getSheet('USUARIOS', 'USUARIOS');
    const dataUsuarios = sheetUsuarios.getDataRange().getDisplayValues();
    const COL_U = CONFIG.COLUMNAS.USUARIOS;
    const mapaCorreos = {};
    for (let i = 1; i < dataUsuarios.length; i++) {
      const rutU = cleanRut(dataUsuarios[i][COL_U.RUT]);
      if (rutU) mapaCorreos[rutU] = dataUsuarios[i][COL_U.CORREO] || "";
    }

    let enviados = 0;
    let sinCorreo = 0;
    let omitidos = 0;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const notifEstado = String(row[7] || "").trim(); // H: NOTIF_CORREO

      // Si la columna H ya tiene valor, esta fila ya fue procesada
      if (notifEstado !== "") {
        omitidos++;
        continue;
      }

      const rutFila     = cleanRut(row[1]);        // B: RUT
      const nombreFila  = row[2] || "Socio";        // C: NOMBRE
      const fechaFila   = row[0] || "";             // A: FECHA_HORA
      const asambleaFila = row[3] || "";            // D: ASAMBLEA
      const tipoFila    = row[4] || "Asistencia";   // E: TIPO_ASISTENCIA

      const correoUsuario = mapaCorreos[rutFila] || "";

      if (correoUsuario && correoUsuario.includes("@")) {
        try {
          enviarCorreoEstilizado(
            correoUsuario,
            "Registro de Asistencia - Sindicato SLIM n°3",
            "Asistencia Registrada",
            "Tu asistencia ha sido registrada en el sistema del sindicato.",
            {
              "Nombre":     nombreFila,
              "Asamblea":   asambleaFila,
              "Tipo":       tipoFila,
              "Fecha/Hora": fechaFila
            },
            "#10b981"
          );
          sheet.getRange(i + 1, 8).setValue("ENVIADO");
          enviados++;
          Logger.log("✅ Notificación enviada a " + correoUsuario + " (RUT: " + rutFila + ")");
        } catch (emailErr) {
          Logger.log("⚠️ Error enviando correo a " + correoUsuario + ": " + emailErr.toString());
          // No se marca para reintento en la próxima ejecución del trigger
        }
      } else {
        sheet.getRange(i + 1, 8).setValue("SIN CORREO");
        sinCorreo++;
        Logger.log("ℹ️ Sin correo registrado para RUT: " + rutFila);
      }
    }

    Logger.log("📊 Resultado: " + enviados + " enviados, " + sinCorreo + " sin correo, " + omitidos + " ya procesados.");

  } catch (e) {
    Logger.log("❌ Error en verificarNotificacionesAsistencia: " + e.toString());
  }
}

// ==========================================
// GESTIÓN SOCIOS - PARA DIRIGENTES
// ==========================================

function obtenerGestionesDirigente(rutDirigente) {
  try {
    const rutLimpio = cleanRut(rutDirigente);
    
    const resultado = {
      prestamos: [],
      justificaciones: [],
      apelaciones: [],
      permisosMedicos: []
    };
    
    // 1. PRÉSTAMOS
    const sheetPrestamos = getSheet('PRESTAMOS', 'PRESTAMOS');
    const dataPrestamos = sheetPrestamos.getDataRange().getDisplayValues();
    const COL_PRES = CONFIG.COLUMNAS.PRESTAMOS;
    
    for (let i = 1; i < dataPrestamos.length; i++) {
      const row = dataPrestamos[i];
      const tipo = row[COL_PRES.TIPO] || "Préstamo";
      const cuotas = row[COL_PRES.CUOTAS] || "S/D";
      const medio = row[COL_PRES.MEDIO_PAGO] || "S/D";
      const monto = row[COL_PRES.MONTO] || "$0";
      const observacion = row[COL_PRES.OBSERVACION] || "";
      
      // --- LIMPIEZA FECHA TÉRMINO ---
      let fechaTerminoStr = "S/D";
      const ftRaw = row[COL_PRES.FECHA_TERMINO];
      if (ftRaw) {
         try {
            const d = new Date(ftRaw);
            if (!isNaN(d.getTime())) {
               fechaTerminoStr = Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy");
            } else {
               fechaTerminoStr = String(ftRaw).split(' ')[0];
            }
         } catch(e) { fechaTerminoStr = String(ftRaw).split(' ')[0]; }
      }

      resultado.prestamos.push({
        id: row[COL_PRES.ID],
        fecha: row[COL_PRES.FECHA],
        rutSocio: row[COL_PRES.RUT],
        nombreSocio: row[COL_PRES.NOMBRE],
        tipo: tipo,
        monto: monto,
        cuotas: cuotas,
        medio: medio,
        estado: row[COL_PRES.ESTADO],
        observacion: observacion,
        fechaTermino: fechaTerminoStr
      });
    }
    
    // JUSTIFICACIONES
    const sheetJustif = getSheet('JUSTIFICACIONES', 'JUSTIFICACIONES');
    const dataJustif = sheetJustif.getDataRange().getDisplayValues();
    const COL_JUST = CONFIG.COLUMNAS.JUSTIFICACIONES;
    
    for (let i = 1; i < dataJustif.length; i++) {
      const row = dataJustif[i];
      const gestion = row[COL_JUST.GESTION];
      
      if (gestion === "Dirigente") {
        resultado.justificaciones.push({
          id: row[COL_JUST.ID],
          fecha: row[COL_JUST.FECHA],
          rutSocio: row[COL_JUST.RUT],
          nombreSocio: row[COL_JUST.NOMBRE],
          tipo: row[COL_JUST.MOTIVO],
          motivo: row[COL_JUST.ARGUMENTO],
          url: row[COL_JUST.RESPALDO],
          estado: row[COL_JUST.ESTADO],
          obs: row[COL_JUST.OBSERVACION],
          asamblea: row[COL_JUST.ASAMBLEA]
        });
      }
    }
    
    // APELACIONES
    const sheetApel = getSheet('APELACIONES', 'APELACIONES');
    const dataApel = sheetApel.getDataRange().getDisplayValues();
    const COL_APEL = CONFIG.COLUMNAS.APELACIONES;
    
    for (let i = 1; i < dataApel.length; i++) {
      const row = dataApel[i];
      const gestion = row[COL_APEL.GESTION];
      
      if (gestion === "Dirigente") {
        resultado.apelaciones.push({
          id: row[COL_APEL.ID],
          fecha: row[COL_APEL.FECHA_SOLICITUD],
          rutSocio: row[COL_APEL.RUT],
          nombreSocio: row[COL_APEL.NOMBRE],
          mesApelacion: row[COL_APEL.MES_APELACION],
          tipoMotivo: row[COL_APEL.TIPO_MOTIVO],
          detalleMotivo: row[COL_APEL.DETALLE_MOTIVO],
          urlComprobante: row[COL_APEL.URL_COMPROBANTE],
          urlLiquidacion: row[COL_APEL.URL_LIQUIDACION],
          estado: row[COL_APEL.ESTADO],
          obs: row[COL_APEL.OBSERVACION],
          urlComprobanteDevolucion: row[COL_APEL.URL_COMPROBANTE_DEVOLUCION] || ""
        });
      }
    }
    
    // PERMISOS MÉDICOS
    const sheetPermisos = getSheet('PERMISOS_MEDICOS', 'PERMISOS_MEDICOS');
    const dataPermisos = sheetPermisos.getDataRange().getDisplayValues();
    const COL_PERM = CONFIG.COLUMNAS.PERMISOS_MEDICOS;
    
    for (let i = 1; i < dataPermisos.length; i++) {
      const row = dataPermisos[i];
      const gestion = row[COL_PERM.GESTION];
      
      if (gestion === "Dirigente") {
        resultado.permisosMedicos.push({
          id: row[COL_PERM.ID],
          fecha: row[COL_PERM.FECHA_SOLICITUD],
          rutSocio: row[COL_PERM.RUT],
          nombreSocio: row[COL_PERM.NOMBRE],
          tipoPermiso: row[COL_PERM.TIPO_PERMISO],
          fechaInicio: row[COL_PERM.FECHA_INICIO],
          motivo: row[COL_PERM.MOTIVO_DETALLE],
          urlDocumento: row[COL_PERM.URL_DOCUMENTO],
          estado: row[COL_PERM.ESTADO]
        });
      }
    }
    
    return { success: true, datos: resultado };
    
  } catch (e) {
    return { success: false, message: "Error: " + e.toString() };
  }
}

// ==========================================
// PANEL ADMINISTRADOR
// ==========================================

function generarInformeAdministrador() {
  try {
    const sheetPrestamos = getSheet('PRESTAMOS', 'PRESTAMOS');
    const data = sheetPrestamos.getDataRange().getValues();
    const COL = CONFIG.COLUMNAS.PRESTAMOS;
    
    const prestamosSolicitados = [];
    const filasActualizar = [];
    
    for (let i = 1; i < data.length; i++) {
      const estado = String(data[i][COL.VALIDACION]);
      
      if (estado === "Solicitado") {
        const obs = data[i][COL.OBSERVACION] || "";
        const partes = obs.split(" - ");
        const tipo = partes[0] || "Préstamo";
        const cuotas = partes[1] || "";
        const medio = partes[2] || "";
        
        prestamosSolicitados.push({
          rut: data[i][COL.RUT],
          nombre: data[i][COL.NOMBRE],
          tipoPrestamo: tipo,
          cuotas: cuotas,
          medioPago: medio
        });
        filasActualizar.push(i + 1);
      }
    }
    
    if (prestamosSolicitados.length === 0) {
      return { success: false, message: "No hay préstamos en estado 'Solicitado' para procesar." };
    }
    
    const ss = getSpreadsheet('PRESTAMOS');
    let sheetInforme = ss.getSheetByName("INFORME_PRESTAMOS_TEMP");
    if (sheetInforme) {
      ss.deleteSheet(sheetInforme);
    }
    sheetInforme = ss.insertSheet("INFORME_PRESTAMOS_TEMP");
    
    const headers = ["RUT", "NOMBRE", "TIPO PRÉSTAMO", "CUOTAS", "MEDIO PAGO"];
    sheetInforme.appendRow(headers);
    
    prestamosSolicitados.forEach(prestamo => {
      sheetInforme.appendRow([
        prestamo.rut,
        prestamo.nombre,
        prestamo.tipoPrestamo,
        prestamo.cuotas,
        prestamo.medioPago
      ]);
    });
    
    const lastRow = sheetInforme.getLastRow();
    const lastCol = sheetInforme.getLastColumn();
    
    sheetInforme.getRange(1, 1, 1, lastCol)
      .setFontWeight("bold")
      .setBackground("#4c1d95")
      .setFontColor("#ffffff");
    
    sheetInforme.setFrozenRows(1);
    sheetInforme.autoResizeColumns(1, lastCol);
    
    const url = "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/export?format=xlsx&gid=" + sheetInforme.getSheetId();
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + token
      }
    });
    
    const blob = response.getBlob();
    blob.setName(`Informe_Prestamos_Solicitados_${new Date().toLocaleDateString('es-CL').replace(/\//g, '-')}.xlsx`);
    
    const sheetUsers = getSheet('USUARIOS', 'USUARIOS');
    const dataUsers = sheetUsers.getDataRange().getDisplayValues();
    const COL_USER = CONFIG.COLUMNAS.USUARIOS;
    let correoAdmin = "admin@sindicato.com";
    
    for (let i = 1; i < dataUsers.length; i++) {
      const rol = String(dataUsers[i][COL_USER.ROL]).toUpperCase();
      if (rol === "ADMIN") {
        correoAdmin = dataUsers[i][COL_USER.CORREO];
        break;
      }
    }
    
    MailApp.sendEmail({
      to: correoAdmin,
      subject: "Informe de Préstamos Solicitados - Sindicato SLIM n°3",
      htmlBody: `
        <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
          <div style="background: linear-gradient(135deg, #4c1d95 0%, #5b21b6 100%); color: white; padding: 30px; text-align: center; border-radius: 10px 10px 0 0;">
            <h1 style="margin: 0; font-size: 24px;">Informe de Préstamos Solicitados</h1>
          </div>
          <div style="background: #f8f9fa; padding: 30px; border-radius: 0 0 10px 10px;">
            <p style="color: #1e293b; font-size: 16px; line-height: 1.6;">
              Se adjunta el informe de préstamos en estado <strong>"Solicitado"</strong> generado el <strong>${new Date().toLocaleDateString('es-CL')}</strong>.
            </p>
            <p style="color: #64748b; font-size: 14px; margin-top: 20px;">
              Total de préstamos procesados: <strong>${prestamosSolicitados.length}</strong>
            </p>
            <p style="color: #dc2626; font-size: 14px; font-weight: bold; margin-top: 15px;">
              ⚠️ Estos préstamos han sido cambiados automáticamente al estado "Enviado".
            </p>
            <hr style="border: none; border-top: 1px solid #e2e8f0; margin: 20px 0;">
            <p style="color: #94a3b8; font-size: 12px; text-align: center;">
              Sindicato SLIM n°3 - Sistema de Gestión
            </p>
          </div>
        </div>
      `,
      attachments: [blob]
    });
    
    filasActualizar.forEach(fila => {
      sheetPrestamos.getRange(fila, COL.VALIDACION + 1).setValue("Enviado");
    });
    
    ss.deleteSheet(sheetInforme);
    
    return { 
      success: true, 
      message: `Informe generado y enviado. ${prestamosSolicitados.length} préstamo(s) cambiado(s) a "Enviado".` 
    };
    
  } catch (e) {
    return { success: false, message: "Error al generar informe: " + e.toString() };
  }
}

// ==========================================
// FUNCIONES AUXILIARES
// ==========================================

// AGREGAR al inicio de Code.gs, después de las funciones auxiliares

/**
 * Verifica si un usuario tiene un rol específico
 * @param {string} rut - RUT del usuario
 * @param {Array} rolesPermitidos - Array de roles permitidos ['ADMIN', 'DIRIGENTE']
 * @returns {Object} {autorizado: boolean, mensaje: string, rol: string}
 */
function verificarRolUsuario(rut, rolesPermitidos) {
  try {
    const usuario = obtenerUsuarioPorRut(rut);
    
    if (!usuario.encontrado) {
      return {
        autorizado: false,
        mensaje: "Usuario no encontrado",
        rol: ""
      };
    }
    
    const rolUsuario = String(usuario.rol || "SOCIO").trim().toUpperCase();
    const tienePermiso = rolesPermitidos.some(function(rol) {
      return rol.toUpperCase() === rolUsuario;
    });
    
    if (!tienePermiso) {
      Logger.log('⚠️ INTENTO DE ACCESO NO AUTORIZADO:');
      Logger.log('   RUT: ' + rut);
      Logger.log('   Rol actual: ' + rolUsuario);
      Logger.log('   Roles requeridos: ' + rolesPermitidos.join(', '));
      
      return {
        autorizado: false,
        mensaje: "No tienes permisos para realizar esta acción",
        rol: rolUsuario
      };
    }
    
    return {
      autorizado: true,
      mensaje: "Acceso autorizado",
      rol: rolUsuario
    };
    
  } catch (e) {
    Logger.log('❌ Error verificando rol: ' + e.toString());
    return {
      autorizado: false,
      mensaje: "Error de validación",
      rol: ""
    };
  }
}

function cleanRut(rut) {
  if (!rut) return "";
  return String(rut).replace(/\./g, '').replace(/-/g, '').toUpperCase().trim();
}

function enviarCorreoEstilizado(destinatario, asunto, titulo, mensaje, detalles, colorTema) {
  try {
    if (!destinatario || !destinatario.includes("@")) {
      console.log("Correo inválido: " + destinatario);
      return;
    }
    
    // Convertir detalles a tabla HTML
    let detallesHtml = "";
    if (detalles && typeof detalles === "object") {
      detallesHtml = "<table style='width: 100%; border-collapse: separate; border-spacing: 0; margin-top: 20px; border: 1px solid #e2e8f0; border-radius: 8px; overflow: hidden;'>";
      
      let isEven = false;
      for (let key in detalles) {
        let valor = detalles[key];
        
        // LÓGICA S/D
        if (valor === null || valor === undefined || valor === "") {
          valor = "<span style='color: #94a3b8; font-style: italic;'>S/D</span>";
        }
        
        const bgRow = isEven ? "#f8fafc" : "#ffffff";
        
        detallesHtml += `
          <tr style="background-color: ${bgRow};">
            <td style='padding: 12px 15px; border-bottom: 1px solid #e2e8f0; color: #64748b; font-weight: 600; font-size: 13px; width: 35%; vertical-align: top; text-transform: uppercase; letter-spacing: 0.05em;'>${key}</td>
            <td style='padding: 12px 15px; border-bottom: 1px solid #e2e8f0; color: #1e293b; font-weight: 500; font-size: 14px; vertical-align: top;'>${valor}</td>
          </tr>
        `;
        isEven = !isEven;
      }
      detallesHtml += "</table>";
    }
    
    // Generamos un ID único para evitar que Gmail agrupe y oculte el footer
    const uniqueId = Utilities.getUuid().slice(0, 8);
    
    const htmlBody = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
      </head>
      <body style="margin: 0; padding: 0; font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; background-color: #f1f5f9;">
        <div style="max-width: 600px; margin: 20px auto; background: white; border-radius: 16px; overflow: hidden; box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1);">
          
          <div style="background: linear-gradient(135deg, ${colorTema} 0%, ${adjustColor(colorTema, -40)} 100%); padding: 40px 30px; text-align: center;">
            <h1 style="margin: 0; color: white; font-size: 24px; font-weight: 800; letter-spacing: -0.5px; text-shadow: 0 2px 4px rgba(0,0,0,0.1);">${titulo}</h1>
            <p style="margin: 10px 0 0 0; color: rgba(255,255,255,0.9); font-size: 14px;">Sindicato SLIM N°3</p>
          </div>
          
          <div style="padding: 40px 30px; background-color: #ffffff;">
            <p style="color: #334155; font-size: 16px; line-height: 1.6; margin: 0 0 25px 0; text-align: left;">
              ${mensaje}
            </p>
            
            ${detallesHtml}
            
            <div style="margin-top: 30px; padding: 15px; background-color: #eff6ff; border-left: 4px solid ${colorTema}; border-radius: 4px;">
              <p style="color: #1e40af; font-size: 12px; line-height: 1.5; margin: 0;">
                <strong>Nota Importante:</strong> Si el campo aparece como "S/D", significa que no hay datos registrados para ese ítem en el momento de la gestión.
              </p>
            </div>
          </div>
          
          <div style="background: #f8fafc; padding: 20px; text-align: center; border-top: 1px solid #e2e8f0;">
            <p style="color: #64748b; font-size: 11px; margin: 0; line-height: 1.4;">
              Este es un mensaje automático. Por favor no respondas a este correo.<br>
              © ${new Date().getFullYear()} Plataforma de Gestión Sindicato SLIM N°3
            </p>
            <p style="color: #cbd5e1; font-size: 9px; margin: 10px 0 0 0;">Ref: ${uniqueId}</p>
          </div>
          
        </div>
      </body>
      </html>
    `;
    
    MailApp.sendEmail({
      to: destinatario,
      subject: asunto,
      htmlBody: htmlBody
    });
    
  } catch (e) {
    console.error("Error enviando correo a " + destinatario + ": " + e.toString());
  }
}

function adjustColor(hexColor, percent) {
  const num = parseInt(hexColor.replace("#", ""), 16);
  const amt = Math.round(2.55 * percent);
  const R = (num >> 16) + amt;
  const G = (num >> 8 & 0x00FF) + amt;
  const B = (num & 0x0000FF) + amt;
  return "#" + (0x1000000 + (R < 255 ? R < 1 ? 0 : R : 255) * 0x10000 +
    (G < 255 ? G < 1 ? 0 : G : 255) * 0x100 +
    (B < 255 ? B < 1 ? 0 : B : 255))
    .toString(16).slice(1);
}

function formatRutServer(rut) {
  // ✅ CORRECCIÓN: Convertir a string primero para evitar error con números
  if (!rut) return "";
  
  // Convertir a string si es número u otro tipo
  let rutString = String(rut).trim();
  
  // Limpiar y formatear
  let value = rutString.replace(/[^0-9kK]/g, '').toUpperCase();
  if (value.length < 2) return value;
  
  const body = value.slice(0, -1);
  const dv = value.slice(-1);
  
  let formattedBody = "";
  for (let i = body.length - 1, j = 0; i >= 0; i--, j++) {
    formattedBody = body.charAt(i) + ((j > 0 && j % 3 === 0) ? "." : "") + formattedBody;
  }
  
  return formattedBody + "-" + dv;
}


// ==========================================
// TRIGGERS Y AUTOMATIZACIONES
// ==========================================

function verificarCambiosJustificaciones() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("JUSTIFICACIONES");
    if (!sheet) return;
    
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const idRegistro = String(row[0]);
      const estadoActual = String(row[8]);
      const estadoNotif = String(row[10]);
      const correo = row[4];
      const nombre = row[3];
      const tipo = row[5];
      const obs = row[9];
      const asamblea = row[11];
      
      if (estadoActual !== estadoNotif) {
        if (correo && correo.includes("@")) {
          let color = "#ea580c";
          let titulo = "Actualización de Justificación";
          
          if (estadoActual.includes("Aceptado")) { 
            color = "#15803d"; 
            titulo = "Justificación Aceptada"; 
          } else if (estadoActual.includes("Rechazado")) { 
            color = "#b91c1c"; 
            titulo = "Justificación Rechazada"; 
          }
          
          enviarCorreoEstilizado(
            correo, 
            titulo + " - Sindicato SLIM n°3", 
            titulo, 
            `Hola ${nombre}, el estado de tu justificación ha cambiado.`, 
            { 
              "ID": idRegistro,
              "Tipo": tipo, 
              "Nuevo Estado": estadoActual, 
              "Observación": obs || "Sin observaciones",
              "Asamblea": asamblea || "Pendiente asignación"
            }, 
            color
          );
        }
        
        sheet.getRange(i + 1, 11).setValue(estadoActual);
      }
    }
  } catch (e) { 
    console.error("Error verificando justificaciones: " + e); 
  }
}

// ==========================================
// SISTEMA QR - VALIDACIÓN Y REGISTRO
// ==========================================

function validarUsuarioQR(rutInput) {
  try {
    const sheet = getSheet('USUARIOS', 'USUARIOS');
    const data = sheet.getDataRange().getDisplayValues();
    const rutLimpio = cleanRut(rutInput);
    const COL = CONFIG.COLUMNAS.USUARIOS;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (cleanRut(row[COL.RUT]) === rutLimpio) {
        const estadoUsuario = String(row[COL.ESTADO]).toUpperCase();
        
        if (estadoUsuario !== "ACTIVO" && estadoUsuario !== "SI" && estadoUsuario !== "TRUE") {
          return { 
            success: false, 
            error: "Usuario desvinculado. Contacta con la directiva." 
          };
        }
        
        return {
          success: true,
          nombre: row[COL.NOMBRE] || "Socio",
          rut: row[COL.RUT]
        };
      }
    }
    
    return { success: false, error: "RUT no encontrado en el sistema." };
    
  } catch (e) {
    return { success: false, error: "Error del servidor: " + e.toString() };
  }
}

// REEMPLAZAR CON ESTO:
function checkinQR(rutInput, nombreAsamblea) {
  var lock = LockService.getScriptLock();
  if (lock.tryLock(30000)) {
    try {
      const rutLimpio = cleanRut(rutInput);
      const sheetUsers = getSheet('USUARIOS', 'USUARIOS');
      const dataUsers = sheetUsers.getDataRange().getDisplayValues();
      const COL = CONFIG.COLUMNAS.USUARIOS;

      let usuario = null;
      for (let i = 1; i < dataUsers.length; i++) {
        if (cleanRut(dataUsers[i][COL.RUT]) === rutLimpio) {
          usuario = {
            rut: dataUsers[i][COL.RUT],
            nombre: dataUsers[i][COL.NOMBRE]
          };
          break;
        }
      }

      if (!usuario) throw new Error("Usuario no encontrado");

      const ssAsistencia = getSpreadsheet('ASISTENCIA');
      let sheetAsistencia = ssAsistencia.getSheetByName(CONFIG.HOJAS.ASISTENCIA);

      if (!sheetAsistencia) {
        sheetAsistencia = ssAsistencia.insertSheet(CONFIG.HOJAS.ASISTENCIA);
        sheetAsistencia.appendRow(["FECHA_HORA", "RUT", "NOMBRE", "ASAMBLEA", "TIPO_ASISTENCIA", "GESTION", "CODIGO_TEMP"]);
      }

      const dataAsistencia = sheetAsistencia.getDataRange().getDisplayValues();
      for (let i = 1; i < dataAsistencia.length; i++) {
        const row = dataAsistencia[i];
        if (cleanRut(row[1]) === rutLimpio && row[3] === nombreAsamblea) {
          throw new Error("Ya registraste tu asistencia en esta asamblea.");
        }
      }

      const fechaHora = new Date();
      const fechaStr = Utilities.formatDate(fechaHora, 'America/Santiago', 'dd/MM/yyyy HH:mm:ss');
      sheetAsistencia.appendRow([
        fechaStr,
        usuario.rut,
        usuario.nombre,
        nombreAsamblea,
        "Asistencia QR",
        "Sistema",
        ""
      ]);

      return {
        success: true,
        nombre: usuario.nombre,
        rut: usuario.rut,
        fecha: fechaStr
      };

    } catch (e) {
      throw new Error(e.message || e.toString());
    } finally {
      lock.releaseLock();
    }
  } else {
    throw new Error("Sistema ocupado, intenta nuevamente.");
  }
}

// ==========================================
// CONFIGURAR TRIGGERS (Ejecutar manualmente UNA VEZ)
// ==========================================

function configurarTriggers() {
  // Eliminar triggers existentes para evitar duplicados
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  // Verificar cambios en justificaciones — cada 8 horas
  ScriptApp.newTrigger('verificarCambiosJustificaciones')
    .timeBased()
    .everyHours(8)
    .create();
  
  // Verificar cambios en apelaciones — cada 8 horas
  ScriptApp.newTrigger('verificarCambiosApelaciones')
    .timeBased()
    .everyHours(8)
    .create();
  
  // Procesar validación de préstamos — diario a las 8 AM
  ScriptApp.newTrigger('procesarValidacionPrestamos')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
  
  // Procesar permisos de comprobantes de devolución — cada 1 hora
  ScriptApp.newTrigger('procesarPermisosComprobantesDevolucion')
    .timeBased()
    .everyHours(1)
    .create();
  
  // Verificar cambios en préstamos — diario a las 8 AM
  ScriptApp.newTrigger('verificarCambiosPrestamos')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
  
  // REEMPLAZAR CON ESTO:
  // Verificar cambios en credenciales — cada domingo a las 8 AM
  ScriptApp.newTrigger('verificarCambiosCredenciales')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(8)
    .create();

  // Verificar notificaciones pendientes de asistencia — diario a las 20 hrs
  ScriptApp.newTrigger('verificarNotificacionesAsistencia')
    .timeBased()
    .everyDays(1)
    .atHour(20)
    .create();

    Logger.log("✅ Triggers configurados exitosamente");
  Logger.log("Total de triggers activos: " + ScriptApp.getProjectTriggers().length);
  
  return {
    success: true,
    message: "Triggers configurados correctamente",
    triggers: [
      "verificarCambiosJustificaciones (cada 8 horas)",
      "verificarCambiosApelaciones (cada 8 horas)",
      "procesarValidacionPrestamos (diario 8 AM)",
      "procesarPermisosComprobantesDevolucion (cada 1 hora)",
      "verificarCambiosPrestamos (diario 8 AM)",
      "verificarCambiosCredenciales (domingos 8 AM)"
    ]
  };
}

    // ==========================================
    // REINTENTO NOTIFICACION REPRESENTANTE LEGAL - PERMISOS MEDICOS
    // Trigger: cada 30 minutos
    // ==========================================
    function reintentarNotificacionRepLegal() {
      var CORREO_REPRESENTANTE_LEGAL = CONFIG.CORREOS.REPRESENTANTE_LEGAL;
      var COL = CONFIG.COLUMNAS.PERMISOS_MEDICOS;

      try {
        var sheetPermisos = getSheet('PERMISOS_MEDICOS', 'PERMISOS_MEDICOS');
        var data = sheetPermisos.getDataRange().getValues();
        var pendientes = 0;
        var exitosos = 0;

        for (var i = 1; i < data.length; i++) {
          var fila = data[i];
          var notificado = fila[COL.NOTIFICADO_REP_LEGAL];
          var estado = String(fila[COL.ESTADO]);

          if (estado === '' || estado === 'Anulado') continue;
          if (notificado === true || String(notificado).toUpperCase() === 'TRUE') continue;

          pendientes++;

          var idPermiso   = String(fila[COL.ID]);
          var nombre      = String(fila[COL.NOMBRE]);
          var rut         = String(fila[COL.RUT]);
          var tipoPermiso = String(fila[COL.TIPO_PERMISO]);
          var urlDoc      = String(fila[COL.URL_DOCUMENTO]);
          var motivo      = String(fila[COL.MOTIVO_DETALLE] !== undefined ? fila[COL.MOTIVO_DETALLE] : fila[COL.MOTIVO]);

          var fechaVal = fila[COL.FECHA_INICIO];
          var fechaInicioStr = (fechaVal instanceof Date)
            ? fechaVal.toLocaleDateString('es-CL', { year: 'numeric', month: 'long', day: 'numeric' })
            : String(fechaVal);

          var tieneDoc = urlDoc && urlDoc !== '' && urlDoc !== 'Sin documento';

          try {
            if (estado === 'Documento Adjuntado') {
              enviarCorreoEstilizado(
                CORREO_REPRESENTANTE_LEGAL,
                "Documento Permiso Medico Adjuntado - Sindicato SLIM n3",
                "Documento de Permiso Medico Disponible",
                "El trabajador <strong>" + nombre + "</strong> ha adjuntado el documento de respaldo para su permiso medico.",
                {
                  "ID": idPermiso,
                  "Trabajador": nombre,
                  "RUT": rut,
                  "Tipo": tipoPermiso,
                  "Fecha Inicio": fechaInicioStr,
                  "Estado": estado,
                  "Documento": tieneDoc
                    ? '<a href="' + urlDoc + '" style="color:#10b981;font-weight:bold;">Ver Documento Adjunto</a>'
                    : "Sin documento"
                },
                "#475569"
              );
            } else {
              enviarCorreoEstilizado(
                CORREO_REPRESENTANTE_LEGAL,
                "Notificacion Permiso Medico - Sindicato SLIM n3",
                "Solicitud de Permiso Medico",
                "Se ha registrado una solicitud de permiso medico para el trabajador <strong>" + nombre + "</strong>.",
                {
                  "ID": idPermiso,
                  "Trabajador": nombre,
                  "RUT": rut,
                  "Tipo": tipoPermiso,
                  "Fecha Inicio": fechaInicioStr,
                  "Motivo": motivo,
                  "Estado": estado,
                  "Documento": tieneDoc
                    ? '<a href="' + urlDoc + '" style="color:#10b981;font-weight:bold;">Ver Documento</a>'
                    : "Pendiente"
                },
                "#10b981"
              );
            }

            sheetPermisos.getRange(i + 1, COL.NOTIFICADO_REP_LEGAL + 1).setValue(true);
            exitosos++;
            Utilities.sleep(600);

          } catch (eEmail) {
            Logger.log("reintentarNotificacionRepLegal - Fila " + (i + 1) + " (" + idPermiso + "): " + eEmail.toString());
          }
        }

        Logger.log("reintentarNotificacionRepLegal: " + pendientes + " pendientes, " + exitosos + " enviados.");

      } catch (e) {
        Logger.log("Error general en reintentarNotificacionRepLegal: " + e.toString());
      }
    }

/**
 * Función para obtener el correo de un usuario por RUT
 */
function obtenerCorreoDeRut(rut) {
  try {
    const sheet = getSheet('USUARIOS', 'USUARIOS');
    const data = sheet.getDataRange().getDisplayValues();
    const rutLimpio = cleanRut(rut);
    const COL = CONFIG.COLUMNAS.USUARIOS;
    
    for (let i = 1; i < data.length; i++) {
      if (cleanRut(data[i][COL.RUT]) === rutLimpio) {
        return data[i][COL.CORREO];
      }
    }
    return "";
  } catch (e) {
    console.error("Error obteniendo correo: " + e);
    return "";
  }
}

/**
 * Función para otorgar permisos a archivos de justificaciones existentes
 * EJECUTAR MANUALMENTE UNA SOLA VEZ
 */
function corregirPermisosJustificacionesExistentes() {
  try {
    const sheet = getSheet('JUSTIFICACIONES', 'JUSTIFICACIONES');
    const data = sheet.getDataRange().getValues();
    const COL = CONFIG.COLUMNAS.JUSTIFICACIONES;
    
    let archivosCorregidos = 0;
    let errores = 0;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const gestion = String(row[COL.GESTION]);
      const urlArchivo = String(row[COL.RESPALDO]);
      const correoDirigente = String(row[COL.CORREO_DIRIGENTE]);
      const correoSocio = obtenerCorreoDeRut(row[COL.RUT]);
      
      if (gestion === "Dirigente" && urlArchivo.includes("drive.google.com") && correoDirigente && correoDirigente.includes("@")) {
        
        try {
          let fileId = "";
          if (urlArchivo.includes("/d/")) {
            fileId = urlArchivo.split("/d/")[1].split("/")[0];
          } else if (urlArchivo.includes("id=")) {
            fileId = urlArchivo.split("id=")[1].split("&")[0];
          }
          
          if (fileId) {
            const file = DriveApp.getFileById(fileId);
            const viewers = file.getViewers();
            const tieneDirigente = viewers.some(viewer => viewer.getEmail() === correoDirigente);
            const tieneSocio = viewers.some(viewer => viewer.getEmail() === correoSocio);
            
            let cambios = [];
            
            if (!tieneDirigente) {
              file.addViewer(correoDirigente);
              cambios.push(`dirigente: ${correoDirigente}`);
            }
            
            if (correoSocio && correoSocio.includes("@") && !tieneSocio) {
              file.addViewer(correoSocio);
              cambios.push(`socio: ${correoSocio}`);
            }
            
            if (cambios.length > 0) {
              archivosCorregidos++;
              Logger.log(`✅ Fila ${i + 1} - Permisos otorgados: ${cambios.join(', ')}`);
            }
          }
        } catch (fileErr) {
          errores++;
          Logger.log(`⚠️ Error en fila ${i + 1}: ${fileErr.toString()}`);
        }
      }
    }
    
    Logger.log(`\n📊 RESUMEN:`);
    Logger.log(`✅ Archivos corregidos: ${archivosCorregidos}`);
    Logger.log(`⚠️ Errores: ${errores}`);
    
    return {
      success: true,
      archivosCorregidos: archivosCorregidos,
      errores: errores
    };
    
  } catch (error) {
    Logger.log('❌ Error: ' + error.message);
    return { success: false, message: error.message };
  }
}

    // ==========================================
    // SISTEMA DE MÉTRICAS Y MONITOREO
    // ==========================================

    /**
     * Registra métricas de uso para análisis de performance
     */
    function registrarMetrica(operacion, duracion, exito) {
      try {
        var properties = PropertiesService.getScriptProperties();
        var fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
        var key = 'metrics_' + fecha + '_' + operacion;
        
        var existing = properties.getProperty(key);
        var metrics = existing ? JSON.parse(existing) : { count: 0, totalDuration: 0, errors: 0 };
        
        metrics.count++;
        metrics.totalDuration += duracion;
        if (!exito) metrics.errors++;
        
        properties.setProperty(key, JSON.stringify(metrics));
      } catch (e) {
        // No hacer nada si falla el logging
        Logger.log('Error logging metrics: ' + e);
      }
    }

    /**
     * Wrapper para medir performance de funciones críticas
     */
    function medirPerformance(nombreFuncion, funcionCallback) {
      var inicio = new Date().getTime();
      var exito = true;
      
      try {
        return funcionCallback();
      } catch (e) {
        exito = false;
        throw e;
      } finally {
        var duracion = new Date().getTime() - inicio;
        registrarMetrica(nombreFuncion, duracion, exito);
      }
    }

    /**
     * Ver métricas del día (ejecutar manualmente desde editor)
     */
    function verMetricasHoy() {
      var properties = PropertiesService.getScriptProperties();
      var fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      var allProps = properties.getProperties();
      
      Logger.log('=== MÉTRICAS ' + fecha + ' ===');
      for (var key in allProps) {
        if (key.startsWith('metrics_' + fecha)) {
          var operacion = key.replace('metrics_' + fecha + '_', '');
          var data = JSON.parse(allProps[key]);
          Logger.log(operacion + ': ' + data.count + ' llamadas, avg: ' + 
                    (data.totalDuration / data.count).toFixed(0) + 'ms, errores: ' + data.errors);
        }
      }
    }

    // ==========================================
    // CONSULTA DE ID CREDENCIAL
    // ==========================================

    /**
     * Consulta el ID Credencial de un usuario por RUT
     * Solo accesible para roles DIRIGENTE y ADMIN
     * @param {string} rutConsultante - RUT del usuario que realiza la consulta
     * @param {string} rutBuscado - RUT del usuario a buscar
     * @returns {Object} {success: boolean, idCredencial: string, nombre: string, ...}
     */
    function consultarIdCredencialBackend(rutConsultante, rutBuscado) {
      try {
        // ==========================================
        // 1. VALIDAR PERMISOS DEL CONSULTANTE
        // ==========================================
        const validacion = verificarRolUsuario(rutConsultante, ['DIRIGENTE', 'ADMIN']);
        
        if (!validacion.autorizado) {
          Logger.log('❌ Acceso denegado a consulta de ID Credencial');
          Logger.log('   RUT consultante: ' + rutConsultante);
          Logger.log('   Rol: ' + validacion.rol);
          return {
            success: false,
            message: 'No tienes permisos para realizar esta consulta.'
          };
        }
        
        Logger.log('✅ Consulta de ID Credencial autorizada');
        Logger.log('   Consultante: ' + rutConsultante + ' (' + validacion.rol + ')');
        Logger.log('   Buscando: ' + rutBuscado);
        
        // ==========================================
        // 2. LIMPIAR Y VALIDAR RUT BUSCADO
        // ==========================================
        const rutLimpio = cleanRut(rutBuscado);
        
        if (!rutLimpio || rutLimpio.length < 7) {
          return {
            success: false,
            message: 'RUT inválido o incompleto.'
          };
        }
        
        // ==========================================
        // 3. BUSCAR USUARIO EN LA BASE DE DATOS
        // ==========================================
        const sheet = getSheet('USUARIOS', 'USUARIOS');
        if (!sheet) {
          Logger.log('❌ No se pudo acceder a la hoja de usuarios');
          return {
            success: false,
            message: 'Error al acceder a la base de datos.'
          };
        }
        
        const COL = CONFIG.COLUMNAS.USUARIOS;
        const lastRow = sheet.getLastRow();
        
        if (lastRow < 2) {
          return {
            success: false,
            message: 'No hay usuarios registrados en el sistema.'
          };
        }
        
        const data = sheet.getRange(2, 1, lastRow - 1, COL.ESTADO_NEG_COLECT + 1).getDisplayValues();

        // Leer las FÓRMULAS de la columna QR para extraer la URL
        const formulasQR = sheet.getRange(2, COL.QR_REGISTRO + 1, lastRow - 1, 1).getFormulas();
        
        // ==========================================
        // 4. BUSCAR COINCIDENCIA
        // ==========================================
        for (let i = 0; i < data.length; i++) {
          if (cleanRut(data[i][COL.RUT]) === rutLimpio) {
            const rolUsuarioBuscado = String(data[i][COL.ROL] || 'SOCIO').trim().toUpperCase();
            
            // ==========================================
            // 4.1 VALIDAR RESTRICCIÓN POR ROL
            // ==========================================
            // Si el consultante es DIRIGENTE, solo puede ver SOCIOS
            if (validacion.rol === 'DIRIGENTE' && rolUsuarioBuscado !== 'SOCIO') {
              Logger.log('⚠️ ACCESO RESTRINGIDO:');
              Logger.log('   Consultante: ' + rutConsultante + ' (DIRIGENTE)');
              Logger.log('   Usuario buscado: ' + data[i][COL.NOMBRE] + ' (' + rolUsuarioBuscado + ')');
              Logger.log('   Motivo: Los dirigentes solo pueden consultar usuarios con rol SOCIO');
              
              return {
                success: false,
                message: 'Acceso restringido: Solo puedes consultar información de usuarios con rol SOCIO.',
                restricted: true
              };
            }
            
            // ==========================================
            // 4.2 USUARIO ENCONTRADO Y AUTORIZADO
            // ==========================================

            // Extraer URL del QR desde la fórmula =IMAGE()
            const formulaQR = formulasQR[i][0]; // Fórmula de la columna QR
            const urlQR = extraerUrlDeImagen(formulaQR);

            const usuario = {
              success: true,
              rut: data[i][COL.RUT],
              nombre: data[i][COL.NOMBRE],
              cargo: data[i][COL.CARGO],
              estado: data[i][COL.ESTADO],
              rol: rolUsuarioBuscado,
              idCredencial: data[i][COL.ID_CREDENCIAL] || 'S/D',
              qrRegistro: urlQR,
              estadoCredencial: obtenerEstadoCredencialPorRut(data[i][COL.RUT])
            };
            
            Logger.log('✅ Usuario encontrado y acceso autorizado:');
            Logger.log('   Consultante: ' + rutConsultante + ' (' + validacion.rol + ')');
            Logger.log('   Usuario: ' + usuario.nombre + ' (' + usuario.rol + ')');
            Logger.log('   ID Credencial: ' + usuario.idCredencial);
            Logger.log('   Fórmula QR: ' + formulaQR);
            Logger.log('   URL QR extraída: ' + urlQR);

            return usuario;
          }
        }
        
        // ==========================================
        // 5. USUARIO NO ENCONTRADO
        // ==========================================
        Logger.log('⚠️ Usuario no encontrado con RUT: ' + rutBuscado);
        return {
          success: false,
          message: 'No se encontró ningún usuario con el RUT ' + formatRutDisplay(rutBuscado) + ' en el sistema.'
        };
        
      } catch (e) {
        Logger.log('❌ ERROR en consultarIdCredencialBackend: ' + e.toString());
        Logger.log('   Stack: ' + e.stack);
        return {
          success: false,
          message: 'Error inesperado: ' + e.toString()
        };
      }
    }

    /**
     * Función auxiliar para formatear RUT para mostrar
     * (Puedes usar la existente si ya tienes una)
     */
    function formatRutDisplay(rut) {
      if (!rut) return '';
      const cleaned = cleanRut(rut);
      if (cleaned.length < 2) return cleaned;
      
      const dv = cleaned.slice(-1);
      const numero = cleaned.slice(0, -1);
      
      // Formatear con puntos
      const formatted = numero.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
      
      return formatted + '-' + dv;
    }

    /**
     * Extrae la URL de una fórmula =IMAGE("URL")
     * @param {string} formula - Fórmula de Google Sheets
     * @returns {string} URL extraída o string vacío
     */
    function extraerUrlDeImagen(formula) {
      if (!formula || typeof formula !== 'string') {
        return '';
      }
      
      // Buscar patrón: =IMAGE("URL")
      const regex = /=IMAGE\s*\(\s*"([^"]+)"\s*\)/i;
      const match = formula.match(regex);
      
      if (match && match[1]) {
        Logger.log('✅ URL extraída de fórmula IMAGE: ' + match[1]);
        return match[1];
      }
      
      // Si no es fórmula IMAGE, puede ser una URL directa
      if (formula.startsWith('http')) {
        return formula;
      }
      
      Logger.log('⚠️ No se pudo extraer URL de: ' + formula);
      return '';
    }

    // ==========================================
    // GENERADOR DE LINKS DE REGISTRO Y CÓDIGOS QR
    // ==========================================

    /**
     * URL base del Web App (con formato /a/~/macros/s/ para Google Workspace).
     * Actualizar el DEPLOYMENT_ID si se crea una nueva implementación.
     */
    const WEBAPP_BASE_URL = 'https://script.google.com/a/~/macros/s/AKfycbzrmy_GgdzMpOLfycvxxUPHU6iyuL9Jv6As_4kxG7mG8oQ4RbV-ALUZw0oeSJnqbvvc/exec';

    /**
     * Genera el "Link Registro" (columna P) y el código QR (columna Q)
     * para todos los usuarios que aún no tienen esos datos.
     *
     * - Solo escribe en filas donde la columna P esté vacía (Link Registro).
     * - Solo escribe el QR en filas donde la columna Q no tenga fórmula IMAGE.
     * - No sobreescribe datos existentes.
     *
     * Ejecutar manualmente desde el editor de Apps Script.
     */
    function generarLinksRegistroYQR() {
      try {
        Logger.log('🚀 Iniciando generarLinksRegistroYQR...');
        Logger.log('🔗 URL base: ' + WEBAPP_BASE_URL);

        // 1. Acceder a la hoja de usuarios
        const sheet = getSheet('USUARIOS', 'USUARIOS');
        if (!sheet) {
          Logger.log('❌ No se pudo acceder a la hoja de usuarios.');
          return;
        }

        const COL       = CONFIG.COLUMNAS.USUARIOS;
        const lastRow   = sheet.getLastRow();

        if (lastRow < 2) {
          Logger.log('⚠️ No hay usuarios en la hoja.');
          return;
        }

        const totalFilas = lastRow - 1;

        // 2. Leer columnas necesarias en un solo bloque para eficiencia
        //    RUT (A=1) → col índice COL.RUT+1
        //    LINK_REGISTRO (P=16) → col índice COL.LINK_REGISTRO+1
        //    QR_REGISTRO (Q=17) → col índice COL.QR_REGISTRO+1
        const rangoRut         = sheet.getRange(2, COL.RUT + 1,          totalFilas, 1).getValues();
        const rangoLinkActual  = sheet.getRange(2, COL.LINK_REGISTRO + 1, totalFilas, 1).getValues();

        // Para el QR necesitamos leer las FÓRMULAS (no los valores) para detectar si ya existe =IMAGE(...)
        const rangoQRFormulas  = sheet.getRange(2, COL.QR_REGISTRO + 1,  totalFilas, 1).getFormulas();

        let generadosLink  = 0;
        let generadosQR    = 0;
        let omitidosLink   = 0;
        let omitidosQR     = 0;
        let sinRut         = 0;

        // 3. Recorrer fila por fila
        for (let i = 0; i < totalFilas; i++) {
          const rutRaw      = String(rangoRut[i][0]).trim();
          const linkActual  = String(rangoLinkActual[i][0]).trim();
          const qrFormula   = String(rangoQRFormulas[i][0]).trim();
          const filaSheet   = i + 2; // Fila real en la hoja (1-based, saltando cabecera)

          // Saltar filas sin RUT válido
          if (!rutRaw || rutRaw === '' || rutRaw === '0' || rutRaw.toLowerCase() === 'false') {
            sinRut++;
            continue;
          }

          // Limpiar RUT usando la función existente del proyecto
          const rutLimpio = cleanRut(rutRaw);
          if (!rutLimpio || rutLimpio.length < 7) {
            Logger.log('⚠️ Fila ' + filaSheet + ': RUT inválido → ' + rutRaw);
            sinRut++;
            continue;
          }

          // ─────────────────────────────────────────────────
          // 3a. LINK REGISTRO (Columna P)
          // ─────────────────────────────────────────────────
          const tieneLink = linkActual !== '' && linkActual !== '0' && linkActual.toLowerCase() !== 'false';

          if (tieneLink) {
            omitidosLink++;
          } else {
            // Construir link con el formato exacto del ejemplo
            const linkRegistro = WEBAPP_BASE_URL + '?action=register&rut=' + rutLimpio;
            sheet.getRange(filaSheet, COL.LINK_REGISTRO + 1).setValue(linkRegistro);
            Logger.log('✅ Link generado | Fila ' + filaSheet + ' | RUT: ' + rutLimpio);
            generadosLink++;
          }

          // ─────────────────────────────────────────────────
          // 3b. CÓDIGO QR (Columna Q) — Fórmula =IMAGE(...)
          // ─────────────────────────────────────────────────
          const tieneQR = qrFormula.toUpperCase().includes('IMAGE');

          if (tieneQR) {
            omitidosQR++;
          } else {
            // Construir el link (puede que se acabe de generar arriba o ya existía)
            const linkParaQR = tieneLink
              ? linkActual  // Usar el que ya estaba en la hoja
              : WEBAPP_BASE_URL + '?action=register&rut=' + rutLimpio;

            // URL-encodear el link para incrustarlo en la URL de quickchart.io
            const linkEncoded = encodeURIComponent(linkParaQR);

            // Construir la fórmula IMAGE igual al ejemplo proporcionado
            const formulaQR = '=IMAGE("https://quickchart.io/qr?size=300&text=' + linkEncoded + '")';

            sheet.getRange(filaSheet, COL.QR_REGISTRO + 1).setFormula(formulaQR);
            Logger.log('✅ QR generado   | Fila ' + filaSheet + ' | RUT: ' + rutLimpio);
            generadosQR++;
          }

          // Pausa cada 100 filas para no superar límites de Apps Script
          if ((generadosLink + generadosQR) % 100 === 0 && (generadosLink + generadosQR) > 0) {
            Utilities.sleep(300);
          }
        }

        // 4. Resumen final en el log
        Logger.log('');
        Logger.log('══════════════════════════════════════════════════');
        Logger.log('📊 RESUMEN — generarLinksRegistroYQR');
        Logger.log('   🔗 Links generados  : ' + generadosLink);
        Logger.log('   📷 QR generados     : ' + generadosQR);
        Logger.log('   ⏭️  Links existentes : ' + omitidosLink);
        Logger.log('   ⏭️  QR existentes    : ' + omitidosQR);
        Logger.log('   ⚠️  Sin RUT válido   : ' + sinRut);
        Logger.log('══════════════════════════════════════════════════');

      } catch (e) {
        Logger.log('❌ Error en generarLinksRegistroYQR: ' + e.toString());
        Logger.log('   Stack: ' + e.stack);
        throw e;
      }
    }

    // ==========================================
// MÓDULO: CREDENCIAL SINDICAL
// ==========================================

/**
 * Obtiene el estado de credencial de un usuario desde BD_CREDENCIALES
 * @param {string} rutInput - RUT del usuario
 * @returns {string} Estado de credencial o "S/D" si no se encuentra
 */
function obtenerEstadoCredencialPorRut(rutInput) {
  try {
    const rutLimpio = cleanRut(String(rutInput));
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEETS.CREDENCIALES);
    const sheet = ss.getSheetByName(CONFIG.HOJAS.CREDENCIALES);
    
    if (!sheet) {
      Logger.log('⚠️ Hoja CREDENCIALES no encontrada');
      return "S/D";
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return "S/D";
    
    // Leer columnas A (RUT=0) y G (ESTADO CREDENCIAL=6)
    const data = sheet.getRange(2, 1, lastRow - 1, 7).getDisplayValues();
    
    for (let i = 0; i < data.length; i++) {
      const rutFila = cleanRut(String(data[i][0]));
      if (rutFila === rutLimpio) {
        const estado = String(data[i][6] || "").trim().toUpperCase();
        return estado || "S/D";
      }
    }
    
    return "S/D";
    
  } catch (e) {
    Logger.log('❌ Error obteniendo estado credencial: ' + e.toString());
    return "S/D";
  }
}

/**
 * Trigger semanal: detecta cambios de estado en credenciales y envía notificaciones
 * Se ejecuta automáticamente cada domingo a las 8am (configurar en configurarTriggers)
 */
function verificarCambiosCredenciales() {
  const ESTADOS_CON_NOTIFICACION = ["ENTREGADO", "DISPONIBLE", "SOLICITADO", "NO VIGENTE", "DATOS INCORRECTOS", "REIMPRIMIR"];
  
  try {
    Logger.log('🔄 Iniciando verificación de cambios en credenciales...');
    
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEETS.CREDENCIALES);
    const sheetImpresion = ss.getSheetByName(CONFIG.HOJAS.CREDENCIALES);
    
    if (!sheetImpresion) {
      Logger.log('❌ No se encontró la hoja IMPRESION en BD_CREDENCIALES');
      return;
    }
    
    // Obtener o crear hoja historial
    let sheetHistorial = ss.getSheetByName(CONFIG.HOJAS.HISTORIAL_CREDENCIALES);
    if (!sheetHistorial) {
      sheetHistorial = ss.insertSheet(CONFIG.HOJAS.HISTORIAL_CREDENCIALES);
      sheetHistorial.appendRow([
        'FECHA', 'RUT', 'NOMBRE', 'CORREO',
        'ESTADO ANTERIOR', 'ESTADO NUEVO', 'EMAIL ENVIADO'
      ]);
      sheetHistorial.getRange(1, 1, 1, 7).setFontWeight('bold')
        .setBackground('#1e293b')
        .setFontColor('#ffffff');
      Logger.log('✅ Hoja HISTORIAL_CREDENCIALES creada');
    }
    
    const lastRow = sheetImpresion.getLastRow();
    if (lastRow < 2) {
      Logger.log('ℹ️ No hay datos en la hoja IMPRESION');
      return;
    }
    
    // Leer columnas A-K (11 columnas)
    // A=RUT(0), B=ID(1), C=ESTADO(2), D=NOMBRE_BD(3), E=NOMBRE_FORM(4),
    // F=CORREO(5), G=ESTADO_CRED(6), H=REGION(7), I=OBS(8), J=ESTADO_ANT(9), K=EMAIL_EST(10)
    const data = sheetImpresion.getRange(2, 1, lastRow - 1, 11).getDisplayValues();
    
    let enviados = 0;
    let errores = 0;
    let inicializados = 0;
    let sinCambios = 0;
    
    for (let i = 0; i < data.length; i++) {
      const fila = i + 2; // Fila real en la hoja (fila 1 es header)
      const rut = String(data[i][0] || "").trim();
      const nombre = String(data[i][3] || data[i][4] || "Socio").trim();
      const correo = String(data[i][5] || "").trim();
      const estadoActual = String(data[i][6] || "").trim().toUpperCase();
      const estadoAnterior = String(data[i][9] || "").trim().toUpperCase();
      
      // Saltar filas sin RUT o sin estado actual
      if (!rut || !estadoActual) continue;
      
      // CASO 1: Primera vez (columna J vacía) → inicializar sin enviar email
      if (!estadoAnterior) {
        sheetImpresion.getRange(fila, 10).setValue(estadoActual); // Columna J
        inicializados++;
        Logger.log(`ℹ️ Fila ${fila} (${rut}): Inicializado con estado "${estadoActual}"`);
        continue;
      }
      
      // CASO 2: Sin cambio → no hacer nada
      if (estadoActual === estadoAnterior) {
        sinCambios++;
        continue;
      }
      
      // CASO 3: Cambio detectado → procesar
      Logger.log(`🔔 Fila ${fila} (${rut}): Cambio detectado "${estadoAnterior}" → "${estadoActual}"`);
      
      let emailEstado = "SIN CORREO";
      
      // Enviar email solo para estados con notificación definida Y si el correo es válido
      if (ESTADOS_CON_NOTIFICACION.indexOf(estadoActual) !== -1 && correo.includes('@')) {
        try {
          enviarNotificacionCredencial(correo, nombre, estadoActual, rut);
          emailEstado = "ENVIADO";
          enviados++;
          Logger.log(`✅ Email enviado a ${correo} (${nombre}) por cambio a "${estadoActual}"`);
        } catch (emailErr) {
          emailEstado = "ERROR: " + emailErr.toString().substring(0, 80);
          errores++;
          Logger.log(`❌ Error enviando email a ${correo}: ${emailErr.toString()}`);
        }
      } else if (!correo.includes('@')) {
        emailEstado = "SIN CORREO VÁLIDO";
        Logger.log(`⚠️ Fila ${fila}: Correo inválido "${correo}"`);
      } else {
        emailEstado = "ESTADO SIN NOTIF.";
      }
      
      const fechaAhora = Utilities.formatDate(new Date(), 'America/Santiago', 'dd/MM/yyyy HH:mm:ss');
      
      // Registrar en historial
      sheetHistorial.appendRow([
        fechaAhora, rut, nombre, correo,
        estadoAnterior, estadoActual, emailEstado
      ]);
      
      // Actualizar col J (estado anterior = estado actual) y col K (estado email)
      sheetImpresion.getRange(fila, 10).setValue(estadoActual);   // Col J
      sheetImpresion.getRange(fila, 11).setValue(emailEstado);    // Col K
    }
    
    Logger.log(`✅ Verificación completada:`);
    Logger.log(`   - Inicializados: ${inicializados}`);
    Logger.log(`   - Sin cambios: ${sinCambios}`);
    Logger.log(`   - Emails enviados: ${enviados}`);
    Logger.log(`   - Errores de envío: ${errores}`);
    
  } catch (e) {
    Logger.log('❌ Error crítico en verificarCambiosCredenciales: ' + e.toString());
  }
}

/**
 * Envía una notificación HTML por correo al usuario sobre el nuevo estado de su credencial
 */
function enviarNotificacionCredencial(correo, nombre, estadoNuevo, rut) {
  const MENSAJES = {
    "ENTREGADO": {
      titulo: "¡Tu credencial sindical ha sido entregada!",
      icono: "🎉",
      color: "#059669",
      colorClaro: "#d1fae5",
      colorBorde: "#6ee7b7",
      mensaje: `Tu tarjeta ha sido entregada y se encuentra disponible para su uso. Te recordamos que la credencial sindical se utiliza para las asambleas presenciales, la cual debes mostrar para la correcta marcación de tu asistencia.`,
      nota: `Si la credencial está desgastada o la extraviaste, debes solicitar una nueva a un dirigente de la organización.`
    },
    "DISPONIBLE": {
      titulo: "¡Tu credencial sindical está lista para retiro!",
      icono: "📦",
      color: "#0d9488",
      colorClaro: "#ccfbf1",
      colorBorde: "#5eead4",
      mensaje: `Tu credencial sindical ya está impresa y disponible para su retiro. Debes acercarte a la oficina sindical o retirarla en la próxima asamblea presencial.`,
      nota: `Recuerda llevar tu RUT al momento de retirarla.`
    },
    "SOLICITADO": {
      titulo: "Solicitud de credencial recibida",
      icono: "⏳",
      color: "#d97706",
      colorClaro: "#fef3c7",
      colorBorde: "#fcd34d",
      mensaje: `Tus datos han sido ingresados al sistema de nuestra organización y se ha solicitado al departamento de comunicaciones la gestión para la impresión de tu credencial sindical.`,
      nota: `Cuando esté disponible, recibirás un correo electrónico notificándote su disponibilidad de retiro.`
    },
    "NO VIGENTE": {
      titulo: "Credencial sindical deshabilitada",
      icono: "⚠️",
      color: "#dc2626",
      colorClaro: "#fee2e2",
      colorBorde: "#fca5a5",
      mensaje: `De acuerdo a nuestros registros, tu credencial sindical ha sido deshabilitada de nuestros sistemas, lo que indica que podrías ya no pertenecer a la empresa o a la organización sindical.`,
      nota: `Si consideras que existe un error, comunícate con algún dirigente de la organización para regularizar tu situación.`
    },
    "DATOS INCORRECTOS": {
      titulo: "Atención: datos incorrectos para tu credencial",
      icono: "❗",
      color: "#ea580c",
      colorClaro: "#ffedd5",
      colorBorde: "#fdba74",
      mensaje: `No se ha podido crear tu credencial sindical debido a que existen datos incorrectos para su fabricación.`,
      nota: `Debes actualizar tus datos en el módulo "Mis Datos" de la aplicación sindical, o comunicarte con un dirigente sindical para que te oriente en el proceso.`
    },
    "REIMPRIMIR": {
      titulo: "Solicitud de reimpresión recibida",
      icono: "🖨️",
      color: "#7c3aed",
      colorClaro: "#ede9fe",
      colorBorde: "#c4b5fd",
      mensaje: `Hemos recibido una solicitud de reimpresión de tu credencial sindical. Nuestro equipo está procesando esta solicitud.`,
      nota: `Una vez que esté disponible, se te notificará a través del correo electrónico registrado o puedes revisar tu estado directamente en la aplicación sindical.`
    }
  };
  
  const info = MENSAJES[estadoNuevo] || {
    titulo: "Actualización de credencial sindical",
    icono: "📋",
    color: "#64748b",
    colorClaro: "#f1f5f9",
    colorBorde: "#cbd5e1",
    mensaje: "El estado de tu credencial sindical ha sido actualizado.",
    nota: "Puedes revisar el estado actual en el módulo 'Mis Datos' de la aplicación sindical."
  };
  
  const fechaActual = Utilities.formatDate(new Date(), 'America/Santiago', "dd 'de' MMMM 'de' yyyy");
  
  const htmlBody = `
<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background-color:#0f172a;font-family:'Helvetica Neue',Arial,sans-serif;">
  <div style="max-width:600px;margin:0 auto;padding:20px;">
    
    <!-- Header -->
    <div style="background:linear-gradient(135deg,${info.color},${info.color}dd);border-radius:16px 16px 0 0;padding:32px 24px;text-align:center;">
      <div style="font-size:48px;margin-bottom:12px;">${info.icono}</div>
      <h1 style="color:#ffffff;font-size:20px;font-weight:700;margin:0 0 8px 0;line-height:1.3;">${info.titulo}</h1>
      <p style="color:rgba(255,255,255,0.85);font-size:13px;margin:0;">Sindicato SLIM N°3 · Credencial Sindical</p>
    </div>
    
    <!-- Body -->
    <div style="background:#ffffff;padding:28px 24px;">
      
      <p style="color:#374151;font-size:15px;margin:0 0 20px 0;">Estimado(a) <strong>${nombre}</strong>,</p>
      
      <!-- Estado badge -->
      <div style="background:${info.colorClaro};border:1px solid ${info.colorBorde};border-radius:12px;padding:16px;text-align:center;margin-bottom:20px;">
        <p style="color:#6b7280;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:1px;margin:0 0 8px 0;">NUEVO ESTADO DE CREDENCIAL</p>
        <span style="display:inline-block;background:${info.color};color:#ffffff;font-size:14px;font-weight:800;text-transform:uppercase;letter-spacing:1px;padding:8px 20px;border-radius:50px;">${estadoNuevo}</span>
      </div>
      
      <!-- Mensaje principal -->
      <div style="background:#f8fafc;border-left:4px solid ${info.color};border-radius:0 8px 8px 0;padding:16px;margin-bottom:16px;">
        <p style="color:#374151;font-size:14px;line-height:1.6;margin:0;">${info.mensaje}</p>
      </div>
      
      <!-- Nota adicional -->
      <div style="background:#fffbeb;border:1px solid #fde68a;border-radius:8px;padding:14px;margin-bottom:20px;">
        <p style="color:#92400e;font-size:12px;line-height:1.6;margin:0;">
          <strong>📌 Nota:</strong> ${info.nota}
        </p>
      </div>
      
      <!-- Datos del socio -->
      <table style="width:100%;border-collapse:collapse;margin-bottom:20px;font-size:13px;">
        <tr style="background:#f1f5f9;">
          <td style="padding:10px 12px;font-weight:700;color:#475569;border-bottom:1px solid #e2e8f0;width:40%;">Nombre</td>
          <td style="padding:10px 12px;color:#1e293b;border-bottom:1px solid #e2e8f0;">${nombre}</td>
        </tr>
        <tr>
          <td style="padding:10px 12px;font-weight:700;color:#475569;border-bottom:1px solid #e2e8f0;">Fecha</td>
          <td style="padding:10px 12px;color:#1e293b;border-bottom:1px solid #e2e8f0;">${fechaActual}</td>
        </tr>
      </table>
      
    </div>
    
    <!-- Footer -->
    <div style="background:#1e293b;border-radius:0 0 16px 16px;padding:20px 24px;text-align:center;">
      <p style="color:#94a3b8;font-size:12px;margin:0 0 4px 0;">Este es un mensaje automático del sistema de gestión</p>
      <p style="color:#64748b;font-size:11px;margin:0;">Sindicato SLIM N°3 · No responder a este correo</p>
    </div>
    
  </div>
</body>
</html>`;
  
  MailApp.sendEmail({
    to: correo,
    subject: `${info.icono} Credencial Sindical: ${estadoNuevo} - Sindicato SLIM N°3`,
    htmlBody: htmlBody,
    name: "Sindicato SLIM N°3"
  });
}
