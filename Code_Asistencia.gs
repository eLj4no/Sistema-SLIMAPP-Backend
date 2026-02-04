// ==========================================
// CONFIGURACIÓN - SISTEMA DE ASISTENCIA QR
// ==========================================

const CONFIG_ASISTENCIA = {
  SPREADSHEET_ID: "1SRQ8Mlc6bBdb0mitAfn4I-EUAS4BOrZRbqS9YAmg3Sk",
  SPREADSHEET_USUARIOS_ID: "1m7KLd3b3BzKOAI10I5E32MVf_L34XWAGFonhTg37TVM",
  HOJAS: {
    ASISTENCIA: "BD_ASISTENCIA",
    PUNTOS_CONTROL: "PUNTOS_CONTROL",
    USUARIOS: "BD_SLIMAPP"
  },
  COLUMNAS: {
    ASISTENCIA: {
      FECHA_HORA: 0,
      RUT: 1,
      NOMBRE: 2,
      ASAMBLEA: 3,
      TIPO_ASISTENCIA: 4,
      GESTION: 5
    },
    USUARIOS: {
      RUT: 0,
      NOMBRE: 3,
      ESTADO: 9
    }
  }
};

// ==========================================
// SERVIR HTML - PUNTO DE ENTRADA
// ==========================================

/**
 * Función principal que sirve la página HTML cuando se escanea un QR
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('QR_Asistencia');
  template.params = e.parameter; // Pasar parámetros del QR
  
  return template.evaluate()
    .setTitle('Control de Asistencia - Sindicato SLIM n°3')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

// ==========================================
// VALIDAR USUARIO POR RUT
// ==========================================

/**
 * Valida que el usuario existe y está activo
 * @param {string} rut - RUT del usuario
 * @returns {Object} {success: boolean, nombre: string, error: string}
 */
function validarUsuarioAsistencia(rut) {
  try {
    const rutLimpio = cleanRut(rut);
    
    if (!rutLimpio || rutLimpio.length < 7) {
      return {
        success: false,
        error: "RUT inválido"
      };
    }
    
    Logger.log('🔍 Validando usuario: ' + rutLimpio);
    
    // Abrir hoja de usuarios del sistema principal
    const ssUsuarios = SpreadsheetApp.openById(CONFIG_ASISTENCIA.SPREADSHEET_USUARIOS_ID);
    const sheetUsuarios = ssUsuarios.getSheetByName(CONFIG_ASISTENCIA.HOJAS.USUARIOS);
    
    if (!sheetUsuarios) {
      Logger.log('❌ No se encontró la hoja de usuarios');
      return {
        success: false,
        error: "No se pudo acceder a la base de datos de usuarios"
      };
    }
    
    const data = sheetUsuarios.getDataRange().getDisplayValues();
    const COL = CONFIG_ASISTENCIA.COLUMNAS.USUARIOS;
    
    Logger.log('📊 Total de usuarios en BD: ' + (data.length - 1));
    
    // Buscar usuario
    for (let i = 1; i < data.length; i++) {
      if (cleanRut(data[i][COL.RUT]) === rutLimpio) {
        const estado = String(data[i][COL.ESTADO] || "").toUpperCase();
        const nombre = data[i][COL.NOMBRE];
        
        Logger.log('✅ Usuario encontrado: ' + nombre);
        Logger.log('   Estado: ' + estado);
        
        // Validar estado DESVINCULADO
        if (estado === "DESVINCULADO") {
          Logger.log('⚠️ Usuario desvinculado');
          return {
            success: false,
            error: "Usuario desvinculado. No puedes registrar asistencia. Contacta con la directiva."
          };
        }
        
        return {
          success: true,
          nombre: nombre,
          rut: data[i][COL.RUT]
        };
      }
    }
    
    Logger.log('❌ Usuario no encontrado en BD');
    return {
      success: false,
      error: "RUT no encontrado en el sistema."
    };
    
  } catch (e) {
    Logger.log('❌ Error en validarUsuarioAsistencia: ' + e.toString());
    return {
      success: false,
      error: "Error del servidor: " + e.toString()
    };
  }
}

// ==========================================
// REGISTRAR ASISTENCIA
// ==========================================

/**
 * Registra la asistencia del usuario en la asamblea/actividad
 * @param {string} rut - RUT del usuario
 * @param {string} nombreControl - Nombre del punto de control (parámetro 'control')
 * @returns {Object} {success: boolean, message: string, ...}
 */
function registrarAsistencia(rut, nombreControl) {
  const lock = LockService.getScriptLock();
  
  try {
    // Obtener lock para evitar registros duplicados
    if (!lock.tryLock(30000)) {
      return {
        success: false,
        message: "Sistema ocupado, intenta nuevamente."
      };
    }
    
    Logger.log('═══════════════════════════════════════');
    Logger.log('🔵 INICIO REGISTRO DE ASISTENCIA');
    Logger.log('   RUT recibido: ' + rut);
    Logger.log('   Punto de control: ' + nombreControl);
    Logger.log('═══════════════════════════════════════');
    
    // ==========================================
    // 1. VALIDAR USUARIO
    // ==========================================
    const usuario = validarUsuarioAsistencia(rut);
    
    if (!usuario.success) {
      Logger.log('❌ Validación fallida: ' + usuario.error);
      return {
        success: false,
        message: usuario.error
      };
    }
    
    Logger.log('✅ Usuario validado: ' + usuario.nombre);
    
    // ==========================================
    // 2. VERIFICAR SI YA REGISTRÓ EN ESTA ACTIVIDAD
    // ==========================================
    const ss = SpreadsheetApp.openById(CONFIG_ASISTENCIA.SPREADSHEET_ID);
    const sheetAsistencia = ss.getSheetByName(CONFIG_ASISTENCIA.HOJAS.ASISTENCIA);
    
    if (!sheetAsistencia) {
      Logger.log('❌ No se encontró la hoja de asistencia');
      return {
        success: false,
        message: "Error: No se encontró la hoja de asistencia"
      };
    }
    
    const data = sheetAsistencia.getDataRange().getDisplayValues();
    const COL = CONFIG_ASISTENCIA.COLUMNAS.ASISTENCIA;
    const rutLimpio = cleanRut(rut);
    
    Logger.log('📊 Total de registros existentes: ' + (data.length - 1));
    
    // Verificar duplicados
    for (let i = 1; i < data.length; i++) {
      const rutRegistrado = cleanRut(data[i][COL.RUT]);
      const asambleaRegistrada = data[i][COL.ASAMBLEA];
      
      if (rutRegistrado === rutLimpio && asambleaRegistrada === nombreControl) {
        Logger.log('⚠️ Ya registrado anteriormente');
        Logger.log('   Fila del registro: ' + (i + 1));
        return {
          success: false,
          message: "Ya registraste tu asistencia en esta actividad.",
          yaRegistrado: true
        };
      }
    }
    
    // ==========================================
    // 3. REGISTRAR ASISTENCIA
    // ==========================================
    const ahora = new Date();
    const fechaHora = Utilities.formatDate(ahora, "America/Santiago", "dd/MM/yyyy HH:mm:ss");
    
    sheetAsistencia.appendRow([
      fechaHora,           // FECHA_HORA
      usuario.rut,         // RUT
      usuario.nombre,      // NOMBRE
      nombreControl,       // ASAMBLEA (nombre del punto de control)
      "",                  // TIPO_ASISTENCIA (vacío, lo llena el admin)
      "Sistema QR"         // GESTION
    ]);
    
    Logger.log('✅ Asistencia registrada exitosamente');
    Logger.log('   Fecha/Hora: ' + fechaHora);
    Logger.log('   Nueva fila: ' + (data.length + 1));
    Logger.log('═══════════════════════════════════════');
    
    return {
      success: true,
      message: "Asistencia registrada exitosamente",
      nombre: usuario.nombre,
      rut: usuario.rut,
      asamblea: nombreControl,
      fecha: fechaHora
    };
    
  } catch (e) {
    Logger.log('❌ ERROR en registrarAsistencia: ' + e.toString());
    Logger.log('   Stack: ' + e.stack);
    return {
      success: false,
      message: "Error inesperado: " + e.toString()
    };
    
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// FUNCIONES AUXILIARES
// ==========================================

/**
 * Limpia el RUT removiendo puntos y guiones
 */
function cleanRut(rut) {
  if (!rut) return "";
  return String(rut).replace(/\./g, '').replace(/-/g, '').toUpperCase().trim();
}

/**
 * Formatea el RUT para mostrar
 */
function formatRutDisplay(rut) {
  if (!rut) return '';
  const cleaned = cleanRut(rut);
  if (cleaned.length < 2) return cleaned;
  
  const dv = cleaned.slice(-1);
  const numero = cleaned.slice(0, -1);
  const formatted = numero.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  
  return formatted + '-' + dv;
}
