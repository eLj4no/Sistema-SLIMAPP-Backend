/**
 * Servir HTML - Modificado para soportar múltiples vistas
 */
function doGet(e) {
  const page = e.parameter.page;
  const action = e.parameter.action;

  // Si se solicita la página QR Access
  if (page === "qr") {
    const tpl = HtmlService.createTemplateFromFile("QR_Access");
    tpl.data = e.parameter;
    return tpl
      .evaluate()
      .setTitle("Control de Asistencia QR - SLIM N°3")
      .addMetaTag(
        "viewport",
        "width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no"
      )
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // Si viene action=register sin page=qr, redirigir a QR
  if (action === "register" && !page) {
    const tpl = HtmlService.createTemplateFromFile("QR_Access");
    tpl.data = e.parameter;
    return tpl
      .evaluate()
      .setTitle("Control de Asistencia QR - SLIM N°3")
      .addMetaTag(
        "viewport",
        "width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no"
      )
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // Vista principal (Dashboard)
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Sindicato SLIM n°3 - App Socios")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag(
      "viewport",
      "width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no"
    );
}

/**
 * Validar Usuario (Login)
 */
function validarUsuario(rutInput, passwordInput) {
  try {
    const sheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BD_USUARIOS");
    if (!sheet)
      return { success: false, message: "Error: Falta hoja BD_USUARIOS" };

    const data = sheet.getDataRange().getDisplayValues();
    const rutLimpioInput = cleanRut(rutInput);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (cleanRut(row[0]) === rutLimpioInput) {
        const passDb = String(row[11]);
        const nombreUsuario = row[3];
        const rolUsuario = row[14];
        const estadoUsuario = String(row[9]).toUpperCase();

        if (passDb === passwordInput) {
          return {
            success: true,
            message: "Login exitoso",
            user: nombreUsuario || "Socio",
            role: rolUsuario || "SOCIO",
            state: estadoUsuario || "ACTIVO",
          };
        } else {
          return { success: false, message: "Contraseña incorrecta." };
        }
      }
    }
    return { success: false, message: "RUT no encontrado." };
  } catch (e) {
    return { success: false, message: "Error Servidor: " + e.toString() };
  }
}

/**
 * Obtener Datos Completos del Usuario
 */
function obtenerDatosUsuario(rutInput) {
  try {
    const sheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BD_USUARIOS");
    const data = sheet.getDataRange().getDisplayValues();
    const rutLimpioInput = cleanRut(rutInput);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (cleanRut(row[0]) === rutLimpioInput) {
        return {
          success: true,
          datos: {
            rut: row[0] || "---",
            nombre: row[3] || "Sin Nombre",
            cargo: row[4] || "---",
            site: row[6] || "---",
            region: row[7],
            estado: String(row[9]).toUpperCase(),
            correo: row[5],
            contacto: row[13],
          },
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
  const lock = LockService.getScriptLock();
  if (lock.tryLock(10000)) {
    try {
      const sheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BD_USUARIOS");
      const data = sheet.getDataRange().getValues();
      const rutLimpioInput = cleanRut(rutInput);

      let colIndex = -1;
      if (campo === "region") colIndex = 7;
      else if (campo === "correo") colIndex = 5;
      else if (campo === "contacto") colIndex = 13;

      if (colIndex === -1) return { success: false, message: "Campo inválido" };

      for (let i = 1; i < data.length; i++) {
        if (cleanRut(String(data[i][0])) === rutLimpioInput) {
          sheet.getRange(i + 1, colIndex + 1).setValue(valor);
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
    const sheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BD_USUARIOS");
    const data = sheet.getDataRange().getDisplayValues();
    const rutLimpio = cleanRut(rutInput);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (cleanRut(row[0]) === rutLimpio) {
        const correo = row[5];
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
    const sheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BD_USUARIOS");
    const data = sheet.getDataRange().getDisplayValues();
    const rutLimpio = cleanRut(rutInput);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (cleanRut(row[0]) === rutLimpio) {
        const nombre = row[3];
        const correo = row[5];
        const password = row[11];

        if (!correo || !correo.includes("@")) {
          return {
            success: false,
            message:
              "No tienes un correo registrado. Contacta con la directiva.",
          };
        }

        enviarCorreoEstilizado(
          correo,
          "Recuperación de Contraseña - Sindicato SLIM n°3",
          "Recuperación de Contraseña",
          `Hola ${nombre}, has solicitado recuperar tu contraseña de acceso al portal.`,
          {
            "Tu contraseña es": password,
            RUT: row[0],
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

function crearSolicitudPrestamo(
  rutGestor,
  tipo,
  cuotas,
  medioPago,
  rutBeneficiario
) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(10000)) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheetUsers = ss.getSheetByName("BD_USUARIOS");
      let sheetLoans = ss.getSheetByName("PRESTAMOS");

      if (!sheetLoans) {
        sheetLoans = ss.insertSheet("PRESTAMOS");
        sheetLoans.appendRow([
          "ID",
          "Fecha",
          "Rut",
          "Nombre",
          "Correo",
          "Tipo Préstamo",
          "Cuotas",
          "Medio Pago",
          "Estado",
          "Fecha Termino",
          "Gestión",
          "Nombre Dirigente",
          "Correo Dirigente",
        ]);
      }

      const dataUsers = sheetUsers.getDataRange().getDisplayValues();

      let gestor = null;
      const rutLimpioGestor = cleanRut(rutGestor);
      for (let i = 1; i < dataUsers.length; i++) {
        if (cleanRut(dataUsers[i][0]) === rutLimpioGestor) {
          gestor = {
            rut: dataUsers[i][0],
            nombre: dataUsers[i][3],
            correo: dataUsers[i][5],
          };
          break;
        }
      }
      if (!gestor) return { success: false, message: "Error de sesión." };

      let rutTarget = rutBeneficiario
        ? cleanRut(rutBeneficiario)
        : rutLimpioGestor;
      let beneficiario = null;

      if (rutTarget === rutLimpioGestor) {
        beneficiario = gestor;
      } else {
        for (let i = 1; i < dataUsers.length; i++) {
          if (cleanRut(dataUsers[i][0]) === rutTarget) {
            beneficiario = {
              rut: dataUsers[i][0],
              nombre: dataUsers[i][3],
              correo: dataUsers[i][5],
            };
            break;
          }
        }
        if (!beneficiario)
          return { success: false, message: "RUT del socio no encontrado." };
      }

      const dataLoans = sheetLoans.getDataRange().getDisplayValues();
      for (let i = 1; i < dataLoans.length; i++) {
        const row = dataLoans[i];
        const rowRut = cleanRut(row[2]);
        const rowTipo = row[5];
        const rowEstado = row[8];
        const estadosActivos = ["Solicitado", "Enviado", "Vigente"];

        if (
          rowRut === cleanRut(beneficiario.rut) &&
          rowTipo === tipo &&
          estadosActivos.includes(rowEstado)
        ) {
          return {
            success: false,
            message: `El socio ${beneficiario.nombre} ya tiene un ${tipo} en estado "${rowEstado}".`,
          };
        }
      }

      const fechaSolicitud = new Date();
      const fechaTermino = new Date(fechaSolicitud);
      fechaTermino.setMonth(fechaTermino.getMonth() + parseInt(cuotas));
      const idUnico = Utilities.getUuid();
      const estadoInicial = "Solicitado";

      let gestion = "Socio";
      let nomDirigente = "";
      let correoDirigente = "";

      if (rutTarget !== rutLimpioGestor) {
        gestion = "Dirigente";
        nomDirigente = gestor.nombre;
        correoDirigente = gestor.correo;
      }

      sheetLoans.appendRow([
        idUnico,
        fechaSolicitud,
        beneficiario.rut,
        beneficiario.nombre,
        beneficiario.correo,
        tipo,
        cuotas,
        medioPago,
        estadoInicial,
        fechaTermino,
        gestion,
        nomDirigente,
        correoDirigente,
      ]);

      if (beneficiario.correo && beneficiario.correo.includes("@")) {
        let mensajeExtra =
          gestion === "Dirigente"
            ? `<br><em>(Ingresado por: ${nomDirigente})</em>`
            : "";
        enviarCorreoEstilizado(
          beneficiario.correo,
          "Confirmación de Solicitud - Sindicato SLIM n°3",
          "¡Solicitud Recibida!",
          `Hola ${beneficiario.nombre}, se ha ingresado una solicitud de préstamo a tu nombre.`,
          {
            ID: idUnico,
            Tipo: tipo,
            Cuotas: cuotas,
            Gestión: gestion + mensajeExtra,
            Fecha: fechaSolicitud.toLocaleDateString(),
          },
          "#2563eb"
        );
      }

      if (
        gestion === "Dirigente" &&
        correoDirigente &&
        correoDirigente.includes("@") &&
        correoDirigente !== beneficiario.correo
      ) {
        enviarCorreoEstilizado(
          correoDirigente,
          "Respaldo de Gestión - Sindicato SLIM n°3",
          "Solicitud Ingresada",
          `Has ingresado exitosamente una solicitud para el socio <strong>${beneficiario.nombre}</strong>.`,
          {
            ID: idUnico,
            Socio: beneficiario.nombre,
            RUT: beneficiario.rut,
            Tipo: tipo,
            Cuotas: cuotas,
          },
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

function obtenerHistorialPrestamos(rutInput) {
  try {
    const sheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PRESTAMOS");
    if (!sheet) return { success: true, registros: [] };

    const data = sheet.getDataRange().getDisplayValues();
    const rutLimpio = cleanRut(rutInput);
    const registros = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (cleanRut(row[2]) === rutLimpio) {
        registros.push({
          id: row[0],
          fecha: row[1],
          tipo: row[5],
          cuotas: row[6],
          medio: row[7],
          estado: row[8],
          fechaTermino: row[9],
          gestion: row[10],
          nomDirigente: row[11],
        });
      }
    }
    registros.reverse();
    return { success: true, registros: registros };
  } catch (e) {
    return { success: false, message: "Error: " + e.toString() };
  }
}

function eliminarSolicitud(idSolicitud) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(10000)) {
    try {
      const sheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PRESTAMOS");
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(idSolicitud)) {
          const estado = String(data[i][8]);
          const estadosPermitidos = ["Solicitado", "Pagado", "Rechazado"];
          if (!estadosPermitidos.includes(estado)) {
            return {
              success: false,
              message: "No se puede eliminar un préstamo activo.",
            };
          }
          sheet.deleteRow(i + 1);
          return {
            success: true,
            message: "Registro eliminado correctamente.",
          };
        }
      }
      return { success: false, message: "No encontrado." };
    } catch (e) {
      return { success: false, message: "Error: " + e.toString() };
    } finally {
      lock.releaseLock();
    }
  } else {
    return { success: false, message: "Servidor ocupado." };
  }
}

function modificarSolicitud(idSolicitud, nuevasCuotas, nuevoMedio) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(10000)) {
    try {
      const sheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PRESTAMOS");
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(idSolicitud)) {
          const estado = String(data[i][8]);
          if (estado !== "Solicitado")
            return {
              success: false,
              message: "No se puede editar. Estado: " + estado,
            };

          sheet.getRange(i + 1, 7).setValue(nuevasCuotas);
          sheet.getRange(i + 1, 8).setValue(nuevoMedio);

          const fechaSolicitud = new Date(data[i][1]);
          const nuevaFechaTermino = new Date(fechaSolicitud);
          nuevaFechaTermino.setMonth(
            nuevaFechaTermino.getMonth() + parseInt(nuevasCuotas)
          );
          sheet.getRange(i + 1, 10).setValue(nuevaFechaTermino);

          return { success: true, message: "Modificado correctamente." };
        }
      }
      return { success: false, message: "No encontrado." };
    } catch (e) {
      return { success: false, message: "Error: " + e.toString() };
    } finally {
      lock.releaseLock();
    }
  } else {
    return { success: false, message: "Servidor ocupado." };
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
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetConfig = ss.getSheetByName("CONFIG_JUSTIFICACIONES");

    if (!sheetConfig) {
      sheetConfig = ss.insertSheet("CONFIG_JUSTIFICACIONES");
      sheetConfig.appendRow(["Habilitado", "Fecha Límite"]);
      sheetConfig.appendRow([false, ""]);
    }

    const data = sheetConfig.getRange(2, 1, 1, 2).getValues();
    const habilitado =
      data[0][0] === true || data[0][0] === "TRUE" || data[0][0] === "true";
    const fechaLimite = data[0][1];

    if (habilitado && fechaLimite) {
      const ahora = new Date();
      const limite = new Date(fechaLimite);

      if (ahora > limite) {
        sheetConfig.getRange(2, 1).setValue(false);
        return { habilitado: false, fechaLimite: fechaLimite };
      }
    }

    return { habilitado: habilitado, fechaLimite: fechaLimite };
  } catch (e) {
    return { habilitado: false, fechaLimite: "", error: e.toString() };
  }
}

/**
 * Actualizar estado del switch de justificaciones
 */
function actualizarSwitchJustificaciones(nuevoEstado, fechaLimite) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(10000)) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let sheetConfig = ss.getSheetByName("CONFIG_JUSTIFICACIONES");

      if (!sheetConfig) {
        sheetConfig = ss.insertSheet("CONFIG_JUSTIFICACIONES");
        sheetConfig.appendRow(["Habilitado", "Fecha Límite"]);
        sheetConfig.appendRow([false, ""]);
      }

      sheetConfig.getRange(2, 1).setValue(nuevoEstado);
      if (fechaLimite) {
        sheetConfig.getRange(2, 2).setValue(fechaLimite);
      }

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
      mensaje:
        "Módulo de justificaciones temporalmente deshabilitado.\nConsulte con la directiva.",
    };
  }

  return { habilitado: true };
}

function enviarJustificacion(
  rutGestor,
  tipo,
  motivo,
  archivoData,
  rutBeneficiario
) {
  const CARPETA_ID = "1Tyx4RTAX59M01uk_zJhsJjyYMydtfH03";

  const disp = verificarDisponibilidadJustificaciones();
  if (!disp.habilitado) return { success: false, message: disp.mensaje };

  const lock = LockService.getScriptLock();
  if (lock.tryLock(30000)) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheetUsers = ss.getSheetByName("BD_USUARIOS");
      let sheetJustif = ss.getSheetByName("JUSTIFICACIONES");

      if (!sheetJustif) {
        sheetJustif = ss.insertSheet("JUSTIFICACIONES");
        sheetJustif.appendRow([
          "ID",
          "Fecha",
          "Rut",
          "Nombre",
          "Correo",
          "Tipo",
          "Motivo",
          "URL Archivo",
          "Estado",
          "Observación",
          "Notificado",
          "ASAMBLEA",
          "Gestión",
          "Nombre Dirigente",
          "Correo Dirigente",
        ]);
      }

      const dataUsers = sheetUsers.getDataRange().getDisplayValues();

      let gestor = null;
      const rutLimpioGestor = cleanRut(rutGestor);
      for (let i = 1; i < dataUsers.length; i++) {
        if (cleanRut(dataUsers[i][0]) === rutLimpioGestor) {
          gestor = {
            rut: dataUsers[i][0],
            nombre: dataUsers[i][3],
            correo: dataUsers[i][5],
          };
          break;
        }
      }
      if (!gestor) return { success: false, message: "Error de sesión." };

      let rutTarget = rutBeneficiario
        ? cleanRut(rutBeneficiario)
        : rutLimpioGestor;
      let beneficiario = null;

      if (rutTarget === rutLimpioGestor) {
        beneficiario = gestor;
      } else {
        for (let i = 1; i < dataUsers.length; i++) {
          if (cleanRut(dataUsers[i][0]) === rutTarget) {
            beneficiario = {
              rut: dataUsers[i][0],
              nombre: dataUsers[i][3],
              correo: dataUsers[i][5],
            };
            break;
          }
        }
        if (!beneficiario)
          return { success: false, message: "RUT del socio no encontrado." };
      }

      const dataJustif = sheetJustif.getDataRange().getDisplayValues();
      for (let i = 1; i < dataJustif.length; i++) {
        if (
          cleanRut(dataJustif[i][2]) === cleanRut(beneficiario.rut) &&
          dataJustif[i][8] === "Enviado"
        ) {
          return {
            success: false,
            message: "Ya tienes una justificación pendiente.",
          };
        }
      }

      const idUnico = Utilities.getUuid();
      let fileUrl = "Sin archivo";

      if (archivoData && archivoData.base64) {
        const sizeInBytes = (archivoData.base64.length * 3) / 4;
        if (sizeInBytes > 5 * 1024 * 1024) {
          return {
            success: false,
            message: "El archivo es demasiado grande (máx 5MB).",
          };
        }

        try {
          const folder = DriveApp.getFolderById(CARPETA_ID);
          const blob = Utilities.newBlob(
            Utilities.base64Decode(archivoData.base64),
            archivoData.mimeType,
            archivoData.fileName
          );

          let extension = "";
          const nameParts = archivoData.fileName.split(".");
          if (nameParts.length > 1) extension = "." + nameParts.pop();

          blob.setName(
            `${idUnico} - ${cleanRut(beneficiario.rut)}${extension}`
          );

          const file = folder.createFile(blob);
          file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);

          if (beneficiario.correo && beneficiario.correo.includes("@")) {
            try {
              file.addViewer(beneficiario.correo);
            } catch (permErr) {
              console.log("Permiso drive error: " + permErr);
            }
          }
          fileUrl = file.getUrl();
        } catch (errDrive) {
          console.error("Error Drive: " + errDrive);
          fileUrl = "Error subida";
        }
      }

      const fechaHoy = new Date();
      const estado = "Enviado";

      let gestion = "Socio";
      let nomDirigente = "";
      let correoDirigente = "";

      if (rutTarget !== rutLimpioGestor) {
        gestion = "Dirigente";
        nomDirigente = gestor.nombre;
        correoDirigente = gestor.correo;
      }

      sheetJustif.appendRow([
        idUnico,
        fechaHoy,
        beneficiario.rut,
        beneficiario.nombre,
        beneficiario.correo,
        tipo,
        motivo,
        fileUrl,
        estado,
        "",
        estado,
        "",
        gestion,
        nomDirigente,
        correoDirigente,
      ]);

      const lastRow = sheetJustif.getLastRow();
      const cellEstado = sheetJustif.getRange(lastRow, 9);
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(
          ["Enviado", "Aceptado", "Aceptado/Obs", "Rechazado"],
          true
        )
        .setAllowInvalid(false)
        .build();
      cellEstado.setDataValidation(rule);

      if (beneficiario.correo && beneficiario.correo.includes("@")) {
        let mensajeExtra =
          gestion === "Dirigente"
            ? `<br><em>(Ingresado por: ${nomDirigente})</em>`
            : "";
        enviarCorreoEstilizado(
          beneficiario.correo,
          "Justificación Recibida - Sindicato SLIM n°3",
          "Justificación Enviada",
          `Hola ${beneficiario.nombre}, hemos recibido tu justificación.`,
          {
            ID: idUnico,
            Tipo: tipo,
            Motivo: motivo,
            Gestión: gestion + mensajeExtra,
            Archivo: fileUrl.includes("http") ? "Adjuntado" : "No adjuntado",
            Estado: estado,
          },
          "#ea580c"
        );
      }

      if (
        gestion === "Dirigente" &&
        correoDirigente &&
        correoDirigente.includes("@") &&
        correoDirigente !== beneficiario.correo
      ) {
        enviarCorreoEstilizado(
          correoDirigente,
          "Respaldo Gestión Justificación - Sindicato SLIM n°3",
          "Justificación Ingresada",
          `Has ingresado exitosamente una justificación para el socio <strong>${beneficiario.nombre}</strong>.`,
          {
            ID: idUnico,
            Socio: beneficiario.nombre,
            Tipo: tipo,
            Motivo: motivo,
          },
          "#475569"
        );
      }

      return { success: true, message: "Justificación enviada exitosamente." };
    } catch (e) {
      return { success: false, message: "Error: " + e.toString() };
    } finally {
      lock.releaseLock();
    }
  } else {
    return { success: false, message: "Servidor ocupado." };
  }
}

function obtenerHistorialJustificaciones(rutInput) {
  try {
    const sheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("JUSTIFICACIONES");
    if (!sheet) return { success: true, registros: [] };

    const data = sheet.getDataRange().getDisplayValues();
    const rutLimpio = cleanRut(rutInput);
    const registros = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (cleanRut(row[2]) === rutLimpio) {
        registros.push({
          id: row[0],
          fecha: row[1],
          tipo: row[5],
          motivo: row[6],
          url: row[7],
          estado: row[8],
          obs: row[9],
          asamblea: row[11],
          gestion: row[12],
          nomDirigente: row[13],
        });
      }
    }
    registros.reverse();
    return { success: true, registros: registros };
  } catch (e) {
    return { success: false, message: "Error: " + e.toString() };
  }
}

function eliminarJustificacion(idJustif) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(10000)) {
    try {
      const sheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("JUSTIFICACIONES");
      const data = sheet.getDataRange().getValues();

      const estadoSwitch = obtenerEstadoSwitchJustificaciones();

      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(idJustif)) {
          const estado = String(data[i][8]);

          if (!estadoSwitch.habilitado && estado === "Enviado") {
            return {
              success: false,
              message:
                "El plazo para agregar o modificar información ha vencido. Si al final del mes aparece con multa puede realizar la apelación.",
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
        mensaje: "No se pueden apelar meses anteriores a Marzo 2025.",
      };
    }

    const mesActual = hoy.getMonth();
    const yearActual = hoy.getFullYear();

    if (yearSel === yearActual && monthSel === mesActual) {
      if (diaActual < 25) {
        return {
          habilitado: false,
          mensaje:
            "Las apelaciones del mes en curso solo están disponibles a partir del día 25.",
        };
      }
    }

    const fechaHoy = new Date(yearActual, mesActual, 1);
    fechaHoy.setHours(0, 0, 0, 0);

    if (fechaSeleccionada > fechaHoy) {
      return {
        habilitado: false,
        mensaje: "No se pueden apelar meses futuros.",
      };
    }

    return { habilitado: true };
  } catch (e) {
    return {
      habilitado: false,
      mensaje: "Error validando disponibilidad: " + e.toString(),
    };
  }
}

function enviarApelacion(
  rutGestor,
  mesApelacion,
  tipoMotivo,
  detalleMotivo,
  archivoComprobante,
  archivoLiquidacion,
  rutBeneficiario
) {
  const CARPETA_COMPROBANTES_ID = "12LJ9YQggDLtHiVismYOMz30OnomY37oP";
  const CARPETA_LIQUIDACIONES_ID = "15B19WnWy-3hDkst730J00qwUhpz4VQrZ";

  const validacion = verificarDisponibilidadApelaciones(mesApelacion);
  if (!validacion.habilitado) {
    return { success: false, message: validacion.mensaje };
  }

  const lock = LockService.getScriptLock();
  if (lock.tryLock(30000)) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheetUsers = ss.getSheetByName("BD_USUARIOS");
      let sheetApelaciones = ss.getSheetByName("APELACIONES");

      if (!sheetApelaciones) {
        sheetApelaciones = ss.insertSheet("APELACIONES");
        sheetApelaciones.appendRow([
          "ID",
          "Fecha",
          "Rut",
          "Nombre",
          "Correo",
          "Mes Apelación",
          "Tipo Motivo",
          "Detalle Motivo",
          "URL Comprobante",
          "URL Liquidación",
          "Estado",
          "Observación",
          "Notificado",
          "Gestión",
          "Nombre Dirigente",
          "Correo Dirigente",
          "URL Comprobante Devolución",
        ]);
      }

      const dataUsers = sheetUsers.getDataRange().getDisplayValues();

      let gestor = null;
      const rutLimpioGestor = cleanRut(rutGestor);
      for (let i = 1; i < dataUsers.length; i++) {
        if (cleanRut(dataUsers[i][0]) === rutLimpioGestor) {
          gestor = {
            rut: dataUsers[i][0],
            nombre: dataUsers[i][3],
            correo: dataUsers[i][5],
          };
          break;
        }
      }
      if (!gestor) return { success: false, message: "Error de sesión." };

      let rutTarget = rutBeneficiario
        ? cleanRut(rutBeneficiario)
        : rutLimpioGestor;
      let beneficiario = null;

      if (rutTarget === rutLimpioGestor) {
        beneficiario = gestor;
      } else {
        for (let i = 1; i < dataUsers.length; i++) {
          if (cleanRut(dataUsers[i][0]) === rutTarget) {
            beneficiario = {
              rut: dataUsers[i][0],
              nombre: dataUsers[i][3],
              correo: dataUsers[i][5],
            };
            break;
          }
        }
        if (!beneficiario)
          return { success: false, message: "RUT del socio no encontrado." };
      }

      const dataApelaciones = sheetApelaciones
        .getDataRange()
        .getDisplayValues();
      for (let i = 1; i < dataApelaciones.length; i++) {
        const row = dataApelaciones[i];
        const estadoActual = String(row[10]);
        const estadosBloqueantes = ["Enviado", "Aceptado", "Aceptado-Obs"];

        if (
          cleanRut(row[2]) === cleanRut(beneficiario.rut) &&
          row[5] === mesApelacion &&
          estadosBloqueantes.includes(estadoActual)
        ) {
          let mensajeError = "";
          if (estadoActual === "Enviado") {
            mensajeError =
              "Ya tienes una apelación pendiente para este mes. Verifica el estado en tu historial.";
          } else if (
            estadoActual === "Aceptado" ||
            estadoActual === "Aceptado-Obs"
          ) {
            mensajeError =
              "Este mes ya fue resuelto favorablemente. Verifica los detalles en tu historial.";
          }

          return { success: false, message: mensajeError };
        }
      }

      if (archivoComprobante && archivoComprobante.base64) {
        const sizeComp = (archivoComprobante.base64.length * 3) / 4;
        if (sizeComp > 5 * 1024 * 1024) {
          return {
            success: false,
            message: "El comprobante es muy pesado (máx 5MB).",
          };
        }
      }

      if (archivoLiquidacion && archivoLiquidacion.base64) {
        const sizeLiq = (archivoLiquidacion.base64.length * 3) / 4;
        if (sizeLiq > 5 * 1024 * 1024) {
          return {
            success: false,
            message: "La liquidación es muy pesada (máx 5MB).",
          };
        }
      }

      const idUnico = Utilities.getUuid();

      let urlComprobante = "Sin archivo";
      if (archivoComprobante && archivoComprobante.base64) {
        try {
          const folderComp = DriveApp.getFolderById(CARPETA_COMPROBANTES_ID);
          const blobComp = Utilities.newBlob(
            Utilities.base64Decode(archivoComprobante.base64),
            archivoComprobante.mimeType,
            archivoComprobante.fileName
          );

          let extensionComp = "";
          const namePartsComp = archivoComprobante.fileName.split(".");
          if (namePartsComp.length > 1)
            extensionComp = "." + namePartsComp.pop();

          blobComp.setName(
            `APEL-COMP-${idUnico}-${cleanRut(beneficiario.rut)}${extensionComp}`
          );

          const fileComp = folderComp.createFile(blobComp);
          fileComp.setSharing(
            DriveApp.Access.PRIVATE,
            DriveApp.Permission.NONE
          );

          if (beneficiario.correo && beneficiario.correo.includes("@")) {
            try {
              fileComp.addViewer(beneficiario.correo);
            } catch (permErr) {
              console.log("Permiso drive comprobante: " + permErr);
            }
          }

          urlComprobante = fileComp.getUrl();
        } catch (errComp) {
          console.error("Error subiendo comprobante: " + errComp);
          urlComprobante = "Error subida comprobante";
        }
      }

      let urlLiquidacion = "Error: No se subió liquidación";
      if (!archivoLiquidacion || !archivoLiquidacion.base64) {
        return {
          success: false,
          message: "La liquidación de sueldo es obligatoria.",
        };
      }

      try {
        const folderLiq = DriveApp.getFolderById(CARPETA_LIQUIDACIONES_ID);
        const blobLiq = Utilities.newBlob(
          Utilities.base64Decode(archivoLiquidacion.base64),
          archivoLiquidacion.mimeType,
          archivoLiquidacion.fileName
        );

        let extensionLiq = "";
        const namePartsLiq = archivoLiquidacion.fileName.split(".");
        if (namePartsLiq.length > 1) extensionLiq = "." + namePartsLiq.pop();

        blobLiq.setName(
          `APEL-LIQ-${idUnico}-${cleanRut(beneficiario.rut)}${extensionLiq}`
        );

        const fileLiq = folderLiq.createFile(blobLiq);
        fileLiq.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);

        if (beneficiario.correo && beneficiario.correo.includes("@")) {
          try {
            fileLiq.addViewer(beneficiario.correo);
          } catch (permErr) {
            console.log("Permiso drive liquidación: " + permErr);
          }
        }

        urlLiquidacion = fileLiq.getUrl();
      } catch (errLiq) {
        console.error("Error subiendo liquidación: " + errLiq);
        return {
          success: false,
          message: "Error al subir la liquidación. Intente nuevamente.",
        };
      }

      const fechaHoy = new Date();
      const estado = "Enviado";

      let gestion = "Socio";
      let nomDirigente = "";
      let correoDirigente = "";

      if (rutTarget !== rutLimpioGestor) {
        gestion = "Dirigente";
        nomDirigente = gestor.nombre;
        correoDirigente = gestor.correo;
      }

      sheetApelaciones.appendRow([
        idUnico,
        fechaHoy,
        beneficiario.rut,
        beneficiario.nombre,
        beneficiario.correo,
        mesApelacion,
        tipoMotivo,
        detalleMotivo || "",
        urlComprobante,
        urlLiquidacion,
        estado,
        "",
        estado,
        gestion,
        nomDirigente,
        correoDirigente,
        "",
      ]);

      const lastRow = sheetApelaciones.getLastRow();
      const cellEstado = sheetApelaciones.getRange(lastRow, 11);
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(
          ["Enviado", "Aceptado", "Aceptado-Obs", "Rechazado"],
          true
        )
        .setAllowInvalid(false)
        .build();
      cellEstado.setDataValidation(rule);

      if (beneficiario.correo && beneficiario.correo.includes("@")) {
        let mensajeExtra =
          gestion === "Dirigente"
            ? `<br><em>(Ingresado por: ${nomDirigente})</em>`
            : "";

        const partesMes = mesApelacion.split("-");
        const fechaMes = new Date(`${mesApelacion}-02`);
        const nombreMes = fechaMes.toLocaleString("es-CL", {
          month: "long",
          year: "numeric",
        });

        enviarCorreoEstilizado(
          beneficiario.correo,
          "Apelación Recibida - Sindicato SLIM n°3",
          "Apelación Enviada",
          `Hola ${beneficiario.nombre}, hemos recibido tu apelación de multa.`,
          {
            ID: idUnico,
            "Mes Apelado": nombreMes.toUpperCase(),
            Motivo: tipoMotivo,
            Detalle: detalleMotivo || "Sin detalle adicional",
            Gestión: gestion + mensajeExtra,
            Liquidación: urlLiquidacion.includes("http")
              ? "Adjuntada"
              : "Error",
            Comprobante: urlComprobante.includes("http")
              ? "Adjuntado"
              : "No adjuntado",
            Estado: estado,
          },
          "#dc2626"
        );
      }

      if (
        gestion === "Dirigente" &&
        correoDirigente &&
        correoDirigente.includes("@") &&
        correoDirigente !== beneficiario.correo
      ) {
        enviarCorreoEstilizado(
          correoDirigente,
          "Respaldo Gestión Apelación - Sindicato SLIM n°3",
          "Apelación Ingresada",
          `Has ingresado exitosamente una apelación para el socio <strong>${beneficiario.nombre}</strong>.`,
          {
            ID: idUnico,
            Socio: beneficiario.nombre,
            Mes: mesApelacion,
            Motivo: tipoMotivo,
          },
          "#475569"
        );
      }

      return { success: true, message: "Apelación enviada exitosamente." };
    } catch (e) {
      return {
        success: false,
        message: "Error al enviar apelación: " + e.toString(),
      };
    } finally {
      lock.releaseLock();
    }
  } else {
    return { success: false, message: "Servidor ocupado." };
  }
}

function obtenerHistorialApelaciones(rutInput) {
  try {
    const sheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("APELACIONES");
    if (!sheet) return { success: true, registros: [] };

    const data = sheet.getDataRange().getDisplayValues();
    const rutLimpio = cleanRut(rutInput);
    const registros = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (cleanRut(row[2]) === rutLimpio) {
        registros.push({
          id: row[0],
          fecha: row[1],
          mesApelacion: row[5],
          tipoMotivo: row[6],
          detalleMotivo: row[7],
          urlComprobante: row[8],
          urlLiquidacion: row[9],
          estado: row[10],
          obs: row[11],
          gestion: row[13],
          nomDirigente: row[14],
          urlComprobanteDevolucion: row[16] || "",
        });
      }
    }

    registros.reverse();
    return { success: true, registros: registros };
  } catch (e) {
    return { success: false, message: "Error: " + e.toString() };
  }
}

function eliminarApelacion(idApelacion) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(10000)) {
    try {
      const sheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("APELACIONES");
      const data = sheet.getDataRange().getValues();

      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(idApelacion)) {
          const estado = String(data[i][10]);

          if (estado !== "Enviado") {
            return {
              success: false,
              message: "No se puede eliminar una apelación procesada.",
            };
          }

          sheet.deleteRow(i + 1);
          return {
            success: true,
            message: "Apelación eliminada correctamente.",
          };
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
    const sheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("APELACIONES");
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const idRegistro = String(row[0]);
      const estadoActual = String(row[10]);
      const estadoNotif = String(row[12]);
      const correo = row[4];
      const nombre = row[3];
      const mesApel = String(row[5]);
      const tipoMotivo = row[6];
      const obs = row[11];

      if (estadoActual !== estadoNotif) {
        if (correo && correo.includes("@")) {
          let color = "#dc2626";
          let titulo = "Actualización de Apelación";

          if (estadoActual.includes("Aceptado")) {
            color = "#15803d";
            titulo = "Apelación Aceptada";
          } else if (estadoActual.includes("Rechazado")) {
            color = "#b91c1c";
            titulo = "Apelación Rechazada";
          }

          const partesMes = mesApel.split("-");
          const fechaMes = new Date(`${mesApel}-02`);
          const nombreMes = fechaMes.toLocaleString("es-CL", {
            month: "long",
            year: "numeric",
          });

          enviarCorreoEstilizado(
            correo,
            titulo + " - Sindicato SLIM n°3",
            titulo,
            `Hola ${nombre}, el estado de tu apelación ha cambiado.`,
            {
              ID: idRegistro,
              "Mes Apelado": nombreMes.toUpperCase(),
              Motivo: tipoMotivo,
              "Nuevo Estado": estadoActual,
              Observación: obs || "Sin observaciones",
            },
            color
          );
        }

        sheet.getRange(i + 1, 13).setValue(estadoActual);
      }
    }
  } catch (e) {
    console.error("Error verificando apelaciones: " + e);
  }
}

function procesarPermisosComprobantesDevolucion() {
  try {
    const sheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("APELACIONES");
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const urlComprobanteDevolucion = String(row[16]);
      const correoUsuario = row[4];

      if (
        urlComprobanteDevolucion &&
        urlComprobanteDevolucion.includes("drive.google.com") &&
        correoUsuario &&
        correoUsuario.includes("@")
      ) {
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
            const hasAccess = viewers.some(
              (viewer) => viewer.getEmail() === correoUsuario
            );

            if (!hasAccess) {
              file.addViewer(correoUsuario);
              console.log(
                `Permiso otorgado a ${correoUsuario} para archivo ${fileId}`
              );
            }
          }
        } catch (fileErr) {
          console.error(
            `Error procesando archivo para fila ${i + 1}: ${fileErr}`
          );
        }
      }
    }
  } catch (e) {
    console.error("Error en procesarPermisosComprobantesDevolucion: " + e);
  }
}

// ==========================================
// MÓDULO: PERMISO MÉDICO
// ==========================================

function solicitarPermisoMedico(
  rutGestor,
  tipoPermiso,
  fechaInicio,
  motivo,
  rutBeneficiario
) {
  const CORREO_REPRESENTANTE_LEGAL = "penailillo.fetrasiss@gmail.com"; // CAMBIAR POR CORREO REAL

  const lock = LockService.getScriptLock();
  if (lock.tryLock(30000)) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheetUsers = ss.getSheetByName("BD_USUARIOS");
      let sheetPermisos = ss.getSheetByName("PERMISO-MEDICO");

      if (!sheetPermisos) {
        sheetPermisos = ss.insertSheet("PERMISO-MEDICO");
        sheetPermisos.appendRow([
          "ID",
          "Fecha Solicitud",
          "Rut",
          "Nombre",
          "Correo",
          "Tipo Permiso",
          "Fecha Inicio Permiso",
          "Motivo/Detalle",
          "URL Documento Respaldo",
          "Estado",
          "Fecha Subida Documento",
          "Notificado Rep. Legal",
          "Gestión",
          "Nombre Dirigente",
          "Correo Dirigente",
        ]);
      }

      const dataUsers = sheetUsers.getDataRange().getDisplayValues();

      let gestor = null;
      const rutLimpioGestor = cleanRut(rutGestor);
      for (let i = 1; i < dataUsers.length; i++) {
        if (cleanRut(dataUsers[i][0]) === rutLimpioGestor) {
          gestor = {
            rut: dataUsers[i][0],
            nombre: dataUsers[i][3],
            correo: dataUsers[i][5],
          };
          break;
        }
      }
      if (!gestor) return { success: false, message: "Error de sesión." };

      let rutTarget = rutBeneficiario
        ? cleanRut(rutBeneficiario)
        : rutLimpioGestor;
      let beneficiario = null;

      if (rutTarget === rutLimpioGestor) {
        beneficiario = gestor;
      } else {
        for (let i = 1; i < dataUsers.length; i++) {
          if (cleanRut(dataUsers[i][0]) === rutTarget) {
            beneficiario = {
              rut: dataUsers[i][0],
              nombre: dataUsers[i][3],
              correo: dataUsers[i][5],
            };
            break;
          }
        }
        if (!beneficiario)
          return { success: false, message: "RUT del socio no encontrado." };
      }

      const idUnico = Utilities.getUuid();
      const fechaHoy = new Date();
      const estado = "Solicitado";

      let gestion = "Socio";
      let nomDirigente = "";
      let correoDirigente = "";

      if (rutTarget !== rutLimpioGestor) {
        gestion = "Dirigente";
        nomDirigente = gestor.nombre;
        correoDirigente = gestor.correo;
      }

      sheetPermisos.appendRow([
        idUnico,
        fechaHoy,
        beneficiario.rut,
        beneficiario.nombre,
        beneficiario.correo,
        tipoPermiso,
        fechaInicio,
        motivo,
        "Sin documento",
        estado,
        "",
        false,
        gestion,
        nomDirigente,
        correoDirigente,
      ]);

      if (beneficiario.correo && beneficiario.correo.includes("@")) {
        let mensajeExtra =
          gestion === "Dirigente"
            ? `<br><em>(Ingresado por: ${nomDirigente})</em>`
            : "";

        const fechaInicioObj = new Date(fechaInicio);
        const fechaInicioStr = fechaInicioObj.toLocaleDateString("es-CL", {
          year: "numeric",
          month: "long",
          day: "numeric",
        });

        enviarCorreoEstilizado(
          beneficiario.correo,
          "Solicitud Permiso Médico - Sindicato SLIM n°3",
          "Permiso Médico Solicitado",
          `Hola ${beneficiario.nombre}, se ha registrado tu solicitud de permiso médico. <strong>IMPORTANTE:</strong> Debes adjuntar el documento de respaldo en el historial del módulo.`,
          {
            ID: idUnico,
            Tipo: tipoPermiso,
            "Fecha Inicio": fechaInicioStr,
            Motivo: motivo,
            Gestión: gestion + mensajeExtra,
            Estado: estado,
            "Acción Requerida": "Adjuntar documento de respaldo",
          },
          "#10b981"
        );
      }

      enviarCorreoEstilizado(
        CORREO_REPRESENTANTE_LEGAL,
        "Notificación Permiso Médico - Sindicato SLIM n°3",
        "Nueva Solicitud de Permiso Médico",
        `Se ha registrado una solicitud de permiso médico para el trabajador <strong>${beneficiario.nombre}</strong>.`,
        {
          ID: idUnico,
          Trabajador: beneficiario.nombre,
          RUT: beneficiario.rut,
          Tipo: tipoPermiso,
          "Fecha Inicio": fechaInicio,
          Motivo: motivo,
          "Fecha Solicitud": fechaHoy.toLocaleDateString(),
        },
        "#475569"
      );

      if (
        gestion === "Dirigente" &&
        correoDirigente &&
        correoDirigente.includes("@") &&
        correoDirigente !== beneficiario.correo
      ) {
        enviarCorreoEstilizado(
          correoDirigente,
          "Respaldo Gestión Permiso Médico - Sindicato SLIM n°3",
          "Permiso Médico Ingresado",
          `Has ingresado exitosamente un permiso médico para el socio <strong>${beneficiario.nombre}</strong>.`,
          { ID: idUnico, Socio: beneficiario.nombre, Tipo: tipoPermiso },
          "#475569"
        );
      }

      return {
        success: true,
        message:
          "Permiso médico solicitado. No olvides adjuntar el documento de respaldo.",
      };
    } catch (e) {
      return { success: false, message: "Error: " + e.toString() };
    } finally {
      lock.releaseLock();
    }
  } else {
    return { success: false, message: "Servidor ocupado." };
  }
}

function adjuntarDocumentoPermiso(idPermiso, archivoData) {
  const CARPETA_ID = "1VBzAXMnKzZ-RdO0Nzyqz4EblKDB9T3Ka";
  const CORREO_REPRESENTANTE_LEGAL = "penailillo.fetrasiss@gmail.com"; // CAMBIAR POR CORREO REAL

  const lock = LockService.getScriptLock();
  if (lock.tryLock(30000)) {
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheetPermisos = ss.getSheetByName("PERMISO-MEDICO");
      if (!sheetPermisos)
        return { success: false, message: "Hoja no encontrada." };

      const data = sheetPermisos.getDataRange().getValues();
      let rowIndex = -1;
      let beneficiario = null;
      let tipoPermiso = "";

      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(idPermiso)) {
          rowIndex = i + 1;
          beneficiario = {
            nombre: data[i][3],
            correo: data[i][4],
            rut: data[i][2],
          };
          tipoPermiso = data[i][5];
          break;
        }
      }

      if (rowIndex === -1)
        return { success: false, message: "Permiso no encontrado." };

      const sizeInBytes = (archivoData.base64.length * 3) / 4;
      if (sizeInBytes > 5 * 1024 * 1024) {
        return {
          success: false,
          message: "El archivo es demasiado grande (máx 5MB).",
        };
      }

      let fileUrl = "Sin archivo";

      try {
        const folder = DriveApp.getFolderById(CARPETA_ID);
        const blob = Utilities.newBlob(
          Utilities.base64Decode(archivoData.base64),
          archivoData.mimeType,
          archivoData.fileName
        );

        let extension = "";
        const nameParts = archivoData.fileName.split(".");
        if (nameParts.length > 1) extension = "." + nameParts.pop();

        blob.setName(
          `PERMISO-${idPermiso}-${cleanRut(beneficiario.rut)}${extension}`
        );

        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);

        // Otorgar permisos al socio
        if (beneficiario.correo && beneficiario.correo.includes("@")) {
          try {
            file.addViewer(beneficiario.correo);
          } catch (permErr) {
            console.log("Permiso drive error socio: " + permErr);
          }
        }

        // Otorgar permisos al representante legal
        if (
          CORREO_REPRESENTANTE_LEGAL &&
          CORREO_REPRESENTANTE_LEGAL.includes("@")
        ) {
          try {
            file.addViewer(CORREO_REPRESENTANTE_LEGAL);
          } catch (permErr) {
            console.log("Permiso drive error rep legal: " + permErr);
          }
        }

        fileUrl = file.getUrl();
      } catch (errDrive) {
        console.error("Error Drive: " + errDrive);
        return {
          success: false,
          message: "Error al subir archivo: " + errDrive.toString(),
        };
      }

      const fechaSubida = new Date();
      const nuevoEstado = "Documento Adjuntado";

      sheetPermisos.getRange(rowIndex, 9).setValue(fileUrl); // URL Documento
      sheetPermisos.getRange(rowIndex, 10).setValue(nuevoEstado); // Estado
      sheetPermisos.getRange(rowIndex, 11).setValue(fechaSubida); // Fecha Subida
      sheetPermisos.getRange(rowIndex, 12).setValue(false); // Notificado Rep. Legal (se notificará ahora)

      // Notificar al socio
      if (beneficiario.correo && beneficiario.correo.includes("@")) {
        enviarCorreoEstilizado(
          beneficiario.correo,
          "Documento Adjuntado - Sindicato SLIM n°3",
          "Documento de Permiso Médico Adjuntado",
          `Hola ${beneficiario.nombre}, tu documento de respaldo ha sido adjuntado exitosamente.`,
          {
            ID: idPermiso,
            "Tipo Permiso": tipoPermiso,
            Estado: nuevoEstado,
            Documento: "Adjuntado correctamente",
          },
          "#10b981"
        );
      }

      // Notificar al representante legal CON ENLACE AL DOCUMENTO
      enviarCorreoEstilizado(
        CORREO_REPRESENTANTE_LEGAL,
        "Documento Permiso Médico Adjuntado - Sindicato SLIM n°3",
        "Documento de Permiso Médico Disponible",
        `El trabajador <strong>${beneficiario.nombre}</strong> ha adjuntado el documento de respaldo para su permiso médico.`,
        {
          ID: idPermiso,
          Trabajador: beneficiario.nombre,
          RUT: beneficiario.rut,
          "Tipo Permiso": tipoPermiso,
          Documento: `<a href="${fileUrl}" style="color: #10b981; font-weight: bold;">Disponible para revisión</a>`,
          "Fecha Adjunto": fechaSubida.toLocaleDateString(),
        },
        "#475569"
      );

      // Marcar como notificado
      sheetPermisos.getRange(rowIndex, 12).setValue(true);

      return {
        success: true,
        message: "Documento adjuntado y notificaciones enviadas.",
      };
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
    const sheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PERMISO-MEDICO");
    if (!sheet) return { success: true, registros: [] };

    const data = sheet.getDataRange().getDisplayValues();
    const rutLimpio = cleanRut(rutInput);
    const registros = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (cleanRut(row[2]) === rutLimpio) {
        registros.push({
          id: row[0],
          fecha: row[1],
          tipoPermiso: row[5],
          fechaInicio: row[6],
          motivo: row[7],
          urlDocumento: row[8],
          estado: row[9],
          gestion: row[12],
          nomDirigente: row[13],
        });
      }
    }

    registros.reverse();
    return { success: true, registros: registros };
  } catch (e) {
    return { success: false, message: "Error: " + e.toString() };
  }
}

function eliminarPermisoMedico(idPermiso) {
  const CORREO_REPRESENTANTE_LEGAL = "penailillo.fetrasiss@gmail.com"; // CAMBIAR POR CORREO REAL

  const lock = LockService.getScriptLock();
  if (lock.tryLock(10000)) {
    try {
      const sheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PERMISO-MEDICO");
      if (!sheet) return { success: false, message: "Hoja no encontrada." };

      const data = sheet.getDataRange().getValues();

      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(idPermiso)) {
          const estado = String(data[i][9]);

          if (estado !== "Solicitado") {
            return {
              success: false,
              message: "Solo se pueden anular permisos en estado 'Solicitado'.",
            };
          }

          const beneficiario = {
            nombre: data[i][3],
            correo: data[i][4],
            rut: data[i][2],
          };
          const tipoPermiso = data[i][5];
          const fechaInicio = data[i][6];

          // Notificar al socio
          if (beneficiario.correo && beneficiario.correo.includes("@")) {
            enviarCorreoEstilizado(
              beneficiario.correo,
              "Permiso Médico Anulado - Sindicato SLIM n°3",
              "Solicitud de Permiso Anulada",
              `Hola ${beneficiario.nombre}, tu solicitud de permiso médico ha sido anulada. No se hará uso de este permiso.`,
              {
                ID: idPermiso,
                "Tipo Permiso": tipoPermiso,
                "Fecha Inicio": fechaInicio,
                Estado: "Anulado",
                Acción: "Solicitud eliminada del sistema",
              },
              "#ef4444"
            );
          }

          // Notificar al representante legal
          enviarCorreoEstilizado(
            CORREO_REPRESENTANTE_LEGAL,
            "Permiso Médico Anulado - Sindicato SLIM n°3",
            "Solicitud de Permiso Anulada",
            `La solicitud de permiso médico del trabajador <strong>${beneficiario.nombre}</strong> ha sido anulada. No se hará uso de este permiso.`,
            {
              ID: idPermiso,
              Trabajador: beneficiario.nombre,
              RUT: beneficiario.rut,
              "Tipo Permiso": tipoPermiso,
              "Fecha Inicio": fechaInicio,
              Estado: "Anulado por el usuario",
            },
            "#475569"
          );

          sheet.deleteRow(i + 1);
          return {
            success: true,
            message: "Permiso anulado y notificaciones enviadas.",
          };
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

// ==========================================
// MÓDULO: REGISTRO ASISTENCIA
// ==========================================

function obtenerHistorialAsistencia(rutInput) {
  try {
    const sheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ASISTENCIA");
    if (!sheet) return { success: true, registros: [] };

    const data = sheet.getDataRange().getDisplayValues();
    const rutLimpio = cleanRut(rutInput);
    const registros = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (cleanRut(row[0]) === rutLimpio) {
        registros.push({
          asamblea: row[1] || "Asamblea",
          tipo: row[2] || "Tipo no especificado",
          gestion: row[3] || "Sistema",
          dirigente: row[4] || "",
        });
      }
    }

    registros.reverse();
    return { success: true, registros: registros };
  } catch (e) {
    return { success: false, message: "Error: " + e.toString() };
  }
}

// ==========================================
// GESTIÓN SOCIOS - PARA DIRIGENTES
// ==========================================

function obtenerGestionesDirigente(rutDirigente) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rutLimpio = cleanRut(rutDirigente);

    const resultado = {
      prestamos: [],
      justificaciones: [],
      apelaciones: [],
      permisosMedicos: [],
    };

    // PRÉSTAMOS
    const sheetPrestamos = ss.getSheetByName("PRESTAMOS");
    if (sheetPrestamos) {
      const dataPrestamos = sheetPrestamos.getDataRange().getDisplayValues();
      for (let i = 1; i < dataPrestamos.length; i++) {
        const row = dataPrestamos[i];
        const gestion = row[10];

        if (gestion === "Dirigente") {
          resultado.prestamos.push({
            id: row[0],
            fecha: row[1],
            rutSocio: row[2],
            nombreSocio: row[3],
            tipo: row[5],
            cuotas: row[6],
            medio: row[7],
            estado: row[8],
            fechaTermino: row[9],
          });
        }
      }
    }

    // JUSTIFICACIONES
    const sheetJustif = ss.getSheetByName("JUSTIFICACIONES");
    if (sheetJustif) {
      const dataJustif = sheetJustif.getDataRange().getDisplayValues();
      for (let i = 1; i < dataJustif.length; i++) {
        const row = dataJustif[i];
        const gestion = row[12];

        if (gestion === "Dirigente") {
          resultado.justificaciones.push({
            id: row[0],
            fecha: row[1],
            rutSocio: row[2],
            nombreSocio: row[3],
            tipo: row[5],
            motivo: row[6],
            url: row[7],
            estado: row[8],
            obs: row[9],
            asamblea: row[11],
          });
        }
      }
    }

    // APELACIONES
    const sheetApel = ss.getSheetByName("APELACIONES");
    if (sheetApel) {
      const dataApel = sheetApel.getDataRange().getDisplayValues();
      for (let i = 1; i < dataApel.length; i++) {
        const row = dataApel[i];
        const gestion = row[13];

        if (gestion === "Dirigente") {
          resultado.apelaciones.push({
            id: row[0],
            fecha: row[1],
            rutSocio: row[2],
            nombreSocio: row[3],
            mesApelacion: row[5],
            tipoMotivo: row[6],
            detalleMotivo: row[7],
            urlComprobante: row[8],
            urlLiquidacion: row[9],
            estado: row[10],
            obs: row[11],
            urlComprobanteDevolucion: row[16] || "",
          });
        }
      }
    }

    // PERMISOS MÉDICOS
    const sheetPermisos = ss.getSheetByName("PERMISO-MEDICO");
    if (sheetPermisos) {
      const dataPermisos = sheetPermisos.getDataRange().getDisplayValues();
      for (let i = 1; i < dataPermisos.length; i++) {
        const row = dataPermisos[i];
        const gestion = row[12];

        if (gestion === "Dirigente") {
          resultado.permisosMedicos.push({
            id: row[0],
            fecha: row[1],
            rutSocio: row[2],
            nombreSocio: row[3],
            tipoPermiso: row[5],
            fechaInicio: row[6],
            motivo: row[7],
            urlDocumento: row[8],
            estado: row[9],
          });
        }
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
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetPrestamos = ss.getSheetByName("PRESTAMOS");

    if (!sheetPrestamos) {
      return { success: false, message: "No se encontró la hoja PRESTAMOS." };
    }

    const data = sheetPrestamos.getDataRange().getValues();

    // Filtrar solo préstamos en estado "Solicitado"
    const prestamosSolicitados = [];
    const filasActualizar = [];

    for (let i = 1; i < data.length; i++) {
      const estado = String(data[i][8]); // Columna Estado (índice 8)

      if (estado === "Solicitado") {
        prestamosSolicitados.push({
          rut: data[i][2], // Columna C: RUT
          nombre: data[i][3], // Columna D: NOMBRE
          tipoPrestamo: data[i][5], // Columna F: TIPO PRÉSTAMO
          cuotas: data[i][6], // Columna G: CUOTAS
          medioPago: data[i][7], // Columna H: MEDIO PAGO
        });
        filasActualizar.push(i + 1); // Guardar índice de fila (base 1)
      }
    }

    if (prestamosSolicitados.length === 0) {
      return {
        success: false,
        message: "No hay préstamos en estado 'Solicitado' para procesar.",
      };
    }

    // Crear nueva hoja temporal para el informe
    let sheetInforme = ss.getSheetByName("INFORME_PRESTAMOS_TEMP");
    if (sheetInforme) {
      ss.deleteSheet(sheetInforme);
    }
    sheetInforme = ss.insertSheet("INFORME_PRESTAMOS_TEMP");

    // Encabezados
    const headers = ["RUT", "NOMBRE", "TIPO PRÉSTAMO", "CUOTAS", "MEDIO PAGO"];
    sheetInforme.appendRow(headers);

    // Agregar datos filtrados
    prestamosSolicitados.forEach((prestamo) => {
      sheetInforme.appendRow([
        prestamo.rut,
        prestamo.nombre,
        prestamo.tipoPrestamo,
        prestamo.cuotas,
        prestamo.medioPago,
      ]);
    });

    // Formatear
    const lastRow = sheetInforme.getLastRow();
    const lastCol = sheetInforme.getLastColumn();

    sheetInforme
      .getRange(1, 1, 1, lastCol)
      .setFontWeight("bold")
      .setBackground("#4c1d95")
      .setFontColor("#ffffff");

    sheetInforme.setFrozenRows(1);
    sheetInforme.autoResizeColumns(1, lastCol);

    // Convertir a Excel (blob)
    const url =
      "https://docs.google.com/spreadsheets/d/" +
      ss.getId() +
      "/export?format=xlsx&gid=" +
      sheetInforme.getSheetId();
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, {
      headers: {
        Authorization: "Bearer " + token,
      },
    });

    const blob = response.getBlob();
    blob.setName(
      `Informe_Prestamos_Solicitados_${new Date()
        .toLocaleDateString("es-CL")
        .replace(/\//g, "-")}.xlsx`
    );

    // Obtener correo del administrador desde BD_USUARIOS
    const sheetUsers = ss.getSheetByName("BD_USUARIOS");
    const dataUsers = sheetUsers.getDataRange().getDisplayValues();
    let correoAdmin = "admin@sindicato.com"; // Correo por defecto

    for (let i = 1; i < dataUsers.length; i++) {
      const rol = String(dataUsers[i][14]).toUpperCase();
      if (rol === "ADMIN") {
        correoAdmin = dataUsers[i][5]; // Columna F: CORREO
        break;
      }
    }

    // Enviar correo con archivo adjunto
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
              Se adjunta el informe de préstamos en estado <strong>"Solicitado"</strong> generado el <strong>${new Date().toLocaleDateString(
                "es-CL"
              )}</strong>.
            </p>
            <p style="color: #64748b; font-size: 14px; margin-top: 20px;">
              Total de préstamos procesados: <strong>${
                prestamosSolicitados.length
              }</strong>
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
      attachments: [blob],
    });

    // Cambiar estado de "Solicitado" a "Enviado"
    filasActualizar.forEach((fila) => {
      sheetPrestamos.getRange(fila, 9).setValue("Enviado"); // Columna I: Estado
    });

    // Eliminar hoja temporal
    ss.deleteSheet(sheetInforme);

    return {
      success: true,
      message: `Informe generado y enviado. ${prestamosSolicitados.length} préstamo(s) cambiado(s) a "Enviado".`,
    };
  } catch (e) {
    return {
      success: false,
      message: "Error al generar informe: " + e.toString(),
    };
  }
}

// ==========================================
// FUNCIONES AUXILIARES
// ==========================================

function cleanRut(rut) {
  if (!rut) return "";
  return String(rut).replace(/\./g, "").replace(/-/g, "").toUpperCase().trim();
}

function enviarCorreoEstilizado(
  destinatario,
  asunto,
  titulo,
  mensaje,
  detalles,
  colorTema
) {
  try {
    if (!destinatario || !destinatario.includes("@")) {
      console.log("Correo inválido: " + destinatario);
      return;
    }

    let detallesHtml = "";
    if (detalles && typeof detalles === "object") {
      detallesHtml =
        "<table style='width: 100%; border-collapse: collapse; margin-top: 20px;'>";
      for (let key in detalles) {
        let valor = detalles[key];

        // Si el valor contiene HTML (como enlaces), no escaparlo
        if (
          typeof valor === "string" &&
          (valor.includes("<a href=") ||
            valor.includes("<br>") ||
            valor.includes("<em>") ||
            valor.includes("<strong>"))
        ) {
          // Ya contiene HTML, usar tal cual
        } else {
          // Escapar caracteres especiales para valores normales
          valor = String(valor).replace(/</g, "&lt;").replace(/>/g, "&gt;");
        }

        detallesHtml += `
          <tr>
            <td style='padding: 12px; border-bottom: 1px solid #e2e8f0; color: #64748b; font-weight: 600; width: 40%;'>${key}</td>
            <td style='padding: 12px; border-bottom: 1px solid #e2e8f0; color: #1e293b; font-weight: 500;'>${valor}</td>
          </tr>
        `;
      }
      detallesHtml += "</table>";
    }

    const htmlBody = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
      </head>
      <body style="margin: 0; padding: 0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f1f5f9;">
        <div style="max-width: 600px; margin: 40px auto; background: white; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
          
          <div style="background: linear-gradient(135deg, ${colorTema} 0%, ${adjustColor(
      colorTema,
      -20
    )} 100%); color: white; padding: 40px 30px; text-align: center;">
            <h1 style="margin: 0; font-size: 28px; font-weight: 700; letter-spacing: -0.5px;">${titulo}</h1>
          </div>
          
          <div style="padding: 40px 30px;">
            <p style="color: #1e293b; font-size: 16px; line-height: 1.8; margin: 0 0 20px 0;">
              ${mensaje}
            </p>
            
            ${detallesHtml}
            
            <div style="margin-top: 30px; padding-top: 30px; border-top: 2px solid #f1f5f9;">
              <p style="color: #64748b; font-size: 14px; line-height: 1.6; margin: 0;">
                <strong style="color: #1e293b;">Importante:</strong> Este es un correo automático del sistema de gestión del Sindicato SLIM n°3. 
                Por favor no respondas a este correo.
              </p>
            </div>
          </div>
          
          <div style="background: #f8fafc; padding: 20px 30px; text-align: center; border-top: 1px solid #e2e8f0;">
            <p style="color: #94a3b8; font-size: 12px; margin: 0;">
              © ${new Date().getFullYear()} Sindicato SLIM n°3 - Sistema de Gestión de Socios
            </p>
          </div>
          
        </div>
      </body>
      </html>
    `;

    MailApp.sendEmail({
      to: destinatario,
      subject: asunto,
      htmlBody: htmlBody,
    });
  } catch (e) {
    console.error(
      "Error enviando correo a " + destinatario + ": " + e.toString()
    );
  }
}

function adjustColor(hexColor, percent) {
  const num = parseInt(hexColor.replace("#", ""), 16);
  const amt = Math.round(2.55 * percent);
  const R = (num >> 16) + amt;
  const G = ((num >> 8) & 0x00ff) + amt;
  const B = (num & 0x0000ff) + amt;
  return (
    "#" +
    (
      0x1000000 +
      (R < 255 ? (R < 1 ? 0 : R) : 255) * 0x10000 +
      (G < 255 ? (G < 1 ? 0 : G) : 255) * 0x100 +
      (B < 255 ? (B < 1 ? 0 : B) : 255)
    )
      .toString(16)
      .slice(1)
  );
}

// ==========================================
// TRIGGERS Y AUTOMATIZACIONES
// ==========================================

function verificarCambiosPrestamos() {
  try {
    const sheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PRESTAMOS");
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const estado = String(row[8]);
      const fechaTermino = new Date(row[9]);
      const hoy = new Date();

      // Si el préstamo está "Vigente" y ya pasó la fecha de término, cambiar a "Pagado"
      if (estado === "Vigente" && hoy > fechaTermino) {
        sheet.getRange(i + 1, 9).setValue("Pagado");

        const correo = row[4];
        const nombre = row[3];
        const tipo = row[5];

        if (correo && correo.includes("@")) {
          enviarCorreoEstilizado(
            correo,
            "Préstamo Completado - Sindicato SLIM n°3",
            "Préstamo Finalizado",
            `Hola ${nombre}, tu préstamo de tipo <strong>${tipo}</strong> ha sido marcado como pagado.`,
            {
              ID: row[0],
              Tipo: tipo,
              Estado: "Pagado",
              "Fecha Término": fechaTermino.toLocaleDateString(),
            },
            "#10b981"
          );
        }
      }
    }
  } catch (e) {
    console.error("Error verificando préstamos: " + e);
  }
}

function verificarCambiosJustificaciones() {
  try {
    const sheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("JUSTIFICACIONES");
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
              ID: idRegistro,
              Tipo: tipo,
              "Nuevo Estado": estadoActual,
              Observación: obs || "Sin observaciones",
              Asamblea: asamblea || "Pendiente asignación",
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

// Las funciones verificarCambiosApelaciones y procesarPermisosComprobantesDevolucion
// ya están definidas arriba en la sección de apelaciones

// ==========================================
// MÓDULO: SISTEMA QR DE ASISTENCIA
// ==========================================

/**
 * Valida si un usuario existe y está ACTIVO (con caché)
 */
function validarUsuarioQR(rut) {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = `user_${cleanRut(rut)}`;

    // Intentar obtener del caché (válido por 10 minutos)
    const cached = cache.get(cacheKey);
    if (cached) {
      const userData = JSON.parse(cached);
      Logger.log(`✅ Usuario ${rut} obtenido desde caché`);
      return userData;
    }

    // Si no está en caché, buscar en hoja
    const sheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BD_USUARIOS");
    if (!sheet)
      return { success: false, error: "Error: Base de datos no encontrada." };

    const data = sheet.getDataRange().getDisplayValues();
    const rutLimpio = cleanRut(rut);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (cleanRut(row[0]) === rutLimpio) {
        const estadoUsuario = String(row[9]).toUpperCase();
        const nombreUsuario = row[3];

        const resultado =
          estadoUsuario === "ACTIVO"
            ? { success: true, nombre: nombreUsuario }
            : {
                success: false,
                error: "Usuario inactivo. Contacta con la directiva.",
              };

        // Guardar en caché solo si está activo
        if (resultado.success) {
          cache.put(cacheKey, JSON.stringify(resultado), 600); // 10 minutos
          Logger.log(`💾 Usuario ${rut} guardado en caché`);
        }

        return resultado;
      }
    }

    return { success: false, error: "RUT no encontrado en el sistema." };
  } catch (e) {
    return { success: false, error: "Error del servidor: " + e.toString() };
  }
}

/**
 * Registra asistencia mediante QR con sistema de bloqueo robusto
 */
function checkinQR(rut, asamblea) {
  const MAX_INTENTOS = 3;
  const TIMEOUT_LOCK = 30000; // 30 segundos

  let intentoActual = 0;
  let ultimoError = null;

  // Reintentos automáticos en caso de conflicto
  while (intentoActual < MAX_INTENTOS) {
    intentoActual++;

    const lock = LockService.getScriptLock();

    try {
      // Intentar obtener el bloqueo con timeout
      const lockObtenido = lock.tryLock(TIMEOUT_LOCK);

      if (!lockObtenido) {
        ultimoError = `Intento ${intentoActual}/${MAX_INTENTOS}: Sistema ocupado`;
        Logger.log(ultimoError);

        // Esperar antes de reintentar (backoff exponencial)
        Utilities.sleep(Math.pow(2, intentoActual) * 100);
        continue;
      }

      // ✅ BLOQUEO OBTENIDO - Ejecutar operación
      try {
        const validacion = validarUsuarioQR(rut);

        if (!validacion.success) {
          throw new Error(validacion.error);
        }

        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName("ASISTENCIA");

        if (!sheet) {
          throw new Error("Error: Hoja ASISTENCIA no encontrada.");
        }

        const now = new Date();
        const rutFormateado = rut;
        const nombreUsuario = validacion.nombre;
        const gestion = "Socio";
        const dirigente = "";

        // Verificar duplicados en los últimos 5 minutos
        const data = sheet.getDataRange().getValues();
        const cincMinutosAtras = new Date(now.getTime() - 5 * 60 * 1000);

        for (let i = data.length - 1; i >= Math.max(0, data.length - 50); i--) {
          if (cleanRut(String(data[i][0])) === cleanRut(rut)) {
            // Verificar si el registro es del mismo evento (asamblea)
            if (String(data[i][2]) === asamblea) {
              Logger.log(`⚠️ Usuario ${rut} ya registrado en ${asamblea}`);
              throw new Error("Ya registraste tu asistencia para este evento.");
            }
          }
        }

        // Registrar asistencia
        sheet.appendRow([
          rutFormateado,
          nombreUsuario,
          asamblea,
          "Asistencia Presencial",
          gestion,
          dirigente,
        ]);

        const fechaFormateada = Utilities.formatDate(
          now,
          "America/Santiago",
          "dd/MM/yyyy HH:mm:ss"
        );

        Logger.log(
          `✅ Asistencia registrada: ${nombreUsuario} (${rut}) - ${asamblea} - ${fechaFormateada}`
        );

        // Liberar bloqueo antes de retornar
        lock.releaseLock();

        return {
          rut: rutFormateado,
          fecha: fechaFormateada,
          nombre: nombreUsuario,
          mensaje: "Asistencia registrada exitosamente",
        };
      } catch (innerError) {
        // Liberar bloqueo en caso de error
        lock.releaseLock();
        throw innerError;
      }
    } catch (e) {
      ultimoError = e.message || e.toString();
      Logger.log(`❌ Error intento ${intentoActual}: ${ultimoError}`);

      // Si es el último intento, lanzar el error
      if (intentoActual >= MAX_INTENTOS) {
        // Registrar fallo en hoja de log (opcional)
        registrarErrorAsistencia(rut, asamblea, ultimoError);
        throw new Error(ultimoError);
      }

      // Esperar antes de reintentar
      Utilities.sleep(Math.pow(2, intentoActual) * 100);
    }
  }

  // Si se agotaron los intentos
  throw new Error(
    `No se pudo registrar después de ${MAX_INTENTOS} intentos. Intenta nuevamente.`
  );
}

/**
 * Registra errores de asistencia para auditoría
 */
function registrarErrorAsistencia(rut, asamblea, error) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetLog = ss.getSheetByName("LOG_ERRORES_QR");

    // Crear hoja de log si no existe
    if (!sheetLog) {
      sheetLog = ss.insertSheet("LOG_ERRORES_QR");
      sheetLog.appendRow(["Fecha", "RUT", "Asamblea", "Error", "Reintentado"]);
      sheetLog
        .getRange(1, 1, 1, 5)
        .setFontWeight("bold")
        .setBackground("#dc2626")
        .setFontColor("#ffffff");
    }

    const now = new Date();
    const fechaFormateada = Utilities.formatDate(
      now,
      "America/Santiago",
      "dd/MM/yyyy HH:mm:ss"
    );

    sheetLog.appendRow([fechaFormateada, rut, asamblea, error, "Sí"]);

    Logger.log(`📝 Error registrado en LOG: ${rut} - ${error}`);
  } catch (logError) {
    Logger.log(`⚠️ No se pudo registrar error en log: ${logError.toString()}`);
  }
}

/**
 * Genera Link de Registro QR para un usuario
 */
function generarLinkRegistroQR(rut) {
  try {
    let webAppUrl = ScriptApp.getService().getUrl();

    // LIMPIAR Y CORREGIR URL
    webAppUrl = webAppUrl.replace(/\/a\/[^\/]+\//, "/");
    webAppUrl = webAppUrl.replace("/macros/", "/a/~/macros/");

    const rutLimpio = cleanRut(rut);
    const linkRegistro = `${webAppUrl}?page=qr&action=register&rut=${rutLimpio}`;

    return { success: true, link: linkRegistro };
  } catch (e) {
    return { success: false, message: "Error: " + e.toString() };
  }
}

/**
 * Genera Link de Punto de Control
 */
function generarLinkPuntoControl(nombrePunto, fecha, datosQR) {
  try {
    let webAppUrl = ScriptApp.getService().getUrl();

    // LIMPIAR Y CORREGIR URL
    webAppUrl = webAppUrl.replace(/\/a\/[^\/]+\//, "/");
    webAppUrl = webAppUrl.replace("/macros/", "/a/~/macros/");

    const asambleaId = `${fecha}-${datosQR}`;
    const linkPuntoControl = `${webAppUrl}?page=qr&action=checkin&asamblea=${encodeURIComponent(
      asambleaId
    )}`;

    return { success: true, link: linkPuntoControl };
  } catch (e) {
    return { success: false, message: "Error: " + e.toString() };
  }
}

/**
 * Obtener URL base de la WebApp
 */
function obtenerUrlWebApp() {
  try {
    let webAppUrl = ScriptApp.getService().getUrl();

    // LIMPIAR Y CORREGIR URL
    webAppUrl = webAppUrl.replace(/\/a\/[^\/]+\//, "/");
    webAppUrl = webAppUrl.replace("/macros/", "/a/~/macros/");

    return { success: true, url: webAppUrl };
  } catch (e) {
    return { success: false, message: "Error: " + e.toString() };
  }
}

/**
 * Generar QRs personales para TODOS los usuarios ACTIVOS
 */
function generarQRsUsuarios() {
  try {
    Logger.log("🚀 INICIANDO generarQRsUsuarios()");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log("✅ Spreadsheet obtenido: " + ss.getName());

    const sheet = ss.getSheetByName("BD_USUARIOS");

    if (!sheet) {
      Logger.log("❌ ERROR: No se encontró la hoja BD_USUARIOS");
      return;
    }

    Logger.log("✅ Hoja BD_USUARIOS encontrada");

    const data = sheet.getDataRange().getValues();
    Logger.log("📊 Total de filas en la hoja: " + data.length);

    // URL FIJA del deployment (exec, NO dev)
    const webAppUrl =
      "https://script.google.com/a/~/macros/s/AKfycbxUJWzgv-fA8XpTCkfBydGbiCikwC4ouJSnLJcuIcxfwwAb4J4B7W-02CXicX5GbSxY8w/exec";

    Logger.log("📍 URL base utilizada: " + webAppUrl);

    let contador = 0;
    let usuariosRevisados = 0;
    let usuariosActivos = 0;
    let usuariosInactivos = 0;

    // Limpiar columnas P y Q primero
    const lastRow = sheet.getLastRow();
    Logger.log("📏 Última fila con datos: " + lastRow);

    if (lastRow > 1) {
      Logger.log("🧹 Limpiando columnas P y Q (filas 2 a " + lastRow + ")");
      sheet.getRange(2, 16, lastRow - 1, 2).clearContent();
      Logger.log("✅ Limpieza completada");
    }

    for (let i = 1; i < data.length; i++) {
      usuariosRevisados++;

      const rut = data[i][0]; // Columna A
      const estado = String(data[i][9]).toUpperCase().trim(); // Columna J

      Logger.log(`\n--- Usuario ${i} ---`);
      Logger.log("RUT: " + rut);
      Logger.log("Estado (columna J): '" + estado + "'");
      Logger.log("Estado original: '" + data[i][9] + "'");

      if (estado === "ACTIVO" && rut) {
        usuariosActivos++;

        const rutLimpio = cleanRut(String(rut));
        Logger.log("RUT limpio: " + rutLimpio);

        // URL completa con page=qr
        const linkRegistro = `${webAppUrl}?page=qr&action=register&rut=${rutLimpio}`;
        Logger.log("Link generado: " + linkRegistro);

        // Fórmula para generar QR
        const formulaQR = `=IMAGE("https://quickchart.io/qr?size=200&text=" & ENCODEURL(P${
          i + 1
        }))`;
        Logger.log("Fórmula QR: " + formulaQR);

        // Escribir en columna P (Link Registro)
        Logger.log("Escribiendo en P" + (i + 1));
        sheet.getRange(i + 1, 16).setValue(linkRegistro);

        // Escribir en columna Q (Código QR)
        Logger.log("Escribiendo en Q" + (i + 1));
        sheet.getRange(i + 1, 17).setFormula(formulaQR);

        contador++;
        Logger.log("✅ QR generado exitosamente para fila " + (i + 1));
      } else {
        usuariosInactivos++;
        Logger.log("⏭️ Usuario omitido - Estado: " + estado + ", RUT: " + rut);
      }
    }

    Logger.log("\n📊 RESUMEN FINAL:");
    Logger.log("Total usuarios revisados: " + usuariosRevisados);
    Logger.log("Usuarios ACTIVOS encontrados: " + usuariosActivos);
    Logger.log("Usuarios INACTIVOS/sin RUT: " + usuariosInactivos);
    Logger.log("QRs generados exitosamente: " + contador);
    Logger.log(`✅ ÉXITO: Se generaron ${contador} códigos QR personales`);
    Logger.log(
      `📊 Ejemplo URL: ${webAppUrl}?page=qr&action=register&rut=166258375`
    );
  } catch (error) {
    Logger.log(`❌ ERROR CRÍTICO: ${error.toString()}`);
    Logger.log(`Stack trace: ${error.stack}`);
  }
}

// ==========================================
// CONFIGURAR TRIGGERS
// ==========================================

function configurarTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((trigger) => ScriptApp.deleteTrigger(trigger));

  ScriptApp.newTrigger("verificarCambiosPrestamos")
    .timeBased()
    .everyHours(1)
    .create();

  ScriptApp.newTrigger("verificarCambiosJustificaciones")
    .timeBased()
    .everyMinutes(30)
    .create();

  ScriptApp.newTrigger("verificarCambiosApelaciones")
    .timeBased()
    .everyMinutes(30)
    .create();

  ScriptApp.newTrigger("procesarPermisosComprobantesDevolucion")
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log("Triggers configurados exitosamente");
}
