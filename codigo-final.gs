const NOMBRE_CARPETA_IMAGENES = "ImagenesPacientes";
const NOMBRE_HOJA_IMAGENES = "ImagenesPacientes";
const HOJA_PACIENTES = "Hoja 1";
const HOJA_DOCUMENTOS = "Documentos";
const HOJA_DIAGNOSTICOS = "Diagnosticos";
const HOJA_LOGS = "Logs";
const DATOS_MEDICO = {
  nombre: "Dr. Nikolaus Vivallo",
  profesion: "Medico Cirujano",
  rut: "19.456.116-4",
  registro: "541933"
};

/**
 * ============================
 * POST
 * ============================
 */
function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var data = JSON.parse((e && e.postData && e.postData.contents) || "{}");

    var accion = txt(data.accion);
    var nombre = txt(data.nombre);

    escribirLog("ACCION DETECTADA: " + accion + " | NOMBRE: " + nombre);

    if (accion === "actualizarDiagnostico") {
      var hojaDiag = obtenerHojaObligatoria(ss, HOJA_DIAGNOSTICOS);

      var datosDiag = hojaDiag.getDataRange().getValues();
      var filaEncontrada = -1;

      for (var i = 1; i < datosDiag.length; i++) {
        if (txt(datosDiag[i][0]) === nombre) {
          filaEncontrada = i + 1;
          break;
        }
      }

      if (filaEncontrada !== -1) {
        hojaDiag.getRange(filaEncontrada, 2).setValue(data.diagnostico || "");
      } else {
        hojaDiag.appendRow([nombre, data.diagnostico || ""]);
      }

      SpreadsheetApp.flush();
      return ContentService.createTextOutput("Diagnostico actualizado");
    }

    if (accion === "editarFichaCompleta") {
      var hoja1 = obtenerHojaObligatoria(ss, HOJA_PACIENTES);

      var dFicha = hoja1.getDataRange().getValues();
      var rutOriginal = txt(data.rutOriginal);

      for (var j = 1; j < dFicha.length; j++) {
        var rutFila = txt(dFicha[j][2]);

        if (rutFila === rutOriginal) {
          hoja1.getRange(j + 1, 2).setValue(data.nombre || "");
          hoja1.getRange(j + 1, 3).setValue(data.rut || "");

          if (data.fechaNac) {
            hoja1.getRange(j + 1, 4).setValue(data.fechaNac);
          }

          hoja1.getRange(j + 1, 6).setValue(data.direccion || "");
          hoja1.getRange(j + 1, 7).setValue(data.telefono || "");
          hoja1.getRange(j + 1, 8).setValue(data.correo || "");
          hoja1.getRange(j + 1, 10).setValue(data.tratamiento_activo || data.tratamientoActivo || "");
        }
      }

      SpreadsheetApp.flush();
      return ContentService.createTextOutput("Ficha actualizada");
    }

    if (accion === "eliminarCompleto") {
      eliminarPacienteCompleto(data.nombre, data.rut);
      return ContentService.createTextOutput("Paciente eliminado completamente");
    }

    if (accion === "eliminar") {
      eliminarFila(data.nombre, data.rut);
      return ContentService.createTextOutput("Paciente eliminado");
    }

    if (accion === "eliminarNota") {
      eliminarNotaPaciente(data.nombre, data.fecha, data.tabla || "Evoluciones");
      return ContentService.createTextOutput("Nota eliminada");
    }

    if (accion === "alta") {
      eliminarPacienteCompleto(data.nombre, data.rut);
      return ContentService.createTextOutput("Paciente eliminado definitivamente");
    }

    if (accion === "guardarEvolucion") {
      var sheetEvolucion = obtenerHojaObligatoria(ss, HOJA_PACIENTES);

      var datosPaciente = buscarPacientePorNombre(nombre);

      sheetEvolucion.appendRow([
        new Date(),
        nombre,
        datosPaciente.rut,
        datosPaciente.fechaNacimiento,
        "",
        datosPaciente.direccion,
        datosPaciente.telefono,
        datosPaciente.correo,
        data.evolucion || "",
        datosPaciente.tratamientoActivo
      ]);

      SpreadsheetApp.flush();
      return ContentService.createTextOutput("Evolucion guardada");
    }

    if (accion === "guardarDocumento") {
      var hojaDocumentos = obtenerHojaObligatoria(ss, HOJA_DOCUMENTOS);

      var datosPacienteDocumento = buscarPacientePorNombre(nombre);

      hojaDocumentos.appendRow([
        new Date(),
        nombre,
        datosPacienteDocumento.rut,
        data.tipo || "Documento",
        data.contenido || ""
      ]);

      SpreadsheetApp.flush();
      return ContentService.createTextOutput("Documento guardado");
    }

    if (accion === "enviarDocumentoCorreo") {
      enviarDocumentoPorCorreo(data);
      return ContentService.createTextOutput("Correo enviado");
    }

    if (accion === "subirImagenPaciente") {
      var resultadoImagen = subirImagenPaciente(data);
      return ContentService
        .createTextOutput(JSON.stringify(resultadoImagen))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var sheetActiva = obtenerHojaObligatoria(ss, HOJA_PACIENTES);

    sheetActiva.appendRow([
      new Date(),
      nombre,
      data.rut || "",
      data.fecha_nacimiento || "",
      "",
      data.direccion || "",
      data.telefono || "",
      data.correo || "",
      data.evolucion || "",
      data.tratamiento_activo || data.tratamientoActivo || ""
    ]);

    SpreadsheetApp.flush();
    return ContentService.createTextOutput("Paciente guardado");

  } catch (error) {
    Logger.log("ERROR doPost: " + error.toString());
    escribirLog("ERROR doPost: " + error.toString());
    return ContentService.createTextOutput("Error: " + error.toString());
  }
}


/**
 * ============================
 * GET
 * ============================
 */
function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if (e.parameter.accion === "listarImagenesPaciente") {
      var resultado = listarImagenesPaciente(e.parameter.rut || "", e.parameter.nombre || "");
      return ContentService
        .createTextOutput(JSON.stringify(resultado))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var nombreHoja = e.parameter.hoja || e.parameter.tabla || HOJA_PACIENTES;
    var sheet = ss.getSheetByName(nombreHoja);

    if (!sheet) {
      return ContentService
        .createTextOutput(JSON.stringify({ error: "Hoja no existe: " + nombreHoja }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var data = sheet.getDataRange().getValues();

    return ContentService
      .createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
function buscarPacientePorNombre(nombre) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(HOJA_PACIENTES);
  var data = hoja.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (txt(data[i][1]) === txt(nombre)) {
      return {
        rut: data[i][2] || "",
        fechaNacimiento: data[i][3] || "",
        direccion: data[i][5] || "",
        telefono: data[i][6] || "",
        correo: data[i][7] || "",
        tratamientoActivo: data[i][9] || ""
      };
    }
  }

  return {
    rut: "",
    fechaNacimiento: "",
    direccion: "",
    telefono: "",
    correo: "",
    tratamientoActivo: ""
  };
}

function enviarDocumentoPorCorreo(data) {
  data = data || {};
  var nombre = txt(data.nombre);
  var correo = txt(data.correo);
  var tipo = txt(data.tipo) || "Documento medico";
  var rut = txt(data.rut);
  var edad = txt(data.edad);
  var fechaNacimiento = txt(data.fechaNacimiento);
  var direccion = txt(data.direccion);
  var contenido = txt(data.contenido);
  var htmlDocumento = data.htmlDocumento || "";

  if (!correo) throw new Error("Falta el correo del destinatario");
  if (!nombre) throw new Error("Falta el nombre del paciente");
  if (!contenido) throw new Error("Falta el contenido del documento");

  var fechaEmision = formatearFechaActual();
  var asunto = tipo + " - " + nombre;
  var documento = construirDocumentoMedico({
    tipo: tipo,
    nombre: nombre,
    rut: rut,
    edad: edad,
    fechaNacimiento: fechaNacimiento,
    direccion: direccion,
    contenido: contenido,
    fechaEmision: fechaEmision
  });

  var cuerpoPlano = "Estimado paciente, se le hace envio de la "
    + (tipo === "RECETA MEDICA" ? "receta medica" : "orden de examenes")
    + " pactada en archivo PDF adjunto.";

  MailApp.sendEmail({
    to: correo,
    subject: asunto,
    body: cuerpoPlano,
    attachments: [crearPdfDocumento(documento, asunto + ".pdf", htmlDocumento)]
  });
}

function autorizarServicios() {
  var correo = Session.getActiveUser().getEmail();
  if (!correo) {
    throw new Error("No se pudo detectar el correo de la cuenta activa para la prueba.");
  }

  enviarDocumentoPorCorreo({
    nombre: "Paciente Prueba",
    correo: correo,
    tipo: "RECETA MEDICA",
    rut: "11.111.111-1",
    edad: "30",
    fechaNacimiento: "01/01/1995",
    direccion: "Direccion de prueba",
    contenido: "Documento de prueba para autorizar MailApp, DriveApp y DocumentApp."
  });
}

function escaparHtml(texto) {
  return (texto || "")
    .toString()
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function txt(valor) {
  return (valor || "").toString().trim();
}

function obtenerHojaObligatoria(ss, nombreHoja) {
  var hoja = ss.getSheetByName(nombreHoja);
  if (!hoja) throw new Error("No existe la hoja '" + nombreHoja + "'");
  return hoja;
}

function formatearFechaActual() {
  return Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone() || "America/Santiago",
    "dd/MM/yyyy"
  );
}

function construirDocumentoMedico(datos) {
  var tipo = txt(datos.tipo) || "Documento medico";
  var nombre = txt(datos.nombre);
  var rut = txt(datos.rut) || "No registrado";
  var edad = txt(datos.edad) || "No registrada";
  var fechaNacimiento = txt(datos.fechaNacimiento) || "No registrada";
  var direccion = txt(datos.direccion) || "No registrada";
  var contenido = txt(datos.contenido);
  var fechaEmision = txt(datos.fechaEmision) || formatearFechaActual();
  var contenidoHtml = escaparHtml(contenido).replace(/\n/g, "<br>");

  var html = ''
    + '<html><head><meta charset="UTF-8"></head><body style="margin:0;padding:24px;font-family:Georgia,Times New Roman,serif;color:#243d34;">'
    + '<div style="max-width:780px;margin:0 auto;border:1px solid #dceee4;">'
    + '<div style="height:16px;background:linear-gradient(90deg,#6fbf9d 0%,#bde3cf 100%);"></div>'
    + '<div style="padding:24px;">'
    + '<div style="display:flex;justify-content:space-between;align-items:flex-start;border-bottom:2px solid #8cc7ab;padding-bottom:16px;margin-bottom:16px;">'
    + '<div><h2 style="margin:0;color:#2f6b58;">' + escaparHtml(DATOS_MEDICO.nombre) + '</h2><p style="margin:4px 0 0;color:#4e6a61;letter-spacing:1px;text-transform:uppercase;">' + escaparHtml(DATOS_MEDICO.profesion) + '</p></div>'
    + '<div style="font-size:32px;color:#6fbf9d;font-weight:bold;">+</div>'
    + '</div>'
    + '<div style="display:flex;justify-content:space-between;gap:12px;margin-bottom:16px;">'
    + '<h3 style="margin:0;color:#2f6b58;font-size:14px;letter-spacing:2px;text-transform:uppercase;">' + escaparHtml(tipo) + '</h3>'
    + '<div style="font-size:12px;color:#4e6a61;">Fecha: ' + escaparHtml(fechaEmision) + '</div>'
    + '</div>'
    + '<table style="width:100%;border-collapse:separate;border-spacing:0 8px;margin-bottom:20px;font-size:13px;">'
    + '<tr><td><strong>Paciente:</strong> ' + escaparHtml(nombre) + '</td><td><strong>RUT:</strong> ' + escaparHtml(rut) + '</td></tr>'
    + '<tr><td><strong>Fecha de nacimiento:</strong> ' + escaparHtml(fechaNacimiento) + '</td><td><strong>Edad:</strong> ' + escaparHtml(edad) + '</td></tr>'
    + '<tr><td colspan="2"><strong>Direccion:</strong> ' + escaparHtml(direccion) + '</td></tr>'
    + '</table>'
    + '<div style="display:flex;gap:12px;">'
    + '<div style="font-size:28px;color:#2f6b58;font-weight:bold;">' + (tipo === "RECETA MEDICA" ? "Rx" : "") + '</div>'
    + '<div style="font-size:16px;line-height:1.7;">' + contenidoHtml + '</div>'
    + '</div>'
    + '<div style="margin-top:32px;padding-top:16px;border-top:1px solid #d8e8df;display:flex;justify-content:space-between;gap:16px;">'
    + '<div style="font-size:12px;color:#4e6a61;">'
    + '<p style="margin:3px 0;">RUT: ' + escaparHtml(DATOS_MEDICO.rut) + '</p>'
    + '<p style="margin:3px 0;">Registro Nacional de Prestadores: ' + escaparHtml(DATOS_MEDICO.registro) + '</p>'
    + '</div>'
    + '<div style="width:42%;text-align:center;font-size:12px;color:#4e6a61;">'
    + '<div style="border-top:1px solid #8cc7ab;margin-bottom:6px;"></div>'
    + '<p style="margin:0;">' + escaparHtml(DATOS_MEDICO.nombre) + '</p>'
    + '<p style="margin:0;">Firma y timbre medico</p>'
    + '</div>'
    + '</div>'
    + '</div>'
    + '</div>'
    + '</body></html>';

  var textoPlano = [
    DATOS_MEDICO.nombre,
    DATOS_MEDICO.profesion,
    "",
    tipo,
    "Fecha: " + fechaEmision,
    "",
    "Paciente: " + nombre,
    "RUT: " + rut,
    "Fecha de nacimiento: " + fechaNacimiento,
    "Edad: " + edad,
    "Direccion: " + direccion,
    "",
    contenido,
    "",
    "RUT: " + DATOS_MEDICO.rut,
    "Registro Nacional de Prestadores: " + DATOS_MEDICO.registro
  ].join("\n");

  return {
    html: html,
    texto: textoPlano,
    tipo: tipo,
    nombre: nombre,
    rut: rut,
    edad: edad,
    fechaNacimiento: fechaNacimiento,
    direccion: direccion,
    contenido: contenido,
    fechaEmision: fechaEmision
  };
}

function crearPdfDocumento(documento, nombreArchivo, htmlDocumento) {
  if (txt(htmlDocumento)) {
    try {
      return HtmlService
        .createHtmlOutput(htmlDocumento)
        .getBlob()
        .getAs(MimeType.PDF)
        .setName(nombreArchivo);
    } catch (error) {
      escribirLog("WARNING PDF desde HTML: " + error.toString());
    }
  }

  var doc = DocumentApp.create(nombreArchivo.replace(/\.pdf$/i, ""));
  try {
    var body = doc.getBody();
    body.clear();

    var colorPrincipal = "#2f6b58";
    var colorTextoSecundario = "#4e6a61";
    var colorLinea = "#8cc7ab";

    body.appendParagraph("")
      .setBackgroundColor("#bde3cf")
      .setSpacingAfter(10);

    var tablaHeader = body.appendTable([
      [
        DATOS_MEDICO.nombre + "\n" + DATOS_MEDICO.profesion.toUpperCase(),
        "+"
      ]
    ]);
    tablaHeader.setBorderWidth(0);
    tablaHeader.getRow(0).getCell(0).setPaddingTop(6).setPaddingBottom(8);
    tablaHeader.getRow(0).getCell(1).setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
    tablaHeader.getRow(0).getCell(1).setPaddingLeft(12);

    var textoHeader = tablaHeader.getCell(0, 0).editAsText();
    textoHeader.setFontFamily("Georgia").setFontSize(16).setBold(true).setForegroundColor(colorPrincipal);
    textoHeader.setFontSize(DATOS_MEDICO.nombre.length + 1, textoHeader.getText().length - 1, 10);

    tablaHeader.getCell(0, 1).editAsText()
      .setFontFamily("Georgia")
      .setFontSize(24)
      .setBold(true)
      .setForegroundColor("#6fbf9d");

    body.appendHorizontalRule();
    body.appendParagraph("").setSpacingAfter(8);

    var tablaTitulo = body.appendTable([
      [documento.tipo, "Fecha: " + documento.fechaEmision]
    ]);
    tablaTitulo.setBorderWidth(0);
    tablaTitulo.getCell(0, 0).editAsText()
      .setFontFamily("Georgia")
      .setFontSize(12)
      .setBold(true)
      .setForegroundColor(colorPrincipal);
    tablaTitulo.getCell(0, 1).editAsText()
      .setFontFamily("Arial")
      .setFontSize(10)
      .setForegroundColor(colorTextoSecundario);

    body.appendParagraph("").setSpacingAfter(4);

    agregarLineaDocumento(body, "Paciente", documento.nombre, "RUT", documento.rut, colorLinea);
    agregarLineaDocumento(body, "Nacimiento", documento.fechaNacimiento, "Edad", documento.edad, colorLinea);
    agregarLineaDocumento(body, "Direccion", documento.direccion, "", "", colorLinea);

    body.appendParagraph("").setSpacingAfter(8);

    var esReceta = documento.tipo === "RECETA MEDICA";
    if (esReceta) {
      var lineasReceta = (documento.contenido || "").split("\n");
      var primeraLinea = lineasReceta.shift() || "";
      var parrafoReceta = body.appendParagraph("Rx  " + primeraLinea);
      var textoReceta = parrafoReceta.editAsText();
      textoReceta
        .setFontFamily("Georgia")
        .setFontSize(13)
        .setForegroundColor("#243d34");
      textoReceta
        .setBold(0, 1, true)
        .setFontSize(0, 1, 22)
        .setForegroundColor(0, 1, colorPrincipal);
      parrafoReceta.setSpacingAfter(2);

      for (var i = 0; i < lineasReceta.length; i++) {
        body.appendParagraph("    " + lineasReceta[i])
          .setFontFamily("Georgia")
          .setFontSize(13)
          .setForegroundColor("#243d34")
          .setSpacingAfter(2);
      }
    } else {
      var tituloExamen = body.appendParagraph("Examenes solicitados");
      tituloExamen.setFontFamily("Georgia")
        .setFontSize(15)
        .setBold(true)
        .setForegroundColor(colorPrincipal)
        .setSpacingAfter(8);

      var lineasExamenes = (documento.contenido || "").split("\n");
      for (var j = 0; j < lineasExamenes.length; j++) {
        body.appendParagraph(lineasExamenes[j])
          .setFontFamily("Georgia")
          .setFontSize(13)
          .setForegroundColor("#243d34")
          .setSpacingAfter(2);
      }
    }

    body.appendParagraph("").setSpacingAfter(22);
    body.appendHorizontalRule();
    body.appendParagraph("").setSpacingAfter(8);

    var tablaFooter = body.appendTable([
      [
        "RUT: " + DATOS_MEDICO.rut + "\nRegistro Nacional de Prestadores: " + DATOS_MEDICO.registro,
        DATOS_MEDICO.nombre + "\nFirma y timbre medico"
      ]
    ]);
    tablaFooter.setBorderWidth(0);
    tablaFooter.getRow(0).getCell(1).setPaddingLeft(20);
    tablaFooter.getCell(0, 0).editAsText()
      .setFontFamily("Arial")
      .setFontSize(9)
      .setForegroundColor(colorTextoSecundario);
    tablaFooter.getCell(0, 1).editAsText()
      .setFontFamily("Arial")
      .setFontSize(9)
      .setForegroundColor(colorTextoSecundario);

    doc.saveAndClose();

    var pdf = DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF);
    pdf.setName(nombreArchivo);
    return pdf;
  } finally {
    try {
      DriveApp.getFileById(doc.getId()).setTrashed(true);
    } catch (error) {
      escribirLog("WARNING limpieza PDF temporal: " + error.toString());
    }
  }
}

function agregarLineaDocumento(body, etiqueta1, valor1, etiqueta2, valor2, colorLinea) {
  var texto = etiqueta1 + "  " + valor1;
  if (etiqueta2 || valor2) {
    texto += "          " + etiqueta2 + "  " + valor2;
  }

  var parrafo = body.appendParagraph(texto);
  parrafo
    .setFontFamily("Georgia")
    .setFontSize(10)
    .setForegroundColor("#243d34")
    .setSpacingAfter(3)
    .setBorderBottom(true);

  try {
    parrafo.setBorderColor(colorLinea);
  } catch (error) {
    // Algunos entornos de Apps Script no exponen este setter; en ese caso usamos el borde por defecto.
  }
}

function eliminarFila(nombre, rut) {
  var sheet = obtenerHojaObligatoria(SpreadsheetApp.getActiveSpreadsheet(), HOJA_PACIENTES);

  var data = sheet.getDataRange().getValues();
  var nombreBuscado = txt(nombre);
  var rutBuscado = txt(rut);

  for (var i = data.length - 1; i >= 1; i--) {
    var nombreFila = txt(data[i][1]);
    var rutFila = txt(data[i][2]);

    if (nombreFila === nombreBuscado && rutFila === rutBuscado) {
      sheet.deleteRow(i + 1);
      break;
    }
  }

  SpreadsheetApp.flush();
}

function eliminarNotaPaciente(nombre, fecha, tabla) {
  var nombreBuscado = txt(nombre);
  var fechaBuscada = fecha ? new Date(fecha).toISOString() : "";
  var sheetName = tabla === HOJA_DOCUMENTOS ? HOJA_DOCUMENTOS : HOJA_PACIENTES;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

  if (!sheet) throw new Error("No existe la hoja '" + sheetName + "'");

  var data = sheet.getDataRange().getValues();

  for (var i = data.length - 1; i >= 1; i--) {
    var fechaFila = data[i][0] ? new Date(data[i][0]).toISOString() : "";
    var nombreFila = txt(data[i][1]);

    if (nombreFila === nombreBuscado && fechaFila === fechaBuscada) {
      sheet.deleteRow(i + 1);
      break;
    }
  }

  SpreadsheetApp.flush();
}

/**
 * ============================
 * IMAGENES PACIENTE - VERSION FINAL
 * ============================
 */
function subirImagenPaciente(data) {
  var nombre = txt(data.nombre);
  var rut = txt(data.rut);
  var descripcion = txt(data.descripcion);
  var mimeType = (data.mimeType || "application/octet-stream").toString();
  var fileBase64 = txt(data.fileBase64);
  var fileNameOriginal = txt(data.fileName) || "imagen";

  if (!rut) throw new Error("Falta el RUT del paciente");

  var fecha = new Date();
  var registro = {
    fecha: fecha,
    nombre: nombre,
    rut: rut,
    descripcion: descripcion || fileNameOriginal || "Imagen sin descripcion",
    fileId: "",
    url: ""
  };

  // 1. REGISTRO MINIMO ASEGURADO EN LA HOJA
  var filaRegistrada = registrarImagenPaciente(registro);
  escribirLog("Registro base creado en fila " + filaRegistrada + " para rut: " + rut);

  // 2. INTENTO DE GUARDAR EN DRIVE
  try {
    if (!fileBase64) {
      throw new Error("No se recibio la imagen");
    }

    var carpeta = obtenerCarpetaImagenes();
    var timestamp = Utilities.formatDate(
      fecha,
      Session.getScriptTimeZone() || "America/Santiago",
      "yyyy-MM-dd_HHmmss"
    );

    var extension = obtenerExtension(fileNameOriginal, mimeType);
    var nombreSeguro = limpiarTextoArchivo(nombre || "paciente");
    var descripcionSegura = limpiarTextoArchivo(descripcion || "imagen");
    var nombreArchivo = rut + "_" + timestamp + "_" + nombreSeguro + "_" + descripcionSegura + extension;

    var bytes = Utilities.base64Decode(fileBase64);
    var blob = Utilities.newBlob(bytes, mimeType, nombreArchivo);
    var file = carpeta.createFile(blob);

    // Si esto falla, no perdemos el registro del sheet
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (sharingError) {
      escribirLog("WARNING setSharing: " + sharingError.toString());
    }

    registro.fileId = file.getId();
    registro.url = file.getUrl();

    var hoja = obtenerHojaImagenesPacientes();
    hoja.getRange(filaRegistrada, 5).setValue(registro.fileId);
    hoja.getRange(filaRegistrada, 6).setValue(registro.url);
    SpreadsheetApp.flush();

    escribirLog("Imagen subida a Drive y fila " + filaRegistrada + " actualizada para rut: " + rut);

    return {
      ok: true,
      nombre: nombre,
      rut: rut,
      fileId: registro.fileId,
      url: registro.url
    };
  } catch (error) {
    escribirLog("ERROR Drive en subirImagenPaciente: " + error.toString());

    return {
      ok: true,
      nombre: nombre,
      rut: rut,
      fileId: "",
      url: "",
      warning: "Se registro en la hoja, pero no se pudo completar Drive"
    };
  }
}

function listarImagenesPaciente(rut, nombre) {
  var hoja = obtenerHojaImagenesPacientes();
  var data = hoja.getDataRange().getValues();
  var rutBuscado = txt(rut).toLowerCase();
  var nombreBuscado = txt(nombre).toLowerCase();
  var resultado = [];

  for (var i = 1; i < data.length; i++) {
    var fila = data[i];
    var rutFila = txt(fila[2]).toLowerCase();
    var nombreFila = txt(fila[1]).toLowerCase();

    if (
      (rutBuscado && rutFila === rutBuscado) ||
      (nombreBuscado && nombreFila === nombreBuscado)
    ) {
      resultado.push({
        fecha: fila[0] || "",
        nombre: fila[1] || "",
        rut: fila[2] || "",
        descripcion: fila[3] || "",
        fileId: fila[4] || "",
        url: fila[5] || ""
      });
    }
  }

  return resultado;
}

function obtenerCarpetaImagenes() {
  var carpetas = DriveApp.getFoldersByName(NOMBRE_CARPETA_IMAGENES);
  if (!carpetas.hasNext()) {
    throw new Error("No existe una carpeta de Drive llamada '" + NOMBRE_CARPETA_IMAGENES + "'");
  }
  return carpetas.next();
}

function obtenerHojaImagenesPacientes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(NOMBRE_HOJA_IMAGENES);

  if (!hoja) {
    hoja = ss.insertSheet(NOMBRE_HOJA_IMAGENES);
    hoja.appendRow(["fecha", "nombre", "rut", "descripcion", "fileId", "url"]);
    SpreadsheetApp.flush();
  }

  return hoja;
}

function registrarImagenPaciente(data) {
  var hoja = obtenerHojaImagenesPacientes();
  hoja.appendRow([
    data.fecha || new Date(),
    data.nombre || "",
    data.rut || "",
    data.descripcion || "",
    data.fileId || "",
    data.url || ""
  ]);
  SpreadsheetApp.flush();
  return hoja.getLastRow();
}

function limpiarTextoArchivo(texto) {
  return (texto || "")
    .toString()
    .trim()
    .replace(/[^\w\-]+/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_+|_+$/g, "") || "archivo";
}

function obtenerExtension(fileName, mimeType) {
  var nombre = (fileName || "").toLowerCase();

  if (nombre.match(/\.[a-z0-9]+$/)) {
    return nombre.slice(nombre.lastIndexOf("."));
  }

  var mapa = {
    "image/jpeg": ".jpg",
    "image/jpg": ".jpg",
    "image/png": ".png",
    "image/webp": ".webp",
    "image/heic": ".heic",
    "application/pdf": ".pdf"
  };

  return mapa[mimeType] || ".bin";
}


/** LOGS */
function escribirLog(texto) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(HOJA_LOGS);

  if (!hoja) {
    hoja = ss.insertSheet(HOJA_LOGS);
    hoja.appendRow(["Fecha", "Mensaje"]);
  }

  hoja.appendRow([new Date(), texto]);
}


/**
 * ============================
 * ELIMINAR PACIENTE COMPLETO
 * ============================
 */
function eliminarPacienteCompleto(nombre, rut) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var nombreBuscado = txt(nombre);
  var rutBuscado = txt(rut);

  var hojaPacientes = ss.getSheetByName(HOJA_PACIENTES);
  if (hojaPacientes) {
    var dataPacientes = hojaPacientes.getDataRange().getValues();
    for (var i = dataPacientes.length - 1; i >= 1; i--) {
      var nombreFila = txt(dataPacientes[i][1]);
      var rutFila = txt(dataPacientes[i][2]);

      if (nombreFila === nombreBuscado && rutFila === rutBuscado) {
        hojaPacientes.deleteRow(i + 1);
      }
    }
  }

  var hojaDiag = ss.getSheetByName(HOJA_DIAGNOSTICOS);
  if (hojaDiag) {
    var dataDiag = hojaDiag.getDataRange().getValues();
    for (var j = dataDiag.length - 1; j >= 1; j--) {
      var nombreDiag = txt(dataDiag[j][0]);
      if (nombreDiag === nombreBuscado) {
        hojaDiag.deleteRow(j + 1);
      }
    }
  }

  var hojaImagenes = ss.getSheetByName(NOMBRE_HOJA_IMAGENES);
  if (hojaImagenes) {
    var dataImagenes = hojaImagenes.getDataRange().getValues();
    for (var n = dataImagenes.length - 1; n >= 1; n--) {
      var nombreImagen = txt(dataImagenes[n][1]);
      var rutImagen = txt(dataImagenes[n][2]);
      if (nombreImagen === nombreBuscado || (rutBuscado && rutImagen === rutBuscado)) {
        hojaImagenes.deleteRow(n + 1);
      }
    }
  }

  SpreadsheetApp.flush();
}
