// @ts-nocheck
/**
 * ============================================================
 * SISTEMA DE ACTAS DE REUNIÓN - OGPL UNMSM
 * Versión 2.5 - Final con Dashboard funcional
 * ============================================================
 */

const CONFIG = {
  FOLDER_ID: '1vaKJ-l-a5k3tLkkMsDXwdiEa24mdG560',
  SPREADSHEET_ID: '1co-LOqTlvhs2bUDMPiL-n3BuKM_La-uLOX6hZMklr18',
};
// ==================== LISTA DE FERIADOS ====================
const FERIADOS_PERU = [
  '2026-01-01', // Año Nuevo
  '2026-04-02', // Jueves Santo
  '2026-04-03', // Viernes Santo
  '2026-05-01', // Día del Trabajador
  '2026-06-29', // San Pedro y San Pablo
  '2026-07-28', // Fiestas Patrias
  '2026-07-29', // Fiestas Patrias
  '2026-08-06', // Batalla de Junín
  '2026-08-30', // Santa Rosa de Lima
  '2026-10-08', // Combate de Angamos
  '2026-11-01', // Día de todos los Santos
  '2026-12-08', // Inmaculada Concepción
  '2026-12-09', // Batalla de Ayacucho
  '2026-12-25'  // Navidad
];
// -------------------- UTILIDADES --------------------
function getSpreadsheet() {
  if (!CONFIG.SPREADSHEET_ID) throw new Error('No está configurado el SPREADSHEET_ID.');
  try {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  } catch(e) {
    throw new Error('No se pudo abrir la base de datos. Error: ' + e.message);
  }
}

function hashPassword(password, salt) {
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password + salt);
  return digest.map(function(b) { return (b < 0 ? b + 256 : b).toString(16).padStart(2, '0'); }).join('');
}

function generarToken(email) {
  var random = Math.random().toString(36).substring(2);
  var timestamp = new Date().getTime();
  var token = hashPassword(email + timestamp + random, 'appscript_salt_2026');
  CacheService.getScriptCache().put(token, email, 7200);
  return token;
}

function verificarToken(token) {
  if (!token) throw new Error('Token de sesión requerido.');
  var email = CacheService.getScriptCache().get(token);
  if (!email) throw new Error('Sesión expirada o inválida.');
  return email;
}

function getUserByEmail(email) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Usuarios');
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      return { email: data[i][0], nombre: data[i][2], rol: data[i][3], unidad: data[i][4] };
    }
  }
  return null;
}

function calcularFechaLimite(fechaBase, plazo, unidad) {
  if (!unidad) return null;
  var uni = unidad.toString().toLowerCase().trim();

  // Si es un acuerdo de decisión sin tiempo asignado
  if (uni === 'sin tiempo' || uni === 'n/a' || uni === 'sin plazo') {
    return null;
  }

  var fecha = new Date(fechaBase);
  var n = parseInt(plazo) || 0;

  // Si la unidad es días hábiles, entramos al bucle de salto de fines de semana y feriados
  if (uni === 'días hábiles' || uni === 'días') {
    var diasAgregados = 0;
    
    while (diasAgregados < n) {
      fecha.setDate(fecha.getDate() + 1); // Suma 1 día natural
      var day = fecha.getDay();
      
      // Formateamos la fecha a YYYY-MM-DD para buscarla en el arreglo
      var dateString = Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      var esFeriado = FERIADOS_PERU.indexOf(dateString) !== -1;
      
      // Si no es domingo (0), ni sábado (6), ni feriado, es un día hábil válido
      if (day !== 0 && day !== 6 && !esFeriado) {
        diasAgregados++;
      }
    }
  } else {
    // Lógica estándar para horas y meses
    switch(uni) {
      case 'horas': fecha.setHours(fecha.getHours() + n); break;
      case 'meses': fecha.setMonth(fecha.getMonth() + n); break;
      default:      fecha.setDate(fecha.getDate() + n);
    }
  }
  
  return fecha;
}

// -------------------- INICIALIZACIÓN --------------------
function inicializarSistema() {
  var ss = getSpreadsheet();
  var hojasRequeridas = [
    { nombre: 'Usuarios', headers: ['Email','PasswordHash','Nombre','Rol','UnidadOficina','FechaRegistro'] },
    { nombre: 'Actas',    headers: ['ID_Acta','Tema','Modalidad','Lugar','Fecha','HoraInicio','HoraFin','URL_PDF','RegistradoPor','Timestamp','CantidadAsistentes','CantidadFotos'] },
    { nombre: 'Asistentes', headers: ['ID_Acta','Nombres','Apellidos','Cargo','Unidad','TieneFirma','Timestamp'] },
    { nombre: 'Acuerdos', headers: ['ID_Acuerdo','ID_Acta','Acuerdo','Responsable','Plazo','UnidadPlazo','FechaLimite','Estado','FechaCumplimiento','DiasRestantes','Indicador','Timestamp'] }
  ];
  
  var creadas = 0;
  hojasRequeridas.forEach(function(h) {
    if (!ss.getSheetByName(h.nombre)) {
      var sheet = ss.insertSheet(h.nombre);
      sheet.appendRow(h.headers);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, h.headers.length).setFontWeight('bold').setBackground('#d9e2f3');
      creadas++;
    }
  });
  
  return '✅ Inicialización completa. Hojas creadas: ' + creadas;
}

function crearUsuarioAdmin() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Usuarios');
  if (!sheet) return '❌ No existe la hoja "Usuarios".';
  
  var email = 'admin@unmsm.edu.pe';
  var password = 'admin123';
  
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === email) return '⚠️ El usuario ' + email + ' ya existe.';
  }
  
  var hash = hashPassword(password, email);
  sheet.appendRow([email, hash, 'Administrador Principal', 'Admin', 'OGPL', new Date()]);
  return '✅ Usuario creado: ' + email;
}

// -------------------- WEB APP --------------------
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Sistema de Actas de Reunión - OGPL UNMSM')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// -------------------- API: AUTENTICACIÓN --------------------
function login(credenciales) {
  try {
    var email = credenciales.email;
    var password = credenciales.password;
    var user = getUserByEmail(email);
    if (!user) return { success: false, error: 'Usuario no encontrado en el sistema.' };
    
    var hash = hashPassword(password, email);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Usuarios');
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === email && data[i][1] === hash) {
        var token = generarToken(email);
        return { success: true, token: token, user: user };
      }
    }
    return { success: false, error: 'Contraseña incorrecta.' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function getSessionUser(token) {
  try {
    var email = verificarToken(token);
    var user = getUserByEmail(email);
    if (!user) throw new Error('Usuario no encontrado en sesión activa.');
    return { success: true, user: user };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// -------------------- API: CREAR USUARIOS --------------------
function crearUsuario(datos, token) {
  try {
    verificarToken(token);
    var email = datos.email, password = datos.password, nombre = datos.nombre;
    if (!email || !password || !nombre) {
      return { success: false, error: 'Email, password y nombre son obligatorios.' };
    }
    if (getUserByEmail(email)) return { success: false, error: 'Ya existe un usuario con ese correo.' };
    
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Usuarios');
    var hash = hashPassword(password, email);
    sheet.appendRow([email, hash, nombre, datos.rol || 'Usuario', datos.unidad || 'OGPL', new Date()]);
    return { success: true, message: 'Usuario creado: ' + email };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// -------------------- API: GENERAR ACTA Y PDF --------------------
function generarActaPDF(datos, token) {
  var lock = LockService.getScriptLock();
  
  try {
    var email = verificarToken(token);
    var carpeta = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    var now = new Date();
    
    lock.waitLock(15000); 
    
    var ss = getSpreadsheet();
    var sheetActas = ss.getSheetByName('Actas');
    var correlativo = sheetActas.getLastRow(); 
    var numeroActa = correlativo.toString().padStart(3, '0'); 
    var anio = now.getFullYear();
    
    var idActaSolo = numeroActa + '-' + anio + '-OR-OGPL/UNMSM';
    var idActa = 'ACTA N° ' + idActaSolo;
    
    // Generación del Código QR dinámico (Se ubicará en el pie de página)
    var fechaEmision = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    var qrText = encodeURIComponent("Documento Original SIGA. " + idActa + " Fecha: " + fechaEmision);
    var qrUrl = "https://quickchart.io/qr?size=120&text=" + qrText;
    
    var htmlStr = '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8">' +
      '<style>' +
      '@page { margin: 40px; }' +
      'body { font-family: Arial, sans-serif; color: #000; font-size: 11px; }' +
      '* { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; color-adjust: exact !important; }' +
      '.institucional { text-align: center; margin-bottom: 8px; }' +
      '.institucional img { height: 65px; margin-bottom: 6px; }' +
      '.institucional .uni-nombre { font-family: Georgia, "Times New Roman", Times, serif; font-weight: bold; font-size: 15px; color: #16305c; }' +
      '.institucional .uni-sub { font-family: Georgia, "Times New Roman", Times, serif; font-size: 11.5px; color: #16305c; }' +
      '.institucional .uni-oficina { font-family: Georgia, "Times New Roman", Times, serif; font-weight: bold; font-size: 12.5px; color: #16305c; }' +
      '.institucional-rule { border: none; border-top: 1pt solid #999; margin: 8px 0 18px 0; }' +
      '.header-table { width: 100%; border-collapse: collapse; margin-bottom: 25px; border: none; }' +
      '.header-table td { border: none; text-align: center; vertical-align: middle; }' +
      '.box-left { border: 1.5pt solid #000 !important; width: 22%; font-weight: bold; font-size: 12px; padding: 8px; }' +
      '.box-title { width: 63%; font-weight: bold; font-size: 24px; }' +
      '.box-logo { border: 1.5pt solid #000 !important; width: 15%; padding: 6px; }' +
      '.box-logo img { max-height: 55px; max-width: 100%; }' +
      '.agenda-table { width: 100%; border-collapse: collapse; margin-bottom: 20px; border: 1pt solid #000; }' +
      '.agenda-table td { border: 1pt solid #000; padding: 6px; }' +
      '.agenda-header { background-color: #d9e2f3 !important; font-weight: bold; }' +
      '.agenda-table ol { margin: 0; padding-left: 20px; }' +
      '.agenda-table li { margin-bottom: 4px; }' +
      '.info-table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }' +
      '.info-table td { border: 1pt solid #000; padding: 6px; }' +
      '.info-label { background-color: #d9e2f3 !important; font-weight: bold; width: 20%; }' +
      '.data-table { width: 100%; border-collapse: collapse; margin-bottom: 25px; text-align: center; border: 1pt solid #000; }' +
      '.data-table th, .data-table td { border: 1pt solid #000; padding: 6px; vertical-align: middle; }' +
      '.data-table th { background-color: #d9e2f3 !important; font-weight: bold; }' +
      '.signature-table { width: 100%; border: none !important; margin: 0; padding: 0; text-align: left; }' +
      '.signature-table td { border: none !important; padding: 1px; line-height: 1.2; font-size: 8.5px; }' +
      '.footer-table { width: 100%; margin-top: 30px; border-top: 1.5pt solid #000 !important; padding-top: 10px; border-collapse: collapse; }' +
      '.footer-table td { border: none; vertical-align: middle; }' +
      '.page-break { page-break-before: always; }' +
      '.foto-container { width: 48%; display: inline-block; margin: 1%; text-align: center; border: 0.5pt solid #ccc; padding: 5px; box-sizing: border-box; }' +
      '</style></head><body>' +
      
      // CABECERA INSTITUCIONAL (Escudo + Universidad + Oficina, centrado)
      '<div class="institucional">' +
      '<img src="' + LOGO_UNMSM + '" alt="Escudo UNMSM"/>' +
      '<div class="uni-nombre">UNIVERSIDAD NACIONAL MAYOR DE SAN MARCOS</div>' +
      '<div class="uni-sub">Universidad del Perú. Decana de América</div>' +
      '<div class="uni-oficina">OFICINA GENERAL DE PLANIFICACIÓN</div>' +
      '</div>' +
      '<hr class="institucional-rule"/>' +

      // CABECERA (Recuadro Acta N°, título y escudo)
      '<table class="header-table"><tr>' +
      '<td class="box-left">ACTA N°<br>' + numeroActa + '-' + anio + '-OR-<br>OGPL/UNMSM</td>' +
      '<td class="box-title">ACTA DE REUNIÓN</td>' +
      '<td class="box-logo"><img src="' + LOGO_UNMSM + '" alt="Escudo UNMSM"/></td>' +
      '</tr></table>' +
      
      // DETALLES DE REUNIÓN (Info Section)
      '<table class="info-table">' +
      '<tr><td class="info-label">Tema:</td><td>' + datos.tema + '</td></tr>' +
      '<tr><td class="info-label">Modalidad:</td><td>' + datos.modalidad + '</td></tr>' +
      '<tr><td class="info-label">Fecha:</td><td>' + datos.fecha + '</td></tr>' +
      '<tr><td class="info-label">Lugar de reunión:</td><td>' + datos.lugar + '</td></tr>' +
      '<tr><td class="info-label">Horario:</td><td>' + datos.horaInicio + ' - ' + datos.horaFin + '</td></tr>' +
      '</table>' +
      
      // ASISTENTES (Data Table)
      '<table class="data-table">' +
      '<thead><tr>' +
      '<th style="width: 5%;">N°</th>' +
      '<th style="width: 35%;">Nombre Y Apellidos</th>' +
      '<th style="width: 25%;">Cargo / Unidad</th>' +
      '<th style="width: 35%;">Firma</th>' +
      '</tr></thead><tbody>';
    
    datos.asistentes.forEach(function(asis, index) {
      var firmaContent = '';
      if (asis.usarFirmaDigital) {
        var apellidosMayus = asis.apellidos.toUpperCase();
        var nombresCap = asis.nombres.split(' ').map(function(w){ return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase(); }).join(' ');
        var fechaStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yy');
        var horaStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');
        
        firmaContent = '<table class="signature-table"><tr>' +
                       '<td style="width:25%; text-align:center;">' +
                       '<img src="' + LOGO_UNMSM + '" style="max-height:35px;"/></td>' +
                       '<td style="width:75%; color:#222;">' +
                       'Firmado digitalmente por <br><b>' + apellidosMayus + ' ' + nombresCap + '</b><br>' +
                       'Motivo: Soy el Autor de la Firma<br>' +
                       'Fecha: ' + fechaStr + ' Hora: ' + horaStr +
                       '</td></tr></table>';
      } else if (asis.firma) {
        firmaContent = '<img src="' + asis.firma + '" style="max-height: 40px; max-width: 100%; display: block; margin: auto;"/>';
      }

      htmlStr += '<tr>' +
        '<td>' + (index + 1) + '</td>' +
        '<td style="text-align: left;">' + asis.nombres + ' ' + asis.apellidos + '</td>' +
        '<td>' + asis.cargo + '<br><small>' + asis.unidad + '</small></td>' +
        '<td style="padding: 2px;">' + firmaContent + '</td></tr>';
    });
    htmlStr += '</tbody></table>';
      
    // AGENDA
    htmlStr += '<table class="agenda-table">' +
      '<tr><td class="agenda-header">Agenda a tratar:</td></tr>' +
      '<tr><td><ol>' +
      (datos.agenda.length > 0 ? datos.agenda.map(function(item) { return '<li>' + item + '</li>'; }).join('') : '<li>---</li>') +
      '</ol></td></tr>' +
      '</table>';
      
    // ACUERDOS (SIN PLAZO)
    if (datos.acuerdos.length > 0) {
      htmlStr += '<table class="data-table">' +
        '<thead><tr>' +
        '<th style="width: 5%;">N°</th>' +
        '<th style="width: 65%;">Acuerdo</th>' +
        '<th style="width: 30%;">Responsable</th>' +
        '</tr></thead><tbody>';
      datos.acuerdos.forEach(function(ac, idx) {
        htmlStr += '<tr>' +
          '<td>' + (idx + 1) + '</td>' +
          '<td style="text-align: left;">' + ac.texto + '</td>' +
          '<td>' + ac.responsable + '</td>' +
          '</tr>';
      });
      htmlStr += '</tbody></table>';
    }

    // COMPROMISOS (CON PLAZO)
    if (datos.compromisos.length > 0) {
      htmlStr += '<table class="data-table">' +
        '<thead><tr>' +
        '<th style="width: 5%;">N°</th>' +
        '<th style="width: 45%;">Compromiso</th>' +
        '<th style="width: 20%;">Responsable</th>' +
        '<th style="width: 15%;">Plazo</th>' +
        '<th style="width: 15%;">Fecha Límite</th>' +
        '</tr></thead><tbody>';
      
      datos.compromisos.forEach(function(co, idx) {
        var fLimite = calcularFechaLimite(now, co.plazo, co.unidad);
        var fStr = Utilities.formatDate(fLimite, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
        htmlStr += '<tr>' +
          '<td>' + (idx + 1) + '</td>' +
          '<td style="text-align: left;">' + co.texto + '</td>' +
          '<td>' + co.responsable + '</td>' +
          '<td>' + co.plazo + ' ' + co.unidad + '</td>' +
          '<td>' + fStr + '</td>' +
          '</tr>';
      });
      htmlStr += '</tbody></table>';
    }
    
    // EVIDENCIAS FOTOGRÁFICAS
    if (datos.fotos && datos.fotos.length > 0) {
      htmlStr += '<div class="page-break"></div>' +
        '<h3 style="text-align: center; border-bottom: 1pt solid #000; padding-bottom: 5px;">ANEXO FOTOGRÁFICO DE LA REUNIÓN</h3>' +
        '<div style="width: 100%;">';
      datos.fotos.forEach(function(foto, idx) {
        htmlStr += '<div class="foto-container">' +
          '<img src="' + foto + '" style="max-width: 100%; max-height: 250px; display: block; margin-bottom: 5px; margin-left: auto; margin-right: auto;"/>' +
          '<span style="font-size: 8px; color: #666;">Evidencia N° ' + (idx + 1) + '</span>' +
          '</div>';
      });
      htmlStr += '</div>';
    }
    
    // PIE DE PÁGINA (Con QR dinámico y número de página manual)
    htmlStr += '<table class="footer-table"><tr>' +
      '<td style="width: 15%; text-align: left;">' +
      '<img src="' + qrUrl + '" style="max-height: 70px;" alt="QR Code"/></td>' +
      '<td style="width: 70%; text-align: justify; font-size: 9.5px; padding: 0 15px; color: #222;">' +
      'Documento certificado y generado digitalmente por el Sistema Integral de Actas (SIGA). ' +
      'Este documento es original y tiene validez conforme a la normativa vigente. Oficina de ' +
      'Racionalización / Oficina General de Planificación- UNMSM. Fecha de emisión: ' + fechaEmision + '.' +
      '</td>' +
      '<td style="width: 15%; text-align: right; vertical-align: bottom; font-size: 11px; font-weight: bold;">' +
      '<span style="border: 1pt solid #000; padding: 3px 10px; display: inline-block;">Pág. 1</span></td>' +
      '</tr></table></body></html>';
    
    var blob = HtmlService.createHtmlOutput(htmlStr).getAs(MimeType.PDF);
    blob.setName(idActaSolo + ' - ' + datos.tema + '.pdf'); 
    var archivo = carpeta.createFile(blob);
    var urlPDF = archivo.getUrl();

    // Registra el acta, los asistentes y los acuerdos/compromisos en la hoja
    guardarActaEnBD(idActa, datos, urlPDF, email);

    // Liberamos el bloqueo de concurrencia
    lock.releaseLock();

    return { success: true, url: urlPDF, idActa: idActa };

  } catch (error) {
    if (lock.hasLock()) lock.releaseLock();
    return { success: false, error: error.toString() };
  }
}
    
// -------------------- FUNCIÓN PARA GUARDAR EN BASE DE DATOS --------------------
function guardarActaEnBD(idActa, datos, urlPDF, email) {
  var ss = getSpreadsheet();
  var now = new Date();
  
  // 1. Guardar en la hoja Actas
  var sheetActas = ss.getSheetByName('Actas');
  sheetActas.appendRow([
    idActa, datos.tema, datos.modalidad, datos.lugar, datos.fecha,
    datos.horaInicio, datos.horaFin, urlPDF, email, now,
    datos.asistentes.length, (datos.fotos || []).length
  ]);
  
  // 2. Guardar en la hoja Asistentes
  var sheetAsist = ss.getSheetByName('Asistentes');
  datos.asistentes.forEach(function(a) {
    sheetAsist.appendRow([idActa, a.nombres, a.apellidos, a.cargo, a.unidad, a.firma || a.usarFirmaDigital ? 'Sí' : 'No', now]);
  });
  
  // 3. Guardar Acuerdos y Compromisos en la hoja Acuerdos
  var sheetAcuerdos = ss.getSheetByName('Acuerdos');
  var correlativoGlobal = 1;
  
  // Acuerdos simples (Se guardan como "Cumplido" para que el Dashboard no los marque como pendientes)
  if (datos.acuerdos && datos.acuerdos.length > 0) {
    datos.acuerdos.forEach(function(ac) {
      var idAcuerdo = idActa + '_AC' + correlativoGlobal.toString().padStart(3, '0');
      sheetAcuerdos.appendRow([
        idAcuerdo, idActa, ac.texto, ac.responsable, '-', 'N/A',
        now, 'Cumplido', now, '', 'ACUERDO', now
      ]);
      correlativoGlobal++;
    });
  }

  // Compromisos con fecha límite
  if (datos.compromisos && datos.compromisos.length > 0) {
    datos.compromisos.forEach(function(co) {
      var idAcuerdo = idActa + '_CO' + correlativoGlobal.toString().padStart(3, '0');
      var fechaLimite = calcularFechaLimite(now, co.plazo, co.unidad);
      sheetAcuerdos.appendRow([
        idAcuerdo, idActa, co.texto, co.responsable, co.plazo, co.unidad,
        fechaLimite, 'Pendiente', '', '', 'PENDIENTE', now
      ]);
      correlativoGlobal++;
    });
  }
}
// -------------------- API: DASHBOARD FINAL --------------------
function obtenerAcuerdos(token, filtroOficina) {
  try {
    verificarToken(token);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Acuerdos');
    
    if (!sheet) {
      return { success: false, error: 'No existe la hoja Acuerdos.' };
    }
    
    var data = sheet.getDataRange().getValues();
    var ahora = new Date();
    var resultados = [];
    
    if (data.length <= 1) {
      return { success: true, acuerdos: [] };
    }
    
    // AQUÍ INICIA EL BUCLE FOR 
    for (var i = 1; i < data.length; i++) {
      try {
        var row = data[i];
        
        // Verificar fila completa
        if (!row || row.length < 12) continue;
        
        var idAcuerdo = row[0];
        var idActa = row[1];
        var acuerdoTexto = row[2];
        var responsable = row[3];
        var plazo = row[4];
        var unidadPlazo = row[5];
        var fechaLimiteRaw = row[6];
        var estado = row[7];
        var indicador = row[10];
        
        if (!idAcuerdo) continue;
        if (filtroOficina && responsable !== filtroOficina) continue;
        
        // Identificar si es un acuerdo sin tiempo
        var esSinTiempo = (fechaLimiteRaw === 'N/A' || fechaLimiteRaw === '-' || estado === 'Decisión' || indicador === 'INFORMATIVO');
        var fechaLimite = null;
        
        if (!esSinTiempo) {
          if (fechaLimiteRaw instanceof Date) {
            fechaLimite = fechaLimiteRaw;
          } else if (typeof fechaLimiteRaw === 'string' && fechaLimiteRaw.trim() !== '') {
            var partes = fechaLimiteRaw.match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})/);
            if (partes) {
              fechaLimite = new Date(
                parseInt(partes[3]), parseInt(partes[2]) - 1, parseInt(partes[1]),
                parseInt(partes[4]), parseInt(partes[5]), parseInt(partes[6])
              );
            } else {
              fechaLimite = new Date(fechaLimiteRaw);
              if (isNaN(fechaLimite.getTime())) fechaLimite = new Date();
            }
          } else {
            fechaLimite = new Date();
          }
        }
        
        var estadoFinal = estado || 'Pendiente';
        var indicadorFinal = indicador || 'PENDIENTE';
        var diasRestantesStr = '';
        var clase = 'azul';
        var fechaLimiteStr = 'N/A';
        
        if (esSinTiempo) {
          estadoFinal = 'Decisión';
          indicadorFinal = 'INFORMATIVO';
          clase = 'azul'; 
          diasRestantesStr = '-';
        } else if (estadoFinal === 'Cumplido') {
          indicadorFinal = 'CUMPLIDO';
          clase = 'verde';
          fechaLimiteStr = Utilities.formatDate(fechaLimite, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
        } else if (ahora > fechaLimite) {
          estadoFinal = 'Vencido';
          indicadorFinal = 'NO CUMPLIDO';
          clase = 'rojo';
          fechaLimiteStr = Utilities.formatDate(fechaLimite, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
        } else {
          var diffMs = fechaLimite - ahora;
          var diffDias = Math.ceil(diffMs / (1000 * 60 * 60 * 24));
          var diffHoras = Math.ceil(diffMs / (1000 * 60 * 60));
          
          diasRestantesStr = diffDias < 1 ? diffHoras + 'h restantes' : diffDias + ' días';
          
          if (diffDias <= 1) {
            indicadorFinal = 'URGENTE';
            clase = 'amarillo';
          } else {
            indicadorFinal = 'EN PLAZO';
            clase = 'azul';
          }
          fechaLimiteStr = Utilities.formatDate(fechaLimite, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
        }
        
        resultados.push({
          idAcuerdo: String(idAcuerdo),
          idActa: String(idActa || ''),
          acuerdo: String(acuerdoTexto || 'Sin descripción'),
          responsable: String(responsable || 'No asignado'),
          plazo: esSinTiempo ? 'Sin tiempo' : String(plazo || '0') + ' ' + String(unidadPlazo || ''),
          fechaLimite: fechaLimiteStr,
          estado: estadoFinal,
          indicador: indicadorFinal,
          clase: clase,
          diasRestantes: diasRestantesStr
        });
        
      } catch (rowError) {
        // Continuar con siguiente fila
      }
    } 
    // FIN DEL BUCLE FOR
    
    // Ordenar (Se añade "INFORMATIVO" para que los acuerdos sin plazo queden al final y no interfieran con las urgencias)
    var orden = { 'Vencido': 1, 'URGENTE': 2, 'EN PLAZO': 3, 'CUMPLIDO': 4, 'INFORMATIVO': 5 };
    resultados.sort(function(a, b) {
      return (orden[a.indicador] || 6) - (orden[b.indicador] || 6);
    });
    
    return { success: true, acuerdos: resultados };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// -------------------- API: HISTORIAL DE ASISTENTES --------------------
/**
 * Devuelve el historial de asistentes registrados para alimentar los
 * datalist de autocompletado del formulario (nombres, apellidos, cargos
 * y unidades). Se llama sin token porque solo expone datos ya presentes
 * en las actas y se usa al abrir el formulario.
 */
function obtenerHistorialAsistentes() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Asistentes');
    if (!sheet) return { success: true, asistentes: [], apellidos: [], cargos: [], unidades: [] };

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: true, asistentes: [], apellidos: [], cargos: [], unidades: [] };
    }

    var vistos = {};
    var asistentes = [];
    var setApellidos = {}, setCargos = {}, setUnidades = {};

    // Columnas: ID_Acta(0), Nombres(1), Apellidos(2), Cargo(3), Unidad(4)
    for (var i = 1; i < data.length; i++) {
      var nombre   = String(data[i][1] || '').trim();
      var apellido = String(data[i][2] || '').trim();
      var cargo    = String(data[i][3] || '').trim();
      var unidad   = String(data[i][4] || '').trim();

      if (!nombre) continue;

      // Una entrada por persona (nombre + apellido), la más reciente gana
      var clave = (nombre + '|' + apellido).toLowerCase();
      if (vistos[clave] === undefined) {
        vistos[clave] = asistentes.length;
        asistentes.push({ nombre: nombre, apellido: apellido, cargo: cargo, unidad: unidad });
      } else {
        var prev = asistentes[vistos[clave]];
        if (cargo)  prev.cargo  = cargo;
        if (unidad) prev.unidad = unidad;
      }

      if (apellido) setApellidos[apellido] = true;
      if (cargo)    setCargos[cargo]       = true;
      if (unidad)   setUnidades[unidad]    = true;
    }

    asistentes.sort(function(a, b) { return a.nombre.localeCompare(b.nombre); });

    return {
      success: true,
      asistentes: asistentes,
      apellidos: Object.keys(setApellidos).sort(),
      cargos:    Object.keys(setCargos).sort(),
      unidades:  Object.keys(setUnidades).sort()
    };
  } catch (e) {
    return { success: false, error: e.toString(), asistentes: [], apellidos: [], cargos: [], unidades: [] };
  }
}

// -------------------- API: OFICINAS RESPONSABLES --------------------
/**
 * Lista sin repetidos las oficinas responsables registradas en la hoja
 * Acuerdos. Alimenta el desplegable de filtro del panel gerencial.
 */
function obtenerOficinas() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Acuerdos');
    if (!sheet) return { success: true, oficinas: [] };

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, oficinas: [] };

    var set = {};
    // Columna Responsable = índice 3
    for (var i = 1; i < data.length; i++) {
      var resp = String(data[i][3] || '').trim();
      if (resp) set[resp] = true;
    }

    return { success: true, oficinas: Object.keys(set).sort() };
  } catch (e) {
    return { success: false, error: e.toString(), oficinas: [] };
  }
}

// -------------------- API: MARCAR ACUERDO COMO CUMPLIDO --------------------
/**
 * Marca un acuerdo o compromiso como cumplido: escribe Estado, la fecha
 * de cumplimiento y el indicador. Usa LockService para evitar escrituras
 * simultáneas sobre la misma fila.
 */
function marcarAcuerdoCumplido(idAcuerdo, token) {
  var lock = LockService.getScriptLock();
  try {
    verificarToken(token);

    if (!idAcuerdo) return { success: false, error: 'ID de acuerdo no válido.' };

    lock.waitLock(15000);

    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Acuerdos');
    if (!sheet) return { success: false, error: 'No existe la hoja Acuerdos.' };

    var data = sheet.getDataRange().getValues();

    // Columnas (base 1): H=8 Estado, I=9 FechaCumplimiento, K=11 Indicador
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(idAcuerdo).trim()) {
        if (String(data[i][7]).trim() === 'Cumplido') {
          lock.releaseLock();
          return { success: false, error: 'Ese acuerdo ya estaba marcado como cumplido.' };
        }

        var fila = i + 1;
        var ahora = new Date();
        sheet.getRange(fila, 8).setValue('Cumplido');
        sheet.getRange(fila, 9).setValue(ahora);
        sheet.getRange(fila, 11).setValue('CUMPLIDO');

        lock.releaseLock();
        return {
          success: true,
          message: 'Acuerdo ' + idAcuerdo + ' marcado como cumplido el ' +
                   Utilities.formatDate(ahora, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') + '.'
        };
      }
    }

    lock.releaseLock();
    return { success: false, error: 'No se encontró el acuerdo ' + idAcuerdo + '.' };

  } catch (e) {
    if (lock.hasLock()) lock.releaseLock();
    return { success: false, error: e.toString() };
  }
}
