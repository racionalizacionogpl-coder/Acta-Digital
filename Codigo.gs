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
      '@page { size: A4; margin: 12mm 16mm 26mm 16mm; }' +
      '* { box-sizing: border-box; }' +
      'html, body {' +
      '  margin: 0; padding: 0 0 6mm 0;' +
      '  font-family: Georgia, "Times New Roman", serif;' +
      '  color: #2e2e2e;' +
      '  background: #ffffff;' +
      '  -webkit-print-color-adjust: exact !important;' +
      '  print-color-adjust: exact !important;' +
      '}' +
      ':root {' +
      '  --principal: #1f3a5f;' +
      '  --principal-oscuro: #132840;' +
      '  --acento: #8a7530;' +
      '  --gris-texto: #2e2e2e;' +
      '  --gris-suave: #6b6b6b;' +
      '  --linea: #dcdcdc;' +
      '  --fondo-suave: #f7f8fa;' +
      '}' +
      '.watermark { position: fixed; top: 44%; left: 50%; width: 120mm; transform: translate(-50%, -50%); opacity: 0.05; z-index: 0; pointer-events: none; }' +
      '.content { position: relative; z-index: 1; }' +
      '.header { display: flex; flex-direction: row; align-items: center; justify-content: flex-start; gap: 8mm; padding-bottom: 5mm; border-bottom: 0.7pt solid var(--linea); text-align: center; }' +
      '.header img.escudo { width: 19mm; height: auto; flex-shrink: 0; order: 1; }' +
      '.header-text { text-align: center; order: 2; flex: 1; }' +
      '.header-text .linea1 { font-size: 12.5pt; font-weight: bold; letter-spacing: 0.6pt; color: var(--principal-oscuro); text-transform: uppercase; line-height: 1.25; }' +
      '.header-text .linea2 { font-size: 8.5pt; font-style: italic; color: var(--gris-suave); margin-top: 1.2mm; }' +
      '.header-text .linea3 { font-size: 8.5pt; font-weight: bold; letter-spacing: 0.5pt; color: var(--principal-oscuro); text-transform: uppercase; margin-top: 1.8mm; padding-top: 1.8mm; border-top: 0.5pt solid var(--linea); display: inline-block; }' +
      '.acta-code { text-align: center; font-size: 8pt; color: var(--gris-suave); line-height: 1.5; margin: 3mm 0 6mm 0; }' +
      '.acta-code .num { font-size: 9.3pt; font-weight: bold; color: var(--principal); display: block; margin-bottom: 0.8mm; }' +
      '.title-band { text-align: center; margin: 6mm 0 4mm 0; }' +
      '.title-band h1 { display: inline-block; margin: 0; font-size: 19pt; letter-spacing: 2.5pt; color: var(--principal-oscuro); font-weight: bold; text-transform: uppercase; padding-bottom: 2mm; border-bottom: 1pt solid var(--acento); }' +
      '.info-table { width: 100%; border-collapse: collapse; margin-bottom: 6mm; font-size: 9.8pt; }' +
      '.info-table tr { border-bottom: 0.6pt solid var(--linea); }' +
      '.info-table tr:last-child { border-bottom: none; }' +
      '.info-table td { padding: 3mm; vertical-align: top; }' +
      '.info-table td.label { width: 36mm; font-weight: bold; color: var(--principal); text-transform: uppercase; font-size: 8.3pt; letter-spacing: 0.3pt; }' +
      '.info-table td.value { color: var(--gris-texto); }' +
      '.section-title { font-size: 9.5pt; font-weight: bold; color: #ffffff; background: var(--principal); text-transform: uppercase; letter-spacing: 0.6pt; padding: 2.3mm 4mm; margin: 0 0 4mm 0; border-radius: 1.2pt; }' +
      '.asistentes { width: 100%; border-collapse: collapse; margin-bottom: 7mm; font-size: 9.3pt; }' +
      '.asistentes th { background: var(--fondo-suave); border: 0.6pt solid var(--linea); padding: 2.6mm 3mm; text-align: left; color: var(--principal-oscuro); font-size: 8.3pt; text-transform: uppercase; letter-spacing: 0.3pt; }' +
      '.asistentes td { border: 0.6pt solid var(--linea); padding: 3mm; vertical-align: middle; }' +
      '.asistentes td.n { text-align: center; width: 8mm; color: var(--gris-suave); }' +
      '.asistentes td.firma { text-align: center; width: 54mm; padding: 2px; }' +
      '.asistentes td.firma .sello { display: inline-flex; align-items: center; gap: 1.8mm; font-size: 6.3pt; line-height: 1.55; color: #333333; text-align: left; white-space: nowrap; }' +
      '.asistentes td.firma .sello img { width: 11mm; height: auto; flex-shrink: 0; }' +
      '.asistentes td.firma .sello b { font-size: 6.7pt; color: #1a1a1a; }' +
      '.asistentes td.firma img.firma-img { max-height: 40px; max-width: 100%; display: block; margin: auto; }' +
      '.agenda-list { margin: 0 0 7mm 0; padding: 0; list-style: none; }' +
      '.agenda-list li { display: flex; align-items: center; gap: 3mm; padding: 2.8mm 0; border-bottom: 0.5pt dashed var(--linea); font-size: 10pt; }' +
      '.agenda-list li:last-child { border-bottom: none; }' +
      '.agenda-list .idx { flex-shrink: 0; width: 6.5mm; height: 6.5mm; border-radius: 50%; background: var(--principal); color: #fff; font-size: 8.5pt; font-weight: bold; display: flex; align-items: center; justify-content: center; }' +
      '.registro-table { width: 100%; border-collapse: collapse; margin-bottom: 7mm; font-size: 9.3pt; }' +
      '.registro-table th { background: var(--fondo-suave); border: 0.6pt solid var(--linea); padding: 2.6mm 3mm; text-align: left; color: var(--principal-oscuro); font-size: 8.3pt; text-transform: uppercase; letter-spacing: 0.3pt; }' +
      '.registro-table td { border: 0.6pt solid var(--linea); padding: 3mm; vertical-align: top; }' +
      '.registro-table td.n { text-align: center; width: 8mm; color: var(--gris-suave); }' +
      '.anexo-titulo { text-align: center; font-size: 12pt; font-weight: bold; color: var(--principal-oscuro); text-transform: uppercase; letter-spacing: 0.5pt; border-bottom: 0.7pt solid var(--linea); padding-bottom: 3mm; margin: 0 0 6mm 0; }' +
      '.foto-container { width: 48%; display: inline-block; margin: 1%; text-align: center; border: 0.6pt solid var(--linea); padding: 5px; box-sizing: border-box; vertical-align: top; }' +
      '.page-break { page-break-before: always; }' +
      '.footer { position: fixed; bottom: 0; left: 0; right: 0; padding-top: 3mm; border-top: 0.7pt solid var(--linea); display: flex; align-items: center; gap: 5mm; font-size: 7.3pt; color: var(--gris-suave); line-height: 1.5; background: #ffffff; }' +
      '.footer img.qr { width: 16mm; height: 16mm; flex-shrink: 0; }' +
      '.footer .txt { flex: 1; }' +
      '.footer .pag { flex-shrink: 0; border: 0.6pt solid var(--linea); padding: 1mm 3mm; font-size: 8pt; font-weight: bold; color: var(--principal); align-self: center; }' +
      '</style></head><body>' +

      '<img class="watermark" src="' + LOGO_UNMSM + '" alt="">' +
      '<div class="content">' +

      // MEMBRETE (escudo + texto institucional)
      '<div class="header">' +
      '<img class="escudo" src="' + LOGO_UNMSM + '" alt="Escudo UNMSM">' +
      '<div class="header-text">' +
      '<div class="linea1">Universidad Nacional Mayor de San Marcos</div>' +
      '<div class="linea2">Universidad del Perú. Decana de América</div>' +
      '<div class="linea3">Oficina General de Planificación</div>' +
      '</div>' +
      '</div>' +

      // TÍTULO + REFERENCIA DEL ACTA
      '<div class="title-band"><h1>Acta de Reunión</h1></div>' +
      '<div class="acta-code">' +
      '<span class="num">N° ' + numeroActa + '-' + anio + '-OR-OGPL/UNMSM</span>' +
      'Emisión: ' + fechaEmision +
      '</div>' +

      // DETALLES DE LA REUNIÓN
      '<table class="info-table">' +
      '<tr><td class="label">Tema</td><td class="value">' + datos.tema + '</td></tr>' +
      '<tr><td class="label">Modalidad</td><td class="value">' + datos.modalidad + '</td></tr>' +
      '<tr><td class="label">Fecha</td><td class="value">' + datos.fecha + '</td></tr>' +
      '<tr><td class="label">Lugar de reunión</td><td class="value">' + datos.lugar + '</td></tr>' +
      '<tr><td class="label">Horario</td><td class="value">' + datos.horaInicio + ' &ndash; ' + datos.horaFin + '</td></tr>' +
      '</table>' +

      // PARTICIPANTES
      '<div class="section-title">Participantes</div>' +
      '<table class="asistentes">' +
      '<tr>' +
      '<th style="width:8mm;">N°</th>' +
      '<th>Nombre y apellidos</th>' +
      '<th style="width:30mm;">Cargo / Unidad</th>' +
      '<th style="width:54mm;">Firma</th>' +
      '</tr>';
    
    datos.asistentes.forEach(function(asis, index) {
      var firmaContent = '';
      if (asis.usarFirmaDigital) {
        var apellidosMayus = asis.apellidos.toUpperCase();
        var nombresCap = asis.nombres.split(' ').map(function(w){ return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase(); }).join(' ');
        var fechaStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yy');
        var horaStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');

        firmaContent = '<span class="sello">' +
                       '<img src="' + LOGO_UNMSM + '" alt="">' +
                       '<span>' +
                       'Firmado digitalmente por<br>' +
                       '<b>' + apellidosMayus + ' ' + nombresCap + '</b><br>' +
                       'Motivo: Soy el Autor de la Firma<br>' +
                       'Fecha: ' + fechaStr + ' Hora: ' + horaStr +
                       '</span></span>';
      } else if (asis.firma) {
        firmaContent = '<img class="firma-img" src="' + asis.firma + '" alt="Firma"/>';
      }

      htmlStr += '<tr>' +
        '<td class="n">' + (index + 1) + '</td>' +
        '<td>' + asis.nombres + ' ' + asis.apellidos + '</td>' +
        '<td style="text-align:center;">' + asis.cargo + '<br>' + asis.unidad + '</td>' +
        '<td class="firma">' + firmaContent + '</td></tr>';
    });
    htmlStr += '</table>';

    // AGENDA
    htmlStr += '<div class="section-title">Agenda tratada</div>' +
      '<ul class="agenda-list">' +
      (datos.agenda.length > 0 ?
        datos.agenda.map(function(item, idx) { return '<li><span class="idx">' + (idx + 1) + '</span><span>' + item + '</span></li>'; }).join('') :
        '<li><span class="idx">1</span><span>&mdash;</span></li>') +
      '</ul>';

    // ACUERDOS (SIN PLAZO)
    if (datos.acuerdos.length > 0) {
      htmlStr += '<div class="section-title">Acuerdos</div>' +
        '<table class="registro-table">' +
        '<tr>' +
        '<th style="width:8mm;">N°</th>' +
        '<th>Acuerdo</th>' +
        '<th style="width:40mm;">Responsable</th>' +
        '</tr>';
      datos.acuerdos.forEach(function(ac, idx) {
        htmlStr += '<tr>' +
          '<td class="n">' + (idx + 1) + '</td>' +
          '<td>' + ac.texto + '</td>' +
          '<td>' + ac.responsable + '</td>' +
          '</tr>';
      });
      htmlStr += '</table>';
    }

    // COMPROMISOS (CON PLAZO)
    if (datos.compromisos.length > 0) {
      htmlStr += '<div class="section-title">Compromisos</div>' +
        '<table class="registro-table">' +
        '<tr>' +
        '<th style="width:8mm;">N°</th>' +
        '<th>Compromiso</th>' +
        '<th style="width:32mm;">Responsable</th>' +
        '<th style="width:22mm;">Plazo</th>' +
        '<th style="width:30mm;">Fecha límite</th>' +
        '</tr>';

      datos.compromisos.forEach(function(co, idx) {
        var fLimite = calcularFechaLimite(now, co.plazo, co.unidad);
        var fStr = Utilities.formatDate(fLimite, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
        htmlStr += '<tr>' +
          '<td class="n">' + (idx + 1) + '</td>' +
          '<td>' + co.texto + '</td>' +
          '<td>' + co.responsable + '</td>' +
          '<td>' + co.plazo + ' ' + co.unidad + '</td>' +
          '<td>' + fStr + '</td>' +
          '</tr>';
      });
      htmlStr += '</table>';
    }

    // EVIDENCIAS FOTOGRÁFICAS
    if (datos.fotos && datos.fotos.length > 0) {
      htmlStr += '<div class="page-break"></div>' +
        '<div class="anexo-titulo">Anexo fotográfico de la reunión</div>' +
        '<div style="width: 100%;">';
      datos.fotos.forEach(function(foto, idx) {
        htmlStr += '<div class="foto-container">' +
          '<img src="' + foto + '" style="max-width: 100%; max-height: 250px; display: block; margin-bottom: 5px; margin-left: auto; margin-right: auto;"/>' +
          '<span style="font-size: 8px; color: #666;">Evidencia N° ' + (idx + 1) + '</span>' +
          '</div>';
      });
      htmlStr += '</div>';
    }
    
    // PIE DE PÁGINA (Con QR dinámico)
    htmlStr += '<div class="footer">' +
      '<img class="qr" src="' + qrUrl + '" alt="Código QR de verificación">' +
      '<div class="txt">' +
      'Documento certificado y generado digitalmente por el Sistema Integral de Actas (SIGA). Este documento es ' +
      'original y tiene validez conforme a la normativa vigente.<br>' +
      'Oficina de Racionalización / Oficina General de Planificación &ndash; UNMSM. Fecha de emisión: ' + fechaEmision + '.' +
      '</div>' +
      '<div class="pag">Pág. 1</div>' +
      '</div>' +

      '</div></body></html>';
    
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
