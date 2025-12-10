// ==========================================
// 1. CONFIGURACIÓN Y ROUTING
// ==========================================
const PROPS = PropertiesService.getScriptProperties();

function doGet(e) {
  const dbId = PROPS.getProperty('DB_SS_ID');
  if (!dbId) {
    return HtmlService.createTemplateFromFile('Setup')
      .evaluate().setTitle('Instalación GMAO').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createTemplateFromFile('Index')
      .evaluate().setTitle('GMAO Universidad').addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// *** NUEVA FUNCIÓN IMPORTANTE ***
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==========================================
// 2. INSTALACIÓN (SETUP)
// ==========================================
function instalarSistema(urlSheet) {
  try {
    let ss;
    try { ss = SpreadsheetApp.openByUrl(urlSheet); } 
    catch(e) { ss = SpreadsheetApp.openById(urlSheet); }
    const ssId = ss.getId();
    
    const tabs = {
      'CONFIG': ['Clave', 'Valor', 'Descripcion'],
      'CAMPUS': ['ID', 'Nombre', 'Provincia', 'Direccion', 'ID_Carpeta_Drive'],
      'EDIFICIOS': ['ID', 'ID_Campus', 'Nombre', 'Contacto', 'ID_Carpeta_Drive', 'ID_Carpeta_Activos'],
      'CAT_INSTALACIONES': ['ID', 'Nombre', 'Normativa_Ref', 'Periodicidad_Dias'],
      'ACTIVOS': ['ID', 'ID_Edificio', 'Tipo', 'Nombre', 'Marca', 'Fecha_Alta', 'ID_Carpeta_Drive'],
      'DOCS_HISTORICO': ['ID_Doc', 'Tipo_Entidad', 'ID_Entidad', 'Nombre_Archivo', 'URL', 'Version', 'Fecha', 'Usuario', 'ID_Archivo_Drive'],
      'PLAN_MANTENIMIENTO': ['ID_Plan', 'ID_Activo', 'Tipo_Revision', 'Fecha_Ultima', 'Fecha_Proxima', 'Periodicidad_Dias', 'Estado'],
      'CONTRATOS': ['ID_Contrato', 'Tipo_Entidad', 'ID_Entidad', 'Proveedor', 'Num_Ref', 'Fecha_Inicio', 'Fecha_Fin']
    };

    const sheets = ss.getSheets().map(s => s.getName());
    for (const [name, headers] of Object.entries(tabs)) {
      let sheet = sheets.includes(name) ? ss.getSheetByName(name) : ss.insertSheet(name);
      if (sheet.getLastRow() === 0) {
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#ddd');
      }
    }

    const sheetConfig = ss.getSheetByName('CONFIG');
    const dataConfig = sheetConfig.getDataRange().getValues();
    let rootId = null;
    for(let row of dataConfig) { if(row[0] === 'ROOT_FOLDER_ID') rootId = row[1]; }

    if(!rootId) {
      const rootFolder = DriveApp.createFolder("GMAO UNIVERSIDAD - GESTIÓN DOCUMENTAL");
      rootId = rootFolder.getId();
      sheetConfig.appendRow(['ROOT_FOLDER_ID', rootId, 'Raíz del sistema']);
    }
    
    const emailAdmin = Session.getActiveUser().getEmail();
    sheetConfig.appendRow(['EMAIL_AVISOS', emailAdmin, 'Email para alertas semanales']);

    PROPS.setProperty('DB_SS_ID', ssId);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ==========================================
// 3. LÓGICA DE DRIVE
// ==========================================
function getRootFolderId() {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const data = ss.getSheetByName('CONFIG').getDataRange().getValues();
  for(let row of data) { if(row[0] === 'ROOT_FOLDER_ID') return row[1]; }
  return null;
}

function crearCarpeta(nombre, idPadre) {
  return DriveApp.getFolderById(idPadre).createFolder(nombre).getId();
}

// ==========================================
// 4. CATÁLOGO Y NORMATIVA
// ==========================================
function getCatalogoInstalaciones() {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const sheet = ss.getSheetByName('CAT_INSTALACIONES');
  
  if (sheet.getLastRow() < 2) {
    sheet.appendRow([Utilities.getUuid(), 'Baja Tensión (REBT)', 'RD 842/2002', 1825]);
    sheet.appendRow([Utilities.getUuid(), 'Ascensores', 'ITC AEM 1', 30]);
    sheet.appendRow([Utilities.getUuid(), 'Instalaciones Térmicas (RITE)', 'RD 1027/2007', 365]);
    sheet.appendRow([Utilities.getUuid(), 'Protección Incendios (PCI)', 'RIPCI', 90]);
  }
  
  const data = sheet.getDataRange().getValues();
  const catalogo = [];
  for(let i=1; i<data.length; i++){
    catalogo.push({ id: data[i][0], nombre: data[i][1], norma: data[i][2], dias: data[i][3] });
  }
  return catalogo;
}

// ==========================================
// 5. CREACIÓN DE ENTIDADES
// ==========================================
function crearCampus(datos) { 
  const rootId = getRootFolderId();
  const folderId = crearCarpeta(datos.nombre, rootId);
  const id = Utilities.getUuid();
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  ss.getSheetByName('CAMPUS').appendRow([id, datos.nombre, datos.provincia, datos.direccion, folderId]);
  return { success: true };
}

function crearEdificio(datos) { 
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const campData = ss.getSheetByName('CAMPUS').getDataRange().getValues();
  let parentFolderId = null;
  for(let i=1; i<campData.length; i++){
    if(String(campData[i][0]) === String(datos.idCampus)) { parentFolderId = campData[i][4]; break; }
  }
  if(!parentFolderId) return { success: false, error: "Campus no encontrado" };
  
  const edifFolderId = crearCarpeta(datos.nombre, parentFolderId);
  const activosFolderId = crearCarpeta("Activos", edifFolderId);
  const id = Utilities.getUuid();
  ss.getSheetByName('EDIFICIOS').appendRow([id, datos.idCampus, datos.nombre, datos.contacto, edifFolderId, activosFolderId]);
  return { success: true };
}

function crearActivo(datos) { 
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const edifData = ss.getSheetByName('EDIFICIOS').getDataRange().getValues();
  let parentFolderId = null;
  for(let i=1; i<edifData.length; i++){
    if(String(edifData[i][0]) === String(datos.idEdificio)) { parentFolderId = edifData[i][5]; break; } 
  }
  const assetFolderId = crearCarpeta(datos.nombre, parentFolderId);
  const idActivo = Utilities.getUuid();
  
  const sheetCat = ss.getSheetByName('CAT_INSTALACIONES');
  const catData = sheetCat.getDataRange().getValues();
  let diasRevision = 365;
  let nombreTipo = datos.tipo; 

  for(let i=1; i<catData.length; i++){
    if(String(catData[i][0]) === String(datos.tipo)) {
       nombreTipo = catData[i][1];
       diasRevision = catData[i][3];
       break;
    }
  }

  ss.getSheetByName('ACTIVOS').appendRow([idActivo, datos.idEdificio, nombreTipo, datos.nombre, datos.marca, new Date(), assetFolderId]);
  
  const sheetPlan = ss.getSheetByName('PLAN_MANTENIMIENTO');
  const fechaAlta = new Date();
  const fechaProx = new Date(fechaAlta.getTime() + (diasRevision * 24 * 60 * 60 * 1000));
  
  sheetPlan.appendRow([Utilities.getUuid(), idActivo, `Mantenimiento Legal (${nombreTipo})`, "", fechaProx, diasRevision, "ACTIVO"]);
  return { success: true };
}

// ==========================================
// 6. GESTIÓN DOCUMENTAL
// ==========================================
function subirArchivo(base64, nombre, mime, idEntidad, tipoEntidad) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
    
    let targetFolderId = null;
    let sheetName = (tipoEntidad === 'EDIFICIO') ? 'EDIFICIOS' : 'ACTIVOS';
    let colIndex = (tipoEntidad === 'EDIFICIO') ? 4 : 6;
    const entityData = ss.getSheetByName(sheetName).getDataRange().getValues();
    for(let i=1; i<entityData.length; i++){
      if(String(entityData[i][0]) === String(idEntidad)) { targetFolderId = entityData[i][colIndex]; break; }
    }
    
    if(!targetFolderId) throw new Error("Carpeta no encontrada");
    const sheetDocs = ss.getSheetByName('DOCS_HISTORICO');
    const docsData = sheetDocs.getDataRange().getValues();
    let maxVer = 0;
    for(let i=1; i<docsData.length; i++){
      if(String(docsData[i][2]) === String(idEntidad) && docsData[i][3] === nombre) {
        if(docsData[i][5] > maxVer) maxVer = docsData[i][5];
      }
    }
    const newVer = maxVer + 1;
    
    const blob = Utilities.newBlob(Utilities.base64Decode(base64), mime, `[v${newVer}] ${nombre}`);
    const file = DriveApp.getFolderById(targetFolderId).createFile(blob);
    
    sheetDocs.appendRow([Utilities.getUuid(), tipoEntidad, idEntidad, nombre, file.getUrl(), newVer, new Date(), Session.getActiveUser().getEmail(), file.getId()]);
    
    return { success: true, version: newVer };
  } catch(e) { return { success: false, error: e.toString() }; } finally { lock.releaseLock(); }
}

function obtenerDocs(idEntidad) {
  try {
    const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
    const sheet = ss.getSheetByName('DOCS_HISTORICO');
    if (sheet.getLastRow() < 2) return [];
    
    const data = sheet.getDataRange().getValues();
    const res = [];
    for(let i = data.length - 1; i >= 1; i--){
      if(String(data[i][2]) === String(idEntidad)) {
        let fechaFormateada = "";
        try {
          let fechaRaw = data[i][6];
          fechaFormateada = (fechaRaw instanceof Date) ? Utilities.formatDate(fechaRaw, Session.getScriptTimeZone(), "dd/MM/yyyy") : String(fechaRaw);
        } catch(e) { fechaFormateada = "--"; }
        res.push({ nombre: data[i][3], url: data[i][4], version: data[i][5], fecha: fechaFormateada });
      }
    }
    return res;
  } catch (e) { throw new Error(e.toString()); }
}

// ==========================================
// 7. GESTIÓN MANTENIMIENTO
// ==========================================
function obtenerPlanMantenimiento(idActivo) {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const sheet = ss.getSheetByName('PLAN_MANTENIMIENTO');
  if(!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getDataRange().getValues();
  const planes = [];
  for(let i=1; i<data.length; i++) {
    if(String(data[i][1]) === String(idActivo)) {
      let fProx = data[i][4];
      let fechaStr = (fProx instanceof Date) ? Utilities.formatDate(fProx, Session.getScriptTimeZone(), "yyyy-MM-dd") : "";
      planes.push({ id: data[i][0], tipo: data[i][2], fechaProxima: fechaStr });
    }
  }
  return planes;
}

function crearRevision(datos) { 
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const sheet = ss.getSheetByName('PLAN_MANTENIMIENTO');
  sheet.appendRow([Utilities.getUuid(), datos.idActivo, datos.tipo, "", new Date(datos.fechaProx), 365, "ACTIVO"]);
  return { success: true };
}

// ==========================================
// 8. GESTIÓN CONTRATOS
// ==========================================
function obtenerContratos(idEntidad) {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const sheet = ss.getSheetByName('CONTRATOS');
  if(!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getDataRange().getValues();
  const contratos = [];
  const hoy = new Date();
  for(let i=1; i<data.length; i++) {
    if(String(data[i][2]) === String(idEntidad)) {
      const fFin = data[i][6] instanceof Date ? data[i][6] : new Date(data[i][6]);
      const fIni = data[i][5] instanceof Date ? data[i][5] : new Date(data[i][5]);
      let estado = 'VIGENTE'; let color = 'verde';
      const diffDays = Math.ceil((fFin - hoy) / (1000 * 60 * 60 * 24));
      if (diffDays < 0) { estado = 'CADUCADO'; color = 'rojo'; }
      else if (diffDays <= 30) { estado = 'PRÓXIMO'; color = 'amarillo'; }
      contratos.push({
        id: data[i][0], proveedor: data[i][3], ref: data[i][4],
        inicio: Utilities.formatDate(fIni, Session.getScriptTimeZone(), "dd/MM/yyyy"),
        fin: Utilities.formatDate(fFin, Session.getScriptTimeZone(), "dd/MM/yyyy"),
        estado: estado, color: color
      });
    }
  }
  return contratos;
}

function crearContrato(datos) {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const sheet = ss.getSheetByName('CONTRATOS');
  sheet.appendRow([Utilities.getUuid(), datos.tipoEntidad, datos.idEntidad, datos.proveedor, datos.ref, new Date(datos.fechaIni), new Date(datos.fechaFin)]);
  return { success: true };
}

// ==========================================
// 9. AUTOMATIZACIÓN DE AVISOS Y DASHBOARD
// ==========================================
function getDatosDashboard() {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const hoy = new Date();
  const sheetAct = ss.getSheetByName('ACTIVOS');
  const totalActivos = sheetAct.getLastRow() - 1; 

  const sheetMant = ss.getSheetByName('PLAN_MANTENIMIENTO');
  const dataMant = sheetMant.getDataRange().getValues();
  let revPendientes = 0, revOk = 0, revVencidas = 0;
  for(let i=1; i<dataMant.length; i++) {
    const fecha = dataMant[i][4] instanceof Date ? dataMant[i][4] : new Date(dataMant[i][4]);
    const diffDays = Math.ceil((fecha - hoy) / (1000 * 60 * 60 * 24));
    if (diffDays < 0) revVencidas++; else if (diffDays <= 30) revPendientes++; else revOk++;
  }

  const sheetCont = ss.getSheetByName('CONTRATOS');
  let contCaducados = 0;
  if(sheetCont && sheetCont.getLastRow() > 1) {
    const dataCont = sheetCont.getDataRange().getValues();
    for(let i=1; i<dataCont.length; i++) {
      const fecha = dataCont[i][6] instanceof Date ? dataCont[i][6] : new Date(dataCont[i][6]);
      if (fecha < hoy) contCaducados++;
    }
  }

  return { activos: totalActivos > 0 ? totalActivos : 0, pendientes: revPendientes, vencidas: revVencidas, ok: revOk, contratosCaducados: contCaducados };
}

function enviarResumenSemanal() {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const sheetConfig = ss.getSheetByName('CONFIG');
  const dataConfig = sheetConfig.getDataRange().getValues();
  let emailDestino = "";
  for(let row of dataConfig) { if(row[0] === 'EMAIL_AVISOS') emailDestino = row[1]; }
  if(!emailDestino) emailDestino = Session.getActiveUser().getEmail();

  const hoy = new Date();
  let alertas = [];
  const sheetMant = ss.getSheetByName('PLAN_MANTENIMIENTO');
  const dataMant = sheetMant.getDataRange().getValues();
  const mapActivos = {};
  const dataAct = ss.getSheetByName('ACTIVOS').getDataRange().getValues();
  for(let i=1; i<dataAct.length; i++) { mapActivos[dataAct[i][0]] = dataAct[i][3]; } 

  for(let i=1; i<dataMant.length; i++) {
    const fecha = new Date(dataMant[i][4]);
    const diffDays = Math.ceil((fecha - hoy) / (1000 * 60 * 60 * 24));
    if(diffDays <= 30) {
      const nombreActivo = mapActivos[dataMant[i][1]] || "Desconocido";
      const estado = diffDays < 0 ? "🔴 VENCIDO" : "🟡 PRÓXIMO";
      alertas.push(`<li><strong>${estado}</strong>: ${nombreActivo} - ${dataMant[i][2]} (${Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM/yyyy")})</li>`);
    }
  }
  const sheetCont = ss.getSheetByName('CONTRATOS');
  if(sheetCont) {
    const dataCont = sheetCont.getDataRange().getValues();
    for(let i=1; i<dataCont.length; i++) {
      const fecha = new Date(dataCont[i][6]);
      if(fecha < hoy) alertas.push(`<li><strong>🔴 CONTRATO CADUCADO</strong>: Proveedor ${dataCont[i][3]} (Ref: ${dataCont[i][4]})</li>`);
    }
  }

  if(alertas.length > 0) {
    const htmlBody = `<h3>Resumen Semanal GMAO</h3><ul>${alertas.join('')}</ul>`;
    MailApp.sendEmail({ to: emailDestino, subject: "⚠️ Alerta GMAO: Vencimientos", htmlBody: htmlBody });
  }
}

// ==========================================
// 10. ÁRBOL UI
// ==========================================
function getArbolDatos() {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const campus = ss.getSheetByName('CAMPUS').getDataRange().getValues().slice(1);
  const edificios = ss.getSheetByName('EDIFICIOS').getDataRange().getValues().slice(1);
  const activos = ss.getSheetByName('ACTIVOS').getDataRange().getValues().slice(1);
  
  return campus.map(c => ({
    id: c[0], nombre: c[1], type: 'CAMPUS',
    hijos: edificios.filter(e => String(e[1]) === String(c[0])).map(e => ({
      id: e[0], nombre: e[2], type: 'EDIFICIO',
      hijos: activos.filter(a => String(a[1]) === String(e[0])).map(a => ({
        id: a[0], nombre: a[3], type: 'ACTIVO'
      }))
    }))
  }));
}
