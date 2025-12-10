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

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==========================================
// 2. HELPERS DE BASE DE DATOS
// ==========================================
function getSheetData(sheetName) {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getDataRange().getValues();
}

function getRootFolderId() {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const data = ss.getSheetByName('CONFIG').getDataRange().getValues();
  for(let row of data) { if(row[0] === 'ROOT_FOLDER_ID') return row[1]; }
  return null;
}
function crearCarpeta(nombre, idPadre) { return DriveApp.getFolderById(idPadre).createFolder(nombre).getId(); }

// ==========================================
// 3. API PARA VISTAS (Listas y Detalles)
// ==========================================
function getListaCampus() {
  const data = getSheetData('CAMPUS');
  return data.slice(1).map(r => ({ id: r[0], nombre: r[1] }));
}

function getEdificiosPorCampus(idCampus) {
  const data = getSheetData('EDIFICIOS');
  return data.slice(1)
    .filter(r => String(r[1]) === String(idCampus))
    .map(r => ({ id: r[0], nombre: r[2] }));
}

function getActivosPorEdificio(idEdificio) {
  const data = getSheetData('ACTIVOS');
  return data.slice(1)
    .filter(r => String(r[1]) === String(idEdificio))
    .map(r => ({ id: r[0], nombre: r[3], tipo: r[2], marca: r[4] }));
}

function getAssetInfo(idActivo) {
  const data = getSheetData('ACTIVOS');
  for(let i=1; i<data.length; i++){
    if(String(data[i][0]) === String(idActivo)){
      let fechaStr = '-';
      if (data[i][5] && data[i][5] instanceof Date) {
        fechaStr = Utilities.formatDate(data[i][5], Session.getScriptTimeZone(), "dd/MM/yyyy");
      }
      return {
        id: data[i][0],
        nombre: data[i][3],
        tipo: data[i][2],
        marca: data[i][4],
        fechaAlta: fechaStr
      };
    }
  }
  throw new Error("Activo no encontrado.");
}

// EDITAR ACTIVO
function updateAsset(datos) {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const sheet = ss.getSheetByName('ACTIVOS');
  const data = sheet.getDataRange().getValues();
  
  for(let i=1; i<data.length; i++){
    if(String(data[i][0]) === String(datos.id)) {
      sheet.getRange(i+1, 3).setValue(datos.tipo);   
      sheet.getRange(i+1, 4).setValue(datos.nombre); 
      sheet.getRange(i+1, 5).setValue(datos.marca);  
      return { success: true };
    }
  }
  return { success: false, error: "ID no encontrado" };
}

// ==========================================
// 4. API PARA TABLAS COMPLETAS
// ==========================================
function getTableData(tipo) {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  if (tipo === 'CAMPUS') {
    const data = getSheetData('CAMPUS');
    return data.slice(1).map(r => ({ id: r[0], nombre: r[1], provincia: r[2], direccion: r[3] }));
  }
  if (tipo === 'EDIFICIOS') {
    const data = getSheetData('EDIFICIOS');
    const dataC = getSheetData('CAMPUS');
    const mapCampus = {};
    dataC.slice(1).forEach(r => mapCampus[r[0]] = r[1]);
    return data.slice(1).map(r => ({ id: r[0], campus: mapCampus[r[1]] || '-', nombre: r[2], contacto: r[3] }));
  }
  return [];
}

// ==========================================
// 5. FUNCIONES DE NEGOCIO
// ==========================================

// --- DOCUMENTOS ---
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
    if(!targetFolderId) throw new Error("Carpeta no encontrada.");
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
  const data = getSheetData('DOCS_HISTORICO');
  const res = [];
  for(let i = data.length - 1; i >= 1; i--){
    if(String(data[i][2]) === String(idEntidad)) {
      let fecha = data[i][6] instanceof Date ? Utilities.formatDate(data[i][6], Session.getScriptTimeZone(), "dd/MM/yyyy") : String(data[i][6]);
      res.push({ 
        id: data[i][0], // <--- ID del documento para poder borrarlo
        nombre: data[i][3], 
        url: data[i][4], 
        version: data[i][5], 
        fecha: fecha 
      });
    }
  }
  return res;
}

// *** NUEVO: ELIMINAR DOCUMENTO ***
function eliminarDocumento(idDoc) {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const sheet = ss.getSheetByName('DOCS_HISTORICO');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(idDoc)) {
      // Borrar de Drive si es posible
      const idDrive = data[i][8]; 
      if (idDrive) {
        try { DriveApp.getFileById(idDrive).setTrashed(true); } catch(e) { console.warn("No se pudo borrar de Drive: " + e); }
      }
      // Borrar fila (i+1 porque es base 1)
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: "Documento no encontrado" };
}

// --- MANTENIMIENTO ---
function obtenerPlanMantenimiento(idActivo) {
  const data = getSheetData('PLAN_MANTENIMIENTO');
  const planes = [];
  for(let i=1; i<data.length; i++) {
    if(String(data[i][1]) === String(idActivo)) {
      let f = data[i][4];
      let fechaStr = (f instanceof Date) ? Utilities.formatDate(f, Session.getScriptTimeZone(), "yyyy-MM-dd") : "";
      planes.push({ id: data[i][0], tipo: data[i][2], fechaProxima: fechaStr });
    }
  }
  return planes;
}

function crearRevision(datos) { 
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  ss.getSheetByName('PLAN_MANTENIMIENTO').appendRow([Utilities.getUuid(), datos.idActivo, datos.tipo, "", new Date(datos.fechaProx), 365, "ACTIVO"]);
  return { success: true };
}

// --- CONTRATOS ---
function obtenerContratos(idEntidad) {
  const data = getSheetData('CONTRATOS');
  const contratos = [];
  const hoy = new Date();
  for(let i=1; i<data.length; i++) {
    if(String(data[i][2]) === String(idEntidad)) {
      const fFin = new Date(data[i][6]);
      const fIni = new Date(data[i][5]);
      let estado = 'VIGENTE'; let color = 'verde';
      const diffDays = Math.ceil((fFin - hoy) / 86400000);
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

function crearContrato(d) {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  ss.getSheetByName('CONTRATOS').appendRow([Utilities.getUuid(), d.tipoEntidad, d.idEntidad, d.proveedor, d.ref, new Date(d.fechaIni), new Date(d.fechaFin)]);
  return { success: true };
}

// ==========================================
// 6. DASHBOARD & CREACIÓN
// ==========================================
function getDatosDashboard() {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const hoy = new Date();
  const cAct = (ss.getSheetByName('ACTIVOS').getLastRow() - 1) || 0;
  const cEdif = (ss.getSheetByName('EDIFICIOS').getLastRow() - 1) || 0;
  const dataMant = getSheetData('PLAN_MANTENIMIENTO');
  let revPend = 0, revVenc = 0, revOk = 0;
  for(let i=1; i<dataMant.length; i++) {
    const diff = Math.ceil((new Date(dataMant[i][4]) - hoy) / 86400000);
    if(diff < 0) revVenc++; else if(diff <= 30) revPend++; else revOk++;
  }
  const dataCont = getSheetData('CONTRATOS');
  let contCad = 0;
  for(let i=1; i<dataCont.length; i++) if(new Date(dataCont[i][6]) < hoy) contCad++;
  return { activos: cAct, edificios: cEdif, pendientes: revPend, vencidas: revVenc, ok: revOk, contratosCaducados: contCad };
}

function crearCampus(d) { 
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const fId = crearCarpeta(d.nombre, getRootFolderId());
  ss.getSheetByName('CAMPUS').appendRow([Utilities.getUuid(), d.nombre, d.provincia, d.direccion, fId]);
  return {success:true};
}
function crearEdificio(d) {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const cData = getSheetData('CAMPUS');
  let pId; for(let i=1; i<cData.length; i++) if(String(cData[i][0])==String(d.idCampus)) pId=cData[i][4];
  const fId = crearCarpeta(d.nombre, pId);
  const aId = crearCarpeta("Activos", fId);
  ss.getSheetByName('EDIFICIOS').appendRow([Utilities.getUuid(), d.idCampus, d.nombre, d.contacto, fId, aId]);
  return {success:true};
}
function crearActivo(d) {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const eData = getSheetData('EDIFICIOS');
  let pId; for(let i=1; i<eData.length; i++) if(String(eData[i][0])==String(d.idEdificio)) pId=eData[i][5];
  const fId = crearCarpeta(d.nombre, pId);
  const id = Utilities.getUuid();
  ss.getSheetByName('ACTIVOS').appendRow([id, d.idEdificio, d.tipo, d.nombre, d.marca, new Date(), fId]);
  return {success:true};
}
function getCatalogoInstalaciones() { return getSheetData('CAT_INSTALACIONES').slice(1).map(r=>({id:r[0], nombre:r[1], dias:r[3]})); }
