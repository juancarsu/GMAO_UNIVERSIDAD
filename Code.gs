// ==========================================
// 1. CONFIGURACIÓN Y ROUTING
// ==========================================
const PROPS = PropertiesService.getScriptProperties();

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate().setTitle('GMAO Universidad').addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }

// ==========================================
// 2. HELPERS DE BASE DE DATOS
// ==========================================
function getSheetData(sheetName) {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  // Leemos todo el rango de datos
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
// 3. API PARA VISTAS
// ==========================================
function getListaCampus() { const data = getSheetData('CAMPUS'); return data.slice(1).map(r => ({ id: r[0], nombre: r[1] })); }
function getEdificiosPorCampus(idCampus) { const data = getSheetData('EDIFICIOS'); return data.slice(1).filter(r => String(r[1]) === String(idCampus)).map(r => ({ id: r[0], nombre: r[2] })); }
function getActivosPorEdificio(idEdificio) { const data = getSheetData('ACTIVOS'); return data.slice(1).filter(r => String(r[1]) === String(idEdificio)).map(r => ({ id: r[0], nombre: r[3], tipo: r[2], marca: r[4] })); }

function getAssetInfo(idActivo) {
  const data = getSheetData('ACTIVOS');
  for(let i=1; i<data.length; i++){
    if(String(data[i][0]) === String(idActivo)){
      let fechaStr = '-';
      if (data[i][5] && data[i][5] instanceof Date) { fechaStr = Utilities.formatDate(data[i][5], Session.getScriptTimeZone(), "dd/MM/yyyy"); }
      return { id: data[i][0], nombre: data[i][3], tipo: data[i][2], marca: data[i][4], fechaAlta: fechaStr };
    }
  }
  throw new Error("Activo no encontrado.");
}

function updateAsset(datos) {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const sheet = ss.getSheetByName('ACTIVOS');
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++){
    if(String(data[i][0]) === String(datos.id)) {
      sheet.getRange(i+1, 3).setValue(datos.tipo); sheet.getRange(i+1, 4).setValue(datos.nombre); sheet.getRange(i+1, 5).setValue(datos.marca); return { success: true };
    }
  }
  return { success: false, error: "ID no encontrado" };
}

// ==========================================
// 4. FUNCIONES DE NEGOCIO
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
    for(let i=1; i<entityData.length; i++){ if(String(entityData[i][0]) === String(idEntidad)) { targetFolderId = entityData[i][colIndex]; break; } }
    if(!targetFolderId) throw new Error("Carpeta no encontrada.");
    const sheetDocs = ss.getSheetByName('DOCS_HISTORICO');
    const docsData = sheetDocs.getDataRange().getValues();
    let maxVer = 0;
    for(let i=1; i<docsData.length; i++){ if(String(docsData[i][2]) === String(idEntidad) && docsData[i][3] === nombre) { if(docsData[i][5] > maxVer) maxVer = docsData[i][5]; } }
    const newVer = maxVer + 1;
    const blob = Utilities.newBlob(Utilities.base64Decode(base64), mime, `[v${newVer}] ${nombre}`);
    const file = DriveApp.getFolderById(targetFolderId).createFile(blob);
    sheetDocs.appendRow([Utilities.getUuid(), tipoEntidad, idEntidad, nombre, file.getUrl(), newVer, new Date(), Session.getActiveUser().getEmail(), file.getId()]);
    return { success: true, version: newVer };
  } catch(e) { return { success: false, error: e.toString() }; } finally { lock.releaseLock(); }
}

function obtenerDocs(idEntidad) {
  const data = getSheetData('DOCS_HISTORICO'); const res = [];
  for(let i = data.length - 1; i >= 1; i--){
    if(String(data[i][2]) === String(idEntidad)) {
      let fecha = data[i][6] instanceof Date ? Utilities.formatDate(data[i][6], Session.getScriptTimeZone(), "dd/MM/yyyy") : String(data[i][6]);
      res.push({ id: data[i][0], nombre: data[i][3], url: data[i][4], version: data[i][5], fecha: fecha });
    }
  }
  return res;
}

function eliminarDocumento(idDoc) {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const sheet = ss.getSheetByName('DOCS_HISTORICO'); const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(idDoc)) {
      const idDrive = data[i][8]; if (idDrive) { try { DriveApp.getFileById(idDrive).setTrashed(true); } catch(e) { console.warn(e); } }
      sheet.deleteRow(i + 1); return { success: true };
    }
  }
  return { success: false, error: "Documento no encontrado" };
}

// *** MANTENIMIENTO BLINDADO ***
// *** MANTENIMIENTO (SÚPER ROBUSTO) ***
function obtenerPlanMantenimiento(idActivo) {
  const data = getSheetData('PLAN_MANTENIMIENTO'); 
  const planes = [];
  const hoy = new Date();
  
  // Normalizamos "hoy" para comparar solo fechas (sin horas)
  hoy.setHours(0,0,0,0);
  
  for(let i=1; i<data.length; i++) {
    if(String(data[i][1]) === String(idActivo)) {
      const rawDate = data[i][4]; // Columna E (Fecha Próxima)
      let fechaObj = null;
      let color = 'gris'; // Color por defecto si falla
      let fechaStr = "--/--/----";

      // 1. INTENTO DE INTERPRETACIÓN DE FECHA
      if (rawDate instanceof Date) {
        // Es un objeto fecha de Google Sheets
        fechaObj = new Date(rawDate);
      } else if (typeof rawDate === 'string' && rawDate.includes('/')) {
        // Es texto tipo "19/12/2025" -> Lo convertimos a mano
        const partes = rawDate.split('/'); // [19, 12, 2025]
        // new Date(año, mes-1, dia)
        if(partes.length === 3) fechaObj = new Date(partes[2], partes[1]-1, partes[0]);
      }

      // 2. CÁLCULO DE SEMÁFORO
      if (fechaObj && !isNaN(fechaObj.getTime())) {
         fechaObj.setHours(0,0,0,0); // Quitamos horas para comparar
         
         // Diferencia en milisegundos
         const diffTime = fechaObj.getTime() - hoy.getTime();
         // Diferencia en días
         const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
         
         if (diffDays < 0) color = 'rojo';        // Vencida
         else if (diffDays <= 30) color = 'amarillo'; // Próxima (30 días)
         else color = 'verde';                    // Lejana
         
         fechaStr = Utilities.formatDate(fechaObj, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
      
      planes.push({ 
          id: data[i][0], 
          tipo: data[i][2], 
          fechaProxima: fechaStr,
          color: color 
      });
    }
  }
  return planes;
}

function crearRevision(datos) { const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); ss.getSheetByName('PLAN_MANTENIMIENTO').appendRow([Utilities.getUuid(), datos.idActivo, datos.tipo, "", new Date(datos.fechaProx), 365, "ACTIVO"]); return { success: true }; }
function updateRevision(datos) { const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const sheet = ss.getSheetByName('PLAN_MANTENIMIENTO'); const data = sheet.getDataRange().getValues(); for(let i=1; i<data.length; i++){ if(String(data[i][0]) === String(datos.idPlan)) { sheet.getRange(i+1, 3).setValue(datos.tipo); sheet.getRange(i+1, 5).setValue(new Date(datos.fechaProx)); return { success: true }; } } return { success: false, error: "Plan no encontrado" }; }
function eliminarRevision(idPlan) { const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const sheet = ss.getSheetByName('PLAN_MANTENIMIENTO'); const data = sheet.getDataRange().getValues(); for(let i=1; i<data.length; i++){ if(String(data[i][0]) === String(idPlan)) { sheet.deleteRow(i+1); return { success: true }; } } return { success: false, error: "Plan no encontrado" }; }

// *** CONTRATOS BLINDADO (Lectura segura de columnas) ***
function obtenerContratos(idEntidad) {
  // Forzamos leer hasta la columna H (8 columnas)
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const sheet = ss.getSheetByName('CONTRATOS');
  // Si no hay datos, devolver vacío
  if (!sheet || sheet.getLastRow() < 2) return [];
  
  // Obtenemos rango específico para asegurar ancho
  const data = sheet.getRange(1, 1, sheet.getLastRow(), 8).getValues();
  
  const contratos = [];
  const hoy = new Date();
  
  for(let i=1; i<data.length; i++) {
    if(String(data[i][2]) === String(idEntidad)) {
      const fFin = data[i][6] instanceof Date ? data[i][6] : null;
      const fIni = data[i][5] instanceof Date ? data[i][5] : null;
      
      // Acceso seguro a la columna 7 (Estado)
      let estadoDB = (data[i].length > 7) ? data[i][7] : 'ACTIVO';
      if(!estadoDB) estadoDB = 'ACTIVO'; // Si está vacía, asumimos activo
      
      let estadoCalc = 'VIGENTE'; 
      let color = 'verde';
      
      if (estadoDB === 'INACTIVO') {
         estadoCalc = 'INACTIVO'; 
         color = 'gris';
      } else if (fFin) {
         const diffTime = fFin.getTime() - hoy.getTime();
         const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
         if (diffDays < 0) { estadoCalc = 'CADUCADO'; color = 'rojo'; }
         else if (diffDays <= 30) { estadoCalc = 'PRÓXIMO'; color = 'amarillo'; }
      } else {
         estadoCalc = 'SIN FECHA'; color = 'gris';
      }
      
      const iniStr = fIni ? Utilities.formatDate(fIni, Session.getScriptTimeZone(), "yyyy-MM-dd") : "";
      const finStr = fFin ? Utilities.formatDate(fFin, Session.getScriptTimeZone(), "yyyy-MM-dd") : "";

      contratos.push({
        id: data[i][0], 
        proveedor: data[i][3], 
        ref: data[i][4],
        inicio: iniStr,
        fin: finStr,
        estado: estadoCalc, 
        color: color,
        estadoDB: estadoDB
      });
    }
  }
  return contratos;
}

function crearContrato(d) { const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); ss.getSheetByName('CONTRATOS').appendRow([Utilities.getUuid(), d.tipoEntidad, d.idEntidad, d.proveedor, d.ref, new Date(d.fechaIni), new Date(d.fechaFin), d.estado]); return { success: true }; }
function updateContrato(datos) { const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const sheet = ss.getSheetByName('CONTRATOS'); const data = sheet.getDataRange().getValues(); for(let i=1; i<data.length; i++){ if(String(data[i][0]) === String(datos.id)) { sheet.getRange(i+1, 4).setValue(datos.proveedor); sheet.getRange(i+1, 5).setValue(datos.ref); sheet.getRange(i+1, 6).setValue(new Date(datos.fechaIni)); sheet.getRange(i+1, 7).setValue(new Date(datos.fechaFin)); sheet.getRange(i+1, 8).setValue(datos.estado); return { success: true }; } } return { success: false, error: "Contrato no encontrado" }; }
function eliminarContrato(id) { const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const sheet = ss.getSheetByName('CONTRATOS'); const data = sheet.getDataRange().getValues(); for(let i=1; i<data.length; i++){ if(String(data[i][0]) === String(id)) { sheet.deleteRow(i+1); return { success: true }; } } return { success: false, error: "Contrato no encontrado" }; }

// ==========================================
// 6. DASHBOARD & CREACIÓN
// ==========================================
function getDatosDashboard() {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const hoy = new Date();
  const cAct = (ss.getSheetByName('ACTIVOS').getLastRow() - 1) || 0; const cEdif = (ss.getSheetByName('EDIFICIOS').getLastRow() - 1) || 0;
  const dataMant = getSheetData('PLAN_MANTENIMIENTO'); let revPend = 0, revVenc = 0, revOk = 0;
  for(let i=1; i<dataMant.length; i++) { const f = dataMant[i][4]; if(f instanceof Date) { const diff = Math.ceil((f - hoy) / 86400000); if(diff < 0) revVenc++; else if(diff <= 30) revPend++; else revOk++; } }
  const dataCont = getSheetData('CONTRATOS'); let contCad = 0;
  for(let i=1; i<dataCont.length; i++) if(dataCont[i][6] instanceof Date && dataCont[i][6] < hoy) contCad++;
  return { activos: cAct, edificios: cEdif, pendientes: revPend, vencidas: revVenc, ok: revOk, contratosCaducados: contCad };
}
function crearCampus(d) { const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const fId = crearCarpeta(d.nombre, getRootFolderId()); ss.getSheetByName('CAMPUS').appendRow([Utilities.getUuid(), d.nombre, d.provincia, d.direccion, fId]); return {success:true}; }
function crearEdificio(d) { const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const cData = getSheetData('CAMPUS'); let pId; for(let i=1; i<cData.length; i++) if(String(cData[i][0])==String(d.idCampus)) pId=cData[i][4]; const fId = crearCarpeta(d.nombre, pId); const aId = crearCarpeta("Activos", fId); ss.getSheetByName('EDIFICIOS').appendRow([Utilities.getUuid(), d.idCampus, d.nombre, d.contacto, fId, aId]); return {success:true}; }
function crearActivo(d) { const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const eData = getSheetData('EDIFICIOS'); let pId; for(let i=1; i<eData.length; i++) if(String(eData[i][0])==String(d.idEdificio)) pId=eData[i][5]; const fId = crearCarpeta(d.nombre, pId); const id = Utilities.getUuid(); const cats = getSheetData('CAT_INSTALACIONES'); let nombreTipo = d.tipo; for(let i=1; i<cats.length; i++) { if(String(cats[i][0]) === String(d.tipo)) { nombreTipo = cats[i][1]; break; } } ss.getSheetByName('ACTIVOS').appendRow([id, d.idEdificio, nombreTipo, d.nombre, d.marca, new Date(), fId]); return {success:true}; }
function getCatalogoInstalaciones() { return getSheetData('CAT_INSTALACIONES').slice(1).map(r=>({id:r[0], nombre:r[1], dias:r[3]})); }
