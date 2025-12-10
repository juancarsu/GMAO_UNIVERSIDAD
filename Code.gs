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
// 2. HELPERS
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

// Convierte Objeto Date a Texto DD/MM/YYYY
function fechaATexto(dateObj) {
  if (!dateObj || !(dateObj instanceof Date)) return "-";
  return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "dd/MM/yyyy");
}

// Convierte YYYY-MM-DD a Objeto Date (Hora 12:00)
function textoAFecha(txt) {
  if (!txt) return null;
  var parts = String(txt).split('-');
  if (parts.length !== 3) return null;
  return new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10), 12, 0, 0);
}

// ==========================================
// 3. API VISTAS
// ==========================================
function getListaCampus() { const data = getSheetData('CAMPUS'); return data.slice(1).map(r => ({ id: r[0], nombre: r[1] })); }
function getEdificiosPorCampus(idCampus) { const data = getSheetData('EDIFICIOS'); return data.slice(1).filter(r => String(r[1]) === String(idCampus)).map(r => ({ id: r[0], nombre: r[2] })); }
function getActivosPorEdificio(idEdificio) { const data = getSheetData('ACTIVOS'); return data.slice(1).filter(r => String(r[1]) === String(idEdificio)).map(r => ({ id: r[0], nombre: r[3], tipo: r[2], marca: r[4] })); }

function getAssetInfo(idActivo) {
  const data = getSheetData('ACTIVOS');
  for(let i=1; i<data.length; i++){
    if(String(data[i][0]) === String(idActivo)){
      const f = data[i][5];
      const fechaStr = (f instanceof Date) ? Utilities.formatDate(f, Session.getScriptTimeZone(), "dd/MM/yyyy") : "-";
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
// 4. MANTENIMIENTO
// ==========================================

// Obtener Plan (Individual)
function obtenerPlanMantenimiento(idActivo) {
  const data = getSheetData('PLAN_MANTENIMIENTO'); 
  const planes = [];
  const hoy = new Date(); hoy.setHours(0,0,0,0);
  
  for(let i=1; i<data.length; i++) {
    if(String(data[i][1]) === String(idActivo)) {
      let f = data[i][4];
      let color = 'gris'; 
      let fechaStr = "-";
      let fechaISO = "";

      if (f instanceof Date) {
         f.setHours(0,0,0,0);
         const diff = Math.ceil((f.getTime() - hoy.getTime()) / (86400000));
         if (diff < 0) color = 'rojo'; else if (diff <= 30) color = 'amarillo'; else color = 'verde';
         fechaStr = Utilities.formatDate(f, Session.getScriptTimeZone(), "dd/MM/yyyy");
         fechaISO = Utilities.formatDate(f, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
      planes.push({ id: data[i][0], tipo: data[i][2], fechaProxima: fechaStr, fechaISO: fechaISO, color: color });
    }
  }
  
  return planes.sort((a, b) => a.fechaISO.localeCompare(b.fechaISO));
}

// Obtener Plan (Global)
function getGlobalMaintenance() {
  const planes = getSheetData('PLAN_MANTENIMIENTO');
  const activos = getSheetData('ACTIVOS');
  const edificios = getSheetData('EDIFICIOS');
  
  const mapActivos = {}; activos.slice(1).forEach(r => mapActivos[r[0]] = { nombre: r[3], idEdif: r[1] });
  const mapEdificios = {}; edificios.slice(1).forEach(r => mapEdificios[r[0]] = r[2]);
  
  const result = [];
  const hoy = new Date(); hoy.setHours(0,0,0,0);
  
  for(let i=1; i<planes.length; i++) {
    const idActivo = planes[i][1];
    const activoInfo = mapActivos[idActivo];
    if(activoInfo) {
       const nombreEdificio = mapEdificios[activoInfo.idEdif] || "-";
       let f = planes[i][4];
       let color = 'gris'; let fechaStr = "-"; let dias = 9999; let fechaISO = "";

       if (f instanceof Date) {
         f.setHours(0,0,0,0);
         dias = Math.ceil((f.getTime() - hoy.getTime()) / (86400000));
         if (dias < 0) color = 'rojo'; else if (dias <= 30) color = 'amarillo'; else color = 'verde';
         fechaStr = Utilities.formatDate(f, Session.getScriptTimeZone(), "dd/MM/yyyy");
         fechaISO = Utilities.formatDate(f, Session.getScriptTimeZone(), "yyyy-MM-dd");
       }
       result.push({ id: planes[i][0], activo: activoInfo.nombre, edificio: nombreEdificio, tipo: planes[i][2], fecha: fechaStr, fechaISO: fechaISO, color: color, dias: dias });
    }
  }
  return result.sort((a,b) => a.dias - b.dias);
}

// CREAR REVISIÓN
function crearRevision(d) {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const sheet = ss.getSheetByName('PLAN_MANTENIMIENTO');
  
  try {
    let fechaActual = textoAFecha(d.fechaProx);
    if (!fechaActual) return { success: false, error: "Fecha inválida" };

    var esRepetitiva = (String(d.esRecursiva) === "true"); 
    var frecuencia = parseInt(d.diasFreq) || 0;
    var fechaLimite = d.fechaFin ? textoAFecha(d.fechaFin) : null;

    // 1. Crear Primera Revisión
    sheet.appendRow([Utilities.getUuid(), d.idActivo, d.tipo, "", new Date(fechaActual), frecuencia, "ACTIVO"]);

    // 2. Crear Repeticiones
    if (esRepetitiva && frecuencia > 0 && fechaLimite && fechaLimite > fechaActual) {
      let contador = 0;
      while (contador < 50) { 
        fechaActual.setDate(fechaActual.getDate() + frecuencia);
        if (fechaActual > fechaLimite) break;
        sheet.appendRow([Utilities.getUuid(), d.idActivo, d.tipo, "", new Date(fechaActual), frecuencia, "ACTIVO"]);
        contador++;
      }
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// UPDATE REVISIÓN
function updateRevision(d) { 
    const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); 
    const sheet = ss.getSheetByName('PLAN_MANTENIMIENTO'); 
    const data = sheet.getDataRange().getValues(); 
    
    let nuevaFecha = textoAFecha(d.fechaProx);
    if (!nuevaFecha) return { success: false, error: "Fecha inválida al editar" };

    for(let i=1; i<data.length; i++){ 
        if(String(data[i][0]) === String(d.idPlan)) { 
            sheet.getRange(i+1, 3).setValue(d.tipo); 
            sheet.getRange(i+1, 5).setValue(nuevaFecha); 
            break;
        } 
    } 
    
    // Crear nuevas repeticiones (Lógica igual a crearRevision)
    var esRepetitiva = (String(d.esRecursiva) === "true"); 
    var frecuencia = parseInt(d.diasFreq) || 0;
    var fechaLimite = d.fechaFin ? textoAFecha(d.fechaFin) : null;
    
    if (esRepetitiva && frecuencia > 0 && fechaLimite && fechaLimite > nuevaFecha) {
         let fechaIter = new Date(nuevaFecha.getTime());
         let contador = 0;
         while (contador < 50) {
            fechaIter.setDate(fechaIter.getDate() + frecuencia);
            if (fechaIter > fechaLimite) break;
            sheet.appendRow([Utilities.getUuid(), d.idActivo, d.tipo, "", new Date(fechaIter), frecuencia, "ACTIVO"]);
            contador++;
         }
    }
    
    return { success: true }; 
}

function eliminarRevision(idPlan) { const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const sheet = ss.getSheetByName('PLAN_MANTENIMIENTO'); const data = sheet.getDataRange().getValues(); for(let i=1; i<data.length; i++){ if(String(data[i][0]) === String(idPlan)) { sheet.deleteRow(i+1); return { success: true }; } } return { success: false, error: "Plan no encontrado" }; }

// ==========================================
// 5. CONTRATOS (VISTA LOCAL Y GLOBAL)
// ==========================================
function obtenerContratos(idEntidad) {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const sheet = ss.getSheetByName('CONTRATOS'); if (!sheet || sheet.getLastRow() < 2) return []; 
  const data = sheet.getRange(1, 1, sheet.getLastRow(), 8).getValues(); 
  const contratos = []; 
  const hoy = new Date(); 
  
  for(let i=1; i<data.length; i++) { 
    if(String(data[i][2]) === String(idEntidad)) { 
      const fFin = data[i][6] instanceof Date ? data[i][6] : null; 
      const fIni = data[i][5] instanceof Date ? data[i][5] : null; 
      let estadoDB = (data[i].length > 7) ? data[i][7] : 'ACTIVO'; 
      let estadoCalc = 'VIGENTE'; let color = 'verde'; 
      
      if (estadoDB === 'INACTIVO') { estadoCalc = 'INACTIVO'; color = 'gris'; } 
      else if (fFin) { 
        const diff = Math.ceil((fFin.getTime() - hoy.getTime()) / (86400000)); 
        if (diff < 0) { estadoCalc = 'CADUCADO'; color = 'rojo'; } 
        else if (diff <= 30) { estadoCalc = 'PRÓXIMO'; color = 'amarillo'; } 
      } else { estadoCalc = 'SIN FECHA'; color = 'gris'; } 
      
      contratos.push({ 
        id: data[i][0], 
        proveedor: data[i][3], 
        ref: data[i][4], 
        inicio: fIni?Utilities.formatDate(fIni,Session.getScriptTimeZone(),"yyyy-MM-dd"):"", 
        fin: fFin?Utilities.formatDate(fFin,Session.getScriptTimeZone(),"yyyy-MM-dd"):"", 
        estado: estadoCalc, color: color, estadoDB: estadoDB 
      }); 
    } 
  } 
  return contratos.sort((a, b) => a.fin.localeCompare(b.fin));
}

// *** NUEVA FUNCIÓN: OBTENER CONTRATOS GLOBAL ***
function obtenerContratosGlobal() {
  const contratos = getSheetData('CONTRATOS');
  const activos = getSheetData('ACTIVOS');
  const edificios = getSheetData('EDIFICIOS');
  
  const mapActivos = {}; activos.slice(1).forEach(r => mapActivos[r[0]] = { nombre: r[3], idEdif: r[1] });
  const mapEdificios = {}; edificios.slice(1).forEach(r => mapEdificios[r[0]] = r[2]);
  
  const result = [];
  const hoy = new Date();
  
  for(let i=1; i<contratos.length; i++) {
    const r = contratos[i];
    const idEntidad = r[2];
    const tipoEntidad = r[1];
    
    let nombreEntidad = "N/A";
    
    if (tipoEntidad === 'ACTIVO' && mapActivos[idEntidad]) {
      const info = mapActivos[idEntidad];
      nombreEntidad = info.nombre + " (" + mapEdificios[info.idEdif] + ")";
    } else if (tipoEntidad === 'EDIFICIO' && mapEdificios[idEntidad]) {
      nombreEntidad = mapEdificios[idEntidad];
    }
    
    const fFin = (r[6] instanceof Date) ? r[6] : null;
    let estadoDB = r[7] || 'ACTIVO';
    let estadoCalc = 'VIGENTE'; let color = 'verde';
    
    if (estadoDB === 'INACTIVO') { estadoCalc = 'INACTIVO'; color = 'gris'; }
    else if (fFin) {
       const diff = Math.ceil((fFin.getTime() - hoy.getTime()) / 86400000);
       if (diff < 0) { estadoCalc = 'CADUCADO'; color = 'rojo'; }
       else if (diff <= 90) { estadoCalc = 'PRÓXIMO'; color = 'amarillo'; } // Usamos 90 días para el filtro global
       
    } else { estadoCalc = 'SIN FECHA'; color = 'gris'; }
    
    result.push({
      id: r[0], nombreEntidad: nombreEntidad, proveedor: r[3], ref: r[4],
      inicio: r[5] ? Utilities.formatDate(r[5], Session.getScriptTimeZone(), "yyyy-MM-dd") : "-",
      fin: fFin ? Utilities.formatDate(fFin, Session.getScriptTimeZone(), "yyyy-MM-dd") : "-",
      estado: estadoCalc, color: color, estadoDB: estadoDB
    });
  }
  
  return result.sort((a, b) => {
      if (a.fin === "-") return 1;
      if (b.fin === "-") return -1;
      return a.fin.localeCompare(b.fin);
  });
}
// FIN NUEVA FUNCIÓN


function crearContrato(d) { const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); ss.getSheetByName('CONTRATOS').appendRow([Utilities.getUuid(), d.tipoEntidad, d.idEntidad, d.proveedor, d.ref, textoAFecha(d.fechaIni), textoAFecha(d.fechaFin), d.estado]); return { success: true }; }
function updateContrato(datos) { const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const sheet = ss.getSheetByName('CONTRATOS'); const data = sheet.getDataRange().getValues(); for(let i=1; i<data.length; i++){ if(String(data[i][0]) === String(datos.id)) { sheet.getRange(i+1, 4).setValue(datos.proveedor); sheet.getRange(i+1, 5).setValue(datos.ref); sheet.getRange(i+1, 6).setValue(textoAFecha(datos.fechaIni)); sheet.getRange(i+1, 7).setValue(textoAFecha(datos.fechaFin)); sheet.getRange(i+1, 8).setValue(datos.estado); return { success: true }; } } return { success: false, error: "Contrato no encontrado" }; }
function eliminarContrato(id) { const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const sheet = ss.getSheetByName('CONTRATOS'); const data = sheet.getDataRange().getValues(); for(let i=1; i<data.length; i++){ if(String(data[i][0]) === String(id)) { sheet.deleteRow(i+1); return { success: true }; } } return { success: false, error: "Contrato no encontrado" }; }

// ==========================================
// 6. DASHBOARD & OTROS
// ==========================================
function getDatosDashboard() { const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const hoy = new Date(); const cAct = (ss.getSheetByName('ACTIVOS').getLastRow() - 1) || 0; const cEdif = (ss.getSheetByName('EDIFICIOS').getLastRow() - 1) || 0; const dataMant = getSheetData('PLAN_MANTENIMIENTO'); let revPend = 0, revVenc = 0, revOk = 0; for(let i=1; i<dataMant.length; i++) { const f = dataMant[i][4]; if(f instanceof Date) { const diff = Math.ceil((f - hoy) / 86400000); if(diff < 0) revVenc++; else if(diff <= 30) revPend++; else revOk++; } } const dataCont = getSheetData('CONTRATOS'); let contCad = 0; for(let i=1; i<dataCont.length; i++) { const f = dataCont[i][6]; if(f instanceof Date && f < hoy) contCad++; } return { activos: cAct, edificios: cEdif, pendientes: revPend, vencidas: revVenc, ok: revOk, contratosCaducados: contCad }; }
function crearCampus(d) { const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const fId = crearCarpeta(d.nombre, getRootFolderId()); ss.getSheetByName('CAMPUS').appendRow([Utilities.getUuid(), d.nombre, d.provincia, d.direccion, fId]); return {success:true}; }
function crearEdificio(d) { const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const cData = getSheetData('CAMPUS'); let pId; for(let i=1; i<cData.length; i++) if(String(cData[i][0])==String(d.idCampus)) pId=cData[i][4]; const fId = crearCarpeta(d.nombre, pId); const aId = crearCarpeta("Activos", fId); ss.getSheetByName('EDIFICIOS').appendRow([Utilities.getUuid(), d.idCampus, d.nombre, d.contacto, fId, aId]); return {success:true}; }
function crearActivo(d) { const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const eData = getSheetData('EDIFICIOS'); let pId; for(let i=1; i<eData.length; i++) if(String(eData[i][0])==String(d.idEdificio)) pId=eData[i][5]; const fId = crearCarpeta(d.nombre, pId); const id = Utilities.getUuid(); const cats = getSheetData('CAT_INSTALACIONES'); let nombreTipo = d.tipo; for(let i=1; i<cats.length; i++) { if(String(cats[i][0]) === String(d.tipo)) { nombreTipo = cats[i][1]; break; } } ss.getSheetByName('ACTIVOS').appendRow([id, d.idEdificio, nombreTipo, d.nombre, d.marca, new Date(), fId]); return {success:true}; }
function getCatalogoInstalaciones() { return getSheetData('CAT_INSTALACIONES').slice(1).map(r=>({id:r[0], nombre:r[1], dias:r[3]})); }
function getTableData(tipo) { const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); if (tipo === 'CAMPUS') { const data = getSheetData('CAMPUS'); return data.slice(1).map(r => ({ id: r[0], nombre: r[1], provincia: r[2], direccion: r[3] })); } if (tipo === 'EDIFICIOS') { const data = getSheetData('EDIFICIOS'); const dataC = getSheetData('CAMPUS'); const mapCampus = {}; dataC.slice(1).forEach(r => mapCampus[r[0]] = r[1]); return data.slice(1).map(r => ({ id: r[0], campus: mapCampus[r[1]] || '-', nombre: r[2], contacto: r[3] })); } return []; }
