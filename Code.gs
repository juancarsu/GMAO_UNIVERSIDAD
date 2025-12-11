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
  var partes = String(txt).split('-');
  if (partes.length !== 3) return null;
  return new Date(parseInt(partes[0], 10), parseInt(partes[1], 10) - 1, parseInt(partes[2], 10), 12, 0, 0);
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
function obtenerPlanMantenimiento(idActivo) {
  const data = getSheetData('PLAN_MANTENIMIENTO'); 
  
  // --- NUEVO: PRECARGAR QUE REVISIONES TIENEN DOCS ---
  const docsData = getSheetData('DOCS_HISTORICO');
  const docsMap = {};
  for(let j=1; j<docsData.length; j++) {
    if(String(docsData[j][1]) === 'REVISION') {
      docsMap[String(docsData[j][2])] = true; // ID_Entidad
    }
  }
  // ----------------------------------------------------

  const planes = [];
  const hoy = new Date(); hoy.setHours(0,0,0,0);
  for(let i=1; i<data.length; i++) {
    if(String(data[i][1]) === String(idActivo)) {
      let f = data[i][4];
      let color = 'gris'; let fechaStr = "-"; let fechaISO = "";
      if (f instanceof Date) {
         f.setHours(0,0,0,0);
         const diff = Math.ceil((f.getTime() - hoy.getTime()) / (86400000));
         if (diff < 0) color = 'rojo'; else if (diff <= 30) color = 'amarillo'; else color = 'verde';
         fechaStr = Utilities.formatDate(f, Session.getScriptTimeZone(), "dd/MM/yyyy");
         fechaISO = Utilities.formatDate(f, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
      planes.push({ 
        id: data[i][0], 
        tipo: data[i][2], 
        fechaProxima: fechaStr, 
        fechaISO: fechaISO, 
        color: color,
        hasDocs: docsMap[String(data[i][0])] || false // <--- NUEVO
      });
    }
  }
  return planes.sort((a, b) => a.fechaISO.localeCompare(b.fechaISO));
}

function getGlobalMaintenance() {
  const planes = getSheetData('PLAN_MANTENIMIENTO');
  const activos = getSheetData('ACTIVOS');
  const edificios = getSheetData('EDIFICIOS');
  const campus = getSheetData('CAMPUS'); 

  // --- NUEVO: MAPA DE DOCS ---
  const docsData = getSheetData('DOCS_HISTORICO');
  const docsMap = {};
  for(let j=1; j<docsData.length; j++) {
    if(String(docsData[j][1]) === 'REVISION') {
      docsMap[String(docsData[j][2])] = true;
    }
  }
  // ---------------------------

  const mapCampus = {}; campus.slice(1).forEach(r => mapCampus[r[0]] = r[1]);
  const mapEdificios = {}; edificios.slice(1).forEach(r => mapEdificios[r[0]] = { nombre: r[2], idCampus: r[1] });
  const mapActivos = {}; 
  activos.slice(1).forEach(r => {
    const edificioInfo = mapEdificios[r[1]] || {};
    mapActivos[r[0]] = { nombre: r[3], idEdif: r[1], idCampus: edificioInfo.idCampus || null }; 
  });
  const result = [];
  const hoy = new Date(); hoy.setHours(0,0,0,0);
  for(let i=1; i<planes.length; i++) {
    const idActivo = planes[i][1];
    const activoInfo = mapActivos[idActivo];
    if(activoInfo) {
       const nombreEdificio = mapEdificios[activoInfo.idEdif] ? mapEdificios[activoInfo.idEdif].nombre : "-";
       let f = planes[i][4];
       let color = 'gris'; let fechaStr = "-"; let dias = 9999; let fechaISO = "";
       if (f instanceof Date) {
         f.setHours(0,0,0,0);
         dias = Math.ceil((f.getTime() - hoy.getTime()) / (86400000));
         if (dias < 0) color = 'rojo'; else if (dias <= 30) color = 'amarillo'; else color = 'verde';
         fechaStr = Utilities.formatDate(f, Session.getScriptTimeZone(), "dd/MM/yyyy");
         fechaISO = Utilities.formatDate(f, Session.getScriptTimeZone(), "yyyy-MM-dd");
       }
       result.push({ 
         id: planes[i][0], 
         idActivo: idActivo, 
         activo: activoInfo.nombre, 
         edificio: nombreEdificio, 
         tipo: planes[i][2], 
         fecha: fechaStr, 
         fechaISO: fechaISO, 
         color: color, 
         dias: dias, 
         edificioId: activoInfo.idEdif, 
         campusId: activoInfo.idCampus,
         hasDocs: docsMap[String(planes[i][0])] || false // <--- NUEVO
       });
    }
  }
  return result.sort((a,b) => a.dias - b.dias);
}

function crearRevision(d) {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const sheet = ss.getSheetByName('PLAN_MANTENIMIENTO');
  try {
    let fechaActual = textoAFecha(d.fechaProx);
    if (!fechaActual) return { success: false, error: "Fecha inválida" };
    var esRepetitiva = (String(d.esRecursiva) === "true"); 
    var frecuencia = parseInt(d.diasFreq) || 0;
    var fechaLimite = d.fechaFin ? textoAFecha(d.fechaFin) : null;
    sheet.appendRow([Utilities.getUuid(), d.idActivo, d.tipo, "", new Date(fechaActual), frecuencia, "ACTIVO"]);
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
  } catch (e) { return { success: false, error: e.toString() }; }
}

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
// 5. CONTRATOS
// ==========================================
function obtenerContratos(idEntidad) {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const sheet = ss.getSheetByName('CONTRATOS'); if (!sheet || sheet.getLastRow() < 2) return []; const data = sheet.getRange(1, 1, sheet.getLastRow(), 8).getValues(); const contratos = []; const hoy = new Date(); for(let i=1; i<data.length; i++) { if(String(data[i][2]) === String(idEntidad)) { const fFin = data[i][6] instanceof Date ? data[i][6] : null; const fIni = data[i][5] instanceof Date ? data[i][5] : null; let estadoDB = (data[i].length > 7) ? data[i][7] : 'ACTIVO'; let estadoCalc = 'VIGENTE'; let color = 'verde'; if (estadoDB === 'INACTIVO') { estadoCalc = 'INACTIVO'; color = 'gris'; } else if (fFin) { const diff = Math.ceil((fFin.getTime() - hoy.getTime()) / (86400000)); if (diff < 0) { estadoCalc = 'CADUCADO'; color = 'rojo'; } else if (diff <= 30) { estadoCalc = 'PRÓXIMO'; color = 'amarillo'; } } else { estadoCalc = 'SIN FECHA'; color = 'gris'; } contratos.push({ id: data[i][0], proveedor: data[i][3], ref: data[i][4], inicio: fIni?fechaATexto(fIni):"-", fin: fFin?fechaATexto(fFin):"-", estado: estadoCalc, color: color, estadoDB: estadoDB }); } } return contratos.sort((a, b) => a.fin.localeCompare(b.fin));
}

function obtenerContratosGlobal() {
  const contratos = getSheetData('CONTRATOS');
  const activos = getSheetData('ACTIVOS');
  const edificios = getSheetData('EDIFICIOS');
  const campus = getSheetData('CAMPUS'); 
  const mapCampus = {}; campus.slice(1).forEach(r => mapCampus[r[0]] = r[1]);
  const mapEdificios = {}; edificios.slice(1).forEach(r => mapEdificios[r[0]] = { nombre: r[2], idCampus: r[1] });
  const mapActivos = {}; 
  activos.slice(1).forEach(r => {
    const edificioInfo = mapEdificios[r[1]] || {};
    mapActivos[r[0]] = { nombre: r[3], idEdif: r[1], idCampus: edificioInfo.idCampus || null }; 
  });
  const result = [];
  const hoy = new Date();
  for(let i=1; i<contratos.length; i++) {
    const r = contratos[i];
    const idEntidad = r[2];
    const tipoEntidad = r[1];
    let nombreEntidad = "N/A"; let edificioId = null; let campusId = null;
    if (tipoEntidad === 'ACTIVO' && mapActivos[idEntidad]) {
      const info = mapActivos[idEntidad];
      nombreEntidad = info.nombre + " (" + (mapEdificios[info.idEdif] ? mapEdificios[info.idEdif].nombre : 'Sin Edif.') + ")";
      edificioId = info.idEdif; campusId = info.idCampus;
    } else if (tipoEntidad === 'EDIFICIO' && mapEdificios[idEntidad]) {
      const info = mapEdificios[idEntidad];
      nombreEntidad = info.nombre;
      edificioId = idEntidad; campusId = info.idCampus;
    }
    const fFin = (r[6] instanceof Date) ? r[6] : null;
    let estadoDB = r[7] || 'ACTIVO';
    let estadoCalc = 'VIGENTE'; let color = 'verde';
    if (estadoDB === 'INACTIVO') { estadoCalc = 'INACTIVO'; color = 'gris'; }
    else if (fFin) {
       const diff = Math.ceil((fFin.getTime() - hoy.getTime()) / 86400000);
       if (diff < 0) { estadoCalc = 'CADUCADO'; color = 'rojo'; }
       else if (diff <= 90) { estadoCalc = 'PRÓXIMO'; color = 'amarillo'; }
    } else { estadoCalc = 'SIN FECHA'; color = 'gris'; }
    result.push({ id: r[0], nombreEntidad: nombreEntidad, proveedor: r[3], ref: r[4], inicio: r[5] ? fechaATexto(r[5]) : "-", fin: fFin ? fechaATexto(fFin) : "-", estado: estadoCalc, color: color, estadoDB: estadoDB, edificioId: edificioId, campusId: campusId });
  }
  return result.sort((a, b) => { if (a.fin === "-") return 1; if (b.fin === "-") return -1; return a.fin.localeCompare(b.fin); });
}
function crearContrato(d) { const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); ss.getSheetByName('CONTRATOS').appendRow([Utilities.getUuid(), d.tipoEntidad, d.idEntidad, d.proveedor, d.ref, textoAFecha(d.fechaIni), textoAFecha(d.fechaFin), d.estado]); return { success: true }; }
function updateContrato(datos) { const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const sheet = ss.getSheetByName('CONTRATOS'); const data = sheet.getDataRange().getValues(); for(let i=1; i<data.length; i++){ if(String(data[i][0]) === String(datos.id)) { sheet.getRange(i+1, 4).setValue(datos.proveedor); sheet.getRange(i+1, 5).setValue(datos.ref); sheet.getRange(i+1, 6).setValue(textoAFecha(datos.fechaIni)); sheet.getRange(i+1, 7).setValue(textoAFecha(datos.fechaFin)); sheet.getRange(i+1, 8).setValue(datos.estado); return { success: true }; } } return { success: false, error: "Contrato no encontrado" }; }
function eliminarContrato(id) { const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const sheet = ss.getSheetByName('CONTRATOS'); const data = sheet.getDataRange().getValues(); for(let i=1; i<data.length; i++){ if(String(data[i][0]) === String(id)) { sheet.deleteRow(i+1); return { success: true }; } } return { success: false, error: "Contrato no encontrado" }; }

// ==========================================
// 6. DASHBOARD & OTROS
// ==========================================
function getDatosDashboard() { 
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); 
  const hoy = new Date(); hoy.setHours(0,0,0,0);
  
  const dataMant = getSheetData('PLAN_MANTENIMIENTO'); 
  let revPend = 0, revVenc = 0, revOk = 0; 
  
  // LOGICA DE GRÁFICO (Próximos 6 meses)
  const mesesNombres = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];
  const countsMap = {};
  const chartLabels = [];
  const chartData = [];
  
  // Inicializar mapa para los próximos 6 meses
  let dIter = new Date(hoy.getFullYear(), hoy.getMonth(), 1); // Primer día del mes actual
  for (let k = 0; k < 6; k++) {
    let key = mesesNombres[dIter.getMonth()] + " " + dIter.getFullYear();
    countsMap[key] = 0;
    chartLabels.push(key);
    dIter.setMonth(dIter.getMonth() + 1);
  }

  for(let i=1; i<dataMant.length; i++) { 
      const f = dataMant[i][4]; 
      if(f instanceof Date) { 
          f.setHours(0,0,0,0);
          const diff = Math.ceil((f.getTime() - hoy.getTime()) / 86400000); 
          if(diff < 0) revVenc++; 
          else if(diff <= 30) revPend++; 
          else revOk++; 
          
          // Llenar gráfico
          let key = mesesNombres[f.getMonth()] + " " + f.getFullYear();
          if (countsMap.hasOwnProperty(key)) {
             countsMap[key]++;
          }
      } 
  } 
  
  // Pasar datos del mapa al array en orden
  chartLabels.forEach(lbl => {
      chartData.push(countsMap[lbl]);
  });
  
  const dataCont = getSheetData('CONTRATOS'); 
  let contCad = 0; 
  for(let i=1; i<dataCont.length; i++) { 
      const f = dataCont[i][6]; 
      if(f instanceof Date && f < hoy) contCad++; 
  }
  
  const cAct = (ss.getSheetByName('ACTIVOS').getLastRow() - 1) || 0;
  const cEdif = (ss.getSheetByName('EDIFICIOS').getLastRow() - 1) || 0;
  
  return { 
      activos: cAct, 
      edificios: cEdif, 
      pendientes: revPend, 
      vencidas: revVenc, 
      ok: revOk, 
      contratosCaducados: contCad,
      chartLabels: chartLabels, 
      chartData: chartData      
  }; 
}

function crearCampus(d) { const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const fId = crearCarpeta(d.nombre, getRootFolderId()); ss.getSheetByName('CAMPUS').appendRow([Utilities.getUuid(), d.nombre, d.provincia, d.direccion, fId]); return {success:true}; }
function crearEdificio(d) { const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const cData = getSheetData('CAMPUS'); let pId; for(let i=1; i<cData.length; i++) if(String(cData[i][0])==String(d.idCampus)) pId=cData[i][4]; const fId = crearCarpeta(d.nombre, pId); const aId = crearCarpeta("Activos", fId); ss.getSheetByName('EDIFICIOS').appendRow([Utilities.getUuid(), d.idCampus, d.nombre, d.contacto, fId, aId]); return {success:true}; }
function crearActivo(d) { const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const eData = getSheetData('EDIFICIOS'); let pId; for(let i=1; i<eData.length; i++) if(String(eData[i][0])==String(d.idEdificio)) pId=eData[i][5]; const fId = crearCarpeta(d.nombre, pId); const id = Utilities.getUuid(); const cats = getSheetData('CAT_INSTALACIONES'); let nombreTipo = d.tipo; for(let i=1; i<cats.length; i++) { if(String(cats[i][0]) === String(d.tipo)) { nombreTipo = cats[i][1]; break; } } ss.getSheetByName('ACTIVOS').appendRow([id, d.idEdificio, nombreTipo, d.nombre, d.marca, new Date(), fId]); return {success:true}; }
function getCatalogoInstalaciones() { return getSheetData('CAT_INSTALACIONES').slice(1).map(r=>({id:r[0], nombre:r[1], dias:r[3]})); }
function getTableData(tipo) { const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); if (tipo === 'CAMPUS') { const data = getSheetData('CAMPUS'); return data.slice(1).map(r => ({ id: r[0], nombre: r[1], provincia: r[2], direccion: r[3] })); } if (tipo === 'EDIFICIOS') { const data = getSheetData('EDIFICIOS'); const dataC = getSheetData('CAMPUS'); const mapCampus = {}; dataC.slice(1).forEach(r => mapCampus[r[0]] = r[1]); return data.slice(1).map(r => ({ id: r[0], campus: mapCampus[r[1]] || '-', nombre: r[2], contacto: r[3] })); } return []; }

// BUSCADOR GLOBAL
function buscarGlobal(texto) {
  if (!texto || texto.length < 3) return []; // Mínimo 3 caracteres
  texto = texto.toLowerCase();
  
  const resultados = [];
  
  // 1. Buscar en ACTIVOS
  const activos = getSheetData('ACTIVOS');
  for(let i=1; i<activos.length; i++) {
    const r = activos[i];
    if (String(r[3]).toLowerCase().includes(texto) || 
        String(r[2]).toLowerCase().includes(texto) || 
        String(r[4]).toLowerCase().includes(texto)) {
      resultados.push({
        id: r[0],
        tipo: 'ACTIVO',
        texto: r[3], 
        subtexto: r[2] + (r[4] ? " - " + r[4] : "") 
      });
    }
  }
  
  // 2. Buscar en EDIFICIOS
  const edificios = getSheetData('EDIFICIOS');
  for(let i=1; i<edificios.length; i++) {
    const r = edificios[i];
    if (String(r[2]).toLowerCase().includes(texto)) {
      resultados.push({
        id: r[0],
        tipo: 'EDIFICIO',
        texto: r[2], 
        subtexto: 'Edificio'
      });
    }
  }
  
  return resultados.slice(0, 10); // Limitar a 10 resultados
}

// ==========================================
// 7. GESTIÓN DOCUMENTAL (ACTIVOS Y REVISIONES)
// ==========================================

function obtenerDocs(idEntidad, tipoEntidad) {
  // Si no se pasa tipo, asumimos ACTIVO por defecto
  const tipo = tipoEntidad || 'ACTIVO';
  const data = getSheetData('DOCS_HISTORICO');
  const docs = [];
  
  // Indices: 0=ID_Doc, 1=Tipo, 2=ID_Entidad, 3=Nombre, 4=URL, 5=Version, 6=Fecha
  for(let i=1; i<data.length; i++) {
    if(String(data[i][2]) === String(idEntidad) && String(data[i][1]) === tipo) {
       const f = data[i][6];
       const fechaStr = (f instanceof Date) ? Utilities.formatDate(f, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : "-";
       docs.push({
         id: data[i][0],
         nombre: data[i][3],
         url: data[i][4],
         version: data[i][5],
         fecha: fechaStr
       });
    }
  }
  return docs.sort((a,b) => b.version - a.version); // Ordenar por versión
}

function subirArchivo(dataBase64, nombreArchivo, mimeType, idEntidad, tipoEntidad) {
  try {
    const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
    let carpetaId = null;
    let activoId = idEntidad; // Por defecto

    // 1. DETERMINAR CARPETA DE DESTINO
    if (tipoEntidad === 'ACTIVO') {
      // Buscar carpeta del activo directamente
      const activos = getSheetData('ACTIVOS');
      for(let i=1; i<activos.length; i++) {
        if(String(activos[i][0]) === String(idEntidad)) {
          carpetaId = activos[i][6]; // Columna G es ID_Carpeta
          break;
        }
      }
    } else if (tipoEntidad === 'REVISION') {
      // Si es revisión, primero buscamos a qué activo pertenece esa revisión
      const planes = getSheetData('PLAN_MANTENIMIENTO');
      for(let i=1; i<planes.length; i++) {
        if(String(planes[i][0]) === String(idEntidad)) {
          activoId = planes[i][1]; // Obtenemos el ID del Activo padre
          break;
        }
      }
      // Ahora buscamos la carpeta de ese activo
      if(activoId) {
        const activos = getSheetData('ACTIVOS');
        for(let i=1; i<activos.length; i++) {
          if(String(activos[i][0]) === String(activoId)) {
            carpetaId = activos[i][6];
            break;
          }
        }
      }
    }

    if (!carpetaId) throw new Error("No se encontró carpeta de destino para este elemento.");

    // 2. SUBIR A DRIVE
    const blob = Utilities.newBlob(Utilities.base64Decode(dataBase64), mimeType, nombreArchivo);
    const folder = DriveApp.getFolderById(carpetaId);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // 3. REGISTRAR EN CSV
    const version = 1; 
    const usuario = Session.getActiveUser().getEmail();
    const fecha = new Date();
    
    ss.getSheetByName('DOCS_HISTORICO').appendRow([
      Utilities.getUuid(),
      tipoEntidad,
      idEntidad,
      nombreArchivo,
      file.getUrl(),
      version,
      fecha,
      usuario,
      file.getId()
    ]);

    return { success: true };

  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function eliminarDocumento(idDoc) {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const sheet = ss.getSheetByName('DOCS_HISTORICO');
  const data = sheet.getDataRange().getValues();
  
  for(let i=1; i<data.length; i++){
    if(String(data[i][0]) === String(idDoc)) {
      // Opcional: Borrar de Drive también usando DriveApp.getFileById(data[i][8]).setTrashed(true);
      sheet.deleteRow(i+1);
      return { success: true };
    }
  }
  return { success: false, error: "Documento no encontrado" };
}

// ==========================================
// 8. GESTIÓN DE EDIFICIOS
// ==========================================

function getBuildingInfo(idEdificio) {
  const data = getSheetData('EDIFICIOS');
  const campus = getSheetData('CAMPUS');
  const mapCampus = {};
  campus.slice(1).forEach(r => mapCampus[r[0]] = r[1]);

  for(let i=1; i<data.length; i++){
    if(String(data[i][0]) === String(idEdificio)){
      return { 
        id: data[i][0], 
        campus: mapCampus[data[i][1]] || "Desconocido",
        nombre: data[i][2], 
        contacto: data[i][3] 
      };
    }
  }
  throw new Error("Edificio no encontrado.");
}

function getObrasPorEdificio(idEdificio) {
  const data = getSheetData('OBRAS');
  const obras = [];
  // Indices: 0=ID, 1=ID_Edif, 2=Nombre, 3=Desc, 4=F.Ini, 5=F.Fin, 6=Estado
  
  // MAPA DE DOCS para saber si tienen adjuntos (Clip)
  const docsData = getSheetData('DOCS_HISTORICO');
  const docsMap = {};
  for(let j=1; j<docsData.length; j++) {
    if(String(docsData[j][1]) === 'OBRA') docsMap[String(docsData[j][2])] = true;
  }

  for(let i=1; i<data.length; i++) {
    if(String(data[i][1]) === String(idEdificio)) {
      obras.push({
        id: data[i][0],
        nombre: data[i][2],
        descripcion: data[i][3],
        fecha: data[i][4] ? fechaATexto(data[i][4]) : "-",
        estado: data[i][6],
        hasDocs: docsMap[String(data[i][0])] || false
      });
    }
  }
  return obras;
}

function crearObra(d) {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  
  // 1. Buscar carpeta del edificio padre para crear la subcarpeta de la obra
  const edifs = getSheetData('EDIFICIOS');
  let idCarpetaEdificio = null;
  for(let i=1; i<edifs.length; i++){
    if(String(edifs[i][0]) === String(d.idEdificio)) {
      idCarpetaEdificio = edifs[i][4]; // Columna E es Folder ID del edificio
      break;
    }
  }
  
  if(!idCarpetaEdificio) return { success: false, error: "Carpeta de edificio no encontrada" };

  // 2. Crear carpeta para la obra (o una general de obras si prefieres, aquí creo una por obra)
  const nombreCarpeta = "OBRA - " + d.nombre;
  const idCarpetaObra = crearCarpeta(nombreCarpeta, idCarpetaEdificio);

  // 3. Guardar en Sheet
  ss.getSheetByName('OBRAS').appendRow([
    Utilities.getUuid(),
    d.idEdificio,
    d.nombre,
    d.descripcion,
    textoAFecha(d.fechaInicio),
    null, // Fecha fin opcional
    "EN CURSO",
    idCarpetaObra
  ]);
  return { success: true };
}

function eliminarObra(id) {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const sheet = ss.getSheetByName('OBRAS');
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++){
    if(String(data[i][0]) === String(id)) {
      // Opcional: Marcar carpeta como papelera
      sheet.deleteRow(i+1);
      return { success: true };
    }
  }
  return { success: false, error: "Obra no encontrada" };
}

// === ACTUALIZACIÓN DE LA FUNCIÓN subirArchivo (REEMPLAZA LA ANTERIOR) ===
function subirArchivo(dataBase64, nombreArchivo, mimeType, idEntidad, tipoEntidad) {
  try {
    const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
    let carpetaId = null;

    // LÓGICA DE CARPETAS SEGÚN TIPO
    if (tipoEntidad === 'ACTIVO') {
      const rows = getSheetData('ACTIVOS');
      for(let i=1; i<rows.length; i++) if(String(rows[i][0]) === String(idEntidad)) { carpetaId = rows[i][6]; break; }
    } 
    else if (tipoEntidad === 'EDIFICIO') {
      const rows = getSheetData('EDIFICIOS');
      for(let i=1; i<rows.length; i++) if(String(rows[i][0]) === String(idEntidad)) { carpetaId = rows[i][4]; break; } // Col 4 = FolderID Edificio
    }
    else if (tipoEntidad === 'OBRA') {
      const rows = getSheetData('OBRAS');
      for(let i=1; i<rows.length; i++) if(String(rows[i][0]) === String(idEntidad)) { carpetaId = rows[i][7]; break; } // Col 7 = FolderID Obra
    }
    else if (tipoEntidad === 'REVISION') {
      // Lógica existente para revisiones... (buscar activo padre)
      const planes = getSheetData('PLAN_MANTENIMIENTO');
      let activoId = null;
      for(let i=1; i<planes.length; i++) if(String(planes[i][0]) === String(idEntidad)) { activoId = planes[i][1]; break; }
      if(activoId) {
         const rows = getSheetData('ACTIVOS');
         for(let i=1; i<rows.length; i++) if(String(rows[i][0]) === String(activoId)) { carpetaId = rows[i][6]; break; }
      }
    }

    if (!carpetaId) throw new Error("No se encontró carpeta de destino.");

    const blob = Utilities.newBlob(Utilities.base64Decode(dataBase64), mimeType, nombreArchivo);
    const folder = DriveApp.getFolderById(carpetaId);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    ss.getSheetByName('DOCS_HISTORICO').appendRow([
      Utilities.getUuid(), tipoEntidad, idEntidad, nombreArchivo, file.getUrl(), 1, new Date(), Session.getActiveUser().getEmail(), file.getId()
    ]);

    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}
