// GMAO
// Universidad de Navarra
// Versión 1.0
// Autor: Juan Carlos Suárez
//
// Licencia: Creative Commons Reconocimiento (CC BY) - creativecommons.org
// Puedes usar, copiar, modificar y distribuir este código (sin fines comerciales),
// siempre que cites a Juan Carlos Suárez como autor original.

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
// 2. SEGURIDAD Y ROLES
// ==========================================
function getMyRole() {
  const email = Session.getActiveUser().getEmail();
  const data = getSheetData('USUARIOS');
  for(let i=1; i<data.length; i++) {
    if(String(data[i][2]).trim().toLowerCase() === email.toLowerCase()) {
      return { email: email, nombre: data[i][1], rol: data[i][3] }; 
    }
  }
  return { email: email, nombre: "Invitado", rol: 'CONSULTA' };
}

function verificarPermiso(accionesPermitidas) {
  const usuario = getMyRole();
  const rol = usuario.rol;
  if (rol === 'ADMIN') return true;
  if (rol === 'CONSULTA') throw new Error("Acceso denegado: Permisos de solo lectura.");
  if (rol === 'TECNICO') {
    if (accionesPermitidas.includes('ADMIN_ONLY')) throw new Error("Acceso denegado: Requiere ser Administrador.");
    if (accionesPermitidas.includes('DELETE')) throw new Error("Acceso denegado: Los técnicos no pueden eliminar registros.");
    return true; 
  }
  throw new Error("Rol desconocido.");
}

// ==========================================
// 3. HELPERS BASE
// ==========================================
function getSheetData(sheetName) {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet && sheetName === 'USUARIOS') {
     sheet = ss.insertSheet('USUARIOS');
     sheet.appendRow(['ID', 'NOMBRE', 'EMAIL', 'ROL', 'RECIBIR_AVISOS']);
  }
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

function fechaATexto(dateObj) {
  if (!dateObj || !(dateObj instanceof Date)) return "-";
  return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "dd/MM/yyyy");
}

function textoAFecha(txt) {
  if (!txt) return null;
  var partes = String(txt).split('-');
  if (partes.length !== 3) return null;
  return new Date(parseInt(partes[0], 10), parseInt(partes[1], 10) - 1, parseInt(partes[2], 10), 12, 0, 0);
}

// ==========================================
// 4. API VISTAS
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
  verificarPermiso(['WRITE']);
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
// 5. EDIFICIOS Y OBRAS
// ==========================================
function getBuildingInfo(idEdificio) {
  const data = getSheetData('EDIFICIOS');
  const campus = getSheetData('CAMPUS');
  const mapCampus = {};
  campus.slice(1).forEach(r => mapCampus[r[0]] = r[1]);
  for(let i=1; i<data.length; i++){
    if(String(data[i][0]) === String(idEdificio)){
      return { id: data[i][0], campus: mapCampus[data[i][1]] || "Desconocido", nombre: data[i][2], contacto: data[i][3] };
    }
  }
  throw new Error("Edificio no encontrado.");
}

function getObrasPorEdificio(idEdificio) {
  const data = getSheetData('OBRAS');
  const obras = [];
  const docsData = getSheetData('DOCS_HISTORICO');
  const docsMap = {};
  for(let j=1; j<docsData.length; j++) { if(String(docsData[j][1]) === 'OBRA') docsMap[String(docsData[j][2])] = true; }
  for(let i=1; i<data.length; i++) {
    if(String(data[i][1]) === String(idEdificio)) {
      obras.push({ id: data[i][0], nombre: data[i][2], descripcion: data[i][3], fecha: data[i][4] ? fechaATexto(data[i][4]) : "-", estado: data[i][6], hasDocs: docsMap[String(data[i][0])] || false });
    }
  }
  return obras;
}

function crearObra(d) {
  verificarPermiso(['WRITE']);
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const edifs = getSheetData('EDIFICIOS');
  let idCarpetaEdificio = null;
  for(let i=1; i<edifs.length; i++){ if(String(edifs[i][0]) === String(d.idEdificio)) { idCarpetaEdificio = edifs[i][4]; break; } }
  if(!idCarpetaEdificio) return { success: false, error: "Carpeta de edificio no encontrada" };
  const nombreCarpeta = "OBRA - " + d.nombre;
  const idCarpetaObra = crearCarpeta(nombreCarpeta, idCarpetaEdificio);
  const newId = Utilities.getUuid();
  ss.getSheetByName('OBRAS').appendRow([newId, d.idEdificio, d.nombre, d.descripcion, textoAFecha(d.fechaInicio), null, "EN CURSO", idCarpetaObra]);
  return { success: true, newId: newId };
}

function finalizarObra(idObra, fechaFin) {
  verificarPermiso(['WRITE']);
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const sheet = ss.getSheetByName('OBRAS');
  const data = sheet.getDataRange().getValues();
  const fechaObj = textoAFecha(fechaFin);
  for(let i=1; i<data.length; i++){
    if(String(data[i][0]) === String(idObra)) {
      sheet.getRange(i+1, 6).setValue(fechaObj); sheet.getRange(i+1, 7).setValue("FINALIZADA"); return { success: true };
    }
  }
  return { success: false, error: "Obra no encontrada" };
}

function eliminarObra(id) {
  verificarPermiso(['DELETE']);
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const sheet = ss.getSheetByName('OBRAS');
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++){
    if(String(data[i][0]) === String(id)) { sheet.deleteRow(i+1); return { success: true }; }
  }
  return { success: false, error: "Obra no encontrada" };
}

// ==========================================
// 6. GESTIÓN DOCUMENTAL
// ==========================================
function obtenerDocs(idEntidad, tipoEntidad) {
  const tipo = tipoEntidad || 'ACTIVO';
  const data = getSheetData('DOCS_HISTORICO');
  const docs = [];
  for(let i=1; i<data.length; i++) {
    if(String(data[i][2]) === String(idEntidad) && String(data[i][1]) === tipo) {
       const f = data[i][6];
       const fechaStr = (f instanceof Date) ? Utilities.formatDate(f, Session.getScriptTimeZone(), "dd/MM/yyyy") : "-";
       docs.push({ id: data[i][0], nombre: data[i][3], url: data[i][4], version: data[i][5], fecha: fechaStr });
    }
  }
  return docs.sort((a,b) => b.version - a.version);
}

function subirArchivo(dataBase64, nombreArchivo, mimeType, idEntidad, tipoEntidad) {
  if (tipoEntidad !== 'INCIDENCIA') verificarPermiso(['WRITE']); 
  try {
    const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
    let carpetaId = null;
    if (tipoEntidad === 'ACTIVO') {
      const rows = getSheetData('ACTIVOS'); for(let i=1; i<rows.length; i++) if(String(rows[i][0]) === String(idEntidad)) { carpetaId = rows[i][6]; break; }
    } else if (tipoEntidad === 'EDIFICIO') {
      const rows = getSheetData('EDIFICIOS'); for(let i=1; i<rows.length; i++) if(String(rows[i][0]) === String(idEntidad)) { carpetaId = rows[i][4]; break; }
    } else if (tipoEntidad === 'OBRA') {
      const rows = getSheetData('OBRAS'); for(let i=1; i<rows.length; i++) if(String(rows[i][0]) === String(idEntidad)) { carpetaId = rows[i][7]; break; }
    } else if (tipoEntidad === 'REVISION') {
      const planes = getSheetData('PLAN_MANTENIMIENTO'); let activoId = null;
      for(let i=1; i<planes.length; i++) if(String(planes[i][0]) === String(idEntidad)) { activoId = planes[i][1]; break; }
      if(activoId) { const rows = getSheetData('ACTIVOS'); for(let i=1; i<rows.length; i++) if(String(rows[i][0]) === String(activoId)) { carpetaId = rows[i][6]; break; } }
    } else { carpetaId = getRootFolderId(); }
    if (!carpetaId) carpetaId = getRootFolderId();
    const blob = Utilities.newBlob(Utilities.base64Decode(dataBase64), mimeType, nombreArchivo);
    const folder = DriveApp.getFolderById(carpetaId);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    if(tipoEntidad !== 'INCIDENCIA') {
       ss.getSheetByName('DOCS_HISTORICO').appendRow([Utilities.getUuid(), tipoEntidad, idEntidad, nombreArchivo, file.getUrl(), 1, new Date(), Session.getActiveUser().getEmail(), file.getId()]);
    }
    return { success: true, fileId: file.getId(), url: file.getUrl() };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function eliminarDocumento(idDoc) {
  verificarPermiso(['DELETE']);
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const sheet = ss.getSheetByName('DOCS_HISTORICO');
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++){ if(String(data[i][0]) === String(idDoc)) { sheet.deleteRow(i+1); return { success: true }; } }
  return { success: false, error: "Documento no encontrado" };
}

// ==========================================
// 7. MANTENIMIENTO Y CALENDARIO
// ==========================================
function getCalendarId() { return Session.getActiveUser().getEmail(); } 

function gestionarEventoCalendario(accion, datos, eventIdExistente) {
  try {
    const cal = CalendarApp.getCalendarById(getCalendarId());
    if (!cal) return null;
    const titulo = `MANT: ${datos.tipo} - ${datos.nombreActivo}`;
    const descripcion = `Activo: ${datos.nombreActivo}\nMarca: ${datos.marca}\nEdificio: ${datos.edificio}\n\nGestión desde GMAO.`;
    const fechaInicio = textoAFecha(datos.fecha);
    if (!fechaInicio) return null;
    if (accion === 'CREAR') {
      const evento = cal.createAllDayEvent(titulo, fechaInicio, { description: descripcion, location: datos.edificio });
      evento.setColor(CalendarApp.EventColor.PALE_RED);
      return evento.getId();
    } else if (accion === 'ACTUALIZAR' && eventIdExistente) {
      const evento = cal.getEventById(eventIdExistente);
      if (evento) { evento.setTitle(titulo); evento.setAllDayDate(fechaInicio); evento.setDescription(descripcion); evento.setLocation(datos.edificio); }
      return eventIdExistente;
    } else if (accion === 'BORRAR' && eventIdExistente) {
      const evento = cal.getEventById(eventIdExistente); if (evento) evento.deleteEvent();
      return null;
    }
  } catch (e) { console.log("Error Calendar: " + e.toString()); return null; }
}

function getInfoParaCalendar(idActivo) {
  const activos = getSheetData('ACTIVOS'); const edificios = getSheetData('EDIFICIOS');
  let nombreActivo = "Activo"; let marca = "-"; let nombreEdificio = "Sin ubicación";
  for(let i=1; i<activos.length; i++){ if(String(activos[i][0]) === String(idActivo)) { nombreActivo = activos[i][3]; marca = activos[i][4]; const idEdif = activos[i][1]; for(let j=1; j<edificios.length; j++){ if(String(edificios[j][0]) === String(idEdif)) { nombreEdificio = edificios[j][2]; break; } } break; } }
  return { nombreActivo, marca, edificio: nombreEdificio };
}

function obtenerPlanMantenimiento(idActivo) {
  const data = getSheetData('PLAN_MANTENIMIENTO'); const docsData = getSheetData('DOCS_HISTORICO');
  const docsMap = {}; for(let j=1; j<docsData.length; j++) if(String(docsData[j][1]) === 'REVISION') docsMap[String(docsData[j][2])] = true;
  const planes = []; const hoy = new Date(); hoy.setHours(0,0,0,0);
  for(let i=1; i<data.length; i++) {
    if(String(data[i][1]) === String(idActivo)) {
      let f = data[i][4]; let color = 'gris'; let fechaStr = "-"; let fechaISO = "";
      if (f instanceof Date) { f.setHours(0,0,0,0); const diff = Math.ceil((f.getTime() - hoy.getTime()) / (86400000)); if (diff < 0) color = 'rojo'; else if (diff <= 30) color = 'amarillo'; else color = 'verde'; fechaStr = Utilities.formatDate(f, Session.getScriptTimeZone(), "dd/MM/yyyy"); fechaISO = Utilities.formatDate(f, Session.getScriptTimeZone(), "yyyy-MM-dd"); }
      let hasEvent = (data[i].length > 7 && data[i][7]) ? true : false;
      planes.push({ id: data[i][0], tipo: data[i][2], fechaProxima: fechaStr, fechaISO: fechaISO, color: color, hasDocs: docsMap[String(data[i][0])] || false, hasCalendar: hasEvent });
    }
  }
  return planes.sort((a, b) => a.fechaISO.localeCompare(b.fechaISO));
}

function getGlobalMaintenance() {
  const planes = getSheetData('PLAN_MANTENIMIENTO'); const activos = getSheetData('ACTIVOS'); const edificios = getSheetData('EDIFICIOS'); const docsData = getSheetData('DOCS_HISTORICO');
  const docsMap = {}; for(let j=1; j<docsData.length; j++) { if(String(docsData[j][1]) === 'REVISION') docsMap[String(docsData[j][2])] = true; }
  const mapEdificios = {}; edificios.slice(1).forEach(r => mapEdificios[r[0]] = { nombre: r[2], idCampus: r[1] });
  const mapActivos = {}; activos.slice(1).forEach(r => { const edificioInfo = mapEdificios[r[1]] || {}; mapActivos[r[0]] = { nombre: r[3], idEdif: r[1], idCampus: edificioInfo.idCampus || null }; });
  const result = []; const hoy = new Date(); hoy.setHours(0,0,0,0);
  for(let i=1; i<planes.length; i++) {
    const idActivo = planes[i][1]; const activoInfo = mapActivos[idActivo];
    if(activoInfo) {
       const nombreEdificio = mapEdificios[activoInfo.idEdif] ? mapEdificios[activoInfo.idEdif].nombre : "-";
       let f = planes[i][4]; let color = 'gris'; let fechaStr = "-"; let dias = 9999; let fechaISO = "";
       if (f instanceof Date) { f.setHours(0,0,0,0); dias = Math.ceil((f.getTime() - hoy.getTime()) / (86400000)); if (dias < 0) color = 'rojo'; else if (dias <= 30) color = 'amarillo'; else color = 'verde'; fechaStr = Utilities.formatDate(f, Session.getScriptTimeZone(), "dd/MM/yyyy"); fechaISO = Utilities.formatDate(f, Session.getScriptTimeZone(), "yyyy-MM-dd"); }
       let hasCalendar = (planes[i].length > 7 && planes[i][7]) ? true : false;
       result.push({ id: planes[i][0], idActivo: idActivo, activo: activoInfo.nombre, edificio: nombreEdificio, tipo: planes[i][2], fecha: fechaStr, fechaISO: fechaISO, color: color, dias: dias, edificioId: activoInfo.idEdif, campusId: activoInfo.idCampus, hasDocs: docsMap[String(planes[i][0])] || false, hasCalendar: hasCalendar });
    }
  }
  return result.sort((a,b) => a.dias - b.dias);
}

function crearRevision(d) {
  verificarPermiso(['WRITE']);
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const sheet = ss.getSheetByName('PLAN_MANTENIMIENTO');
  try {
    let fechaActual = textoAFecha(d.fechaProx); if (!fechaActual) return { success: false, error: "Fecha inválida" };
    var esRepetitiva = (String(d.esRecursiva) === "true"); var frecuencia = parseInt(d.diasFreq) || 0; var fechaLimite = d.fechaFin ? textoAFecha(d.fechaFin) : null; var syncCal = (String(d.syncCalendar) === "true");
    const infoExtra = syncCal ? getInfoParaCalendar(d.idActivo) : {};
    let eventId = null; if (syncCal) { eventId = gestionarEventoCalendario('CREAR', { ...infoExtra, tipo: d.tipo, fecha: d.fechaProx }); }
    sheet.appendRow([Utilities.getUuid(), d.idActivo, d.tipo, "", new Date(fechaActual), frecuencia, "ACTIVO", eventId]);
    if (esRepetitiva && frecuencia > 0 && fechaLimite && fechaLimite > fechaActual) {
      let contador = 0;
      while (contador < 50) { 
        fechaActual.setDate(fechaActual.getDate() + frecuencia); if (fechaActual > fechaLimite) break;
        let eventIdLoop = null; if (syncCal) { let fStr = Utilities.formatDate(fechaActual, Session.getScriptTimeZone(), "yyyy-MM-dd"); eventIdLoop = gestionarEventoCalendario('CREAR', { ...infoExtra, tipo: d.tipo, fecha: fStr }); }
        sheet.appendRow([Utilities.getUuid(), d.idActivo, d.tipo, "", new Date(fechaActual), frecuencia, "ACTIVO", eventIdLoop]); contador++;
      }
    }
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function updateRevision(d) { 
    verificarPermiso(['WRITE']); 
    const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const sheet = ss.getSheetByName('PLAN_MANTENIMIENTO'); const data = sheet.getDataRange().getValues(); 
    let nuevaFecha = textoAFecha(d.fechaProx); if (!nuevaFecha) return { success: false, error: "Fecha inválida" };
    var syncCal = (String(d.syncCalendar) === "true");
    for(let i=1; i<data.length; i++){ 
        if(String(data[i][0]) === String(d.idPlan)) { 
            sheet.getRange(i+1, 3).setValue(d.tipo); sheet.getRange(i+1, 5).setValue(nuevaFecha); 
            let currentEventId = (data[i].length > 7) ? data[i][7] : null;
            if (syncCal) {
                const infoExtra = getInfoParaCalendar(d.idActivo);
                if (currentEventId) { gestionarEventoCalendario('ACTUALIZAR', { ...infoExtra, tipo: d.tipo, fecha: d.fechaProx }, currentEventId); } 
                else { const newId = gestionarEventoCalendario('CREAR', { ...infoExtra, tipo: d.tipo, fecha: d.fechaProx }); sheet.getRange(i+1, 8).setValue(newId); }
            }
            break;
        } 
    } 
    return { success: true }; 
}

function eliminarRevision(idPlan) { 
  verificarPermiso(['DELETE']);
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const sheet = ss.getSheetByName('PLAN_MANTENIMIENTO'); const data = sheet.getDataRange().getValues(); 
  for(let i=1; i<data.length; i++){ 
    if(String(data[i][0]) === String(idPlan)) { 
       let currentEventId = (data[i].length > 7) ? data[i][7] : null; if (currentEventId) { gestionarEventoCalendario('BORRAR', {}, currentEventId); }
       sheet.deleteRow(i+1); return { success: true }; 
    } 
  } 
  return { success: false, error: "Plan no encontrado" }; 
}

// ==========================================
// 8. CONTRATOS
// ==========================================
function obtenerContratos(idEntidad) {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const sheet = ss.getSheetByName('CONTRATOS'); if (!sheet || sheet.getLastRow() < 2) return []; const data = sheet.getRange(1, 1, sheet.getLastRow(), 8).getValues(); const contratos = []; const hoy = new Date(); for(let i=1; i<data.length; i++) { if(String(data[i][2]) === String(idEntidad)) { const fFin = data[i][6] instanceof Date ? data[i][6] : null; const fIni = data[i][5] instanceof Date ? data[i][5] : null; let estadoDB = (data[i].length > 7) ? data[i][7] : 'ACTIVO'; let estadoCalc = 'VIGENTE'; let color = 'verde'; if (estadoDB === 'INACTIVO') { estadoCalc = 'INACTIVO'; color = 'gris'; } else if (fFin) { const diff = Math.ceil((fFin.getTime() - hoy.getTime()) / (86400000)); if (diff < 0) { estadoCalc = 'CADUCADO'; color = 'rojo'; } else if (diff <= 30) { estadoCalc = 'PRÓXIMO'; color = 'amarillo'; } } else { estadoCalc = 'SIN FECHA'; color = 'gris'; } contratos.push({ id: data[i][0], proveedor: data[i][3], ref: data[i][4], inicio: fIni?fechaATexto(fIni):"-", fin: fFin?fechaATexto(fFin):"-", estado: estadoCalc, color: color, estadoDB: estadoDB }); } } return contratos.sort((a, b) => a.fin.localeCompare(b.fin));
}
function obtenerContratosGlobal() {
  const contratos = getSheetData('CONTRATOS'); const activos = getSheetData('ACTIVOS'); const edificios = getSheetData('EDIFICIOS'); const campus = getSheetData('CAMPUS'); 
  const mapCampus = {}; campus.slice(1).forEach(r => mapCampus[r[0]] = r[1]); const mapEdificios = {}; edificios.slice(1).forEach(r => mapEdificios[r[0]] = { nombre: r[2], idCampus: r[1] }); const mapActivos = {}; activos.slice(1).forEach(r => { const edificioInfo = mapEdificios[r[1]] || {}; mapActivos[r[0]] = { nombre: r[3], idEdif: r[1], idCampus: edificioInfo.idCampus || null }; });
  const result = []; const hoy = new Date();
  for(let i=1; i<contratos.length; i++) {
    const r = contratos[i]; const idEntidad = r[2]; const tipoEntidad = r[1];
    let nombreEntidad = "N/A"; let edificioId = null; let campusId = null;
    if (tipoEntidad === 'ACTIVO' && mapActivos[idEntidad]) { const info = mapActivos[idEntidad]; nombreEntidad = info.nombre + " (" + (mapEdificios[info.idEdif] ? mapEdificios[info.idEdif].nombre : 'Sin Edif.') + ")"; edificioId = info.idEdif; campusId = info.idCampus; } 
    else if (tipoEntidad === 'EDIFICIO' && mapEdificios[idEntidad]) { const info = mapEdificios[idEntidad]; nombreEntidad = info.nombre; edificioId = idEntidad; campusId = info.idCampus; }
    const fFin = (r[6] instanceof Date) ? r[6] : null; let estadoDB = r[7] || 'ACTIVO'; let estadoCalc = 'VIGENTE'; let color = 'verde';
    if (estadoDB === 'INACTIVO') { estadoCalc = 'INACTIVO'; color = 'gris'; } else if (fFin) { const diff = Math.ceil((fFin.getTime() - hoy.getTime()) / 86400000); if (diff < 0) { estadoCalc = 'CADUCADO'; color = 'rojo'; } else if (diff <= 90) { estadoCalc = 'PRÓXIMO'; color = 'amarillo'; } } else { estadoCalc = 'SIN FECHA'; color = 'gris'; }
    result.push({ id: r[0], nombreEntidad: nombreEntidad, proveedor: r[3], ref: r[4], inicio: r[5] ? fechaATexto(r[5]) : "-", fin: fFin ? fechaATexto(fFin) : "-", estado: estadoCalc, color: color, estadoDB: estadoDB, edificioId: edificioId, campusId: campusId });
  }
  return result.sort((a, b) => { if (a.fin === "-") return 1; if (b.fin === "-") return -1; return a.fin.localeCompare(b.fin); });
}
function crearContrato(d) { verificarPermiso(['WRITE']); const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); ss.getSheetByName('CONTRATOS').appendRow([Utilities.getUuid(), d.tipoEntidad, d.idEntidad, d.proveedor, d.ref, textoAFecha(d.fechaIni), textoAFecha(d.fechaFin), d.estado]); return { success: true }; }
function updateContrato(datos) { verificarPermiso(['WRITE']); const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const sheet = ss.getSheetByName('CONTRATOS'); const data = sheet.getDataRange().getValues(); for(let i=1; i<data.length; i++){ if(String(data[i][0]) === String(datos.id)) { sheet.getRange(i+1, 4).setValue(datos.proveedor); sheet.getRange(i+1, 5).setValue(datos.ref); sheet.getRange(i+1, 6).setValue(textoAFecha(datos.fechaIni)); sheet.getRange(i+1, 7).setValue(textoAFecha(datos.fechaFin)); sheet.getRange(i+1, 8).setValue(datos.estado); return { success: true }; } } return { success: false, error: "Contrato no encontrado" }; }
function eliminarContrato(id) { verificarPermiso(['DELETE']); const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const sheet = ss.getSheetByName('CONTRATOS'); const data = sheet.getDataRange().getValues(); for(let i=1; i<data.length; i++){ if(String(data[i][0]) === String(id)) { sheet.deleteRow(i+1); return { success: true }; } } return { success: false, error: "Contrato no encontrado" }; }

// ==========================================
// 9. DASHBOARD Y CRUD GENERAL (ACTUALIZADO V5)
// ==========================================
function getDatosDashboard() { 
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); 
  const hoy = new Date(); hoy.setHours(0,0,0,0);
  
  // 1. MANTENIMIENTO
  const dataMant = getSheetData('PLAN_MANTENIMIENTO');
  const activos = getSheetData('ACTIVOS');
  
  // Mapa de nombres de activos para el calendario
  const mapActivos = {};
  activos.slice(1).forEach(r => mapActivos[r[0]] = r[3]); // ID -> Nombre

  let revPend = 0, revVenc = 0, revOk = 0; 
  const calendarEvents = [];

  // Gráfico Lineal (Meses)
  const mesesNombres = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];
  const countsMap = {}; const chartLabels = []; const chartData = [];
  let dIter = new Date(hoy.getFullYear(), hoy.getMonth(), 1); 
  for (let k = 0; k < 6; k++) { 
    let key = mesesNombres[dIter.getMonth()] + " " + dIter.getFullYear(); 
    countsMap[key] = 0; chartLabels.push(key); 
    dIter.setMonth(dIter.getMonth() + 1); 
  }

  for(let i=1; i<dataMant.length; i++) { 
    const f = dataMant[i][4]; 
    if(f instanceof Date) { 
      f.setHours(0,0,0,0); 
      const diff = Math.ceil((f.getTime() - hoy.getTime()) / 86400000); 
      let status = 'verde';
      
      if(diff < 0) { revVenc++; status = 'rojo'; }
      else if(diff <= 30) { revPend++; status = 'amarillo'; }
      else { revOk++; status = 'verde'; }
      
      // Datos para gráfico lineal
      let key = mesesNombres[f.getMonth()] + " " + f.getFullYear(); 
      if (countsMap.hasOwnProperty(key)) { countsMap[key]++; }
      
      // Datos para Calendario
      let nombreActivo = mapActivos[dataMant[i][1]] || 'Activo';
      let colorEvento = (status === 'rojo') ? '#dc3545' : (status === 'amarillo' ? '#ffc107' : '#198754');
      let textColor = (status === 'amarillo') ? '#000' : '#fff';
      
      calendarEvents.push({
        id: dataMant[i][0], // ID Plan
        title: `${dataMant[i][2]} - ${nombreActivo}`,
        start: Utilities.formatDate(f, Session.getScriptTimeZone(), "yyyy-MM-dd"),
        backgroundColor: colorEvento,
        borderColor: colorEvento,
        textColor: textColor,
        extendedProps: {
           id: dataMant[i][0],
           tipo: dataMant[i][2],
           fechaISO: Utilities.formatDate(f, Session.getScriptTimeZone(), "yyyy-MM-dd"),
           idActivo: dataMant[i][1],
           hasCalendar: (dataMant[i].length > 7 && dataMant[i][7]) ? true : false
        }
      });
    } 
  } 
  
  chartLabels.forEach(lbl => { chartData.push(countsMap[lbl]); });

  // 2. OTROS CONTEOS
  const dataCont = getSheetData('CONTRATOS'); 
  let contCount = (dataCont.length > 1) ? dataCont.length - 1 : 0;
  
  const dataInc = getSheetData('INCIDENCIAS');
  let incCount = 0;
  // Contamos solo pendientes o en proceso
  for(let i=1; i<dataInc.length; i++) {
     if(dataInc[i][6] !== 'RESUELTA') incCount++;
  }

  const cAct = (activos.length > 1) ? activos.length - 1 : 0; 
  const cEdif = (getSheetData('EDIFICIOS').length - 1) || 0;
  const cCampus = (getSheetData('CAMPUS').length - 1) || 0;

  return { 
    activos: cAct, 
    edificios: cEdif, 
    pendientes: revPend, 
    vencidas: revVenc, 
    ok: revOk, 
    contratos: contCount,
    incidencias: incCount,
    campus: cCampus,
    chartLabels: chartLabels, 
    chartData: chartData,
    calendarEvents: calendarEvents // <--- Nuevo para el calendario
  }; 
}

function crearCampus(d) { verificarPermiso(['WRITE']); const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const fId = crearCarpeta(d.nombre, getRootFolderId()); ss.getSheetByName('CAMPUS').appendRow([Utilities.getUuid(), d.nombre, d.provincia, d.direccion, fId]); return {success:true}; }
function crearEdificio(d) { verificarPermiso(['WRITE']); const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const cData = getSheetData('CAMPUS'); let pId; for(let i=1; i<cData.length; i++) if(String(cData[i][0])==String(d.idCampus)) pId=cData[i][4]; const fId = crearCarpeta(d.nombre, pId); const aId = crearCarpeta("Activos", fId); ss.getSheetByName('EDIFICIOS').appendRow([Utilities.getUuid(), d.idCampus, d.nombre, d.contacto, fId, aId]); return {success:true}; }
function crearActivo(d) { verificarPermiso(['WRITE']); const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const eData = getSheetData('EDIFICIOS'); let pId; for(let i=1; i<eData.length; i++) if(String(eData[i][0])==String(d.idEdificio)) pId=eData[i][5]; const fId = crearCarpeta(d.nombre, pId); const id = Utilities.getUuid(); const cats = getSheetData('CAT_INSTALACIONES'); let nombreTipo = d.tipo; for(let i=1; i<cats.length; i++) { if(String(cats[i][0]) === String(d.tipo)) { nombreTipo = cats[i][1]; break; } } ss.getSheetByName('ACTIVOS').appendRow([id, d.idEdificio, nombreTipo, d.nombre, d.marca, new Date(), fId]); return {success:true}; }
function getCatalogoInstalaciones() { return getSheetData('CAT_INSTALACIONES').slice(1).map(r=>({id:r[0], nombre:r[1], dias:r[3]})); }
function getTableData(tipo) { const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); if (tipo === 'CAMPUS') { const data = getSheetData('CAMPUS'); return data.slice(1).map(r => ({ id: r[0], nombre: r[1], provincia: r[2], direccion: r[3] })); } if (tipo === 'EDIFICIOS') { const data = getSheetData('EDIFICIOS'); const dataC = getSheetData('CAMPUS'); const mapCampus = {}; dataC.slice(1).forEach(r => mapCampus[r[0]] = r[1]); return data.slice(1).map(r => ({ id: r[0], campus: mapCampus[r[1]] || '-', nombre: r[2], contacto: r[3] })); } return []; }
function buscarGlobal(texto) { if (!texto || texto.length < 3) return []; texto = texto.toLowerCase(); const resultados = []; const activos = getSheetData('ACTIVOS'); for(let i=1; i<activos.length; i++) { const r = activos[i]; if (String(r[3]).toLowerCase().includes(texto) || String(r[2]).toLowerCase().includes(texto) || String(r[4]).toLowerCase().includes(texto)) { resultados.push({ id: r[0], tipo: 'ACTIVO', texto: r[3], subtexto: r[2] + (r[4] ? " - " + r[4] : "") }); } } const edificios = getSheetData('EDIFICIOS'); for(let i=1; i<edificios.length; i++) { const r = edificios[i]; if (String(r[2]).toLowerCase().includes(texto)) { resultados.push({ id: r[0], tipo: 'EDIFICIO', texto: r[2], subtexto: 'Edificio' }); } } return resultados.slice(0, 10); }

// ==========================================
// 10. GESTIÓN USUARIOS Y CONFIG (ADMIN ONLY)
// ==========================================
function getConfigCatalogo() { const data = getSheetData('CAT_INSTALACIONES'); return data.slice(1).map(r => ({ id: r[0], nombre: r[1], normativa: r[2], dias: r[3] })); }
function saveConfigCatalogo(d) { verificarPermiso(['ADMIN_ONLY']); const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const sheet = ss.getSheetByName('CAT_INSTALACIONES'); if(!d.nombre) return { success: false, error: "Nombre obligatorio" }; if (d.id) { const data = sheet.getDataRange().getValues(); for(let i=1; i<data.length; i++) { if(String(data[i][0]) === String(d.id)) { sheet.getRange(i+1, 2).setValue(d.nombre); sheet.getRange(i+1, 3).setValue(d.normativa); sheet.getRange(i+1, 4).setValue(d.dias); return { success: true }; } } return { success: false, error: "ID no encontrado" }; } else { sheet.appendRow([Utilities.getUuid(), d.nombre, d.normativa, d.dias]); return { success: true }; } }
function deleteConfigCatalogo(id) { verificarPermiso(['ADMIN_ONLY']); const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const sheet = ss.getSheetByName('CAT_INSTALACIONES'); const data = sheet.getDataRange().getValues(); for(let i=1; i<data.length; i++){ if(String(data[i][0]) === String(id)) { sheet.deleteRow(i+1); return { success: true }; } } return { success: false, error: "No encontrado" }; }

function getListaUsuarios() { const data = getSheetData('USUARIOS'); return data.slice(1).map(r => ({ id: r[0], nombre: r[1], email: r[2], rol: r[3], avisos: r[4] })); }
function saveUsuario(u) { verificarPermiso(['ADMIN_ONLY']); const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const sheet = ss.getSheetByName('USUARIOS'); if (!u.nombre || !u.email) return { success: false, error: "Datos incompletos" }; if (u.id) { const data = sheet.getDataRange().getValues(); for(let i=1; i<data.length; i++){ if(String(data[i][0]) === String(u.id)) { sheet.getRange(i+1, 2).setValue(u.nombre); sheet.getRange(i+1, 3).setValue(u.email); sheet.getRange(i+1, 4).setValue(u.rol); sheet.getRange(i+1, 5).setValue(u.avisos); return { success: true }; } } return { success: false, error: "Usuario no encontrado" }; } else { sheet.appendRow([Utilities.getUuid(), u.nombre, u.email, u.rol, u.avisos]); return { success: true }; } }
function deleteUsuario(id) { verificarPermiso(['ADMIN_ONLY']); const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const sheet = ss.getSheetByName('USUARIOS'); const data = sheet.getDataRange().getValues(); for(let i=1; i<data.length; i++){ if(String(data[i][0]) === String(id)) { sheet.deleteRow(i+1); return { success: true }; } } return { success: false, error: "Usuario no encontrado" }; }

// ==========================================
// 11. NOTIFICACIONES AUTOMÁTICAS
// ==========================================
function enviarResumenSemanal() {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const sheetUsers = ss.getSheetByName('USUARIOS');
  if(!sheetUsers) return;
  const usuarios = sheetUsers.getDataRange().getValues().slice(1);
  const emailsDestino = usuarios.filter(u => String(u[4]).toUpperCase() === 'SI' && u[2]).map(u => u[2]).join(",");
  if(!emailsDestino) return;
  const mantenimiento = getGlobalMaintenance(); 
  const mantCriticos = mantenimiento.filter(m => m.color === 'rojo' || m.color === 'amarillo');
  const contratos = obtenerContratosGlobal(); 
  const contCriticos = contratos.filter(c => c.color === 'rojo' || c.color === 'amarillo');
  if (mantCriticos.length === 0 && contCriticos.length === 0) return;
  let html = `<div style="font-family: Arial, sans-serif; color: #333;"><h2 style="color: #0d6efd;">Resumen GMAO</h2><p>Estado de activos y contratos:</p>`;
  if (mantCriticos.length > 0) { html += `<h3>🔧 Revisiones Pendientes (${mantCriticos.length})</h3><ul>`; mantCriticos.forEach(m => { html += `<li><b>${m.color === 'rojo' ? 'VENCIDO' : 'PRÓXIMO'}:</b> ${m.activo} (${m.edificio}) - ${m.tipo} - ${m.fecha}</li>`; }); html += `</ul>`; }
  if (contCriticos.length > 0) { html += `<h3>📄 Contratos por Vencer (${contCriticos.length})</h3><ul>`; contCriticos.forEach(c => { html += `<li><b>${c.estado}:</b> ${c.proveedor} (${c.fin})</li>`; }); html += `</ul>`; }
  html += `<p style="color: #666; font-size: 12px;">Generado automáticamente.</p></div>`;
  MailApp.sendEmail({ to: emailsDestino, subject: `[GMAO] Alerta: ${mantCriticos.length + contCriticos.length} incidencias`, htmlBody: html });
}

// ==========================================
// 12. GESTIÓN DE INCIDENCIAS
// ==========================================
function getIncidencias() {
  const data = getSheetData('INCIDENCIAS');
  const list = [];
  for(let i=1; i<data.length; i++) {
    if(data[i][0]) {
       let f = data[i][7];
       let fechaStr = (f instanceof Date) ? Utilities.formatDate(f, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : "-";
       let urlFoto = "";
       if(data[i][9]) { try { urlFoto = DriveApp.getFileById(data[i][9]).getUrl(); } catch(e) {} }
       list.push({ id: data[i][0], tipoOrigen: data[i][1], nombreOrigen: data[i][3], descripcion: data[i][4], prioridad: data[i][5], estado: data[i][6], fecha: fechaStr, solicitante: data[i][8], urlFoto: urlFoto });
    }
  }
  return list.sort((a,b) => {
    if(a.estado === 'RESUELTA' && b.estado !== 'RESUELTA') return 1;
    if(a.estado !== 'RESUELTA' && b.estado === 'RESUELTA') return -1;
    return b.fecha.localeCompare(a.fecha);
  });
}

function crearIncidencia(d) {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const sheet = ss.getSheetByName('INCIDENCIAS');
  const usuario = getMyRole().email;
  const fecha = new Date();
  try {
    let idFoto = "";
    if (d.fotoBase64) {
       let carpetaId = getRootFolderId(); 
       if (d.idOrigen) {
          const activos = getSheetData('ACTIVOS');
          for(let i=1; i<activos.length; i++) if(String(activos[i][0])===String(d.idOrigen)) { carpetaId = activos[i][6]; break; }
          if (carpetaId === getRootFolderId()) { 
             const edifs = getSheetData('EDIFICIOS');
             for(let i=1; i<edifs.length; i++) if(String(edifs[i][0])===String(d.idOrigen)) { carpetaId = edifs[i][4]; break; }
          }
       }
       const blob = Utilities.newBlob(Utilities.base64Decode(d.fotoBase64), d.mimeType, "INCIDENCIA_" + d.nombreArchivo);
       const file = DriveApp.getFolderById(carpetaId).createFile(blob);
       file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
       idFoto = file.getId();
    }
    sheet.appendRow([Utilities.getUuid(), d.tipoOrigen, d.idOrigen, d.nombreOrigen, d.descripcion, d.prioridad, "PENDIENTE", fecha, usuario, idFoto]);
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function actualizarEstadoIncidencia(id, nuevoEstado) {
  verificarPermiso(['WRITE']); 
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const sheet = ss.getSheetByName('INCIDENCIAS'); const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++){ if(String(data[i][0]) === String(id)) { sheet.getRange(i+1, 7).setValue(nuevoEstado); return { success: true }; } }
  return { success: false, error: "Incidencia no encontrada" };
}

function getIncidenciaDetalle(id) {
  const data = getSheetData('INCIDENCIAS'); const activos = getSheetData('ACTIVOS'); const edificios = getSheetData('EDIFICIOS');
  for(let i=1; i<data.length; i++) {
    if(String(data[i][0]) === String(id)) {
       const r = data[i];
       let idCampus = null; let idEdificio = null; let idActivo = null;
       if (r[1] === 'ACTIVO') {
          idActivo = r[2]; for(let a=1; a<activos.length; a++) { if(String(activos[a][0]) === String(idActivo)) { idEdificio = activos[a][1]; break; } }
       } else { idEdificio = r[2]; }
       if(idEdificio) { for(let e=1; e<edificios.length; e++) { if(String(edificios[e][0]) === String(idEdificio)) { idCampus = edificios[e][1]; break; } } }
       let urlFoto = null; if (r[9]) { try { urlFoto = DriveApp.getFileById(r[9]).getUrl(); } catch(e) {} }
       return { id: r[0], tipoOrigen: r[1], idOrigen: r[2], nombreOrigen: r[3], descripcion: r[4], prioridad: r[5], idCampus: idCampus, idEdificio: idEdificio, idActivo: idActivo, urlFoto: urlFoto };
    }
  }
  throw new Error("Incidencia no encontrada");
}

function updateIncidenciaData(d) {
  verificarPermiso(['WRITE']); 
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const sheet = ss.getSheetByName('INCIDENCIAS'); const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++){
    if(String(data[i][0]) === String(d.id)) {
       sheet.getRange(i+1, 2).setValue(d.tipoOrigen); sheet.getRange(i+1, 3).setValue(d.idOrigen); sheet.getRange(i+1, 4).setValue(d.nombreOrigen);
       sheet.getRange(i+1, 5).setValue(d.descripcion); sheet.getRange(i+1, 6).setValue(d.prioridad);
       return { success: true };
    }
  }
  return { success: false, error: "ID no encontrado" };
}

// ==========================================
// 13. GESTIÓN DE CAMPUS (EDITAR/BORRAR)
// ==========================================

function updateCampus(d) {
  verificarPermiso(['WRITE']);
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const sheet = ss.getSheetByName('CAMPUS');
  const data = sheet.getDataRange().getValues();
  
  for(let i=1; i<data.length; i++){
    if(String(data[i][0]) === String(d.id)) {
      sheet.getRange(i+1, 2).setValue(d.nombre);
      sheet.getRange(i+1, 3).setValue(d.provincia);
      sheet.getRange(i+1, 4).setValue(d.direccion);
      // Nota: No cambiamos la carpeta de Drive (Columna 5) para no romper enlaces
      return { success: true };
    }
  }
  return { success: false, error: "Campus no encontrado" };
}

function eliminarCampus(id) {
  verificarPermiso(['DELETE']); // Solo Admin
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const sheet = ss.getSheetByName('CAMPUS');
  const data = sheet.getDataRange().getValues();
  
  // Opcional: Verificar si tiene edificios antes de borrar para evitar huérfanos
  
  for(let i=1; i<data.length; i++){
    if(String(data[i][0]) === String(id)) {
      sheet.deleteRow(i+1);
      return { success: true };
    }
  }
  return { success: false, error: "Campus no encontrado" };
}

// ==========================================
// 14. GESTIÓN DE EDIFICIOS (EDITAR/BORRAR)
// ==========================================

function updateEdificio(d) {
  verificarPermiso(['WRITE']);
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const sheet = ss.getSheetByName('EDIFICIOS');
  const data = sheet.getDataRange().getValues();
  
  for(let i=1; i<data.length; i++){
    if(String(data[i][0]) === String(d.id)) {
      // Col B: ID_Campus, Col C: Nombre, Col D: Contacto
      sheet.getRange(i+1, 2).setValue(d.idCampus);
      sheet.getRange(i+1, 3).setValue(d.nombre);
      sheet.getRange(i+1, 4).setValue(d.contacto);
      return { success: true };
    }
  }
  return { success: false, error: "Edificio no encontrado" };
}

function eliminarEdificio(id) {
  verificarPermiso(['DELETE']);
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const sheet = ss.getSheetByName('EDIFICIOS');
  const data = sheet.getDataRange().getValues();
  
  for(let i=1; i<data.length; i++){
    if(String(data[i][0]) === String(id)) {
      sheet.deleteRow(i+1);
      return { success: true };
    }
  }
  return { success: false, error: "Edificio no encontrado" };
}
