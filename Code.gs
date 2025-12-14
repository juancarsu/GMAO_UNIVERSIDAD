// GMAO
// Universidad de Navarra
// Versión 1.2
// Autor: Juan Carlos Suárez
//
// Licencia: Creative Commons Reconocimiento (CC BY) - creativecommons.org
// Puedes usar, copiar, modificar y distribuir este código (sin fines comerciales),
// siempre que cites a Juan Carlos Suárez como autor original.

// ==========================================
// 1. CONFIGURACIÓN Y ROUTING
// ==========================================
const PROPS = PropertiesService.getScriptProperties();

// OPTIMIZACIÓN: SINGLETON PARA SPREADSHEET
let _SS_INSTANCE = null;

function getDB() {
  if (!_SS_INSTANCE) {
    // Abrimos el archivo UNA sola vez por ejecución
    _SS_INSTANCE = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  }
  return _SS_INSTANCE;
}

function getSheetData(sheetName) {
  const ss = getDB(); // Usamos la instancia cacheada en memoria
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet && sheetName === 'USUARIOS') {
     sheet = ss.insertSheet('USUARIOS');
     sheet.appendRow(['ID', 'NOMBRE', 'EMAIL', 'ROL', 'RECIBIR_AVISOS']);
  }
  
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getDataRange().getValues();
}

function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  template.urlParams = e ? e.parameter : {};
  return template.evaluate().setTitle('GMAO Universidad').addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }

/**
 * OBTENER TODO DE UNA VEZ PARA EL FRONTEND
 * Devuelve: Usuario, Dashboard, Filtros Campus, Catálogo
 */
function getAppInitData() {
  const t1 = new Date();
  
  // 1. Obtener Usuario
  const usuario = getMyRole();
  
  // 2. Obtener Dashboard (reutiliza la lógica existente)
  // Nota: getDatosDashboard ya usa cachés, así que es rápido
  const dashboard = getDatosDashboard(null); 
  
  // 3. Obtener Listas para Filtros (Campus y Catálogo)
  const campusRaw = getCachedSheetData('CAMPUS');
  const listaCampus = campusRaw.slice(1).map(r => ({ id: r[0], nombre: r[1] }));
  
  const catalogoRaw = getCachedSheetData('CAT_INSTALACIONES');
  const catalogo = catalogoRaw.slice(1).map(r=>({id:r[0], nombre:r[1], dias:r[3]}));

  // Log de tiempo para depuración
  Logger.log("Tiempo InitData: " + (new Date() - t1) + "ms");

  return {
    usuario: usuario,
    dashboard: dashboard,
    listaCampus: listaCampus,
    catalogo: catalogo
  };
}

// ==========================================
// SISTEMA DE CACHÉ OPTIMIZADO
// ==========================================

const CACHE = CacheService.getScriptCache();
const CACHE_TIME = 300; // 5 minutos

/**
 * Obtener datos con caché automática
 */
function getCachedSheetData(sheetName, forceRefresh = false) {
  const cacheKey = 'SHEET_' + sheetName;
  
  if (!forceRefresh) {
    const cached = CACHE.get(cacheKey);
    if (cached) {
      try {
        return JSON.parse(cached);
      } catch(e) {
        Logger.log('Error parsing cache: ' + e);
      }
    }
  }
  
  // Si no hay caché o falló, cargar desde Sheet
  const data = getSheetData(sheetName);
  
  // Guardar en caché (solo si no es muy grande)
  try {
    const serialized = JSON.stringify(data);
    if (serialized.length < 90000) { // Límite de CacheService
      CACHE.put(cacheKey, serialized, CACHE_TIME);
    }
  } catch(e) {
    Logger.log('No se pudo cachear ' + sheetName + ': ' + e);
  }
  
  return data;
}

/**
 * Invalidar caché cuando se modifiquen datos
 */
function invalidateCache(sheetName) {
  CACHE.remove('SHEET_' + sheetName);
  // También invalida cachés relacionados
  if (sheetName === 'ACTIVOS') {
    CACHE.remove('INDEX_ACTIVOS');
  }
  if (sheetName === 'EDIFICIOS') {
    CACHE.remove('INDEX_EDIFICIOS');
  }
}

/**
 * Crear índices optimizados para búsquedas rápidas
 */
function buildActivosIndex() {
  const cacheKey = 'INDEX_ACTIVOS';
  
  try {
    const cached = CACHE.get(cacheKey);
    if (cached) {
      return JSON.parse(cached);
    }
  } catch(e) {
    Logger.log('Error leyendo caché de índice: ' + e);
  }
  
  try {
    const activos = getCachedSheetData('ACTIVOS');
    const edificios = getCachedSheetData('EDIFICIOS');
    const campus = getCachedSheetData('CAMPUS');
    
    // Crear mapas rápidos
    const mapCampus = {};
    if (campus && campus.length > 1) {
      campus.slice(1).forEach(r => {
        if (r && r[0] && r[1]) {
          mapCampus[String(r[0])] = r[1];
        }
      });
    }
    
    const mapEdificios = {};
    if (edificios && edificios.length > 1) {
      edificios.slice(1).forEach(r => {
        if (r && r[0] && r[1] && r[2]) {
          mapEdificios[String(r[0])] = {
            nombre: r[2],
            idCampus: String(r[1]),
            campusNombre: mapCampus[String(r[1])] || '-'
          };
        }
      });
    }
    
    // Índice principal de activos
    const index = {
      byId: {},
      byEdificio: {},
      byCampus: {},
      searchable: []
    };
    
    if (activos && activos.length > 1) {
      activos.slice(1).forEach(r => {
        if (!r || !r[0]) return; // Skip filas vacías
        
        const id = String(r[0]);
        const idEdif = String(r[1] || '');
        const edifInfo = mapEdificios[idEdif] || {};
        
        const activo = {
          id: id,
          idEdificio: idEdif,
          tipo: r[2] || '-',
          nombre: r[3] || 'Sin nombre',
          marca: r[4] || '',
          fechaAlta: r[5] || null,
          edificio: edifInfo.nombre || '-',
          campus: edifInfo.campusNombre || '-',
          idCampus: edifInfo.idCampus || null
        };
        
        // Por ID
        index.byId[id] = activo;
        
        // Por Edificio
        if (idEdif) {
          if (!index.byEdificio[idEdif]) index.byEdificio[idEdif] = [];
          index.byEdificio[idEdif].push(activo);
        }
        
        // Por Campus
        if (edifInfo.idCampus) {
          if (!index.byCampus[edifInfo.idCampus]) index.byCampus[edifInfo.idCampus] = [];
          index.byCampus[edifInfo.idCampus].push(activo);
        }
        
        // Para búsquedas de texto
        const searchText = (activo.nombre + ' ' + activo.tipo + ' ' + activo.marca).toLowerCase();
        index.searchable.push({
          id: id,
          text: searchText
        });
      });
    }
    
    // Guardar índice en caché
    try {
      const serialized = JSON.stringify(index);
      if (serialized.length < 90000) {
        CACHE.put(cacheKey, serialized, CACHE_TIME);
      } else {
        Logger.log('Índice demasiado grande para caché: ' + serialized.length);
      }
    } catch(e) {
      Logger.log('Error guardando índice en caché: ' + e);
    }
    
    return index;
    
  } catch(e) {
    Logger.log('Error construyendo índice: ' + e.toString());
    // Devolver índice vacío en caso de error
    return {
      byId: {},
      byEdificio: {},
      byCampus: {},
      searchable: []
    };
  }
}

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
  // 1. Intentar leer de propiedades (Memoria ultra rápida)
  let id = PROPS.getProperty('ROOT_FOLDER_CACHE');
  if (id) return id;

  // 2. Si no está, leer de la hoja (Lento, solo la primera vez)
  const data = getSheetData('CONFIG');
  for(let row of data) { 
    if(row[0] === 'ROOT_FOLDER_ID') {
      id = row[1];
      // Guardamos en caché persistente
      PROPS.setProperty('ROOT_FOLDER_CACHE', id);
      return id;
    } 
  }
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

function getActivosPorEdificio(idEdificio) {
  const index = buildActivosIndex();
  return index.byEdificio[String(idEdificio)] || [];
}

function getAllAssetsList() {
  const index = buildActivosIndex();
  const allIds = Object.keys(index.byId);
  const list = [];
  
  // 1. Cargar datos necesarios
  const docsData = getCachedSheetData('DOCS_HISTORICO');
  const planData = getCachedSheetData('PLAN_MANTENIMIENTO');

  // 2. Mapear qué IDs (Activos o Revisiones) tienen documentos
  //    docsSet: Set con los IDs de las entidades que tienen docs
  const docsSet = new Set();
  if (docsData && docsData.length > 1) {
    docsData.slice(1).forEach(r => {
      // r[1] = TIPO (ACTIVO, REVISION...), r[2] = ID
      docsSet.add(String(r[2])); 
    });
  }

  // 3. Encontrar la ÚLTIMA revisión LEGAL REALIZADA con documentos para cada activo
  const mapLegalDocs = {}; // { idActivo: true/false }

  if (planData && planData.length > 1) {
    // Ordenamos cronológicamente ascendente para procesar fechas
    // (O simplemente recorremos y comparamos fechas)
    for (let i = 1; i < planData.length; i++) {
      const r = planData[i];
      const idPlan = String(r[0]);
      const idActivo = String(r[1]);
      const tipo = r[2];
      const fecha = r[4]; // Fecha programada o realizada
      const estado = r[6]; // Estado (REALIZADA, ACTIVO...)

      // Criterio: Debe ser LEGAL y estar REALIZADA
      if (tipo === 'Legal' && estado === 'REALIZADA') {
        const fechaObj = (fecha instanceof Date) ? fecha : new Date(fecha);
        
        // Si esta revisión tiene documentos
        const tieneDocs = docsSet.has(idPlan);

        // Lógica de "la última":
        // Si no tenemos dato previo para este activo, o esta fecha es más reciente que la guardada
        if (!mapLegalDocs[idActivo] || mapLegalDocs[idActivo].fecha < fechaObj) {
          mapLegalDocs[idActivo] = {
            fecha: fechaObj,
            hasDocs: tieneDocs
          };
        }
      }
    }
  }

  // 4. Construir la lista final
  for (const id of allIds) {
    const a = index.byId[id];
    // Verificar si existe registro legal y si tiene docs
    const legalInfo = mapLegalDocs[id];
    const legalOk = legalInfo ? legalInfo.hasDocs : false;

    list.push({
      id: a.id,
      nombre: a.nombre,
      tipo: a.tipo,
      marca: a.marca,
      idEdificio: a.idEdificio,
      edificioNombre: a.edificio,
      idCampus: a.idCampus,
      campusNombre: a.campus,
      hasDocs: docsSet.has(a.id), // Documentos generales del activo
      hasLegalDocs: legalOk       // NUEVO: Documentos de la última legal
    });
  }
  
  // Ordenar alfabéticamente por nombre
  return list.sort((a,b) => a.nombre.localeCompare(b.nombre));
}

function getAssetInfo(idActivo) {
  const index = buildActivosIndex();
  const activo = index.byId[String(idActivo)];
  
  if (!activo) return null;
  
  return {
    id: activo.id,
    nombre: activo.nombre,
    tipo: activo.tipo,
    marca: activo.marca,
    fechaAlta: activo.fechaAlta instanceof Date ? 
               Utilities.formatDate(activo.fechaAlta, Session.getScriptTimeZone(), "dd/MM/yyyy") : "-",
    edificio: activo.edificio,
    campus: activo.campus
  };
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
  invalidateCache('ACTIVOS');
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
  registrarLog("CREAR OBRA", "Nombre: " + d.nombre + " | Edificio ID: " + d.idEdificio);
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
      sheet.getRange(i+1, 6).setValue(fechaObj); sheet.getRange(i+1, 7).setValue("FINALIZADA"); 
      registrarLog("FINALIZAR OBRA", "ID Obra: " + idObra + " | Fecha Fin: " + fechaFin);
      return { success: true };
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
    if(String(data[i][0]) === String(id)) { sheet.deleteRow(i+1); 
    registrarLog("ELIMINAR OBRA", "ID Obra eliminada: " + id);
    return { success: true }; }
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

function gestionarEventoCalendario(accion, datos, eventIdExistente) {
  try {
    const cal = CalendarApp.getCalendarById(Session.getActiveUser().getEmail());
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
  const data = getSheetData('PLAN_MANTENIMIENTO'); 
  const docsData = getSheetData('DOCS_HISTORICO');
  const docsMap = {}; 
  for(let j=1; j<docsData.length; j++) if(String(docsData[j][1]) === 'REVISION') docsMap[String(docsData[j][2])] = true;
  
  const planes = []; 
  const hoy = new Date(); hoy.setHours(0,0,0,0);
  
  for(let i=1; i<data.length; i++) {
    // FILTRO NUEVO: Si está realizada, saltar
    if (data[i][6] === 'REALIZADA') continue; 

    if(String(data[i][1]) === String(idActivo)) {
      let f = data[i][4]; let color = 'gris'; let fechaStr = "-"; let fechaISO = "";
      if (f instanceof Date) { f.setHours(0,0,0,0); const diff = Math.ceil((f.getTime() - hoy.getTime()) / (86400000)); if (diff < 0) color = 'rojo'; else if (diff <= 30) color = 'amarillo'; else color = 'verde'; fechaStr = Utilities.formatDate(f, Session.getScriptTimeZone(), "dd/MM/yyyy"); fechaISO = Utilities.formatDate(f, Session.getScriptTimeZone(), "yyyy-MM-dd"); }
      let hasEvent = (data[i].length > 7 && data[i][7]) ? true : false;
      planes.push({ id: data[i][0], tipo: data[i][2], fechaProxima: fechaStr, fechaISO: fechaISO, color: color, hasDocs: docsMap[String(data[i][0])] || false, hasCalendar: hasEvent });
    }
  }
  return planes.sort((a, b) => a.fechaISO.localeCompare(b.fechaISO));
}

// En Code.gs

function getGlobalMaintenance() {
  const planes = getCachedSheetData('PLAN_MANTENIMIENTO');
  const index = buildActivosIndex();
  const docsData = getCachedSheetData('DOCS_HISTORICO');
  
  // Mapa de documentos
  const docsMap = {};
  for(let j = 1; j < docsData.length; j++) {
    if(String(docsData[j][1]) === 'REVISION') {
      docsMap[String(docsData[j][2])] = true;
    }
  }
  
  const result = [];
  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);
  
  for(let i = 1; i < planes.length; i++) {
    const idActivo = String(planes[i][1]);
    const activo = index.byId[idActivo];
    
    if (!activo) continue;
    
    // --- CAMBIO PRINCIPAL AQUÍ ---
    const estado = planes[i][6]; // Columna G (Estado)
    
    let f = planes[i][4];
    let color = 'gris';
    let fechaStr = "-";
    let dias = 0;
    let fechaISO = "";
    
    // Arreglo de fechas (caché vs objeto)
    if (typeof f === 'string' && f.length > 5) {
      f = new Date(f);
    }

    if (f instanceof Date && !isNaN(f.getTime())) {
      fechaStr = Utilities.formatDate(f, Session.getScriptTimeZone(), "dd/MM/yyyy");
      fechaISO = Utilities.formatDate(f, Session.getScriptTimeZone(), "yyyy-MM-dd");
      
      // Solo calculamos semáforos si NO está realizada
      if (estado !== 'REALIZADA') {
        f.setHours(0, 0, 0, 0);
        dias = Math.ceil((f.getTime() - hoy.getTime()) / 86400000);
        
        if (dias < 0) color = 'rojo';
        else if (dias <= 30) color = 'amarillo';
        else color = 'verde';
      } else {
        // Si está realizada, le ponemos color azul (histórico)
        color = 'azul';
        dias = 99999; // Para que salgan al final al ordenar
      }
    }
    
    let hasCalendar = (planes[i].length > 7 && planes[i][7]) ? true : false;
    
    result.push({
      id: planes[i][0],
      idActivo: idActivo,
      activo: activo.nombre,
      edificio: activo.edificio,
      campusNombre: activo.campus,
      campusId: activo.idCampus,
      tipo: planes[i][2],
      fecha: fechaStr,
      fechaISO: fechaISO,
      color: color, // rojo, amarillo, verde o AZUL
      dias: dias,
      edificioId: activo.idEdificio,
      hasDocs: docsMap[String(planes[i][0])] || false,
      hasCalendar: hasCalendar,
      estadoReal: estado // Guardamos el estado real para usarlo en el frontend
    });
  }
  
  // Ordenar: primero las urgentes (negativo), al final las históricas (muy positivo)
  return result.sort((a, b) => a.dias - b.dias);
}

// En Code.gs

function crearRevision(d) {
  verificarPermiso(['WRITE']);
  const ss = getDB(); 
  const sheet = ss.getSheetByName('PLAN_MANTENIMIENTO');
  
  try {
    let fechaActual = textoAFecha(d.fechaProx); 
    if (!fechaActual) return { success: false, error: "Fecha inválida" };
    
    var esRepetitiva = (String(d.esRecursiva) === "true"); 
    var frecuencia = parseInt(d.diasFreq) || 0; 
    var syncCal = (String(d.syncCalendar) === "true");
    
    const infoExtra = syncCal ? getInfoParaCalendar(d.idActivo) : {};
    let eventId = null; 
    if (syncCal) { 
        eventId = gestionarEventoCalendario('CREAR', { ...infoExtra, tipo: d.tipo, fecha: d.fechaProx }); 
    }
    
    const newId = Utilities.getUuid();

    // 1. Crear la revisión "padre"
    sheet.appendRow([newId, d.idActivo, d.tipo, "", new Date(fechaActual), frecuencia, "ACTIVO", eventId]);
    
    // 2. Lógica de repetición BLINDADA (Automática)
    if (esRepetitiva && frecuencia > 0) {
      // --- SEGURIDAD ---
      const MAX_REVISIONES = 12;
      const HORIZONTE_ANIOS = 10;
      
      const hoy = new Date();
      const fechaTope = new Date(hoy.getFullYear() + HORIZONTE_ANIOS, hoy.getMonth(), hoy.getDate());
      
      let fechaSiguiente = new Date(fechaActual);
      let contador = 0;

      while (contador < MAX_REVISIONES) { 
        // Sumar frecuencia
        fechaSiguiente.setDate(fechaSiguiente.getDate() + frecuencia);
        
        // Freno de emergencia por fecha
        if (fechaSiguiente > fechaTope) break;
        
        // Crear evento calendario si corresponde
        let eventIdLoop = null; 
        if (syncCal) { 
            let fStr = Utilities.formatDate(fechaSiguiente, Session.getScriptTimeZone(), "yyyy-MM-dd"); 
            eventIdLoop = gestionarEventoCalendario('CREAR', { ...infoExtra, tipo: d.tipo, fecha: fStr }); 
        }
        
        // Insertar fila
        sheet.appendRow([Utilities.getUuid(), d.idActivo, d.tipo, "", new Date(fechaSiguiente), frecuencia, "ACTIVO", eventIdLoop]); 
        contador++;
      }
      registrarLog("CREAR REVISION", `Activo: ${d.idActivo} | Tipo: ${d.tipo} (+${contador} futuras)`);
    } else {
      registrarLog("CREAR REVISION", `Activo: ${d.idActivo} | Tipo: ${d.tipo}`);
    }
    
    invalidateCache('PLAN_MANTENIMIENTO');
    return { success: true, newId: newId }; 
    
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
    registrarLog("EDITAR REVISION", "ID Revisión: " + d.idPlan + " | Nueva Fecha: " + d.fechaProx);
    return { success: true }; 
}

function completarRevision(id) {
  verificarPermiso(['WRITE']);
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const sheet = ss.getSheetByName('PLAN_MANTENIMIENTO');
  const data = sheet.getDataRange().getValues();
  
  for(let i=1; i<data.length; i++){
    if(String(data[i][0]) === String(id)) {
      // La columna 7 (índice 6) es el ESTADO. Lo cambiamos de 'ACTIVO' a 'REALIZADA'
      sheet.getRange(i+1, 7).setValue("REALIZADA"); 
      
      // Opcional: Si tenía evento de calendario, se podría borrar o actualizar, 
      // pero por ahora lo dejamos así para mantener el histórico.
      registrarLog("COMPLETAR REVISION", "Revisión finalizada ID: " + id);
      return { success: true };
    }
  }
  
  return { success: false, error: "Revisión no encontrada" };
}

function eliminarRevision(idPlan) { 
  verificarPermiso(['DELETE']);
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const sheet = ss.getSheetByName('PLAN_MANTENIMIENTO'); const data = sheet.getDataRange().getValues(); 
  for(let i=1; i<data.length; i++){ 
    if(String(data[i][0]) === String(idPlan)) { 
       let currentEventId = (data[i].length > 7) ? data[i][7] : null; if (currentEventId) { gestionarEventoCalendario('BORRAR', {}, currentEventId); }
       sheet.deleteRow(i+1); 
       registrarLog("ELIMINAR REVISION", "ID Eliminado: " + idPlan);
       return { success: true }; 
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
  const contratos = getSheetData('CONTRATOS'); 
  const activos = getSheetData('ACTIVOS'); 
  const edificios = getSheetData('EDIFICIOS'); 
  const campus = getSheetData('CAMPUS'); 

  // --- 1. MAPEO ROBUSTO (Todo a String y sin espacios) ---
  const mapCampus = {}; 
  campus.slice(1).forEach(r => {
    // Clave: ID Campus limpio
    mapCampus[String(r[0]).trim()] = r[1];
  }); 

  const mapEdificios = {}; 
  edificios.slice(1).forEach(r => {
    mapEdificios[String(r[0]).trim()] = { 
      nombre: r[2], 
      idCampus: String(r[1]).trim() 
    }; 
  }); 
  
  const mapActivos = {}; 
  activos.slice(1).forEach(r => { 
    const idEdif = String(r[1]).trim();
    const edificioInfo = mapEdificios[idEdif] || {}; 
    
    mapActivos[String(r[0]).trim()] = { 
      nombre: r[3], 
      idEdif: idEdif, 
      idCampus: edificioInfo.idCampus || null 
    }; 
  });
  
  const result = []; 
  const hoy = new Date();
  
  // --- 2. PROCESAR CONTRATOS ---
  for(let i=1; i<contratos.length; i++) {
    const r = contratos[i]; 
    const idEntidad = String(r[2] || "").trim(); // ID de la entidad (limpio)
    const tipoEntidad = String(r[1]).trim();     // ACTIVO, EDIFICIO, CAMPUS
    
    let nombreEntidad = "Sin Asignar"; // Texto por defecto si no hay enlace
    let edificioId = null; 
    let campusId = null;
    let campusNombre = "-"; 
    
    // LÓGICA DE BÚSQUEDA
    // A. Es un ACTIVO
    if (tipoEntidad === 'ACTIVO' && mapActivos[idEntidad]) { 
      const info = mapActivos[idEntidad]; 
      const nombreEdif = mapEdificios[info.idEdif] ? mapEdificios[info.idEdif].nombre : 'Sin Edif.';
      nombreEntidad = `${info.nombre} (${nombreEdif})`; 
      edificioId = info.idEdif; 
      campusId = info.idCampus; 
    } 
    // B. Es un EDIFICIO
    else if (tipoEntidad === 'EDIFICIO' && mapEdificios[idEntidad]) { 
      const info = mapEdificios[idEntidad]; 
      nombreEntidad = info.nombre; 
      edificioId = idEntidad; 
      campusId = info.idCampus; 
    }
    // C. Es un CAMPUS
    else if (tipoEntidad === 'CAMPUS' && mapCampus[idEntidad]) {
       campusId = idEntidad;
       nombreEntidad = "General Campus"; // O el nombre del campus si prefieres
    }

    // RESOLVER NOMBRE DEL CAMPUS
    if (campusId && mapCampus[campusId]) {
        campusNombre = mapCampus[campusId];
    } else if (campusId) {
        // Si tenemos ID pero no nombre, es que el campus se borró o el ID está corrupto
        campusNombre = "Desconocido"; 
    }

    // ESTADOS Y FECHAS
    const fFin = (r[6] instanceof Date) ? r[6] : null; 
    let estadoDB = r[7] || 'ACTIVO'; 
    let estadoCalc = 'VIGENTE'; 
    let color = 'verde';
    
    if (estadoDB === 'INACTIVO') { estadoCalc = 'INACTIVO'; color = 'gris'; } 
    else if (fFin) { 
      const diff = Math.ceil((fFin.getTime() - hoy.getTime()) / 86400000); 
      if (diff < 0) { estadoCalc = 'CADUCADO'; color = 'rojo'; } 
      else if (diff <= 90) { estadoCalc = 'PRÓXIMO'; color = 'amarillo'; } 
    } else { estadoCalc = 'SIN FECHA'; color = 'gris'; }
    
    result.push({ 
      id: r[0], 
      nombreEntidad: nombreEntidad, 
      proveedor: r[3], 
      ref: r[4], 
      inicio: r[5] ? fechaATexto(r[5]) : "-", 
      fin: fFin ? fechaATexto(fFin) : "-", 
      estado: estadoCalc, 
      color: color, 
      estadoDB: estadoDB, 
      edificioId: edificioId, 
      campusId: campusId,
      campusNombre: campusNombre 
    });
  }
  
  return result.sort((a, b) => { if (a.fin === "-") return 1; if (b.fin === "-") return -1; return a.fin.localeCompare(b.fin); });
}

function crearContrato(d) {
  verificarPermiso(['WRITE']);
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  ss.getSheetByName('CONTRATOS').appendRow([Utilities.getUuid(), d.tipoEntidad, d.idEntidad, d.proveedor, d.ref, textoAFecha(d.fechaIni), textoAFecha(d.fechaFin), d.estado]); registrarLog("CREAR CONTRATO", "Proveedor: " + d.proveedor + " | Ref: " + d.ref);
  invalidateCache('CONTRATOS');
  return { success: true };
}

function updateContrato(datos) {
  verificarPermiso(['WRITE']);
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const sheet = ss.getSheetByName('CONTRATOS');
  const data = sheet.getDataRange().getValues();

  for(let i=1; i<data.length; i++){
    if(String(data[i][0]) === String(datos.id)) {
      
      sheet.getRange(i+1, 2).setValue(datos.tipoEntidad); // Columna B (Tipo)
      sheet.getRange(i+1, 3).setValue(datos.idEntidad);   // Columna C (ID)

      sheet.getRange(i+1, 4).setValue(datos.proveedor);
      sheet.getRange(i+1, 5).setValue(datos.ref);
      sheet.getRange(i+1, 6).setValue(textoAFecha(datos.fechaIni));
      sheet.getRange(i+1, 7).setValue(textoAFecha(datos.fechaFin));
      sheet.getRange(i+1, 8).setValue(datos.estado);
      
      registrarLog("EDITAR CONTRATO", "Proveedor: " + datos.proveedor + " | ID: " + datos.id);
      return { success: true };
    }
  }
  return { success: false, error: "Contrato no encontrado" };
}

function eliminarContrato(id) {
  verificarPermiso(['DELETE']);
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const sheet = ss.getSheetByName('CONTRATOS');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      registrarLog("ELIMINAR_CONTRATO", "ID Contrato eliminado: " + id);;
      return { success: true };
    }
  }
  return { success: false, error: "Contrato no encontrado" };
}

// ==========================================
// 9. DASHBOARD Y CRUD GENERAL (V5)
// ==========================================

function getDatosDashboard(idCampus) { 
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); 
  const hoy = new Date(); hoy.setHours(0,0,0,0);
  
  // 1. OBTENER DATOS MASIVOS
  const dataMant = getSheetData('PLAN_MANTENIMIENTO');
  const activos = getSheetData('ACTIVOS');
  const edificios = getSheetData('EDIFICIOS');
  const campus = getSheetData('CAMPUS');
  const dataInc = getSheetData('INCIDENCIAS');
  const dataCont = getSheetData('CONTRATOS');

  // 2. PREPARAR FILTROS (Sets para búsqueda rápida)
  let validBuildingIds = new Set();
  let validAssetIds = new Set();
  
  // Si hay filtro de campus, identificamos qué edificios y activos son válidos
  if (idCampus) {
    // Filtrar edificios del campus
    for(let i=1; i<edificios.length; i++) {
      if (String(edificios[i][1]) === String(idCampus)) {
        validBuildingIds.add(String(edificios[i][0]));
      }
    }
    // Filtrar activos de esos edificios
    for(let i=1; i<activos.length; i++) {
        if (validBuildingIds.has(String(activos[i][1]))) {
            validAssetIds.add(String(activos[i][0]));
        }
    }
  }

  // Mapa de nombres de activos para el calendario
  const mapActivos = {}; 
  activos.slice(1).forEach(r => mapActivos[r[0]] = r[3]);
  
  // 3. PROCESAR MANTENIMIENTOS
  let revPend = 0, revVenc = 0, revOk = 0; const calendarEvents = [];
  const mesesNombres = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];
  const countsMap = {}; const chartLabels = []; const chartData = [];
  let dIter = new Date(hoy.getFullYear(), hoy.getMonth(), 1); 
  for (let k = 0; k < 6; k++) { let key = mesesNombres[dIter.getMonth()] + " " + dIter.getFullYear(); countsMap[key] = 0; chartLabels.push(key); dIter.setMonth(dIter.getMonth() + 1); }
  
  for(let i=1; i<dataMant.length; i++) { 
    // FILTROS
    if (dataMant[i][6] === 'REALIZADA') continue;
    if (idCampus && !validAssetIds.has(String(dataMant[i][1]))) continue; // Si es un activo de otro campus, saltar

    const f = dataMant[i][4]; 
    if(f instanceof Date) { 
      f.setHours(0,0,0,0); const diff = Math.ceil((f.getTime() - hoy.getTime()) / 86400000); 
      let status = 'verde';
      if(diff < 0) { revVenc++; status = 'rojo'; } else if(diff <= 30) { revPend++; status = 'amarillo'; } else { revOk++; status = 'verde'; }
      let key = mesesNombres[f.getMonth()] + " " + f.getFullYear(); if (countsMap.hasOwnProperty(key)) { countsMap[key]++; }
      let nombreActivo = mapActivos[dataMant[i][1]] || 'Activo';
      let colorEvento = (status === 'rojo') ? '#dc3545' : (status === 'amarillo' ? '#ffc107' : '#198754');
      let textColor = (status === 'amarillo') ? '#000' : '#fff';
      calendarEvents.push({
        id: dataMant[i][0], title: `${dataMant[i][2]} - ${nombreActivo}`, start: Utilities.formatDate(f, Session.getScriptTimeZone(), "yyyy-MM-dd"),
        backgroundColor: colorEvento, borderColor: colorEvento, textColor: textColor,
        extendedProps: { id: dataMant[i][0], tipo: dataMant[i][2], fechaISO: Utilities.formatDate(f, Session.getScriptTimeZone(), "yyyy-MM-dd"), idActivo: dataMant[i][1], hasCalendar: (dataMant[i].length > 7 && dataMant[i][7]) ? true : false }
      });
    } 
  } 
  chartLabels.forEach(lbl => { chartData.push(countsMap[lbl]); });

  // 4. PROCESAR OTROS CONTADORES CON FILTROS
  
  // Contratos
  let contCount = 0;
  if (!idCampus) {
      contCount = (dataCont.length > 1) ? dataCont.length - 1 : 0;
  } else {
      for(let i=1; i<dataCont.length; i++) {
        // Asumiendo contrato vinculado a Edificio (col 2) o Campus (col 1 - check indices logic later if needed, assuming simple count for now or based on hierarchy)
        // Revisando estructura contratos: [id, idCampus, idEdificio, ...]
        // Si tiene idEdificio y está en validBuildingIds -> cuenta
        // Si NO tiene idEdificio pero tiene idCampus y coincide -> cuenta
        let cCampusId = String(dataCont[i][1]);
        let cEdifId = String(dataCont[i][2]);
        if (cEdifId && validBuildingIds.has(cEdifId))  { contCount++; continue; }
        if (!cEdifId && cCampusId === String(idCampus)) { contCount++; }
      }
  }

  // Incidencias
  let incCount = 0; 
  for(let i=1; i<dataInc.length; i++) { 
    if(dataInc[i][6] !== 'RESUELTA') {
        if (!idCampus) {
            incCount++;
        } else {
            // Filtrar por activo o edificio
            let tipoRel = dataInc[i][1]; // ACTIVO o EDIFICIO
            let idRel = String(dataInc[i][2]);
            if (tipoRel === 'ACTIVO' && validAssetIds.has(idRel)) incCount++;
            else if (tipoRel === 'EDIFICIO' && validBuildingIds.has(idRel)) incCount++;
        }
    } 
  }

  // Totales Entidades
  let cAct = 0, cEdif = 0, cCampus = 0;
  
  if (!idCampus) {
      cAct = (activos.length > 1) ? activos.length - 1 : 0; 
      cEdif = (getSheetData('EDIFICIOS').length - 1) || 0;
      cCampus = (getSheetData('CAMPUS').length - 1) || 0;
  } else {
      cAct = validAssetIds.size;
      cEdif = validBuildingIds.size;
      cCampus = 1;
  }

  return { activos: cAct, edificios: cEdif, pendientes: revPend, vencidas: revVenc, ok: revOk, contratos: contCount, incidencias: incCount, campus: cCampus, chartLabels: chartLabels, chartData: chartData, calendarEvents: calendarEvents }; 
}

function crearCampus(d) {
  verificarPermiso(['WRITE']);
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); 
  const fId = crearCarpeta(d.nombre, getRootFolderId()); ss.getSheetByName('CAMPUS').appendRow([Utilities.getUuid(), d.nombre, d.provincia, d.direccion, fId]); 
  registrarLog("CREAR CAMPUS", "Nombre: " + d.nombre);
  invalidateCache('CAMPUS');    
  invalidateCache('EDIFICIOS');
  invalidateCache('ACTIVOS');
  return { success: true };
}

function crearEdificio(d) { 
  verificarPermiso(['WRITE']); 
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); 
  const cData = getSheetData('CAMPUS'); 
  let pId; 
  for(let i=1; i<cData.length; i++) if(String(cData[i][0])==String(d.idCampus)) pId=cData[i][4]; 
  
  const fId = crearCarpeta(d.nombre, pId); 
  const aId = crearCarpeta("Activos", fId); 
  
  // Añadimos d.lat y d.lng al final
  ss.getSheetByName('EDIFICIOS').appendRow([Utilities.getUuid(), d.idCampus, d.nombre, d.contacto, fId, aId, d.lat, d.lng]); 
  
  registrarLog("CREAR EDIFICIO", "Nombre: " + d.nombre); // Auditoría
  invalidateCache('EDIFICIOS');
  invalidateCache('ACTIVOS');
  return {success:true}; 
}

function crearActivo(d) {
  verificarPermiso(['WRITE']);
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const eData = getSheetData('EDIFICIOS');
  let pId;
  for (let i = 1; i < eData.length; i++)
    if (String(eData[i][0]) == String(d.idEdificio)) pId = eData[i][5];
  const fId = crearCarpeta(d.nombre, pId);
  const id = Utilities.getUuid();
  const cats = getSheetData('CAT_INSTALACIONES');
  let nombreTipo = d.tipo;
  for (let i = 1; i < cats.length; i++) {
    if (String(cats[i][0]) === String(d.tipo)) {
      nombreTipo = cats[i][1]; break;
    }
  }
  ss.getSheetByName('ACTIVOS').appendRow([id, d.idEdificio, nombreTipo, d.nombre, d.marca, new Date(), fId]);
  invalidateCache('ACTIVOS');
  return { success: true };
}

function getCatalogoInstalaciones() { return getSheetData('CAT_INSTALACIONES').slice(1).map(r=>({id:r[0], nombre:r[1], dias:r[3]})); }

function getTableData(tipo) { 
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); 
  if (tipo === 'CAMPUS') { 
    const data = getSheetData('CAMPUS'); 
    return data.slice(1).map(r => ({ 
      id: r[0], 
      nombre: r[1], 
      provincia: r[2], 
      direccion: r[3] })); 
  } 
  if (tipo === 'EDIFICIOS') { 
    const data = getSheetData('EDIFICIOS'); 
    const dataC = getSheetData('CAMPUS'); 
    const mapCampus = {}; 
    dataC.slice(1).forEach(r => mapCampus[r[0]] = r[1]); 
    // Ahora leemos también col 6 (LAT) y 7 (LNG) -> índices del array
    return data.slice(1).map(r => ({ 
      id: r[0], 
      campus: mapCampus[r[1]] || '-', 
      nombre: r[2], 
      contacto: r[3],
      lat: r[6], // Nueva columna
      lng: r[7]  // Nueva columna
    })); 
  } 
  return []; 
}

/**
 * buscarGlobal - OPTIMIZADO Y ROBUSTO
 */
function buscarGlobal(texto) {
  if (!texto || texto.length < 3) return [];
  
  try {
    const index = buildActivosIndex();
    const textoLower = texto.toLowerCase();
    const resultados = [];
    
    // Búsqueda en activos
    for (let i = 0; i < index.searchable.length && resultados.length < 10; i++) {
      const item = index.searchable[i];
      if (item.text.includes(textoLower)) {
        const activo = index.byId[item.id];
        if (activo) {
          resultados.push({
            id: activo.id,
            tipo: 'ACTIVO',
            texto: activo.nombre,
            subtexto: activo.tipo + (activo.marca ? " - " + activo.marca : "")
          });
        }
      }
    }
    
    // Búsqueda en edificios (si no hay suficientes resultados)
    if (resultados.length < 10) {
      const edificios = getCachedSheetData('EDIFICIOS');
      for (let i = 1; i < edificios.length && resultados.length < 10; i++) {
        const nombre = String(edificios[i][2]).toLowerCase();
        if (nombre.includes(textoLower)) {
          resultados.push({
            id: edificios[i][0],
            tipo: 'EDIFICIO',
            texto: edificios[i][2],
            subtexto: 'Edificio'
          });
        }
      }
    }
    
    return resultados;
    
  } catch(e) {
    Logger.log('Error en buscarGlobal: ' + e.toString());
    return [];
  }
}

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
// 15. SISTEMA DE FEEDBACK
// ==========================================

// 1. Enviar Feedback + Notificación Email
function enviarFeedback(datos) {
  try {
    const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
    let sheet = ss.getSheetByName('FEEDBACK');
    if (!sheet) { 
      sheet = ss.insertSheet('FEEDBACK'); 
      sheet.appendRow(['ID', 'FECHA', 'USUARIO', 'TIPO', 'MENSAJE', 'ESTADO']); 
    }
    
    const usuario = getMyRole().email || "Anónimo"; 
    const fecha = new Date();
    
    // Guardar en Excel
    sheet.appendRow([Utilities.getUuid(), fecha, usuario, datos.tipo, datos.mensaje, 'NUEVO']);
    
    // --- NUEVO: ENVIAR EMAIL AL ADMIN ---
    // (Asegúrate de cambiar este email por el tuyo real o cogerlo de la config)
    const emailAdmin = "jcsuarez@unav.es"; // O pon "tu_email@dominio.com"
    const asunto = `[GMAO] Nuevo Feedback: ${datos.tipo}`;
    const cuerpo = `
      Has recibido una nueva sugerencia en la App:
      
      Tipo: ${datos.tipo}
      Usuario: ${usuario}
      Mensaje: ${datos.mensaje}
      
      Entra en la App para gestionarla.
    `;
    
    MailApp.sendEmail(emailAdmin, asunto, cuerpo);
    // ------------------------------------

    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

// 2. Leer Feedback (Solo Admin)
function getFeedbackList() {
  verificarPermiso(['ADMIN_ONLY']);
  const data = getSheetData('FEEDBACK');
  // Devolvemos objetos limpios invertidos (lo más nuevo primero)
  return data.slice(1).reverse().map(r => ({
    id: r[0],
    fecha: r[1] ? fechaATexto(r[1]) : "-",
    usuario: r[2],
    tipo: r[3],
    mensaje: r[4],
    estado: r[5]
  }));
}

// 3. Actualizar Estado (Marcar como leído/borrar)
function updateFeedbackStatus(id, nuevoEstado) {
  verificarPermiso(['ADMIN_ONLY']);
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const sheet = ss.getSheetByName('FEEDBACK');
  const data = sheet.getDataRange().getValues();
  
  for(let i=1; i<data.length; i++){
    if(String(data[i][0]) === String(id)) {
      if(nuevoEstado === 'DELETE') {
        sheet.deleteRow(i+1);
      } else {
        sheet.getRange(i+1, 6).setValue(nuevoEstado); // Columna 6 es ESTADO
      }
      return { success: true };
    }
  }
  return { success: false, error: "No encontrado" };
}

// ==========================================
// 16. EDICIÓN Y BORRADO (CAMPUS Y EDIFICIOS)
// ==========================================
function updateCampus(d) {
  verificarPermiso(['WRITE']);
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const sheet = ss.getSheetByName('CAMPUS'); const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++){
    if(String(data[i][0]) === String(d.id)) {
      sheet.getRange(i+1, 2).setValue(d.nombre); sheet.getRange(i+1, 3).setValue(d.provincia); sheet.getRange(i+1, 4).setValue(d.direccion); 
      registrarLog("EDITAR CAMPUS", "Nombre: " + d.nombre + " | ID: " + d.id);
      return { success: true };
    }
  }
  
  return { success: false, error: "No encontrado" };
}

function eliminarCampus(id) {
  verificarPermiso(['DELETE']);
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const sheet = ss.getSheetByName('CAMPUS'); const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++){ if(String(data[i][0]) === String(id)) { sheet.deleteRow(i+1); 
  registrarLog("ELIMINAR CAMPUS", "ID Campus eliminado: " + id);  
  return { success: true }; } }
  return { success: false, error: "No encontrado" };
}

function updateEdificio(d) {
  verificarPermiso(['WRITE']);
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); 
  const sheet = ss.getSheetByName('EDIFICIOS'); 
  const data = sheet.getDataRange().getValues();
  
  for(let i=1; i<data.length; i++){
    if(String(data[i][0]) === String(d.id)) {
      sheet.getRange(i+1, 2).setValue(d.idCampus); 
      sheet.getRange(i+1, 3).setValue(d.nombre); 
      sheet.getRange(i+1, 4).setValue(d.contacto); 
      // Guardamos Lat/Lng en columnas 7 y 8
      sheet.getRange(i+1, 7).setValue(d.lat); 
      sheet.getRange(i+1, 8).setValue(d.lng);

      registrarLog("EDITAR EDIFICIO", "Nombre: " + d.nombre + " | ID: " + d.id);
      return { success: true };
    }
  }
  return { success: false, error: "No encontrado" };
}

function eliminarEdificio(id) {
  verificarPermiso(['DELETE']);
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const sheet = ss.getSheetByName('EDIFICIOS'); const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++){ if(String(data[i][0]) === String(id)) { sheet.deleteRow(i+1); 
  registrarLog("ELIMINAR_EDIFICIO", "ID Edificio eliminado: " + id);
  return { success: true }; } }
  return { success: false, error: "No encontrado" };
}

// ==========================================
// 17. CARGA MASIVA DE ACTIVOS (VERSIÓN CON CARPETAS PROPIAS)
// ==========================================
function procesarCargaMasiva(filas) {
  verificarPermiso(['WRITE']);
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const sheetActivos = ss.getSheetByName('ACTIVOS');
  
  // 1. Cargar mapas de IDs para búsqueda rápida
  const campusData = getSheetData('CAMPUS');
  const edifData = getSheetData('EDIFICIOS');
  
  const mapaCampus = {}; // Nombre -> ID
  campusData.slice(1).forEach(r => mapaCampus[String(r[1]).trim().toLowerCase()] = r[0]);
  
  const mapaEdificios = {}; // NombreEdificio + IDCampus -> ID
  edifData.slice(1).forEach(r => {
    const key = String(r[2]).trim().toLowerCase() + "_" + String(r[1]);
    mapaEdificios[key] = r[0]; // ID del edificio
  });

  const nuevosActivos = [];
  const errores = [];
  const folderCache = {}; // Cacheamos la ID de la carpeta PADRE del edificio (no la del activo)

  // 2. Procesar fila a fila
  // Importante: Usamos un bucle for tradicional para manejar mejor los tiempos si hay muchos
  for (let index = 0; index < filas.length; index++) {
    const fila = filas[index];
    if (fila.length < 4) continue; // Saltar filas incompletas

    const nombreCampus = String(fila[0]).trim();
    const nombreEdif = String(fila[1]).trim();
    const tipo = String(fila[2]).trim();
    const nombreActivo = String(fila[3]).trim();
    const marca = fila[4] ? String(fila[4]).trim() : "-";

    // 3. Resolver IDs
    const idCampus = mapaCampus[nombreCampus.toLowerCase()];
    if (!idCampus) {
      errores.push(`Fila ${index + 1}: Campus '${nombreCampus}' no existe.`);
      continue;
    }

    const idEdificio = mapaEdificios[nombreEdif.toLowerCase() + "_" + idCampus];
    if (!idEdificio) {
      errores.push(`Fila ${index + 1}: Edificio '${nombreEdif}' no encontrado en ese Campus.`);
      continue;
    }

    // 4. Localizar la carpeta "Madre" del Edificio (donde se guardan los activos)
    let idCarpetaPadreEdificio = folderCache[idEdificio];
    if (!idCarpetaPadreEdificio) {
      for(let k=1; k<edifData.length; k++) {
        if(String(edifData[k][0]) === String(idEdificio)) {
          // La columna 6 (índice 5) en EDIFICIOS es la carpeta "aId" (Assets Folder)
          idCarpetaPadreEdificio = edifData[k][5]; 
          folderCache[idEdificio] = idCarpetaPadreEdificio;
          break;
        }
      }
    }

    // 5. CREAR CARPETA INDIVIDUAL PARA EL ACTIVO
    // Esto hace que cada activo tenga su propio espacio, igual que al crear manual.
    let idCarpetaActivo = "";
    if (idCarpetaPadreEdificio) {
      try {
        // Usamos tu función helper existente 'crearCarpeta'
        // DriveApp puede ser lento, así que esto añade tiempo de proceso.
        idCarpetaActivo = crearCarpeta(nombreActivo, idCarpetaPadreEdificio);
      } catch (e) {
        // Si falla Drive (raro), guardamos el error pero creamos el activo sin carpeta para no perder datos
        errores.push(`Fila ${index + 1}: Activo creado pero falló al crear carpeta Drive (${e.message}).`);
      }
    } else {
       errores.push(`Fila ${index + 1}: No se encontró carpeta del edificio. Activo creado sin carpeta.`);
    }

    // 6. Añadir a la lista
    nuevosActivos.push([
      Utilities.getUuid(),
      idEdificio,
      tipo,
      nombreActivo,
      marca,
      new Date(),
      idCarpetaActivo // ¡Aquí va la ID nueva específica para este activo!
    ]);
  }

  // 7. Volcar a la hoja de cálculo
  if (nuevosActivos.length > 0) {
    sheetActivos.getRange(
      sheetActivos.getLastRow() + 1, 
      1, 
      nuevosActivos.length, 
      nuevosActivos[0].length
    ).setValues(nuevosActivos);
  }
  registrarLog("IMPORTACIÓN MASIVA", "Se importaron " + nuevosActivos.length + " activos.");
  return { 
    success: true, 
    procesados: nuevosActivos.length, 
    errores: errores 
  };
}

// ==========================================
// 18. SISTEMA DE AUDITORÍA (LOGS)
// ==========================================

// Función principal para registrar acciones
function registrarLog(accion, detalles) {
  try {
    const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
    let sheet = ss.getSheetByName('LOGS');
    
    // Si no existe la hoja, la crea y la oculta para que no moleste
    if (!sheet) {
      sheet = ss.insertSheet('LOGS');
      sheet.appendRow(['FECHA', 'USUARIO', 'ACCIÓN', 'DETALLES']);
      sheet.setColumnWidth(1, 150); // Fecha
      sheet.setColumnWidth(2, 200); // Usuario
      sheet.setColumnWidth(3, 150); // Acción
      sheet.setColumnWidth(4, 400); // Detalles
      sheet.hideSheet(); // Ocultar visualmente en el Spreadsheet
    }
    
    const usuario = Session.getActiveUser().getEmail();
    const fecha = new Date();
    
    // Guardamos el log
    sheet.appendRow([fecha, usuario, accion, detalles]);
    
  } catch (e) {
    console.error("Error al registrar log: " + e.toString());
  }
}

// Función para que el Admin vea los logs en la App
function getLogsAuditoria() {
  verificarPermiso(['ADMIN_ONLY']); // ¡Solo Admins!
  const data = getSheetData('LOGS');
  // Devolvemos los últimos 100 registros (invertimos el orden para ver los nuevos primero)
  const logs = data.slice(1).reverse().slice(0, 100).map(r => {
    let fechaStr = "-";
    if (r[0] instanceof Date) {
      fechaStr = Utilities.formatDate(r[0], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    }
    return { fecha: fechaStr, usuario: r[1], accion: r[2], detalles: r[3] };
  });
  return logs;
}

// ==========================================
// 19. SISTEMA DE NOVEDADES (CHANGELOG)
// ==========================================
function getNovedadesApp() {
  // Permitimos que cualquiera (Consultas, Tecnicos, Admins) vea esto
  const data = getSheetData('NOVEDADES');
  
  // Devolvemos las filas (excepto cabecera), invertidas para ver las nuevas primero
  return data.slice(1).reverse().map(r => ({
    fecha: r[0] ? fechaATexto(r[0]) : "",
    version: r[1],
    tipo: r[2],
    titulo: r[3],
    descripcion: r[4]
  }));
}
function crearNovedad(d) {
  verificarPermiso(['ADMIN_ONLY']); // Solo Admins
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  let sheet = ss.getSheetByName('NOVEDADES');
  
  if (!sheet) {
    sheet = ss.insertSheet('NOVEDADES');
    sheet.appendRow(['FECHA', 'VERSION', 'TIPO', 'TITULO', 'DESCRIPCION']);
  }

  // TRUCO: Concatenamos "'" al principio de la versión
  // Esto obliga a Google Sheets a tratar "01.01" como texto y no como 1 de Enero.
  const versionTexto = "'" + d.version; 

  sheet.appendRow([
    new Date(),     // Fecha actual
    versionTexto,   // Versión (Texto forzado)
    d.tipo, 
    d.titulo, 
    d.descripcion
  ]);

  registrarLog("NUEVA VERSIÓN", "Publicada versión " + d.version);
  return { success: true };
}

// ==========================================
// 20. GENERADOR DE INFORMES PDF
// ==========================================

function generarInformePDF(tipoReporte) {
  // Verificamos permisos (Técnicos y Admins pueden descargar)
  verificarPermiso(['READ']); 
  
  let html = "";
  const fechaHoy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  let filename = "Informe.pdf";

  // Estilos CSS inline para el PDF (Google PDF engine prefiere estilos simples)
  const css = `
    <style>
      body { font-family: sans-serif; font-size: 10pt; color: #333; }
      h1 { color: #CC0605; font-size: 16pt; margin-bottom: 5px; }
      .header { border-bottom: 2px solid #CC0605; padding-bottom: 10px; margin-bottom: 20px; }
      .meta { font-size: 9pt; color: #666; margin-bottom: 20px; }
      table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
      th { background-color: #f3f4f6; color: #CC0605; border-bottom: 1px solid #ddd; padding: 8px; text-align: left; font-size: 9pt; }
      td { border-bottom: 1px solid #eee; padding: 8px; font-size: 9pt; vertical-align: top; }
      .badge { padding: 2px 6px; border-radius: 4px; font-size: 8pt; font-weight: bold; display: inline-block; }
      .bg-verde { color: #10b981; border: 1px solid #10b981; }
      .bg-rojo { color: #ef4444; border: 1px solid #ef4444; }
      .bg-amarillo { color: #f59e0b; border: 1px solid #f59e0b; }
      .footer { position: fixed; bottom: 0; width: 100%; text-align: center; font-size: 8pt; color: #aaa; border-top: 1px solid #eee; padding-top: 10px; }
    </style>
  `;

  if (tipoReporte === 'LEGAL') {
    filename = "Auditoria_Revisiones_Legales.pdf";
    const datos = getGlobalMaintenance(); // Reutilizamos tu función existente
    // Filtramos solo las LEGALES
    const legales = datos.filter(r => r.tipo === 'Legal');

    html += `
      <html><body>${css}
      <div class="header">
        <h1>Informe de Revisiones Legales</h1>
        <div class="meta">GMAO Universidad de Navarra | Generado el: ${fechaHoy}</div>
      </div>
      <div class="meta">
        <strong>Total Registros:</strong> ${legales.length} revisiones normativas.
      </div>
      <table>
        <thead>
          <tr>
            <th width="20%">Campus / Edificio</th>
            <th width="30%">Activo</th>
            <th width="15%">Tipo</th>
            <th width="15%">Fecha Límite</th>
            <th width="20%">Estado</th>
          </tr>
        </thead>
        <tbody>
    `;

    if(legales.length === 0) {
      html += `<tr><td colspan="5" style="text-align:center; padding:20px;">No hay revisiones pendientes.</td></tr>`;
    } else {
      legales.forEach(r => {
        let estadoTxt = r.color === 'rojo' ? 'VENCIDA' : (r.color === 'amarillo' ? 'PRÓXIMA' : 'AL DÍA');
        let claseColor = 'bg-' + r.color; // bg-rojo, bg-verde...
        
        // Usamos los nombres resueltos que arreglamos antes
        let ubicacion = `<strong>${r.edificio}</strong><br><small>${r.campusNombre}</small>`;

        html += `
          <tr>
            <td>${ubicacion}</td>
            <td>${r.activo}</td>
            <td>${r.tipo}</td>
            <td>${r.fecha}</td>
            <td><span class="badge ${claseColor}">${estadoTxt}</span></td>
          </tr>
        `;
      });
    }
    html += `</tbody></table><div class="footer">Documento generado automáticamente para auditoría interna/externa.</div></body></html>`;

  } else if (tipoReporte === 'CONTRATOS') {
    filename = "Informe_Contratos.pdf";
    const datos = obtenerContratosGlobal(); // Reutilizamos tu función existente

    html += `
      <html><body>${css}
      <div class="header">
        <h1>Listado de Contratos Vigentes</h1>
        <div class="meta">GMAO Universidad de Navarra | Generado el: ${fechaHoy}</div>
      </div>
      <table>
        <thead>
          <tr>
            <th width="15%">Estado</th>
            <th width="25%">Proveedor</th>
            <th width="20%">Referencia</th>
            <th width="25%">Vigencia</th>
            <th width="15%">Activo/Edif</th>
          </tr>
        </thead>
        <tbody>
    `;

    if(datos.length === 0) {
      html += `<tr><td colspan="5">No hay contratos registrados.</td></tr>`;
    } else {
      datos.forEach(c => {
        // Formateo de fechas
        let vigencia = `${c.inicio} - ${c.fin}`;
        
        html += `
          <tr>
            <td>${c.estado}</td>
            <td><strong>${c.proveedor}</strong></td>
            <td>${c.ref}</td>
            <td>${vigencia}</td>
            <td>${c.nombreEntidad}</td>
          </tr>
        `;
      });
    }
    html += `</tbody></table><div class="footer">Documento oficial de control de contratos.</div></body></html>`;
  }

  // Conversión mágica a PDF
  const blob = Utilities.newBlob(html, MimeType.HTML).getAs(MimeType.PDF);
  blob.setName(filename);
  
  // Devolvemos el Base64 para descarga directa
  return { 
    base64: Utilities.base64Encode(blob.getBytes()), 
    filename: filename 
  };
}

// ==========================================
// 21. DETALLE DE CONTRATO PARA EDICIÓN
// ==========================================
function getContratoFullDetails(idContrato) {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const sheet = ss.getSheetByName('CONTRATOS');
  const data = sheet.getDataRange().getValues();
  
  let contrato = null;
  // Buscamos el contrato
  for(let i=1; i<data.length; i++){
    if(String(data[i][0]) === String(idContrato)) {
      contrato = {
        id: data[i][0],
        tipoEntidad: data[i][1],
        idEntidad: data[i][2],
        proveedor: data[i][3],
        ref: data[i][4],
        inicio: data[i][5] ? Utilities.formatDate(data[i][5], Session.getScriptTimeZone(), "yyyy-MM-dd") : "",
        fin: data[i][6] ? Utilities.formatDate(data[i][6], Session.getScriptTimeZone(), "yyyy-MM-dd") : "",
        estado: data[i][7]
      };
      break;
    }
  }
  
  if(!contrato) throw new Error("Contrato no encontrado");

  // AHORA RESOLVEMOS LA JERARQUÍA (Campus -> Edificio -> Activo)
  let idCampus = "";
  let idEdificio = "";
  let idActivo = "";

  if (contrato.tipoEntidad === 'CAMPUS') {
    idCampus = contrato.idEntidad;
  } 
  else if (contrato.tipoEntidad === 'EDIFICIO') {
    idEdificio = contrato.idEntidad;
    // Buscamos a qué campus pertenece este edificio
    const edifs = getSheetData('EDIFICIOS');
    for(let i=1; i<edifs.length; i++) {
       if(String(edifs[i][0]) === String(idEdificio)) {
         idCampus = edifs[i][1];
         break;
       }
    }
  } 
  else if (contrato.tipoEntidad === 'ACTIVO') {
    idActivo = contrato.idEntidad;
    // Buscamos activo -> edificio -> campus
    const activos = getSheetData('ACTIVOS');
    const edifs = getSheetData('EDIFICIOS');
    
    for(let i=1; i<activos.length; i++) {
       if(String(activos[i][0]) === String(idActivo)) {
         idEdificio = activos[i][1];
         break;
       }
    }
    if (idEdificio) {
      for(let i=1; i<edifs.length; i++) {
         if(String(edifs[i][0]) === String(idEdificio)) {
           idCampus = edifs[i][1];
           break;
         }
      }
    }
  }

  // Devolvemos todo junto
  return {
    ...contrato,
    idCampus: idCampus,
    idEdificio: idEdificio,
    idActivo: idActivo
  };
}

// ==========================================
// 22. SISTEMA DE NOTIFICACIONES AUTOMÁTICAS
// ==========================================

/**
 * FUNCIÓN PRINCIPAL - Se ejecuta automáticamente cada día
 * Configurar en: Activadores (Triggers) > Añadir activador
 * - Función: enviarNotificacionesAutomaticas
 * - Tipo: Controlado por tiempo
 * - Tipo de activador: Temporizador diario
 * - Hora: 08:00 - 09:00 (recomendado)
 */
function enviarNotificacionesAutomaticas() {
  try {
    Logger.log("=== INICIO ENVÍO NOTIFICACIONES AUTOMÁTICAS ===");
    
    // 1. Obtener usuarios con notificaciones activadas
    const usuariosConAvisos = obtenerUsuariosConAvisos();
    
    if (usuariosConAvisos.length === 0) {
      Logger.log("No hay usuarios con avisos activados. Finalizando.");
      return;
    }
    
    // 2. Detectar revisiones próximas a vencer (7 días)
    const revisionesProximas = detectarRevisionesProximas();
    
    // 3. Detectar revisiones vencidas
    const revisionesVencidas = detectarRevisionesVencidas();
    
    // 4. Detectar contratos próximos a caducar (60 días)
    const contratosProximos = detectarContratosProximos();
    
    // 5. Si hay algo que notificar, enviar emails
    if (revisionesProximas.length > 0 || revisionesVencidas.length > 0 || contratosProximos.length > 0) {
      usuariosConAvisos.forEach(usuario => {
        enviarEmailResumen(usuario, revisionesProximas, revisionesVencidas, contratosProximos);
      });
      
      Logger.log(`Emails enviados a ${usuariosConAvisos.length} usuario(s)`);
    } else {
      Logger.log("No hay alertas que notificar hoy.");
    }
    
    Logger.log("=== FIN ENVÍO NOTIFICACIONES ===");
    
  } catch (e) {
    Logger.log("ERROR en notificaciones automáticas: " + e.toString());
    // Enviar email de error al admin
    MailApp.sendEmail(
      "jcsuarez@unav.es", 
      "[GMAO] Error en Notificaciones Automáticas", 
      "Se ha producido un error al enviar las notificaciones automáticas:\n\n" + e.toString()
    );
  }
}

// Obtener usuarios que tienen avisos activados
function obtenerUsuariosConAvisos() {
  const data = getSheetData('USUARIOS');
  const usuarios = [];
  
  for(let i = 1; i < data.length; i++) {
    if (data[i][4] === 'SI') { // Columna 5 = AVISOS
      usuarios.push({
        nombre: data[i][1],
        email: data[i][2],
        rol: data[i][3]
      });
    }
  }
  
  return usuarios;
}

// Detectar revisiones que vencen en los próximos 7 días
function detectarRevisionesProximas() {
  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);
  
  const en7Dias = new Date(hoy);
  en7Dias.setDate(en7Dias.getDate() + 7);
  
  const dataMant = getSheetData('PLAN_MANTENIMIENTO');
  const activos = getSheetData('ACTIVOS');
  const edificios = getSheetData('EDIFICIOS');
  
  // Mapas para resolver nombres
  const mapActivos = {};
  activos.slice(1).forEach(r => mapActivos[r[0]] = { nombre: r[3], edificioId: r[1] });
  
  const mapEdificios = {};
  edificios.slice(1).forEach(r => mapEdificios[r[0]] = r[2]);
  
  const proximas = [];
  
  for(let i = 1; i < dataMant.length; i++) {
    // Saltar si ya está realizada
    if (dataMant[i][6] === 'REALIZADA') continue;
    
    const fechaRevision = dataMant[i][4];
    
    if (fechaRevision instanceof Date) {
      fechaRevision.setHours(0, 0, 0, 0);
      
      // Si está entre hoy y 7 días
      if (fechaRevision >= hoy && fechaRevision <= en7Dias) {
        const idActivo = dataMant[i][1];
        const activo = mapActivos[idActivo];
        
        if (activo) {
          const diasRestantes = Math.ceil((fechaRevision - hoy) / (1000 * 60 * 60 * 24));
          
          proximas.push({
            tipo: dataMant[i][2],
            activo: activo.nombre,
            edificio: mapEdificios[activo.edificioId] || 'N/A',
            fecha: Utilities.formatDate(fechaRevision, Session.getScriptTimeZone(), "dd/MM/yyyy"),
            diasRestantes: diasRestantes
          });
        }
      }
    }
  }
  
  return proximas;
}

// Detectar revisiones vencidas
function detectarRevisionesVencidas() {
  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);
  
  const dataMant = getSheetData('PLAN_MANTENIMIENTO');
  const activos = getSheetData('ACTIVOS');
  const edificios = getSheetData('EDIFICIOS');
  
  const mapActivos = {};
  activos.slice(1).forEach(r => mapActivos[r[0]] = { nombre: r[3], edificioId: r[1] });
  
  const mapEdificios = {};
  edificios.slice(1).forEach(r => mapEdificios[r[0]] = r[2]);
  
  const vencidas = [];
  
  for(let i = 1; i < dataMant.length; i++) {
    if (dataMant[i][6] === 'REALIZADA') continue;
    
    const fechaRevision = dataMant[i][4];
    
    if (fechaRevision instanceof Date) {
      fechaRevision.setHours(0, 0, 0, 0);
      
      // Si está vencida
      if (fechaRevision < hoy) {
        const idActivo = dataMant[i][1];
        const activo = mapActivos[idActivo];
        
        if (activo) {
          const diasVencida = Math.ceil((hoy - fechaRevision) / (1000 * 60 * 60 * 24));
          
          vencidas.push({
            tipo: dataMant[i][2],
            activo: activo.nombre,
            edificio: mapEdificios[activo.edificioId] || 'N/A',
            fecha: Utilities.formatDate(fechaRevision, Session.getScriptTimeZone(), "dd/MM/yyyy"),
            diasVencida: diasVencida
          });
        }
      }
    }
  }
  
  return vencidas;
}

// Detectar contratos que caducan en los próximos 60 días
function detectarContratosProximos() {
  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);
  
  const en60Dias = new Date(hoy);
  en60Dias.setDate(en60Dias.getDate() + 60);
  
  const dataContratos = getSheetData('CONTRATOS');
  const proximos = [];
  
  for(let i = 1; i < dataContratos.length; i++) {
    const estado = dataContratos[i][7]; // Estado del contrato
    
    // Solo contratos activos
    if (estado !== 'ACTIVO') continue;
    
    const fechaFin = dataContratos[i][6];
    
    if (fechaFin instanceof Date) {
      fechaFin.setHours(0, 0, 0, 0);
      
      // Si caduca en los próximos 60 días
      if (fechaFin >= hoy && fechaFin <= en60Dias) {
        const diasRestantes = Math.ceil((fechaFin - hoy) / (1000 * 60 * 60 * 24));
        
        proximos.push({
          proveedor: dataContratos[i][3],
          referencia: dataContratos[i][4],
          fechaFin: Utilities.formatDate(fechaFin, Session.getScriptTimeZone(), "dd/MM/yyyy"),
          diasRestantes: diasRestantes
        });
      }
    }
  }
  
  return proximos;
}

// Enviar email resumen a un usuario
function enviarEmailResumen(usuario, revisionesProximas, revisionesVencidas, contratosProximos) {
  const fechaHoy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
  
  // Construir HTML del email
  let htmlBody = `
    <html>
      <head>
        <style>
          body { font-family: Arial, sans-serif; color: #333; line-height: 1.6; }
          .header { background-color: #CC0605; color: white; padding: 20px; text-align: center; }
          .content { padding: 20px; }
          .section { margin-bottom: 30px; }
          .section-title { color: #CC0605; font-size: 18px; font-weight: bold; border-bottom: 2px solid #CC0605; padding-bottom: 5px; margin-bottom: 15px; }
          table { width: 100%; border-collapse: collapse; margin-top: 10px; }
          th { background-color: #f3f4f6; color: #CC0605; padding: 10px; text-align: left; font-size: 12px; }
          td { padding: 10px; border-bottom: 1px solid #eee; font-size: 13px; }
          .urgente { color: #dc3545; font-weight: bold; }
          .proximo { color: #f59e0b; font-weight: bold; }
          .footer { background-color: #f9fafb; padding: 15px; text-align: center; font-size: 12px; color: #666; margin-top: 30px; }
          .badge { padding: 3px 8px; border-radius: 4px; font-size: 11px; font-weight: bold; }
          .badge-danger { background-color: #fef2f2; color: #dc3545; }
          .badge-warning { background-color: #fffbeb; color: #f59e0b; }
        </style>
      </head>
      <body>
        <div class="header">
          <h1>🔔 Notificación Automática GMAO</h1>
          <p>Universidad de Navarra - ${fechaHoy}</p>
        </div>
        
        <div class="content">
          <p>Hola <strong>${usuario.nombre}</strong>,</p>
          <p>Este es tu resumen automático de mantenimiento y contratos:</p>
  `;
  
  // SECCIÓN: Revisiones Vencidas
  if (revisionesVencidas.length > 0) {
    htmlBody += `
      <div class="section">
        <div class="section-title">⚠️ REVISIONES VENCIDAS (${revisionesVencidas.length})</div>
        <table>
          <thead>
            <tr>
              <th>Activo</th>
              <th>Edificio</th>
              <th>Tipo</th>
              <th>Fecha Límite</th>
              <th>Días Vencida</th>
            </tr>
          </thead>
          <tbody>
    `;
    
    revisionesVencidas.forEach(r => {
      htmlBody += `
        <tr>
          <td>${r.activo}</td>
          <td>${r.edificio}</td>
          <td>${r.tipo}</td>
          <td>${r.fecha}</td>
          <td class="urgente">${r.diasVencida} días</td>
        </tr>
      `;
    });
    
    htmlBody += `
          </tbody>
        </table>
      </div>
    `;
  }
  
  // SECCIÓN: Revisiones Próximas
  if (revisionesProximas.length > 0) {
    htmlBody += `
      <div class="section">
        <div class="section-title">📅 REVISIONES PRÓXIMAS (${revisionesProximas.length})</div>
        <p style="font-size: 13px; color: #666;">Vencen en los próximos 7 días</p>
        <table>
          <thead>
            <tr>
              <th>Activo</th>
              <th>Edificio</th>
              <th>Tipo</th>
              <th>Fecha Límite</th>
              <th>Días Restantes</th>
            </tr>
          </thead>
          <tbody>
    `;
    
    revisionesProximas.forEach(r => {
      htmlBody += `
        <tr>
          <td>${r.activo}</td>
          <td>${r.edificio}</td>
          <td>${r.tipo}</td>
          <td>${r.fecha}</td>
          <td class="proximo">${r.diasRestantes} días</td>
        </tr>
      `;
    });
    
    htmlBody += `
          </tbody>
        </table>
      </div>
    `;
  }
  
  // SECCIÓN: Contratos Próximos a Caducar
  if (contratosProximos.length > 0) {
    htmlBody += `
      <div class="section">
        <div class="section-title">📄 CONTRATOS PRÓXIMOS A CADUCAR (${contratosProximos.length})</div>
        <p style="font-size: 13px; color: #666;">Caducan en los próximos 60 días</p>
        <table>
          <thead>
            <tr>
              <th>Proveedor</th>
              <th>Referencia</th>
              <th>Fecha Fin</th>
              <th>Días Restantes</th>
            </tr>
          </thead>
          <tbody>
    `;
    
    contratosProximos.forEach(c => {
      htmlBody += `
        <tr>
          <td>${c.proveedor}</td>
          <td>${c.referencia}</td>
          <td>${c.fechaFin}</td>
          <td class="proximo">${c.diasRestantes} días</td>
        </tr>
      `;
    });
    
    htmlBody += `
          </tbody>
        </table>
      </div>
    `;
  }
  
  // Footer
  htmlBody += `
          <div class="footer">
            <p><strong>GMAO - Sistema de Gestión de Mantenimiento</strong></p>
            <p>Universidad de Navarra | Servicio de Obras y Mantenimiento</p>
            <p style="font-size: 11px; color: #999;">Este es un email automático. Para gestionar tus preferencias de notificaciones, accede a la aplicación.</p>
          </div>
        </div>
      </body>
    </html>
  `;
  
  // Construir asunto dinámico
  let asunto = "[GMAO] Resumen de Mantenimiento - " + fechaHoy;
  
  if (revisionesVencidas.length > 0) {
    asunto = `⚠️ [GMAO] ${revisionesVencidas.length} Revisión(es) Vencida(s) - ${fechaHoy}`;
  } else if (revisionesProximas.length > 0) {
    asunto = `📅 [GMAO] ${revisionesProximas.length} Revisión(es) Próxima(s) - ${fechaHoy}`;
  }
  
  // Enviar email
  MailApp.sendEmail({
    to: usuario.email,
    subject: asunto,
    htmlBody: htmlBody
  });
  
  Logger.log(`Email enviado a: ${usuario.email}`);
}

// ==========================================
// 23. FUNCIONES DE PRUEBA Y CONFIGURACIÓN
// ==========================================

/**
 * Función de PRUEBA - Ejecutar manualmente para probar el sistema
 * NO configurar como trigger automático
 */
function probarNotificaciones() {
  Logger.log("=== PRUEBA DE NOTIFICACIONES ===");
  enviarNotificacionesAutomaticas();
  Logger.log("Revisa tu email y los logs para ver el resultado");
}

/**
 * Función para DESACTIVAR todas las notificaciones de un usuario
 */
function desactivarNotificacionesUsuario(email) {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const sheet = ss.getSheetByName('USUARIOS');
  const data = sheet.getDataRange().getValues();
  
  for(let i = 1; i < data.length; i++) {
    if (String(data[i][2]).toLowerCase() === email.toLowerCase()) {
      sheet.getRange(i + 1, 5).setValue('NO'); // Columna AVISOS
      Logger.log(`Notificaciones desactivadas para: ${email}`);
      return { success: true };
    }
  }
  
  return { success: false, error: "Usuario no encontrado" };
}

// ==========================================
// 24. SISTEMA DE CÓDIGOS QR PARA ACTIVOS
// ==========================================

/**
 * Generar código QR para un activo específico
 * Usa la API gratuita de Google Charts
 */
function generarQRActivo(idActivo) {
  try {
    // Obtener datos del activo
    const activo = getAssetInfo(idActivo);
    
    if (!activo) {
      return { success: false, error: "Activo no encontrado" };
    }
    
    // Obtener URL de la aplicación
    const urlApp = ScriptApp.getService().getUrl();
    
    // Crear URL con parámetro para abrir directamente el activo
    const urlQR = `${urlApp}?activo=${idActivo}`;
    
    // Generar QR usando API QRServer (muy fiable)
    const qrImageUrl = `https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=${encodeURIComponent(urlQR)}`;
    
    // Descargar imagen y convertir a Base64 para evitar bloqueos
    let qrBase64 = qrImageUrl;
    try {
      const resp = UrlFetchApp.fetch(qrImageUrl);
      if (resp.getResponseCode() === 200) {
        const blob = resp.getBlob();
        qrBase64 = "data:image/png;base64," + Utilities.base64Encode(blob.getBytes());
      }
    } catch(e) {
      console.error("Error fetching QR: " + e);
    }

    return {
      success: true,
      qrUrl: qrBase64,
      targetUrl: urlQR,
      activo: {
        id: activo.id,
        nombre: activo.nombre,
        tipo: activo.tipo,
        edificio: activo.edificio,
        campus: activo.campus
      }
    };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Generar QR para todos los activos de un edificio
 * Devuelve un array con todos los QR para imprimir
 */
function generarQRsEdificio(idEdificio) {
  try {
    const activos = getSheetData('ACTIVOS');
    const edificios = getSheetData('EDIFICIOS');
    const campus = getSheetData('CAMPUS');
    
    // Buscar nombre del edificio
    let nombreEdificio = '';
    let idCampus = '';
    for(let i = 1; i < edificios.length; i++) {
      if (String(edificios[i][0]) === String(idEdificio)) {
        nombreEdificio = edificios[i][2];
        idCampus = edificios[i][1];
        break;
      }
    }
    
    // Buscar nombre del campus
    let nombreCampus = '';
    for(let i = 1; i < campus.length; i++) {
      if (String(campus[i][0]) === String(idCampus)) {
        nombreCampus = campus[i][1];
        break;
      }
    }
    
    const urlApp = ScriptApp.getService().getUrl();
    const qrs = [];
    
    // Generar QR para cada activo del edificio
    for(let i = 1; i < activos.length; i++) {
      if (String(activos[i][1]) === String(idEdificio)) {
        const idActivo = activos[i][0];
        const urlQR = `${urlApp}?activo=${idActivo}`;
        const qrImageUrl = `https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=${encodeURIComponent(urlQR)}`;
        
        qrs.push({
          id: idActivo,
          nombre: activos[i][3],
          tipo: activos[i][2],
          marca: activos[i][4] || '-',
          qrUrl: qrImageUrl,
          targetUrl: urlQR
        });
      }
    }
    
    return {
      success: true,
      edificio: nombreEdificio,
      campus: nombreCampus,
      totalActivos: qrs.length,
      qrs: qrs
    };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Generar PDF con etiquetas QR para imprimir
 * Formato: 2 columnas, con nombre del activo y QR
 */
function generarPDFEtiquetasQR(idEdificio) {
  try {
    const data = generarQRsEdificio(idEdificio);
    
    if (!data.success) {
      return data;
    }

    // Pre-fetch de imágenes QR para convertirlas a Base64
    // Esto soluciona el problema de imágenes rotas en el PDF
    const requests = data.qrs.map(qr => ({
      url: qr.qrUrl,
      method: "GET",
      muteHttpExceptions: true
    }));

    try {
      const responses = UrlFetchApp.fetchAll(requests);
      
      data.qrs.forEach((qr, index) => {
        const response = responses[index];
        if (response.getResponseCode() === 200) {
          const blob = response.getBlob();
          const base64 = Utilities.base64Encode(blob.getBytes());
          qr.qrUrl = "data:image/png;base64," + base64;
        }
      });
    } catch (err) {
      console.error("Error fetching QR images: " + err);
      // Si falla, intentamos usar la URL original, aunque puede que no cargue en PDF
    }
    
    const fechaHoy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    
    // Construir HTML para el PDF
    let html = `
      <html>
        <head>
          <style>
            @page { margin: 1cm; size: A4; }
            body { font-family: Arial, sans-serif; }
            .header { text-align: center; margin-bottom: 20px; border-bottom: 2px solid #CC0605; padding-bottom: 10px; }
            .header h1 { color: #CC0605; margin: 0; font-size: 18pt; }
            .header p { margin: 5px 0; font-size: 10pt; color: #666; }
            .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; }
            .etiqueta { 
              border: 2px solid #CC0605; 
              padding: 15px; 
              text-align: center; 
              page-break-inside: avoid;
              border-radius: 8px;
              display: flex;
              flex-direction: column;
              align-items: center;
              justify-content: center;
            }
            .etiqueta img { width: 200px; height: 200px; margin: 10px 0; object-fit: contain; }
            .etiqueta h3 { color: #CC0605; margin: 5px 0; font-size: 14pt; }
            .etiqueta p { margin: 3px 0; font-size: 10pt; color: #666; }
            .etiqueta .tipo { font-weight: bold; color: #333; font-size: 11pt; }
            .footer { 
              position: fixed; 
              bottom: 0; 
              width: 100%; 
              text-align: center; 
              font-size: 8pt; 
              color: #999; 
              border-top: 1px solid #eee; 
              padding-top: 5px; 
            }
          </style>
        </head>
        <body>
          <div class="header">
            <h1>Códigos QR - Activos</h1>
            <p><strong>${data.edificio}</strong> | ${data.campus}</p>
            <p>Generado el: ${fechaHoy} | Total: ${data.totalActivos} activos</p>
          </div>
          
          <div class="grid">
    `;
    
    data.qrs.forEach(qr => {
      html += `
        <div class="etiqueta">
          <h3>${qr.nombre}</h3>
          <p class="tipo">${qr.tipo}</p>
          ${qr.marca !== '-' ? `<p>Marca: ${qr.marca}</p>` : ''}
          <img src="${qr.qrUrl}" alt="QR ${qr.nombre}">
          <p style="font-size: 8pt; color: #999;">Escanea para ver detalles</p>
        </div>
      `;
    });
    
    html += `
          </div>
          <div class="footer">
            GMAO - Universidad de Navarra | Sistema de Gestión de Mantenimiento
          </div>
        </body>
      </html>
    `;
    
    // Convertir a PDF
    const blob = Utilities.newBlob(html, MimeType.HTML).getAs(MimeType.PDF);
    const filename = `QR_${data.edificio.replace(/\s+/g, '_')}_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd")}.pdf`;
    blob.setName(filename);
    
    return {
      success: true,
      base64: Utilities.base64Encode(blob.getBytes()),
      filename: filename,
      totalEtiquetas: data.totalActivos
    };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Obtener información del activo cuando se escanea el QR
 * Esta función se llama cuando alguien accede a la URL del QR
 */
function getActivoByQR(idActivo) {
  try {
    const activo = getAssetInfo(idActivo);
    
    if (!activo) {
      return { success: false, error: "Activo no encontrado" };
    }
    
    // Obtener últimas revisiones
    const dataMant = getSheetData('PLAN_MANTENIMIENTO');
    const revisiones = [];
    
    for(let i = 1; i < dataMant.length; i++) {
      if (String(dataMant[i][1]) === String(idActivo)) {
        revisiones.push({
          tipo: dataMant[i][2],
          fecha: dataMant[i][4] ? Utilities.formatDate(dataMant[i][4], Session.getScriptTimeZone(), "dd/MM/yyyy") : '-',
          estado: dataMant[i][6] || 'PENDIENTE'
        });
      }
    }
    
    // Obtener últimas incidencias
    const dataInc = getSheetData('INCIDENCIAS');
    const incidencias = [];
    
    for(let i = 1; i < dataInc.length; i++) {
      if (String(dataInc[i][2]) === String(idActivo) && dataInc[i][1] === 'ACTIVO') {
        incidencias.push({
          descripcion: dataInc[i][4],
          prioridad: dataInc[i][5],
          estado: dataInc[i][6],
          fecha: dataInc[i][7] ? Utilities.formatDate(dataInc[i][7], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : '-'
        });
      }
    }
    
    return {
      success: true,
      activo: activo,
      revisiones: revisiones.slice(0, 5), // Últimas 5
      incidencias: incidencias.slice(0, 3) // Últimas 3
    };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ==========================================
// 25. SISTEMA DE RELACIONES ENTRE ACTIVOS
// ==========================================

/**
 * Obtener relaciones de un activo específico
 */
function getRelacionesActivo(idActivo) {
  const data = getSheetData('ACTIVOS');
  
  for(let i = 1; i < data.length; i++) {
    if(String(data[i][0]) === String(idActivo)) {
      // La columna 8 (índice 7) guardará las relaciones como JSON
      const relacionesRaw = data[i][7] || "[]";
      
      try {
        const relaciones = JSON.parse(relacionesRaw);
        
        // Enriquecer con nombres de los activos relacionados
        return relaciones.map(rel => {
          const activoInfo = getAssetInfo(rel.idActivoRelacionado);
          return {
            id: rel.id,
            idActivoRelacionado: rel.idActivoRelacionado,
            nombreActivo: activoInfo ? activoInfo.nombre : "Desconocido",
            tipoActivo: activoInfo ? activoInfo.tipo : "-",
            tipoRelacion: rel.tipoRelacion,
            descripcion: rel.descripcion || ""
          };
        });
      } catch(e) {
        return [];
      }
    }
  }
  
  return [];
}

/**
 * Guardar nueva relación entre activos (bidireccional)
 */
function crearRelacionActivo(datos) {
  verificarPermiso(['WRITE']);
  
  try {
    const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
    const sheet = ss.getSheetByName('ACTIVOS');
    const dataActivos = sheet.getDataRange().getValues();
    
    // 1. Añadir relación en Activo A -> B
    actualizarRelacionEnActivo(sheet, dataActivos, datos.idActivoA, {
      id: Utilities.getUuid(),
      idActivoRelacionado: datos.idActivoB,
      tipoRelacion: datos.tipoRelacion,
      descripcion: datos.descripcion
    });
    
    // 2. Añadir relación inversa en Activo B -> A (si es bidireccional)
    if(datos.bidireccional) {
      const tipoInverso = obtenerTipoInverso(datos.tipoRelacion);
      actualizarRelacionEnActivo(sheet, dataActivos, datos.idActivoB, {
        id: Utilities.getUuid(),
        idActivoRelacionado: datos.idActivoA,
        tipoRelacion: tipoInverso,
        descripcion: datos.descripcion
      });
    }
    
    registrarLog("CREAR RELACIÓN", `Entre activos ${datos.idActivoA} y ${datos.idActivoB} (${datos.tipoRelacion})`);
    
    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Helper: Actualizar relaciones en un activo específico
 */
function actualizarRelacionEnActivo(sheet, data, idActivo, nuevaRelacion) {
  for(let i = 1; i < data.length; i++) {
    if(String(data[i][0]) === String(idActivo)) {
      const relacionesActuales = data[i][7] || "[]";
      let relaciones = [];
      
      try {
        relaciones = JSON.parse(relacionesActuales);
      } catch(e) {
        relaciones = [];
      }
      
      // Añadir nueva relación
      relaciones.push(nuevaRelacion);
      
      // Guardar en Excel
      sheet.getRange(i + 1, 8).setValue(JSON.stringify(relaciones));
      break;
    }
  }
}

/**
 * Helper: Obtener tipo de relación inversa
 */
function obtenerTipoInverso(tipo) {
  const mapInversos = {
    "DEPENDE_DE": "ES_REQUERIDO_POR",
    "ES_REQUERIDO_POR": "DEPENDE_DE",
    "ALIMENTA": "ES_ALIMENTADO_POR",
    "ES_ALIMENTADO_POR": "ALIMENTA",
    "PERTENECE_A": "CONTIENE",
    "CONTIENE": "PERTENECE_A"
  };
  
  return mapInversos[tipo] || tipo;
}

/**
 * Eliminar una relación específica
 */
function eliminarRelacionActivo(idActivo, idRelacion) {
  verificarPermiso(['DELETE']);
  
  try {
    const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
    const sheet = ss.getSheetByName('ACTIVOS');
    const data = sheet.getDataRange().getValues();
    
    for(let i = 1; i < data.length; i++) {
      if(String(data[i][0]) === String(idActivo)) {
        const relacionesRaw = data[i][7] || "[]";
        let relaciones = JSON.parse(relacionesRaw);
        
        // Filtrar la relación a eliminar
        relaciones = relaciones.filter(r => r.id !== idRelacion);
        
        sheet.getRange(i + 1, 8).setValue(JSON.stringify(relaciones));
        
        registrarLog("ELIMINAR RELACIÓN", `Relación ${idRelacion} del activo ${idActivo}`);
        return { success: true };
      }
    }
    
    return { success: false, error: "Activo no encontrado" };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Obtener alertas de activos relacionados con problemas
 */
function getAlertasActivosRelacionados(idActivo) {
  const relaciones = getRelacionesActivo(idActivo);
  const alertas = [];
  
  // Revisar si algún activo relacionado tiene incidencias abiertas
  const incidencias = getSheetData('INCIDENCIAS');
  
  relaciones.forEach(rel => {
    for(let i = 1; i < incidencias.length; i++) {
      if(String(incidencias[i][2]) === String(rel.idActivoRelacionado) && 
         incidencias[i][6] !== 'RESUELTA') {
        alertas.push({
          nombreActivo: rel.nombreActivo,
          tipoRelacion: rel.tipoRelacion,
          problema: incidencias[i][4],
          prioridad: incidencias[i][5]
        });
      }
    }
  });
  
  return alertas;
}

/**
 * Buscar activos disponibles para relacionar (excluyendo el actual y ya relacionados)
 */
function buscarActivosParaRelacionar(idActivo, textoBusqueda) {
  const activos = getSheetData('ACTIVOS');
  const edificios = getSheetData('EDIFICIOS');
  const relacionesActuales = getRelacionesActivo(idActivo);
  
  const idsExcluidos = new Set([idActivo, ...relacionesActuales.map(r => r.idActivoRelacionado)]);
  
  const mapEdificios = {};
  edificios.slice(1).forEach(e => mapEdificios[e[0]] = e[2]);
  
  const resultados = [];
  const textoLower = (textoBusqueda || "").toLowerCase();
  
  for(let i = 1; i < activos.length; i++) {
    const id = activos[i][0];
    const nombre = activos[i][3];
    const tipo = activos[i][2];
    
    if(idsExcluidos.has(String(id))) continue;
    
    if(!textoBusqueda || 
       nombre.toLowerCase().includes(textoLower) || 
       tipo.toLowerCase().includes(textoLower)) {
      
      const edificio = mapEdificios[activos[i][1]] || "Sin edificio";
      
      resultados.push({
        id: id,
        nombre: nombre,
        tipo: tipo,
        edificio: edificio
      });
      
      if(resultados.length >= 20) break; // Limitar resultados
    }
  }
  
  return resultados;
}

/**
 * Función wrapper para obtener relaciones + alertas (llamada desde frontend)
 */
function obtenerDatosRelaciones(idActivo) {
  return {
    relaciones: getRelacionesActivo(idActivo),
    alertas: getAlertasActivosRelacionados(idActivo)
  };
}

function limpiarCache() {
  const CACHE = CacheService.getScriptCache();
  CACHE.removeAll(['SHEET_ACTIVOS', 'SHEET_EDIFICIOS', 'SHEET_CAMPUS', 'INDEX_ACTIVOS']);
  Logger.log('Caché limpiada');
}

// ==========================================
// 26. PLANIFICADOR (CALENDARIO GLOBAL)
// ==========================================

/**
 * Obtiene todos los eventos combinados para el planificador
 */
function getPlannerEvents() {
  const eventos = [];
  const hoy = new Date(); 
  hoy.setHours(0,0,0,0);

  // 1. REVISIONES DE MANTENIMIENTO
  const mant = getGlobalMaintenance(); // Reutilizamos tu función existente
  mant.forEach(r => {
    eventos.push({
      id: r.id,
      resourceId: 'MANTENIMIENTO', // Para filtrado
      title: `🔧 ${r.tipo} - ${r.activo}`,
      start: r.fechaISO, // YYYY-MM-DD
      color: r.color === 'rojo' ? '#dc3545' : (r.color === 'amarillo' ? '#ffc107' : '#198754'),
      textColor: r.color === 'amarillo' ? '#000' : '#fff',
      extendedProps: {
        tipo: 'MANTENIMIENTO',
        descripcion: `${r.edificio} | ${r.campusNombre}`,
        status: r.color,
        editable: true // Permitir drag & drop
      }
    });
  });

  // 2. OBRAS EN CURSO
  const dataObras = getSheetData('OBRAS');
  for(let i=1; i<dataObras.length; i++) {
    if(dataObras[i][6] === 'EN CURSO' && dataObras[i][4] instanceof Date) {
      eventos.push({
        id: dataObras[i][0],
        resourceId: 'OBRA',
        title: `🏗️ ${dataObras[i][2]}`,
        start: Utilities.formatDate(dataObras[i][4], Session.getScriptTimeZone(), "yyyy-MM-dd"),
        color: '#0d6efd', // Azul
        extendedProps: {
          tipo: 'OBRA',
          descripcion: dataObras[i][3],
          editable: true
        }
      });
    }
  }

  // 3. INCIDENCIAS PENDIENTES
  // Nota: Las incidencias suelen tener fecha de creación, no de "ejecución".
  // Las mostramos en el día que se reportaron para referencia.
  const dataInc = getSheetData('INCIDENCIAS');
  for(let i=1; i<dataInc.length; i++) {
    if(dataInc[i][6] !== 'RESUELTA' && dataInc[i][7] instanceof Date) {
      eventos.push({
        id: dataInc[i][0],
        resourceId: 'INCIDENCIA',
        title: `⚠️ ${dataInc[i][5]} - ${dataInc[i][3]}`,
        start: Utilities.formatDate(dataInc[i][7], Session.getScriptTimeZone(), "yyyy-MM-dd"),
        color: '#fd7e14', // Naranja
        extendedProps: {
          tipo: 'INCIDENCIA',
          descripcion: dataInc[i][4],
          editable: false // No solemos reprogramar la "creación" de una incidencia
        }
      });
    }
  }

  return eventos;
}

/**
 * Actualiza la fecha de un evento tras Drag & Drop
 */
function updateEventDate(id, tipo, nuevaFechaISO) {
  verificarPermiso(['WRITE']);
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const nuevaFecha = textoAFecha(nuevaFechaISO); // Tu helper espera YYYY-MM-DD

  if (tipo === 'MANTENIMIENTO') {
    const sheet = ss.getSheetByName('PLAN_MANTENIMIENTO');
    const data = sheet.getDataRange().getValues();
    for(let i=1; i<data.length; i++){
      if(String(data[i][0]) === String(id)) {
        sheet.getRange(i+1, 5).setValue(nuevaFecha); // Columna fecha
        registrarLog("REPROGRAMAR", `Revisión ${id} movida a ${nuevaFechaISO}`);
        return { success: true };
      }
    }
  } 
  else if (tipo === 'OBRA') {
    const sheet = ss.getSheetByName('OBRAS');
    const data = sheet.getDataRange().getValues();
    for(let i=1; i<data.length; i++){
      if(String(data[i][0]) === String(id)) {
        sheet.getRange(i+1, 5).setValue(nuevaFecha); // Columna fecha inicio
        registrarLog("REPROGRAMAR", `Obra ${id} movida a ${nuevaFechaISO}`);
        return { success: true };
      }
    }
  }

  return { success: false, error: "Elemento no encontrado o tipo no editable" };
}

// ==========================================
// 27. MÓDULO DE EXPORTACIÓN PARA AUDITORÍA
// ==========================================

/**
 * Obtiene los años disponibles en el histórico de mantenimiento
 */
function getAniosAuditoria() {
  const planes = getCachedSheetData('PLAN_MANTENIMIENTO');
  const anios = new Set();
  
  // Empezamos en 1 para saltar cabecera
  for(let i=1; i<planes.length; i++) {
    const fecha = planes[i][4]; // Columna Fecha
    if (fecha instanceof Date) {
      anios.add(fecha.getFullYear());
    }
  }
  // Convertir a array y ordenar descendente
  return Array.from(anios).sort((a,b) => b - a);
}

/**
 * Genera una carpeta en Drive con toda la documentación copiada
 */
function generarPaqueteAuditoria(anio, tipoFiltro) {
  // tipoFiltro puede ser 'Legal', 'Todos', etc.
  
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const docsData = getSheetData('DOCS_HISTORICO');
  const planData = getSheetData('PLAN_MANTENIMIENTO');
  const activosData = getCachedSheetData('ACTIVOS');
  const edifData = getCachedSheetData('EDIFICIOS');
  
  // 1. Mapear Activos y Edificios para obtener nombres rápidos
  const mapEdificios = {};
  edifData.slice(1).forEach(r => mapEdificios[String(r[0])] = r[2]); // ID -> Nombre
  
  const mapActivos = {};
  activosData.slice(1).forEach(r => {
    mapActivos[String(r[0])] = {
      nombre: r[3],
      nombreEdificio: mapEdificios[String(r[1])] || "Sin Ubicación"
    };
  });

  // 2. Filtrar Revisiones que cumplen el criterio (Año y Tipo)
  const revisionesValidas = {}; // Map ID_PLAN -> Info para nombre archivo
  
  for(let i=1; i<planData.length; i++) {
    const row = planData[i];
    const fecha = row[4];
    const tipo = row[2]; // Legal, Periódica...
    const idActivo = row[1];
    const idPlan = String(row[0]);
    
    // Filtro de Estado: Solo las REALIZADAS (o todas si prefieres auditar lo pendiente también)
    // Para auditoría, generalmente se buscan evidencias de lo hecho.
    // Si quieres incluir todo, quita el if de estado.
    // if (row[6] !== 'REALIZADA') continue; 

    if (fecha instanceof Date && String(fecha.getFullYear()) === String(anio)) {
      if (tipoFiltro === 'TODOS' || tipo === tipoFiltro) {
        const activoInfo = mapActivos[String(idActivo)] || { nombre: "Activo Borrado", nombreEdificio: "_" };
        
        revisionesValidas[idPlan] = {
          fechaStr: Utilities.formatDate(fecha, Session.getScriptTimeZone(), "yyyy-MM-dd"),
          tipo: tipo,
          activo: activoInfo.nombre,
          edificio: activoInfo.nombreEdificio
        };
      }
    }
  }
  
  // 3. Buscar Documentos asociados a esas revisiones
  const archivosACopiar = [];
  
  for(let i=1; i<docsData.length; i++) {
    const r = docsData[i];
    const tipoEntidad = r[1];
    const idEntidad = String(r[2]); // En caso de REVISION, es el ID del Plan
    const fileId = r[8]; // Columna I (índice 8) es el ID de archivo de Drive
    
    if (tipoEntidad === 'REVISION' && revisionesValidas[idEntidad] && fileId) {
      archivosACopiar.push({
        fileId: fileId,
        info: revisionesValidas[idEntidad],
        originalName: r[3]
      });
    }
  }
  
  if (archivosACopiar.length === 0) {
    return { success: false, error: "No se encontraron documentos para ese año y criterio." };
  }
  
  // 4. Crear Estructura en Drive
  try {
    const rootFolderId = getRootFolderId();
    const parentFolder = DriveApp.getFolderById(rootFolderId);
    
    const folderName = `AUDITORIA_${anio}_${tipoFiltro}_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HHmm")}`;
    const auditFolder = parentFolder.createFolder(folderName);
    
    let copiados = 0;
    let errores = 0;
    
    // 5. Copiar y renombrar archivos
    archivosACopiar.forEach(item => {
      try {
        const file = DriveApp.getFileById(item.fileId);
        
        // Nomenclatura ordenada: [FECHA]_[EDIFICIO]_[ACTIVO]_[TIPO].pdf
        // Esto permite que al ordenar por nombre en Windows/Mac, salgan cronológicos y por sitio.
        const cleanName = `${item.info.fechaStr}_${item.info.edificio}_${item.info.activo}_${item.info.tipo}`;
        // Limpiar caracteres raros del nombre
        const safeName = cleanName.replace(/[^a-zA-Z0-9_\-\.]/g, '_') + ".pdf"; // Asumimos PDF, o preservar extensión original
        
        file.makeCopy(safeName, auditFolder);
        copiados++;
      } catch(e) {
        console.log("Error copiando archivo auditoría: " + e);
        errores++;
      }
    });
    
    return { 
      success: true, 
      url: auditFolder.getUrl(), 
      total: copiados,
      errores: errores,
      folderName: folderName
    };
    
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Procesa la subida rápida clasificando el archivo y generando revisiones con seguridad.
 */
function procesarArchivoRapido(data) {
  try {
    const ss = getDB();
    let idEntidadDestino = data.idActivo;
    let tipoEntidadDestino = 'ACTIVO';
    let nombreFinal = data.nombreArchivo;

    // A. LÓGICA OCA (Ya implementada)
    if (data.categoria === 'OCA') {
      const sheetMant = ss.getSheetByName('PLAN_MANTENIMIENTO');
      const idRevision = Utilities.getUuid();
      const fechaReal = data.fechaOCA ? textoAFecha(data.fechaOCA) : new Date();
      
      sheetMant.appendRow([
        idRevision, data.idActivo, 'Legal', 'OCA - Documento Histórico/Carga',
        fechaReal, data.freqDias || 365, 'REALIZADA', null
      ]);

      if (data.crearSiguientes && data.freqDias > 0) {
        const frecuencia = parseInt(data.freqDias);
        const MAX_REVISIONES = 10;
        const HORIZONTE_ANIOS = 10;
        const hoy = new Date();
        const fechaTope = new Date(hoy.getFullYear() + HORIZONTE_ANIOS, hoy.getMonth(), hoy.getDate());
        let fechaSiguiente = new Date(fechaReal);
        let contador = 0;

        while (contador < MAX_REVISIONES) {
          fechaSiguiente.setDate(fechaSiguiente.getDate() + frecuencia);
          if (fechaSiguiente > fechaTope) break;
          sheetMant.appendRow([
            Utilities.getUuid(), data.idActivo, 'Legal', 'Próxima Inspección Reglamentaria',
            new Date(fechaSiguiente), frecuencia, 'ACTIVO', null
          ]);
          contador++;
        }
      }
      idEntidadDestino = idRevision;
      tipoEntidadDestino = 'REVISION';
      nombreFinal = "OCA_" + nombreFinal;
    } 
    
    // B. LÓGICA CONTRATO (NUEVA)
    else if (data.categoria === 'CONTRATO') {
      // 1. Crear el registro en la hoja CONTRATOS
      const sheetCont = ss.getSheetByName('CONTRATOS');
      const idContrato = Utilities.getUuid();
      
      // Datos obligatorios o por defecto
      const prov = data.contProveedor || "Proveedor Desconocido";
      const ref = data.contRef || "S/N";
      const ini = data.contIni ? textoAFecha(data.contIni) : new Date();
      // Si no pone fin, ponemos 1 año por defecto
      let fin = data.contFin ? textoAFecha(data.contFin) : new Date();
      if (!data.contFin) fin.setFullYear(fin.getFullYear() + 1);

      sheetCont.appendRow([
        idContrato,
        'ACTIVO',       // Tipo Entidad
        data.idActivo,  // ID Entidad
        prov,
        ref,
        ini,
        fin,
        'ACTIVO'        // Estado DB
      ]);

      // 2. El archivo se guarda asociado al activo, pero con nombre claro
      nombreFinal = "CONTRATO_" + prov + "_" + nombreFinal;
      
      // Opcional: Podríamos vincular el doc al contrato si tuvieras esa estructura, 
      // pero por ahora lo dejamos en el activo para que sea fácil de encontrar.
      registrarLog("SUBIDA RÁPIDA", "Contrato creado para activo " + data.idActivo);
    }

    return subirArchivo(data.base64, nombreFinal, data.mimeType, idEntidadDestino, tipoEntidadDestino);
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}
