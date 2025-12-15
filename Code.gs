// GMAO
// Universidad de Navarra
// Versión 1.3
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
 */
function getAppInitData() {
  const usuario = getMyRole();
  
  // 1. Campus ordenados
  const campusRaw = getCachedSheetData('CAMPUS');
  const listaCampus = campusRaw.slice(1)
    .map(r => ({ id: r[0], nombre: r[1] }))
    .sort((a, b) => a.nombre.localeCompare(b.nombre, 'es', { sensitivity: 'base' })); // <--- AÑADIDO
  
  // 2. Catálogo ordenado
  const catalogoRaw = getCachedSheetData('CAT_INSTALACIONES');
  const catalogo = catalogoRaw.slice(1)
    .map(r => ({ id: r[0], nombre: r[1], dias: r[3] }))
    .sort((a, b) => a.nombre.localeCompare(b.nombre, 'es', { sensitivity: 'base' })); // <--- AÑADIDO

  return {
    usuario: usuario,
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
  
  // 1. INTENTAR LEER DESDE CACHÉ
  try {
    const cached = CACHE.get(cacheKey);
    if (cached) {
      const parsed = JSON.parse(cached);
      Logger.log('✅ Índice cargado desde caché');
      return parsed;
    }
  } catch(e) {
    Logger.log('⚠️ Error leyendo caché de índice: ' + e);
  }
  
  // 2. CONSTRUIR ÍNDICE DESDE CERO
  try {
    Logger.log('🔨 Construyendo índice desde hojas...');
    
    const activos = getCachedSheetData('ACTIVOS');
    const edificios = getCachedSheetData('EDIFICIOS');
    const campus = getCachedSheetData('CAMPUS');
    
    if (!activos || !Array.isArray(activos) || activos.length < 2) {
      Logger.log('⚠️ ADVERTENCIA: Hoja ACTIVOS vacía');
      return {
        byId: {},
        byEdificio: {},
        byCampus: {},
        searchable: []
      };
    }
    
    // Crear mapas de Campus
    const mapCampus = {};
    if (campus && campus.length > 1) {
      campus.slice(1).forEach(r => {
        if (r && r[0] && r[1]) {
          mapCampus[String(r[0])] = r[1];
        }
      });
    }
    Logger.log(`📍 ${Object.keys(mapCampus).length} campus mapeados`);
    
    // Crear mapas de Edificios
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
    Logger.log(`🏢 ${Object.keys(mapEdificios).length} edificios mapeados`);
    
    // Índice principal
    const index = {
      byId: {},
      byEdificio: {},
      byCampus: {},
      searchable: []
    };
    
    let procesados = 0;
    let errores = 0;
    
    // Procesar cada activo
    activos.slice(1).forEach((r, idx) => {
      try {
        if (!r || !r[0]) return;
        
        const id = String(r[0]);
        const idEdif = String(r[1] || '');
        const edifInfo = mapEdificios[idEdif] || {};
        
        // ✅ CONVERTIR FECHA A STRING (CRÍTICO)
        let fechaAltaStr = null;
        if (r[5]) {
          if (r[5] instanceof Date) {
            fechaAltaStr = Utilities.formatDate(r[5], Session.getScriptTimeZone(), "dd/MM/yyyy");
          } else {
            fechaAltaStr = String(r[5]);
          }
        }
        
        const activo = {
          id: id,
          idEdificio: idEdif,
          tipo: r[2] || '-',
          nombre: r[3] || 'Sin nombre',
          marca: r[4] || '',
          fechaAlta: fechaAltaStr, // ✅ STRING, no Date
          edificio: edifInfo.nombre || '-',
          campus: edifInfo.campusNombre || '-',
          idCampus: edifInfo.idCampus || null
        };
        
        // Por ID
        index.byId[id] = activo;
        
        // Por Edificio
        if (idEdif) {
          if (!index.byEdificio[idEdif]) {
            index.byEdificio[idEdif] = [];
          }
          index.byEdificio[idEdif].push(activo);
        }
        
        // Por Campus
        if (edifInfo.idCampus) {
          if (!index.byCampus[edifInfo.idCampus]) {
            index.byCampus[edifInfo.idCampus] = [];
          }
          index.byCampus[edifInfo.idCampus].push(activo);
        }
        
        // Para búsquedas
        const searchText = (activo.nombre + ' ' + activo.tipo + ' ' + activo.marca).toLowerCase();
        index.searchable.push({
          id: id,
          text: searchText
        });
        
        procesados++;
        
      } catch(e) {
        Logger.log(`❌ Error procesando activo en fila ${idx + 2}: ${e.toString()}`);
        errores++;
      }
    });
    
    Logger.log(`✅ Índice construido: ${procesados} activos procesados, ${errores} errores`);
    Logger.log(`📊 Edificios con activos: ${Object.keys(index.byEdificio).length}`);
    
    const ejemploEdif = Object.keys(index.byEdificio)[0];
    if (ejemploEdif) {
      Logger.log(`🔍 Ejemplo - Edificio ${ejemploEdif} tiene ${index.byEdificio[ejemploEdif].length} activos`);
    }
    
    // Guardar en caché
    try {
      const serialized = JSON.stringify(index);
      if (serialized.length < 90000) {
        CACHE.put(cacheKey, serialized, CACHE_TIME);
        Logger.log('💾 Índice guardado en caché');
      } else {
        Logger.log('⚠️ Índice demasiado grande para caché: ' + serialized.length + ' bytes');
      }
    } catch(e) {
      Logger.log('⚠️ Error guardando índice en caché: ' + e);
    }
    
    return index;
    
  } catch(e) {
    Logger.log('💥 ERROR CRÍTICO construyendo índice: ' + e.toString());
    Logger.log('Stack trace: ' + e.stack);
    
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
// SUSTITUIR EN Code.gs

function getMyRole() {
  const email = Session.getActiveUser().getEmail();
  const cacheKey = 'USER_ROLE_' + email;
  const cache = CacheService.getUserCache(); // Caché privada del usuario
  
  // 1. Intentar leer de memoria rápida
  const cachedRole = cache.get(cacheKey);
  if (cachedRole) {
    return JSON.parse(cachedRole);
  }

  // 2. Si no está, leer de la hoja (Lento)
  // Usamos getCachedSheetData para aprovechar la caché general si ya existe
  const data = getCachedSheetData('USUARIOS'); 
  
  let usuario = { email: email, nombre: "Invitado", rol: 'CONSULTA' }; // Default

  // Buscamos (empezando en 1 para saltar cabecera)
  for(let i=1; i<data.length; i++) {
    if(String(data[i][2]).trim().toLowerCase() === email.toLowerCase()) {
      usuario = { email: email, nombre: data[i][1], rol: data[i][3] }; 
      break;
    }
  }

  // 3. Guardar en memoria rápida por 20 minutos
  cache.put(cacheKey, JSON.stringify(usuario), 1200);
  
  return usuario;
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

function getListaCampus() { 
  const data = getSheetData('CAMPUS'); 
  return data.slice(1)
    .map(r => ({ id: r[0], nombre: r[1] }))
    .sort((a, b) => a.nombre.localeCompare(b.nombre, 'es', { sensitivity: 'base' })); 
}

function getEdificiosPorCampus(idCampus) { 
  const data = getSheetData('EDIFICIOS'); 
  return data.slice(1)
    .filter(r => String(r[1]) === String(idCampus))
    .map(r => ({ id: r[0], nombre: r[2] }))
    .sort((a, b) => a.nombre.localeCompare(b.nombre, 'es', { sensitivity: 'base' })); // <--- AÑADIDO
}

function getActivosPorEdificio(idEdificio) {
  try {
    Logger.log('═══════════════════════════════════════');
    Logger.log('🚀 INICIO: getActivosPorEdificio');
    Logger.log('═══════════════════════════════════════');
    
    // 1. VALIDACIÓN DEL PARÁMETRO
    if (!idEdificio || String(idEdificio).trim() === '') {
      Logger.log('❌ ERROR: idEdificio está vacío o inválido');
      Logger.log('Valor recibido: ' + JSON.stringify(idEdificio));
      return [];
    }
    
    const idEdificioStr = String(idEdificio).trim();
    Logger.log('✅ ID Edificio válido: ' + idEdificioStr);
    
    // 2. CONSTRUIR ÍNDICE
    Logger.log('📊 Llamando a buildActivosIndex()...');
    const index = buildActivosIndex();
    
    // 3. VALIDAR EL ÍNDICE
    if (!index) {
      Logger.log('💥 ERROR CRÍTICO: buildActivosIndex() devolvió null/undefined');
      return [];
    }
    Logger.log('✅ Índice recibido correctamente');
    Logger.log('📋 Estructura del índice: ' + JSON.stringify(Object.keys(index)));
    
    if (!index.byEdificio) {
      Logger.log('❌ ERROR: index.byEdificio no existe');
      Logger.log('Propiedades disponibles: ' + Object.keys(index).join(', '));
      return [];
    }
    Logger.log('✅ index.byEdificio existe');
    
    // 4. MOSTRAR IDS DE EDIFICIOS DISPONIBLES
    const edificiosDisponibles = Object.keys(index.byEdificio);
    Logger.log(`🏢 Total de edificios con activos: ${edificiosDisponibles.length}`);
    Logger.log('🔑 IDs disponibles: ' + edificiosDisponibles.join(', '));
    
    // 5. BUSCAR ACTIVOS
    const activos = index.byEdificio[idEdificioStr];
    
    if (!activos) {
      Logger.log('⚠️ No se encontró el edificio en el índice');
      Logger.log('🔍 Buscando coincidencias parciales...');
      
      // Intentar encontrar coincidencias
      const coincidencias = edificiosDisponibles.filter(id => 
        id.includes(idEdificioStr) || idEdificioStr.includes(id)
      );
      
      if (coincidencias.length > 0) {
        Logger.log('💡 Posibles coincidencias: ' + coincidencias.join(', '));
      } else {
        Logger.log('🚫 Sin coincidencias encontradas');
      }
      
      return [];
    }
    
    if (activos.length === 0) {
      Logger.log('📭 El edificio existe pero no tiene activos');
      return [];
    }
    
    // 6. ÉXITO
    Logger.log('✅ ÉXITO: ' + activos.length + ' activos encontrados');
    Logger.log('📦 Primer activo: ' + JSON.stringify(activos[0]));
    Logger.log('═══════════════════════════════════════');
    
    return activos;
    
  } catch(e) {
    Logger.log('═══════════════════════════════════════');
    Logger.log('💥 EXCEPCIÓN CAPTURADA');
    Logger.log('Error: ' + e.toString());
    Logger.log('Stack: ' + e.stack);
    Logger.log('═══════════════════════════════════════');
    return [];
  }
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
  return list.sort((a, b) => {
    // Ordenar alfabéticamente por nombre (insensible a mayúsculas/minúsculas)
    return a.nombre.localeCompare(b.nombre, 'es', { sensitivity: 'base' });
});
}


function getAssetInfo(idActivo) {
  // Aseguramos que el índice esté fresco
  const index = buildActivosIndex();
  const activo = index.byId[String(idActivo)];
  
  if (!activo) return null;

  // Debug para verificar
  Logger.log('Recuperando activo recién creado: ' + activo.nombre);
  
  return {
    id: activo.id,
    nombre: activo.nombre,
    tipo: activo.tipo,
    marca: activo.marca,
    fechaAlta: activo.fechaAlta instanceof Date ? 
               Utilities.formatDate(activo.fechaAlta, Session.getScriptTimeZone(), "dd/MM/yyyy") : "-",
    // --- CORRECCIÓN AQUÍ ---
    // El frontend espera 'edificioNombre' para la lista, pero 'edificio' para el detalle.
    // Enviamos ambos para evitar el "undefined".
    edificio: activo.edificio,       
    edificioNombre: activo.edificio, 
    // -----------------------
    idEdificio: activo.idEdificio,
    campus: activo.campus,
    idCampus: activo.idCampus
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

    // A. LÓGICA ESTÁNDAR (Entidades directas)
    if (tipoEntidad === 'ACTIVO') {
      const rows = getSheetData('ACTIVOS'); for(let i=1; i<rows.length; i++) if(String(rows[i][0]) === String(idEntidad)) { carpetaId = rows[i][6]; break; }
    } 
    else if (tipoEntidad === 'EDIFICIO') {
      const rows = getSheetData('EDIFICIOS'); for(let i=1; i<rows.length; i++) if(String(rows[i][0]) === String(idEntidad)) { carpetaId = rows[i][4]; break; }
    } 
    else if (tipoEntidad === 'OBRA') {
      const rows = getSheetData('OBRAS'); for(let i=1; i<rows.length; i++) if(String(rows[i][0]) === String(idEntidad)) { carpetaId = rows[i][7]; break; }
    } 
    else if (tipoEntidad === 'REVISION') {
      const planes = getSheetData('PLAN_MANTENIMIENTO'); let activoId = null;
      for(let i=1; i<planes.length; i++) if(String(planes[i][0]) === String(idEntidad)) { activoId = planes[i][1]; break; }
      if(activoId) { const rows = getSheetData('ACTIVOS'); for(let i=1; i<rows.length; i++) if(String(rows[i][0]) === String(activoId)) { carpetaId = rows[i][6]; break; } }
    } 
    
    // B. LÓGICA ESPECIAL PARA CONTRATOS
    else if (tipoEntidad === 'CONTRATO') {
      const contratos = getSheetData('CONTRATOS');
      for(let i=1; i<contratos.length; i++) {
        if(String(contratos[i][0]) === String(idEntidad)) {
           const tipoTarget = contratos[i][1]; // CAMPUS, EDIFICIO, ACTIVO, ACTIVOS
           const idTarget = contratos[i][2];
           
           // 1. Contrato de ACTIVO ÚNICO
           if(tipoTarget === 'ACTIVO') {
              const rows = getSheetData('ACTIVOS'); 
              for(let k=1; k<rows.length; k++) if(String(rows[k][0]) === String(idTarget)) { carpetaId = rows[k][6]; break; }
           } 
           // 2. Contrato de EDIFICIO
           else if(tipoTarget === 'EDIFICIO') {
              const rows = getSheetData('EDIFICIOS'); 
              for(let k=1; k<rows.length; k++) if(String(rows[k][0]) === String(idTarget)) { carpetaId = rows[k][4]; break; }
           }
           // 3. Contrato de CAMPUS
           else if(tipoTarget === 'CAMPUS') {
              const rows = getSheetData('CAMPUS'); 
              for(let k=1; k<rows.length; k++) if(String(rows[k][0]) === String(idTarget)) { carpetaId = rows[k][4]; break; } // Col 4 es carpetaID en Campus
           }
           // 4. Contrato de MÚLTIPLES ACTIVOS (JSON)
           else if(tipoTarget === 'ACTIVOS') {
              try {
                const ids = JSON.parse(idTarget);
                if (ids && ids.length > 0) {
                   const primerId = ids[0]; // Usamos la carpeta del primer activo de la lista
                   const rows = getSheetData('ACTIVOS'); 
                   for(let k=1; k<rows.length; k++) if(String(rows[k][0]) === String(primerId)) { carpetaId = rows[k][6]; break; }
                }
              } catch(e) { Logger.log("Error parseando activos contrato: " + e); }
           }
           break;
        }
      }
    }
    
    // Si falla todo, a la raíz
    if (!carpetaId) carpetaId = getRootFolderId();
    
    const blob = Utilities.newBlob(Utilities.base64Decode(dataBase64), mimeType, nombreArchivo);
    const folder = DriveApp.getFolderById(carpetaId);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    if(tipoEntidad !== 'INCIDENCIA') {
       ss.getSheetByName('DOCS_HISTORICO').appendRow([
         Utilities.getUuid(), 
         tipoEntidad, 
         idEntidad, 
         nombreArchivo, 
         file.getUrl(), 
         1, 
         new Date(), 
         Session.getActiveUser().getEmail(), 
         file.getId()
       ]);
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

// SUSTITUIR EN Code.gs

function obtenerContratos(idEntidad) {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); 
  const sheet = ss.getSheetByName('CONTRATOS'); 
  if (!sheet || sheet.getLastRow() < 2) return []; 
  const data = sheet.getRange(1, 1, sheet.getLastRow(), 8).getValues(); 
  
  // Mapa Proveedores
  const provData = getSheetData('PROVEEDORES');
  const mapProv = {};
  if (provData.length > 1) {
    for (let k = 1; k < provData.length; k++) {
       if(provData[k][0]) mapProv[String(provData[k][0]).trim()] = provData[k][1];
    }
  }

  // Mapa Archivos
  const docsData = getSheetData('DOCS_HISTORICO');
  const fileMap = {};
  if (docsData.length > 1) {
    for(let j=1; j<docsData.length; j++) {
       if(String(docsData[j][1]) === 'CONTRATO') {
          fileMap[String(docsData[j][2]).trim()] = docsData[j][4]; 
       }
    }
  }

  const contratos = []; 
  const hoy = new Date(); 
  const idEntidadStr = String(idEntidad).trim(); // Limpiamos el ID que buscamos
  
  for(let i=1; i<data.length; i++) { 
    if(!data[i][0]) continue;

    const tipo = String(data[i][1]).trim();
    const idVinculado = String(data[i][2]).trim();
    let esDeEsteActivo = false;

    if (idVinculado === idEntidadStr) {
        esDeEsteActivo = true;
    } else if (tipo === 'ACTIVOS' && idVinculado.includes(idEntidadStr)) {
        esDeEsteActivo = true; 
    }

    if(esDeEsteActivo) { 
      const fFin = data[i][6] instanceof Date ? data[i][6] : null; 
      const fIni = data[i][5] instanceof Date ? data[i][5] : null; 
      let estadoDB = (data[i].length > 7) ? data[i][7] : 'ACTIVO'; 
      let estadoCalc = 'VIGENTE'; 
      let color = 'verde'; 
      
      if (estadoDB === 'INACTIVO') { estadoCalc = 'INACTIVO'; color = 'gris'; } 
      else if (fFin) { 
        const diff = Math.ceil((fFin.getTime() - hoy.getTime()) / (86400000)); 
        if (diff < 0) { estadoCalc = 'CADUCADO'; color = 'rojo'; } 
        else if (diff <= 30) { estadoCalc = 'PRÓXIMO'; color = 'amarillo'; } 
      } else { estadoCalc = 'SIN FECHA'; color = 'gris'; } 
      
      const idProv = String(data[i][3]).trim();
      const nombreProv = mapProv[idProv] || idProv;

      contratos.push({ 
        id: data[i][0], 
        proveedor: nombreProv, 
        ref: data[i][4], 
        inicio: fIni?fechaATexto(fIni):"-", 
        fin: fFin?fechaATexto(fFin):"-", 
        estado: estadoCalc, 
        color: color, 
        estadoDB: estadoDB,
        fileUrl: fileMap[String(data[i][0]).trim()] || null
      }); 
    } 
  } 
  return contratos.sort((a, b) => a.fin.localeCompare(b.fin));
}

function updateContratoV2(d) {
  verificarPermiso(['WRITE']);
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const sheet = ss.getSheetByName('CONTRATOS');
  const data = sheet.getDataRange().getValues();

  // Determinar ID Entidad (Si es array de activos, lo convertimos a JSON string)
  let idEntidadFinal = d.idEntidad;
  if (d.tipoEntidad === 'ACTIVOS' && d.idsActivos && d.idsActivos.length > 0) {
      idEntidadFinal = JSON.stringify(d.idsActivos);
  }

  for(let i=1; i<data.length; i++){
    if(String(data[i][0]) === String(d.id)) {
      
      // Actualizar columnas
      sheet.getRange(i+1, 2).setValue(d.tipoEntidad); // B: Tipo
      sheet.getRange(i+1, 3).setValue(idEntidadFinal); // C: ID Entidad
      sheet.getRange(i+1, 4).setValue(d.idProveedor); // D: Proveedor (Ahora ID)
      sheet.getRange(i+1, 5).setValue(d.ref);         // E: Ref
      sheet.getRange(i+1, 6).setValue(textoAFecha(d.fechaIni)); // F: Inicio
      sheet.getRange(i+1, 7).setValue(textoAFecha(d.fechaFin)); // G: Fin
      sheet.getRange(i+1, 8).setValue(d.estado);       // H: Estado
      
      registrarLog("EDITAR CONTRATO", "ID: " + d.id + " | Ref: " + d.ref);
      invalidateCache('CONTRATOS');
      return { success: true };
    }
  }
  return { success: false, error: "Contrato no encontrado para editar" };
}

function obtenerContratosGlobal() {
  const contratos = getSheetData('CONTRATOS'); 
  // Forzamos recarga de caché de activos por si hay nuevos
  const activos = getCachedSheetData('ACTIVOS', true); 
  const edificios = getSheetData('EDIFICIOS'); 
  const campus = getSheetData('CAMPUS'); 
  const proveedores = getListaProveedores();

  // 1. MAPA DE PROVEEDORES (BLINDADO)
  const mapProveedores = {};
  if (proveedores && proveedores.length > 0) {
      proveedores.forEach(p => {
          // Clave: ID limpio y en minúsculas para asegurar match
          if (p.id) mapProveedores[String(p.id).trim()] = p.nombre;
      });
  }
  
  // 2. MAPA DE UBICACIONES
  const mapEdificios = {}; 
  edificios.slice(1).forEach(r => {
    if(r[0]) mapEdificios[String(r[0]).trim()] = { nombre: r[2], idCampus: String(r[1]).trim() }; 
  }); 
  
  const mapCampus = {}; 
  campus.slice(1).forEach(r => { if(r[0]) mapCampus[String(r[0]).trim()] = r[1]; }); 

  const mapActivos = {}; 
  activos.slice(1).forEach(r => { 
    if (r[0]) {
        const idA = String(r[0]).trim();
        const idE = String(r[1]).trim();
        const infoEdif = mapEdificios[idE] || {};
        mapActivos[idA] = { 
          nombre: r[3], 
          idEdif: idE, 
          idCampus: infoEdif.idCampus,
          nombreEdificio: infoEdif.nombre 
        }; 
    }
  });
  
  // Mapa de Archivos
  const docsData = getSheetData('DOCS_HISTORICO');
  const fileMap = {};
  for(let j=1; j<docsData.length; j++) {
     if(String(docsData[j][1]) === 'CONTRATO') fileMap[String(docsData[j][2]).trim()] = docsData[j][4]; 
  }

  const result = []; 
  const hoy = new Date();
  
  // 3. ITERAR CONTRATOS
  for(let i=1; i<contratos.length; i++) {
    const r = contratos[i]; 
    if (!r[0]) continue; // Saltar filas vacías

    const idContrato = String(r[0]).trim();
    const tipoEntidad = String(r[1]).trim();
    const idEntidadRaw = r[2]; // ¡NO convertir a String aún! Puede ser Array
    const idProveedor = String(r[3] || "").trim();
    
    let nombreEntidad = "Sin Asignar";
    let edificioId = null; 
    let campusId = null;
    let campusNombre = "-"; 

    // --- LÓGICA ACTIVOS MÚLTIPLES (CORREGIDA) ---
    if (tipoEntidad === 'ACTIVOS') {
       let ids = [];
       // Detectar si Google Sheets ya nos dio un Array o un String JSON
       if (Array.isArray(idEntidadRaw)) {
           ids = idEntidadRaw;
       } else if (typeof idEntidadRaw === 'string' && idEntidadRaw.trim().startsWith('[')) {
           try { ids = JSON.parse(idEntidadRaw); } catch(e) { Logger.log("Error JSON: " + e); }
       }

       if (ids && ids.length > 0) {
           nombreEntidad = ids.length + " Activos";
           // Buscar el primer activo válido para sacar el edificio
           for(let k=0; k<ids.length; k++) {
               const idAct = String(ids[k]).trim();
               if(mapActivos[idAct]) {
                   const info = mapActivos[idAct];
                   campusId = info.idCampus;
                   edificioId = info.idEdif;
                   nombreEntidad += ` (${info.nombreEdificio || 'Sin Edif.'})`;
                   break; // Ya tenemos ubicación, salimos
               }
           }
       }
    }
    // Lógica Activo Único
    else if (tipoEntidad === 'ACTIVO') { 
      const info = mapActivos[String(idEntidadRaw).trim()]; 
      if (info) {
          nombreEntidad = `${info.nombre} (${info.nombreEdificio || ''})`; 
          edificioId = info.idEdif; 
          campusId = info.idCampus; 
      }
    } 
    // Lógica Edificio
    else if (tipoEntidad === 'EDIFICIO') { 
      const info = mapEdificios[String(idEntidadRaw).trim()]; 
      if (info) {
          nombreEntidad = info.nombre; 
          edificioId = String(idEntidadRaw).trim(); 
          campusId = info.idCampus; 
      }
    }
    // Lógica Campus
    else if (tipoEntidad === 'CAMPUS') {
       campusId = String(idEntidadRaw).trim();
       nombreEntidad = "General Campus";
    }

    if (campusId && mapCampus[campusId]) campusNombre = mapCampus[campusId];

    // Estados
    const fFin = (r[6] instanceof Date) ? r[6] : null; 
    let estadoDB = r[7] || 'ACTIVO'; 
    let estadoCalc = 'VIGENTE'; 
    let color = 'verde';
    if (estadoDB === 'INACTIVO') { estadoCalc = 'INACTIVO'; color = 'gris'; } 
    else if (fFin) { 
      const diff = Math.ceil((fFin.getTime() - hoy.getTime()) / (86400000)); 
      if (diff < 0) { estadoCalc = 'CADUCADO'; color = 'rojo'; } 
      else if (diff <= 30) { estadoCalc = 'PRÓXIMO'; color = 'amarillo'; } 
    } else { estadoCalc = 'SIN FECHA'; color = 'gris'; } 
    
    // --- TRADUCCIÓN PROVEEDOR ---
    const nombreProveedor = mapProveedores[idProveedor] || idProveedor;

    result.push({ 
      id: idContrato, 
      nombreEntidad: nombreEntidad, 
      proveedor: nombreProveedor, 
      idProveedor: idProveedor,
      ref: r[4], 
      inicio: r[5] ? fechaATexto(r[5]) : "-", 
      fin: fFin ? fechaATexto(fFin) : "-", 
      estado: estadoCalc, 
      color: color, 
      estadoDB: estadoDB, 
      edificioId: edificioId, 
      campusId: campusId,
      campusNombre: campusNombre,
      fileUrl: fileMap[idContrato] || null
    });
  }
  
  return result.sort((a, b) => {
    // Primero por nombre de Proveedor
    const compareProv = a.proveedor.localeCompare(b.proveedor, 'es', { sensitivity: 'base' });
    
    // Si es el mismo proveedor, ordenar por fecha fin (lo más urgente primero)
    if (compareProv === 0) {
        if (a.fin === "-") return 1;
        if (b.fin === "-") return -1;
        return a.fin.localeCompare(b.fin);
    }
    
    return compareProv;
});
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
  try {
    verificarPermiso(['WRITE']);
    if (!d.nombre) throw new Error("El nombre es obligatorio.");

    const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
    
    // Intentar crear carpeta, si falla usar raíz
    let fId = getRootFolderId();
    try {
      fId = crearCarpeta(d.nombre, fId); // Crea carpeta dentro de la raíz
    } catch (err) {
      console.warn("Error creando carpeta de Campus: " + err);
    }

    ss.getSheetByName('CAMPUS').appendRow([
      Utilities.getUuid(), 
      d.nombre, 
      d.provincia, 
      d.direccion, 
      fId
    ]);
    
    SpreadsheetApp.flush(); // Forzar guardado
    invalidateCache('CAMPUS');
    registrarLog("CREAR CAMPUS", "Nombre: " + d.nombre);
    
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function crearEdificio(d) {
  try {
    verificarPermiso(['WRITE']);
    if (!d.nombre || !d.idCampus) throw new Error("Faltan datos obligatorios.");

    const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
    
    // 1. Buscar carpeta del Campus padre
    const cData = getCachedSheetData('CAMPUS');
    let pId = getRootFolderId(); // Por defecto raíz
    
    for(let i=1; i<cData.length; i++) {
      if(String(cData[i][0]) === String(d.idCampus)) {
        if(cData[i][4]) pId = cData[i][4]; // Usar carpeta del campus si existe
        break;
      }
    }

    // 2. Crear carpetas (Edificio y Activos)
    let fId = "", aId = "";
    try {
      fId = crearCarpeta(d.nombre, pId);
      aId = crearCarpeta("Activos - " + d.nombre, fId);
    } catch (err) {
      console.warn("Error carpetas edificio: " + err);
      if(!fId) fId = pId; // Fallback
      if(!aId) aId = pId;
    }

    // 3. Guardar
    ss.getSheetByName('EDIFICIOS').appendRow([
      Utilities.getUuid(), 
      d.idCampus, 
      d.nombre, 
      d.contacto, 
      fId, 
      aId, 
      d.lat, 
      d.lng
    ]);

    SpreadsheetApp.flush();
    invalidateCache('EDIFICIOS');
    registrarLog("CREAR EDIFICIO", "Nombre: " + d.nombre); 
    
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// SUSTITUIR EN Code.gs

function crearActivo(d) {
  try {
    verificarPermiso(['WRITE']);
    
    // LÓGICA ACTIVOS DE CAMPUS (NUEVO)
    // Si viene idCampus pero NO idEdificio, asignamos al edificio "General"
    if (!d.idEdificio && d.idCampus) {
       const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
       const sheetEdif = ss.getSheetByName('EDIFICIOS');
       const datosEdif = sheetEdif.getDataRange().getValues();
       const nombreGeneral = "ZONAS EXTERIORES / GENERAL";
       
       let idEdificioGeneral = null;
       let carpetaPadreId = null;

       // 1. Buscar si ya existe el edificio "General" en este campus
       for(let i=1; i<datosEdif.length; i++) {
         if (String(datosEdif[i][1]) === String(d.idCampus) && 
             String(datosEdif[i][2]).toUpperCase() === nombreGeneral) {
             idEdificioGeneral = datosEdif[i][0];
             carpetaPadreId = datosEdif[i][5]; // Carpeta activos
             break;
         }
       }

       // 2. Si no existe, lo creamos automáticamente
       if (!idEdificioGeneral) {
         Logger.log("Creando edificio virtual General para el campus " + d.idCampus);
         idEdificioGeneral = Utilities.getUuid();
         // Buscamos carpeta del campus para anidar bien
         const campus = getCachedSheetData('CAMPUS');
         let idCarpetaCampus = getRootFolderId();
         for(let c=1; c<campus.length; c++) {
            if(String(campus[c][0]) === String(d.idCampus)) { idCarpetaCampus = campus[c][4]; break; }
         }
         // Crear carpetas
         let fIdEdif = crearCarpeta(nombreGeneral, idCarpetaCampus);
         let fIdActivos = crearCarpeta("Activos", fIdEdif);
         
         sheetEdif.appendRow([
           idEdificioGeneral, d.idCampus, nombreGeneral, "Gestión Campus", fIdEdif, fIdActivos, "", ""
         ]);
         invalidateCache('EDIFICIOS'); // Limpiar caché para que aparezca luego
         carpetaPadreId = fIdActivos;
       }
       
       // Asignamos el ID encontrado/creado
       d.idEdificio = idEdificioGeneral;
       // Pasamos la carpeta padre explícitamente para ahorrar búsqueda
       d._carpetaPadreCache = carpetaPadreId; 
    }

    // --- VALIDACIÓN ESTÁNDAR ---
    if (!d.idEdificio) throw new Error("No has seleccionado un Edificio ni Campus.");
    if (!d.nombre) throw new Error("El nombre es obligatorio.");

    const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
    const eData = getCachedSheetData('EDIFICIOS');
    
    // 1. Buscar carpeta (si no la hemos pre-calculado arriba)
    let pId = d._carpetaPadreCache || null;
    let nombreEdificio = "Desconocido";

    if (!pId) {
      for (let i = 1; i < eData.length; i++) {
        if (String(eData[i][0]) === String(d.idEdificio)) {
          pId = eData[i][5]; 
          nombreEdificio = eData[i][2];
          break;
        }
      }
    }

    // 2. Crear carpeta del activo
    let fId = "";
    try {
      if (!pId) pId = getRootFolderId();
      fId = crearCarpeta(d.nombre, pId);
    } catch (errDrive) {
      fId = "ERROR_DRIVE"; 
    }

    // 3. Resolver nombre Tipo
    const cats = getCachedSheetData('CAT_INSTALACIONES');
    let nombreTipo = d.tipo;
    for (let i = 1; i < cats.length; i++) {
      if (String(cats[i][0]) === String(d.tipo)) { nombreTipo = cats[i][1]; break; }
    }

    // 4. Guardar
    const newId = Utilities.getUuid();
    ss.getSheetByName('ACTIVOS').appendRow([
      newId, d.idEdificio, nombreTipo, d.nombre, d.marca || '', new Date(), fId
    ]);

    SpreadsheetApp.flush(); 
    invalidateCache('ACTIVOS');
    
    return { success: true, newId: newId };

  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function getCatalogoInstalaciones() { return getSheetData('CAT_INSTALACIONES').slice(1).map(r=>({id:r[0], nombre:r[1], dias:r[3]})); }

function getTableData(tipo) {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  
  // --- CAMPUS ---
  if (tipo === 'CAMPUS') {
    const data = getCachedSheetData('CAMPUS');
    return data.slice(1)
      .map(r => ({
        id: r[0], nombre: r[1], provincia: r[2], direccion: r[3]
      }))
      .sort((a, b) => a.nombre.localeCompare(b.nombre, 'es', { sensitivity: 'base' })); // <--- ORDENAR A-Z
  }
  
  // --- EDIFICIOS ---
  if (tipo === 'EDIFICIOS') {
    const data = getCachedSheetData('EDIFICIOS');
    const dataC = getCachedSheetData('CAMPUS');
    
    // ... (Mantén aquí la lógica de contadores que añadimos antes: countActivos, countInc) ...
    // Para simplificar te pongo la versión compacta, pero mantén tus contadores si los tienes
    const dataActivos = getCachedSheetData('ACTIVOS');
    const dataInc = getCachedSheetData('INCIDENCIAS');
    
    const mapCampus = {};
    dataC.slice(1).forEach(r => mapCampus[r[0]] = r[1]);
    
    const countActivos = {};
    dataActivos.slice(1).forEach(r => { const id = String(r[1]); countActivos[id] = (countActivos[id] || 0) + 1; });

    const countInc = {};
    dataInc.slice(1).forEach(r => { if(r[6]!=='RESUELTA') { const id=String(r[2]); if(r[1]==='EDIFICIO' || r[1]==='ACTIVO') countInc[id]=(countInc[id]||0)+1; } }); // Simplificado

    return data.slice(1)
      .map(r => ({
        id: r[0],
        campus: mapCampus[r[1]] || '-',
        nombre: r[2],
        contacto: r[3],
        lat: r[6],
        lng: r[7],
        nActivos: countActivos[String(r[0])] || 0,
        nIncidencias: countInc[String(r[0])] || 0
      }))
      .sort((a, b) => {
         // Primero ordenamos por Campus
         const cmpCampus = a.campus.localeCompare(b.campus, 'es', { sensitivity: 'base' });
         if (cmpCampus !== 0) return cmpCampus;
         // Si es el mismo campus, por nombre de Edificio
         return a.nombre.localeCompare(b.nombre, 'es', { sensitivity: 'base' });
      });
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
    
    // Guardar en la hoja
    sheet.appendRow([Utilities.getUuid(), d.tipoOrigen, d.idOrigen, d.nombreOrigen, d.descripcion, d.prioridad, "PENDIENTE", fecha, usuario, idFoto]);
    
    // --- NUEVO: ENVIAR EMAIL DE ALERTA ---
    // Lo envolvemos en un try/catch interno para que, si falla el email, 
    // la incidencia se guarde igualmente y no dé error al usuario.
    try {
      enviarAlertaIncidencia({
        nombreOrigen: d.nombreOrigen,
        prioridad: d.prioridad,
        descripcion: d.descripcion,
        usuario: usuario,
        fecha: Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")
      });
    } catch(mailErr) {
      console.log("No se pudo enviar el email: " + mailErr);
    }
    // -------------------------------------

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
  const mant = getGlobalMaintenance(); 
  mant.forEach(r => {
    eventos.push({
      id: r.id,
      resourceId: 'MANTENIMIENTO',
      title: `🔧 ${r.tipo} - ${r.activo}`,
      start: r.fechaISO,
      color: r.color === 'rojo' ? '#dc3545' : (r.color === 'amarillo' ? '#ffc107' : '#198754'),
      textColor: r.color === 'amarillo' ? '#000' : '#fff',
      extendedProps: {
        tipo: 'MANTENIMIENTO',
        descripcion: `${r.edificio} | ${r.campusNombre}`,
        status: r.color,
        editable: true
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
          editable: false // Las incidencias son fecha de reporte, no se mueven
        }
      });
    }
  }

  // 4. VENCIMIENTOS DE CONTRATOS (NUEVO)
  const dataCont = getSheetData('CONTRATOS');
  for(let i=1; i<dataCont.length; i++) {
    const estado = dataCont[i][7]; // Col H
    const fFin = dataCont[i][6];   // Col G
    
    // Solo contratos ACTIVOS que tengan fecha de fin
    if(estado === 'ACTIVO' && fFin instanceof Date) {
       eventos.push({
         id: dataCont[i][0],
         resourceId: 'CONTRATO',
         title: `📄 Fin: ${dataCont[i][3]}`, // Proveedor
         start: Utilities.formatDate(fFin, Session.getScriptTimeZone(), "yyyy-MM-dd"),
         color: '#6f42c1', // Morado
         extendedProps: {
           tipo: 'CONTRATO',
           descripcion: `Ref: ${dataCont[i][4]}`,
           editable: true // Permitimos mover la fecha de fin arrastrando
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
  const ss = getDB();
  const nuevaFecha = textoAFecha(nuevaFechaISO); 

  if (tipo === 'MANTENIMIENTO') {
    const sheet = ss.getSheetByName('PLAN_MANTENIMIENTO');
    const data = sheet.getDataRange().getValues();
    for(let i=1; i<data.length; i++){
      if(String(data[i][0]) === String(id)) {
        sheet.getRange(i+1, 5).setValue(nuevaFecha);
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
        sheet.getRange(i+1, 5).setValue(nuevaFecha);
        registrarLog("REPROGRAMAR", `Obra ${id} movida a ${nuevaFechaISO}`);
        return { success: true };
      }
    }
  }
  else if (tipo === 'CONTRATO') { // NUEVO: Actualizar fecha fin contrato
    const sheet = ss.getSheetByName('CONTRATOS');
      const data = sheet.getDataRange().getValues();
      for(let i=1; i<data.length; i++){
        if(String(data[i][0]) === String(id)) {
          sheet.getRange(i+1, 7).setValue(nuevaFecha); // Columna 7 es Fecha Fin
          registrarLog("REPROGRAMAR", `Contrato ${id} extendido/movido a ${nuevaFechaISO}`);
          invalidateCache('CONTRATOS'); // Limpiar caché para refrescar listas
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
  /**
   * Obtiene los años disponibles en el histórico de mantenimiento
   * CORREGIDA: Maneja fechas que vienen como String desde el caché JSON
   */
  function getAniosAuditoria() {
    const planes = getCachedSheetData('PLAN_MANTENIMIENTO');
    const anios = new Set();
    
    // Empezamos en 1 para saltar cabecera
    for(let i=1; i<planes.length; i++) {
      let fecha = planes[i][4]; // Columna Fecha (índice 4)
      
      // --- CORRECCIÓN CLAVE ---
      // Si viene del caché, es un String. Lo convertimos a Date.
      if (typeof fecha === 'string') {
          // Intentar crear fecha. Si es formato ISO, funciona directo.
          const fechaObj = new Date(fecha);
          // Verificar que es una fecha válida
          if (!isNaN(fechaObj.getTime())) {
              fecha = fechaObj;
          }
      }
      // ------------------------

      if (fecha instanceof Date && !isNaN(fecha.getTime())) {
        anios.add(fecha.getFullYear());
      }
    }
    
    // Convertir a array y ordenar descendente (2025, 2024...)
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

// ==========================================
// NOTIFICACIONES POR EMAIL (NUEVO)
// ==========================================

/**
 * Obtiene la lista de emails de usuarios (Admins y Técnicos) que quieren recibir avisos
 */
function getDestinatariosAvisos() {
  const data = getSheetData('USUARIOS');
  const emails = [];
  
  // Empezamos en 1 para saltar cabecera
  for (let i = 1; i < data.length; i++) {
    const email = String(data[i][2]).trim();
    const rol = String(data[i][3]);
    const recibeAvisos = String(data[i][4]); // Columna 'RECIBIR_AVISOS'
    
    // Filtro: Debe tener email, ser ADMIN o TECNICO, y tener AVISOS = SI
    if (email && (rol === 'ADMIN' || rol === 'TECNICO') && recibeAvisos === 'SI') {
      emails.push(email);
    }
  }
  return emails;
}

/**
 * Envía el email formateado
 */
function enviarAlertaIncidencia(datos) {
  const destinatarios = getDestinatariosAvisos();
  
  if (destinatarios.length === 0) {
    console.log("No hay destinatarios configurados para recibir alertas.");
    return;
  }

  // Definir colores según prioridad
  let colorHeader = "#6c757d"; // Gris (Baja)
  if (datos.prioridad === 'MEDIA') colorHeader = "#0d6efd"; // Azul
  if (datos.prioridad === 'ALTA') colorHeader = "#fd7e14";  // Naranja
  if (datos.prioridad === 'URGENTE') colorHeader = "#dc3545"; // Rojo

  const asunto = `[GMAO] Nueva Incidencia: ${datos.nombreOrigen}`;
  
  const htmlBody = `
    <div style="font-family: Arial, sans-serif; color: #333; max-width: 600px; border: 1px solid #ddd; border-radius: 8px; overflow: hidden;">
      <div style="background-color: ${colorHeader}; color: white; padding: 15px; text-align: center;">
        <h2 style="margin: 0;">Nueva Incidencia Reportada</h2>
        <p style="margin: 5px 0 0 0; font-size: 14px;">Prioridad: <strong>${datos.prioridad}</strong></p>
      </div>
      
      <div style="padding: 20px;">
        <p>Se ha registrado una nueva avería en el sistema:</p>
        
        <table style="width: 100%; border-collapse: collapse;">
          <tr>
            <td style="padding: 8px 0; color: #666; width: 120px;"><strong>Ubicación/Activo:</strong></td>
            <td style="padding: 8px 0;">${datos.nombreOrigen}</td>
          </tr>
          <tr>
            <td style="padding: 8px 0; color: #666;"><strong>Reportado por:</strong></td>
            <td style="padding: 8px 0;">${datos.usuario}</td>
          </tr>
          <tr>
            <td style="padding: 8px 0; color: #666;"><strong>Fecha:</strong></td>
            <td style="padding: 8px 0;">${datos.fecha}</td>
          </tr>
          <tr>
            <td style="padding: 8px 0; color: #666; vertical-align: top;"><strong>Descripción:</strong></td>
            <td style="padding: 8px 0; background-color: #f8f9fa; border-radius: 4px; padding: 10px;">${datos.descripcion}</td>
          </tr>
        </table>

        <div style="text-align: center; margin-top: 25px;">
          <a href="${ScriptApp.getService().getUrl()}" style="background-color: #333; color: white; text-decoration: none; padding: 10px 20px; border-radius: 5px; font-size: 14px;">Acceder al GMAO</a>
        </div>
      </div>
      
      <div style="background-color: #f1f1f1; padding: 10px; text-align: center; font-size: 11px; color: #888;">
        GMAO Universidad de Navarra | Notificación Automática
      </div>
    </div>
  `;

  try {
    MailApp.sendEmail({
      to: destinatarios.join(','), // Enviar a todos separados por coma (o usar bcc si son muchos)
      subject: asunto,
      htmlBody: htmlBody
    });
    console.log(`Email de incidencia enviado a ${destinatarios.length} destinatarios.`);
  } catch (e) {
    console.error("Error enviando email: " + e.toString());
  }
}

// ==========================================
// GESTIÓN DE PROVEEDORES
// ==========================================

/**
 * Obtener lista completa de proveedores
 */
// En Code.gs -> function getListaProveedores()

function getListaProveedores() {
  const data = getSheetData('PROVEEDORES');
  
  if (!data || data.length < 2) {
    return [];
  }
  
  const lista = data.slice(1).map(r => ({
    id: r[0],
    nombre: r[1],
    cif: r[2] || '-',
    contacto: r[3] || '-',
    telefono: r[4] || '-',
    email: r[5] || '-',
    activo: r[6] !== 'NO' 
  }));

  // AÑADIR ESTE ORDENAMIENTO:
  return lista.sort((a, b) => a.nombre.localeCompare(b.nombre, 'es', { sensitivity: 'base' }));
}

/**
 * Crear o actualizar proveedor
 */
function saveProveedor(datos) {
  verificarPermiso(['WRITE']);
  
  const ss = getDB();
  let sheet = ss.getSheetByName('PROVEEDORES');
  
  // Si no existe la hoja, crearla
  if (!sheet) {
    sheet = ss.insertSheet('PROVEEDORES');
    sheet.appendRow(['ID', 'NOMBRE', 'CIF', 'CONTACTO', 'TELEFONO', 'EMAIL', 'ACTIVO']);
    sheet.getRange(1, 1, 1, 7).setBackground('#CC0605').setFontColor('#FFFFFF').setFontWeight('bold');
  }
  
  // Validaciones
  if (!datos.nombre || datos.nombre.trim() === '') {
    return { success: false, error: "El nombre es obligatorio" };
  }
  
  const nombre = datos.nombre.trim();
  const cif = (datos.cif || '').trim();
  const contacto = (datos.contacto || '').trim();
  const telefono = (datos.telefono || '').trim();
  const email = (datos.email || '').trim();
  const activo = datos.activo === false ? 'NO' : 'SI';
  
  // EDICIÓN
  if (datos.id) {
    const dataSheet = sheet.getDataRange().getValues();
    
    for (let i = 1; i < dataSheet.length; i++) {
      if (String(dataSheet[i][0]) === String(datos.id)) {
        sheet.getRange(i + 1, 2).setValue(nombre);
        sheet.getRange(i + 1, 3).setValue(cif);
        sheet.getRange(i + 1, 4).setValue(contacto);
        sheet.getRange(i + 1, 5).setValue(telefono);
        sheet.getRange(i + 1, 6).setValue(email);
        sheet.getRange(i + 1, 7).setValue(activo);
        
        registrarLog("EDITAR PROVEEDOR", `${nombre} (${cif})`);
        
        return { success: true };
      }
    }
    
    return { success: false, error: "Proveedor no encontrado" };
  }
  
  // CREACIÓN
  const newId = Utilities.getUuid();
  sheet.appendRow([newId, nombre, cif, contacto, telefono, email, activo]);
  
  registrarLog("CREAR PROVEEDOR", `${nombre} (${cif})`);
  
  return { success: true, newId: newId };
}

/**
 * Eliminar proveedor (solo si no tiene contratos activos)
 */
function eliminarProveedor(id) {
  verificarPermiso(['DELETE']);
  
  const ss = getDB();
  
  // Verificar que no tenga contratos activos
  const contratos = getSheetData('CONTRATOS');
  const tieneContratos = contratos.slice(1).some(r => 
    String(r[3]) === String(id) || // Si el ID del proveedor coincide (antigua columna)
    obtenerNombreProveedor(r[3]) === obtenerNombreProveedor(id) // O si el nombre coincide
  );
  
  if (tieneContratos) {
    return { 
      success: false, 
      error: "No se puede eliminar: tiene contratos asociados. Márcalo como inactivo en su lugar." 
    };
  }
  
  const sheet = ss.getSheetByName('PROVEEDORES');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      const nombre = data[i][1];
      sheet.deleteRow(i + 1);
      
      registrarLog("ELIMINAR PROVEEDOR", nombre);
      
      return { success: true };
    }
  }
  
  return { success: false, error: "Proveedor no encontrado" };
}

/**
 * Helper: Obtener nombre de proveedor por ID
 */
function obtenerNombreProveedor(id) {
  const proveedores = getSheetData('PROVEEDORES');
  
  for (let i = 1; i < proveedores.length; i++) {
    if (String(proveedores[i][0]) === String(id)) {
      return proveedores[i][1];
    }
  }
  
  return id; // Si no se encuentra, devolver el ID tal cual (puede ser un nombre antiguo)
}

/**
 * Buscar proveedor por nombre (para autocompletar)
 */
function buscarProveedores(texto) {
  if (!texto || texto.length < 2) {
    return [];
  }
  
  const proveedores = getListaProveedores();
  const textoLower = texto.toLowerCase();
  
  return proveedores
    .filter(p => p.activo) // Solo activos
    .filter(p => 
      p.nombre.toLowerCase().includes(textoLower) ||
      (p.cif && p.cif.toLowerCase().includes(textoLower))
    )
    .slice(0, 10) // Máximo 10 resultados
    .map(p => ({
      id: p.id,
      nombre: p.nombre,
      cif: p.cif,
      displayText: p.cif ? `${p.nombre} (${p.cif})` : p.nombre
    }));
}

// --- PEGAR ESTO AL FINAL DE Code.gs ---

function getContratoFullDetailsV2(idContrato) {
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const sheet = ss.getSheetByName('CONTRATOS');
  const data = sheet.getDataRange().getValues();
  
  let contrato = null;
  
  for(let i=1; i<data.length; i++){
    if(String(data[i][0]) === String(idContrato)) {
      
      const tipoEntidad = data[i][1];
      const idEntidadRaw = data[i][2];
      let idsActivos = [];

      // Si es de tipo "ACTIVOS", parseamos el JSON
      if (tipoEntidad === 'ACTIVOS') {
        try { idsActivos = JSON.parse(idEntidadRaw); } catch(e) {}
      }

      contrato = {
        id: data[i][0],
        tipoEntidad: tipoEntidad,
        idEntidad: idEntidadRaw,
        idsActivos: idsActivos,
        idProveedor: data[i][3], // Columna D (ID Proveedor)
        ref: data[i][4],         // Columna E
        inicio: data[i][5] ? Utilities.formatDate(data[i][5], Session.getScriptTimeZone(), "yyyy-MM-dd") : "",
        fin: data[i][6] ? Utilities.formatDate(data[i][6], Session.getScriptTimeZone(), "yyyy-MM-dd") : "",
        estado: data[i][7]
      };
      break;
    }
  }
  
  if(!contrato) throw new Error("Contrato no encontrado");

  // Resolver jerarquía para rellenar los selects
  let idCampus = "";
  let idEdificio = "";
  let idActivo = "";

  if (contrato.tipoEntidad === 'CAMPUS') {
    idCampus = contrato.idEntidad;
  } 
  else if (contrato.tipoEntidad === 'EDIFICIO') {
    idEdificio = contrato.idEntidad;
    const edifs = getSheetData('EDIFICIOS');
    for(let i=1; i<edifs.length; i++) {
       if(String(edifs[i][0]) === String(idEdificio)) { idCampus = edifs[i][1]; break; }
    }
  } 
  else if (contrato.tipoEntidad === 'ACTIVO') {
    idActivo = contrato.idEntidad;
    const activos = getSheetData('ACTIVOS');
    const edifs = getSheetData('EDIFICIOS');
    for(let i=1; i<activos.length; i++) {
       if(String(activos[i][0]) === String(idActivo)) { idEdificio = activos[i][1]; break; }
    }
    if (idEdificio) {
      for(let i=1; i<edifs.length; i++) {
         if(String(edifs[i][0]) === String(idEdificio)) { idCampus = edifs[i][1]; break; }
      }
    }
  }
  else if (contrato.tipoEntidad === 'ACTIVOS' && contrato.idsActivos.length > 0) {
    const primerId = contrato.idsActivos[0];
    const activos = getSheetData('ACTIVOS');
    const edifs = getSheetData('EDIFICIOS');
    
    for(let i=1; i<activos.length; i++) {
       if(String(activos[i][0]) === String(primerId)) { idEdificio = activos[i][1]; break; }
    }
    if (idEdificio) {
      for(let i=1; i<edifs.length; i++) {
         if(String(edifs[i][0]) === String(idEdificio)) { idCampus = edifs[i][1]; break; }
      }
    }
  }

  return {
    ...contrato,
    idCampus: idCampus,
    idEdificio: idEdificio,
    idActivo: idActivo
  };
}

function crearContratoV2(d) {
  verificarPermiso(['WRITE']);
  const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID'));
  const sheet = ss.getSheetByName('CONTRATOS');

  // Determinar ID Entidad (Si es array de activos, lo convertimos a JSON string)
  let idEntidadFinal = d.idEntidad;
  let tipoEntidadFinal = d.tipoEntidad;

  if (d.idsActivos && d.idsActivos.length > 0) {
      tipoEntidadFinal = 'ACTIVOS';
      idEntidadFinal = JSON.stringify(d.idsActivos);
  }

  // Columnas: [ID, TIPO, ID_ENTIDAD, ID_PROVEEDOR, REF, INICIO, FIN, ESTADO]
  const newId = Utilities.getUuid();
  sheet.appendRow([
      newId,
      tipoEntidadFinal,
      idEntidadFinal,
      d.idProveedor,
      d.ref,
      textoAFecha(d.fechaIni),
      textoAFecha(d.fechaFin),
      d.estado
  ]);
  
  registrarLog("CREAR CONTRATO", "Ref: " + d.ref);
  invalidateCache('CONTRATOS');
  return { success: true, newId: newId };
}

function DIAGNOSTICO_CONTRATOS() {
  const sheetC = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')).getSheetByName('CONTRATOS');
  const datosC = sheetC.getDataRange().getValues();
  
  const sheetP = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')).getSheetByName('PROVEEDORES');
  const datosP = sheetP.getDataRange().getValues();

  Logger.log("=== 1. ANALIZANDO PROVEEDORES ===");
  const mapaProv = {};
  for(let i=1; i<datosP.length; i++) {
    const id = String(datosP[i][0]).trim();
    const nombre = datosP[i][1];
    mapaProv[id] = nombre;
    Logger.log(`Prov [${i}]: ID="${id}" | Nombre="${nombre}"`);
  }

  Logger.log("=== 2. ANALIZANDO CONTRATOS (Primeros 5) ===");
  // Empezamos en 1 (saltar cabecera)
  for(let i=1; i<Math.min(datosC.length, 6); i++) {
    const fila = datosC[i];
    const tipo = fila[1]; // Columna B
    const entidad = fila[2]; // Columna C (El JSON o ID)
    const provID = String(fila[3]).trim(); // Columna D (El ID del proveedor)

    Logger.log(`--- Contrato Fila ${i+1} ---`);
    Logger.log(`Tipo: ${tipo}`);
    
    // Test Proveedor
    const nombreProv = mapaProv[provID];
    if (nombreProv) {
      Logger.log(`✅ Proveedor ENCONTRADO: ID="${provID}" -> "${nombreProv}"`);
    } else {
      Logger.log(`❌ Proveedor NO encontrado en el mapa. ID en contrato: "${provID}"`);
    }

    // Test Entidad (JSON)
    Logger.log(`Entidad Raw (Tipo: ${typeof entidad}): ${entidad}`);
    if (tipo === 'ACTIVOS') {
      if (Array.isArray(entidad)) {
         Logger.log(`ℹ️ Google ya lo convirtió en Array. Longitud: ${entidad.length}`);
         Logger.log(`Primer ID: ${entidad[0]}`);
      } else {
         Logger.log(`ℹ️ Es texto. Intentando parsear...`);
         try {
           const parsed = JSON.parse(entidad);
           Logger.log(`✅ Parseo OK. Primer ID: ${parsed[0]}`);
         } catch(e) {
           Logger.log(`❌ ERROR PARSEO JSON: ${e.message}`);
         }
      }
    }
  }
}

function getManualHTML() {
  // Aquí puedes guardar el manual en HTML
  // Por ahora lo devolvemos como string, pero podrías guardarlo en una hoja de cálculo
  
  return `
    <div class="manual-toc">
      <h3><i class="bi bi-list-ul me-2"></i>Índice de Contenidos</h3>
      <ul>
        <li><a href="#intro">1. Introducción</a></li>
        <li><a href="#acceso">2. Acceso al Sistema</a></li>
        <li><a href="#permisos">3. Permisos y Roles</a></li>
        <li><a href="#navegacion">4. Navegación Principal</a></li>
        <li><a href="#campus">5. Gestión de Campus y Edificios</a></li>
        <li><a href="#activos">6. Gestión de Activos</a></li>
        <li><a href="#mantenimiento">7. Plan de Mantenimiento</a></li>
        <li><a href="#contratos">8. Gestión de Contratos</a></li>
        <li><a href="#incidencias">9. Incidencias</a></li>
        <li><a href="#qr">10. Códigos QR</a></li>
        <li><a href="#adicionales">11. Funciones Adicionales</a></li>
        <li><a href="#soporte">12. Soporte</a></li>
      </ul>
    </div>

    <section id="intro" class="manual-section">
      <h1>📋 Introducción</h1>
      <p>El <strong>GMAO (Sistema de Gestión de Mantenimiento Asistido por Ordenador)</strong> de la Universidad de Navarra es una aplicación web diseñada para gestionar de forma integral todos los activos, instalaciones y mantenimientos de los diferentes campus universitarios.</p>
      
      <h3>Características Principales</h3>
      <ul>
        <li>✅ Gestión centralizada de activos e instalaciones</li>
        <li>✅ Programación automática de mantenimientos legales y periódicos</li>
        <li>✅ Control de contratos con proveedores</li>
        <li>✅ Reportes de incidencias con fotografías</li>
        <li>✅ Códigos QR para identificación rápida</li>
        <li>✅ Alertas automáticas por email</li>
        <li>✅ Historial completo de documentación</li>
      </ul>
    </section>

    <section id="acceso" class="manual-section">
      <h1>🔐 Acceso al Sistema</h1>
      
      <h3>Primer Acceso</h3>
      <ol>
        <li>Abra el enlace proporcionado por el administrador</li>
        <li>Inicie sesión con su cuenta de Google corporativa (@unav.es)</li>
        <li>El sistema detectará automáticamente sus permisos</li>
      </ol>

      <h3>Interfaz Principal</h3>
      <p>La aplicación se divide en tres zonas:</p>
      
      <div class="alert-box info">
        <strong>💡 Barra Lateral Izquierda (Menú)</strong><br>
        Dashboard, Campus, Edificios, Activos, Mantenimiento, Incidencias, Contratos, Planificador, Proveedores, Configuración
      </div>
      
      <div class="alert-box success">
        <strong>📱 Botones Flotantes (esquina inferior derecha)</strong><br>
        • 💬 Botón azul: Enviar sugerencias/reportar errores<br>
        • 🔴 Botón rojo: Reportar avería urgente
      </div>
    </section>

    <section id="permisos" class="manual-section">
      <h1>👥 Permisos y Roles</h1>
      <p>El sistema tiene tres niveles de acceso:</p>

      <h3>👁️ CONSULTA (Solo lectura)</h3>
      <ul>
        <li>✅ Ver toda la información</li>
        <li>✅ Descargar documentos</li>
        <li>✅ Reportar averías</li>
        <li>❌ No puede crear ni modificar</li>
      </ul>

      <h3>🔧 TÉCNICO (Operativo)</h3>
      <ul>
        <li>✅ Todo lo de Consulta</li>
        <li>✅ Crear y editar activos</li>
        <li>✅ Programar mantenimientos</li>
        <li>✅ Subir documentación</li>
        <li>✅ Gestionar contratos</li>
        <li>❌ No puede eliminar registros</li>
        <li>❌ No puede gestionar usuarios</li>
      </ul>

      <h3>👑 ADMINISTRADOR (Control total)</h3>
      <ul>
        <li>✅ Acceso completo</li>
        <li>✅ Eliminar registros</li>
        <li>✅ Gestionar usuarios</li>
        <li>✅ Configurar catálogo de instalaciones</li>
        <li>✅ Ver logs de auditoría</li>
      </ul>
    </section>

    <section id="navegacion" class="manual-section">
      <h1>📊 Dashboard</h1>
      
      <p>El <strong>Dashboard</strong> muestra un resumen general del estado del sistema.</p>
      
      <h3>Tarjetas de Estado</h3>
      <ul>
        <li><strong>Activos:</strong> Número total de equipos registrados</li>
        <li><strong>Vencidas:</strong> Revisiones que no se han realizado a tiempo (🔴 rojo)</li>
        <li><strong>Pendientes:</strong> Revisiones próximas a vencer en 30 días (🟡 amarillo)</li>
        <li><strong>Incidencias:</strong> Averías sin resolver</li>
        <li><strong>Contratos:</strong> Contratos vigentes</li>
      </ul>

      <div class="alert-box success">
        <strong>💡 Truco:</strong> Haga clic en cualquier tarjeta para ir directamente a esa sección
      </div>

      <h3>Calendario</h3>
      <ul>
        <li><strong style="color: #10b981;">Verde:</strong> Mantenimiento al día</li>
        <li><strong style="color: #f59e0b;">Amarillo:</strong> Próximo a vencer (≤30 días)</li>
        <li><strong style="color: #ef4444;">Rojo:</strong> Vencido</li>
      </ul>
    </section>

    <section id="campus" class="manual-section">
      <h1>🏛️ Campus y Edificios</h1>
      
      <h2>Gestión de Campus</h2>
      
      <h3>Crear Nuevo Campus</h3>
      <ol>
        <li>Clic en <strong>"+ Nuevo Campus"</strong></li>
        <li>Complete los campos:
          <ul>
            <li><strong>Nombre:</strong> Ej. "Campus de Pamplona"</li>
            <li><strong>Provincia:</strong> Ej. "Navarra"</li>
            <li><strong>Dirección:</strong> Dirección completa</li>
          </ul>
        </li>
        <li>Clic en <strong>"Guardar"</strong></li>
      </ol>

      <div class="alert-box warning">
        <strong>⚠️ Importante:</strong> Al crear un campus, se crea automáticamente una carpeta en Google Drive
      </div>

      <h2>Gestión de Edificios</h2>
      
      <h3>Crear Nuevo Edificio</h3>
      <ol>
        <li>Clic en <strong>"+ Nuevo Edificio"</strong></li>
        <li>Complete los datos requeridos</li>
        <li>Opcionalmente, añada coordenadas para visualización en mapa</li>
      </ol>

      <div class="alert-box info">
        <strong>💡 Para obtener coordenadas:</strong><br>
        1. Abra Google Maps<br>
        2. Clic derecho sobre el edificio<br>
        3. Copie las coordenadas que aparecen
      </div>

      <h3>Ficha de Edificio</h3>
      <p>Al hacer clic en el ojo (👁️) de un edificio, accede a:</p>
      <ul>
        <li><strong>📋 Información:</strong> Datos básicos</li>
        <li><strong>📁 Documentación:</strong> Licencias, planos, certificados</li>
        <li><strong>🏗️ Obras:</strong> Historial de reformas</li>
        <li><strong>🔧 Activos:</strong> Lista de equipos instalados</li>
      </ul>
    </section>

    <section id="activos" class="manual-section">
      <h1>📦 Gestión de Activos</h1>
      
      <p>Los <strong>activos</strong> son todos los equipos e instalaciones: calderas, ascensores, aire acondicionado, cuadros eléctricos, etc.</p>

      <h2>Crear Nuevo Activo</h2>
      
      <h3>Método Manual</h3>
      <ol>
        <li>Clic en <strong>"+ Crear Activo"</strong></li>
        <li>Seleccione ubicación (Campus + Edificio)</li>
        <li>Elija tipo del catálogo</li>
        <li>Asigne nombre único</li>
        <li>Indique marca/fabricante</li>
      </ol>

      <h3>Método Masivo (Importación)</h3>
      <p>Para dar de alta muchos activos a la vez:</p>
      <ol>
        <li>Clic en <strong>"Importar"</strong></li>
        <li>Prepare sus datos en Excel: <code>Campus | Edificio | Tipo | Nombre | Marca</code></li>
        <li>Copie las filas (sin cabeceras)</li>
        <li>Pegue en el cuadro de texto</li>
        <li>Clic en <strong>"Procesar Importación"</strong></li>
      </ol>

      <h2>Ficha Completa de Activo</h2>
      
      <h3>📁 Documentación</h3>
      <p>Aquí se guardan manuales, certificados, fichas técnicas, etc.</p>
      
      <div class="alert-box success">
        <strong>📤 Subida Rápida (botón nube ☁️):</strong><br>
        • Permite subir varios archivos a la vez<br>
        • Clasifica automáticamente OCAs y contratos<br>
        • Programa revisiones futuras automáticamente
      </div>

      <h3>🔧 Mantenimiento</h3>
      <p>Programar nueva revisión:</p>
      <ol>
        <li>Clic en <strong>"+ Programar Revisión"</strong></li>
        <li>Seleccione tipo (Legal, Periódica, Reparación, Extraordinaria)</li>
        <li>Si es Legal, elija normativa (se autocompleta frecuencia)</li>
        <li>Active "Repetir" para crear futuras automáticamente</li>
        <li>Adjunte evidencia si ya la tiene</li>
        <li>Opcionalmente sincronice con Google Calendar</li>
      </ol>

      <h3>Marcar Revisión como Realizada</h3>
      <ol>
        <li>Clic en botón ✅ (check verde)</li>
        <li>Confirmar</li>
      </ol>
      <p>La revisión pasará a "Histórico" (azul) y no aparecerá en alertas.</p>
    </section>

    <section id="mantenimiento" class="manual-section">
      <h1>🔧 Plan de Mantenimiento</h1>
      
      <p>Vista global con <strong>todas las revisiones programadas</strong> de todos los activos.</p>

      <h2>Sistema de Filtros</h2>
      
      <h3>Filtros de Ubicación</h3>
      <ul>
        <li><strong>Campus:</strong> Filtra por campus específico</li>
        <li><strong>Edificio:</strong> Filtra por edificio</li>
      </ul>

      <h3>Filtros de Estado</h3>
      <ul>
        <li><strong>Todas:</strong> Muestra todas excepto históricas</li>
        <li><strong>Vencidas (🔴):</strong> Ya pasaron su fecha</li>
        <li><strong>Próximas (🟡):</strong> Vencen en ≤30 días</li>
        <li><strong>Al día (🟢):</strong> Bien de fecha</li>
        <li><strong>Histórico (🔵):</strong> Ya realizadas</li>
      </ul>

      <h2>Informe PDF</h2>
      <p>Clic en <strong>"Informe Legal PDF"</strong> genera un documento con todas las revisiones reglamentarias, ideal para auditorías externas.</p>
    </section>

    <section id="contratos" class="manual-section">
      <h1>📑 Gestión de Contratos</h1>
      
      <h2>Estados de Contratos</h2>
      <ul>
        <li>🟢 <strong>Vigente:</strong> Contrato activo</li>
        <li>🟡 <strong>Próximo:</strong> Caduca en ≤90 días</li>
        <li>🔴 <strong>Caducado:</strong> Ya venció</li>
        <li>⚪ <strong>Inactivo:</strong> Desactivado manualmente</li>
      </ul>

      <h2>Crear Nuevo Contrato</h2>
      
      <h3>Paso 1: Proveedor</h3>
      <p>Seleccione de la lista o clic en <strong>"+ Nuevo"</strong> para crear uno</p>

      <h3>Paso 2: Ubicación (¿A qué aplica?)</h3>
      <ul>
        <li><strong>Todo el Campus:</strong> No seleccione edificio ni activo</li>
        <li><strong>Todo un Edificio:</strong> Seleccione edificio, no activo</li>
        <li><strong>Un Activo Concreto:</strong> Seleccione hasta activo específico</li>
        <li><strong>Varios Activos:</strong> Active casilla y seleccione múltiples</li>
      </ul>

      <h3>Paso 3: Datos del Contrato</h3>
      <ul>
        <li>Referencia/Nº de contrato</li>
        <li>Fechas de inicio y fin</li>
        <li>Estado (Activo/Inactivo)</li>
      </ul>

      <h3>Paso 4: Adjuntar PDF</h3>
      <p>Suba el documento del contrato firmado (opcional)</p>
    </section>

    <section id="incidencias" class="manual-section">
      <h1>⚠️ Sistema de Incidencias</h1>
      
      <h2>Reportar una Avería</h2>
      
      <h3>Botón Flotante Rojo (Acceso Rápido)</h3>
      <ol>
        <li>Clic en botón 🔴 (esquina inferior derecha)</li>
        <li>Complete:
          <ul>
            <li>Ubicación (Campus, Edificio, Activo)</li>
            <li>Descripción del problema</li>
            <li>Prioridad: Baja, Media, Alta, ¡Urgente!</li>
            <li>Foto (opcional)</li>
          </ul>
        </li>
        <li>Clic en <strong>"Enviar Reporte"</strong></li>
      </ol>

      <div class="alert-box warning">
        <strong>📧 Notificación automática:</strong> Se envía email a todos los técnicos y administradores con avisos activados
      </div>

      <h2>Estados de Incidencias</h2>
      <ul>
        <li>🔴 <strong>Pendiente:</strong> Recién creada</li>
        <li>🔵 <strong>En Proceso:</strong> Ya se está trabajando</li>
        <li>🟢 <strong>Resuelta:</strong> Cerrada</li>
      </ul>
    </section>

    <section id="qr" class="manual-section">
      <h1>📱 Códigos QR</h1>
      
      <p>Los códigos QR permiten acceso instantáneo a la ficha de un activo desde el móvil.</p>

      <h2>Generar QR Individual</h2>
      <ol>
        <li>Entre en la ficha del activo</li>
        <li>Clic en <strong>"Descargar QR"</strong></li>
        <li>Se descarga imagen PNG</li>
        <li>Imprima y pegue en el equipo físico</li>
      </ol>

      <h2>Generar QR de Edificio Completo (PDF)</h2>
      <ol>
        <li>Desde ficha del edificio</li>
        <li>Botón <strong>"QR Edificio (PDF)"</strong></li>
        <li>Se genera PDF con etiquetas de todos los activos</li>
        <li>Listo para imprimir (2 columnas/página)</li>
      </ol>

      <h2>Escanear un QR</h2>
      <ol>
        <li>Abra cámara del móvil</li>
        <li>Enfoque el código QR</li>
        <li>Toque el enlace</li>
        <li>Se abre Vista Móvil Optimizada con:
          <ul>
            <li>Datos del activo</li>
            <li>Estado de mantenimiento</li>
            <li>Botones: Reportar avería, Realizar revisión, Ver manuales</li>
          </ul>
        </li>
      </ol>
    </section>

    <section id="adicionales" class="manual-section">
      <h1>➕ Funciones Adicionales</h1>
      
      <h2>📅 Planificador</h2>
      <p>Vista de calendario unificada con:</p>
      <ul>
        <li>🔧 Mantenimientos programados</li>
        <li>🏗️ Obras en curso</li>
        <li>⚠️ Incidencias pendientes</li>
        <li>📄 Vencimientos de contratos</li>
      </ul>
      <p><strong>Función de arrastrar:</strong> Puede cambiar fechas arrastrando eventos en el calendario</p>

      <h2>📊 Auditoría (Exportación Masiva)</h2>
      <ol>
        <li>Seleccione año</li>
        <li>Seleccione tipo (Solo Legales o Todo)</li>
        <li>Clic en <strong>"Generar Paquete"</strong></li>
      </ol>
      <p>Se crea carpeta en Drive con copia de todos los certificados, archivos renombrados automáticamente.</p>

      <h2>🔔 Alertas Automáticas</h2>
      <p>Si tiene activadas las alertas, recibirá emails diarios con:</p>
      <ul>
        <li>⚠️ Revisiones vencidas</li>
        <li>📅 Revisiones próximas (≤7 días)</li>
        <li>📄 Contratos próximos a caducar (≤60 días)</li>
      </ul>

      <h2>💬 Buzón de Sugerencias</h2>
      <p>Botón flotante azul para enviar:</p>
      <ul>
        <li>💡 Ideas de mejora</li>
        <li>🐛 Reportes de errores</li>
        <li>💬 Comentarios generales</li>
      </ul>
    </section>

    <section id="soporte" class="manual-section">
      <h1>📞 Soporte y Contacto</h1>
      
      <div class="alert-box info">
        <strong>Administrador del Sistema:</strong><br>
        Email: jcsuarez@unav.es<br>
        Departamento: Servicio de Obras y Mantenimiento
      </div>

      <h3>Para solicitar:</h3>
      <ul>
        <li>✅ Cambio de permisos</li>
        <li>✅ Alta de nuevos usuarios</li>
        <li>✅ Resolución de incidencias técnicas</li>
        <li>✅ Formación adicional</li>
      </ul>

      <h2>🆘 Resolución de Problemas</h2>
      
      <h3>El sistema no carga</h3>
      <ol>
        <li>Verifique conexión a internet</li>
        <li>Cierre y vuelva a abrir la pestaña</li>
        <li>Borre caché del navegador</li>
        <li>Contacte con administrador si persiste</li>
      </ol>

      <h3>No puedo crear/editar</h3>
      <p>Probablemente tiene permisos de Solo Consulta. Contacte con administrador.</p>

      <h3>No encuentro un activo</h3>
      <ol>
        <li>Seleccione Campus</li>
        <li>Seleccione Edificio</li>
        <li>Use el buscador (mínimo 3 caracteres)</li>
      </ol>
    </section>

    <section class="manual-section">
      <h1>📝 Consejos Finales</h1>
      
      <div class="alert-box success">
        <ul style="margin-bottom: 0; padding-left: 20px;">
          <li>✅ <strong>Use los códigos QR:</strong> Ahorra muchísimo tiempo en campo</li>
          <li>✅ <strong>Suba las OCAs con función rápida:</strong> Programa automáticamente siguientes revisiones</li>
          <li>✅ <strong>Active alertas por email:</strong> No se le pasará ningún mantenimiento</li>
          <li>✅ <strong>Use el Planificador:</strong> Visión global de toda la carga de trabajo</li>
          <li>✅ <strong>Reporte todas las averías:</strong> Ayuda a detectar patrones</li>
          <li>✅ <strong>Revise Dashboard regularmente:</strong> Los números en rojo necesitan atención urgente</li>
        </ul>
      </div>

      <hr style="margin: 30px 0; border-color: #e5e7eb;">
      
      <p style="text-align: center; color: #9ca3af; font-size: 0.9rem;">
        <strong>Versión del Manual:</strong> 1.0 | 
        <strong>Última Actualización:</strong> Diciembre 2025<br>
        <strong>Sistema GMAO</strong> - Universidad de Navarra
      </p>
    </section>
  `;
}

function generarPDFManual() {
  const contenido = getManualHTML();
  
  // Envolvemos el contenido en una estructura HTML completa con estilos para PDF
  const html = `
    <html>
      <head>
        <style>
          body { font-family: 'Helvetica', sans-serif; font-size: 10pt; color: #333; line-height: 1.5; }
          h1 { color: #CC0605; font-size: 18pt; border-bottom: 2px solid #CC0605; padding-bottom: 5px; margin-top: 20px; }
          h2 { color: #1f2937; font-size: 14pt; margin-top: 15px; background-color: #f3f4f6; padding: 5px; }
          h3 { color: #4b5563; font-size: 12pt; border-left: 4px solid #CC0605; padding-left: 10px; }
          .alert-box { padding: 10px; border: 1px solid #ddd; background-color: #f9fafb; margin: 10px 0; border-radius: 4px; font-size: 9pt; }
          ul, ol { margin-bottom: 10px; }
          li { margin-bottom: 5px; }
          .manual-toc { page-break-after: always; background-color: #f8f9fa; padding: 20px; }
          a { text-decoration: none; color: #333; }
        </style>
      </head>
      <body>
        <div style="text-align: center; margin-bottom: 30px;">
          <h1 style="border:none; font-size: 24pt;">Manual de Usuario GMAO</h1>
          <p style="color: #666;">Universidad de Navarra</p>
        </div>
        ${contenido}
      </body>
    </html>
  `;

  const blob = Utilities.newBlob(html, MimeType.HTML).getAs(MimeType.PDF);
  blob.setName("Manual_Usuario_GMAO.pdf");

  return {
    base64: Utilities.base64Encode(blob.getBytes()),
    filename: "Manual_Usuario_GMAO.pdf"
  };
}
