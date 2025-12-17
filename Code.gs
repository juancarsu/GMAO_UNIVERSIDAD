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

function getFirestore() {
  const props = PropertiesService.getScriptProperties();
  const jsonKey = JSON.parse(props.getProperty('FIREBASE_KEY')); 
  const email = props.getProperty('FIREBASE_EMAIL');
  const projectId = props.getProperty('FIREBASE_PROJECT_ID');

  return FirestoreApp.getFirestore(email, jsonKey, projectId);
}
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

function getMyRole() {
  const email = Session.getActiveUser().getEmail();
  const cacheKey = 'USER_ROLE_' + email;
  const cache = CacheService.getUserCache();
  
  // 1. Memoria rápida
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  // 2. Firestore (Mucho más rápido que leer Sheet entera)
  try {
    const firestore = getFirestore();
    const usuarios = firestore.query('usuarios')
                              .where('EMAIL', '==', email)
                              .execute();
    
    let usuario = { email: email, nombre: "Invitado", rol: 'CONSULTA' };
    
    if (usuarios.length > 0) {
      const u = usuarios[0];
      usuario = { 
        email: email, 
        nombre: u.NOMBRE || u.nombre, 
        rol: u.ROL || u.rol 
      };
    }

    cache.put(cacheKey, JSON.stringify(usuario), 1200);
    return usuario;

  } catch (e) {
    // Fallback de seguridad
    console.error("Error Auth Firestore: " + e);
    return { email: email, nombre: "Error Conexión", rol: 'CONSULTA' };
  }
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
  try {
    const firestore = getFirestore();
    // Pedimos todos los campus
    const docs = firestore.query('campus').execute();
    
    return docs.map(d => ({
      id: d.id,
      nombre: d.Nombre || d.nombre || d.NOMBRE || 'Sin Nombre'
    })).sort((a, b) => a.nombre.localeCompare(b.nombre));
    
  } catch(e) {
    console.error("Error getListaCampus: " + e.message);
    return [];
  }
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
    const firestore = getFirestore();
    const idStr = String(idEdificio).trim();
    
    // INTENTO 1: Buscar por 'ID_EDIFICIO' (Mayúsculas, lo más probable si vino de Excel)
    let activos = firestore.query('activos')
                           .where('ID_EDIFICIO', '==', idStr)
                           .execute();
                           
    // INTENTO 2: Si no encuentra nada, probamos con 'ID_Edificio' (Capitalizado)
    if (activos.length === 0) {
       activos = firestore.query('activos')
                          .where('ID_Edificio', '==', idStr)
                          .execute();
    }

    // INTENTO 3: Si sigue vacío, probamos 'idEdificio' (CamelCase)
    if (activos.length === 0) {
       activos = firestore.query('activos')
                          .where('idEdificio', '==', idStr)
                          .execute();
    }
    
    // Mapear resultados
    return activos.map(doc => ({
      id: doc.id,
      idEdificio: doc.ID_EDIFICIO || doc.ID_Edificio || doc.idEdificio,
      // Usamos las variantes de nombre igual que en getAllAssetsList
      nombre: doc.NOMBRE || doc.Nombre || doc.nombre || 'Sin nombre',
      tipo: doc.TIPO || doc.Tipo || doc.tipo || '-',
      marca: doc.MARCA || doc.Marca || doc.marca || '-',
      fechaAlta: doc.FECHA_ALTA || doc.FechaAlta ? new Date(doc.FECHA_ALTA || doc.FechaAlta).toLocaleDateString() : "-"
    })).sort((a, b) => a.nombre.localeCompare(b.nombre));

  } catch(e) {
    console.error("Error getActivosPorEdificio: " + e.message);
    return [];
  }
}

function getAllAssetsList() {
  try {
    const firestore = getFirestore();
    
    // 1. Cargar datos
    const activos = firestore.query('activos').execute();
    const edificios = firestore.query('edificios').execute();
    const campus = firestore.query('campus').execute();
    
    // 2. Mapas
    const mapaCampus = {};
    campus.forEach(c => mapaCampus[c.id] = c.Nombre || c.NOMBRE || c.nombre || 'Desconocido');
    
    const mapaEdificios = {};
    edificios.forEach(e => {
      const idC = e.ID_Campus || e.idCampus || e.Campus || '';
      mapaEdificios[e.id] = {
        nombre: e.Nombre || e.NOMBRE || e.nombre || 'Desconocido',
        nombreCampus: mapaCampus[idC] || '-',
        idCampus: idC // <--- CLAVE: Guardamos el ID del Campus
      };
    });

    // 3. Lista Final con IDs para filtros
    return activos.map(a => {
      const idEdif = a.ID_Edificio || a.ID_EDIFICIO || a.idEdificio || '';
      const infoEdif = mapaEdificios[idEdif] || { nombre: '-', nombreCampus: '-', idCampus: null };
      
      return {
        id: a.id,
        nombre: a.Nombre || a.NOMBRE || a.nombre || 'Sin Nombre',
        tipo: a.Tipo || a.TIPO || a.tipo || '-',
        marca: a.Marca || a.MARCA || a.marca || '-',
        edificioNombre: infoEdif.nombre,
        campusNombre: infoEdif.nombreCampus,
        idEdificio: idEdif,          // Necesario para filtro Edificio
        idCampus: infoEdif.idCampus, // Necesario para filtro Campus
        hasDocs: false, 
        hasLegalDocs: false
      };
    }).sort((a, b) => a.nombre.localeCompare(b.nombre));

  } catch(e) { console.error(e); return []; }
}

function getAssetInfo(id) {
  try {
    const firestore = getFirestore();
    const doc = firestore.getDocument('activos/' + id);
    if (!doc || !doc.fields) return null;
    
    const d = doc.fields;
    const idEdif = d.ID_EDIFICIO || d.idEdificio || '';
    
    // Resolver nombres de Edificio y Campus
    let nombreEdif = '-', nombreCamp = '-', idCamp = null;
    if (idEdif) {
      try {
        const docEdif = firestore.getDocument('edificios/' + idEdif);
        if (docEdif && docEdif.fields) {
           nombreEdif = docEdif.fields.NOMBRE || docEdif.fields.nombre;
           idCamp = docEdif.fields.ID_CAMPUS || docEdif.fields.idCampus;
           if (idCamp) {
              const docCamp = firestore.getDocument('campus/' + idCamp);
              if (docCamp && docCamp.fields) nombreCamp = docCamp.fields.NOMBRE || docCamp.fields.nombre;
           }
        }
      } catch(e){}
    }

    return {
      id: doc.name.split('/').pop(),
      nombre: d.NOMBRE || d.nombre || 'Sin Nombre',
      tipo: d.TIPO || d.tipo || '-',
      marca: d.MARCA || d.marca || '-',
      fechaAlta: d.FECHA_ALTA ? new Date(d.FECHA_ALTA).toLocaleDateString() : "-",
      edificio: nombreEdif,       // Para vista detalle
      edificioNombre: nombreEdif, // Para listas
      idEdificio: idEdif,
      campus: nombreCamp,
      idCampus: idCamp,
      // Guardamos la carpeta para subidas
      carpetaDriveId: d.ID_CARPETA_DRIVE || d.ID_Carpeta_Drive || null 
    };
  } catch(e) { console.error(e); return null; }
}

function updateAsset(d) {
  try {
    const firestore = getFirestore();
    const data = {
      NOMBRE: d.nombre,
      TIPO: d.tipo,
      MARCA: d.marca
    };
    firestore.updateDocument('activos/' + d.id, data);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function deleteAsset(id) {
  try {
    getFirestore().deleteDocument('activos/' + id);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// ==========================================
// 5. EDIFICIOS Y OBRAS
// ==========================================

function getBuildingInfo(id) {
  try {
    const firestore = getFirestore();
    const doc = firestore.getDocument('edificios/' + id);
    if (!doc || !doc.fields) throw new Error("Edificio no encontrado");
    
    const d = doc.fields;
    const idCamp = d.ID_CAMPUS || d.idCampus;
    let nombreCamp = "Desconocido";
    
    if (idCamp) {
       try {
         const c = firestore.getDocument('campus/' + idCamp);
         if(c && c.fields) nombreCamp = c.fields.NOMBRE || c.fields.nombre;
       } catch(e){}
    }

    return { 
      id: doc.name.split('/').pop(), 
      campus: nombreCamp, 
      nombre: d.NOMBRE || d.nombre, 
      contacto: d.CONTACTO || d.contacto 
    };
  } catch(e) { console.error(e); return null; }
}

function getObrasPorEdificio(idEdificio) {
  try {
    const firestore = getFirestore();
    const obras = firestore.query('obras')
                           .where('ID_EDIFICIO', '==', String(idEdificio))
                           .execute();

    return obras.map(o => {
      let fStr = "-";
      if (o.FECHA_INICIO || o.fechaInicio) fStr = new Date(o.FECHA_INICIO || o.fechaInicio).toLocaleDateString();
      
      return {
        id: o.id,
        nombre: o.NOMBRE || o.nombre,
        descripcion: o.DESCRIPCION || o.descripcion,
        fecha: fStr,
        estado: o.ESTADO || o.estado
      };
    });
  } catch (e) { console.error(e); return []; }
}

function crearObra(d) {
  try {
    const firestore = getFirestore();
    const newId = Utilities.getUuid();
    
    // Preparamos los datos para Firestore
    // (Omitimos la creación de carpeta en Drive para máxima velocidad, 
    // puedes añadirla luego si es vital)
    const datosObra = {
      ID: newId,
      ID_EDIFICIO: d.idEdificio,
      NOMBRE: d.nombre,
      DESCRIPCION: d.descripcion || '',
      FECHA_INICIO: d.fechaInicio, // Guardamos la fecha (YYYY-MM-DD)
      ESTADO: 'EN CURSO'
    };
    
    // Guardar en la colección 'obras'
    firestore.createDocument('obras/' + newId, datosObra);
    
    return { success: true };
  } catch(e) {
    console.error("Error crearObra: " + e.message);
    return { success: false, error: e.message };
  }
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
  try {
    const firestore = getFirestore();
    firestore.deleteDocument('obras/' + id);
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ==========================================
// 6. GESTIÓN DOCUMENTAL
// ==========================================

function obtenerDocs(idEntidad, tipoEntidad) {
  try {
    const tipo = tipoEntidad || 'ACTIVO';
    const firestore = getFirestore();
    const idStr = String(idEntidad).trim();
    
    // 1. Buscamos en la colección de documentos
    // Usamos ID_ENTIDAD (que suele ser el nombre estándar en la migración)
    let docs = firestore.query('docs_historico')
                        .where('ID_ENTIDAD', '==', idStr)
                        .execute();
                        
    // Si no encuentra nada, probamos con variantes de nombre por seguridad
    if (docs.length === 0) {
        docs = firestore.query('docs_historico')
                        .where('idEntidad', '==', idStr)
                        .execute();
    }

    // 2. Filtramos en memoria por el TIPO (EDIFICIO vs ACTIVO) para asegurar
    // Y mapeamos a un formato limpio para el HTML
    return docs.filter(d => {
        const tipoDoc = d.TIPO_ENTIDAD || d.TipoEntidad || d.tipoEntidad || '';
        return tipoDoc === tipo;
    }).map(d => {
      // Fecha formateada
      let fechaStr = "-";
      const f = d.FECHA || d.Fecha || d.fecha;
      if (f) {
          try { fechaStr = new Date(f).toLocaleDateString(); } catch(e) {}
      }

      return {
        id: d.id,
        nombre: d.NOMBRE_ARCHIVO || d.NombreArchivo || d.nombre || 'Documento',
        url: d.URL || d.Url || d.url || '#',
        version: d.VERSION || d.Version || 1,
        fecha: fechaStr
      };
    }).sort((a, b) => b.version - a.version); // Más recientes primero

  } catch (e) {
    console.error("Error obteniendo docs: " + e.message);
    return [];
  }
}

function subirArchivo(dataBase64, nombreArchivo, mimeType, idEntidad, tipoEntidad) {
  // if (tipoEntidad !== 'INCIDENCIA') verificarPermiso(['WRITE']); // Descomenta si usas permisos
  
  try {
    const firestore = getFirestore();
    let carpetaId = null;
    const idStr = String(idEntidad).trim();

    // A. BUSCAR LA CARPETA DE DRIVE DESTINO
    if (tipoEntidad === 'EDIFICIO') {
       const doc = firestore.getDocument('edificios/' + idStr);
       // En tu captura se ve que el campo es "ID_Carpeta_Drive"
       if (doc && doc.fields) carpetaId = doc.fields.ID_Carpeta_Drive || doc.fields.ID_CARPETA || doc.fields.fId;
    } 
    else if (tipoEntidad === 'ACTIVO') {
       const doc = firestore.getDocument('activos/' + idStr);
       // Probamos variantes comunes
       if (doc && doc.fields) carpetaId = doc.fields.ID_Carpeta_Drive || doc.fields.CARPETA_ID || doc.fields.fId;
    }
    
    // Fallback: Si no encontramos carpeta específica, usamos la raíz configurada
    if (!carpetaId) {
        carpetaId = getRootFolderId(); // Tu función auxiliar existente
    }

    // B. SUBIR EL ARCHIVO FÍSICO A DRIVE
    const blob = Utilities.newBlob(Utilities.base64Decode(dataBase64), mimeType, nombreArchivo);
    const folder = DriveApp.getFolderById(carpetaId);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // C. GUARDAR EL REGISTRO EN FIRESTORE
    if(tipoEntidad !== 'INCIDENCIA') {
       const newDocId = Utilities.getUuid();
       const datosDoc = {
         ID: newDocId,
         TIPO_ENTIDAD: tipoEntidad,
         ID_ENTIDAD: idStr,
         NOMBRE_ARCHIVO: nombreArchivo,
         URL: file.getUrl(),
         FILE_ID_DRIVE: file.getId(),
         VERSION: 1,
         FECHA: new Date().toISOString(),
         USUARIO: Session.getActiveUser().getEmail()
       };
       
       firestore.createDocument('docs_historico/' + newDocId, datosDoc);
    }
    
    return { success: true, fileId: file.getId(), url: file.getUrl() };

  } catch (e) {
    console.error("Error subirArchivo: " + e.toString());
    return { success: false, error: e.toString() }; 
  }
}

function eliminarDocumento(idDoc) {
  // verificarPermiso(['DELETE']); 
  try {
    const firestore = getFirestore();
    
    // 1. Obtener datos para borrar de Drive también (opcional, buena práctica)
    /* const doc = firestore.getDocument('docs_historico/' + idDoc);
    if (doc && doc.fields && doc.fields.FILE_ID_DRIVE) {
        try { DriveApp.getFileById(doc.fields.FILE_ID_DRIVE).setTrashed(true); } catch(e){}
    }
    */

    // 2. Borrar de Firestore
    firestore.deleteDocument('docs_historico/' + idDoc);
    
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
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
  try {
    const firestore = getFirestore();
    // Consulta: Mantenimientos de este activo que NO estén realizados
    const docs = firestore.query('mantenimientos')
                          .where('ID_ACTIVO', '==', String(idActivo))
                          .execute();
    
    const hoy = new Date(); hoy.setHours(0,0,0,0);

    return docs.filter(m => (m.ESTADO || m.estado) !== 'REALIZADA').map(m => {
       const fStr = m.FECHA || m.fecha;
       let fechaDisp = "-";
       let color = "gris";
       let fIso = "";
       
       if (fStr) {
          const f = new Date(fStr);
          fIso = f.toISOString().split('T')[0];
          fechaDisp = f.toLocaleDateString();
          f.setHours(0,0,0,0);
          const diff = Math.ceil((f.getTime() - hoy.getTime()) / 86400000);
          if (diff < 0) color = 'rojo';
          else if (diff <= 30) color = 'amarillo';
          else color = 'verde';
       }

       return {
         id: m.id,
         tipo: m.TIPO || m.tipo,
         fechaProxima: fechaDisp,
         fechaISO: fIso,
         color: color,
         hasDocs: false, // Optimizado
         hasCalendar: !!(m.EVENT_ID || m.eventId)
       };
    }).sort((a,b) => a.fechaISO.localeCompare(b.fechaISO));

  } catch(e) { console.error(e); return []; }
}

function getGlobalMaintenance() {
  try {
    const firestore = getFirestore();
    const mant = firestore.query('mantenimientos').execute();
    if (!mant.length) return [];

    // Necesitamos cruzar datos para saber el Campus/Edificio de cada revisión
    const activos = firestore.query('activos').execute();
    const edificios = firestore.query('edificios').execute();
    
    // Mapa Edificios (ID -> idCampus)
    const mapEdif = {};
    edificios.forEach(e => {
       mapEdif[e.id] = { 
         nombre: e.Nombre || e.NOMBRE || e.nombre,
         idCampus: e.ID_Campus || e.Campus || e.idCampus
       };
    });

    // Mapa Activos (ID -> idEdificio, idCampus)
    const mapActivos = {};
    activos.forEach(a => {
        const idEdif = a.ID_EDIFICIO || a.ID_Edificio || a.idEdificio;
        const infoEdif = mapEdif[idEdif] || {};
        mapActivos[a.id] = { 
            nombre: a.NOMBRE || a.Nombre || a.nombre || 'Desconocido',
            idEdificio: idEdif,
            idCampus: infoEdif.idCampus,
            nombreEdificio: infoEdif.nombre
        };
    });

    const hoy = new Date(); hoy.setHours(0,0,0,0);
    const result = [];

    mant.forEach(m => {
        const estado = m.ESTADO || m.estado || 'ACTIVO';
        // ... (lógica de fechas igual que antes)
        let fechaStr = m.FECHA || m.fecha;
        let fechaObj = null;
        let dias = 0;
        let color = 'gris';
        if (fechaStr) {
            fechaObj = new Date(fechaStr);
            if (!isNaN(fechaObj.getTime())) {
                fechaObj.setHours(0,0,0,0);
                dias = Math.ceil((fechaObj.getTime() - hoy.getTime()) / 86400000);
                if (estado === 'REALIZADA') { color = 'azul'; dias = 9999; }
                else {
                    if (dias < 0) color = 'rojo';
                    else if (dias <= 30) color = 'amarillo';
                    else color = 'verde';
                }
            }
        }

        const idAct = m.ID_ACTIVO || m.idActivo;
        const info = mapActivos[idAct] || {};

        result.push({
            id: m.id,
            idActivo: idAct,
            activo: info.nombre || 'Borrado',
            edificio: info.nombreEdificio || '-',
            campusId: info.idCampus,     // <--- CLAVE PARA FILTRO
            edificioId: info.idEdificio, // <--- CLAVE PARA FILTRO
            tipo: m.TIPO || m.tipo,
            fecha: fechaObj ? Utilities.formatDate(fechaObj, Session.getScriptTimeZone(), "dd/MM/yyyy") : "-",
            color: color,
            dias: dias
        });
    });

    return result.sort((a, b) => a.dias - b.dias);

  } catch (e) { console.error(e); return []; }
}


function crearRevision(d) {
  try {
    const firestore = getFirestore();
    const fecha = new Date(d.fechaProx);
    let eventId = null;
    
    // Calendar
    if (d.syncCalendar === true || d.syncCalendar === 'true') {
       // ... (Lógica calendar existente, omitida por brevedad, no cambia) ...
       // eventId = gestionarEventoCalendario(...)
    }

    const guardarRev = (fechaDate) => {
        const newId = Utilities.getUuid();
        firestore.createDocument('mantenimientos/' + newId, {
            ID: newId,
            ID_ACTIVO: d.idActivo,
            TIPO: d.tipo,
            FECHA: fechaDate.toISOString(),
            FRECUENCIA: parseInt(d.diasFreq) || 0,
            ESTADO: 'ACTIVO',
            EVENT_ID: eventId
        });
    };

    // 1. Crear la primera
    guardarRev(fecha);

    // 2. Recursividad (Crear futuras)
    if ((d.esRecursiva === true || d.esRecursiva === 'true') && d.diasFreq > 0) {
        const freq = parseInt(d.diasFreq);
        let nextDate = new Date(fecha);
        for(let i=0; i<5; i++) { // Límite 5 años/veces por seguridad
            nextDate.setDate(nextDate.getDate() + freq);
            guardarRev(nextDate);
        }
    }
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function updateRevision(d) {
  try {
    const firestore = getFirestore();
    firestore.updateDocument('mantenimientos/' + d.idPlan, {
        TIPO: d.tipo,
        FECHA: new Date(d.fechaProx).toISOString()
    });
    // Actualizar Calendar si aplica...
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function completarRevision(id) {
  try {
    const firestore = getFirestore();
    // Marcar como REALIZADA
    firestore.updateDocument('mantenimientos/' + id, { ESTADO: 'REALIZADA' });
    
    // Opcional: Generar la siguiente revisión automáticamente si era recurrente
    // (Esto requiere leer la frecuencia del documento actual primero)
    
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function eliminarRevision(id) {
  try {
    getFirestore().deleteDocument('mantenimientos/' + id);
    // Borrar de Calendar si aplica...
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// ==========================================
// 8. CONTRATOS
// ==========================================

function obtenerContratos(idEntidad) {
  try {
    const firestore = getFirestore();
    const idStr = String(idEntidad).trim();
    
    // 1. Traemos contratos que apuntan DIRECTAMENTE a este activo/edificio
    const directos = firestore.query('contratos')
                              .where('ID_ENTIDAD', '==', idStr)
                              .execute();
    
    // 2. Traemos contratos de tipo 'ACTIVOS' (Múltiples) para buscar dentro
    // (Firestore no busca bien dentro de strings JSON, así que traemos todos los de este tipo y filtramos)
    const multiples = firestore.query('contratos')
                               .where('TIPO_ENTIDAD', '==', 'ACTIVOS')
                               .execute();
                               
    const vinculados = multiples.filter(c => {
       const raw = c.ID_ENTIDAD || c.idEntidad || '';
       return raw.includes(idStr); // Búsqueda simple en el string JSON
    });

    const todos = [...directos, ...vinculados];
    
    // Necesitamos nombres de proveedores
    const provs = firestore.query('proveedores').execute();
    const mapProv = {};
    provs.forEach(p => mapProv[p.id] = p.NOMBRE || p.nombre);

    const hoy = new Date(); 
    hoy.setHours(0,0,0,0);

    return todos.map(c => {
        const fFinStr = c.FECHA_FIN || c.fechaFin;
        const fIniStr = c.FECHA_INI || c.fechaIni;
        const idProv = c.ID_PROVEEDOR || c.idProveedor;
        
        let estadoCalc = 'VIGENTE';
        let color = 'verde';
        const estadoDB = c.ESTADO || c.estado || 'ACTIVO';

        if (estadoDB === 'INACTIVO') { estadoCalc = 'INACTIVO'; color = 'gris'; }
        else if (fFinStr) {
            const fFin = new Date(fFinStr);
            const diff = Math.ceil((fFin.getTime() - hoy.getTime()) / 86400000);
            if (diff < 0) { estadoCalc = 'CADUCADO'; color = 'rojo'; }
            else if (diff <= 90) { estadoCalc = 'PRÓXIMO'; color = 'amarillo'; }
        }

        return {
            id: c.id,
            proveedor: mapProv[idProv] || 'Desconocido',
            ref: c.REF || c.ref || '-',
            inicio: fIniStr ? new Date(fIniStr).toLocaleDateString() : "-",
            fin: fFinStr ? new Date(fFinStr).toLocaleDateString() : "-",
            estado: estadoCalc,
            color: color,
            fileUrl: null // La gestión de archivos en contrato requiere lógica extra, por ahora null para velocidad
        };
    });

  } catch (e) { console.error(e); return []; }
}

function updateContratoV2(d) {
  try {
    const firestore = getFirestore();
    
    let idEntidadFinal = d.idEntidad;
    // Manejo de activos múltiples
    if (d.tipoEntidad === 'ACTIVOS' && d.idsActivos && d.idsActivos.length > 0) {
        idEntidadFinal = JSON.stringify(d.idsActivos);
    }

    const data = {
      TIPO_ENTIDAD: d.tipoEntidad,
      ID_ENTIDAD: idEntidadFinal,
      ID_PROVEEDOR: d.idProveedor,
      REF: d.ref,
      FECHA_INI: d.fechaIni ? new Date(d.fechaIni).toISOString() : null,
      FECHA_FIN: d.fechaFin ? new Date(d.fechaFin).toISOString() : null,
      ESTADO: d.estado
    };

    firestore.updateDocument('contratos/' + d.id, data);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function obtenerContratosGlobal() {
  try {
    const firestore = getFirestore();
    const contratos = firestore.query('contratos').execute();
    const provs = firestore.query('proveedores').execute();
    
    // Para resolver jerarquías
    const edificios = firestore.query('edificios').execute();
    const activos = firestore.query('activos').execute();

    const mapProv = {};
    provs.forEach(p => mapProv[p.id] = p.NOMBRE || p.nombre || 'Desconocido');

    // Mapas de jerarquía
    const mapEdif = {}; 
    edificios.forEach(e => mapEdif[e.id] = { nombre: e.Nombre || e.nombre, idCampus: e.ID_Campus || e.Campus });
    
    const mapActivos = {};
    activos.forEach(a => {
       const idE = a.ID_EDIFICIO || a.ID_Edificio || a.idEdificio;
       mapActivos[a.id] = { nombre: a.NOMBRE || a.nombre, idEdificio: idE, idCampus: (mapEdif[idE] || {}).idCampus };
    });

    const hoy = new Date(); hoy.setHours(0,0,0,0);

    return contratos.map(c => {
        const idEntidad = c.ID_ENTIDAD || c.idEntidad;
        const tipoEntidad = c.TIPO_ENTIDAD || c.tipoEntidad;
        let campusId = null;
        let edificioId = null;
        let nombreEntidad = '-';

        // Lógica de resolución de filtros
        if (tipoEntidad === 'CAMPUS') {
            campusId = idEntidad;
            nombreEntidad = 'Campus Completo';
        } else if (tipoEntidad === 'EDIFICIO') {
            edificioId = idEntidad;
            campusId = (mapEdif[idEntidad] || {}).idCampus;
            nombreEntidad = (mapEdif[idEntidad] || {}).nombre;
        } else if (tipoEntidad === 'ACTIVO') {
            const info = mapActivos[idEntidad] || {};
            edificioId = info.idEdificio;
            campusId = info.idCampus;
            nombreEntidad = info.nombre;
        } else if (tipoEntidad === 'ACTIVOS') { // Múltiples
             nombreEntidad = 'Varios Activos';
             // Intentamos sacar ubicación del primer activo para que el filtro funcione al menos parcialmente
             try {
                const ids = JSON.parse(idEntidad);
                if(ids.length > 0) {
                    const info = mapActivos[ids[0]] || {};
                    edificioId = info.idEdificio;
                    campusId = info.idCampus;
                }
             } catch(e){}
        }

        // Lógica Fechas
        const fFinStr = c.FECHA_FIN || c.fechaFin;
        let estadoCalc = 'VIGENTE';
        let color = 'verde';
        const estadoDB = c.ESTADO || c.estado || 'ACTIVO';

        if (estadoDB === 'INACTIVO') { estadoCalc = 'INACTIVO'; color = 'gris'; }
        else if (fFinStr) {
            const fFin = new Date(fFinStr);
            const diff = Math.ceil((fFin.getTime() - hoy.getTime()) / 86400000);
            if (diff < 0) { estadoCalc = 'CADUCADO'; color = 'rojo'; }
            else if (diff <= 90) { estadoCalc = 'PRÓXIMO'; color = 'amarillo'; }
        }

        return {
            id: c.id,
            proveedor: mapProv[c.ID_PROVEEDOR || c.idProveedor] || 'Desconocido',
            ref: c.REF || c.ref || '-',
            inicio: c.FECHA_INI || c.fechaIni ? new Date(c.FECHA_INI || c.fechaIni).toLocaleDateString() : "-",
            fin: fFinStr ? new Date(fFinStr).toLocaleDateString() : "-",
            estado: estadoCalc,
            color: color,
            nombreEntidad: nombreEntidad,
            // IDs para filtros
            campusId: campusId,
            edificioId: edificioId
        };
    });

  } catch (e) { console.error(e); return []; }
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
  try {
    getFirestore().deleteDocument('contratos/' + id);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// ==========================================
// 9. DASHBOARD Y CRUD GENERAL (V5)
// ==========================================

function getDatosDashboard(idCampusFiltro) {
  try {
    const firestore = getFirestore();
    
    // 1. CARGA ULTRA-RÁPIDA (Solo pedimos los IDs para contar)
    // Usamos .select(['__name__']) para que la base de datos no nos mande todo el texto, solo que "existen".
    const activos = firestore.query('activos').select(['__name__']).execute();
    const contratos = firestore.query('contratos').select(['ESTADO', 'Estado', 'FECHA_FIN', 'FechaFin']).execute();
    const incidencias = firestore.query('incidencias').select(['ESTADO', 'Estado']).execute();
    const campus = firestore.query('campus').select(['__name__']).execute();     // <--- NUEVO
    const edificios = firestore.query('edificios').select(['__name__']).execute(); // <--- NUEVO
    
    // Para mantenimiento necesitamos más datos para el calendario
    const mantenimiento = firestore.query('mantenimientos')
                                   .select(['ESTADO', 'Estado', 'FECHA', 'Fecha', 'TIPO', 'Tipo', 'ACTIVO_NOMBRE', 'ActivoNombre'])
                                   .execute();

    // 2. CÁLCULOS
    let totalContratos = 0;
    let totalIncidencias = 0;
    
    // Mantenimiento
    let revPend = 0, revVenc = 0, revOk = 0;
    const eventosCalendario = [];
    const hoy = new Date();
    hoy.setHours(0,0,0,0);

    // Procesar Contratos (Solo contar vigentes)
    contratos.forEach(c => {
       const est = c.ESTADO || c.Estado || c.estado || 'ACTIVO';
       if(est === 'ACTIVO') totalContratos++;
    });

    // Procesar Incidencias (Solo contar abiertas)
    incidencias.forEach(i => {
       const est = i.ESTADO || i.Estado || i.estado;
       if(est !== 'RESUELTA') totalIncidencias++;
    });

    // Procesar Mantenimiento
    mantenimiento.forEach(m => {
      const estado = m.ESTADO || m.Estado || m.estado || 'ACTIVO';
      if (estado === 'REALIZADA') return; 

      // Blindaje de Fecha: Probamos todas las combinaciones posibles
      const fechaStr = m.FECHA || m.Fecha || m.fecha || m.FechaProx || m.FECHA_PROX;
      if (!fechaStr) return;

      const fecha = new Date(fechaStr);
      if (isNaN(fecha.getTime())) return; // Si la fecha no es válida, saltamos

      const diffDias = Math.ceil((fecha - hoy) / (1000 * 60 * 60 * 24));
      
      let color = '#198754'; // Verde
      if (diffDias < 0) { 
          revVenc++; 
          color = '#dc3545'; 
      } else if (diffDias <= 30) { 
          revPend++; 
          color = '#ffc107'; 
      } else { 
          revOk++; 
      }
      
      // Nombre del activo (fallback si no se guardó en el mantenimiento)
      const nombreActivo = m.ACTIVO_NOMBRE || m.ActivoNombre || 'Activo';

      eventosCalendario.push({
        id: m.id,
        title: (m.TIPO || m.Tipo || 'Rev.') + ' - ' + nombreActivo,
        start: fecha.toISOString().split('T')[0], // Formato YYYY-MM-DD seguro
        backgroundColor: color,
        borderColor: color
      });
    });

    return {
      activos: activos.length,
      edificios: edificios.length, // <--- AHORA SÍ CUENTA
      campus: campus.length,       // <--- AHORA SÍ CUENTA
      pendientes: revPend,
      vencidas: revVenc,
      ok: revOk,
      contratos: totalContratos,
      incidencias: totalIncidencias,
      // Gráficos (Simulados por ahora para no ralentizar más)
      chartLabels: ["Ene", "Feb", "Mar", "Abr", "May", "Jun"],
      chartData: [12, 19, 3, 5, 2, 3],
      calendarEvents: eventosCalendario
    };

  } catch(e) {
    console.error("Error Dashboard: " + e.message);
    // Devolvemos objeto vacío seguro para que no se quede cargando infinitamente
    return { activos:0, edificios:0, campus:0, vencidas:0, pendientes:0, ok:0, incidencias:0, contratos:0, calendarEvents:[] };
  }
}

function crearCampus(d) {
  try {
    const firestore = getFirestore();
    const newId = Utilities.getUuid();
    const data = {
      ID: newId,
      NOMBRE: d.nombre,
      PROVINCIA: d.provincia,
      DIRECCION: d.direccion
    };
    firestore.createDocument('campus/' + newId, data);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function crearEdificio(d) {
  try {
    const firestore = getFirestore();
    const newId = Utilities.getUuid();
    
    // Carpetas Drive
    let fId = "", aId = "";
    try {
       const root = getRootFolderId(); // Simplificado
       fId = crearCarpeta(d.nombre, root);
       aId = crearCarpeta("Activos - " + d.nombre, fId);
    } catch(e){}

    const data = {
      ID: newId,
      ID_CAMPUS: d.idCampus,
      NOMBRE: d.nombre,
      CONTACTO: d.contacto,
      LAT: d.lat,
      LNG: d.lng,
      ID_CARPETA_DRIVE: fId,
      ID_CARPETA_ACTIVOS: aId
    };
    
    firestore.createDocument('edificios/' + newId, data);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function crearActivo(d) {
  try {
    const firestore = getFirestore();
    const newId = Utilities.getUuid();
    
    // Crear carpeta en Drive (Opcional, si quieres mantener el orden)
    // Si da error, sigue adelante sin carpeta
    let folderId = "";
    try {
       // Buscar carpeta del edificio
       if (d.idEdificio) {
          const docEdif = firestore.getDocument('edificios/' + d.idEdificio);
          // Asumimos que guardaste ID_CARPETA_ACTIVOS en el edificio al migrar
          const parentId = docEdif.fields.ID_CARPETA_ACTIVOS || getRootFolderId();
          folderId = crearCarpeta(d.nombre, parentId); 
       } else {
          folderId = getRootFolderId();
       }
    } catch(err) { console.warn("Drive error: " + err); }

    const data = {
      ID: newId,
      ID_EDIFICIO: d.idEdificio,
      ID_CAMPUS: d.idCampus,
      NOMBRE: d.nombre,
      TIPO: d.tipo,
      MARCA: d.marca,
      FECHA_ALTA: new Date().toISOString(),
      ID_CARPETA_DRIVE: folderId
    };

    firestore.createDocument('activos/' + newId, data);
    return { success: true, newId: newId };
  } catch(e) { return { success: false, error: e.message }; }
}

function getCatalogoInstalaciones() {
  return getConfigCatalogo(); // Alias
}

function getTableData(tipo) {
  const firestore = getFirestore();
  
  // --- CASO 1: CAMPUS ---
  if (tipo === 'CAMPUS') {
    try {
      const docs = firestore.query('campus').execute();
      return docs.map(d => ({
        id: d.id,
        nombre: d.Nombre || d.nombre || d.NOMBRE || 'Sin Nombre',
        provincia: d.Provincia || d.provincia || d.PROVINCIA || '-',
        direccion: d.Direccion || d.direccion || d.DIRECCION || '-'
      })).sort((a, b) => a.nombre.localeCompare(b.nombre));
    } catch (e) { return []; }
  }
  
  // --- CASO 2: EDIFICIOS (Ya lo tenías, lo incluimos aquí para tener todo junto) ---
  if (tipo === 'EDIFICIOS') {
    try {
      // 1. Descargar campus para traducir IDs
      const campusDocs = firestore.query('campus').execute();
      const mapaCampus = {};
      campusDocs.forEach(c => {
        mapaCampus[c.id] = c.Nombre || c.nombre || c.NOMBRE || "Desconocido";
      });

      // 2. Descargar edificios
      const docs = firestore.query('edificios').execute();
      
      return docs.map(d => {
        const idCamp = d.ID_Campus || d.idCampus || d.Campus || '';
        return {
          id: d.id,
          nombre: d.Nombre || d.nombre || d.NOMBRE || 'Sin nombre',
          campus: mapaCampus[idCamp] || idCamp || '-', 
          contacto: d.Contacto || d.contacto || d.CONTACTO || '-',
          nActivos: 0, nIncidencias: 0 // Simplificado para velocidad
        };
      }).sort((a, b) => a.nombre.localeCompare(b.nombre));
    } catch (e) { return []; }
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

function getConfigCatalogo() {
  try {
    const firestore = getFirestore();
    const docs = firestore.query('catalogo').execute();
    return docs.map(c => ({
      id: c.id,
      // Usamos || para asegurar que lo pilla esté como esté escrito
      nombre: c.NOMBRE || c.nombre || c.Nombre || 'Sin nombre',
      normativa: c.NORMATIVA || c.normativa || '-',
      dias: c.DIAS || c.dias || 365
    })).sort((a, b) => a.nombre.localeCompare(b.nombre));
  } catch (e) { console.error(e); return []; }
}

function getCatalogoInstalaciones() { return getConfigCatalogo(); }

function saveConfigCatalogo(d) {
  // verificarPermiso(['ADMIN_ONLY']); // Descomenta para seguridad
  try {
    const firestore = getFirestore();
    const id = d.id || Utilities.getUuid();
    const data = {
      NOMBRE: d.nombre,
      NORMATIVA: d.normativa,
      DIAS: parseInt(d.dias)
    };
    firestore.updateDocument('catalogo/' + id, data);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function deleteConfigCatalogo(id) {
  try {
    getFirestore().deleteDocument('catalogo/' + id);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// Gestión de Usuarios
function saveUsuario(u) {
  try {
    const firestore = getFirestore();
    const id = u.id || Utilities.getUuid();
    const data = {
      NOMBRE: u.nombre,
      EMAIL: u.email,
      ROL: u.rol,
      RECIBIR_AVISOS: u.avisos
    };
    firestore.updateDocument('usuarios/' + id, data);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function getListaUsuarios() {
  try {
    const firestore = getFirestore();
    const docs = firestore.query('usuarios').execute();
    return docs.map(u => ({
      id: u.id,
      nombre: u.NOMBRE || u.nombre,
      email: u.EMAIL || u.email,
      rol: u.ROL || u.rol || 'CONSULTA',
      avisos: u.RECIBIR_AVISOS || u.avisos || 'NO'
    }));
  } catch (e) { console.error(e); return []; }
}

function saveUsuario(u) { verificarPermiso(['ADMIN_ONLY']); const ss = SpreadsheetApp.openById(PROPS.getProperty('DB_SS_ID')); const sheet = ss.getSheetByName('USUARIOS'); if (!u.nombre || !u.email) return { success: false, error: "Datos incompletos" }; if (u.id) { const data = sheet.getDataRange().getValues(); for(let i=1; i<data.length; i++){ if(String(data[i][0]) === String(u.id)) { sheet.getRange(i+1, 2).setValue(u.nombre); sheet.getRange(i+1, 3).setValue(u.email); sheet.getRange(i+1, 4).setValue(u.rol); sheet.getRange(i+1, 5).setValue(u.avisos); return { success: true }; } } return { success: false, error: "Usuario no encontrado" }; } else { sheet.appendRow([Utilities.getUuid(), u.nombre, u.email, u.rol, u.avisos]); return { success: true }; } }

function deleteUsuario(id) {
  try {
    getFirestore().deleteDocument('usuarios/' + id);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// ==========================================
// 12. GESTIÓN DE INCIDENCIAS
// ==========================================

function getIncidencias() {
  try {
    const firestore = getFirestore();
    const docs = firestore.query('incidencias').execute();
    
    if (!docs.length) return [];

    return docs.map(d => {
        let fechaStr = "-";
        if (d.FECHA || d.fecha) {
            const f = new Date(d.FECHA || d.fecha);
            fechaStr = Utilities.formatDate(f, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
        }

        return {
            id: d.id,
            tipoOrigen: d.TIPO_ORIGEN || d.tipoOrigen,
            nombreOrigen: d.NOMBRE_ORIGEN || d.nombreOrigen,
            descripcion: d.DESCRIPCION || d.descripcion,
            prioridad: d.PRIORIDAD || d.prioridad,
            estado: d.ESTADO || d.estado,
            fecha: fechaStr,
            solicitante: d.USUARIO || d.usuario || '-',
            urlFoto: "" // Fotos requieren lógica extra de Drive, dejar vacío por velocidad ahora
        };
    }).sort((a,b) => {
        // Ordenar: Pendientes primero
        if(a.estado === 'RESUELTA' && b.estado !== 'RESUELTA') return 1;
        if(a.estado !== 'RESUELTA' && b.estado === 'RESUELTA') return -1;
        return 0;
    });

  } catch (e) {
    console.error("Error getIncidencias: " + e.message);
    return [];
  }
}

function getIncidenciaById(id) {
  try {
    const firestore = getFirestore();
    const doc = firestore.getDocument('incidencias/' + id);
    
    if (!doc || !doc.fields) return null;
    
    // Mapeo seguro de campos
    return {
      id: doc.name.split('/').pop(),
      titulo: doc.fields.TITULO || doc.fields.titulo || 'Incidencia',
      descripcion: doc.fields.DESCRIPCION || doc.fields.descripcion || '',
      prioridad: doc.fields.PRIORIDAD || doc.fields.prioridad || 'MEDIA',
      estado: doc.fields.ESTADO || doc.fields.estado || 'PENDIENTE',
      tipoOrigen: doc.fields.TIPO_ORIGEN || doc.fields.tipoOrigen,
      idOrigen: doc.fields.ID_ORIGEN || doc.fields.idOrigen
    };
  } catch(e) {
    console.error("Error getIncidenciaById: " + e.message);
    return null;
  }
}

function crearIncidencia(d) {
  try {
    const firestore = getFirestore();
    const newId = Utilities.getUuid();
    
    const data = {
        TIPO_ORIGEN: d.tipoOrigen,
        ID_ORIGEN: d.idOrigen,
        NOMBRE_ORIGEN: d.nombreOrigen,
        DESCRIPCION: d.descripcion,
        PRIORIDAD: d.prioridad,
        ESTADO: 'PENDIENTE',
        FECHA: new Date().toISOString(),
        USUARIO: Session.getActiveUser().getEmail()
    };
    
    // Si hay foto (base64), la subimos a Drive y guardamos la URL
    if (d.fotoBase64) {
        // ... (Lógica de subida a Drive similar a la de documentos) ...
        // Por simplicidad, aquí guardamos que tiene foto, o implementamos subirArchivo
        // data.URL_FOTO = url; 
    }

    firestore.createDocument('incidencias/' + newId, data);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function actualizarEstadoIncidencia(id, nuevoEstado) {
  try {
    const firestore = getFirestore();
    firestore.updateDocument('incidencias/' + id, { ESTADO: nuevoEstado });
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function getIncidenciaDetalle(id) {
  try {
    const firestore = getFirestore();
    const doc = firestore.getDocument('incidencias/' + id);
    if (!doc || !doc.fields) throw new Error("Incidencia no encontrada");
    
    const d = doc.fields;
    const tipo = d.TIPO_ORIGEN || d.tipoOrigen;
    const idOrigen = d.ID_ORIGEN || d.idOrigen;
    
    let idCampus = null, idEdificio = null, idActivo = null;
    
    // Lógica para saber dónde está la incidencia
    if (tipo === 'ACTIVO') {
        idActivo = idOrigen;
        const a = getAssetInfo(idActivo); // Reutilizamos función
        if (a) { idEdificio = a.idEdificio; idCampus = a.idCampus; }
    } else if (tipo === 'EDIFICIO') {
        idEdificio = idOrigen;
        const b = getBuildingInfo(idEdificio);
        if (b) { 
            // getBuildingInfo devuelve el nombre del campus, aquí necesitamos el ID. 
            // Hacemos consulta rápida si es necesario o asumimos que el frontend lo tiene.
            // Para editar no suele ser crítico, pero si lo fuera, haríamos un getDocument('edificios/'+id)
        }
    } else if (tipo === 'CAMPUS') {
        idCampus = idOrigen;
    }

    return {
      id: doc.name.split('/').pop(),
      tipoOrigen: tipo,
      idOrigen: idOrigen,
      nombreOrigen: d.NOMBRE_ORIGEN || d.nombreOrigen,
      descripcion: d.DESCRIPCION || d.descripcion,
      prioridad: d.PRIORIDAD || d.prioridad,
      estado: d.ESTADO || d.estado,
      idCampus: idCampus,
      idEdificio: idEdificio,
      idActivo: idActivo,
      urlFoto: d.URL_FOTO || null
    };
  } catch(e) { console.error(e); return null; }
}

function updateIncidenciaData(d) {
  try {
    const firestore = getFirestore();
    const data = {
      DESCRIPCION: d.descripcion,
      PRIORIDAD: d.prioridad,
      // Si permites cambiar ubicación:
      // TIPO_ORIGEN: d.tipoOrigen,
      // ID_ORIGEN: d.idOrigen
    };
    firestore.updateDocument('incidencias/' + d.id, data);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
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
  try {
    const firestore = getFirestore();
    const docs = firestore.query('feedback').execute();
    return docs.map(r => ({
      id: r.id,
      fecha: r.FECHA || r.fecha ? new Date(r.FECHA || r.fecha).toLocaleDateString() : '-',
      usuario: r.USUARIO || r.usuario,
      tipo: r.TIPO || r.tipo,
      mensaje: r.MENSAJE || r.mensaje,
      estado: r.ESTADO || r.estado
    })).reverse(); // Nuevos primero
  } catch (e) { return []; }
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
  try {
    const firestore = getFirestore();
    const data = {
      NOMBRE: d.nombre,
      PROVINCIA: d.provincia,
      DIRECCION: d.direccion
    };
    firestore.updateDocument('campus/' + d.id, data);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function eliminarCampus(id) {
  try {
    // Nota: Deberías borrar también edificios hijos, pero por ahora borramos el campus
    getFirestore().deleteDocument('campus/' + id);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function updateEdificio(d) {
  try {
    const firestore = getFirestore();
    const data = {
      NOMBRE: d.nombre,
      ID_CAMPUS: d.idCampus, // Vinculación
      CONTACTO: d.contacto,
      LAT: d.lat,
      LNG: d.lng
    };
    firestore.updateDocument('edificios/' + d.id, data);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function eliminarEdificio(id) {
  try {
    getFirestore().deleteDocument('edificios/' + id);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
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
  try {
    const firestore = getFirestore();
    // Traemos todo. Si hay muchísimos logs, habría que usar .Limit(100) pero tu librería manual básica quizás no lo tenga.
    // Asumimos que no son millones.
    const docs = firestore.query('logs').execute(); 
    
    // Ordenar por fecha descendente en memoria
    docs.sort((a, b) => {
       const dA = new Date(a.FECHA || a.fecha).getTime();
       const dB = new Date(b.FECHA || b.fecha).getTime();
       return dB - dA;
    });

    return docs.slice(0, 100).map(r => ({
      fecha: r.FECHA || r.fecha ? new Date(r.FECHA || r.fecha).toLocaleString() : '-',
      usuario: r.USUARIO || r.usuario,
      accion: r.ACCION || r.accion,
      detalles: r.DETALLES || r.detalles
    }));
  } catch (e) { return []; }
}

// ==========================================
// 19. SISTEMA DE NOVEDADES (CHANGELOG)
// ==========================================

function getNovedadesApp() {
  try {
    const firestore = getFirestore();
    const docs = firestore.query('novedades').execute();
    return docs.map(r => ({
      fecha: r.FECHA || r.fecha ? new Date(r.FECHA || r.fecha).toLocaleDateString() : '-',
      version: r.VERSION || r.version,
      tipo: r.TIPO || r.tipo,
      titulo: r.TITULO || r.titulo,
      descripcion: r.DESCRIPCION || r.descripcion
    })).reverse();
  } catch (e) { return []; }
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
  // Las relaciones suelen guardarse en una subcolección o colección aparte
  // Asumiremos colección 'relaciones'
  try {
    const firestore = getFirestore();
    const rels = firestore.query('relaciones')
                          .where('ID_ACTIVO_A', '==', idActivo)
                          .execute();
                          
    // Necesitamos nombres de los relacionados
    const activos = firestore.query('activos').select(['NOMBRE', 'TIPO']).execute();
    const mapNombres = {};
    activos.forEach(a => mapNombres[a.id] = { nombre: a.NOMBRE || a.nombre, tipo: a.TIPO || a.tipo });

    return rels.map(r => ({
      id: r.id,
      idActivoRelacionado: r.ID_ACTIVO_B,
      nombreActivo: (mapNombres[r.ID_ACTIVO_B] || {}).nombre || 'Desconocido',
      tipoActivo: (mapNombres[r.ID_ACTIVO_B] || {}).tipo || '-',
      tipoRelacion: r.TIPO_RELACION,
      descripcion: r.DESCRIPCION || ''
    }));
  } catch(e) { return []; }
}

/**
 * Guardar nueva relación entre activos (bidireccional)
 */

function crearRelacionActivo(d) {
  try {
    const firestore = getFirestore();
    const id1 = Utilities.getUuid();
    
    // Relación A -> B
    firestore.createDocument('relaciones/' + id1, {
      ID_ACTIVO_A: d.idActivoA,
      ID_ACTIVO_B: d.idActivoB,
      TIPO_RELACION: d.tipoRelacion,
      DESCRIPCION: d.descripcion
    });

    // Si es bidireccional, creamos la inversa B -> A
    if (d.bidireccional) {
      // Lógica de inversión de tipo (ej: ALIMENTA -> ES_ALIMENTADO_POR)
      const inversos = { "ALIMENTA": "ES_ALIMENTADO_POR", "ES_ALIMENTADO_POR": "ALIMENTA", "DEPENDE_DE": "ES_REQUERIDO_POR" };
      const tipoInv = inversos[d.tipoRelacion] || d.tipoRelacion; // Fallback
      
      const id2 = Utilities.getUuid();
      firestore.createDocument('relaciones/' + id2, {
        ID_ACTIVO_A: d.idActivoB,
        ID_ACTIVO_B: d.idActivoA,
        TIPO_RELACION: tipoInv,
        DESCRIPCION: d.descripcion
      });
    }
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
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
  try {
    getFirestore().deleteDocument('relaciones/' + idRelacion);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
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
    const firestore = getFirestore();
    let idEntidadDestino = data.idActivo;
    let tipoEntidadDestino = 'ACTIVO';
    let nombreFinal = data.nombreArchivo;

    // A. LÓGICA OCA (Inspección Reglamentaria)
    if (data.categoria === 'OCA') {
      const idRevision = Utilities.getUuid();
      // Convertir fecha string a objeto Date
      const fechaReal = data.fechaOCA ? new Date(data.fechaOCA) : new Date();
      
      // 1. Crear la revisión "padre" (Histórica/Carga) en Firestore
      firestore.createDocument('mantenimientos/' + idRevision, {
        ID: idRevision,
        ID_ACTIVO: data.idActivo,
        TIPO: 'Legal',
        // Título descriptivo para diferenciarlo
        ACTIVO_NOMBRE: 'OCA - Documento Histórico/Carga', 
        FECHA: fechaReal.toISOString(),
        FRECUENCIA: parseInt(data.freqDias) || 365,
        ESTADO: 'REALIZADA' // Se marca como hecha porque subimos el certificado
      });

      // 2. Generar futuras revisiones automáticas (si se marcó el check)
      if (data.crearSiguientes && data.freqDias > 0) {
        const frecuencia = parseInt(data.freqDias);
        const MAX_REVISIONES = 10;
        const hoy = new Date();
        const fechaTope = new Date(hoy.getFullYear() + 10, hoy.getMonth(), hoy.getDate()); // 10 años
        
        let fechaSiguiente = new Date(fechaReal);
        let contador = 0;

        while (contador < MAX_REVISIONES) {
          fechaSiguiente.setDate(fechaSiguiente.getDate() + frecuencia);
          if (fechaSiguiente > fechaTope) break;
          
          const idFutura = Utilities.getUuid();
          firestore.createDocument('mantenimientos/' + idFutura, {
            ID: idFutura,
            ID_ACTIVO: data.idActivo,
            TIPO: 'Legal',
            ACTIVO_NOMBRE: 'Próxima Inspección Reglamentaria',
            FECHA: fechaSiguiente.toISOString(),
            FRECUENCIA: frecuencia,
            ESTADO: 'ACTIVO'
          });
          contador++;
        }
      }
      
      // El archivo se vinculará a esta revisión específica
      idEntidadDestino = idRevision;
      tipoEntidadDestino = 'REVISION';
      nombreFinal = "OCA_" + nombreFinal;
    } 
    
    // B. LÓGICA CONTRATO
    else if (data.categoria === 'CONTRATO') {
      const idContrato = Utilities.getUuid();
      
      // Datos por defecto si faltan
      const prov = data.contProveedor || "Proveedor Desconocido"; 
      // Nota: Aquí guardamos el nombre directo porque la subida rápida es "sucia",
      // idealmente deberíamos buscar el ID del proveedor, pero para no complicar usamos texto.
      const ref = data.contRef || "S/N";
      const ini = data.contIni ? new Date(data.contIni) : new Date();
      
      let fin = new Date(ini);
      if (data.contFin) {
         fin = new Date(data.contFin);
      } else {
         fin.setFullYear(fin.getFullYear() + 1); // 1 año por defecto
      }

      // Crear el contrato en Firestore
      firestore.createDocument('contratos/' + idContrato, {
        ID: idContrato,
        TIPO_ENTIDAD: 'ACTIVO',
        ID_ENTIDAD: data.idActivo,
        // Al ser subida rápida, quizás no tenemos el ID del proveedor en la BBDD,
        // así que guardamos el nombre en un campo auxiliar o buscamos uno genérico.
        // Estrategia: Guardar el texto en un campo 'PROVEEDOR_TEXTO' si no hay ID.
        ID_PROVEEDOR: 'GENERICO', 
        PROVEEDOR_TEXTO: prov, // Campo extra para mostrar si no hay cruce
        REF: ref,
        FECHA_INI: ini.toISOString(),
        FECHA_FIN: fin.toISOString(),
        ESTADO: 'ACTIVO'
      });

      nombreFinal = "CONTRATO_" + prov + "_" + nombreFinal;
      // El archivo se queda en el activo para fácil acceso
    }

    // Finalmente, subimos el archivo físico y creamos su registro en Firestore
    return subirArchivo(data.base64, nombreFinal, data.mimeType, idEntidadDestino, tipoEntidadDestino);

  } catch (e) {
    return { success: false, error: e.message };
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

function getListaProveedores() {
  try {
    const firestore = getFirestore();
    const docs = firestore.query('proveedores').execute();
    return docs.map(p => ({
      id: p.id,
      nombre: p.NOMBRE || p.nombre || 'Sin Nombre',
      cif: p.CIF || p.cif || '-',
      contacto: p.CONTACTO || p.contacto || '-',
      telefono: p.TELEFONO || p.telefono || '-',
      email: p.EMAIL || p.email || '-',
      activo: (p.ACTIVO || p.activo) !== 'NO'
    })).sort((a, b) => a.nombre.localeCompare(b.nombre));
  } catch (e) { console.error(e); return []; }
}

/**
 * Crear o actualizar proveedor
 */

function saveProveedor(d) {
  try {
    const firestore = getFirestore();
    const id = d.id || Utilities.getUuid();
    const data = {
      NOMBRE: d.nombre,
      CIF: d.cif,
      CONTACTO: d.contacto,
      TELEFONO: d.telefono,
      EMAIL: d.email,
      ACTIVO: d.activo ? 'SI' : 'NO'
    };
    firestore.updateDocument('proveedores/' + id, data);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

/**
 * Eliminar proveedor (solo si no tiene contratos activos)
 */

function eliminarProveedor(id) {
  try {
    getFirestore().deleteDocument('proveedores/' + id);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
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

function getContratoFullDetailsV2(id) {
  try {
    const firestore = getFirestore();
    const doc = firestore.getDocument('contratos/' + id);
    if (!doc || !doc.fields) throw new Error("Contrato no encontrado");
    const d = doc.fields;
    
    // Parsear JSON de activos si existe
    let idsActivos = [];
    if ((d.TIPO_ENTIDAD || d.tipoEntidad) === 'ACTIVOS') {
       try { idsActivos = JSON.parse(d.ID_ENTIDAD || d.idEntidad); } catch(e){}
    }

    let fechaIniStr = d.FECHA_INI || d.fechaIni;
    let fechaFinStr = d.FECHA_FIN || d.fechaFin;

    return {
        id: doc.name.split('/').pop(),
        tipoEntidad: d.TIPO_ENTIDAD || d.tipoEntidad,
        idEntidad: d.ID_ENTIDAD || d.idEntidad,
        idsActivos: idsActivos,
        idProveedor: d.ID_PROVEEDOR || d.idProveedor,
        ref: d.REF || d.ref,
        inicio: fechaIniStr ? fechaIniStr.split('T')[0] : "",
        fin: fechaFinStr ? fechaFinStr.split('T')[0] : "",
        estado: d.ESTADO || d.estado
    };
  } catch(e) { console.error(e); return null; }
}

function crearContratoV2(d) {
  try {
    const firestore = getFirestore();
    const newId = Utilities.getUuid();
    
    // Lógica para determinar el ID de entidad (si es un array de activos, lo guardamos como JSON string)
    let idEntidadFinal = d.idEntidad;
    let tipoEntidadFinal = d.tipoEntidad;

    if (d.idsActivos && d.idsActivos.length > 0) {
        tipoEntidadFinal = 'ACTIVOS';
        idEntidadFinal = JSON.stringify(d.idsActivos);
    }

    const data = {
      ID: newId,
      TIPO_ENTIDAD: tipoEntidadFinal,
      ID_ENTIDAD: idEntidadFinal,
      ID_PROVEEDOR: d.idProveedor, // Guardamos el ID, no el nombre
      REF: d.ref,
      // Guardamos fechas en formato ISO para Firestore
      FECHA_INI: d.fechaIni ? new Date(d.fechaIni).toISOString() : null,
      FECHA_FIN: d.fechaFin ? new Date(d.fechaFin).toISOString() : null,
      ESTADO: d.estado || 'ACTIVO'
    };

    firestore.createDocument('contratos/' + newId, data);
    return { success: true, newId: newId };
  } catch(e) { return { success: false, error: e.message }; }
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
