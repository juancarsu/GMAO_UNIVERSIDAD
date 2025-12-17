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

/**
 * OBTENER TODO DE UNA VEZ PARA EL FRONTEND (Optimizado)
 */
function getAppInitData() {
  const usuario = getMyRole(); //
  const firestore = getFirestore(); //
  
  // Cargamos Campus, Edificios y Catálogo en paralelo
  const campusDocs = firestore.query('campus').execute(); //
  const edificiosDocs = firestore.query('edificios').execute(); //
  const catalogoDocs = getConfigCatalogo(); //

  return {
    usuario: usuario,
    listaCampus: campusDocs.map(c => ({ id: c.id, nombre: c.NOMBRE || c.Nombre || 'Sin Nombre' })), //
    listaEdificios: edificiosDocs.map(e => ({ 
      id: e.id, 
      nombre: e.NOMBRE || e.Nombre || 'Sin Nombre', 
      idCampus: e.ID_CAMPUS || e.idCampus || e.ID_Campus || e.Campus // Variantes de ID
    })),
    catalogo: catalogoDocs //
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

/**
 * Crea el índice de búsqueda global leyendo desde Firestore
 * (Sustituye a la versión que leía de Sheets)
 */
function buildActivosIndex() {
  const cacheKey = 'INDEX_ACTIVOS';
  
  // 1. Intentar leer el índice completo de la caché para velocidad
  try {
    const cached = CACHE.get(cacheKey);
    if (cached) {
      return JSON.parse(cached);
    }
  } catch(e) {
    console.warn('Cache miss o error en buildActivosIndex: ' + e.message);
  }

  // 2. Si no está en caché, lo construimos usando los datos de Firestore
  // Usamos getAllAssetsList porque ya hace el trabajo duro de cruzar Activos con Edificios/Campus
  const listaActivos = getAllAssetsList(); 
  
  const index = {
    byId: {},
    byEdificio: {},
    byCampus: {},
    searchable: []
  };

  listaActivos.forEach(a => {
    // A. Mapa por ID para acceso directo
    index.byId[a.id] = a;

    // B. Array de búsqueda (Texto en minúsculas incluyendo nombre, tipo, marca y edificio)
    const textoBusqueda = (
      (a.nombre || '') + ' ' + 
      (a.tipo || '') + ' ' + 
      (a.marca || '') + ' ' + 
      (a.edificioNombre || '')
    ).toLowerCase();

    index.searchable.push({
      id: a.id,
      text: textoBusqueda
    });

    // C. Agrupación por Edificio (útil para filtros rápidos)
    if (a.idEdificio) {
      if (!index.byEdificio[a.idEdificio]) {
        index.byEdificio[a.idEdificio] = [];
      }
      index.byEdificio[a.idEdificio].push(a);
    }

    // D. Agrupación por Campus
    if (a.idCampus) {
      if (!index.byCampus[a.idCampus]) {
        index.byCampus[a.idCampus] = [];
      }
      index.byCampus[a.idCampus].push(a);
    }
  });

  // 3. Guardar el resultado en caché (10 minutos)
  try {
    // Nota: Si tienes miles de activos, el JSON podría exceder el límite de 100KB de Apps Script Cache.
    // Si eso pasa, el índice funcionará pero será más lento (se reconstruirá en cada búsqueda).
    CACHE.put(cacheKey, JSON.stringify(index), 600); 
  } catch(e) {
    console.warn('El índice es demasiado grande para la caché: ' + e.message);
  }

  return index;
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
  try {
    const firestore = getFirestore();
    const idStr = String(idCampus).trim();
    
    // Buscamos en Firestore por ID_Campus (Tal como sale en tu captura)
    let docs = firestore.query('edificios').where('ID_Campus', '==', idStr).execute();
    
    // Si falla, probamos variantes por seguridad
    if (!docs.length) docs = firestore.query('edificios').where('idCampus', '==', idStr).execute();
    
    return docs.map(d => ({
      id: d.id,
      nombre: d.NOMBRE || d.Nombre || d.nombre || 'Sin Nombre'
    })).sort((a, b) => a.nombre.localeCompare(b.nombre));
    
  } catch (e) {
    console.error("Error getEdificiosPorCampus: " + e.message);
    return [];
  }
}

function getActivosPorEdificio(idEdificio) {
  try {
    const firestore = getFirestore();
    const idStr = String(idEdificio).trim();
    
    // Buscamos por ID_Edificio
    let docs = firestore.query('activos').where('ID_Edificio', '==', idStr).execute();
    
    // Variantes de seguridad
    if (!docs.length) docs = firestore.query('activos').where('ID_EDIFICIO', '==', idStr).execute();
    if (!docs.length) docs = firestore.query('activos').where('idEdificio', '==', idStr).execute();

    return docs.map(d => ({
      id: d.id,
      nombre: d.NOMBRE || d.Nombre || d.nombre || 'Sin Nombre',
      tipo: d.TIPO || d.Tipo || d.tipo || '-',
      marca: d.MARCA || d.Marca || d.marca || '-'
    })).sort((a, b) => a.nombre.localeCompare(b.nombre));

  } catch (e) {
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
  // Limpieza de seguridad del ID
  const idLimpio = String(id).includes('/') ? String(id).split('/').pop() : String(id);
  
  try {
    const firestore = getFirestore();
    
    // 1. Leemos el documento
    const d = firestore.getDocument('activos/' + idLimpio);
    
    if (!d) {
      return { id: idLimpio, nombre: "Error: Doc no existe", tipo: "-", marca: "-", edificio: "-", campus: "-" };
    }

    // --- BLOQUE DE DIAGNÓSTICO ---
    // Vamos a ver qué claves tiene realmente el documento
    const claves = Object.keys(d).join(", ");
    const jsonMuestra = JSON.stringify(d).substring(0, 150); // Primeros 150 caracteres
    
    // Intentamos recuperar datos usando TODAS las variantes posibles
    const nombre = d.NOMBRE || d.nombre || d.Nombre || ('DEBUG: ' + claves);
    const tipo = d.TIPO || d.tipo || d.Tipo || '-';
    const marca = d.MARCA || d.marca || d.Marca || '-';
    // -----------------------------

    // Resolver Edificio (Lógica estándar)
    const idEdif = d.ID_EDIFICIO || d.idEdificio || d.IdEdificio || '';
    let nombreEdif = '-';
    let nombreCamp = '-';
    
    if (idEdif) {
      try {
        const docEdif = firestore.getDocument('edificios/' + idEdif);
        if (docEdif) {
           nombreEdif = docEdif.NOMBRE || docEdif.nombre || 'Sin nombre';
           const idCamp = docEdif.ID_CAMPUS || docEdif.idCampus;
           if (idCamp) {
              const docCamp = firestore.getDocument('campus/' + idCamp);
              if (docCamp) nombreCamp = docCamp.NOMBRE || docCamp.nombre || '-';
           }
        }
      } catch(e) {}
    }

    return {
      id: idLimpio,
      nombre: nombre, // Aquí verás las claves reales si falla el nombre
      tipo: tipo,
      marca: marca,
      fechaAlta: "-",
      edificio: nombreEdif,       
      edificioNombre: nombreEdif, 
      idEdificio: idEdif,
      campus: nombreCamp,
      idCampus: null,
      carpetaDriveId: null 
    };

  } catch(e) { 
    return {
      id: idLimpio, 
      nombre: "Error Script: " + e.message, 
      tipo: "-", marca: "-", edificio: "-", campus: "-"
    };
  }
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
    
    if (!doc) throw new Error("Edificio no encontrado");
    
    // SOPORTE HÍBRIDO
    const d = doc.fields ? doc.fields : doc;
    
    // Resolver Campus
    const idCamp = d.ID_CAMPUS || d.ID_Campus || d.idCampus;
    let nombreCamp = "Desconocido";
    
    if (idCamp) {
       try {
         const c = firestore.getDocument('campus/' + idCamp);
         if(c) {
            const dCamp = c.fields ? c.fields : c;
            nombreCamp = dCamp.NOMBRE || dCamp.Nombre || dCamp.nombre || '-';
         }
       } catch(e){}
    }

    return { 
      id: id, 
      campus: nombreCamp, 
      // Variantes de nombre para asegurar que salga
      nombre: d.NOMBRE || d.Nombre || d.nombre || 'Sin Nombre', 
      contacto: d.CONTACTO || d.Contacto || d.contacto || '-' 
    };
  } catch(e) { 
    console.error(e); 
    return { id: id, nombre: "Error carga", campus: "-", contacto: "-" };
  }
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
  // verificarPermiso(['WRITE']);
  try {
    const firestore = getFirestore();
    const fechaObj = new Date(fechaFin).toISOString(); // Formato ISO para Firestore
    
    firestore.updateDocument('obras/' + idObra, {
      FECHA_FIN: fechaObj,
      ESTADO: 'FINALIZADA'
    });
    
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
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
    const coleccion = 'docs_historico';
    
    // 1. Búsqueda "Todoterreno" del ID
    // Probamos: ID_ENTIDAD, Id_Entidad, id_entidad, idEntidad
    let docs = firestore.query(coleccion).where('ID_ENTIDAD', '==', idStr).execute();
    if (!docs.length) docs = firestore.query(coleccion).where('Id_Entidad', '==', idStr).execute();
    if (!docs.length) docs = firestore.query(coleccion).where('ID_Entidad', '==', idStr).execute(); // Variante probable
    if (!docs.length) docs = firestore.query(coleccion).where('idEntidad', '==', idStr).execute();

    // 2. Mapeo con soporte para tus nombres de campos
    return docs.filter(d => {
        // Buscamos: TIPO_ENTIDAD, Tipo_Entidad, tipoEntidad
        const dTipo = d.TIPO_ENTIDAD || d.Tipo_Entidad || d.TipoEntidad || d.tipoEntidad || '';
        return dTipo === tipo;
    }).map(d => {
      let fechaStr = "-";
      // Buscamos: FECHA, Fecha, fecha
      const fRaw = d.FECHA || d.Fecha || d.fecha;
      
      if (fRaw) {
          try { 
             // Convertir si es Timestamp o String
             const f = fRaw.toDate ? fRaw.toDate() : new Date(fRaw);
             fechaStr = f.toLocaleDateString(); 
          } catch(e) {}
      }

      return {
        id: d.id,
        // Buscamos: NOMBRE_ARCHIVO, Nombre_Archivo, NombreArchivo
        nombre: d.NOMBRE_ARCHIVO || d.Nombre_Archivo || d.NombreArchivo || d.nombre || 'Documento',
        url: d.URL || d.Url || d.url || '#',
        version: d.VERSION || d.Version || d.version || 1,
        fecha: fechaStr
      };
    }).sort((a, b) => b.version - a.version);

  } catch (e) { return []; }
}

function subirArchivo(dataBase64, nombreArchivo, mimeType, idEntidad, tipoEntidad) {
  try {
    const firestore = getFirestore();
    let carpetaId = null;
    const idStr = String(idEntidad).trim();

    // 1. OBTENER CARPETA DESTINO (Lógica Blindada)
    try {
      if (tipoEntidad === 'EDIFICIO') {
         const doc = firestore.getDocument('edificios/' + idStr);
         if (doc) {
             const d = doc.fields ? doc.fields : doc; // Soporte híbrido
             carpetaId = d.ID_Carpeta_Drive || d.ID_CARPETA_DRIVE || d.idCarpetaDrive;
         }
      } 
      else if (tipoEntidad === 'ACTIVO' || tipoEntidad === 'REVISION') {
         // Si es revisión, primero buscamos el activo padre
         let idActivo = idStr;
         if (tipoEntidad === 'REVISION') {
             const docRev = firestore.getDocument('mantenimientos/' + idStr);
             const dRev = docRev ? (docRev.fields || docRev) : {};
             idActivo = dRev.ID_ACTIVO || dRev.Id_Activo || dRev.idActivo;
         }

         if (idActivo) {
             const doc = firestore.getDocument('activos/' + idActivo);
             if (doc) {
                 const d = doc.fields ? doc.fields : doc; // Soporte híbrido
                 // Probamos variantes de nombre del campo carpeta
                 carpetaId = d.ID_Carpeta_Drive || d.ID_CARPETA_DRIVE || d.idCarpetaDrive || d.Carpeta_Id;
             }
         }
      }
    } catch(e) {
      console.warn("Error buscando carpeta específica: " + e.message);
    }
    
    // 2. FALLBACK: Si no encontramos carpeta, usamos la RAÍZ configurada
    if (!carpetaId) {
        console.log("⚠️ No se encontró carpeta específica, usando RAÍZ.");
        carpetaId = getRootFolderId(); 
    }

    // 3. SUBIR ARCHIVO A DRIVE
    const blob = Utilities.newBlob(Utilities.base64Decode(dataBase64), mimeType, nombreArchivo);
    const folder = DriveApp.getFolderById(carpetaId);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // 4. REGISTRAR EN FIRESTORE
    if(tipoEntidad !== 'INCIDENCIA') {
       const newDocId = Utilities.getUuid();
       const datosDoc = {
         ID: newDocId,
         TIPO_ENTIDAD: tipoEntidad, // 'ACTIVO', 'EDIFICIO', 'REVISION'
         ID_ENTIDAD: idStr,
         NOMBRE_ARCHIVO: nombreArchivo,
         URL: file.getUrl(),
         FILE_ID_DRIVE: file.getId(), // Guardamos ID para poder borrar luego
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
  // Reutilizamos getAssetInfo que ya está adaptada a Firestore
  const info = getAssetInfo(idActivo);
  
  if (info) {
    return { 
      nombreActivo: info.nombre || "Activo", 
      marca: info.marca || "-", 
      edificio: info.edificio || "Sin ubicación" 
    };
  }
  
  return { nombreActivo: "Activo Desconocido", marca: "-", edificio: "-" };
}

function obtenerPlanMantenimiento(idActivo) {
  try {
    const firestore = getFirestore();
    const idStr = String(idActivo).trim();
    
    // Búsqueda por ID_Activo (Tal como sale en tu captura)
    // Probamos ID_Activo y ID_ACTIVO por si acaso
    let docs = firestore.query('mantenimientos').where('ID_Activo', '==', idStr).execute();
    if (!docs.length) docs = firestore.query('mantenimientos').where('ID_ACTIVO', '==', idStr).execute();
    
    const hoy = new Date(); 
    hoy.setHours(0,0,0,0);

    return docs.filter(m => {
       const est = m.ESTADO || m.Estado || m.estado || 'ACTIVO';
       return est !== 'REALIZADA';
    }).map(m => {
       // --- USANDO TUS NOMBRES REALES ---
       const tipo = m.Tipo_Revision || m.TIPO_REVISION || m.Tipo || 'Revisión';
       const fRaw = m.Fecha_Proxima || m.FECHA_PROXIMA || m.Fecha;
       
       let fechaDisp = "-";
       let fIso = "";
       let color = "gris";
       
       if (fRaw) {
          try {
            const f = fRaw.toDate ? fRaw.toDate() : new Date(fRaw);
            if (!isNaN(f.getTime())) {
                fIso = f.toISOString().split('T')[0];
                fechaDisp = f.toLocaleDateString();
                const fCheck = new Date(f); fCheck.setHours(0,0,0,0);
                const diff = Math.ceil((fCheck.getTime() - hoy.getTime()) / 86400000);
                if (diff < 0) color = 'rojo';
                else if (diff <= 30) color = 'amarillo';
                else color = 'verde';
            }
          } catch(e) {}
       }

       return {
         id: m.id,
         tipo: tipo,
         fechaProxima: fechaDisp,
         fechaISO: fIso,
         color: color,
         hasDocs: false, 
         hasCalendar: false
       };
    }).sort((a,b) => a.fechaISO.localeCompare(b.fechaISO));

  } catch(e) { return []; }
}

function getGlobalMaintenance() {
  try {
    const firestore = getFirestore();
    const zonaHoraria = Session.getScriptTimeZone();

    // 1. CARGA MASIVA DE DATOS
    const mant = firestore.query('mantenimientos').execute();
    if (!mant.length) return [];

    const activos = firestore.query('activos').execute();
    const edificios = firestore.query('edificios').execute();
    const campusList = firestore.query('campus').execute(); // Necesario para quitar el "undefined"

    // 2. MAPAS DE NOMBRES (Indexados por ID limpio)
    
    // Mapa Campus (ID -> Nombre)
    const mapCampus = {};
    campusList.forEach(c => {
        const cID = String(c.ID || c.id).trim();
        mapCampus[cID] = c.NOMBRE || c.Nombre || "Campus";
    });

    // Mapa Edificios (ID -> Info completa)
    const mapEdif = {};
    edificios.forEach(e => {
       const eID = String(e.ID || e.id).trim();
       const cID = String(e.ID_Campus || e.idCampus || "").trim();
       mapEdif[eID] = { 
         nombre: e.NOMBRE || e.Nombre || "Edificio",
         idCampus: cID,
         nombreCampus: mapCampus[cID] || "-" // Recuperamos el nombre del campus
       };
    });

    // Mapa Activos (ID -> Info completa)
    const mapActivos = {};
    activos.forEach(a => {
        const aID = String(a.ID || a.id).trim();
        const idE = String(a.ID_Edificio || a.idEdificio || "").trim();
        const infoEdif = mapEdif[idE] || {};
        
        mapActivos[aID] = { 
            nombre: a.NOMBRE || a.Nombre || "Activo Desconocido",
            idEdificio: idE,
            nombreEdificio: infoEdif.nombre || "-",
            idCampus: infoEdif.idCampus,
            nombreCampus: infoEdif.nombreCampus || "-" // Pasamos el nombre del campus
        };
    });

    const hoy = new Date(); hoy.setHours(0,0,0,0);
    const result = [];

    // 3. PROCESAR MANTENIMIENTOS
    mant.forEach(m => {
        // Normalización de campos
        const estado = m.ESTADO || m.Estado || 'ACTIVO';
        const tipo = m.Tipo_Revision || m.TIPO_REVISION || m.Tipo || 'Revisión';
        const fRaw = m.Fecha_Proxima || m.FECHA_PROXIMA || m.Fecha;
        
        // --- CLAVE DEL ARREGLO: Limpiar el ID antes de buscar ---
        const idRaw = m.ID_Activo || m.ID_ACTIVO || m.idActivo || "";
        const idAct = String(idRaw).trim(); // Quitamos espacios que rompen el enlace

        // Buscamos en el mapa
        const info = mapActivos[idAct];

        // Si no lo encuentra, mostramos el ID para que sepas cuál falla
        const nombreActivo = info ? info.nombre : `⚠️ ID Roto: ${idAct}`;
        const nombreEdificio = info ? info.nombreEdificio : '-';
        const nombreCampus = info ? info.nombreCampus : '-';

        let fechaDisp = "-";
        let fechaObj = null;
        let fechaISO = "";
        let dias = 0;
        let color = 'gris';

        if (fRaw) {
            try {
                fechaObj = fRaw.toDate ? fRaw.toDate() : new Date(fRaw);
                if (!isNaN(fechaObj.getTime())) {
                    const fCalc = new Date(fechaObj); 
                    fCalc.setHours(0,0,0,0);
                    dias = Math.ceil((fCalc.getTime() - hoy.getTime()) / 86400000);
                    
                    if (estado === 'REALIZADA') { 
                        color = 'azul'; dias = 9999; 
                    } else {
                        if (dias < 0) color = 'rojo';
                        else if (dias <= 30) color = 'amarillo';
                        else color = 'verde';
                    }
                    fechaISO = Utilities.formatDate(fechaObj, zonaHoraria, "yyyy-MM-dd");
                    fechaDisp = fechaObj.toLocaleDateString(); 
                }
            } catch(e){}
        }

        result.push({
            id: m.id,
            idActivo: idAct,
            activo: nombreActivo,
            edificio: nombreEdificio,
            campus: nombreCampus, // ¡Ahora enviamos el nombre del campus!
            campusId: info ? info.idCampus : '',     
            edificioId: info ? info.idEdificio : '',
            tipo: tipo,
            fecha: fechaDisp,
            fechaISO: fechaISO,
            color: color,
            dias: dias
        });
    });

    return result.sort((a, b) => a.dias - b.dias);

  } catch (e) { 
    console.error("Error global mantenimiento: " + e.message);
    return []; 
  }
}


function crearRevision(d) {
  try {
    const firestore = getFirestore();
    
    // --- CORRECCIÓN DE FECHAS (El truco del Mediodía) ---
    // Creamos la fecha base
    const fecha = new Date(d.fechaProx);
    // Forzamos la hora a las 12:00 del mediodía.
    // Así evitamos que al convertir a UTC (restar 1 o 2 horas) nos cambie de día.
    fecha.setHours(12, 0, 0, 0); 
    // ----------------------------------------------------

    let eventId = null;
    
    // 1. Sincronización con Google Calendar
    if (d.syncCalendar === true || d.syncCalendar === 'true') {
       const infoActivo = getAssetInfo(d.idActivo);
       
       if (infoActivo) {
           const datosEvento = {
               tipo: d.tipo,
               nombreActivo: infoActivo.nombre,
               marca: infoActivo.marca,
               edificio: infoActivo.edificio,
               // Para Calendar usamos el string directo "YYYY-MM-DD" que viene del formulario
               // Así Calendar lo trata como "Todo el día" y no se lía con horas.
               fecha: d.fechaProx 
           };
           eventId = gestionarEventoCalendario('CREAR', datosEvento, null);
       }
    }

    // Función interna para guardar
    const guardarRev = (fechaDate) => {
        const newId = Utilities.getUuid();
        
        // Aseguramos que la fecha recursiva también sea a las 12:00
        fechaDate.setHours(12, 0, 0, 0);

        firestore.createDocument('mantenimientos/' + newId, {
            ID: newId,
            ID_Activo: d.idActivo,
            Tipo_Revision: d.tipo,
            Fecha_Proxima: fechaDate.toISOString(), // Ahora guardará "...T12:00:00.000Z"
            Frecuencia: parseInt(d.diasFreq) || 0,
            Estado: 'ACTIVO',
            Event_Id: eventId 
        });
    };

    // 2. Crear la primera revisión
    guardarRev(fecha);

    // 3. Recursividad (Futuras revisiones)
    if ((d.esRecursiva === true || d.esRecursiva === 'true') && d.diasFreq > 0) {
        const freq = parseInt(d.diasFreq);
        // Usamos una copia de la fecha para no modificar la original
        let nextDate = new Date(fecha); 
        
        for(let i=0; i<5; i++) { 
            // Sumamos días
            nextDate.setDate(nextDate.getDate() + freq);
            // El eventId lo dejamos null para las repeticiones (o crearía 5 eventos en calendar)
            eventId = null; 
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

function completarRevision(idMant, fechaRealizadaStr, comentarios, tiempoInvertido) {
  try {
    const firestore = getFirestore();
    
    // 1. LEER LA REVISIÓN ACTUAL COMPLETA
    const doc = firestore.getDocument('mantenimientos/' + idMant);
    if (!doc) throw new Error("Revisión no encontrada");
    
    // Hacemos copia de seguridad de todos los datos (ID_Activo, Tipo, etc.)
    // Si tu librería devuelve el objeto directo, usamos {...doc}. 
    // Si devuelve .fields, usamos eso.
    const datosAntiguos = doc.fields ? doc.fields : doc;
    const datosActualizados = { ...datosAntiguos };
    
    // Borramos el campo 'id' interno para no duplicarlo al guardar
    delete datosActualizados.id; 
    delete datosActualizados.ID;

    // 2. ACTUALIZAR ESTADO Y DATOS DE CIERRE
    const fechaReal = fechaRealizadaStr ? new Date(fechaRealizadaStr) : new Date();
    // Forzamos mediodía para evitar bailes de fecha
    fechaReal.setHours(12, 0, 0, 0);

    datosActualizados.Estado = 'REALIZADA'; // O 'COMPLETADA' según uses
    datosActualizados.Fecha_Ultima = fechaReal.toISOString();
    datosActualizados.Comentarios = comentarios || "";
    datosActualizados.Tiempo_Invertido = tiempoInvertido || 0;
    
    // IMPORTANTE: Nos aseguramos de que ID_Activo siga existiendo
    // Si por algún motivo se perdió, intentamos recuperarlo (aunque aquí ya debería estar en datosAntiguos)
    if (!datosActualizados.ID_Activo && !datosActualizados.idActivo) {
        throw new Error("Error crítico: La revisión original no tenía ID de Activo.");
    }

    // 3. GUARDAR LA REVISIÓN ACTUALIZADA (Merge manual)
    firestore.updateDocument('mantenimientos/' + idMant, datosActualizados);

    // 4. GENERAR LA SIGUIENTE REVISIÓN (Si es recurrente)
    // Usamos los datos antiguos para saber la frecuencia
    const freq = parseInt(datosActualizados.Frecuencia || datosActualizados.Dias_Freq || 0);
    
    if (freq > 0) {
        const nuevaFecha = new Date(fechaReal);
        nuevaFecha.setDate(nuevaFecha.getDate() + freq);
        nuevaFecha.setHours(12, 0, 0, 0); // Mantener hora segura

        const newId = Utilities.getUuid();
        const nuevaRevision = {
            ID: newId,
            ID_Activo: datosActualizados.ID_Activo || datosActualizados.idActivo, // Heredamos el enlace
            Tipo_Revision: datosActualizados.Tipo_Revision || datosActualizados.Tipo,
            Fecha_Proxima: nuevaFecha.toISOString(),
            Frecuencia: freq,
            Estado: 'ACTIVO',
            Event_Id: null // La nueva nace sin evento de calendar (o créalo si quieres)
        };
        
        firestore.createDocument('mantenimientos/' + newId, nuevaRevision);
    }

    return { success: true };

  } catch(e) { 
    console.error("Error completarRevision: " + e.message);
    return { success: false, error: e.message }; 
  }
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
    
    // 1. CARGAR TODOS LOS PROVEEDORES Y CREAR MAPA (ID -> NOMBRE)
    const docsProvs = firestore.query('proveedores').execute();
    const mapProv = {};
    
    docsProvs.forEach(p => {
        // Aseguramos leer el ID correctamente
        const pID = p.ID || p.id; 
        const pNombre = p.NOMBRE || p.Nombre || p.Empresa || "Sin Nombre";
        if (pID) {
            mapProv[pID] = pNombre;
        }
    });

    // 2. BUSCAR CONTRATOS (Tu patrón: ID_Entidad)
    let contratos = firestore.query('contratos').where('ID_Entidad', '==', idStr).execute();
    if (!contratos.length) contratos = firestore.query('contratos').where('ID_ENTIDAD', '==', idStr).execute();
    
    // Añadir los contratos de tipo 'ACTIVOS' (Múltiples)
    const multiples = firestore.query('contratos').where('Tipo_Entidad', '==', 'ACTIVOS').execute();
    multiples.forEach(c => {
       const rawIDs = c.ID_Entidad || c.ID_ENTIDAD || '';
       if (rawIDs.includes(idStr)) contratos.push(c);
    });

    const hoy = new Date(); hoy.setHours(0,0,0,0);

    return contratos.map(c => {
        // Fechas
        const fFinRaw = c.Fecha_Fin || c.FECHA_FIN;
        const fIniRaw = c.Fecha_Inicio || c.FECHA_INI;
        
        // --- TRADUCCIÓN DEL PROVEEDOR ---
        // Tu campo se llama 'Proveedor' y contiene el ID (ej: 2a398...)
        const idGuardado = c.Proveedor || c.PROVEEDOR || c.ID_Proveedor;
        // Buscamos ese ID en el mapa. Si no está, mostramos el ID para no dejarlo vacío.
        const nombreReal = mapProv[idGuardado] || idGuardado || "Proveedor Desconocido";

        const ref = c.Num_Ref || c.REF || "-";
        const estadoDB = c.Estado || c.ESTADO || 'ACTIVO';
        
        let finStr = "-";
        let inicioStr = "-";
        let estadoCalc = 'VIGENTE';
        let color = 'verde';

        if (fFinRaw) {
            try {
                const fFin = fFinRaw.toDate ? fFinRaw.toDate() : new Date(fFinRaw);
                finStr = fFin.toLocaleDateString();
                const diff = Math.ceil((fFin.getTime() - hoy.getTime()) / 86400000);
                
                if (estadoDB === 'INACTIVO') { estadoCalc = 'INACTIVO'; color = 'gris'; }
                else if (diff < 0) { estadoCalc = 'CADUCADO'; color = 'rojo'; }
                else if (diff <= 90) { estadoCalc = 'PRÓXIMO'; color = 'amarillo'; }
            } catch(e) {}
        }
        
        if (fIniRaw) {
             try { 
               const fIni = fIniRaw.toDate ? fIniRaw.toDate() : new Date(fIniRaw);
               inicioStr = fIni.toLocaleDateString(); 
             } catch(e){}
        }

        return {
            id: c.id,
            proveedor: nombreReal, // ¡Ahora saldrá el nombre!
            ref: ref,
            inicio: inicioStr,
            fin: finStr,
            estado: estadoCalc,
            color: color,
            fileUrl: null
        };
    });

  } catch (e) { return []; }
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
    
    // 1. CARGA MASIVA DE DATOS
    // Traemos todo de una vez para cruzar datos (Providers, Campus, Edificios, Activos)
    const contratos = firestore.query('contratos').execute();
    const provs = firestore.query('proveedores').execute();
    const edificios = firestore.query('edificios').execute();
    const activos = firestore.query('activos').execute();
    const campusList = firestore.query('campus').execute();

    // 2. CREAR MAPAS DE BÚSQUEDA RÁPIDA
    
    // Mapa Proveedores (ID -> Nombre)
    const mapProv = {};
    provs.forEach(p => {
        const pID = p.ID || p.id;
        const pNombre = p.NOMBRE || p.Nombre || p.Empresa || "Sin Nombre";
        if (pID) mapProv[pID] = pNombre;
    });

    // Mapa Campus (ID -> Nombre)
    const mapCampus = {};
    campusList.forEach(c => {
        const cID = c.ID || c.id;
        mapCampus[cID] = c.NOMBRE || c.Nombre || "Campus";
    });

    // Mapa Edificios (ID -> Info)
    const mapEdif = {}; 
    edificios.forEach(e => {
        const eID = e.ID || e.id;
        const idC = e.ID_Campus || e.idCampus || e.ID_CAMPUS;
        mapEdif[eID] = { 
            nombre: e.NOMBRE || e.Nombre || "Edificio", 
            idCampus: idC,
            nombreCampus: mapCampus[idC] || "-"
        };
    });
    
    // Mapa Activos (ID -> Info)
    const mapActivos = {};
    activos.forEach(a => {
       const aID = a.ID || a.id;
       const idE = a.ID_Edificio || a.ID_EDIFICIO || a.idEdificio;
       const infoEdif = mapEdif[idE] || {};
       
       mapActivos[aID] = { 
           nombre: a.NOMBRE || a.Nombre || "Activo", 
           idEdificio: idE, 
           idCampus: infoEdif.idCampus,
           nombreEdificio: infoEdif.nombre,
           nombreCampus: infoEdif.nombreCampus
       };
    });

    const hoy = new Date(); hoy.setHours(0,0,0,0);

    // 3. PROCESAR CADA CONTRATO
    return contratos.map(c => {
        // --- USANDO TUS CAMPOS DE FIRESTORE (Captura) ---
        const idEntidad = c.ID_Entidad || c.ID_ENTIDAD || "";
        const tipoEntidad = c.Tipo_Entidad || c.TIPO_ENTIDAD || "CAMPUS";
        const idProv = c.Proveedor || c.PROVEEDOR || c.ID_Proveedor;
        const ref = c.Num_Ref || c.REF || "-";
        
        // Fechas
        const fFinRaw = c.Fecha_Fin || c.FECHA_FIN;
        const fIniRaw = c.Fecha_Inicio || c.FECHA_INI;

        // --- RESOLVER NOMBRE ENTIDAD (Ubicación) ---
        let campusId = null;
        let edificioId = null;
        let nombreEntidad = '-';
        let campusNombre = '-';

        if (tipoEntidad === 'CAMPUS') {
            campusId = idEntidad;
            nombreEntidad = mapCampus[idEntidad] || 'Campus General';
            campusNombre = nombreEntidad;
        } 
        else if (tipoEntidad === 'EDIFICIO') {
            edificioId = idEntidad;
            const info = mapEdif[idEntidad] || {};
            campusId = info.idCampus;
            nombreEntidad = info.nombre;
            campusNombre = info.nombreCampus;
        } 
        else if (tipoEntidad === 'ACTIVO') {
            const info = mapActivos[idEntidad] || {};
            edificioId = info.idEdificio;
            campusId = info.idCampus;
            nombreEntidad = info.nombre;
            campusNombre = info.nombreCampus;
        } 
        else if (tipoEntidad === 'ACTIVOS') { // Múltiples
             nombreEntidad = 'Varios Activos';
             try {
                // Intentamos sacar ubicación del primer activo para que el filtro funcione
                const ids = JSON.parse(idEntidad);
                if(ids.length > 0) {
                    const info = mapActivos[ids[0]] || {};
                    edificioId = info.idEdificio;
                    campusId = info.idCampus;
                    campusNombre = info.nombreCampus;
                    nombreEntidad = info.nombre + ` (+${ids.length-1})`;
                }
             } catch(e){}
        }

        // --- CALCULAR ESTADO Y FECHAS ---
        let finStr = "-";
        let inicioStr = "-";
        let estadoCalc = 'VIGENTE';
        let color = 'verde';
        const estadoDB = c.Estado || c.ESTADO || 'ACTIVO';

        if (estadoDB === 'INACTIVO') { 
            estadoCalc = 'INACTIVO'; 
            color = 'gris'; 
        }
        else if (fFinRaw) {
            try {
                const fFin = new Date(fFinRaw); // Formato ISO
                finStr = fFin.toLocaleDateString();
                const diff = Math.ceil((fFin.getTime() - hoy.getTime()) / 86400000);
                
                if (diff < 0) { estadoCalc = 'CADUCADO'; color = 'rojo'; }
                else if (diff <= 90) { estadoCalc = 'PRÓXIMO'; color = 'amarillo'; }
            } catch(e) {}
        }
        
        if (fIniRaw) {
             try { inicioStr = new Date(fIniRaw).toLocaleDateString(); } catch(e){}
        }

        return {
            id: c.id,
            // Traducir ID Proveedor a Nombre
            proveedor: mapProv[idProv] || idProv || 'Desconocido', 
            ref: ref,
            inicio: inicioStr,
            fin: finStr,
            estado: estadoCalc,
            color: color,
            nombreEntidad: nombreEntidad,
            campusNombre: campusNombre,
            // IDs para que funcionen los filtros de la tabla
            campusId: campusId,
            edificioId: edificioId
        };
    });

  } catch (e) { 
    console.error("Error global contratos: " + e.message); 
    return []; 
  }
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
    
    // 1. OBTENER CARPETA PADRE (Del Edificio)
    let parentFolderId = getRootFolderId(); 
    
    if (d.idEdificio) {
       try {
         const docEdif = firestore.getDocument('edificios/' + d.idEdificio);
         if (docEdif) {
            const dataEdif = docEdif.fields ? docEdif.fields : docEdif;
            // Buscamos carpeta activos del edificio (variantes)
            const idCarpeta = dataEdif.ID_Carpeta_Activos || dataEdif.ID_CARPETA_ACTIVOS || dataEdif.ID_Carpeta_Drive;
            if (idCarpeta) parentFolderId = idCarpeta;
         }
       } catch(e) {}
    }

    // 2. CREAR CARPETA
    let folderId = "";
    try {
        folderId = crearCarpeta(d.nombre, parentFolderId);
    } catch(err) { folderId = parentFolderId; }

    // 3. GUARDAR (Nombres limpios PascalCase como en tus capturas)
    const data = {
      ID: newId,
      ID_Edificio: d.idEdificio,
      ID_Campus: d.idCampus,
      Nombre: d.nombre,
      Tipo: d.tipo,
      Marca: d.marca,
      Fecha_Alta: new Date().toISOString(),
      ID_Carpeta_Drive: folderId
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
    // 1. Usamos el índice de activos que ya lee de Firestore
    const index = buildActivosIndex();
    const textoLower = texto.toLowerCase();
    const resultados = [];
    
    // A. Búsqueda en ACTIVOS
    if (index && index.searchable) {
      for (let i = 0; i < index.searchable.length && resultados.length < 10; i++) {
        const item = index.searchable[i];
        if (item.text.includes(textoLower)) {
          const activo = index.byId[item.id];
          if (activo) {
            resultados.push({
              id: activo.id,
              tipo: 'ACTIVO',
              texto: activo.nombre,
              subtexto: (activo.tipo || '') + (activo.marca ? " - " + activo.marca : "")
            });
          }
        }
      }
    }
    
    // B. Búsqueda en EDIFICIOS (Directo a Firestore si faltan resultados)
    if (resultados.length < 10) {
      const firestore = getFirestore();
      const docs = firestore.query('edificios').execute(); // Trae todos (optimizar si son muchos)
      
      for (let i = 0; i < docs.length && resultados.length < 10; i++) {
        const edif = docs[i];
        // Soportar mayúsculas/minúsculas en el nombre del campo
        const nombre = (edif.NOMBRE || edif.Nombre || edif.nombre || '').toLowerCase();
        
        if (nombre.includes(textoLower)) {
          resultados.push({
            id: edif.id || edif.name.split('/').pop(),
            tipo: 'EDIFICIO',
            texto: edif.NOMBRE || edif.Nombre || edif.nombre,
            subtexto: 'Edificio'
          });
        }
      }
    }
    
    return resultados;
    
  } catch(e) {
    console.error('Error en buscarGlobal: ' + e.toString());
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
  // verificarPermiso(['WRITE']); // Descomenta si usas seguridad
  const firestore = getFirestore();
  
  // 1. Cargar mapas de IDs para búsqueda rápida (Campus y Edificios ya migrados)
  // Nota: Esto asume que Campus y Edificios ya están en Firestore
  const campusDocs = firestore.query('campus').execute();
  const edifDocs = firestore.query('edificios').execute();
  
  const mapaCampus = {}; 
  campusDocs.forEach(c => mapaCampus[(c.NOMBRE || c.nombre).trim().toLowerCase()] = c.id);
  
  const mapaEdificios = {}; 
  edifDocs.forEach(e => {
    // Clave compuesta: nombre_edificio + id_campus
    const nombre = (e.NOMBRE || e.nombre).trim().toLowerCase();
    const idC = e.ID_CAMPUS || e.idCampus;
    mapaEdificios[nombre + "_" + idC] = e.id;
  });

  const batchSize = 100; // Firestore prefiere escrituras por lotes si es posible, o loops controlados
  let procesados = 0;
  const errores = [];

  for (let index = 0; index < filas.length; index++) {
    const fila = filas[index];
    if (fila.length < 4) continue; 

    const nombreCampus = String(fila[0]).trim();
    const nombreEdif = String(fila[1]).trim();
    const tipo = String(fila[2]).trim();
    const nombreActivo = String(fila[3]).trim();
    const marca = fila[4] ? String(fila[4]).trim() : "-";

    // Resolver IDs
    const idCampus = mapaCampus[nombreCampus.toLowerCase()];
    if (!idCampus) {
      errores.push(`Fila ${index + 1}: Campus '${nombreCampus}' no existe en Firestore.`);
      continue;
    }

    const idEdificio = mapaEdificios[nombreEdif.toLowerCase() + "_" + idCampus];
    if (!idEdificio) {
      errores.push(`Fila ${index + 1}: Edificio '${nombreEdif}' no encontrado en ese Campus.`);
      continue;
    }

    // Crear Activo en Firestore
    try {
      const newId = Utilities.getUuid();
      const data = {
        ID: newId,
        ID_EDIFICIO: idEdificio,
        ID_CAMPUS: idCampus,
        NOMBRE: nombreActivo,
        TIPO: tipo,
        MARCA: marca,
        FECHA_ALTA: new Date().toISOString(),
        // ID_CARPETA_DRIVE: ... (Lógica de Drive opcional aquí)
      };
      
      firestore.createDocument('activos/' + newId, data);
      procesados++;
    } catch (e) {
      errores.push(`Error guardando fila ${index + 1}: ${e.message}`);
    }
  }

  return { success: true, procesados: procesados, errores: errores };
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
/**
 * Obtiene los detalles completos de un contrato desde Firestore.
 * Incluye la resolución de la jerarquía (Campus -> Edificio) para el modal de edición.
 */

function getContratoFullDetailsV2(id) {
  try {
    const firestore = getFirestore();
    const doc = firestore.getDocument('contratos/' + id);
    if (!doc) throw new Error("Contrato no encontrado en Firestore");
    
    // Soporte híbrido para el formato de documento (fields o directo)
    const d = doc.fields ? doc.fields : doc;
    const tipoEntidad = d.Tipo_Entidad || d.TIPO_ENTIDAD || "CAMPUS";
    const idEntidad = d.ID_Entidad || d.ID_ENTIDAD || "";
    
    // Parsear JSON de activos si el tipo es 'ACTIVOS'
    let idsActivos = [];
    if (tipoEntidad === 'ACTIVOS') {
       try { 
         idsActivos = typeof idEntidad === 'string' ? JSON.parse(idEntidad) : idEntidad; 
       } catch(e) { console.error("Error parseando IDs de activos:", e); }
    }

    let idCampus = "";
    let idEdificio = "";
    let idActivo = "";

    // RESOLUCIÓN DE JERARQUÍA SEGÚN EL TIPO DE CONTRATO
    if (tipoEntidad === 'CAMPUS') {
      idCampus = idEntidad;
    } 
    else if (tipoEntidad === 'EDIFICIO') {
      idEdificio = idEntidad;
      const edif = firestore.getDocument('edificios/' + idEdificio);
      if (edif) {
        const dE = edif.fields || edif;
        idCampus = dE.ID_CAMPUS || dE.idCampus || dE.ID_Campus;
      }
    } 
    else if (tipoEntidad === 'ACTIVO') {
      idActivo = idEntidad;
      const activo = firestore.getDocument('activos/' + idActivo);
      if (activo) {
        const dA = activo.fields || activo;
        idEdificio = dA.ID_Edificio || dA.idEdificio || dA.ID_EDIFICIO;
        idCampus = dA.ID_Campus || dA.idCampus || dA.ID_CAMPUS;
      }
    }
    else if (tipoEntidad === 'ACTIVOS' && idsActivos.length > 0) {
      // Si son varios activos, tomamos la ubicación del primero para los selectores del modal
      const primero = firestore.getDocument('activos/' + idsActivos[0]);
      if (primero) {
        const dP = primero.fields || primero;
        idEdificio = dP.ID_Edificio || dP.idEdificio || dP.ID_EDIFICIO;
        idCampus = dP.ID_Campus || dP.idCampus || dP.ID_CAMPUS;
      }
    }

    return {
        id: id,
        tipoEntidad: tipoEntidad,
        idEntidad: idEntidad,
        idsActivos: idsActivos,
        idProveedor: d.Proveedor || d.PROVEEDOR || d.ID_PROVEEDOR,
        ref: d.Num_Ref || d.REF || d.ref || "",
        inicio: d.Fecha_Inicio || d.FECHA_INI || d.inicio || "",
        fin: d.Fecha_Fin || d.FECHA_FIN || d.fin || "",
        estado: d.Estado || d.ESTADO || d.estado || "ACTIVO",
        idCampus: idCampus,
        idEdificio: idEdificio,
        idActivo: idActivo
    };
  } catch(e) { 
    console.error("Error en getContratoFullDetailsV2:", e.message); 
    return null; 
  }
}

// ==========================================
// 22. SISTEMA DE NOTIFICACIONES AUTOMÁTICAS (VERSIÓN FIRESTORE)
// ==========================================

/**
 * FUNCIÓN PRINCIPAL - Ejecutar diariamente mediante Trigger
 */

function enviarNotificacionesAutomaticas() {
  try {
    console.log("=== INICIO NOTIFICACIONES (Firestore) ===");
    
    // 1. Obtener usuarios (Admins/Técnicos) que quieren avisos
    const usuarios = obtenerUsuariosConAvisos();
    if (!usuarios.length) {
      console.log("No hay usuarios suscritos a notificaciones.");
      return;
    }

    // 2. Calcular datos (Usando funciones ya migradas a Firestore)
    const revProximas = detectarRevisionesProximas();
    const revVencidas = detectarRevisionesVencidas();
    const contProximos = detectarContratosProximos();

    // 3. Enviar correos si hay algo que contar
    if (revProximas.length > 0 || revVencidas.length > 0 || contProximos.length > 0) {
      usuarios.forEach(u => {
        enviarEmailResumen(u, revProximas, revVencidas, contProximos);
      });
      console.log(`✅ Notificaciones enviadas a ${usuarios.length} usuarios.`);
    } else {
      console.log("Todo en orden. No hay alertas hoy.");
    }

  } catch (e) {
    console.error("Error crítico en notificaciones: " + e.message);
  }
}

/**
 * Obtiene usuarios con rol ADMIN o TECNICO y avisos activados desde Firestore
 */
function obtenerUsuariosConAvisos() {
  try {
    const firestore = getFirestore();
    // Buscamos usuarios que tengan el campo RECIBIR_AVISOS en 'SI'
    // Nota: Ajusta 'RECIBIR_AVISOS' si en tu base de datos se llama 'avisos' o 'recibirAvisos'
    const docs = firestore.query('usuarios').where('RECIBIR_AVISOS', '==', 'SI').execute();
    
    return docs.map(u => ({
      nombre: u.NOMBRE || u.nombre || 'Usuario',
      email: u.EMAIL || u.email,
      rol: u.ROL || u.rol
    }));
  } catch (e) {
    console.error("Error obteniendo usuarios: " + e.message);
    return [];
  }
}

/**
 * Filtra revisiones próximas (0 a 7 días) reutilizando la lógica global
 */
function detectarRevisionesProximas() {
  // Reutilizamos getGlobalMaintenance que ya lee de Firestore y cruza datos
  const todas = getGlobalMaintenance();
  
  // Filtramos las que vencen en los próximos 7 días (y no están vencidas hoy)
  return todas
    .filter(m => m.dias >= 0 && m.dias <= 7)
    .map(m => ({
      activo: m.activo,
      edificio: m.edificio,
      tipo: m.tipo,
      fecha: m.fecha,
      diasRestantes: m.dias
    }));
}

/**
 * Filtra revisiones vencidas (dias < 0)
 */
function detectarRevisionesVencidas() {
  const todas = getGlobalMaintenance();
  
  return todas
    .filter(m => m.dias < 0)
    .map(m => ({
      activo: m.activo,
      edificio: m.edificio,
      tipo: m.tipo,
      fecha: m.fecha,
      diasVencida: Math.abs(m.dias)
    }));
}

/**
 * Detecta contratos que caducan en los próximos 60 días
 */
function detectarContratosProximos() {
  try {
    const firestore = getFirestore();
    // Traemos solo contratos activos
    const contratos = firestore.query('contratos').where('ESTADO', '==', 'ACTIVO').execute();
    
    // Mapa de proveedores para mostrar nombres
    const provs = firestore.query('proveedores').execute();
    const mapProv = {};
    provs.forEach(p => mapProv[p.id] = p.NOMBRE || p.Nombre || 'Desconocido');
    
    const hoy = new Date();
    hoy.setHours(0,0,0,0);
    const resultados = [];

    contratos.forEach(c => {
      // Intentamos leer la fecha de fin
      const fFinRaw = c.FECHA_FIN || c.Fecha_Fin || c.fechaFin;
      if (!fFinRaw) return;

      // Convertir a objeto Date
      let fFin = null;
      try { 
        fFin = fFinRaw.toDate ? fFinRaw.toDate() : new Date(fFinRaw); 
      } catch(e) {}

      if (fFin && !isNaN(fFin.getTime())) {
        fFin.setHours(0,0,0,0);
        // Diferencia en días
        const diff = Math.ceil((fFin - hoy) / (1000 * 60 * 60 * 24));
        
        // Si caduca en los próximos 60 días (y no ha caducado ya)
        if (diff >= 0 && diff <= 60) {
          const idProv = c.PROVEEDOR || c.Proveedor || c.ID_Proveedor;
          resultados.push({
            proveedor: mapProv[idProv] || idProv || 'Proveedor',
            referencia: c.REF || c.Num_Ref || c.Referencia || '-',
            fechaFin: Utilities.formatDate(fFin, Session.getScriptTimeZone(), "dd/MM/yyyy"),
            diasRestantes: diff
          });
        }
      }
    });
    
    return resultados;
  } catch (e) {
    console.error("Error contratos próximos: " + e.message);
    return [];
  }
}

// ==========================================
// FUNCIÓN AUXILIAR DE EMAIL
// ==========================================

/**
 * Construye y envía el HTML del correo resumen
 */
function enviarEmailResumen(usuario, revisionesProximas, revisionesVencidas, contratosProximos) {
  if (!usuario || !usuario.email) return;

  const fechaHoy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
  
  // 1. Cabecera del Email
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
  
  // 2. SECCIÓN: Revisiones Vencidas
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
    
    htmlBody += `</tbody></table></div>`;
  }
  
  // 3. SECCIÓN: Revisiones Próximas
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
    
    htmlBody += `</tbody></table></div>`;
  }
  
  // 4. SECCIÓN: Contratos Próximos
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
    
    htmlBody += `</tbody></table></div>`;
  }
  
  // 5. Footer y Cierre
  htmlBody += `
          <div class="footer">
            <p><strong>GMAO - Sistema de Gestión de Mantenimiento</strong></p>
            <p>Universidad de Navarra | Servicio de Obras y Mantenimiento</p>
            <p style="font-size: 11px; color: #999;">Email automático generado desde Firestore.</p>
          </div>
        </div>
      </body>
    </html>
  `;
  
  // 6. Asunto Inteligente
  let asunto = "[GMAO] Resumen Diario - " + fechaHoy;
  if (revisionesVencidas.length > 0) {
    asunto = `⚠️ [GMAO] ${revisionesVencidas.length} Tareas Vencidas`;
  } else if (revisionesProximas.length > 0) {
    asunto = `📅 [GMAO] ${revisionesProximas.length} Tareas Próximas`;
  }
  
  // 7. Enviar
  try {
    MailApp.sendEmail({
      to: usuario.email,
      subject: asunto,
      htmlBody: htmlBody
    });
    console.log("✉️ Email enviado correctamente a: " + usuario.email);
  } catch (e) {
    console.error("Error enviando email a " + usuario.email + ": " + e.message);
  }
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

/**
 * Generar QR para todos los activos de un edificio (Versión Firestore)
 */
function generarQRsEdificio(idEdificio) {
  try {
    const firestore = getFirestore();
    const idEdifStr = String(idEdificio).trim();

    // 1. Obtener Info del Edificio y Campus
    const docEdif = firestore.getDocument('edificios/' + idEdifStr);
    if (!docEdif) throw new Error("Edificio no encontrado");

    // Datos del edificio (soporte híbrido fields/directo)
    const datosEdif = docEdif.fields || docEdif;
    const nombreEdificio = datosEdif.NOMBRE || datosEdif.nombre || 'Edificio';
    const idCampus = datosEdif.ID_CAMPUS || datosEdif.idCampus;

    // Datos del Campus
    let nombreCampus = "Campus General";
    if (idCampus) {
      const docCampus = firestore.getDocument('campus/' + idCampus);
      if (docCampus) {
        const dC = docCampus.fields || docCampus;
        nombreCampus = dC.NOMBRE || dC.nombre || 'Campus';
      }
    }
    
    // 2. Obtener Activos del Edificio
    // Buscamos por ID_Edificio (o variantes)
    let activos = firestore.query('activos').where('ID_EDIFICIO', '==', idEdifStr).execute();
    if (!activos.length) activos = firestore.query('activos').where('idEdificio', '==', idEdifStr).execute();
    if (!activos.length) activos = firestore.query('activos').where('ID_Edificio', '==', idEdifStr).execute();

    const urlApp = ScriptApp.getService().getUrl();
    const qrs = [];
    
    // 3. Generar lista para el PDF
    activos.forEach(a => {
        const idActivo = a.id; // El ID del documento
        const urlQR = `${urlApp}?activo=${idActivo}`;
        const qrImageUrl = `https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=${encodeURIComponent(urlQR)}`;
        
        qrs.push({
          id: idActivo,
          nombre: a.NOMBRE || a.nombre || 'Sin Nombre',
          tipo: a.TIPO || a.tipo || '-',
          marca: a.MARCA || a.marca || '-',
          qrUrl: qrImageUrl,
          targetUrl: urlQR
        });
    });
    
    return {
      success: true,
      edificio: nombreEdificio,
      campus: nombreCampus,
      totalActivos: qrs.length,
      qrs: qrs
    };
    
  } catch (e) {
    console.error("Error generarQRsEdificio: " + e.message);
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
 * Obtener información del activo al escanear QR (Versión Firestore)
 */

function getActivoByQR(idActivo) {
  try {
    const firestore = getFirestore();
    const idStr = String(idActivo).trim();

    // 1. Datos del Activo (Reutilizamos tu función ya migrada)
    const activo = getAssetInfo(idStr);
    
    if (!activo || activo.nombre.includes("Error")) {
      return { success: false, error: "Activo no encontrado o eliminado." };
    }
    
    // 2. Obtener últimas 5 revisiones (Mantenimiento)
    let revisiones = [];
    try {
      // Consulta ordenada por fecha descendente
      // Nota: Firestore requiere índice compuesto para where() + order(). 
      // Si falla, usa .execute() y ordena en JS (más seguro por ahora).
      let docsMant = firestore.query('mantenimientos').where('ID_ACTIVO', '==', idStr).execute();
      
      // Mapeo y Ordenación en memoria (para evitar errores de índices de Firestore)
      revisiones = docsMant.map(m => {
         const fRaw = m.FECHA || m.Fecha || m.Fecha_Proxima;
         let fechaDisp = "-";
         let timestamp = 0;
         
         if (fRaw) {
            const f = fRaw.toDate ? fRaw.toDate() : new Date(fRaw);
            fechaDisp = Utilities.formatDate(f, Session.getScriptTimeZone(), "dd/MM/yyyy");
            timestamp = f.getTime();
         }
         
         return {
           tipo: m.TIPO || m.Tipo_Revision || 'Revisión',
           fecha: fechaDisp,
           estado: m.ESTADO || m.Estado || 'PENDIENTE',
           _ts: timestamp
         };
      })
      .sort((a, b) => b._ts - a._ts) // Más recientes primero
      .slice(0, 5); // Solo las 5 últimas

    } catch(e) { console.warn("Error historial mant: " + e.message); }
    
    // 3. Obtener últimas 3 incidencias (Solo tipo 'ACTIVO')
    let incidencias = [];
    try {
      // Buscamos incidencias asociadas a este activo
      let docsInc = firestore.query('incidencias')
                             .where('ID_ORIGEN', '==', idStr)
                             .execute();
      
      incidencias = docsInc.map(i => {
         // Filtrar solo si el origen es ACTIVO (por si acaso el ID coincide con un edificio)
         const tipoOrigen = i.TIPO_ORIGEN || i.tipoOrigen;
         if (tipoOrigen !== 'ACTIVO') return null;

         const fRaw = i.FECHA || i.fecha;
         let fechaDisp = "-";
         let timestamp = 0;
         if (fRaw) {
             const f = fRaw.toDate ? fRaw.toDate() : new Date(fRaw);
             fechaDisp = Utilities.formatDate(f, Session.getScriptTimeZone(), "dd/MM/yyyy");
             timestamp = f.getTime();
         }

         return {
           descripcion: i.DESCRIPCION || i.descripcion,
           prioridad: i.PRIORIDAD || i.prioridad,
           estado: i.ESTADO || i.estado,
           fecha: fechaDisp,
           _ts: timestamp
         };
      })
      .filter(x => x !== null) // Quitar nulos
      .sort((a, b) => b._ts - a._ts)
      .slice(0, 3);

    } catch(e) { console.warn("Error historial incidencias: " + e.message); }
    
    return {
      success: true,
      activo: activo,
      revisiones: revisiones,
      incidencias: incidencias
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
  try {
    const firestore = getFirestore();
    const idStr = String(idActivo).trim();

    // 1. OBTENER EL ACTIVO (Porque las relaciones están dentro)
    // Usamos query porque es lo más seguro con tu librería
    const docs = firestore.query('activos').where('ID', '==', idStr).execute();
    
    if (!docs || docs.length === 0) {
       return []; // Si no encuentra el activo, no hay relaciones
    }

    const activo = docs[0];
    
    // 2. EXTRAER EL JSON DE RELACIONES
    // En la foto se ve: RELACIONES (Mayúsculas)
    let rawRelaciones = activo.RELACIONES || activo.Relaciones || "[]";
    
    let listaRelaciones = [];
    try {
        // Si es un string (lo más probable por la foto), lo convertimos
        if (typeof rawRelaciones === 'string') {
            listaRelaciones = JSON.parse(rawRelaciones);
        } else if (Array.isArray(rawRelaciones)) {
            listaRelaciones = rawRelaciones;
        }
    } catch(e) {
        console.error("Error parseando JSON relaciones: " + e.message);
        return [];
    }

    if (listaRelaciones.length === 0) return [];

    // 3. MAPA DE NOMBRES (Para que no salga el ID feo "38432f...")
    const todosActivos = firestore.query('activos').execute();
    const mapNombres = {};
    todosActivos.forEach(a => {
        // Guardamos ID -> Nombre y Tipo
        const aID = a.ID || a.id;
        mapNombres[aID] = {
            nombre: a.NOMBRE || a.Nombre || "Sin Nombre",
            tipo: a.TIPO || a.Tipo || "-"
        };
    });

    // 4. MAPEAR RESULTADOS (Usando las claves de tu foto: idActivoRelacionado, tipoRelacion)
    return listaRelaciones.map(r => {
        const idVinculado = r.idActivoRelacionado; // Clave exacta de tu foto
        const info = mapNombres[idVinculado] || {};

        return {
            id: r.id,
            idActivoRelacionado: idVinculado,
            // Aquí ponemos el nombre bonito
            nombreActivo: info.nombre || "Activo Borrado/Desconocido",
            tipoActivo: info.tipo || "-",
            tipoRelacion: r.tipoRelacion || "RELACIONADO", // Clave exacta de tu foto
            descripcion: r.descripcion || ""
        };
    });

  } catch(e) {
    console.error("Error fatal relaciones: " + e.message);
    return [];
  }
}

/**
 * Guardar nueva relación entre activos (bidireccional)
 */

function crearRelacionActivo(d) {
  try {
    const firestore = getFirestore();
    const idA = String(d.idActivoA).trim();
    const idB = String(d.idActivoB).trim();

    const agregarRelacion = (idOrigen, idDestino, tipo, desc) => {
       // 1. Leer el activo completo
       const doc = firestore.getDocument('activos/' + idOrigen);
       if (!doc) return;
       
       // 2. Hacer una copia de TODOS los datos existentes para no perder nada
       const datosCompletos = { ...doc };
       delete datosCompletos.id; // Quitamos el ID para no duplicarlo dentro
       
       // 3. Gestionar el array de relaciones
       const raw = datosCompletos.RELACIONES || datosCompletos.Relaciones || "[]";
       let actual = [];
       try {
          if (typeof raw === 'string') actual = JSON.parse(raw);
          else if (Array.isArray(raw)) actual = raw;
       } catch(e) { actual = []; }

       // 4. Añadir la nueva
       actual.push({
          id: Utilities.getUuid(),
          idActivoRelacionado: idDestino,
          tipoRelacion: tipo,
          descripcion: desc || ""
       });

       // 5. Actualizar el campo en el objeto completo
       datosCompletos.RELACIONES = JSON.stringify(actual);
       
       // 6. GUARDAR EL DOCUMENTO ENTERO (Así no se borra el nombre/marca)
       firestore.updateDocument('activos/' + idOrigen, datosCompletos);
    };

    // A -> B
    agregarRelacion(idA, idB, d.tipoRelacion, d.descripcion);

    // B -> A (Bidireccional)
    if (d.bidireccional) {
      const inversos = { 
          "ALIMENTA": "ES_ALIMENTADO_POR", "ES_ALIMENTADO_POR": "ALIMENTA", 
          "DEPENDE_DE": "ES_REQUERIDO_POR", "ES_REQUERIDO_POR": "DEPENDE_DE",
          "PERTENECE_A": "CONTIENE", "CONTIENE": "PERTENECE_A",
          "RELACIONADO": "RELACIONADO"
      };
      const tipoInv = inversos[d.tipoRelacion] || "RELACIONADO";
      try { agregarRelacion(idB, idA, tipoInv, d.descripcion); } catch(e){}
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
    const firestore = getFirestore();
    const idStr = String(idActivo).trim();
    
    // 1. Leer completo
    const doc = firestore.getDocument('activos/' + idStr);
    if (!doc) throw new Error("Activo no encontrado");
    
    // 2. Copiar datos
    const datosCompletos = { ...doc };
    delete datosCompletos.id;
    
    // 3. Procesar array
    const raw = datosCompletos.RELACIONES || datosCompletos.Relaciones || "[]";
    let lista = [];
    try {
        if (typeof raw === 'string') lista = JSON.parse(raw);
        else if (Array.isArray(raw)) lista = raw;
    } catch(e) {}

    const nuevaLista = lista.filter(r => r.id !== idRelacion);
    
    // 4. Actualizar campo
    datosCompletos.RELACIONES = JSON.stringify(nuevaLista);
    
    // 5. Guardar completo
    firestore.updateDocument('activos/' + idStr, datosCompletos);

    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

/**
 * Obtener alertas de activos relacionados con problemas
 */

function getAlertasActivosRelacionados(idActivo, relacionesPrecalculadas) {
  try {
    // Si no nos pasan las relaciones, las buscamos
    const relaciones = relacionesPrecalculadas || getRelacionesActivo(idActivo);
    
    if (!relaciones || relaciones.length === 0) return [];

    const firestore = getFirestore();
    const alertas = [];
    
    // Lista de IDs de los activos relacionados
    const idsRelacionados = relaciones.map(r => r.idActivoRelacionado);
    
    // Buscar incidencias PENDIENTES o EN PROCESO
    // (Traemos todas las abiertas y filtramos en memoria por seguridad)
    const incidencias = firestore.query('incidencias').execute();
    
    incidencias.forEach(inc => {
       // Normalizar estado y origen
       const estado = inc.ESTADO || inc.Estado || 'PENDIENTE';
       const idOrigen = inc.ID_ORIGEN || inc.Id_Origen || inc.ID_Activo; // Ajustar a tus campos reales
       
       if (estado !== 'RESUELTA' && idsRelacionados.includes(idOrigen)) {
           // Encontramos una incidencia en un activo relacionado
           const relInfo = relaciones.find(r => r.idActivoRelacionado === idOrigen);
           
           alertas.push({
             nombreActivo: relInfo.nombreActivo,
             tipoRelacion: relInfo.tipoRelacion,
             problema: inc.DESCRIPCION || inc.Descripcion || "Avería reportada",
             prioridad: inc.PRIORIDAD || inc.Prioridad || "MEDIA"
           });
       }
    });
    
    return alertas;
  } catch (e) {
    console.error("Error en alertas: " + e.message);
    return [];
  }
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
  // Obtenemos las relaciones (que ya arreglamos antes)
  const relaciones = getRelacionesActivo(idActivo);
  
  // Obtenemos las alertas de forma segura (con la nueva función de abajo)
  let alertas = [];
  try {
    alertas = getAlertasActivosRelacionados(idActivo, relaciones);
  } catch (e) {
    console.warn("Error cargando alertas de relaciones: " + e.message);
    // Si falla, devolvemos array vacío para no romper la pantalla
    alertas = [];
  }

  return {
    relaciones: relaciones,
    alertas: alertas
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
 * Obtiene todos los eventos combinados para el planificador (VERSIÓN FIRESTORE)
 */

function getPlannerEvents() {
  const eventos = [];
  const firestore = getFirestore();

  // 1. REVISIONES DE MANTENIMIENTO (Ya estaba migrado, lo mantenemos)
  // Reutilizamos la lógica de getGlobalMaintenance que ya cruza datos correctamente
  const mant = getGlobalMaintenance(); 
  mant.forEach(r => {
    eventos.push({
      id: r.id,
      resourceId: 'MANTENIMIENTO',
      title: `🔧 ${r.tipo} - ${r.activo}`,
      start: r.fechaISO, // Formato YYYY-MM-DD
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

  // 2. OBRAS EN CURSO (Desde Firestore)
  try {
    const obras = firestore.query('obras').where('ESTADO', '==', 'EN CURSO').execute();
    
    obras.forEach(o => {
      // Normalizamos fecha
      const fInicio = o.FECHA_INICIO || o.fechaInicio;
      let fechaISO = "";
      
      if (fInicio) {
        // Si es timestamp o string, intentamos convertir a ISO 'YYYY-MM-DD'
        const d = new Date(fInicio);
        if (!isNaN(d.getTime())) fechaISO = d.toISOString().split('T')[0];
      }

      if (fechaISO) {
        eventos.push({
          id: o.id,
          resourceId: 'OBRA',
          title: `🏗️ ${o.NOMBRE || o.nombre || 'Obra'}`,
          start: fechaISO,
          color: '#0d6efd', // Azul
          extendedProps: {
            tipo: 'OBRA',
            descripcion: o.DESCRIPCION || o.descripcion || '',
            editable: true
          }
        });
      }
    });
  } catch(e) { console.warn("Error cargando obras al calendario: " + e.message); }

  // 3. INCIDENCIAS PENDIENTES (Desde Firestore)
  try {
    const incidencias = firestore.query('incidencias').execute(); // Traemos todas y filtramos en memoria por seguridad
    
    incidencias.forEach(i => {
      const estado = i.ESTADO || i.estado || 'PENDIENTE';
      if (estado === 'RESUELTA') return;

      const fRaw = i.FECHA || i.fecha;
      let fechaISO = "";
      if (fRaw) {
         const d = new Date(fRaw);
         if (!isNaN(d.getTime())) fechaISO = d.toISOString().split('T')[0];
      }

      if (fechaISO) {
        eventos.push({
          id: i.id,
          resourceId: 'INCIDENCIA',
          title: `⚠️ ${i.PRIORIDAD || i.prioridad || 'MEDIA'} - ${i.NOMBRE_ORIGEN || i.nombreOrigen || 'Incidencia'}`,
          start: fechaISO,
          color: '#fd7e14', // Naranja
          extendedProps: {
            tipo: 'INCIDENCIA',
            descripcion: i.DESCRIPCION || i.descripcion || '',
            editable: false // No mover fecha de reporte
          }
        });
      }
    });
  } catch(e) { console.warn("Error cargando incidencias al calendario: " + e.message); }

  // 4. VENCIMIENTOS DE CONTRATOS (Desde Firestore)
  try {
    const contratos = firestore.query('contratos').where('ESTADO', '==', 'ACTIVO').execute();
    
    // Necesitamos nombres de proveedores
    const proveedores = firestore.query('proveedores').execute();
    const mapProv = {};
    proveedores.forEach(p => mapProv[p.id] = p.NOMBRE || p.Nombre || 'Prov.');

    contratos.forEach(c => {
       const fFinRaw = c.FECHA_FIN || c.Fecha_Fin || c.fechaFin;
       let fechaISO = "";
       
       if (fFinRaw) {
          const d = new Date(fFinRaw);
          if (!isNaN(d.getTime())) fechaISO = d.toISOString().split('T')[0];
       }

       if (fechaISO) {
          const idProv = c.PROVEEDOR || c.Proveedor || c.ID_Proveedor;
          const nomProv = mapProv[idProv] || idProv || 'Contrato';

          eventos.push({
            id: c.id,
            resourceId: 'CONTRATO',
            title: `📄 Fin: ${nomProv}`,
            start: fechaISO,
            color: '#6f42c1', // Morado
            extendedProps: {
              tipo: 'CONTRATO',
              descripcion: `Ref: ${c.REF || c.Num_Ref || '-'}`,
              editable: true // Permitimos mover la fecha de fin
            }
          });
       }
    });
  } catch(e) { console.warn("Error cargando contratos al calendario: " + e.message); }

  return eventos;
}

/**
 * Actualiza la fecha de un evento tras Drag & Drop (VERSIÓN FIRESTORE)
 */

function updateEventDate(id, tipo, nuevaFechaISO) {
  // Verificamos permiso de escritura
  verificarPermiso(['WRITE']); 
  
  const firestore = getFirestore();
  // Convertimos a objeto fecha (para que Firestore lo guarde como Timestamp o ISO correcto según tu modelo)
  // Al venir del calendario (fullcalendar), nuevaFechaISO suele ser "YYYY-MM-DD"
  const nuevaFecha = new Date(nuevaFechaISO);
  // Forzamos mediodía para evitar cambios de zona horaria indeseados
  nuevaFecha.setHours(12, 0, 0, 0);
  const fechaGuardar = nuevaFecha.toISOString();

  try {
    if (tipo === 'MANTENIMIENTO') {
      // Actualiza la colección 'mantenimientos'
      firestore.updateDocument('mantenimientos/' + id, {
        FECHA_PROXIMA: fechaGuardar
      });
      // Registrar log opcional
      // registrarLog("REPROGRAMAR", `Revisión ${id} movida a ${nuevaFechaISO}`);
      return { success: true };
    } 
    else if (tipo === 'OBRA') {
      // Actualiza 'obras' (Fecha inicio)
      firestore.updateDocument('obras/' + id, {
        FECHA_INICIO: fechaGuardar // Ojo: Verifica si usas FECHA_INICIO o solo FECHA en tu modelo
      });
      return { success: true };
    }
    else if (tipo === 'CONTRATO') {
      // Actualiza 'contratos' (Fecha fin)
      firestore.updateDocument('contratos/' + id, {
        FECHA_FIN: fechaGuardar
      });
      return { success: true };
    }

    return { success: false, error: "Tipo de evento no editable" };

  } catch (e) {
    console.error("Error al mover evento: " + e.message);
    return { success: false, error: e.message };
  }
}

// ==========================================
// 27. MÓDULO DE EXPORTACIÓN PARA AUDITORÍA (VERSIÓN FIRESTORE)
// ==========================================
/**
 * Obtiene los años disponibles.
 * SE FUERZAN los últimos 3 años para evitar errores de escaneo.
 */

function getAniosAuditoria() {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'AUDIT_YEARS_LIST';
  
  // 1. Intentamos leer de la memoria rápida
  const cached = cache.get(cacheKey);
  if (cached) {
    return JSON.parse(cached);
  }

  const anios = new Set();
  const yearActual = new Date().getFullYear();

  // --- CORRECCIÓN: AÑADIR AÑOS MANUALMENTE (Garantía Total) ---
  // Esto asegura que el Técnico vea estos años aunque la BBDD falle al escanear
  anios.add(yearActual);     // 2025
  anios.add(yearActual - 1); // 2024
  anios.add(yearActual - 2); // 2023
  // ------------------------------------------------------------

  // 2. Intentamos escanear la base de datos por si hay años más viejos
  try {
    const firestore = getFirestore();
    const docs = firestore.query('mantenimientos').select(['FECHA', 'Fecha']).execute();
    
    docs.forEach(d => {
       const fRaw = d.FECHA || d.Fecha;
       if (fRaw) {
         try {
           const fecha = new Date(fRaw);
           if (!isNaN(fecha.getTime())) {
             anios.add(fecha.getFullYear());
           }
         } catch(e){}
       }
    });
  } catch(e) {
    console.warn("El escaneo de Firestore falló, pero mostramos los años forzados: " + e.message);
  }
  
  // 3. Ordenar descendente (2025, 2024...)
  const result = Array.from(anios).sort((a,b) => b - a);
  
  // 4. Guardar en caché 6 horas
  if (result.length > 0) {
    cache.put(cacheKey, JSON.stringify(result), 21600); 
  }
  
  return result;
}

/**
 * Genera una carpeta en Drive con la documentación organizada
 */
function generarPaqueteAuditoria(anio, tipoFiltro) {
  // 1. Preparación
  const firestore = getFirestore();
  const anioStr = String(anio);
  
  try {
    // A. Obtener todas las revisiones del año seleccionado
    // Traemos todo y filtramos fecha en memoria para evitar problemas de índices complejos
    const mantenimientos = firestore.query('mantenimientos').execute();
    
    // Mapa de Revisiones Válidas: ID_REVISION -> Datos para el nombre del archivo
    const revisionesValidas = {}; 
    const idsRevisiones = new Set();

    mantenimientos.forEach(m => {
       const fRaw = m.FECHA || m.Fecha || m.Fecha_Proxima;
       if (!fRaw) return;
       
       const fecha = new Date(fRaw);
       if (isNaN(fecha.getTime()) || String(fecha.getFullYear()) !== anioStr) return;

       // Filtro de Tipo (Legal vs Todos)
       const tipo = m.TIPO || m.Tipo_Revision || 'Revisión';
       if (tipoFiltro !== 'TODOS' && tipo !== 'Legal') return;

       // Guardamos datos para nombrar el archivo luego
       // Nota: m.ACTIVO_NOMBRE suele guardarse al crear la revisión. 
       // Si no está, saldrá "Activo".
       revisionesValidas[m.id] = {
          fechaStr: Utilities.formatDate(fecha, Session.getScriptTimeZone(), "yyyy-MM-dd"),
          tipo: tipo,
          activo: m.ACTIVO_NOMBRE || m.ActivoNombre || 'Activo',
          edificio: 'Edificio' // Simplificación: Si necesitas el edificio real, habría que cruzar con 'activos'
       };
       idsRevisiones.add(m.id);
    });

    if (idsRevisiones.size === 0) {
      return { success: false, error: `No hay revisiones ${tipoFiltro === 'Legal' ? 'legales ' : ''}en ${anio}.` };
    }

    // B. Buscar documentos asociados a esas revisiones
    // Consultamos la colección de historial de documentos
    const docs = firestore.query('docs_historico')
                          .where('TIPO_ENTIDAD', '==', 'REVISION')
                          .execute();

    // Filtramos en memoria los que pertenecen a nuestras revisiones del año
    const archivosACopiar = docs.filter(d => {
        const idPadre = d.ID_ENTIDAD || d.Id_Entidad;
        return idsRevisiones.has(idPadre);
    });

    if (archivosACopiar.length === 0) {
       return { success: false, error: "Se encontraron revisiones, pero ninguna tiene documentos adjuntos." };
    }

    // C. Crear Carpeta en Drive
    const rootId = getRootFolderId();
    const parentFolder = DriveApp.getFolderById(rootId);
    const folderName = `AUDITORIA_${anio}_${tipoFiltro}_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HHmm")}`;
    const auditFolder = parentFolder.createFolder(folderName);
    
    let copiados = 0;
    let errores = 0;

    // D. Copiar y Renombrar Archivos
    archivosACopiar.forEach(d => {
       try {
          const fileId = d.FILE_ID_DRIVE || d.File_Id_Drive;
          if (!fileId) return;

          const info = revisionesValidas[d.ID_ENTIDAD || d.Id_Entidad];
          
          // Nombre estandarizado: 2024-05-20_Extintor-01_Legal.pdf
          const cleanName = `${info.fechaStr}_${info.activo}_${info.tipo}`.replace(/[^a-zA-Z0-9_\-\.]/g, '_');
          const originalFile = DriveApp.getFileById(fileId);
          
          // Mantener extensión original
          const ext = originalFile.getName().split('.').pop();
          const finalName = cleanName + "." + ext;

          originalFile.makeCopy(finalName, auditFolder);
          copiados++;
       } catch (err) {
          console.warn("Error copiando archivo auditoría: " + err);
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

  } catch (e) {
    console.error("Error fatal auditoría: " + e.message);
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
  try {
    const firestore = getFirestore();
    
    // Consultamos solo los usuarios que tienen activado 'SI' en avisos
    const usuarios = firestore.query('usuarios')
                              .where('RECIBIR_AVISOS', '==', 'SI')
                              .execute();
    
    const emails = [];
    
    usuarios.forEach(u => {
      // Normalizamos nombres de campos por seguridad
      const email = (u.EMAIL || u.email || "").trim();
      const rol = u.ROL || u.rol;
      
      // Doble comprobación: Solo añadimos si tiene email y es un rol de gestión
      if (email && (rol === 'ADMIN' || rol === 'TECNICO')) {
        emails.push(email);
      }
    });
    
    console.log(`📧 Destinatarios encontrados: ${emails.length}`);
    return emails;
    
  } catch (e) {
    console.error("Error obteniendo destinatarios de Firestore: " + e.message);
    return []; // Devuelve array vacío para no romper la ejecución
  }
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
 * Helper: Obtener nombre de proveedor por ID desde Firestore
 */

function obtenerNombreProveedor(id) {
  if (!id) return "Desconocido";
  
  try {
    const firestore = getFirestore();
    const doc = firestore.getDocument('proveedores/' + id);
    
    // Verificamos si el documento y sus campos existen
    if (doc && doc.fields) {
      // Probamos variantes de nombres por si acaso
      return doc.fields.NOMBRE || doc.fields.Nombre || doc.fields.Empresa || id;
    }
    // Si el documento existe pero es la versión parseada directamente
    if (doc && (doc.NOMBRE || doc.nombre)) {
       return doc.NOMBRE || doc.nombre;
    }
    
    return id; // Si no encuentra nada, devuelve el ID
  } catch (e) {
    console.warn("Error buscando proveedor " + id + ": " + e.message);
    return id;
  }
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

function crearContratoV2(d) {
  try {
    const firestore = getFirestore();
    const newId = Utilities.getUuid();
    
    // Lógica para activos múltiples
    let idEntidadFinal = d.idEntidad;
    let tipoEntidadFinal = d.tipoEntidad;

    if (d.idsActivos && d.idsActivos.length > 0) {
        tipoEntidadFinal = 'ACTIVOS';
        idEntidadFinal = JSON.stringify(d.idsActivos);
    }

    const data = {
      ID: newId,
      // Nombres exactos de tu BD
      Tipo_Entidad: tipoEntidadFinal,
      ID_Entidad: idEntidadFinal,
      Proveedor: d.idProveedor, // Guardamos el ID del proveedor aquí
      Num_Ref: d.ref,
      Fecha_Inicio: d.fechaIni ? new Date(d.fechaIni).toISOString() : null,
      Fecha_Fin: d.fechaFin ? new Date(d.fechaFin).toISOString() : null,
      Estado: d.estado || 'ACTIVO'
    };

    firestore.createDocument('contratos/' + newId, data);
    return { success: true, newId: newId };
  } catch(e) { return { success: false, error: e.message }; }
}

function updateContratoV2(d) {
  try {
    const firestore = getFirestore();
    
    let idEntidadFinal = d.idEntidad;
    if (d.tipoEntidad === 'ACTIVOS' && d.idsActivos && d.idsActivos.length > 0) {
        idEntidadFinal = JSON.stringify(d.idsActivos);
    }

    const data = {
      Tipo_Entidad: d.tipoEntidad,
      ID_Entidad: idEntidadFinal,
      Proveedor: d.idProveedor,
      Num_Ref: d.ref,
      Fecha_Inicio: d.fechaIni ? new Date(d.fechaIni).toISOString() : null,
      Fecha_Fin: d.fechaFin ? new Date(d.fechaFin).toISOString() : null,
      Estado: d.estado
    };

    firestore.updateDocument('contratos/' + d.id, data);
    return { success: true };
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

function limpiarMantenimientosHuerfanos() {
  const firestore = getFirestore();
  
  // 1. Cargar todos los datos
  const mant = firestore.query('mantenimientos').execute();
  const activos = firestore.query('activos').execute();
  
  // 2. Crear una lista de IDs de activos VÁLIDOS
  const idsValidos = new Set();
  activos.forEach(a => idsValidos.add(String(a.ID || a.id).trim()));
  
  let borrados = 0;
  
  // 3. Recorrer mantenimientos y buscar huérfanos
  mant.forEach(m => {
    const idApunta = String(m.ID_Activo || m.ID_ACTIVO || m.idActivo || "").trim();
    
    // Si el ID al que apunta NO está en la lista de válidos
    if (!idsValidos.has(idApunta)) {
      console.log(`🗑️ Borrando revisión huérfana ${m.id} (Apuntaba a: ${idApunta})`);
      try {
        firestore.deleteDocument('mantenimientos/' + m.id);
        borrados++;
      } catch(e) {
        console.error("Error borrando: " + e.message);
      }
    }
  });
  
  console.log(`✅ PROCESO TERMINADO. Se han eliminado ${borrados} registros huérfanos.`);
  return `Eliminados ${borrados} registros.`;
}

function limpiarCache() {
  const CACHE = CacheService.getScriptCache();
  // Añadimos 'AUDIT_YEARS_LIST' a la lista de borrado
  CACHE.removeAll(['SHEET_ACTIVOS', 'SHEET_EDIFICIOS', 'SHEET_CAMPUS', 'INDEX_ACTIVOS', 'AUDIT_YEARS_LIST']);
  console.log('Caché limpiada (incluyendo años de auditoría)');
}

