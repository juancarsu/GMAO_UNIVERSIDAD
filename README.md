Manual de Usuario - GMAO Universidad de Navarra
📋 Índice

Introducción
Acceso al Sistema
Permisos y Roles
Navegación Principal
Gestión de Campus y Edificios
Gestión de Activos
Plan de Mantenimiento
Gestión de Contratos
Incidencias
Códigos QR
Funciones Adicionales


Introducción
El GMAO (Sistema de Gestión de Mantenimiento Asistido por Ordenador) de la Universidad de Navarra es una aplicación web diseñada para gestionar de forma integral todos los activos, instalaciones y mantenimientos de los diferentes campus universitarios.
Características Principales

✅ Gestión centralizada de activos e instalaciones
✅ Programación automática de mantenimientos legales y periódicos
✅ Control de contratos con proveedores
✅ Reportes de incidencias con fotografías
✅ Códigos QR para identificación rápida
✅ Alertas automáticas por email
✅ Historial completo de documentación


Acceso al Sistema
Primer Acceso

Abra el enlace proporcionado por el administrador
Inicie sesión con su cuenta de Google corporativa (@unav.es)
El sistema detectará automáticamente sus permisos

Interfaz Principal
La aplicación se divide en tres zonas:
Barra Lateral Izquierda (Menú)

Dashboard
Campus
Edificios
Activos
Mantenimiento
Incidencias
Contratos
Planificador
Proveedores
Configuración (solo Admin)

Contenido Central
Muestra la información según la sección seleccionada
Botones Flotantes (esquina inferior derecha)

💬 Botón azul: Enviar sugerencias/reportar errores
🔴 Botón rojo: Reportar avería urgente


Permisos y Roles
El sistema tiene tres niveles de acceso:
👁️ CONSULTA (Solo lectura)

✅ Ver toda la información
✅ Descargar documentos
✅ Reportar averías
❌ No puede crear ni modificar

🔧 TÉCNICO (Operativo)

✅ Todo lo de Consulta
✅ Crear y editar activos
✅ Programar mantenimientos
✅ Subir documentación
✅ Gestionar contratos
❌ No puede eliminar registros
❌ No puede gestionar usuarios

👑 ADMINISTRADOR (Control total)

✅ Acceso completo
✅ Eliminar registros
✅ Gestionar usuarios
✅ Configurar catálogo de instalaciones
✅ Ver logs de auditoría


Navegación Principal
📊 Dashboard
El Dashboard muestra un resumen general del estado del sistema:
Tarjetas de Estado

Activos: Número total de equipos registrados
Vencidas: Revisiones que no se han realizado a tiempo (🔴 rojo)
Pendientes: Revisiones próximas a vencer en 30 días (🟡 amarillo)
Incidencias: Averías sin resolver
Contratos: Contratos vigentes
Edificios y Campus: Total de ubicaciones


💡 Truco: Haga clic en cualquier tarjeta para ir directamente a esa sección

Gráfico de Evolución
Muestra el número de mantenimientos programados para los próximos 6 meses
Calendario

Verde: Mantenimiento al día
Amarillo: Próximo a vencer (≤30 días)
Rojo: Vencido

Acciones del calendario:

Clic en una fecha vacía: Programar nueva revisión
Clic en un evento: Ver detalles o editarlo

Filtro por Campus
Use el selector superior derecho para filtrar los datos por campus específico.

Gestión de Campus y Edificios
🏛️ Campus
Ver Lista de Campus

Haga clic en "Campus" en el menú lateral
Verá la lista completa con:

Nombre del campus
Provincia
Dirección



Crear Nuevo Campus

Clic en "+ Nuevo Campus"
Complete los campos:

Nombre: Ej. "Campus de Pamplona"
Provincia: Ej. "Navarra"
Dirección: Dirección completa


Clic en "Guardar"


⚠️ Importante: Al crear un campus, se crea automáticamente una carpeta en Google Drive para almacenar documentación

Editar Campus

Clic en el botón ✏️ (lápiz) junto al campus
Modifique los datos necesarios
Clic en "Guardar"


🏢 Edificios
Ver Edificios
Vista Lista (por defecto)

Muestra tabla con todos los edificios
Columnas: Edificio, Campus, Contacto

Vista Mapa

Clic en el botón "Mapa" (superior derecha)
Ver ubicación geográfica de los edificios
Clic en un marcador para ver detalles

Crear Nuevo Edificio

Clic en "+ Nuevo Edificio"
Complete:

Nombre: Ej. "Edificio Amigos"
Campus: Seleccione el campus correspondiente
Contacto: Responsable o persona de contacto
Latitud/Longitud (opcional): Para visualización en mapa




💡 Truco: Para obtener coordenadas:

Abra Google Maps
Clic derecho sobre el edificio
Copie las coordenadas que aparecen



Clic en "Guardar"

Ficha de Edificio
Al hacer clic en el ojo (👁️) de un edificio, accede a su ficha completa con pestañas:
📋 Información

Datos básicos del edificio
Campus al que pertenece
Contacto

📁 Documentación Legal/Planos

Licencias de actividad
Planos arquitectónicos
Certificados energéticos

Para subir documentos:

Clic en "Seleccionar archivo"
Elegir el archivo (PDF, imagen, etc.)
Clic en "Subir Documento"

🏗️ Obras y Reformas

Historial de obras realizadas
Estado: "En curso" o "Finalizada"

Para registrar una obra:

Clic en "+ Nueva Obra"
Complete:

Nombre: Ej. "Reforma cubierta norte"
Descripción: Detalles de la obra
Fecha inicio


Clic en "Crear"

🔧 Activos Instalados

Lista todos los equipos del edificio
Filtro por tipo de instalación
Acceso directo a la ficha de cada activo


Gestión de Activos
📦 Activos
Los activos son todos los equipos e instalaciones: calderas, ascensores, aire acondicionado, cuadros eléctricos, etc.
Ver Activos
Sistema de Filtros (recomendado):

Paso 1: Seleccione Campus
Paso 2: Seleccione Edificio
Aparecerán todos los activos de ese edificio
Use el buscador para filtrar por nombre, marca o tipo

Crear Nuevo Activo
Método Manual

Clic en "+ Crear Activo"
Complete:

Campus: Seleccione ubicación
Edificio: Seleccione edificio específico
Tipo: Elija del catálogo (Ej. "Caldera Gas Natural")
Nombre: Identificador único (Ej. "Caldera Edif. Central Planta -1")
Marca: Fabricante (Ej. "Junkers")


Clic en "Guardar"


✅ Se crea automáticamente una carpeta en Drive para este activo

Método Masivo (Importación)
Para dar de alta muchos activos a la vez:

Clic en "Importar"
Prepare sus datos en Excel con estas columnas:

   Campus | Edificio | Tipo | Nombre | Marca

Copie las filas de Excel (sin cabeceras)
Pegue en el cuadro de texto de la app
Clic en "Procesar Importación"


⚠️ Los nombres de Campus y Edificios deben coincidir exactamente con los registrados

Ficha Completa de Activo
Al hacer clic en "Ver Ficha" de un activo:
📋 Pestaña Información

Datos del activo
Edificio y campus donde está
Fecha de alta
Botón "Editar" para modificar

📁 Pestaña Documentación
Aquí se guardan:

Manuales de usuario
Certificados de instalación
Fichas técnicas
Certificados OCA/legalizaciones

Subir un documento:

Clic en "Seleccionar archivo"
Elegir archivo
Clic en "Subir"

📤 Subida Rápida (botón nube ☁️):

Permite subir varios archivos a la vez
Clasifica automáticamente si detecta "OCA", "contrato", etc.
Para OCA/Legalizaciones:

Indica la fecha de inspección
Marca "Programar siguientes" para crear revisiones automáticas
Establece frecuencia (ej. 365 días)



🔧 Pestaña Mantenimiento
Ver plan de mantenimiento:

🟢 Verde: Al día
🟡 Amarillo: Próxima (≤30 días)
🔴 Rojo: Vencida

Programar nueva revisión:

Clic en "+ Programar Revisión"
Seleccione tipo:

Legal: Reglamentaria (OCA, RITE, etc.)
Periódica: Mantenimiento preventivo
Reparación: Correctivo
Extraordinaria: Puntual


Si es Legal:

Seleccione normativa del desplegable
Se autocompletar frecuencia
Active "Repetir esta revisión" para crear futuras automáticamente


Indique fecha próxima
Sincronizar con Google Calendar (opcional)
Adjunte evidencia/certificado si ya lo tiene
Clic en "Guardar"

Marcar como realizada:

Clic en botón ✅ (check verde)
Confirmar


💡 La revisión pasará a "Histórico" (azul) y no aparecerá en alertas

📄 Pestaña Contratos
Lista de contratos de mantenimiento asociados al activo
🔗 Pestaña Relaciones
Define dependencias entre activos (Ej. "Caldera → Alimenta → Radiadores")

Plan de Mantenimiento
🔧 Vista Global de Mantenimiento
Acceso: Menú "Mantenimiento"
¿Qué muestra?
Tabla completa con todas las revisiones programadas de todos los activos
Sistema de Filtros
Filtros de Ubicación

Campus: Filtra por campus específico
Edificio: Filtra por edificio (tras seleccionar campus)

Filtros de Estado (botones)

Todas: Muestra todas excepto históricas
Vencidas (🔴): Solo las que ya pasaron su fecha
Próximas (🟡): Vencen en ≤30 días
Al día (🟢): Bien de fecha
Histórico (🔵): Ya realizadas (para consulta)

Filtros de Tipo

Todos
Legal: Reglamentarias
Periódica: Preventivas
Reparación: Correctivas
Extraordinaria: Puntuales

Acciones Disponibles
☁️ Subida Rápida
Permite adjuntar documentación sin entrar en la ficha del activo
✏️ Editar

Modificar fecha
Cambiar tipo
Adjuntar documentos

✅ Completar
Marca la revisión como realizada (desaparece de pendientes)
Generar Informe PDF
Informe de Revisiones Legales

Clic en "Informe Legal PDF" (esquina superior derecha)
Se descarga un PDF con:

Todas las revisiones reglamentarias
Estado actual
Ideal para auditorías externas




Gestión de Contratos
📑 Contratos
Acceso: Menú "Contratos"
Ver Contratos
La tabla muestra:

Estado:

🟢 Vigente: Contrato activo
🟡 Próximo: Caduca en ≤90 días
🔴 Caducado: Ya venció
⚪ Inactivo: Desactivado manualmente


Activo/Edificio: A qué se aplica
Proveedor: Empresa contratada
Referencia: Nº de contrato
Vigencia: Fechas inicio - fin

Filtros

Campus/Edificio: Filtrar por ubicación
Estado: Vigente, Próximo, Caducado, Inactivo

Crear Nuevo Contrato

Clic en "+ Nuevo Contrato"

Paso 1: Proveedor

Seleccione de la lista o clic en "+ Nuevo" para crear uno

Paso 2: Ubicación (¿A qué aplica?)
Opción A - Todo el Campus: No seleccione edificio ni activo
Opción B - Todo un Edificio:

Seleccione campus
Seleccione edificio
No seleccione activo específico

Opción C - Un Activo Concreto:

Seleccione campus
Seleccione edificio
Seleccione el activo en el desplegable

Opción D - Varios Activos (mantenimiento múltiple):

Seleccione campus y edificio
Active la casilla "Aplicar a varios activos del edificio"
Marque los activos en la lista
Contador muestra cuántos lleva seleccionados

Paso 3: Datos del Contrato

Referencia: Nº de contrato (Ej. "CTR-2025-001")
Estado: Activo / Inactivo
Fecha Inicio: DD/MM/AAAA
Fecha Fin: DD/MM/AAAA

Paso 4: Adjuntar PDF (opcional)
Suba el documento del contrato firmado

Clic en "Guardar Contrato"

Gestión de Proveedores
Para acceder a la lista completa: Menú "Proveedores"
Crear Proveedor

Clic en "+ Nuevo Proveedor"
Complete:

Nombre Empresa: Obligatorio
CIF/NIF
Persona de Contacto
Teléfono
Email
Estado: Activo/Inactivo


Guardar


💡 Proveedores inactivos no aparecen al crear contratos nuevos, pero se mantienen en los existentes


Incidencias
⚠️ Sistema de Incidencias
Reportar una Avería
Método 1: Botón Flotante Rojo (acceso rápido)

Clic en el botón 🔴 (esquina inferior derecha)
Complete el formulario:

Ubicación:

Campus
Edificio
Activo específico (opcional)


Descripción: Detalle del problema
Prioridad:

🟢 Baja
🔵 Media
🟠 Alta
🔴 ¡Urgente!


Foto (opcional): Adjunte imagen del problema


Clic en "Enviar Reporte"


📧 Se envía automáticamente un email a todos los técnicos y administradores con avisos activados

Método 2: Desde Ficha de Activo
Si está viendo un activo específico, el botón flotante rojo pre-selecciona automáticamente ese activo
Ver Incidencias
Acceso: Menú "Incidencias"
Filtros de Estado

Todas: Todas las incidencias
Pendientes: Sin resolver
Resueltas: Ya cerradas

Estados Posibles

🔴 Pendiente: Recién creada
🔵 En Proceso: Ya se está trabajando
🟢 Resuelta: Cerrada

Gestionar Incidencia (Técnicos/Admin)
Cambiar Estado

▶️ En Proceso: Marca que ya se está atendiendo
✅ Resolver: Cierra la incidencia

Editar

Modificar descripción o prioridad
Ver foto adjunta (si existe)


Códigos QR
📱 Sistema de Identificación Rápida
Los códigos QR permiten acceso instantáneo a la ficha de un activo desde el móvil.
Generar QR de un Activo Individual

Entre en la ficha del activo
Clic en "Descargar QR" (botón superior derecho)
Se descarga una imagen PNG
Imprímala y pégueal en el equipo físico

Generar QR de Todo un Edificio (PDF)
Desde la ficha del edificio:

Botón "QR Edificio (PDF)"
Se genera un PDF con:

Etiquetas de todos los activos del edificio
2 columnas por página
Listas para imprimir y etiquetar




💡 Uso recomendado: Imprimir en papel adhesivo o plastificar

Escanear un Código QR
Desde el móvil:

Abra la cámara del teléfono
Enfoque el código QR
Toque el enlace que aparece
Se abre la Vista Móvil Optimizada:

✅ Datos del activo
✅ Estado de mantenimiento
✅ Incidencias abiertas
✅ Botones grandes para:

🔴 Reportar avería
🟢 Realizar revisión
📄 Ver manuales
📋 Historial completo






Funciones Adicionales
📅 Planificador
Acceso: Menú "Planificador"
Vista de Calendario Unificada que muestra:

🔧 Mantenimientos programados
🏗️ Obras en curso
⚠️ Incidencias pendientes
📄 Vencimientos de contratos

Funciones

Clic en evento: Ver detalles o editar
Arrastrar evento: Cambiar fecha (solo mantenimientos y contratos)
Filtros: Mostrar/ocultar categorías con las casillas superiores

📊 Auditoría (Exportación Masiva)
Acceso: Menú "Auditoría"
Para auditorías externas o inspecciones:

Seleccione año
Seleccione tipo:

🔍 Solo Revisiones Legales
📂 Toda la documentación


Clic en "Generar Paquete"

Resultado:

Se crea una carpeta en Google Drive
Contiene copia de todos los certificados/evidencias
Archivos renombrados automáticamente:

YYYY-MM-DD_Edificio_Activo_Tipo.pdf


Ideal para entregas a organismos reguladores

🎁 Novedades de la App
Acceso: Menú "Novedades" (icono regalo)

Historial de versiones
Nuevas funcionalidades añadidas
Correcciones de errores

💬 Buzón de Sugerencias
Botón Flotante Azul (esquina inferior derecha)
Para enviar:

💡 Ideas de mejora
🐛 Reportes de errores
💬 Comentarios generales


📬 Los administradores reciben y gestionan todas las sugerencias

🔔 Alertas Automáticas por Email
Para Técnicos y Administradores:
Si tiene activadas las alertas en su perfil, recibirá emails diarios con:

⚠️ Revisiones vencidas
📅 Revisiones próximas (≤7 días)
📄 Contratos próximos a caducar (≤60 días)

Gestionar preferencias:

Contacte con el Administrador
Solicite activar/desactivar avisos automáticos

📂 Importación Masiva
Ya explicada en Gestión de Activos, pero permite:

Migrar datos desde Excel antiguo
Dar de alta decenas de activos en segundos
Evita tener que crear uno por uno manualmente


🆘 Resolución de Problemas
El sistema no carga

Verifique su conexión a internet
Cierre y vuelva a abrir la pestaña
Borre caché del navegador:

Chrome: Ctrl + Shift + Supr
Seleccione "Imágenes y archivos en caché"


Si persiste, contacte con el administrador

No puedo crear/editar
Probablemente tiene permisos de Solo Consulta. Contacte con el administrador para solicitar permisos de Técnico.
No encuentro un activo
Use los filtros en este orden:

Seleccione Campus
Seleccione Edificio
Use el buscador (busca en nombre, tipo y marca)

La búsqueda global no funciona
La búsqueda del menú lateral requiere mínimo 3 caracteres para activarse.
Un documento no se sube

Verifique que el archivo no supere 5 MB
Formatos aceptados: PDF, JPG, PNG, XLSX, DOCX
Verifique su conexión a internet


📞 Soporte
Administrador del Sistema:

Email: jcsuarez@unav.es
Departamento: Servicio de Obras y Mantenimiento

Para solicitar:

✅ Cambio de permisos
✅ Alta de nuevos usuarios
✅ Resolución de incidencias técnicas
✅ Formación adicional


📝 Consejos Finales
✅ Use los códigos QR: Ahorra muchísimo tiempo en el trabajo de campo
✅ Suba las OCAs con la función rápida: Programa automáticamente las siguientes revisiones
✅ Active las alertas por email: No se le pasará ningún mantenimiento
✅ Use el Planificador: Visión global de toda la carga de trabajo
✅ Reporte todas las averías: Aunque sean pequeñas, ayuda a detectar patrones
✅ Revise el Dashboard regularmente: Los números en rojo necesitan atención urgente

Versión del Manual: 1.0
Última Actualización: Diciembre 2024
Sistema GMAO - Universidad de Navarra
