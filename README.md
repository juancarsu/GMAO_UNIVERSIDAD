📘 Manual de Usuario - GMAO Universidad
Sistema de Gestión de Mantenimiento Asistido por Ordenador

1. Introducción
Bienvenido a la aplicación de gestión de mantenimiento de la Universidad. Esta herramienta permite gestionar de forma centralizada los activos, edificios, planes de mantenimiento, incidencias y documentación legal de todos los campus.

La aplicación es accesible vía web y está integrada con Google Drive (documentos), Google Calendar (agenda) y Gmail (notificaciones).

2. Acceso y Roles
Para acceder, simplemente abra el enlace proporcionado en su navegador. El sistema detectará automáticamente su identidad mediante su cuenta de Google corporativa.

Existen tres niveles de acceso:

Administrador (ADMIN): Control total. Puede crear, editar, borrar y gestionar usuarios y configuraciones.

Técnico (TECNICO): Puede crear y editar activos, revisiones, obras e incidencias. No puede borrar registros ni acceder a la gestión de usuarios.

Consulta / Invitado: Solo puede visualizar datos y descargar documentos. No puede modificar nada, salvo utilizar el botón de "Reportar Avería".

3. El Cuadro de Mandos (Dashboard)
Al entrar, verá la pantalla principal con una visión global del estado de las instalaciones:

Tarjetas Superiores: Resumen numérico (Activos, Revisiones Vencidas, Incidencias abiertas, etc.). Al hacer clic en ellas, le llevará al apartado correspondiente.

Calendario Interactivo: Muestra las revisiones programadas.

Puede hacer clic en un día vacío para programar una nueva revisión.

Puede hacer clic en un evento existente para editarlo.

Gráficos: Evolución de la carga de trabajo y estado actual de cumplimiento (Quesito: Verde=Al día, Amarillo=Próximo, Rojo=Vencido).

4. Gestión de Inventario (Campus, Edificios y Activos)
La estructura de la información es jerárquica: Campus > Edificio > Activo.

4.1. Campus y Edificios
En el menú lateral, acceda a estas secciones para dar de alta las sedes.

Crear: Pulse el botón azul + Nuevo.

Documentación: Al entrar en la ficha de un edificio, puede adjuntar planos, licencias de apertura o legalizaciones en la pestaña "Documentación".

Filtros: En el listado de edificios, use la barra superior para filtrar por Campus o buscar por nombre.

4.2. Activos (Equipos)
Es el corazón del sistema. Aquí se registran calderas, cuadros eléctricos, ascensores, etc.

Alta de Activo: Seleccione primero el Campus y el Edificio. Luego pulse + Crear Activo.

Ficha del Activo: Al pulsar en "Ver Ficha", accederá al detalle donde podrá gestionar:

Información: Datos técnicos y marca.

Documentación: Manuales, fichas técnicas.

Mantenimiento: El plan de revisiones específico de ese equipo.

Contratos: Contratos de mantenimiento asociados.

5. Mantenimiento Preventivo
5.1. Programar Revisiones
Puede programar una revisión desde la ficha de un activo o desde el calendario del Dashboard.

Tipo: Legal (normativa), Periódica, Reparación, etc.

Frecuencia: Si marca "Repetir esta revisión", el sistema generará automáticamente las revisiones futuras (ej: cada 365 días).

Sincronización con Calendar: Marque la casilla ☑ Sincronizar con Google Calendar para que el aviso aparezca automáticamente en su agenda de Google.

5.2. Semáforo de Estado
El sistema le avisa visualmente de la urgencia:

🔴 Rojo: Revisión vencida (fecha pasada).

🟡 Amarillo: Faltan 30 días o menos.

🟢 Verde: Al día.

5.3. Subir Evidencias (OCAs)
Cuando realice una revisión, entre en ella (lápiz de editar) y use la zona de "Evidencias / Documentos" para subir el PDF del certificado o informe técnico. Este documento quedará archivado en Drive automáticamente.

6. Mantenimiento Correctivo (Incidencias)
Este módulo sirve para reportar averías imprevistas (ej: "Puerta atascada").

6.1. Reportar una Avería
Cualquier usuario puede hacerlo pulsando el botón rojo flotante (megáfono) en la esquina inferior derecha.

Seleccione dónde está el problema (Campus > Edificio > Activo).

Describa qué pasa.

Adjunte una foto: Puede subirla desde el móvil. Verá una previsualización antes de enviar.

Pulse "Enviar Reporte".

6.2. Gestión de Tickets (Para Técnicos)
En el menú "Incidencias", los técnicos verán la lista de problemas.

Estados: PENDIENTE ➝ EN PROCESO ➝ RESUELTA.

Acciones: Use los botones para cambiar el estado. Al marcarla como RESUELTA (botón verde), el ticket se cierra.

7. Obras y Reformas
Dentro de la ficha de cada Edificio, encontrará la pestaña "Obras y Reformas".

Utilícela para registrar intervenciones mayores (reformas de cubiertas, pintura, obras civiles).

Puede adjuntar actas de obra y presupuestos.

Al terminar la obra, pulse el botón "Finalizar" para registrar la fecha de fin.

8. Configuración y Usuarios (Solo Administradores)
8.1. Gestión de Usuarios
En el menú "Usuarios", el administrador puede dar de alta al personal.

Importante: El email debe ser la cuenta de Google (Gmail o corporativa) con la que el usuario accede.

Roles: Asigne ADMIN, TECNICO o CONSULTA.

Alertas: Marque "SI" si quiere que ese usuario reciba el resumen semanal por correo.

8.2. Catálogo de Instalaciones
En el menú "Configuración", defina los tipos de activos (ej: "Baja Tensión", "Climatización") y la frecuencia de revisión por defecto (días) según normativa. Esto agiliza el alta de nuevos activos.

9. Notificaciones Automáticas
El sistema envía automáticamente un correo electrónico todos los lunes a las 08:00 AM a los usuarios configurados. Este correo incluye un resumen de:

Revisiones vencidas o próximas a vencer.

Contratos que van a caducar.

10. Preguntas Frecuentes (FAQ)
P: ¿Dónde se guardan los archivos que subo? R: Todos los archivos se guardan en una estructura de carpetas organizada en Google Drive, dentro de la carpeta raíz configurada para la aplicación. Nunca perderá un documento aunque se borre de la app.

P: Soy Técnico y no veo el botón de borrar (papelera). R: Es correcto. Por seguridad, el perfil técnico no puede eliminar registros definitivos. Si necesita borrar algo por error, contacte con un Administrador.

P: He subido una foto a una incidencia y me he equivocado. R: Antes de guardar, puede pulsar la "X" sobre la foto para quitarla. Si ya la ha guardado, entre en "Editar" (lápiz) y vuelva a gestionar la incidencia.

P: ¿Cómo busco un activo concreto rápidamente? R: Use la barra de búsqueda global situada en la parte superior del menú lateral izquierdo. Escriba el nombre o marca y pulse la lupa.
