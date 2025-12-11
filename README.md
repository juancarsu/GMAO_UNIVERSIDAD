# Manual GMAO Universidad - Dashboard

## 📊 Tarjetas de Resumen

El dashboard muestra **6 tarjetas** con métricas clave:

| Tarjeta | Descripción | Acción al hacer clic |
|---------|-------------|---------------------|
| **Activos** | Total de instalaciones registradas | Va a la sección Activos |
| **Vencidas** | Revisiones con fecha superada | Va a Mantenimiento (filtro rojo) |
| **Pendientes** | Revisiones en los próximos 30 días | Va a Mantenimiento (filtro amarillo) |
| **Incidencias** | Averías pendientes o en proceso | Va a Incidencias |
| **Contratos** | Total de contratos activos | Va a Contratos |
| **Campus** | Número de sedes registradas | Va a Campus |

---

## 📅 Calendario de Revisiones

### Funcionalidades del calendario

- **Vista mensual** de todas las revisiones programadas
- **Código de colores**:
  - 🔴 **Rojo**: Revisión vencida
  - 🟡 **Amarillo**: Revisión próxima (≤30 días)
  - 🟢 **Verde**: Revisión al día (>30 días)

### Acciones disponibles

1. **Crear revisión desde el calendario**:
   - Haz clic en cualquier **día vacío** del calendario
   - Se abrirá el formulario de nueva revisión con la fecha preseleccionada
   - Selecciona Campus → Edificio → Activo
   - Completa los datos y guarda

2. **Ver/Editar revisión existente**:
   - Haz clic sobre un **evento** (revisión programada)
   - Se abrirá el formulario con los datos para editar

3. **Cambiar vista**:
   - **Mes**: Vista general por días
   - **Lista**: Listado cronológico de eventos

---

## 📈 Gráficos de Evolución

### Gráfico lineal (6 meses)

Muestra la **tendencia de revisiones** programadas para los próximos 6 meses:
- Eje X: Meses (Ene 2025, Feb 2025...)
- Eje Y: Número de revisiones por mes
- Útil para prever carga de trabajo

### Gráfico circular (Estado actual)

Distribución del estado de las revisiones:
- 🟡 **Amarillo**: Pendientes (≤30 días)
- 🔴 **Rojo**: Vencidas
- 🟢 **Verde**: Al día

---

## 🧭 Menú de Navegación Lateral

### Secciones principales

| Icono | Sección | Descripción |
|-------|---------|-------------|
| 📊 | Dashboard | Panel principal (esta página) |
| 🏛️ | Campus | Gestión de sedes universitarias |
| 🏢 | Edificios | Gestión de inmuebles por campus |
| 📦 | Activos | Instalaciones y equipos |
| 🔧 | Mantenimiento | Plan global de revisiones |
| ⚠️ | Incidencias | Reportes de averías |
| 📄 | Contratos | Gestión de proveedores |

### Secciones de administración

| Icono | Sección | Acceso |
|-------|---------|--------|
| ⚙️ | Configuración | Solo ADMIN |
| 👥 | Usuarios | Solo ADMIN |

---

## 🔍 Barra de Búsqueda Global

**Ubicación**: Parte superior del menú lateral

### Cómo usar la búsqueda

1. Escribe al menos **3 caracteres** en el cuadro de búsqueda
2. El sistema buscará automáticamente en:
   - Nombres de activos
   - Tipos de instalación
   - Marcas de equipos
   - Nombres de edificios

3. **Resultados**:
   - 📦 **Activos**: Con icono azul
   - 🏢 **Edificios**: Con icono verde

4. Haz clic en un resultado para ir directamente a su ficha

> **💡 Tip**: Presiona `ESC` para cerrar los resultados

---

## 👤 Perfil de Usuario

**Ubicación**: Parte inferior del menú lateral

Muestra:
- **Nombre** del usuario logueado
- **Rol** asignado (ADMIN / TECNICO / CONSULTA)

---

## 🚀 Acciones Rápidas desde el Dashboard

### Crear nuevo activo
Botón superior derecho **"+ Nuevo Activo"**

### Reportar avería urgente
Botón rojo **flotante inferior derecho** (icono megáfono)
- Disponible desde cualquier sección
- Accesible para todos los roles (incluido CONSULTA)
