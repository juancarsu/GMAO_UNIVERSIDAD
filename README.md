# 📘 Manual de Usuario GMAO Universidad de Navarra

## Sistema de Gestión de Mantenimiento, Activos y Obras

**Versión 1.0** | Autor: Juan Carlos Suárez  
**Licencia**: Creative Commons Reconocimiento (CC BY)

---

## 🎯 Índice Rápido

- [1. Introducción](#1-introducción)
- [2. Acceso y Roles](#2-acceso-y-roles)
- [3. Dashboard](#3-dashboard)
- [4. Campus](#4-campus)
- [5. Edificios](#5-edificios)
- [6. Activos](#6-activos)
- [7. Mantenimiento](#7-mantenimiento)
- [8. Incidencias](#8-incidencias)
- [9. Contratos](#9-contratos)
- [10. Configuración](#10-configuración)
- [11. Usuarios](#11-usuarios)
- [12. FAQ](#12-faq)

---

# 1. Introducción

## ¿Qué es GMAO Universidad?

Sistema integral de **Gestión de Mantenimiento Asistido por Ordenador** diseñado para:

✅ Gestionar activos en múltiples campus  
✅ Planificar mantenimiento preventivo  
✅ Documentar obras y reformas  
✅ Registrar incidencias en tiempo real  
✅ Administrar contratos con proveedores  
✅ Cumplir normativas de mantenimiento  

## Arquitectura del Sistema

```
CAMPUS (Sedes)
  └─ EDIFICIOS (Inmuebles)
      └─ ACTIVOS (Instalaciones/Equipos)
          ├─ Mantenimiento (Revisiones)
          ├─ Documentación (Archivos)
          └─ Contratos (Proveedores)
```

---

# 2. Acceso y Roles

## Inicio de Sesión

1. Acceder a la URL proporcionada por tu administrador
2. Autorizar con tu cuenta Google corporativa
3. El sistema identifica tu rol automáticamente

## Roles y Permisos

| Rol | Crear | Editar | Eliminar | Config | Usuarios |
|-----|-------|--------|----------|--------|----------|
| **🔴 ADMIN** | ✅ | ✅ | ✅ | ✅ | ✅ |
| **🔵 TECNICO** | ✅ | ✅ | ❌ | ❌ | ❌ |
| **⚪ CONSULTA** | ❌ | ❌ | ❌ | ❌ | ❌ |

> **Nota**: CONSULTA puede reportar incidencias

---

# 3. Dashboard

## Tarjetas de Resumen

| Tarjeta | Descripción | Clic = |
|---------|-------------|--------|
| **Activos** | Total de instalaciones | → Activos |
| **Vencidas** | Revisiones atrasadas | → Mantenimiento (rojo) |
| **Pendientes** | Revisiones ≤30 días | → Mantenimiento (amarillo) |
| **Incidencias** | Averías activas | → Incidencias |
| **Contratos** | Total activos | → Contratos |
| **Campus** | Sedes registradas | → Campus |

## Calendario de Revisiones

**Crear revisión**:
- Clic en día vacío → Formulario con fecha preseleccionada

**Editar revisión**:
- Clic en evento existente → Formulario con datos

**Colores**:
- 🔴 Rojo = Vencida
- 🟡 Amarillo = Próxima (≤30 días)
- 🟢 Verde = Al día

## Gráficos

**Lineal (6 meses)**: Evolución de revisiones programadas  
**Circular**: Distribución por estado (Vencidas/Pendientes/Al día)

## Búsqueda Global

**Ubicación**: Barra superior del menú lateral

**Uso**:
1. Escribe ≥3 caracteres
2. Busca en: activos, edificios, tipos, marcas
3. Clic en resultado → Ficha detallada
4. `ESC` para cerrar

---

# 4. Campus

## Gestión de Sedes

### Ver Campus

**Tabla muestra**:
- Nombre
- Provincia
- Dirección

### Crear Campus

**Requisito**: TECNICO o ADMIN

1. Clic **"+ Nuevo Campus"**
2. Rellenar:
   - **Nombre** (obligatorio)
   - **Provincia**
   - **Dirección**
3. Guardar

> Se crea carpeta automática en Google Drive

### Editar Campus

1. Clic en ✏️
2. Modificar datos
3. Guardar

### Eliminar Campus

**Requisito**: Solo ADMIN

1. Clic en 🗑️
2. Confirmar

> ⚠️ Verificar que no tenga edificios asociados

---

# 5. Edificios

## Gestión de Inmuebles

### Filtros

**Por Campus**: Desplegable  
**Por texto**: Buscar nombre/contacto

### Ver Edificio

**Tabla muestra**:
- Edificio (nombre)
- Campus
- Contacto/Responsable

### Crear Edificio

**Requisito**: TECNICO o ADMIN

1. Clic **"+ Nuevo Edificio"**
2. Completar:
   - **Nombre** (obligatorio)
   - **Campus** (obligatorio)
   - **Contacto**
3. Guardar

> Se crean 2 carpetas: Edificio + subcarpeta "Activos"

## Ficha Detallada

**4 Pestañas**:

### 📋 Información

- Nombre
- Campus
- Contacto

### 📄 Documentación Legal/Planos

**Subir**:
1. Elegir archivo
2. Clic **"📤 Subir Documento"**

**Tipos sugeridos**:
- Planos arquitectónicos
- Licencias
- Certificados (gas, electricidad, PCI)
- Seguros

### 🏗️ Obras y Reformas

**Crear obra**:
1. Clic **"+ Nueva Obra"**
2. Completar:
   - **Nombre** (ej: "Reforma Cubierta Norte")
   - **Descripción**
   - **Fecha Inicio**
   - *Opcional*: Adjuntar documento
3. Guardar

**Estados**:
- 🟡 EN CURSO
- 🟢 FINALIZADA

**Acciones**:
- **Finalizar**: Botón ✓ → introducir fecha fin
- **Adjuntar evidencias**: Botón "+" en tarjeta
- **Ver documentos**: Clic "Ver documentos"
- **Eliminar** (solo ADMIN): 🗑️

### 📦 Activos Instalados

**Filtro**: Por tipo de instalación  
**Acción**: Clic "Ir a Activo" → Ficha completa

---

# 6. Activos

## Gestión de Instalaciones

### Sistema de Filtrado

**Cascada en 3 pasos**:
1. Seleccionar **Campus**
2. Seleccionar **Edificio** (se activa)
3. **Filtro texto** (opcional, se activa)

### Crear Activo

**Requisito**: TECNICO o ADMIN

1. Clic **"+ Crear Activo"**
2. **Paso 1**: Campus + Edificio
3. **Paso 2**: Tipo de Instalación (del catálogo)
4. **Paso 3**: Nombre + Marca
5. Guardar

> Se crea carpeta en "Activos" del edificio

## Ficha de Activo

**4 Pestañas**:

### 📋 Información

**Modo Vista**:
- Nombre, Tipo, Marca, Fecha alta

**Modo Edición**:
1. Clic **"✏️ Editar"**
2. Modificar: Nombre, Tipo, Marca
3. Guardar o Cancelar

> 🔒 Fecha alta no modificable

### 📄 Documentación

**Subir**:
1. Elegir archivo
2. Clic **"📤 Subir"**

**Control de versiones**: Automático por nombre

**Acciones**:
- Ver: 👁️
- Eliminar (ADMIN): 🗑️

### 🔧 Mantenimiento

**Tabla de revisiones**:
- **Estado**: 🔴/🟡/🟢 (semáforo)
- **Tipo**: Legal / Periódica / Reparación / Extraordinaria
- **Próxima**: Fecha límite
- **Iconos**: 📎 (docs) 📅 (calendar)

**Crear revisión**:
1. Clic **"+ Programar Revisión"**
2. Ver sección [7. Mantenimiento](#7-mantenimiento)

**Editar**: ✏️  
**Eliminar** (ADMIN): 🗑️

### 📄 Contratos

**Estados**:
- 🟢 VIGENTE
- 🟡 PRÓXIMO (≤90 días)
- 🔴 CADUCADO
- ⚪ INACTIVO / SIN FECHA

**Crear**: **"+ Añadir Contrato"**  
**Ver sección**: [9. Contratos](#9-contratos)

---

# 7. Mantenimiento

## Vista Global

**Acceso**: Menú lateral → Mantenimiento

### Filtros Avanzados

**Por ubicación**:
- Campus (desplegable)
- Edificio (se carga tras campus)

**Por estado** (botones):
- Todas (predeterminado)
- 🔴 Vencidas
- 🟡 Próximas
- 🟢 Al día

**Por tipo** (botones):
- Todos
- Legal
- Periódica
- Reparación
- Extraordinaria

### Tabla Global

| # | Edificio | Activo | Tipo | Fecha Límite | Acciones |
|---|----------|--------|------|--------------|----------|
| 🔴 | ... | ... | ... | ... | ✏️ 🗑️ |

> Ordenado por urgencia (vencidas primero)

## Formulario de Revisión

### 1. Selector de Activo

**Visible solo desde**:
- Dashboard (calendario)
- Vista global

**Cascada**: Campus → Edificio → Activo

> Oculto si accedes desde ficha de activo

### 2. Tipo de Revisión

**Opciones**:
- **Legal**: Obligatoria por normativa
- **Periódica**: Preventiva programada
- **Reparación**: Correctiva
- **Extraordinaria**: Puntual

**Si Legal**:
- Aparece: Desplegable "Normativa"
- Autorellena frecuencia

**Si Legal o Periódica**:
- Aparece: Checkbox "Repetir"

### 3. Fecha Próxima

Calendario `YYYY-MM-DD`

### 4. Sincronizar Google Calendar

✅ **Activado por defecto**

**Crea evento**:
- Título: `MANT: [Tipo] - [Activo]`
- Día completo
- Color rojo
- Descripción con datos del activo

**Actualización**:
- Editar revisión → actualiza evento
- Eliminar revisión → borra evento

### 5. Repetir Revisión

**Solo**: Legal y Periódica

**Campos**:
- **Cada (días)**: Frecuencia (ej: 365)
- **Repetir hasta**: Fecha límite

**Funcionamiento**:
- Crea revisiones automáticas espaciadas
- Máximo 50 repeticiones
- Ejemplo: Cada 365 días hasta 2028 → 4 revisiones

### 6. Evidencias / Documentos

**Importante**: Solo tras guardar revisión por primera vez

**Subir evidencias**:
1. Guardar revisión (primera vez)
2. Editar revisión (✏️)
3. Sección "EVIDENCIAS"
4. Elegir archivo → Subir

**Tipos útiles**:
- Certificados de revisión
- Fotos de intervención
- Informes técnicos

**Eliminar**: Botón ✕

---

# 8. Incidencias

## Gestión de Averías

### Botón Flotante

**Ubicación**: 🔴 Esquina inferior derecha (megáfono)

**Acceso**: Todos los roles (incluido CONSULTA)

**Función**: Reportar avería desde cualquier sección

### Tabla de Incidencias

| Estado | Prioridad | Elemento | Descripción | Reportado | Acción |
|--------|-----------|----------|-------------|-----------|--------|
| Badge | BAJA/MEDIA/ALTA/URGENTE | ... | ... | Fecha + Usuario | Botones |

### Filtros

- **Todas**
- **Pendientes**
- **Resueltas**

## Reportar Incidencia

1. Clic botón 🔴 flotante

### Formulario

**1. ¿Dónde?**
- Campus (obligatorio)
- Edificio (obligatorio)
- Activo (opcional)

**2. Descripción**
- Texto libre explicando problema

**3. Prioridad**
- BAJA / MEDIA / ALTA / ¡URGENTE!

**4. Foto** (opcional)
- Elegir archivo
- Vista previa
- Botón ✕ para quitar

5. **"Enviar Reporte"**

> 📸 Recomendación: Adjuntar fotos siempre que sea posible

## Gestionar Incidencias

**Requisito**: TECNICO o ADMIN

### Cambiar Estado

**Desde PENDIENTE**:
- ▶️ **En Proceso**
- ✅ **Resolver**

**Desde EN PROCESO**:
- ✅ **Resolver**

> No se puede reabrir una incidencia RESUELTA

### Editar Datos

**Botón ✏️** (solo si no resuelta):
- Modificar: Campus, Edificio, Activo, Descripción, Prioridad
- **No se puede**: Cambiar foto, fecha, usuario

---

# 9. Contratos

## Gestión de Proveedores

### Vista Global

**Acceso**: Menú lateral → Contratos

### Filtros

**Por ubicación**:
- Campus
- Edificio

**Por estado**:
- Todos
- 🟢 Vigente
- 🟡 Próximo (≤90 días)
- 🔴 Caducado
- ⚪ Inactivo

### Tabla

| Estado | Activo/Edificio | Proveedor | Ref | Vigencia | Acciones |
|--------|-----------------|-----------|-----|----------|----------|
| Badge | ... | ... | ... | Inicio - Fin | ✏️ 🗑️ |

## Crear Contrato

**Desde**:
- Ficha de activo (pestaña Contratos)
- Vista global (botón "+ Nuevo Contrato")

### Formulario

| Campo | Obligatorio | Descripción |
|-------|-------------|-------------|
| **Proveedor** | ✅ | Nombre empresa |
| **Referencia** | ❌ | Nº contrato/pedido |
| **Estado** | ✅ | ACTIVO / INACTIVO |
| **Inicio** | ✅ | Fecha inicio |
| **Fin** | ✅ | Fecha finalización |

### Cálculo Automático

**Si ACTIVO**:
- Días hasta fin **<0** → 🔴 CADUCADO
- Días hasta fin **≤90** → 🟡 PRÓXIMO
- Días hasta fin **>90** → 🟢 VIGENTE
- Sin fecha fin → ⚪ SIN FECHA

**Si INACTIVO** → ⚪ Badge gris

## Editar/Eliminar

- **Editar**: ✏️ (todos los roles con permisos)
- **Eliminar**: 🗑️ (solo ADMIN)

---

# 10. Configuración

## Catálogo de Instalaciones

**Requisito**: Solo ADMIN

### Función

- Estandarizar tipos de activos
- Asociar normativas
- Definir frecuencias predeterminadas

### Tabla

| Nombre | Ref. Normativa | Frecuencia (días) | Acciones |
|--------|----------------|-------------------|----------|
| Baja Tensión | REBT 2002 | 365 | ✏️ 🗑️ |

### Crear Tipo

1. Clic **"+ Nuevo Tipo"**
2. Completar:
   - **Nombre** (obligatorio, ej: "Climatización")
   - **Ref. Normativa** (opcional, ej: "RITE 2021")
   - **Frecuencia** (días, opcional)
3. Guardar

### Uso Posterior

- Aparece en desplegable al crear activo
- Al programar revisión Legal, autorellena frecuencia

### Editar/Eliminar

- **Editar**: ✏️
- **Eliminar**: 🗑️

> ⚠️ No eliminar tipos en uso

---

# 11. Usuarios

## Gestión de Accesos

**Requisito**: Solo ADMIN

### Tabla

| Nombre | Email | Rol | Alertas | Acciones |
|--------|-------|-----|---------|----------|
| ... | ... | Badge | 🔔/🔕 | ✏️ 🗑️ |

### Crear Usuario

1. Clic **"+ Nuevo Usuario"**
2. Completar:
   - **Nombre** (obligatorio)
   - **Email** (obligatorio, Google corporativo)
   - **Rol** (obligatorio):
     - ADMIN: Permisos completos
     - TECNICO: Crear/editar
     - CONSULTA: Solo lectura
   - **Alertas**:
     - Sí: Recibe emails semanales
     - No: Sin notificaciones
3. Guardar

### Activación

Usuario debe acceder con su cuenta Google → reconocido por email → aplica rol

### Editar/Eliminar

- **Editar**: ✏️ (cambiar rol/alertas)
- **Eliminar**: 🗑️ (pierde acceso inmediato)

---

# 12. FAQ

## General

**¿Acceso desde móvil?**  
✅ Sí, interfaz responsive

**¿Autoguardado?**  
❌ No, siempre hacer clic en "Guardar"

**¿Deshacer acción?**  
❌ No, eliminaciones permanentes (con confirmación)

## Campus/Edificios

**¿Mover edificio a otro campus?**  
✅ Editar edificio → cambiar campus

**¿Eliminar campus con edificios?**  
⚠️ Sin verificación, eliminar edificios primero

## Activos

**¿Cambiar activo de edificio?**  
❌ Crear nuevo + copiar datos + eliminar antiguo

**¿Cambiar fecha de alta?**  
❌ Inmutable

## Mantenimiento

**¿Revisión sin activo?**  
❌ Siempre debe asociarse a activo

**¿Revisiones repetitivas se crean todas?**  
✅ Sí, al guardar (máx 50)

**¿Editar una actualiza todas?**  
❌ Cada revisión es independiente

**¿Sincronizar revisiones antiguas con Calendar?**  
❌ Solo al crear/editar

## Incidencias

**¿Reabrir incidencia resuelta?**  
❌ Crear nueva si reaparece problema

**¿Cambiar foto?**  
❌ Solo al crear

**¿Asignación automática de tareas?**  
❌ Gestión manual

## Contratos

**¿Contrato sin fecha fin?**  
✅ Posible, aparece como "SIN FECHA"

**¿Alertas automáticas de vencimiento?**  
✅ Si usuario tiene "Alertas: Sí" → email semanal

## Notificaciones

**¿Cómo activar emails automáticos?**  
⚙️ Configurar trigger en Apps Script: función `enviarResumenSemanal()` → frecuencia semanal

**¿Qué contiene el email?**  
📧 Revisiones vencidas/próximas + Contratos por vencer

## Almacenamiento

**¿Dónde se guardan los archivos?**  
🗂️ Google Drive, estructura automática: Campus → Edificios → Activos

**¿Límite de archivos?**  
❌ Sin límite del sistema, sujeto a cuota de Drive

**¿Permisos de archivos?**  
🔗 Cualquiera con enlace (Vista)

## Revisiones Completadas

**¿Qué pasa al marcar una revisión como "Completada"?**  
🟢 Desaparece de la lista de pendientes y del dashboard

**¿Se puede deshacer?**  
❌ No, pero puedes crear una nueva revisión para la próxima fecha

**¿Dónde ver el histórico de revisiones completadas?**  
📊 Actualmente no hay vista de histórico (próxima versión)

## Importación Masiva

**¿Puedo importar muchos activos a la vez?**  
✅ Sí, desde Excel/CSV usando copiar-pegar

**¿Cómo funciona?**  
1. Preparar Excel con columnas: Campus | Edificio | Tipo | Nombre | Marca
2. Copiar datos (sin cabeceras)
3. Pegar en modal de importación
4. Procesar

**¿Qué errores pueden ocurrir?**  
❌ Campus o Edificio no existente  
❌ Columnas mal ordenadas  
❌ Filas incompletas

---

# PARTE 13: FUNCIONES AVANZADAS ADICIONALES

---

## 🔔 Sistema de Feedback

### Reportar Bugs o Sugerencias

**Acceso**: Botón flotante azul (💬) inferior derecho

**Función**: Enviar feedback sobre la aplicación

### Tipos de Feedback

**💡 Sugerencia / Idea**
- Propuestas de mejora
- Nuevas funcionalidades
- Cambios en la interfaz

**🪲 Reporte de Fallo (Bug)**
- Errores encontrados
- Comportamientos inesperados
- Problemas de rendimiento

**💬 Otro comentario**
- Comentarios generales
- Dudas sobre uso
- Agradecimientos

### Cómo Enviar Feedback

1. Clic en botón flotante azul (💬)
2. Seleccionar tipo de mensaje
3. Escribir descripción detallada
4. Clic en **"Enviar"**

> **📝 Nota**: El feedback se guarda en una hoja "FEEDBACK" de la base de datos para revisión del administrador.

### Buenas Prácticas

**Para reportar bugs**:
- Describir qué intentabas hacer
- Indicar qué pasó en su lugar
- Mencionar navegador y dispositivo
- Adjuntar captura si es posible (por email)

**Para sugerencias**:
- Explicar el problema que resolvería
- Describir el comportamiento esperado
- Priorizar según necesidad

---

## ✅ Completar Revisiones

### Marcar Revisión como Realizada

**Ubicación**: Vista Global de Mantenimiento

**Función**: Indicar que una revisión se ha completado

### Cómo Completar una Revisión

1. Ir a **Mantenimiento** (vista global)
2. Localizar la revisión en la tabla
3. Clic en botón **✓ verde** (Completar)
4. Confirmar acción

### Efectos de Completar

**Cambios inmediatos**:
- ✅ Estado cambia a "REALIZADA"
- 📊 Desaparece del dashboard (contadores)
- 📅 Se oculta del calendario
- 🔍 No aparece en filtros de pendientes

**Permanece en**:
- 📁 Base de datos (histórico)
- 📎 Documentación asociada

### Diferencia: Completar vs Eliminar

| Acción | Completar ✓ | Eliminar 🗑️ |
|--------|------------|-------------|
| **Registra ejecución** | ✅ Sí | ❌ No |
| **Mantiene histórico** | ✅ Sí | ❌ No |
| **Elimina evento Calendar** | ❌ No | ✅ Sí |
| **Recuperable** | ⚠️ Manual | ❌ No |
| **Rol mínimo** | TECNICO | ADMIN |

### Caso de Uso

**Situación**: Revisión anual de caldera completada el 15/03/2025

**Pasos**:
1. Técnico realiza la revisión físicamente
2. Sube certificado/evidencias al plan
3. Marca como **"Completada"**
4. Sistema la oculta de pendientes
5. Crea nueva revisión para 15/03/2026

> **💡 Tip**: Siempre sube evidencias ANTES de completar la revisión para mantener trazabilidad.

---

## 📊 Importación Masiva de Activos

### Para Qué Sirve

Permite **migrar rápidamente** activos desde hojas de cálculo existentes (Excel, Google Sheets, CSV).

**Casos de uso**:
- Migración desde sistema antiguo
- Alta inicial de muchos activos
- Actualización masiva tras auditoría

### Requisitos Previos

**Antes de importar**:
1. ✅ Todos los **Campus** deben existir
2. ✅ Todos los **Edificios** deben existir
3. ✅ Los nombres deben coincidir **exactamente**

### Formato de Datos

**Orden de columnas** (obligatorio):

| Columna 1 | Columna 2 | Columna 3 | Columna 4 | Columna 5 |
|-----------|-----------|-----------|-----------|-----------|
| **Campus** | **Edificio** | **Tipo** | **Nombre Activo** | **Marca** |

**Ejemplo**:
```
Campus Central    Edificio A    Baja Tensión    Cuadro Principal    Schneider
Campus Central    Edificio A    Climatización    Caldera 1    Vaillant
Campus Tecnológico    Lab 3    Ascensor    Ascensor Principal    Otis
```

### Paso a Paso

#### 1. Preparar Excel

**En tu hoja de cálculo**:
- Organiza datos en 5 columnas (orden correcto)
- **NO incluyas fila de cabeceras**
- Verifica nombres de Campus/Edificios

#### 2. Copiar Datos

1. Selecciona **solo las celdas con datos** (sin cabecera)
2. Presiona **Ctrl+C** (Cmd+C en Mac)

#### 3. Abrir Modal de Importación

1. Ir a **Activos** (menú lateral)
2. Clic en botón **"📊 Importar"** (superior derecho)
3. Se abre modal "Importación Masiva"

#### 4. Pegar Datos

1. Clic en el área de texto grande
2. Presiona **Ctrl+V** (Cmd+V en Mac)
3. Los datos aparecen con tabulaciones

#### 5. Procesar Importación

1. Revisar datos pegados
2. Clic en **"Procesar Importación"**
3. Confirmar cantidad de activos
4. Esperar procesamiento (puede tardar)

### Resultado

**Si todo va bien**:
- ✅ Mensaje: "¡Éxito! Se han creado X activos"
- 🗂️ Cada activo tiene su carpeta en Drive
- 📊 Dashboard actualizado automáticamente

**Si hay errores**:
- ⚠️ Mensaje: "Importación parcial"
- 📝 Lista de errores abajo del área de texto
- ✅ Los activos válidos SÍ se crearon
- ❌ Los erróneos NO se crearon

### Errores Comunes

| Error | Causa | Solución |
|-------|-------|----------|
| "Campus 'X' no existe" | Nombre no coincide | Crear campus primero o corregir nombre |
| "Edificio 'Y' no encontrado" | No existe en ese campus | Verificar edificio y campus |
| "Fila incompleta" | Faltan columnas | Completar todas las 5 columnas |
| "No se detectan datos válidos" | Formato incorrecto | Usar tabulaciones (copiar de Excel) |

### Limitaciones

**Restricciones técnicas**:
- ⚠️ Puede tardar con **muchos activos** (>100)
- ⚠️ Drive tiene límites de carpetas/minuto
- ❌ No actualiza activos existentes, solo crea nuevos
- ❌ No permite importar mantenimientos o contratos

### Recomendaciones

**Mejores prácticas**:
1. 🧪 **Prueba primero** con 5-10 activos
2. 📝 **Documenta** nombres exactos de Campus/Edificios
3. 🔄 **Procesa por lotes** si son muchos (50-100 cada vez)
4. ✅ **Verifica** en la tabla tras cada importación
5. 🗂️ **Revisa Drive** que las carpetas se crearon

### Ejemplo Completo

**Escenario**: Importar 3 calderas

**Excel original**:
```
Campus       | Edificio  | Tipo          | Nombre        | Marca
-------------|-----------|---------------|---------------|----------
Campus Norte | Edif. A   | Climatización | Caldera Norte | Vaillant
Campus Norte | Edif. B   | Climatización | Caldera Sur   | Junkers
Campus Sur   | Edif. C   | Climatización | Caldera Este  | Baxi
```

**Pasos**:
1. Seleccionar solo datos (sin fila Campus|Edificio|...)
2. Copiar (Ctrl+C)
3. GMAO → Activos → Importar
4. Pegar (Ctrl+V) en área de texto
5. Procesar Importación
6. Confirmar

**Resultado**:
```
✅ ¡Éxito! Se han creado 3 activos.
```

---

## 🎨 Sistema de Notificaciones Mejorado

El sistema ahora incluye **notificaciones visuales elegantes** con SweetAlert2.

### Tipos de Notificaciones

**🟢 Éxito (Toast verde)**
- Aparece en esquina superior derecha
- Desaparece automáticamente en 3 segundos
- Ejemplos: "Activo guardado", "Revisión completada"

**🔴 Error (Alerta modal)**
- Ventana central bloqueante
- Requiere clic en "OK" para cerrar
- Ejemplos: "Campus no encontrado", "Campos incompletos"

**⚠️ Confirmación (Alerta modal)**
- Antes de acciones destructivas
- Botones: "Sí, proceder" / "Cancelar"
- Ejemplos: "¿Eliminar edificio?", "¿Marcar como completada?"

### Diferencias con Sistema Anterior

| Antes | Ahora |
|-------|-------|
| `alert()` nativo | Modal elegante SweetAlert2 |
| `confirm()` simple | Modal con estilos personalizados |
| Sin notificaciones de éxito | Toasts verdes informativos |
| Interrumpe flujo de trabajo | Notificaciones no invasivas |

---

# 📞 Soporte y Feedback

## Reportar Problemas

### Sistema Integrado de Feedback

**Acceso rápido**: Botón flotante azul (💬) siempre visible

**Proceso recomendado**:
1. Clic en botón flotante azul
2. Seleccionar tipo:
   - 🪲 **Bug**: Para errores técnicos
   - 💡 **Sugerencia**: Para mejoras
   - 💬 **Otro**: Para comentarios generales
3. Describir claramente el problema/idea
4. Enviar

**Ventajas**:
- ✅ No necesitas email del administrador
- ✅ Queda registrado en el sistema
- ✅ Accesible para todos los roles (incluido CONSULTA)
- ✅ Contexto automático (usuario, fecha)

### Contacto Directo

Para incidencias críticas o urgentes:

**Administrador GMAO**: [Insertar contacto]  
**Soporte IT**: [Insertar contacto]

### Información Útil al Reportar

**Para cualquier tipo de reporte**:
- 🖥️ Navegador y versión (Chrome 120, Firefox 115...)
- 📱 Dispositivo (PC, móvil, tablet)
- 👤 Rol de usuario (ADMIN, TECNICO, CONSULTA)
- 📝 Descripción paso a paso
- 📸 Capturas de pantalla (si aplica)

**Para bugs específicos**:
- ❌ Mensaje de error exacto
- 🔄 Pasos para reproducir
- ✅ Comportamiento esperado vs real

---

**Fin del Manual Actualizado**

*Versión 1.1 | Última actualización: Diciembre 2024*  
*Nuevas funcionalidades: Completar Revisiones, Importación Masiva, Sistema de Feedback*
