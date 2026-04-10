# 📂 Automatización de socios (flujo actual)

Este proyecto se opera **desde la hoja `Inscripción`** y usa los scripts de la carpeta `Archivos` como flujo principal.

---

## 🚀 0. Arranque rápido

Si solo quieres correr el proceso normal:

1. Verifica requisitos de la sección 1.
2. Ejecuta, en este orden:
	 - `crearCarpetasSocios()`
	 - `renombrarYAgregarHojas()`
	 - `actualizarPermisosProteccion()`
3. Deja `procesarAvales()` con trigger diario.
4. Revisa `Ejecuciones` y `Logs` en Apps Script.

---

## ✅ 1. Requisitos reales

- La hoja activa debe tener pestaña **`Inscripción`**.
- El script lee socios desde:
	- **A:** número de socio
	- **B:** nombre
	- **C:** primer apellido
	- **D:** segundo apellido
	- **F:** valor que se copia a `F1` de `Tarjeta Ahorro`
- Debe existir esta estructura en Drive:
	- Carpeta raíz: **`GA0452 METAMORFOSIS`**
	- Carpeta de socios: **`GA0452-SOCIOS AS`**
- Debe existir archivo base en la carpeta del script con nombre que contenga:
	- **`04 TARJETA AHORRO Y PRESTAMO`**
- Primera ejecución de `IMPORTRANGE`: puede requerir autorización manual.

---

## 🧭 2. Flujo recomendado (carpeta Archivos)

> Los scripts están pensados para ejecución incremental y registran todo en logs.

### 2.1 `0.2-Carpetas.gs` → `crearCarpetasSocios()`

Crea carpetas de socios a partir de `Inscripción`.

- Detecta existentes por **número + nombre completo** (ignora variaciones en iniciales).
- Solo crea faltantes.
- Formato de carpeta:
	- `[Número] [INICIALES] [Nombre] [Primer Apellido] [Segundo Apellido]`
	- Iniciales = primera letra de nombre + primera letra apellido 1 + primera letra apellido 2 (mayúsculas)

### 2.2 `0.3-CreaciónHoja.gs` → `renombrarYAgregarHojas()`

Crea o actualiza tarjetas de socios desde el archivo base.

- Busca carpetas de socios (ignora variaciones en formato con/sin iniciales).
- Extrae las iniciales desde la carpeta del socio `[Número] [INI...] [...]`.
- Crea/reutiliza archivo con nombre:
	- `[INICIALES] - TARJETA AHORRO Y PRESTAMO`
- Agrega solo hojas faltantes (no duplica hojas ya existentes).
- Actualiza campos clave:
	- `Tarjeta Ahorro` y `Fondo de Emergencia`: `B1`, `D1`
	- `Tarjeta Ahorro`: `F1` desde columna `F` de `Inscripción`
	- `Tarjeta Prestamo #n`: `C47` (nombre), `G47` (código)

### 2.3 `0.4-PermisosHoja.gs` → `actualizarPermisosProteccion()`

Sincroniza protecciones del archivo base en tarjetas de socios.

- Compara por hoja/rango y actualiza solo diferencias.
- Copia editores, descripción y tipo de protección.
- Elimina protecciones que ya no existan en el base.
- Aplica en:
	- `Tarjeta Ahorro`
	- `Fondo de Emergencia`
	- `Tarjeta Prestamo #n`

### 2.4 `0.1-AñosAnteriores.gs` → `moverYOrganizarAnios()`

Organización histórica anual.

- Mueve listas antiguas a `01 LISTAS AÑOS ANTERIORES`.
- Ordena ahorros de años previos dentro de cada socio.

> No es parte de la operación diaria; usar en cierre/orden anual.

---

## 🤝 3. Avales (Préstamos)

Archivo: `Archivos/Préstamos/1-Avales.gs`  
Función: `procesarAvales()`

Incluye:

- Validación y normalización de hojas `Tarjeta Prestamo #n`.
- Control de préstamos activos para evitar inconsistencias.
- Sincronización de protecciones y formato con el base.
- Procesamiento bidireccional:
	- **Aval → Prestatario**
	- **Prestatario → Aval** (búsqueda inversa con normalización de nombres)

### ⏰ Trigger diario recomendado (paso a paso)

Para dejar `procesarAvales()` automático todos los días:

1. Abre el archivo de Apps Script donde está `1-Avales.gs`.
2. En el menú lateral, entra a **Activadores** (ícono de reloj).
3. Haz clic en **+ Agregar activador**.
4. Configura así:
	- **Función que se ejecutará:** `procesarAvales`
	- **Origen del evento:** `Basado en tiempo`
	- **Tipo de activador:** `Diario`
	- **Hora:** elige una ventana de baja actividad (por ejemplo mañana temprano).
5. Guarda el activador.
6. Si es la primera vez, autoriza permisos cuando Google lo solicite.

Verificación rápida:

- Revisa en **Activadores** que aparezca `procesarAvales` con estado activo.
- Ve a **Ejecuciones** para confirmar que corre sin error.
- Si falla por permisos o `IMPORTRANGE`, autoriza y vuelve a ejecutar manualmente una vez.

---

## 📊 4. Concentrados (carpeta `Concentrados/Actualizado`)

Los scripts de esta carpeta generan y actualizan los concentrados semanales en hojas resumen.

### 4.1 `1-InformeInicial.gs` → `registrarSociosCondensado()`

Inicializa y mantiene el padrón base en los concentrados.

- Trabaja sobre hojas:
	- `Ahorros y Retiros`
	- `Fondo de Emergencia`
- Inserta socios nuevos desde `Inscripción` sin duplicar existentes.
- Actualiza fórmula de nombre completo y totales por fila.
- Ordena alfabéticamente por nombre para mantener consistencia.

### 4.2 `2-InformeAhorroSemanal.gs` → `llenarCondensadoAhorros()`

Consolida movimientos semanales de ahorro/retiro en `Ahorros y Retiros`.

- Extrae fechas desde `Tarjeta Ahorro` rango **A12:A54**.
- Detecta semanas por fecha y escribe encabezados en fila 3.
- Agrega fórmulas semanales por socio con `IMPORTRANGE` + `QUERY` sobre rango **A12:D54**.
- Completa columnas faltantes sin duplicar semanas existentes.
- Reaplica formato condicional al crear columnas nuevas.

### 4.3 `3-InformeEmergenciaSemanal.gs` → `llenarCondensadoEmergencias()`

Consolida movimientos semanales del fondo de emergencia.

- Extrae fechas desde `Fondo de Emergencia` rango **A4:A25**.
- Detecta semanas por fecha y escribe encabezados en fila 3.
- Llena fórmula de suma semanal por socio con `IMPORTRANGE` + `QUERY` sobre rango **A4:D25**.
- Si la semana ya existe, rellena solo celdas faltantes.
- Escribe fórmula de suma en fila 2 para columnas nuevas.
- Actualiza automáticamente la hoja `Resumen Fondo de Emergencia` con fechas y totales.

### 4.4 `4-InformePrestamoSemanal.gs` → `llenarCondensadoPrestamos()`

Consolida abonos semanales de préstamos en la hoja `Préstamos`.

- Detecta hojas `Tarjeta Prestamo #n` por socio.
- Extrae pagos desde rango **B13:D23** de cada tarjeta de préstamo.
- Carga columnas base (E:K) con `IMPORTRANGE` para nuevos préstamos.
- Crea semanas en fila 3 y abonos semanales mediante `QUERY` dinámico.
- Escribe suma en fila 2 solo para semanas nuevas (sin sobreescribir existentes).
- Actualiza solo columnas nuevas en préstamos ya registrados.

### ⏰ Ejecución recomendada

- Esta parte está configurada para correr **semanalmente**.
- Mantén la misma estructura operativa del proyecto (insumos en `Inscripción`, tarjetas por socio y nombres de hojas estándar).
- Orden sugerido de ejecución manual:
	1. `registrarSociosCondensado()`
	2. `llenarCondensadoAhorros()`
	3. `llenarCondensadoEmergencias()`
	4. `llenarCondensadoPrestamos()`
- Si no usas activador semanal, ejecuta ese orden una vez por semana.

---

## 🗓️ 5. Operación diaria mínima

1. Mantén actualizada la hoja `Inscripción`.
2. Si hay altas/cambios masivos, ejecuta:
	 - `crearCarpetasSocios()`
	 - `renombrarYAgregarHojas()`
	 - `actualizarPermisosProteccion()`
3. Mantén `procesarAvales()` en trigger diario.
4. Revisa errores y resumen en `Ejecuciones` + `Logs`.

---

## 🛡️ 6. Garantías operativas

### ✅ Replicable

- El flujo se puede volver a ejecutar desde `Inscripción` sin rehacer todo manualmente.
- Si agregas socios nuevos, se integran en la siguiente ejecución.

### ✅ Sin duplicados (condiciones normales)

- `crearCarpetasSocios()` evita duplicar carpetas existentes.
- `renombrarYAgregarHojas()` agrega solo hojas faltantes.
- `procesarAvales()` valida y evita repetir avales ya registrados.

### ✅ Sin sobreescritura masiva

- `0.3-CreaciónHoja.gs` actualiza solo campos clave puntuales.
- `0.4-PermisosHoja.gs` cambia únicamente protecciones/rangos diferentes.

### ⚠️ Límites y precauciones

- Los scripts buscan carpetas por **número + nombre completo**, ignorando variaciones en iniciales.
- Si cambias manualmente nombres de hojas fuera del formato `Tarjeta Prestamo #n`, el script puede renombrarlas para normalizar.
- Antes de correr masivo, prueba con 1–2 socios y revisa logs.
- Si falta carpeta base o archivo base, el proceso se detiene y lo reporta en `Logger`.