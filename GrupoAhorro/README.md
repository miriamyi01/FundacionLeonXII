# 📂 Instrucciones para automatización de carpetas y archivos de socios

Este proyecto automatiza la creación y llenado de carpetas y archivos para socios usando Google Apps Script y Google Drive.

---


## 1. Requisitos previos

- **Permisos de acceso:**  
  Antes de ejecutar cualquier script, asegúrate de abrir manualmente cada archivo base (`04 TARJETA AHORRO Y PRESTAMO`) y el archivo de concentrado (la hoja de cálculo principal) para otorgar permisos de acceso a `IMPORTRANGE` y edición si es necesario.  
  > ⚠️ **Nota:** La primera vez que uses `IMPORTRANGE` en el concentrado, deberás autorizar el acceso a cada archivo de socio. Hazlo manualmente para cada fórmula nueva, después ya no será necesario.

- **Archivos base:**  
  El archivo base debe llamarse exactamente:  
  **`04 TARJETA AHORRO Y PRESTAMO`**  
  y debe estar en la carpeta principal de Drive.  
  > 🔒 **Importante:** Antes de ejecutar el script, asegúrate de que las celdas protegidas (con fórmulas) en este archivo base estén desbloqueadas o que tengas permisos para editarlas, ya que las copias heredarán estos permisos.

- **Estructura de la hoja de cálculo:**  
  - La hoja de cálculo principal debe estar dentro de una carpeta en Google Drive.
  - 🚨 **IMPORTANTE:** Las carpetas de los socios deben estar en la misma carpeta principal donde está la hoja de cálculo (sin subcarpetas intermedias).
  - En la hoja de cálculo, la celda A8 debe tener el código base (por ejemplo, `2024`).
  - Los datos de los socios deben estar en las columnas:
    - A: Número de socio (desde la fila 8)
    - B: Nombre
    - C: Primer apellido
    - D: Segundo apellido
  - **Meses y años:**  
    Para que los reportes funcionen correctamente, los encabezados de los meses en las hojas deben tener el formato `Mes Año` (por ejemplo, `Marzo 2025`, `Abril 2026`).  
    > ⚠️ **No pongas solo el nombre del mes** (ejemplo: `Marzo`). Siempre debe incluir el año, ya que los scripts buscan y procesan por mes y año.

- **Precisión de datos:**  
  Todas las hojas y reportes usan **dos decimales** para mayor precisión en los cálculos y presentación de montos.

---


## 2. Flujo de automatización


> 📝 **NOTA IMPORTANTE:** Los scripts `1-Carpetas.gs`, `2-InformeInicial.gs`, `3-InformeAhorroSemanal.gs`, `4-InformePrestamoSemanal.gs` y `4.1-Avales.gs` están optimizados para **ejecución múltiple sin duplicados**. Puedes ejecutarlos tantas veces como necesites de forma segura - detectan automáticamente elementos existentes y solo procesan información nueva. Todos proporcionan reportes detallados en los logs para monitorear el proceso.


### 📁 Paso 1: Crear carpetas y archivos de socios (`1-Carpetas.gs`)

1. **Coloca la hoja de cálculo y el archivo base en la misma carpeta de Google Drive.**
2. **Abre el editor de Apps Script y pega el código de `1-Carpetas.gs`.**
3. **Ejecuta la función `crearCarpetasSocios`.**
   - El script creará carpetas de socios directamente en la carpeta principal (misma carpeta donde está la hoja de cálculo).
   - Para cada socio, creará una carpeta con el formato:  
     `[Número de socio] [INICIALES] [Nombre completo]`
   - Dentro de cada carpeta de socio, copiará el archivo base y lo renombrará con las iniciales.
   - En cada copia, pondrá el nombre completo (capitalizado) en B1 y el número de socio en D1.


### 📝 Paso 2: Registrar socios en el concentrado (`2-InformeInicial.gs`)

1. **Abre el editor de Apps Script y pega el código de `2-InformeInicial.gs`.**
2. **Ejecuta la función `registrarSociosCondensado`.**
   - El script llenará la hoja `Ahorros y Retiros` con los datos de los socios desde el archivo `03 LISTA DE INSCRIPCION`.
   - Ajusta automáticamente el número de filas y coloca fórmulas `IMPORTRANGE` para importar los datos de inscripción.
   - Solo agrega socios nuevos que no estén ya registrados en el condensado.


### 💰 Paso 3: Llenar el informe semanal de ahorros (`3-InformeAhorroSemanal.gs`)

1. **Abre el editor de Apps Script y pega el código de `3-InformeAhorroSemanal.gs`.**
2. **Ejecuta la función `llenarCondensadoAhorros`.**
   - El script detecta los bloques de semanas y meses en la hoja.
   - Busca la carpeta de cada socio (por número de socio y nombre) directamente en la carpeta principal.
   - Busca el archivo de ahorro correspondiente (por iniciales y nombre).
   - Llena las fórmulas de cada semana en el concentrado usando `IMPORTRANGE`, mostrando vacío si hay error o #N/A.
   - Solo procesa filas de socios (omite las últimas 3 filas de la hoja).
   - Solo llena hasta la última columna donde aparece "QUINTA" en las semanas de cada mes.
   - **Optimización:** Usa la carpeta principal directamente sin buscar subcarpetas, mejorando el rendimiento.
   - **Formato de mes y año:**  
     Los encabezados de los meses deben tener el formato `Mes Año` (por ejemplo, `Marzo 2025`, `Abril 2026`). El script detecta automáticamente el año y el mes para cada bloque y solo así funcionará correctamente.
   - Las fechas generadas en las fórmulas de semana siempre usan el último día válido del mes (por ejemplo, nunca pondrá el 31 de septiembre), evitando errores en las consultas de Google Sheets.


### 💸 Paso 4: Llenar el informe de préstamos semanales (`4-InformePrestamoSemanal.gs`)

1. **Abre el editor de Apps Script y pega el código de `4-InformePrestamoSemanal.gs`.**
2. **Ejecuta la función `llenarCondensadoPrestamos`.**
   - El script busca todas las hojas de préstamos (`Tarjeta Prestamo #1`, `Tarjeta Prestamo #2`, etc.) en cada archivo de socio.
   - **Renombrado automático:** Si alguna hoja no sigue el formato `'Tarjeta Prestamo #n'` y no es `'Tarjeta Ahorro'` ni `'Ahorros - No Activa'`, el script la renombra automáticamente como `'Tarjeta Prestamo #n'` usando el siguiente número disponible.  
   - **Esto es necesario para mantener la secuencia de préstamos tanto en el condensado como en las hojas individuales de cada socio.**
   - **Sistema de detección de duplicados mejorado:** Identifica préstamos únicos usando la combinación `código_socio#número_préstamo` y mantiene un registro interno de préstamos existentes.
   - Solo procesa préstamos nuevos que no están ya registrados en la hoja `Prestamos`.
   - Llena los datos básicos del préstamo (número, código, nombre, fecha, cantidad, pago pendiente, destino, interés, tipo de pago).
   - Calcula automáticamente los pagos mensuales (intereses y abonos) para cada mes del año.
   - Calcula la semana del mes del último abono realizado para cada mes.
   - **Optimización:** Busca carpetas de socios directamente en la carpeta principal.
   - **Formato de mes y año:**  
     Los reportes de préstamos también requieren que los encabezados de los meses estén en formato `Mes Año` (por ejemplo, `Marzo 2025`). El script busca y procesa por mes y año, no solo por nombre de mes.

3. **Configuración del Trigger Automático (MUY IMPORTANTE):**
   
   Para que el script detecte automáticamente nuevos préstamos cada día:
   
  a) **En el editor de Apps Script, ve al menú lateral izquierdo y selecciona "Activadores" (⏰ Triggers).**
   
   b) **Haz clic en "+ Agregar activador".**
   
   c) **Configura el activador así:**
   - **Función que se ejecutará:** `llenarCondensadoPrestamos`
   - **Evento de activación:** `Basado en tiempo`
   - **Tipo de activador basado en tiempo:** `Activador diario`
   - **Hora del día:** `9:00 a.m. - 10:00 a.m.` (o la hora que prefieras)
   
   d) **Haz clic en "Guardar".**
   
   e) **La primera vez te pedirá autorización. Haz clic en "Revisar permisos" y autoriza el acceso.**

4. **Verificación del funcionamiento:**
   - Después de configurar el trigger, el script se ejecutará automáticamente cada día.
   - Puedes verificar que funciona revisando los logs en Apps Script después de cada ejecución.
   - También puedes ejecutar manualmente la función cuando quieras actualizar inmediatamente.

5. **Ventajas del sistema automatizado:**
   - **Incremental:** Solo agrega préstamos nuevos, manteniendo el historial completo.
   - **Automático:** Se ejecuta sin intervención manual cada día.
   - **Eficiente:** Usa rangos específicos (B13:B23, D13:D23, etc.) para mejor rendimiento.

> **Nota:** Este script (`4-InformePrestamoSemanal.gs`) está diseñado para ejecutarse diariamente mediante un trigger automático. Es necesario configurar el trigger para que la función se ejecute cada día y así mantener actualizado el reporte de préstamos. Si no tienes el trigger activado, deberás ejecutarlo manualmente.

### 🤝 Paso 5: Procesar avales (`4.1-Avales.gs`)

1. **Abre el editor de Apps Script y pega el código de `4.1-Avales.gs`.**
2. **Ejecuta la función `procesarAvales`.**
   - **Procesamiento bidireccional mejorado:** El script ahora procesa avales en ambos sentidos:
     - **Aval → Prestatario:** Lee información de avales desde las hojas `Tarjeta Ahorro` y la registra en las hojas de préstamo correspondientes.
     - **Prestatario → Aval:** Busca inversamente desde las hojas de préstamo hacia las tarjetas de ahorro de los avales para completar información faltante.
   - **Detección inteligente de duplicados:** Verifica que los avales no se agreguen múltiples veces en ambos sentidos de procesamiento.
   - Agrega automáticamente la información del aval en las hojas de préstamo correspondientes.
   - Calcula el saldo pendiente del préstamo usando fórmulas `IMPORTRANGE` con filtros avanzados.
   - Obtiene la fecha de compromiso del préstamo.

3. **Funcionamiento del proceso bidireccional:**
   
   **Sentido Aval → Prestatario:**
   - Lee las últimas 3 filas de cada hoja `Tarjeta Ahorro` buscando información de avales.
   - Extrae: número de préstamo, código del prestatario, nombre del prestatario, y monto avalado.
   - Busca la hoja de préstamo correspondiente (`Tarjeta Prestamo #X`) en la tarjeta del prestatario.
   - Agrega el nombre del aval y el monto en las filas 4-7 de la columna H e I respectivamente.
   
   **Sentido Prestatario → Aval (Búsqueda Inversa):**
   - Recorre todas las hojas de préstamo de cada socio.
   - Identifica los avales registrados en las hojas de préstamo (filas 4-7, columna H).
   - Usa **normalización avanzada de nombres** (elimina acentos, convierte a mayúsculas) para buscar la carpeta del aval.
   - **Búsqueda flexible por palabras:** Busca coincidencias con al menos 2 palabras significativas del nombre del aval.
   - Agrega la información del préstamo en la hoja `Tarjeta Ahorro` del aval si no existe ya.
   - Evita duplicados verificando registros existentes antes de agregar nuevos.

4. **Configuración de ID dinámico:**
   - El script usa automáticamente el ID específico de cada tarjeta del prestatario (`prestatarioInfo.tarjetaId`)
   - Esto permite referenciar correctamente cada archivo individual de préstamo.
   - Las fórmulas se generan dinámicamente para cada socio y préstamo específico.

5. **Reportes detallados:**
   - El script registra en logs ambos sentidos de procesamiento por separado.
   - Muestra contadores independientes para:
     - Avales procesados de aval a prestatario
     - Avales procesados de prestatario a aval (búsqueda inversa)
   - Informa sobre duplicados detectados y omitidos.

> **Nota:** Este script (`4.1-Avales.gs`) también debe ejecutarse diariamente mediante un trigger para que cualquier cambio en los avales o préstamos se refleje automáticamente en los reportes y hojas correspondientes.

---

## 0. Migración desde la estructura antigua (scripts 0.x)

Si estás migrando desde una estructura o archivos antiguos hacia la estructura actual del proyecto, hay tres scripts auxiliares numerados como `0.x` cuyo objetivo es facilitar esa migración automática cuando los archivos nuevos no existen todavía.

Archivos y propósito rápido:
- `0.1-CarpetasInscripción.gs` — Escanea la carpeta principal (la misma que contiene el concentrado) y rellena la hoja `Inscripción` con la lista de carpetas de socios existentes (carpetas que empiezan por `GA...` o por el código de socio). Extrae número, nombres y apellidos desde el nombre de la carpeta y los escribe empezando en la fila 8. Útil si las carpetas ya existen pero no el `03 LISTA DE INSCRIPCION` o la hoja `Inscripción` está vacía.
- `0.2-ActualizacionHoja.gs` — Abre el archivo `03 LISTA DE INSCRIPCION` (debe estar en la misma carpeta principal), recorre las filas de la hoja `Inscripción` y para cada socio busca la carpeta correspondiente. Dentro de la carpeta del socio copia las hojas base (por ejemplo `Tarjeta Ahorro`) desde el concentrado hacia el archivo individual del socio, evita duplicados, actualiza B1/D1 en `Tarjeta Ahorro` y renombra el archivo de cada socio con el formato `(INICIALES) - TARJETA AHORRO Y PRESTAMO`.
- `0.3-ActualizacionPermisos.gs` — Copia las protecciones de rango (rango, descripción, warningOnly y editores) de las hojas `Tarjeta Ahorro` y `Tarjeta Prestamo #n` del archivo base (`04 TARJETA AHORRO Y PRESTAMO`) hacia los archivos de cada socio. Compara para no recrear protecciones idénticas y elimina protecciones en los archivos de socio que ya no existan en el archivo base.

Orden recomendado de ejecución (cuando migras desde lo antiguo):
1. Abrir el archivo base (`04 TARJETA AHORRO Y PRESTAMO`) y el concentrado en el editor de Apps Script para autorizar accesos si es necesario.
2. Ejecutar `0.1-CarpetasInscripción.gs` → función `llenarHojaInscripcion()` para generar/llenar la hoja `Inscripción` con las carpetas existentes.
3. Ejecutar `0.2-ActualizacionHoja.gs` → función `renombrarYAgregarHojas()` para copiar las hojas base a cada archivo de socio y renombrar archivos.
4. Ejecutar `0.3-ActualizacionPermisos.gs` → función `actualizarPermisosProteccion()` para propagar las protecciones de rango desde el archivo base a los archivos de socio.

Precauciones y notas:
- Haz un respaldo antes de ejecutar los scripts en bloque. Los scripts editan archivos y protecciones en masa.
- Asegúrate de abrir manualmente los archivos que usan `IMPORTRANGE` la primera vez para autorizar el acceso.
- Los scripts asumen que las carpetas de los socios están en la misma carpeta principal que el concentrado (sin subcarpetas intermedias).
- Si tu estructura de nombres es distinta (por ejemplo las carpetas no empiezan por el número de socio seguido de un espacio), ajusta las condiciones de búsqueda en los scripts antes de ejecutarlos.
- Ejecuta los scripts en el orden indicado y revisa los logs de Apps Script para ver el detalle de lo realizado.


## 2.1. Actualización de permisos de rangos protegidos en archivos de socios

Si necesitas modificar los permisos de los rangos protegidos (por ejemplo, cambiar o agregar editores) en todos los archivos de los socios, puedes automatizar este proceso usando el script `ActualizacionPermisos.gs`.


### ⏰ ¿Cuándo usarlo?
- Cuando cambias los permisos de los rangos protegidos en el archivo base (`04 TARJETA AHORRO Y PRESTAMO`) y quieres que esos mismos permisos se apliquen en todos los archivos de los socios.
- Cuando necesitas agregar o quitar editores de los rangos protegidos de las hojas "Tarjeta Ahorro" y "Tarjeta Prestamo #n" en todos los archivos de los socios.


### ⚡ ¿Qué hace el script?
- **Optimización:** El script ahora compara las protecciones existentes en cada archivo de socio con las del archivo base.  
  Solo elimina y vuelve a crear las protecciones que realmente son diferentes o faltan.  
  Las protecciones que ya coinciden (mismo rango, descripción, advertencia y editores) se dejan intactas, ahorrando tiempo y evitando agotar los límites de Google Apps Script.
- Elimina protecciones que no existen en el archivo base.
- Crea nuevas protecciones exactamente iguales a las del archivo base (`04 TARJETA AHORRO Y PRESTAMO`), incluyendo los mismos rangos, editores, advertencias y descripciones.
- **No elimina ni modifica el contenido de las celdas, solo actualiza las reglas de protección.**


### 🛠️ ¿Cómo usarlo?
1. **Actualiza los permisos de los rangos protegidos en el archivo base (`04 TARJETA AHORRO Y PRESTAMO`) como desees.**
2. **Abre el archivo base y ve a Extensiones > Apps Script.**
3. **Pega el código de `ActualizacionPermisos.gs` y ejecuta la función `actualizarPermisosProteccion()`.**
4. **El script recorrerá todas las carpetas de socios y actualizará los rangos protegidos en las hojas correspondientes.**
5. **Revisa los logs para ver el resumen de archivos y protecciones actualizadas.**

> **Nota:** Solo se modifican las hojas que se llamen exactamente "Tarjeta Ahorro" o que empiecen con "Tarjeta Prestamo" (por ejemplo, "Tarjeta Prestamo #1", "Tarjeta Prestamo #2", etc.).

---


## 3. Consideraciones y recomendaciones

- **Estructura de carpetas simplificada:**  
  Las carpetas de los socios deben estar en la misma carpeta principal donde está la hoja de cálculo, sin subcarpetas intermedias. Esto mejora significativamente el rendimiento de los scripts.
- **Triggers automáticos:**  
  Si configuras triggers para automatización, revisa periódicamente que estén funcionando correctamente en la sección "Ejecuciones" del editor de Apps Script.
- **Gestión de préstamos:**  
  El sistema de préstamos está diseñado para ser incremental con detección avanzada de duplicados. Si necesitas regenerar completamente la hoja de préstamos, elimina manualmente el contenido y ejecuta la función nuevamente.
- **Sistema de avales bidireccional:**  
  El procesamiento de avales ahora funciona en ambos sentidos, asegurando que toda la información esté sincronizada tanto en las tarjetas de los avales como en las hojas de préstamo de los prestatarios.
- **Normalización de nombres:**  
  El script de avales usa normalización avanzada (eliminación de acentos, conversión a mayúsculas) para mejorar la búsqueda y coincidencia de nombres, reduciendo errores por diferencias de formato.
- **Rendimiento optimizado:**  
  Los scripts usan rangos específicos (como B13:B23) en lugar de columnas completas para mejorar el rendimiento y evitar timeouts.
- **Permisos de edición y acceso:**  
  Si el archivo base tiene celdas protegidas, asegúrate de tener permisos para editarlas antes de hacer las copias.
- **IMPORTRANGE:**  
  La primera vez que uses `IMPORTRANGE` para un archivo nuevo, deberás autorizar el acceso manualmente en la celda correspondiente.
- **Logs y reportes:**  
  Los scripts optimizados proporcionan reportes detallados en los logs para monitorear el proceso y verificar qué elementos fueron creados vs. existentes.
- **Nombres y formatos:**  
  El script capitaliza automáticamente el nombre completo en las tarjetas de ahorro.
- **Búsqueda flexible de avales:**  
  El sistema de búsqueda inversa de avales tolera variaciones en los nombres, buscando coincidencias con al menos 2 palabras significativas para mayor precisión.
- **Precisión decimal:**  
  Todos los reportes y hojas usan **dos decimales** para mostrar montos y cálculos con mayor exactitud.

### 📅 ¿Cómo configurar un trigger automático en Apps Script?

Para que los scripts de préstamos y avales se ejecuten solos cada día:

1. Abre el editor de Apps Script del archivo correspondiente.
2. Ve al menú lateral izquierdo y selecciona **"Activadores"** (ícono de reloj).
3. Haz clic en **"+ Agregar activador"**.
4. Elige la función que quieres automatizar (`llenarCondensadoPrestamos` o `procesarAvales`).
5. En "¿Con qué frecuencia?" selecciona **"Basado en tiempo"** y luego **"Diariamente"**.
6. Elige la hora que prefieras para la ejecución automática.
7. Guarda y acepta los permisos si es la primera vez.

Así, los reportes de préstamos y avales se mantendrán siempre actualizados sin intervención manual.