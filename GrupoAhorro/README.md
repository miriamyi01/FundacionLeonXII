# 📂 Instrucciones para automatización de carpetas y archivos de socios

Este proyecto automatiza la creación y llenado de carpetas y archivos para socios usando Google Apps Script y Google Drive.

---



## Guía para la automatización de carpetas y archivos de socios

### 1. Requisitos previos

- **Permisos de acceso:**
  - Antes de ejecutar cualquier script, abre manualmente el archivo base (`04 TARJETA AHORRO Y PRESTAMO`) y el archivo de concentrado para otorgar permisos de acceso a `IMPORTRANGE` y edición.
  - La primera vez que uses `IMPORTRANGE` en el concentrado, deberás autorizar el acceso a cada archivo de socio. Hazlo manualmente para cada fórmula nueva.

- **Archivos base:**
  - El archivo base debe llamarse exactamente `04 TARJETA AHORRO Y PRESTAMO` y estar en la carpeta principal de Drive.
  - Asegúrate de que las celdas protegidas en este archivo base estén desbloqueadas o tengas permisos para editarlas, ya que las copias heredarán estos permisos.

- **Estructura de la hoja de cálculo:**
  - La hoja de cálculo principal debe estar dentro de una carpeta en Google Drive.
  - En la hoja de cálculo, la celda A8 debe tener el código base (por ejemplo, `2024`).
  - Los datos de los socios deben estar en las columnas:
    - A: Número de socio (desde la fila 8)
    - B: Nombre
    - C: Primer apellido
    - D: Segundo apellido
  - **Meses y años:** Los encabezados de los meses en las hojas deben tener el formato `Mes Año` (por ejemplo, `Marzo 2025`, `Abril 2026`).
    - No pongas solo el nombre del mes (ejemplo: `Marzo`). Siempre debe incluir el año.

- **Precisión de datos:**
  - Todas las hojas y reportes usan dos decimales para mayor precisión en los cálculos y presentación de montos.

- **Organización de carpetas:**
  - Las carpetas de los socios ya no tienen que estar en la raíz de la carpeta principal. Los scripts buscan y procesan las carpetas de socios dentro de la estructura organizada (`GA0452-SOCIOS AS`), permitiendo mejor gestión y orden.

---

---


## 2. Flujo de automatización


> 📝 **NOTA IMPORTANTE:** Los scripts `1-Carpetas.gs`, `2-InformeInicial.gs`, `3-InformeAhorroSemanal.gs`, `4-InformePrestamoSemanal.gs` y `4.1-Avales.gs` están optimizados para **ejecución múltiple sin duplicados**. Puedes ejecutarlos tantas veces como necesites de forma segura - detectan automáticamente elementos existentes y solo procesan información nueva. Todos proporcionan reportes detallados en los logs para monitorear el proceso.



### 📁 Paso 1: Crear carpetas de socios (`1-Carpetas.gs`)

1. **Coloca la hoja de cálculo y el archivo base en la misma carpeta de Google Drive.**
2. **Abre el editor de Apps Script y pega el código de `1-Carpetas.gs`.**
3. **Ejecuta la función `crearCarpetasSocios`.**
  - El script busca la carpeta raíz `GA0452 METAMORFOSIS` y dentro de ella la subcarpeta `GA0452-SOCIOS AS`.
  - Para cada socio, crea una carpeta con el formato `[Número de socio] [INICIALES] [Nombre completo]` dentro de `GA0452-SOCIOS AS` si no existe.
  - No copia archivos ni hojas en este paso, solo crea carpetas.


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
  - Busca la carpeta de cada socio dentro de `GA0452-SOCIOS AS` (por número de socio y nombre).
  - Busca el archivo de ahorro correspondiente (por iniciales y nombre).
  - Llena las fórmulas de cada semana en el concentrado usando `IMPORTRANGE`, mostrando vacío si hay error o #N/A.
  - Solo procesa filas de socios (omite las últimas 3 filas de la hoja).
  - Solo llena hasta la última columna donde aparece "QUINTA" en las semanas de cada mes.
  - **Optimización:** Usa la estructura organizada de carpetas, mejorando el rendimiento.
  - **Formato de mes y año:** Los encabezados de los meses deben tener el formato `Mes Año` (por ejemplo, `Marzo 2025`, `Abril 2026`). El script detecta automáticamente el año y el mes para cada bloque y solo así funcionará correctamente.
  - Las fechas generadas en las fórmulas de semana siempre usan el último día válido del mes (por ejemplo, nunca pondrá el 31 de septiembre), evitando errores en las consultas de Google Sheets.



### 💸 Paso 4: Llenar el informe de préstamos semanales (`4-InformePrestamoSemanal.gs`)

1. **Abre el editor de Apps Script y pega el código de `4-InformePrestamoSemanal.gs`.**
2. **Ejecuta la función `llenarCondensadoPrestamos`.**
  - El script busca todas las hojas de préstamos (`Tarjeta Prestamo #1`, `Tarjeta Prestamo #2`, etc.) en cada archivo de socio dentro de `GA0452-SOCIOS AS`.
  - **Renombrado automático:** Si alguna hoja no sigue el formato `'Tarjeta Prestamo #n'` y no es `'Tarjeta Ahorro'` ni `'Ahorros - No Activa'`, el script la renombra automáticamente como `'Tarjeta Prestamo #n'` usando el siguiente número disponible.
  - **Esto es necesario para mantener la secuencia de préstamos tanto en el condensado como en las hojas individuales de cada socio.**
  - **Sistema de detección de duplicados mejorado:** Identifica préstamos únicos usando la combinación `código_socio#número_préstamo` y mantiene un registro interno de préstamos existentes.
  - Solo procesa préstamos nuevos que no están ya registrados en la hoja `Prestamos`.
  - Llena los datos básicos del préstamo (número, código, nombre, fecha, cantidad, pago pendiente, destino, interés, tipo de pago).
  - Calcula automáticamente los pagos mensuales (intereses y abonos) para cada mes del año.
  - Calcula la semana del mes del último abono realizado para cada mes.
  - **Optimización:** Usa la estructura organizada de carpetas, mejorando el rendimiento.
  - **Formato de mes y año:** Los reportes de préstamos también requieren que los encabezados de los meses estén en formato `Mes Año` (por ejemplo, `Marzo 2025`). El script busca y procesa por mes y año, no solo por nombre de mes.

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


## 0. Scripts de migración y actualización (Archivos/)

### 0.1-AñosAnteriores.gs
**Función:** moverYOrganizarAnios
- Mueve archivos de años anteriores (excepto el año actual y el archivo base) a la carpeta `01 LISTAS AÑOS ANTERIORES`.
- Organiza el ahorro de socios en carpetas por años previos dentro de cada carpeta de socio (`01 AHORRO AÑOS ANTERIORES`), creando subcarpetas para los dos años anteriores.
- Mueve archivos de ahorro antiguos a la subcarpeta correspondiente según el año.

### 0.2-Carpetas.gs
**Función:** crearCarpetasSocios
- Crea carpetas de socios dentro de `GA0452-SOCIOS AS` usando la hoja `Inscripción`.
- Detecta si la carpeta ya existe (por número y nombre completo, ignorando iniciales).
- Si no existe, la crea con el formato: `[Número de socio] [INICIALES] [Nombre completo]`.
- Registra en logs el total de socios procesados, carpetas nuevas y existentes.

### 0.3-CreaciónHoja.gs
**Función:** renombrarYAgregarHojas
- Copia el archivo base `04 TARJETA AHORRO Y PRESTAMO` a la carpeta de cada socio, usando las iniciales para el nombre del archivo.
- Si el archivo ya existe, lo actualiza.
- Copia las hojas del archivo base al archivo del socio, evitando duplicados.
- Actualiza datos clave en las hojas (`B1` con el nombre completo, `D1` con el número de socio, `F1` con fórmula IMPORTRANGE).
- Renombra el archivo si es necesario.

### 0.4-PermisosHoja.gs
**Función:** actualizarPermisosProteccion
- Propaga los rangos protegidos y permisos del archivo base a los archivos de cada socio.
- Solo actualiza protecciones que sean diferentes o faltantes, evitando duplicados.
- Elimina protecciones que ya no existen en el archivo base.
- Aplica los mismos editores, advertencias y descripciones.

### 0.5-ActualizaciónHojas.gs
**Función:** reemplazarHojaPrestamo1Socios
- Elimina la hoja `Tarjeta Prestamo #1` en cada archivo de socio.
- Copia la hoja `Tarjeta Prestamo #1` del archivo base al archivo de cada socio.
- Reordena las hojas principales: `Tarjeta Ahorro`, `Tarjeta Prestamo #1`, `Fondo de Emergencia`.
- Registra en logs los archivos actualizados.

---

## Cambios y flujo actualizado

- **Las funciones de los scripts en Archivos/ ahora están especializadas en migración, organización, creación y actualización masiva de carpetas, archivos y hojas de socios.**
- **La estructura de carpetas de socios se mantiene dentro de `GA0452-SOCIOS AS` en la carpeta principal, pero los scripts ya no requieren que las carpetas de socios estén en la raíz. Ahora buscan y procesan las carpetas de socios dentro de la estructura organizada, mejorando la gestión y el orden.**
- **Los scripts detectan y evitan duplicados, actualizan solo lo necesario y optimizan el rendimiento.**
- **Todos los procesos registran logs detallados para monitoreo.**
- **La actualización de permisos y hojas es incremental y solo modifica lo que es diferente respecto al archivo base.**

---


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