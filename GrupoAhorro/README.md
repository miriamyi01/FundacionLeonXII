# Instrucciones para automatización de carpetas y archivos de socios

Este proyecto automatiza la creación y llenado de carpetas y archivos para socios usando Google Apps Script y Google Drive.

---

## 1. Requisitos previos

- **Permisos de acceso:**  
  Antes de ejecutar cualquier script, asegúrate de abrir manualmente cada archivo base (`04 TARJETA AHORRO Y PRESTAMO`) y el archivo de concentrado (la hoja de cálculo principal) para otorgar permisos de acceso a `IMPORTRANGE` y edición si es necesario.  
  > **Nota:** La primera vez que uses `IMPORTRANGE` en el concentrado, deberás autorizar el acceso a cada archivo de socio. Hazlo manualmente para cada fórmula nueva, después ya no será necesario.

- **Archivos base:**  
  El archivo base debe llamarse exactamente:  
  **`04 TARJETA AHORRO Y PRESTAMO`**  
  y debe estar en la carpeta principal de Drive.  
  > **Importante:** Antes de ejecutar el script, asegúrate de que las celdas protegidas (con fórmulas) en este archivo base estén desbloqueadas o que tengas permisos para editarlas, ya que las copias heredarán estos permisos.

- **Estructura de la hoja de cálculo:**  
  - La hoja de cálculo principal debe estar dentro de una carpeta en Google Drive.
  - **IMPORTANTE:** Las carpetas de los socios deben estar en la misma carpeta principal donde está la hoja de cálculo (sin subcarpetas intermedias).
  - En la hoja de cálculo, la celda A8 debe tener el código base (por ejemplo, `2024`).
  - Los datos de los socios deben estar en las columnas:
    - A: Número de socio (desde la fila 8)
    - B: Nombre
    - C: Primer apellido
    - D: Segundo apellido

---

## 2. Flujo de automatización

> **📝 NOTA IMPORTANTE:** Los scripts `1-Carpetas.gs`, `2-InformeInicial.gs`, `3-InformeAhorroSemanal.gs`, `4-InformePrestamoSemanal.gs` y `4.1-Avales.gs` están optimizados para **ejecución múltiple sin duplicados**. Puedes ejecutarlos tantas veces como necesites de forma segura - detectan automáticamente elementos existentes y solo procesan información nueva. Todos proporcionan reportes detallados en los logs para monitorear el proceso.

### Paso 1: Crear carpetas y archivos de socios (`1-Carpetas.gs`)

1. **Coloca la hoja de cálculo y el archivo base en la misma carpeta de Google Drive.**
2. **Abre el editor de Apps Script y pega el código de `1-Carpetas.gs`.**
3. **Ejecuta la función `crearCarpetasSocios`.**
   - El script creará carpetas de socios directamente en la carpeta principal (misma carpeta donde está la hoja de cálculo).
   - Para cada socio, creará una carpeta con el formato:  
     `[Número de socio] [INICIALES] [Nombre completo]`
   - Dentro de cada carpeta de socio, copiará el archivo base y lo renombrará con las iniciales.
   - En cada copia, pondrá el nombre completo (capitalizado) en B1 y el número de socio en D1.

### Paso 2: Registrar socios en el concentrado (`2-InformeInicial.gs`)

1. **Abre el editor de Apps Script y pega el código de `2-InformeInicial.gs`.**
2. **Ejecuta la función `registrarSociosCondensado`.**
   - El script llenará la hoja `Ahorros y Retiros` con los datos de los socios desde el archivo `03 LISTA DE INSCRIPCION`.
   - Ajusta automáticamente el número de filas y coloca fórmulas `IMPORTRANGE` para importar los datos de inscripción.
   - Solo agrega socios nuevos que no estén ya registrados en el condensado.

### Paso 3: Llenar el informe semanal de ahorros (`3-InformeAhorroSemanal.gs`)

1. **Abre el editor de Apps Script y pega el código de `3-InformeAhorroSemanal.gs`.**
2. **Ejecuta la función `llenarCondensadoAhorros`.**
   - El script detecta los bloques de semanas y meses en la hoja.
   - Busca la carpeta de cada socio (por número de socio y nombre) directamente en la carpeta principal.
   - Busca el archivo de ahorro correspondiente (por iniciales y nombre).
   - Llena las fórmulas de cada semana en el concentrado usando `IMPORTRANGE`, mostrando vacío si hay error o #N/A.
   - Solo procesa filas de socios (omite las últimas 3 filas de la hoja).
   - Solo llena hasta la última columna donde aparece "QUINTA" en las semanas de cada mes.
   - **Optimización:** Usa la carpeta principal directamente sin buscar subcarpetas, mejorando el rendimiento.

### Paso 4: Llenar el informe de préstamos semanales (`4-InformePrestamoSemanal.gs`)

1. **Abre el editor de Apps Script y pega el código de `4-InformePrestamoSemanal.gs`.**
2. **Ejecuta la función `llenarCondensadoPrestamos`.**
   - El script busca todas las hojas de préstamos (`Tarjeta Prestamo #1`, `Tarjeta Prestamo #2`, etc.) en cada archivo de socio.
   - **Sistema de detección de duplicados mejorado:** Identifica préstamos únicos usando la combinación `código_socio#número_préstamo` y mantiene un registro interno de préstamos existentes.
   - Solo procesa préstamos nuevos que no están ya registrados en la hoja `Prestamos`.
   - Llena los datos básicos del préstamo (número, código, nombre, fecha, cantidad, pago pendiente, destino, interés, tipo de pago).
   - Calcula automáticamente los pagos mensuales (intereses y abonos) para cada mes del año.
   - Calcula la semana del mes del último abono realizado para cada mes.
   - **Optimización:** Busca carpetas de socios directamente en la carpeta principal.

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

### Paso 5: Procesar avales (`4.1-Avales.gs`)

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

### Paso 6: Generar condensado final (`5-CondensadoFinal.gs`)

1. **Abre el editor de Apps Script y pega el código de `5-CondensadoFinal.gs`.**
2. **Ejecuta la función correspondiente para generar el reporte final.**
   - Este paso consolida toda la información procesada en los pasos anteriores.
   - Genera reportes finales con los datos de ahorros, préstamos y avales.

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

---

## 4. Contacto

Para dudas o mejoras, contacta a miriam08.mr@gmail.com