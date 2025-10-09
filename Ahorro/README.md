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
  - En la hoja de cálculo, la celda A8 debe tener el código base (por ejemplo, `2024`).
  - Los datos de los socios deben estar en las columnas:
    - A: Número de socio (desde la fila 8)
    - B: Nombre
    - C: Primer apellido
    - D: Segundo apellido

---

## 2. Flujo de automatización

### Paso 1: Crear carpetas y archivos de socios (`1-Carpetas.gs`)

1. **Coloca la hoja de cálculo y el archivo base en la misma carpeta de Google Drive.**
2. **Abre el editor de Apps Script y pega el código de `1-Carpetas.gs`.**
3. **Ejecuta la función `crearCarpetasSocios`.**
   - El script creará una subcarpeta `[XXXX-SOCIOS AS]` (donde `XXXX` son las primeras 4 letras del código base de A8).
   - Para cada socio, creará una carpeta con el formato:  
     `[Número de socio] [INICIALES] [Nombre completo]`
   - Dentro de cada carpeta de socio, copiará el archivo base y lo renombrará con las iniciales.
   - En cada copia, pondrá el nombre completo (capitalizado) en B1 y el número de socio en D1.

### Paso 2: Registrar socios en el concentrado (`2-InformeInicial.gs`)

1. **Abre el editor de Apps Script y pega el código de `2-InformeInicial.gs`.**
2. **Ejecuta la función `registrarSociosCondensado`.**
   - El script llenará la hoja `Ahorros y Retiros` con los datos de los socios desde el archivo `03 LISTA DE INSCRIPCION`.
   - Ajusta automáticamente el número de filas y coloca fórmulas `IMPORTRANGE` para importar los datos de inscripción.

### Paso 3: Llenar el informe semanal (`3-InformeSemanal.gs`)

1. **Abre el editor de Apps Script y pega el código de `3-InformeSemanal.gs`.**
2. **Ejecuta la función `llenarCondensadoAhorros`.**
   - El script detecta los bloques de semanas y meses en la hoja.
   - Busca la carpeta de cada socio (por número de socio y nombre) dentro de la subcarpeta `[XXXX-SOCIOS AS]`.
   - Busca el archivo de ahorro correspondiente (por iniciales y nombre).
   - Llena las fórmulas de cada semana en el concentrado usando `IMPORTRANGE`, mostrando vacío si hay error o #N/A.
   - Solo procesa filas de socios (omite las últimas 3 filas de la hoja).
   - Solo llena hasta la última columna donde aparece "QUINTA" en las semanas de cada mes.

### Paso 4: Llenar el informe de préstamos semanales (`4-InformePrestamoSemanal.gs`)

1. **Abre el editor de Apps Script y pega el código de `4-InformePrestamoSemanal.gs`.**
2. **Ejecuta la función `llenarCondensadoPrestamos`.**
   - El script busca todas las hojas de préstamos (`Tarjeta Prestamo #1`, `Tarjeta Prestamo #2`, etc.) en cada archivo de socio.
   - **Detección inteligente de préstamos nuevos:** Solo procesa préstamos que no están ya registrados en la hoja `Prestamos`.
   - Identifica préstamos únicos usando la combinación `código_socio#número_préstamo`.
   - Llena los datos básicos del préstamo (número, código, nombre, fecha, cantidad, pago pendiente, destino, interés, tipo de pago).
   - Calcula automáticamente los pagos mensuales (intereses y abonos) para cada mes del año.
   - Calcula la semana del mes del último abono realizado para cada mes.
   - **Preserva datos existentes:** No borra ni sobrescribe préstamos que ya están en el informe.

3. **Configuración del Trigger Automático (MUY IMPORTANTE):**
   
   Para que el script detecte automáticamente nuevos préstamos cada semana:
   
   a) **En el editor de Apps Script, ve al menú lateral izquierdo y selecciona "Activadores" (⏰ Triggers).**
   
   b) **Haz clic en "+ Agregar activador".**
   
   c) **Configura el activador así:**
   - **Función que se ejecutará:** `llenarCondensadoPrestamos`
   - **Evento de activación:** `Basado en tiempo`
   - **Tipo de activador basado en tiempo:** `Activador semanal`
   - **Día de la semana:** `Lunes` (recomendado)
   - **Hora:** `9:00 a.m. - 10:00 a.m.` (o la hora que prefieras)
   
   d) **Haz clic en "Guardar".**
   
   e) **La primera vez te pedirá autorización. Haz clic en "Revisar permisos" y autoriza el acceso.**

4. **Verificación del funcionamiento:**
   - Después de configurar el trigger, el script se ejecutará automáticamente cada semana.
   - Puedes verificar que funciona revisando los logs en Apps Script después de cada ejecución.
   - También puedes ejecutar manualmente la función cuando quieras actualizar inmediatamente.

5. **Ventajas del sistema automatizado:**
   - **Sin duplicados:** No vuelve a procesar préstamos que ya están registrados.
   - **Incremental:** Solo agrega préstamos nuevos, manteniendo el historial completo.
   - **Automático:** Se ejecuta sin intervención manual cada semana.
   - **Eficiente:** Usa rangos específicos (B13:B23, D13:D23, etc.) para mejor rendimiento.

---

## 3. Consideraciones y recomendaciones

- **Triggers automáticos:**  
  Si configuras triggers para automatización, revisa periódicamente que estén funcionando correctamente en la sección "Ejecuciones" del editor de Apps Script.
- **Gestión de préstamos:**  
  El sistema de préstamos está diseñado para ser incremental. Si necesitas regenerar completamente la hoja de préstamos, elimina manualmente el contenido y ejecuta la función nuevamente.
- **Rendimiento optimizado:**  
  Los scripts usan rangos específicos (como B13:B23) en lugar de columnas completas para mejorar el rendimiento y evitar timeouts.
- **Permisos de edición y acceso:**  
  Si el archivo base tiene celdas protegidas, asegúrate de tener permisos para editarlas antes de hacer las copias.
- **IMPORTRANGE:**  
  La primera vez que uses `IMPORTRANGE` para un archivo nuevo, deberás autorizar el acceso manualmente en la celda correspondiente.
- **Logs:**  
  Puedes revisar los logs en el editor de Apps Script para ver detalles del proceso y errores.
- **Nombres y formatos:**  
  El script capitaliza automáticamente el nombre completo en las tarjetas de ahorro.

---

## 4. Contacto

Para dudas o mejoras, contacta a miriam08.mr@gmail.com