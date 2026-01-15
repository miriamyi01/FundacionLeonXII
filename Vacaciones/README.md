# 🏖️ Guía rápida para automatización de vacaciones y aniversarios

Esta carpeta contiene los scripts y plantillas para automatizar la gestión de vacaciones y el envío de correos de aniversario en la Fundación León XIII. El objetivo es que el proceso sea claro, eficiente y repetible cada año. 🚀

## 📂 ¿Qué hay en esta carpeta?

- **1-ActualizaciónBase.gs**: Agrega automáticamente las columnas y fórmulas del año actual en la hoja "Base Vacaciones". Corre cada mes (el día 1).
- **2-ArchivosAnuales.gs**: Genera archivos individuales de vacaciones para cada colaborador y actualiza el índice anual. Corre cada mes (el día 1), no genera duplicados, solo actualiza datos vacíos si el archivo ya existe y elimina los días sobrantes en las hojas individuales para mejor visualización.
 - **3-PermisosDiarios.gs**: Asigna permisos y protecciones en los archivos individuales el día del aniversario de cada colaborador, o si ya pasó y alguno no fue asignado. Corre diariamente y solo realiza cambios cuando corresponde.
- **Aniversario/CorreoAniversario.html**: Plantilla de correo para felicitar en aniversarios, mostrando nombre, años y periodo vacacional.
- **Aniversario/TriggerAniversario.gs**: Envía correos de aniversario y notifica al jefe directo, calculando días de vacaciones según antigüedad.

---

## ⚡ ¿Cómo usar los scripts?

1. **Actualización mensual de la base:**  
   El script `1-ActualizaciónBase.gs` se encarga de agregar automáticamente las columnas y fórmulas necesarias para el año actual en la hoja "Base Vacaciones".  
   - Corre automáticamente el primer día de cada mes mediante un trigger, por lo que normalmente no necesitas ejecutarlo manualmente.
   - Si agregas un colaborador nuevo o cambias la estructura de la hoja, puedes ejecutarlo manualmente para asegurarte de que todo esté actualizado.

2. **Generación y actualización de archivos individuales:**  
   El script `2-ArchivosAnuales.gs` crea los archivos individuales de vacaciones para cada colaborador y actualiza el índice anual.  
   - También corre automáticamente el primer día de cada mes.
   - Si el archivo de vacaciones de un colaborador ya existe, solo actualiza los datos vacíos y nunca genera duplicados.
   - Elimina los días sobrantes en las hojas individuales para que solo se muestren los días realmente asignados, facilitando la visualización.
   - Actualiza la hoja "Aniversarios_AAAA" eliminando hojas antiguas (más de 10 años) y agregando solo colaboradores nuevos.
   - Ajusta el formato y las filas de la hoja de aniversarios para evitar errores.

3. **Actualización de la hoja de aniversarios:**  
   La hoja "Aniversarios_AAAA" (AAA es el año actual) se crea o actualiza automáticamente dentro del archivo `001 - Índice: Archivo de vacaciones` al ejecutar `2-ArchivosAnuales.gs`.  
   Incluye los siguientes datos:
   - Correo
   - Fecha de ingreso
   - Link al archivo de vacaciones
   - Correo Jefe Directo (opcional)

4. **Asignación diaria de permisos:**  
   El script `3-PermisosDiarios.gs` corre diariamente y asigna permisos y protecciones en los archivos individuales el día del aniversario de cada colaborador, o si ya pasó y alguno no fue asignado.  
   - No necesitas ejecutarlo manualmente, salvo que el trigger se haya borrado.

5. **Envío automático de correos de aniversario:**  
   El script `TriggerAniversario.gs` envía automáticamente correos de aniversario y notifica al jefe directo, calculando los días de vacaciones según la antigüedad.  
   - El trigger ya está activo y solo debes ejecutarlo manualmente si se borra o desactiva.

---

## 🔄 ¿Qué hacer si agregas un colaborador nuevo?

- Agrega el colaborador en la hoja "Base Vacaciones" con todos sus datos.
- Ejecuta `1-ActualizaciónBase.gs` para actualizar las fórmulas y columnas del año actual (o espera a la siguiente ejecución automática).
- Ejecuta `2-ArchivosAnuales.gs` para generar su archivo individual y actualizar la hoja índice y la hoja de aniversarios (o espera a la siguiente ejecución automática).
- Verifica que los datos estén completos (correo, fecha de ingreso, entre otros) para evitar que el script lo omita.

---

## 🚨 Advertencias y consideraciones importantes

- La hoja índice `001 - Índice: Archivo de vacaciones` debe existir antes de correr `2-ArchivosAnuales.gs`. Si no existe, el script no la crea y mostrará un error en el log.
- Si la hoja base o la hoja índice no tienen los encabezados esperados, los scripts pueden fallar o no procesar correctamente los datos.
- Los colaboradores sin fecha de ingreso no tendrán archivo ni aparecerán en la hoja de aniversarios.
- Si falta algún dato clave, el script puede omitir al colaborador sin notificarlo.
- Prueba los scripts con pocos datos antes de hacer envíos masivos.
- Revisa el log de Apps Script para ver el estado y posibles errores. Ejemplo de error común:
  - `No se encontró archivo índice para duplicar`
  - `Error creando archivo socio idx ...`
- El cálculo de días de vacaciones está automatizado según la antigüedad, pero puedes ajustar la tabla en el script si cambian las políticas.
- Personaliza los textos y la plantilla HTML si lo deseas.
- Consulta los comentarios en cada script para entender y adaptar el proceso.
- Los triggers para los scripts mensuales y diarios ya están configurados para ejecutarse automáticamente (mensuales el día 1 de cada mes, diarios para aniversarios y permisos). Solo revisa o reactiva los triggers si se borran o desactivan.

---

## 🛠️ Sobre los triggers automáticos

- Los triggers para ejecución mensual (actualización de base y archivos individuales) y diaria (permisos y correos de aniversario) ya están activos en el proyecto de Apps Script.
- Solo necesitas ejecutar manualmente los scripts si el trigger fue borrado o desactivado.
- Para revisar o reactivar los triggers:
  1. Abre el editor de Apps Script.
  2. Ve a "Triggers" en el menú izquierdo.
  3. Verifica que existan triggers para las funciones correspondientes (mensual y diaria).
  4. Si no existen, crea uno nuevo y selecciona la función y el tipo de disparador (por tiempo, mensual o diario).

---

## 🚀 Ventajas

- Automatización real y eficiente de la gestión de vacaciones y aniversarios.
- Archivos y correos personalizados para cada colaborador.
- Visualización clara y sin filas/días sobrantes en los archivos individuales.
- Proceso reutilizable y fácil de adaptar año con año.