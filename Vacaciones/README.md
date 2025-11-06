# 🏖️ Guía rápida para automatización de vacaciones y aniversarios

Esta carpeta contiene los scripts y plantillas para automatizar la gestión de vacaciones y el envío de correos de aniversario en la Fundación León XIII. El objetivo es que el proceso sea claro, eficiente y repetible cada año. 🚀

## 📂 ¿Qué hay en esta carpeta?

- **1-ActualizaciónBase.gs**: Agrega automáticamente las columnas y fórmulas del año siguiente en la hoja "Base Vacaciones".
- **2-ArchivosAnuales.gs**: Genera archivos individuales de vacaciones para cada colaborador y actualiza el índice anual.
- **Aniversario/CorreoAniversario.html**: Plantilla de correo para felicitar en aniversarios, mostrando nombre, años y periodo vacacional.
- **Aniversario/TriggerAniversario.gs**: Envía correos de aniversario y notifica al jefe directo, calculando días de vacaciones según antigüedad.

---

## ⚡ ¿Cómo usar los scripts?

1. Ejecuta `1-ActualizaciónBase.gs` en la hoja "Base Vacaciones" para preparar el año siguiente.
2. Ejecuta `2-ArchivosAnuales.gs` para crear los archivos individuales y actualizar el índice anual.
3. La hoja "Aniversarios_AAAA" (AAA es el año actual) se crea automáticamente dentro del archivo `001 - Índice: Archivo de vacaciones` de la carpeta principal al ejecutar `2-ArchivosAnuales.gs`. Incluye los datos:
   - Correo
   - Fecha de ingreso
   - Link al archivo de vacaciones
   - Correo Jefe Directo (opcional)
4. No es necesario ejecutar manualmente `TriggerAniversario.gs`, ya que el trigger para envío automático de correos de aniversario ya está activo. Solo ejecuta el script si el trigger fue borrado o desactivado.

## 🔄 ¿Qué hacer si agregas un colaborador nuevo?

- Agrega el colaborador en la hoja "Base Vacaciones" con todos sus datos.
- Ejecuta `1-ActualizaciónBase.gs` para actualizar las fórmulas y columnas del año siguiente.
- Ejecuta `2-ArchivosAnuales.gs` para generar su archivo individual y actualizar la hoja índice y la hoja de aniversarios.
- Verifica que los datos estén completos (correo, fecha de ingreso, etc.) para evitar que el script lo omita.

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

## 🛠️ Sobre el trigger de aniversario

- El trigger para envío automático de correos de aniversario ya está activo en el proyecto de Apps Script.
- Solo necesitas ejecutar manualmente `TriggerAniversario.gs` si el trigger fue borrado o desactivado.
- Para revisar o reactivar el trigger:
  1. Abre el editor de Apps Script.
  2. Ve a "Relojes" o "Triggers" en el menú izquierdo.
  3. Verifica que exista un trigger para la función `enviarCorreosAniversario` (ejecución diaria).
  4. Si no existe, crea uno nuevo y selecciona la función y el tipo de disparador (por tiempo, diario).

---

## 🚀 Ventajas

- Automatización real y eficiente de la gestión de vacaciones y aniversarios.
- Archivos y correos personalizados para cada colaborador.
- Proceso reutilizable y fácil de adaptar año con año.

## 💬 Soporte

Si tienes dudas, contacta al responsable de vacaciones o revisa los comentarios en los scripts para adaptar el proceso a tus necesidades.
