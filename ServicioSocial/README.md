
# 📋 Servicio Social - Fundación León XIII

Este módulo automatiza el aviso de término de servicio social y la gestión de eventos en el calendario institucional. Incluye:

- Un script que agrega automáticamente eventos al calendario institucional solo si no existen para ese registro y fecha.
- Un correo de aviso de término que utiliza una plantilla HTML moderna con colores institucionales y logotipo.
- Formato de fecha en español y mensaje personalizado para cada usuario.
- El mensaje de correo y el evento de calendario incluyen información adicional: nombre, universidad, licenciatura, tipo de programa y fechas relevantes.
- Variables dinámicas en la plantilla: `nombre`, `universidad`, `fechaFormateada`, `diffDays`.
- Mejoras en la validación de fechas y control de duplicados en eventos.

---

## Estructura requerida de los archivos para extraer enlaces

Para que el script pueda extraer correctamente los enlaces y datos de los registros, la hoja de cálculo debe tener la siguiente estructura (columnas mínimas recomendadas):

- **Nombre** (columna A)
- **Universidad** (columna E)
- **Licenciatura/Carrera** (columna G)
- **Tipo de programa** (columna H)
- **Fecha de inicio** (columna L)
- **Fecha de término** (columna M)

Asegúrate de que las fechas estén en formato de fecha y no como texto. Si agregas nuevas columnas, actualiza los índices en el script para que correspondan a la nueva estructura.

---

## Archivos incluidos

- **AvisoTerminoMensaje.html**
  - Plantilla HTML para el aviso de término de servicio social.
  - Incluye diseño moderno, colores institucionales y logotipo.
  - Variables dinámicas: `nombre`, `universidad`, `fechaFormateada`, `diffDays`.

- **AvisoTermino.gs**
  - Script de Google Apps Script para procesar todos los registros de la hoja de cálculo.
  - Envía correos automáticos cuando faltan 30, 15 o 7 días para el término del servicio social.
  - Crea eventos en el calendario institucional si no existen para cada registro.
  - Formatea fechas en español y personaliza el mensaje.

---

## ¿Cómo funciona?

1. El script lee todos los registros de la hoja activa en Google Sheets (excepto el encabezado).
2. Para cada registro, calcula los días restantes hasta la fecha de término.
3. Si faltan 30, 15 o 7 días, envía un correo personalizado usando la plantilla HTML.
4. Si no existe un evento en el calendario para ese registro y fecha, lo crea automáticamente.

---

## Personalización

- **Correo destinatario y calendario:**
  - Modifica las variables `email` y `calendarId` en el script para cambiar el destinatario y el calendario.
- **Variables de la plantilla:**
  - Puedes agregar más variables en el script y usarlas en la plantilla HTML.
- **Colores y logotipo:**
  - Los colores y el logotipo pueden ajustarse en la plantilla HTML para adaptarse a la identidad visual de la organización.

---

## Requisitos

- Google Apps Script vinculado a una hoja de cálculo de Google Sheets.
- Acceso al correo y calendario institucional.
- Configuración de permisos para enviar correos y crear eventos en el calendario.

---

## Ejemplo de uso

1. Actualiza la hoja de cálculo con los registros de servicio social.
2. Ejecuta el script `avisoTerminoServicioSocial()` en el editor de Apps Script.
3. Los correos y eventos se generarán automáticamente según las fechas configuradas.

---

## Buenas prácticas

- Verifica que las fechas estén en formato de fecha y no texto.
- Revisa el log de Apps Script para monitorear envíos y eventos creados.
- Personaliza el mensaje y los colores según la comunicación institucional.