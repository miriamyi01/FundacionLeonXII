# 📋 Servicio Social - Fundación León XIII

Este módulo automatiza el aviso de término de servicio social y la gestión de eventos en el calendario institucional. Incluye una plantilla visual para el correo y un script para el procesamiento de registros desde Google Sheets.

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