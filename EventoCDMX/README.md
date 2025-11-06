# 📧 Guía para envío automatizado de invitaciones a eventos 🎉

Bienvenido/a a la carpeta de eventos "one shot" de la Fundación León XIII. Este espacio está diseñado para facilitar el envío masivo y personalizado de invitaciones a eventos únicos, optimizando el proceso para futuros organizadores. 🚀

## ❓ ¿Qué es un evento "one shot"?
Un "one shot" es un evento especial que se realiza una sola vez, sin continuidad programada. 🗓️

---

## 📂 ¿Qué contiene esta carpeta?
- **CorreoEvento.html**: Plantilla de correo adaptable para invitar a los participantes. El nombre de cada destinatario se inserta automáticamente para personalizar el mensaje. ✉️
- **TriggerEvento.gs**: Script de Google Apps Script que automatiza el envío de correos personalizados a todos los participantes registrados en la hoja de cálculo. 🤖
- **README.md**: Esta guía de uso y recomendaciones. 📄

## ⚡ ¿Cómo funciona el envío automatizado?
1. Prepara una hoja de cálculo en Google Sheets con los datos de los participantes. **Puedes usar cualquier archivo o pestaña, siempre que contenga las columnas requeridas:**
   - `Correo` (email de cada participante)
   - `Nombre corto` (nombre para personalizar el mensaje)
2. Personaliza el contenido de la plantilla `CorreoEvento.html` si lo deseas. Puedes agregar más variables si quieres personalizar aún más el mensaje (por ejemplo, fecha, lugar, etc.).
3. Ejecuta el script `TriggerEvento.gs` en Google Apps Script. El script buscará las columnas mencionadas y enviará un correo personalizado usando la plantilla.
   - Si tu hoja no se llama `2025`, puedes cambiar el nombre de la hoja en la línea:
     ```js
     var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('2025');
     ```
     Solo reemplaza `'2025'` por el nombre de tu hoja.
   - El script valida que existan las columnas necesarias. Si falta alguna, mostrará un error y no enviará los correos.
   - El asunto del correo se puede modificar en la variable `asunto` dentro del script.
4. El proceso es reutilizable para futuros eventos: solo actualiza la hoja de cálculo y la plantilla según el nuevo evento. 🔁

> ⚠️ **Importante:** Si la hoja no contiene las columnas `Correo` y `Nombre corto`, el script no podrá enviar los correos correctamente. Revisa los comentarios en el script para adaptar el proceso a tus necesidades.

---

### 🚦 Ejemplo de uso rápido
1. Crea o abre una hoja de cálculo con las columnas `Correo` y `Nombre corto`.
2. Personaliza la plantilla `CorreoEvento.html` si lo deseas. 📝
3. Abre el editor de Apps Script, pega el contenido de `TriggerEvento.gs` y ajusta el nombre de la hoja si es necesario.
4. Haz clic en "Ejecutar" para enviar los correos personalizados. 📤

---

### 👀 Ejemplo visual de la hoja de cálculo
| Correo                | Nombre corto |
|-----------------------|--------------|
| juan@email.com        | Juan         |
| maria@email.com       | María        |
| pedro@email.com       | Pedro        |


Puedes agregar más columnas si quieres personalizar otros datos (por ejemplo, `Fecha`, `Lugar`, `Mensaje`). ✨

---

### 🛠️ Personalización avanzada
**Para agregar más datos personalizados:**
1. Añade nuevas columnas en la hoja de cálculo (por ejemplo, `Fecha`, `Lugar`).
2. Modifica el script `TriggerEvento.gs` para leer esas columnas y pasarlas a la plantilla:
	```js
	var idxFecha = encabezados.indexOf('Fecha');
	var fecha = idxFecha !== -1 ? datos[i][idxFecha] : '';
	html.fecha = fecha;
	```
3. En la plantilla `CorreoEvento.html`, usa la variable:
	```html
	<?= fecha ?>
	```

**Para cambiar el asunto del correo:**
Modifica la línea:
```js
var asunto = "Invitación - Encuentro de promotores 2025 ✨";
```

**Para cambiar el nombre de la hoja:**
Modifica la línea:
```js
var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('2025');
```
Reemplaza `'2025'` por el nombre de tu hoja.

---

### ⚠️ Advertencias y buenas prácticas
- Verifica que los nombres de las columnas estén escritos exactamente igual que en el script.
- Si falta alguna columna requerida, el script mostrará un error y no enviará los correos.
- Puedes probar el envío con tu propio correo antes de hacerlo masivo.
- Revisa el log de Apps Script para ver el estado de los envíos y posibles errores.

---

## 🚀 Ventajas
- Envío masivo y personalizado en minutos.
- Reutilizable para cualquier evento futuro.
- Fácil de adaptar el mensaje y la lista de destinatarios.

## 📝 Recomendaciones
- Revisa y personaliza los textos antes de enviar.
- Verifica que los datos en la hoja de cálculo estén completos y correctos.
- Consulta la documentación interna o los comentarios en el script para dudas sobre la automatización.