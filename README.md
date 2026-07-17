# Academia PT — Oposición Pedagogía Terapéutica (Comunidad de Madrid)

Dos páginas independientes conectadas a un mismo Google Sheet mediante Google Apps Script:

- **profesor.html** — gestión de alumnos, calendario, revisión de prácticos (texto y foto) con feedback manual, temario y cultura general.
- **alumno.html** — accesible mediante enlace personalizado (`alumno.html?id=...`), muestra el calendario propio, permite enviar prácticos (texto o foto) y ver el feedback, y marcar el progreso del temario.

## 1. Crear el Google Sheet y el backend

1. Crea un Google Sheet nuevo y vacío (las pestañas se crean solas la primera vez que se use).
2. Copia su ID: en la URL `https://docs.google.com/spreadsheets/d/AQUI_ESTA_EL_ID/edit`.
3. En el propio Sheet: **Extensiones > Apps Script**.
4. Borra el contenido de `Code.gs` que aparece por defecto y pega el contenido de `Code.gs` de este repositorio.
5. Sustituye `PON_AQUI_TU_SHEET_ID` por el ID que copiaste en el paso 2.
6. **Implementar > Nueva implementación**:
   - Tipo: **Aplicación web**
   - Ejecutar como: **Yo**
   - Quién tiene acceso: **Cualquier usuario**
7. Autoriza los permisos (Sheets + Drive) cuando te lo pida.
8. Copia la URL que termina en `/exec`. Esa es tu `API_URL`.

## 2. Configurar las páginas

En **profesor.html** y en **alumno.html**, sustituye:

```js
const API_URL = 'PON_AQUI_TU_URL_DE_APPS_SCRIPT';
```

por la URL `/exec` que copiaste.

En **profesor.html**, cambia también la clave de acceso si quieres:

```js
const COACH_KEY = 'cgr-pt-2026';
```

Entrarás en tu panel como `profesor.html?key=TU_CLAVE` o escribiéndola en la pantalla de bloqueo.

## 3. Publicar en GitHub Pages

Una vez el repositorio esté en GitHub: **Settings > Pages > Deploy from branch > main / (root)**.
Tus páginas quedarán en:

- `https://tuusuario.github.io/tu-repo/profesor.html`
- `https://tuusuario.github.io/tu-repo/alumno.html`

## 4. Enlaces de alumno

En **profesor.html > Ajustes**, guarda la URL base de tu `alumno.html` publicada (por ejemplo `https://tuusuario.github.io/tu-repo/alumno.html`). A partir de ahí, en la pestaña **Alumnos** verás el enlace personalizado listo para copiar y enviar a cada alumno.

## Notas

- Las fotos de los prácticos se guardan en una carpeta de Google Drive llamada `PT_Practicos_Fotos`, dentro de la cuenta donde implementaste el Apps Script.
- No hay corrección automática por IA: el profesor revisa el texto o la foto y escribe el feedback manualmente en `profesor.html`, que el alumno ve al instante en `alumno.html`.
- Si en el futuro quieres añadir corrección automática, haría falta una clave de API de Anthropic de pago, integrada en el propio `Code.gs` (Apps Script puede llamar a la API de Anthropic con `UrlFetchApp`, incluida la lectura de fotos).
