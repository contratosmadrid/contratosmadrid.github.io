# Mapa de contratación municipal · Ayuntamiento de Madrid

Excel maestro → GitHub → validación y conversión automática → GitHub Pages (mapa público).

## Estructura

| Carpeta / archivo | Qué es | ¿Se toca? |
|---|---|---|
| `datos/Excel_Contratos_Madrid_Definitivo.xlsx` | **Fuente única de datos** (hoja «Datos») | **Sí: es lo único que se actualiza** |
| `datos/particularidades_heredadas.json` | Textos de «Particularidades» que mostraba la web v2 cuando la columna *Excepciones / limites* está vacía. Si se rellena la celda en el Excel, manda el Excel | No |
| `herramientas/excel_a_datos.py` | Valida el Excel y genera los datos de la web | Solo si cambian columnas |
| `web/index.html` | El mapa (diseño, filtros, fichas). No contiene datos | Solo para cambios de diseño |
| `web/data/contratos.js` / `.json` | Datos generados automáticamente. **No editar a mano** | No |
| `web/data/informe.md` | Resumen de la última conversión: cifras, cambios y avisos | Para consultar |
| `.github/workflows/publicar.yml` | Proceso automático de publicación | No |

## Puesta en marcha (una sola vez)

1. Crea una cuenta en github.com y un repositorio **público** nuevo (por ejemplo `mapa-contratos-madrid`).
2. Sube todo el contenido de esta carpeta (en la web: *Add file → Upload files*, arrastrando las carpetas; o con GitHub Desktop).
   La carpeta `.github` es oculta: si arrastras desde el explorador, comprueba que se ha subido.
3. En el repositorio: **Settings → Pages → Build and deployment → Source: «GitHub Actions»**.
4. Ve a **Actions → Publicar mapa → Run workflow**. En 1–2 minutos aparece la dirección pública
   (`https://TU-USUARIO.github.io/mapa-contratos-madrid/`).

## Actualizar un contrato

1. Abre el Excel, modifica los datos en la hoja **Datos** y guarda (sin cambiar el nombre de las columnas).
2. Sube el Excel al repositorio **sustituyendo** el anterior, en `datos/` (mismo nombre o distinto: debe quedar **un solo** `.xlsx` en esa carpeta).
3. GitHub ejecuta solo «Publicar mapa». En 1–2 minutos la web muestra los cambios.

> Guardar el Excel en tu ordenador **no** actualiza nada. Hay que subirlo a GitHub (web, GitHub Desktop o una sincronización configurada aparte).

## Qué ocurre si…

| Caso | Resultado |
|---|---|
| Modificas adjudicatario, fechas, estado, fuentes… | Se publica; `informe.md` lista qué campos cambiaron en cada REL |
| Añades una fila con un REL nuevo (p. ej. `REL-0493`) | Se publica como relación nueva. Si es otro expediente, sube el recuento |
| Un contrato deja de estar vigente | **No borres la fila**: cambia «ESTADO DE VIGENCIA» y la fecha de fin. El mapa lo clasifica en «Fin registrado ya pasado» |
| Borras una fila | Desaparece de la web y queda anotado en el informe como «Eliminada» |
| REL repetido, REL vacío, distrito mal escrito, fecha ilegible, columna renombrada | **No se publica nada**; la versión anterior sigue en línea. El error aparece en Actions (icono rojo) |

## Comprobar una actualización

- **Actions**: icono verde = publicado; rojo = abre la ejecución y lee el paso «1 · Validar el Excel».
- El resumen de cada ejecución muestra expedientes, relaciones, cambios y avisos.
- En la web, la tarjeta informativa indica la fecha de generación y el Excel de origen.
- Cifras de control con el Excel del 23/10/2026: 67 expedientes · 492 relaciones · Retiro 20 · «Casa de Campo» 2 · Acciona 20 · Residuos 127 · grupo ETRA 27 · ámbito municipal 46. Si el Excel cambia, estas cifras pueden cambiar con razón.

## Recursos externos

- Límites de distritos: se descargan en la publicación y se guardan en `web/data/madrid-districts.geojson` (el mapa ya no depende de enlaces externos para la geometría).
- Logo: el proceso intenta guardar una copia del escudo publicado en madrid.es en `web/img/`. Si no puede, se muestra el escudo esquemático. Uso sujeto a las normas de identidad del Ayuntamiento.
- Mapa base de la vista geográfica: © OpenStreetMap · © CARTO (se carga al vuelo).
- Fotografías de distritos: Wikimedia Commons, con enlace a autoría y licencia.
