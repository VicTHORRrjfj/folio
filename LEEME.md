# Folio — Word, Excel y PowerPoint sin conexión

App web instalable (PWA) para abrir, editar y guardar `.docx`, `.xlsx` y `.pptx` en la tablet y en el PC, sin internet y sin anuncios.

## Por qué no deforma los archivos

Folio **no reconstruye** tu documento. Edita el texto y los valores directamente dentro del XML del archivo original y deja todo lo demás intacto: estilos, imágenes, gráficos, tablas, temas y todo lo que no tocaste vuelven a Office exactamente como estaban.

## Instalar en la tablet (Samsung Galaxy Tab S10)

La app necesita publicarse en una dirección `https://` una sola vez; después funciona sin conexión.

**Con GitHub Pages (gratis):**

1. Crea un repositorio nuevo en tu cuenta de GitHub, por ejemplo `folio`.
2. Sube **el contenido** de esta carpeta (`index.html`, `app.css`, `sw.js`, `manifest.webmanifest` y las carpetas `js`, `vendor`, `assets`, `icons`).
3. En el repositorio: *Settings → Pages → Branch: main / (root) → Save*.
4. En un par de minutos queda en `https://TU-USUARIO.github.io/folio/`.
5. Ábrela en Chrome de la tablet → menú ⋮ → **Agregar a pantalla principal / Instalar app**.
6. Ábrela una vez con internet; desde ahí funciona en modo avión.

**En el PC:** abre la misma dirección en Chrome o Edge y usa el ícono **Instalar** de la barra de direcciones. Instalada, puedes abrir `.docx/.xlsx/.pptx` con Folio desde el explorador de archivos y "Guardar" sobrescribe el archivo original.

## Cómo se usa

La barra superior tiene **Deshacer / Rehacer** (Ctrl+Z / Ctrl+Y), **Buscar y reemplazar** (Ctrl+B o Ctrl+F; Ctrl+L o Ctrl+H para reemplazar), zoom y **Guardar** (Ctrl+G o Ctrl+S). En el menú ⋮ están *Imprimir o guardar PDF* (Ctrl+P), *Guardar una copia* y *Compartir*. Debajo va la cinta con pestañas, como en Office.

- **Word** — *Inicio*: fuente, tamaño, negrita (Ctrl+N), cursiva (Ctrl+K), subrayado (Ctrl+U), color, resaltado, alineación, viñetas, numeración, sangría, interlineado y estilos (Título, Título 1–3, Subtítulo, Cita…). *Insertar*: imagen, tabla nueva (elige filas × columnas), salto de página, agregar/quitar filas y columnas, eliminar tabla, **encabezado**, **pie de página** y **número de página** («Página 1 de N»).
  - El documento se **pagina solo** como en Word: cuando el texto llena una hoja pasa a la siguiente, y el encabezado y el pie se repiten en cada página con su número. Abajo se ve la página actual y el conteo de palabras.
  - El encabezado y el pie se editan tocándolos directamente. En *Encabezado* puedes activar **primera página diferente** y **pares e impares diferentes**, como en Word.
  - Las **tablas largas** se parten fila por fila entre páginas y los párrafos largos se parten por líneas (dejando al menos 2 líneas a cada lado).
  - Toca una imagen → **Ajuste**: en línea, cuadrado a la izquierda o derecha (el texto la rodea), arriba y abajo, o **delante del texto**. Las imágenes flotantes se **arrastran**: delante del texto quedan donde las sueltes; si el texto las rodea, cambian de párrafo y de lado.
  - **Notas al pie**: *Insertar → Nota al pie* pone la llamada en el cursor; las notas se numeran solas, van al pie de su página y se editan ahí mismo.
- **Excel** — *Inicio*: fuente, tamaño, colores, relleno, bordes, alineación, ajustar texto, combinar celdas, formato de número, Autosuma y **formato condicional** (mayor que, menor que, entre, igual a, texto que contiene, duplicados, 10 superiores/inferiores, sobre/bajo el promedio, barras de datos, escalas de color e iconos; también *Administrar reglas…* para cambiar el rango, la prioridad, borrar reglas o crear una con fórmula, y *Borrar reglas*). El formato condicional que ya traen tus archivos se dibuja igual que en Excel. *Insertar*: imagen, **gráficos** (columnas, barras, líneas, circular; se actualizan solos al cambiar los datos), filas, columnas (también dentro de tablas con formato de Excel) y hojas. *Datos*: ordenar A→Z / Z→A, **filtros** con ▾ en los encabezados (lista de valores o *Filtros de número / de texto*: mayor que, entre, contiene, comienza por, sobre el promedio, 10 mejores…) y rellenar hacia abajo/derecha (Ctrl+D / Ctrl+R). *Vista*: inmovilizar paneles, ancho de columna, alto de fila, autoajustar y formato regional.
  - Las fórmulas se escriben **en español con punto y coma** (`=SUMA(A1:A5; 10)`, `=SI(B2>0; "Sí"; "No")`), con sugerencias mientras escribes; dentro del archivo se guardan como Excel las espera.
  - El **cuadrito verde** de la celda seleccionada se arrastra para rellenar series: 1, 2, 3…; lunes, martes…; enero, febrero…; Producto 1, Producto 2…; y fórmulas que se ajustan solas.
  - Al seleccionar un rango, abajo aparecen el **promedio, el recuento y la suma**.
  - Hojas: **+** agrega; clic derecho (o mantener presionado) en una pestaña para cambiarle el nombre, duplicarla o eliminarla; doble clic para renombrar.
  - Arrastra el borde de un encabezado de columna o fila para cambiar su tamaño; doble clic en el borde de la columna la autoajusta.
- **PowerPoint** — A la izquierda están las **miniaturas** (clic para ir; clic derecho para duplicar, mover o eliminar). *Inicio*: nueva diapositiva (eliges el diseño), fuente, formato, viñetas, numeración, sangría e interlineado. *Insertar*: cuadro de texto, formas, imagen y tabla. *Diseño*: fondo (de una o de todas), duplicar, subir, bajar, ocultar y eliminar diapositiva.
  - Toca un cuadro de texto para escribir; aparece **⠿ Mover** para arrastrarlo, **×** para borrarlo y un círculo en la esquina para cambiar su tamaño. Imágenes, formas, tablas, gráficos y **grupos** se mueven arrastrándolos; el círculo de arriba **gira** el objeto (con Mayús salta de 15° en 15°) y los objetos girados también cambian de tamaño. En un grupo, el primer toque lo selecciona entero y el segundo, el objeto de adentro.
  - Los **SmartArt** se ven como en PowerPoint y su **texto se edita** tocándolo. La presentación reproduce las **animaciones** de entrada, salida, énfasis y **trayectorias**, clic a clic.
  - ▶ inicia la presentación (toca a la derecha para avanzar, a la izquierda para retroceder).
- **Guardado automático**: mientras editas, Folio guarda un borrador cada pocos segundos. Si cierras sin guardar, el archivo aparece en *Recientes* como «Cambios sin guardar» y al abrirlo se recupera tal cual. Para tener el archivo .docx/.xlsx/.pptx, toca **Guardar**.
- **Plantillas** en el inicio: factura con IVA, cotización, presupuesto mensual, control de inventario, informe y presentación de proyecto.
- **Compartir a Folio (tablet):** desde WhatsApp, Drive, Gmail o *Mis archivos* toca *Compartir* → **Folio**.
- **Copiar y pegar con formato** desde Word, PowerPoint, Excel y Google Sheets (negritas, colores, tamaños, formatos de número).

## Archivos antiguos

- `.xls`, `.ods` y `.csv` se abren convirtiéndolos a `.xlsx` (valores, fórmulas y formatos de número; colores y bordes del archivo antiguo pueden perderse). *Guardar* crea el `.xlsx`. La conversión corre aislada del resto de la app.
- `.doc` y `.ppt` (Office 97–2003) no se pueden leer sin Office: ábrelos una vez en Office, Google Drive o LibreOffice y guárdalos como `.docx` / `.pptx`.

## Límites de esta versión

- `.doc` y `.ppt` antiguos (ver arriba).
- SmartArt: se edita el texto de las cajas; agregar o quitar cajas se hace en PowerPoint.

## Licencias de las librerías incluidas

- JSZip (MIT), docx-preview (Apache-2.0), SSF y SheetJS 0.18.5 (Apache-2.0; solo para convertir .xls/.ods/.csv). Esa versión tiene fallos conocidos con archivos manipulados; por eso corre en un hilo aparte, sin acceso a la app, y se corta si tarda más de 30 s.
- HyperFormula (motor de fórmulas de Excel) se usa con licencia **GPLv3**: bien para uso personal. Si algún día vendes Folio, necesitas la licencia comercial de HyperFormula o reemplazar ese módulo.
