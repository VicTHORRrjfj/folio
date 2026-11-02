# Folio — Word, Excel y PowerPoint sin conexión

App web instalable (PWA) para abrir, editar y guardar `.docx`, `.xlsx` y `.pptx` en la tablet y en el PC, sin internet y sin anuncios.

## Por qué no deforma los archivos

Folio **no reconstruye** tu documento. Edita el texto y los valores directamente dentro del XML del archivo original y deja todo lo demás intacto: estilos, imágenes, gráficos, tablas, encabezados, temas y formatos condicionales vuelven a Office exactamente como estaban.

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

- **Word:** escribe directamente sobre la página. Selecciona texto y usa la barra para negrita, cursiva, subrayado, color, resaltado, tamaño y alineación. *Imagen* inserta una foto o imagen bajo el párrafo donde estás; tócala para achicarla, agrandarla o eliminarla. *Tabla* agrega o quita filas en la tabla donde está el cursor.
- **Excel:** toca una celda para seleccionarla y vuelve a tocar para editar (o escribe en la barra `fx`). Para un rango, activa *Rango* y toca la última celda, o toca el encabezado de una fila/columna (en el PC también sirve Mayús+clic). Con la barra das negrita, color de letra, relleno, tamaño, alineación, ajuste de texto y formato de número (moneda, %, fecha). *Filas* y *Columnas* insertan o eliminan, y Folio ajusta las fórmulas, celdas combinadas, validaciones, formato condicional, gráficos e imágenes igual que Excel. Las fórmulas se recalculan al instante.
- **PowerPoint:** toca un cuadro de texto para editarlo y usa la barra para dar formato. *Diapositiva* crea una nueva (con el mismo diseño de la actual), duplica, mueve o elimina. *Imagen* agrega una imagen a la diapositiva actual; arrástrala para moverla y usa el círculo de la esquina para cambiar su tamaño. ▶ inicia la presentación.
- **Guardar:** en la tablet guarda en *Descargas*; en el PC sobrescribe el archivo. *Compartir* (menú ⋮) lo envía por WhatsApp, Drive, correo, etc.
- Los últimos 20 archivos quedan en *Recientes*, disponibles sin conexión.
- **Compartir a Folio (tablet):** con la app instalada, desde WhatsApp, Drive, Gmail o *Mis archivos* toca *Compartir* → **Folio** y el documento se abre directo.
- **Copiar y pegar con formato:** al pegar texto copiado de Word o PowerPoint se mantienen negrita, cursiva, subrayado, color, tamaño, fuente y alineación. Al pegar celdas copiadas de Excel o Google Sheets se mantienen valores, formato de número (moneda, %, fecha), negrita, colores, relleno y alineación. Si copias y pegas dentro de Folio, las fórmulas se ajustan a su nueva posición. En el PC también puedes copiar un archivo en el explorador y pegarlo (Ctrl+V) en la pantalla de inicio para abrirlo.
- **Atajos (PC):** Ctrl+N o Ctrl+B negrita, Ctrl+K o Ctrl+I cursiva, Ctrl+U subrayado, Ctrl+S guardar.

## Límites de esta versión

- Word: encabezados, pies y cuadros de texto son de solo lectura. Las páginas se cortan solo en saltos de página manuales. Las imágenes se insertan en línea (no flotantes).
- Excel: los gráficos se muestran como un recuadro (se conservan). No se pueden insertar columnas dentro de una tabla con formato de Excel. Deshacer (Ctrl+Z) cubre los valores escritos, no los cambios de formato o de filas.
- PowerPoint: gráficos y SmartArt se muestran como un recuadro (se conservan); las animaciones no se reproducen. Solo las imágenes se mueven con el dedo; las demás formas mantienen su posición.
- Al pegar en Word o PowerPoint se pega el texto con su formato; las imágenes y tablas copiadas se pegan como texto (usa *Imagen* y *Tabla* para eso).
- Formatos antiguos `.doc/.xls/.ppt` no se abren: guárdalos antes como `.docx/.xlsx/.pptx`.

## Licencias de las librerías incluidas

- JSZip (MIT), docx-preview (Apache-2.0), SSF de SheetJS (Apache-2.0).
- HyperFormula (motor de fórmulas de Excel) se usa con licencia **GPLv3**: bien para uso personal. Si algún día vendes Folio, necesitas la licencia comercial de HyperFormula o reemplazar ese módulo.
