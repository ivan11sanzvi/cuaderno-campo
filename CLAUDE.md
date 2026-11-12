# Cuaderno de Campo

App móvil (PWA) de uso personal para una explotación agrícola de secano: cereal (trigo, cebada), leguminosas (titarro, lenteja) y cultivos alternativos (anís dulce, cilantro, esparceta). El usuario la usa en un Android, instalada desde Chrome. Publicada con GitHub Pages desde la rama `main`.

Habla con el usuario en español, con lenguaje sencillo y sin tecnicismos: es agricultor, no programador.

## Estructura
- `index.html`: estructura. `styles.css`: diseño (claro y oscuro). `app.js`: toda la lógica, en un único IIFE.
- `sw.js`: service worker. **Sube `CACHE` (cuaderno-campo-vN) en cada cambio** de `app.js` o `styles.css`, si no los móviles no se actualizan.
- `vendor/`: Leaflet 1.9.4 y SheetJS 0.18.5, locales para que funcione sin cobertura.
- Sin build ni dependencias: se sirve tal cual (`python3 -m http.server`).

## Datos
- Todo en `localStorage` (clave `cuadernoCampo.v1`) del propio teléfono: fincas, siembras (cultivo por finca y campaña), labores, movimientos (gastos e ingresos), cosechas.
- Las campañas van de septiembre a agosto («2026/27»).
- Copias de seguridad en JSON (compartir o descargar) e importación.
- **El repositorio es público: nunca subas datos del usuario** (Excel, copias JSON). `.gitignore` excluye hojas de cálculo.

## SIGPAC y mapa
- Recintos: `https://sigpac-hubcloud.es/ogcapi/collections/recintos/items?f=json&provincia=&municipio=&agregado=&zona=&poligono=&parcela=[&recinto=]`, con geometrías en lon/lat.
- Recinto en un punto: `/servicioconsultassigpac/query/recinfobypoint/4258/{lon}/{lat}.json`. Devuelve la superficie, `uso_sigpac` y el contorno en WKT.
- Linderos: WMS `https://sigpac-hubcloud.es/wms`, capa `recinto`.
- Ortofoto: `https://tms-pnoa-ma.idee.es/1.0.0/pnoa-ma/{z}/{x}/{y}.jpeg` (TMS: `tms: true` en Leaflet).
- Todos admiten CORS. La capa de SIGPAC es la vigente, así que referencias de años anteriores pueden haber cambiado (recintos unidos, concentraciones parcelarias).

## Parcelas y fincas agrupadas
- Cada registro de `db.fincas` es una **parcela**, con `titular` y `grupo` (nombre de la finca agrupada). En la interfaz se habla de «parcela» y de «finca agrupada».
- Las siembras guardan `laboreo` (C = siembra directa con complemento, T = tradicional, D = siembra directa) y `estiercol` (true si se echó «basura»).
- Las labores, gastos y cosechas hechos en varias parcelas a la vez comparten `lote`. Los importes, kilos y costes se reparten por superficie, en las listas salen en una sola línea y se editan juntos.
- «Lindan» se calcula con los contornos SIGPAC guardados en `f.geo` (vértice a menos de 5 m del borde de la otra parcela).

## Almacén, albaranes y facturas
- `db.productos` es el catálogo de productos: `{nombre, tipo: Fitosanitario|Abono|Semilla|Otro, unidad: L|kg|t|ud, regId, registro}`.
- `db.documentos` guarda los albaranes y facturas: `{tipo: albaran|factura, titular, fecha, proveedor, numero, lineas: [{productoId, cantidad, precio, importe}], archivos: [ids], albaranes (solo facturas), base, iva, total}`.
- `db.ajustes` guarda las regularizaciones de existencias.
- El inventario es **por titular**: entradas de albaranes − `labor.consumo` + ajustes. En `movimientosProducto(pid, tit)`, `null` significa todos los titulares y `''` sin titular.
- Las fotos y los PDF van en IndexedDB (`cuadernoCampoArchivos`, almacén `archivos`). Las fotos se reducen a 1800 px. La copia de seguridad JSON los incluye en base64 (`archivos`).
- Una factura cuadra si la base (sin IVA) coincide con la suma de sus albaranes (±0,50 €) y las cantidades por producto también.

## Fitosanitarios y normativa
- `fitos.json` se genera con `herramientas/actualizar_fitos.py`. Lee el Registro del MAPA (`servicio.mapa.gob.es/regfiweb`): la lista de productos de cada cultivo y la tabla «Usos y Dosis Autorizados» de cada ficha PDF.
- Los cultivos de la app se comparan con los del Registro mediante `CULTIVO_REG`. El titarro no está en el Registro: se compara con almorta y leguminosas de grano.
- En una labor de Tratamiento el producto sale **solo del almacén** del titular. Se comprueban la autorización en el cultivo, la dosis máxima, el estado y la caducidad, y si hay existencias suficientes. Si hay problemas se pide confirmar, y la labor guarda `avisos` y sale marcada «⚠ Revisar normativa».
- Conviene regenerar `fitos.json` de vez en cuando, porque el Registro cambia cada semana.

## Asistente de voz
- Es el botón amarillo del micrófono, abajo a la izquierda (`hojaVoz`). El dictado usa Web Speech API en `es-ES`, que necesita cobertura; sin cobertura sirve el micrófono del teclado.
- Funciona **sin IA**: `interpretar(texto)` aplica `normVoz` (números dichos con palabras → cifras), palabras clave (`VERBOS`) y búsqueda aproximada de nombres (`buscarNombres`, Levenshtein) de fincas, titulares, cultivos y productos.
- Entiende labores, cosechas, albaranes, facturas, fincas nuevas, gastos, ingresos y semilla propia.
- Fincas: se aplica la finca agrupada entera, salvo lo que vaya tras «excepto / menos / salvo» (parcelas por número, titulares u otras fincas). También entiende «las de trigo de Tomás».
- Siempre enseña una tarjeta de confirmación. **Guardar** usa `registrarLabor` / `registrarCosecha` / `registrarMov`, con las mismas comprobaciones de normativa que el formulario. **Corregir** abre el formulario relleno.
- Siembra: asigna el cultivo (`cultivoSiembra`) y descuenta la semilla «<Cultivo> para sembrar» del almacén de cada titular. La semilla propia entra como ajuste.
- Historial: `empujar()` espera si hay un `history.go` pendiente (`pendientes`). Sin eso, al cerrar una ventana y abrir otra seguidas, la app se salía al pulsar atrás.

## Excel
- La importación reconoce títulos parecidos («Paraje», «Pol.», «Sup. (ha)», «Propietario», «Superficie declarable PAC», «Término»…) y admite varias campañas en columnas: bloques CULTIVO / BASURA / LABOREO con el año encima.
- Las filas se juntan por parcela SIGPAC. Las parcelas con el mismo nombre en el mismo pueblo forman una finca agrupada, salvo los nombres de la lista «no agrupar» (por defecto «Cambio»).
- Si falta la provincia se elige por pueblo y se recuerda en `localStorage` (`.provPueblo`). Ojo: «Espinosa» (municipio 70) del usuario es Espinosa de Cerrato, en **Palencia (34)**.
- Las fincas que ya existen se emparejan por parcela SIGPAC y se actualizan sin perder lo apuntado.
- Columnas de la plantilla: Nombre, Finca agrupada, Titular, Superficie (ha), Municipio, Provincia, Cód. municipio, Agregado, Zona, Polígono, Parcela, Recinto, Otras parcelas, Cultivo, Variedad, Laboreo, Basura, Campaña, Notas.

## Pruebas
Con Playwright, emulando un Pixel 7. Para llegar a SIGPAC desde Chromium hay que usar `proxy: { server: process.env.HTTPS_PROXY, bypass: 'localhost,127.0.0.1' }`.
