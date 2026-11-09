# Cuaderno de Campo

App para el móvil (PWA) para llevar una explotación de cereal, leguminosas y cultivos alternativos:

- **Fincas**: superficie, municipio, referencia SIGPAC y cultivo de cada campaña, con el historial de rotaciones.
- **Cuaderno de campo**: labrado, siembra, abonado, tratamientos (producto, dosis, nº de registro, plaga), cosecha… Si una labor lleva coste, se apunta sola como gasto.
- **Cuentas**: gastos e ingresos por finca o generales (seguro, gasóleo, PAC). En ventas de grano calcula el importe a partir de los kilos y el precio.
- **Cosechas**: kilos por finca (cada remolque o el total) y rendimiento en kg/ha.
- **Resumen**: superficie, kg/ha y margen €/ha por cultivo y por finca, y comparativa entre campañas.

Funciona sin cobertura. Los datos se guardan en el propio teléfono. Desde **Ajustes → Copia de seguridad** se crea un archivo `.json` que se puede mandar a Google Drive, WhatsApp o al correo, e importar en otro móvil.

Las campañas van de septiembre a agosto (por ejemplo, 2026/27).

## Instalar en Android

1. Abre la dirección donde esté publicada la app en **Chrome**.
2. Menú ⋮ → **Instalar aplicación** (o «Añadir a pantalla de inicio»).
3. Aparecerá el icono «Campo» y se abrirá a pantalla completa, también sin conexión.

## Publicarla gratis con GitHub Pages

1. Sube esta carpeta a un repositorio de GitHub.
2. En el repositorio: **Settings → Pages → Source: Deploy from a branch → `main` / root**.
3. A los pocos minutos estará en `https://TU-USUARIO.github.io/NOMBRE-DEL-REPO/`.

## Ficheros

| Fichero | Qué es |
|---|---|
| `index.html` | Estructura de la pantalla |
| `styles.css` | Diseño (modo claro y oscuro) |
| `app.js` | Toda la lógica y los datos |
| `sw.js` | Guarda la app para usarla sin cobertura |
| `manifest.webmanifest`, `icons/` | Nombre, icono y colores al instalarla |

Al cambiar `app.js` o `styles.css`, sube el número de `CACHE` en `sw.js` para que los móviles descarguen la versión nueva.
