# Reporte Vs Anual

App web para **evaluar las ventas de todos los años** (por defecto desde 2022) a partir de archivos de Excel o CSV. Subes los archivos y la app detecta las columnas sola. El corte de evaluación se toma siempre de **la fecha más reciente que haya en los archivos**: el último año, mes y semana con venta.

Publicada en: https://reporte-vs-anual.onrender.com

## Lectura de archivos

- **Sube varios a la vez**: `.xlsx`, `.xls`, `.xlsm`, `.xlsb` y `.csv`/`.txt`. Puede ser un archivo por año o uno solo con todos.
- Los archivos se leen en el navegador. A la base solo van las ventas resumidas.
- **Detecta las columnas automáticamente**: Venta, Fecha (o Año + Mes, o Año + Semana), Semana, Cadena, Tienda/PDV, Marca y Unidades. Si alguna sale mal, la cambias en la lista y la app recuerda tu elección.
- **Lee todas las hojas del libro**: si el archivo trae una hoja por mes (ENERO, FEBRERO… DICIEMBRE), las suma todas. Las hojas sin datos de venta utilizables, como un RESUMEN, se omiten. También puedes elegir una sola hoja.
- Si una hoja no trae fecha, toma **el mes del nombre de la hoja** y **el año de la columna Año o del nombre del archivo** (por ejemplo, "ANUAL 2024.xlsb").
- Antes de guardar muestra **los meses detectados por año**, y avisa si algún año trae un solo mes.
- **Encuentra la fila de encabezados** aunque haya títulos arriba.
- También lee reportes en **formato ancho** (los meses como columnas). Las filas de TOTAL se omiten.
- **CSV**: detecta el separador y la codificación, y acepta los formatos `1,234.56` y `1.234,56`. Las fechas se leen como día/mes/año.

## Corte de evaluación

Se calcula solo, a partir del dato más reciente en los archivos. Ejemplo: último dato el 21/11/2026 → corte **2026 · noviembre · semana 47**. Con ese corte, cada año se mide de cuatro formas:

| Medida | Qué suma |
|---|---|
| **Acumulado al corte** | Desde la semana 1 hasta la semana del corte (o de enero al mes del corte, si el archivo no trae fecha diaria). Es la comparación justa entre años |
| **Año completo** | Todo el año. En el año en curso llega solo hasta el corte |
| **Mes de corte** | Solo el mes del corte, en cada año |
| **Semana de corte** | Solo la semana del corte, en cada año |

Las semanas van de lunes a domingo; la semana 1 es la que contiene el 1 de enero.

## Pestañas

- **Resumen** (se abre por defecto): el dashboard inicial, con todo de un vistazo. Cada bloque lleva un titular que se escribe solo con la conclusión, y casi todo es clicable para entrar al detalle en la Comparativa ya filtrada.
  - Titular del acumulado al corte, mini barras por año e indicadores: acumulado, mes de corte (parejo con el año anterior), semana de corte, participación Paraíso y tiendas con venta.
  - **Mes a mes** (año actual frente al anterior; el mes en curso va punteado) y **cómo se reparte la venta** por cadena (barras al 100 % de los dos años, con venta, Var % y puntos de participación).
  - **El año día por día**: un calendario de calor donde cada día se compara con el mismo día de la semana del año anterior.
  - **Tiendas que más crecen** y **que más caen**, lado a lado.
  - **Paraíso frente a la competencia** (dona) y **últimas semanas** (Var % de las últimas 8 semanas).

- **Comparativa**: igual que la tabla dinámica de Excel "COMPARATIVAS MES ACTUAL". Tiene segmentadores de **Mes**, **Semanas** y **Día** (clic para elegir uno; Ctrl + clic o "Selección múltiple" para varios) y tablas de **Soles** y **Unidades** del año actual frente al anterior, por cadena, por marca o por Paraíso vs Competencia, con su Var %. Por defecto muestra el mes de corte hasta el día de corte. Debajo, **por semana**: tablas de Soles y Unidades con columnas Sem-N × año, más su Var % y el total, y un gráfico de barras por semana para los últimos tres años, con la Var % del año actual frente al anterior sobre cada semana. Usa el detalle diario guardado en la tabla `ventas_dia`.

- **Por año**: indicadores del año en curso frente al anterior (acumulado, mes de corte y semana de corte), gráficos por año y una tabla con todos los años, cada uno con su variación contra el año anterior.
- Filtros: **desde el año**, **cadena** y **marca**.
- **Descargar Excel**: un libro con las tablas de la Comparativa (soles, unidades y por semana) y de Por año. Cada tabla también tiene **Copiar tabla** para pegarla en Excel.

## Comparativa por tienda y por coordinador

- En la **Comparativa**, "Ver por" ofrece **Cadena, Tienda, Coordinador, Marca y Paraíso vs Competencia**. Tienda y coordinador se calculan en el servidor con el detalle diario por tienda (tabla `ventas_tienda_dia`), así que respetan los segmentadores de Mes, Semana y Día.
- La app lee la columna **COORDINADOR** (o SUPERVISOR) de los archivos.
- **Reemplazo de nombres de coordinador**: en **Datos → Coordinadores**. Cada reemplazo cambia el nombre en lo ya guardado y en las próximas cargas. Ya está cargado **DIEGO AMIAS → HERLES CORNELIO**.
- Los nombres de cadena, tienda, marca y coordinador se guardan en **MAYÚSCULAS**, para que "Tottus" y "TOTTUS" cuenten como la misma cadena.

## Análisis del Vs

Arriba de las tablas de la Comparativa aparece un análisis escrito automáticamente para la selección actual:

- resultado general en soles y %;
- unidades y precio promedio, y si la variación se explica más por volumen o por precio/mix;
- cuántos crecen y cuántos caen, y quiénes más suman y más restan;
- quién concentra más venta;
- la participación de Paraíso;
- la mejor y la peor semana.

El análisis se incluye en la hoja "Análisis" del Excel descargado.

## Cobertura

Los archivos traen una columna de cobertura, por ejemplo **COBERTURA** o **COBERTURADO**. La app la detecta sola y entiende valores como SÍ/NO, COBERTURADO/NO COBERTURADO, 1/0 o X.

- **Se guarda todo**, cada fila con su marca de cobertura.
- **El reporte muestra por defecto solo Coberturados.** El filtro "Cobertura" permite ver también No coberturados, Sin dato o Todos.
- **Al subir**, cada archivo muestra cuánto de su venta es coberturada. Si un archivo no trae la columna, la app avisa, porque esos datos no aparecerán en "Solo coberturados".
- **En Datos**, la tabla de la base muestra la venta coberturada de cada año.

## Venta Paraíso vs Competencia

Cada marca se clasifica en uno de dos grupos:

- **Venta Paraíso**: PARAISO, VIVA-HOME, CASA JOVEN, KREA y KOMFORT. Se reconocen aunque vengan escritas distinto, por ejemplo "Paraíso" o "Viva Home".
- **Venta Competencia**: todas las demás marcas.

Lo que agrega esta clasificación:

- **Filtro "Tipo de venta"**: Todas, Solo venta Paraíso o Solo venta Competencia.
- **En "Por año"**: un gráfico y una tabla de Paraíso frente a Competencia, con la participación de Paraíso en el acumulado al corte, en el mes y la semana de corte y en el año completo. La variación de la participación se muestra en puntos porcentuales (pp).
- **Indicador** de participación Paraíso.
- **En "Datos"**: la lista de marcas detectadas y el grupo de cada una, para revisarla.

La lista de marcas Paraíso está en `app.js`, en la constante `MARCAS_PARAISO`.

## Base de datos compartida

Los datos se guardan en PostgreSQL (base **Proyectos_Below_Db**, esquema `reporte_vs_anual`), así que todas las personas ven el mismo reporte.

- **Ver**: cualquiera con el enlace, sin iniciar sesión.
- **Subir y eliminar datos**: solo usuarios con DNI y contraseña. Los administradores crean usuarios desde el botón **Usuarios**.
- **Al guardar**, cada año que venga en los archivos **reemplaza por completo** a ese año en la base. Así, subir el archivo actualizado de 2026 no duplica las ventas. Cada carga queda registrada: quién, cuándo, qué archivos y qué años.
- A la base solo se envían las ventas ya resumidas por cadena, tienda, marca, año, mes y semana, no los archivos completos.

### Arquitectura

| Parte | Servicio en Render | Carpeta |
|---|---|---|
| Página (sitio estático) | `reporte-vs-anual` → https://reporte-vs-anual.onrender.com | raíz del repo |
| API (Node + Express) | `reporte-vs-anual-api` → https://reporte-vs-anual-api.onrender.com | `api/` |
| Base de datos | `Proyectos_Below_Db`, esquema `reporte_vs_anual` | — |

Variables de entorno de la API:

| Variable | Para qué |
|---|---|
| `DATABASE_URL` | *Internal Database URL* de Proyectos_Below_Db |
| `SESSION_SECRET` | Clave para firmar las sesiones |
| `ADMIN_DNI`, `ADMIN_PASSWORD`, `ADMIN_NOMBRE` | Primer administrador; se crea al iniciar si no existe. La contraseña debe tener al menos 8 caracteres y se puede cambiar luego desde la app |
| `ALLOWED_ORIGINS` | Dominios que pueden usar la API (por defecto, la página publicada) |

Las tablas se crean solas al iniciar la API: `usuarios`, `cargas`, `anios` y `ventas`.

La API está en el plan gratuito de Render. Si nadie la usa por un rato entra en reposo, y la primera carga siguiente puede tardar hasta un minuto.

## Cómo usarla en tu computadora

1. API: en `api/`, ejecuta `npm install` y luego `npm start`, con las variables de entorno de arriba (la API queda en http://localhost:3000).
2. Página: en la raíz, ejecuta `python -m http.server 8000` y entra a http://localhost:8000. En local, la página usa la API de `localhost:3000`.

## Publicarla en Render

El archivo `render.yaml` ya define un **sitio estático**:

1. En Render: **New → Blueprint**. Elige este repositorio.
2. Render crea el servicio `reporte-vs-anual` y te da una URL del tipo `https://reporte-vs-anual.onrender.com`.

Otra forma: **New → Static Site**, con *Build command* vacío y *Publish directory* `.`.

## Archivos

| Archivo | Para qué sirve |
|---|---|
| `index.html` | Estructura de la app |
| `styles.css` | Diseño |
| `app.js` | Lectura de archivos, detección de columnas, corte de evaluación, cálculos, gráficos y exportación |
| `vendor/` | Copias locales de SheetJS (lector de Excel) y Chart.js. SheetJS se intenta cargar primero en su versión más reciente desde su CDN oficial |
| `render.yaml` | Configuración para publicar en Render |
