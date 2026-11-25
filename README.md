# Reporte Vs Anual

App web para **evaluar las ventas de todos los años** (por defecto desde 2022) a partir de archivos de Excel o CSV. Subes los archivos y la app detecta las columnas sola. El corte de evaluación se toma siempre de **la fecha más reciente que haya en los archivos**: el último año, mes y semana con venta.

Publicada en: https://reporte-vs-anual.onrender.com

## Lectura de archivos

- **Sube varios a la vez**: `.xlsx`, `.xls`, `.xlsm`, `.xlsb` y `.csv`/`.txt`. Puede ser un archivo por año o uno solo con todos.
- **Los archivos no salen de tu equipo**: todo se procesa en el navegador.
- **Detecta las columnas automáticamente**: Venta, Fecha (o Año + Mes, o Año + Semana), Semana, Cadena, Tienda/PDV, Marca y Unidades. Si alguna sale mal, la cambias en la lista y la app recuerda tu elección.
- **Encuentra la fila de encabezados** aunque haya títulos arriba, y elige la hoja correcta del libro.
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

- **Por año**: indicadores del año en curso frente al anterior (acumulado, mes de corte y semana de corte), gráficos por año y una tabla con todos los años, cada uno con su variación contra el año anterior.
- **Por mes**: gráfico de líneas con un año por línea y la tabla de meses × años, más la variación del último año contra el anterior.
- **Por semana**: lo mismo, semana por semana.
- **Por cadena** y **Por tienda**: todos los años en columnas, con la medida que elijas, la variación del último año contra el anterior, el estado (crece, cae, nueva o sin venta) y la participación. Las columnas se ordenan con un clic y las tiendas se pueden buscar.
- Filtros: **desde el año**, **cadena** y **marca**.
- **Descargar Excel**: un libro con las hojas Por año, Por mes, Por semana, Por cadena y Por tienda. Cada tabla también tiene **Copiar tabla** para pegarla en Excel.

## Cómo usarla en tu computadora

No necesita instalación. Abre una terminal en la carpeta del proyecto y ejecuta:

```bash
python -m http.server 8000
```

Luego entra a http://localhost:8000.

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
