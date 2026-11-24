# Reporte Vs Anual

App web para comparar ventas **año contra año** a partir de archivos de Excel o CSV. Parte de la idea de *Reporte 30K Paraíso*, pero muestra el informe completo, sin filtrar por el umbral de 30K: subes tus archivos, la app detecta las columnas sola y te da el comparativo por cadena, mes y tienda.

## Qué hace

- **Sube varios archivos a la vez**: `.xlsx`, `.xls`, `.xlsm`, `.xlsb` y `.csv`/`.txt`. Por ejemplo, un archivo por año (2025 y 2026) o uno solo con todo.
- **Los archivos no salen de tu equipo**: todo se procesa en el navegador.
- **Detecta las columnas automáticamente**: Venta, Fecha (o Año + Mes), Cadena, Tienda/PDV, Marca y Unidades. Si alguna sale mal, la cambias en la lista y todo se recalcula. La app recuerda tu elección para archivos con los mismos encabezados.
- **Encuentra la fila de encabezados** aunque el archivo tenga títulos arriba, y elige la hoja correcta del libro.
- **Formato ancho**: también lee reportes con los meses como columnas (`Ene-25`, `2025-01` o fechas en el encabezado). Las filas de TOTAL se omiten.
- **CSV**: detecta el separador (`;`, `,`, tabulación o `|`), la codificación (UTF-8 o Windows-1252) y los formatos `1,234.56` y `1.234,56`. Las fechas se leen como día/mes/año.

### Comparativo

- Eliges el **año base** y el **año a comparar**.
- **Período**: *Mismos meses (YTD)*, que es la opción por defecto y compara solo los meses que ya tienen venta en el año actual; *Año completo*; o un *Rango de meses*.
- Filtros por **cadena** y **marca**.
- Muestra **el informe completo**: todas las cadenas, todos los meses del período y todas las tiendas. No hay umbral ni corte por monto.
- Indicadores: venta de cada año, variación en S/ y en %, tiendas que crecen o caen, y tiendas nuevas o sin venta.
- Gráficos: venta mensual de un año frente al otro y variación % de todas las cadenas.
- Tablas: resumen por cadena, detalle mensual y **detalle por tienda**. Esta última tiene búsqueda, filtro por estado, orden por columna y, si quieres, las columnas de cada mes.
- **Descargar Excel** (hojas Resumen, Por cadena, Mensual y Por tienda) y **Copiar tabla** para pegarla en Excel.

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
| `app.js` | Lectura de archivos, detección de columnas, cálculos, gráficos y exportación |
| `vendor/` | Copias locales de SheetJS (lector de Excel) y Chart.js. SheetJS se intenta cargar primero en su versión más reciente desde su CDN oficial |
| `render.yaml` | Configuración para publicar en Render |
