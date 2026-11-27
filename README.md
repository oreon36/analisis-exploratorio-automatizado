# Análisis exploratorio automatizado

Herramienta de línea de comandos en Python que convierte un CSV, TSV o XLSX en un informe HTML autocontenido. Está pensada para revisar un conjunto de datos al inicio de un análisis.

## Ejecutar la demo

Requiere Python 3.11 o superior. Abre una terminal en la carpeta del proyecto y ejecuta:

```powershell
python -m pip install -r requirements.txt
python eda.py ventas_demo.csv -o informe.html
```

Abre `informe.html` en el navegador. El CSV de la demo contiene datos sintéticos de ventas, con faltantes, una fila repetida y valores extremos introducidos de forma intencional.

## Analizar otros archivos

```powershell
python eda.py "ventas.xlsx" --hoja "Ventas" -o ventas.html
python eda.py "exportacion.csv" --separador ";" --codificacion cp1252 -o exportacion.html
```

En Excel se puede seleccionar una hoja por nombre. Si se omite `--hoja`, se lee la primera. El informe se genera localmente y puede abrirse sin conexión.

## Qué entrega

- Perfil de tipos, valores distintos y faltantes por columna.
- Conteo de duplicados de fila completa.
- Estadísticas descriptivas, histogramas y posibles atípicos por rango intercuartílico.
- Frecuencias categóricas, rangos de fechas y correlaciones de Pearson.

La lectura usa los tipos inferidos por pandas. Los identificadores numéricos pueden aparecer en las estadísticas y las correlaciones. Las fechas de texto se reconocen cuando predominan los valores con año al inicio. Los resultados orientan la revisión y requieren contexto para decidir si un valor extremo o una fila repetida es un error.

Antes de compartir un informe, revisa sus etiquetas y frecuencias, ya que pueden incluir valores del archivo original.
