"""Generate a standalone exploratory data report from a local file."""

import argparse
import csv
import html
from itertools import combinations
from pathlib import Path
from tempfile import NamedTemporaryFile

import numpy as np
import pandas as pd


def load_data(path, sheet, separator, encoding):
    if path.suffix.lower() == ".xlsx":
        with pd.ExcelFile(path, engine="openpyxl") as book:
            selected = book.sheet_names[0] if sheet is None else sheet
            return book.parse(selected), selected
    if path.suffix.lower() not in {".csv", ".tsv"}:
        raise ValueError("Formatos compatibles: CSV, TSV y XLSX.")
    if sheet is not None:
        raise ValueError("--hoja se usa con archivos XLSX.")
    with path.open(encoding=encoding, newline="") as source:
        sample = source.read(8192)
    if separator is None:
        try:
            separator = csv.Sniffer().sniff(sample, delimiters=",;\t|").delimiter
        except csv.Error:
            separator = "\t" if path.suffix.lower() == ".tsv" else ","
    return pd.read_csv(path, sep=separator, encoding=encoding), None


def table(headers, rows):
    head = "".join(f"<th scope='col'>{html.escape(str(cell))}</th>" for cell in headers)
    body = "".join(
        "<tr>" + "".join(f"<td>{html.escape(str(cell))}</td>" for cell in row) + "</tr>"
        for row in rows
    )
    return f"<div class='scroll'><table><thead><tr>{head}</tr></thead><tbody>{body}</tbody></table></div>"


def number(value):
    return f"{value:,.4g}" if pd.notna(value) else "—"


def distribution(name, values):
    counts, edges = np.histogram(values, bins=min(12, values.nunique()))
    bars = "".join(
        f"<li><span>{number(left)} a {number(right)}</span>"
        f"<meter min='0' max='{counts.max()}' value='{count}'></meter><strong>{count}</strong></li>"
        for left, right, count in zip(edges[:-1], edges[1:], counts)
    )
    return f"<article><h3>{html.escape(str(name))}</h3><ul class='bars'>{bars}</ul></article>"


def report(frame, source, sheet):
    total = len(frame)
    blanks = frame.select_dtypes(include=["object", "string"]).apply(
        lambda series: series.str.fullmatch(r"\s*", na=False)
    )
    frame = frame.copy()
    frame[blanks.columns] = frame[blanks.columns].mask(blanks)
    missing = frame.isna().sum()
    duplicates = int(frame.duplicated().sum())
    numeric = frame.select_dtypes(include="number").replace([np.inf, -np.inf], np.nan)
    quality = []
    stats = []
    charts = []
    categories = []
    dates = []
    for name, series in frame.items():
        label = str(name)
        valid = series.dropna()
        kind = str(series.dtype)
        if name in numeric:
            values = numeric[name].dropna()
            q1, q3 = values.quantile([0.25, 0.75])
            spread = q3 - q1
            outliers = "—"
            if len(values) >= 4 and spread > 0:
                outliers = int(((values < q1 - 1.5 * spread) | (values > q3 + 1.5 * spread)).sum())
            stats.append((label, len(values), len(valid) - len(values), number(values.min()),
                          number(values.median()), number(values.mean()), number(values.max()), outliers))
            if len(values) and len(charts) < 8:
                charts.append(distribution(name, values))
        elif len(valid):
            if pd.api.types.is_datetime64_any_dtype(series):
                parsed = valid
            else:
                date_like = valid.astype(str).str.match(r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}(?:$|[ T])")
                parsed = pd.to_datetime(valid.where(date_like), format="mixed", errors="coerce", utc=True)
            if parsed.notna().mean() >= 0.8:
                kind = "Fecha y hora"
                dates.append((label, str(parsed.min()), str(parsed.max()), int(parsed.isna().sum())))
            elif len(categories) < 12 and valid.nunique() <= min(50, max(12, len(valid) // 2)):
                counts = valid.value_counts().head(5)
                categories.append((label, ", ".join(f"{value} ({count})" for value, count in counts.items())))
        quality.append((label, kind, int(missing[name]), f"{missing[name] / total:.1%}" if total else "—",
                        series.nunique()))

    eligible = numeric.loc[:, numeric.nunique() > 1]
    correlations = eligible.corr(min_periods=4)
    pairs = []
    for left, right in combinations(eligible.columns, 2):
        value = correlations.at[left, right]
        if pd.notna(value):
            pairs.append((left, right, value))
    pairs.sort(key=lambda pair: abs(pair[2]), reverse=True)
    correlation_rows = [
        (left, right, f"{value:.3f}", int(eligible[[left, right]].dropna().shape[0]))
        for left, right, value in pairs[:10]
    ]
    highlights = [
        f"{name} tiene {count} valores faltantes ({count / total:.1%})."
        for name, count in missing[missing > 0].sort_values(ascending=False).head(5).items()
    ]
    if duplicates:
        highlights.append(f"Hay {duplicates} filas repetidas después de su primera aparición.")
    if not highlights:
        highlights.append("No se detectaron valores faltantes ni filas repetidas.")
    findings = "".join(f"<li>{html.escape(line)}</li>" for line in highlights)
    source_label = source.name if sheet is None else f"{source.name} · Hoja {sheet}"
    sections = []
    if stats:
        sections.append("<h2>Variables numéricas</h2>" + table(
            ("Columna", "Valores finitos", "Infinitos", "Mínimo", "Mediana", "Media", "Máximo", "Atípicos IQR"), stats))
    if charts:
        sections.append("<h2>Distribuciones</h2><p>Primeras ocho columnas numéricas con datos, según el orden de origen.</p>"
                        "<div class='charts'>" + "".join(charts) + "</div>")
    if categories:
        sections.append("<h2>Categorías frecuentes</h2><p>Hasta doce columnas de baja cardinalidad y cinco valores por columna.</p>"
                        + table(("Columna", "Valor (frecuencia)"), categories))
    if dates:
        sections.append("<h2>Fechas</h2>" + table(("Columna", "Primera", "Última", "Sin interpretar"), dates))
    if correlation_rows:
        sections.append("<h2>Correlaciones de Pearson</h2><p>Hasta diez pares ordenados por magnitud absoluta.</p>"
                        + table(("Variable A", "Variable B", "r", "Observaciones en común"), correlation_rows))

    return f"""<!doctype html>
<html lang="es"><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Informe exploratorio · {html.escape(source_label)}</title>
<style>
*{{box-sizing:border-box}} body{{margin:0;background:#f4f6f7;color:#18262e;font:16px/1.55 system-ui,sans-serif}}
header{{background:#12343b;color:#fff;padding:3rem max(5vw,calc((100vw - 1050px)/2));overflow-wrap:anywhere}}
h1{{font-size:clamp(2rem,5vw,3.2rem);margin:0;line-height:1.15}} header p{{margin-bottom:0}}
main{{max-width:1100px;margin:auto;padding:2rem 1.5rem 3rem}} h2{{margin:2rem 0 .55rem}}
.metrics,.charts{{display:grid;grid-template-columns:repeat(auto-fit,minmax(min(100%,215px),1fr));gap:1rem}}
article,.scroll,details{{background:white;border:1px solid #d3dfe2;border-radius:10px;padding:1rem}}
.metrics strong{{display:block;font-size:1.7rem}} .scroll{{overflow:auto;padding:0}}
table{{width:100%;border-collapse:collapse;text-align:left}} th,td{{padding:.65rem .9rem;vertical-align:top}}
th{{background:#e2eeed}} td{{border-top:1px solid #e0e8ea}} h3{{margin:0 0 .7rem;overflow-wrap:anywhere}}
.bars{{padding:0;list-style:none;margin:0;font-size:.84rem}} .bars li{{display:grid;grid-template-columns:1fr 52px auto;gap:.4rem;align-items:center}}
meter{{width:100%}} meter::-webkit-meter-optimum-value{{background:#19746b}} details{{margin-top:2rem}} summary{{cursor:pointer}}
</style>
<header><h1>Informe exploratorio</h1><p>{html.escape(source_label)}</p></header>
<main><div class="metrics">
<article>Filas<strong>{total}</strong></article><article>Columnas<strong>{len(frame.columns)}</strong></article>
<article>Celdas faltantes<strong>{int(missing.sum())}</strong></article><article>Filas repetidas<strong>{duplicates}</strong></article>
</div><h2>Hallazgos</h2><ul>{findings}</ul><h2>Perfil de columnas</h2>
{table(("Columna", "Tipo detectado", "Faltantes", "% faltante", "Valores distintos"), quality)}
{''.join(sections)}
<details><summary>Criterios de cálculo</summary>
<p>Los faltantes incluyen celdas vacías, espacios en blanco y los marcadores predeterminados de pandas. Las filas repetidas se cuentan sobre todas las columnas después de normalizar los espacios en blanco.</p>
<p>Los atípicos se cuentan fuera de Q1 − 1,5 × IQR y Q3 + 1,5 × IQR, con cuatro valores finitos como mínimo e IQR positivo. Las correlaciones usan al menos cuatro observaciones en común.</p>
<p>Las fechas de texto se detectan cuando al menos el 80 % de los valores presentes se interpretan con año al inicio. Esas fechas se muestran en UTC. Los límites de los histogramas están redondeados.</p>
</details></main></html>"""


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("entrada", type=Path)
    parser.add_argument("-o", "--salida", type=Path, default=Path("informe.html"))
    parser.add_argument("--hoja", help="Nombre de la hoja de Excel. Por defecto se lee la primera.")
    parser.add_argument("--separador", choices=[",", ";", "|", "\t"])
    parser.add_argument("--codificacion", default="utf-8-sig")
    args = parser.parse_args()
    temporary_path = None
    try:
        if args.salida.suffix.lower() != ".html" or args.entrada.resolve() == args.salida.resolve():
            raise ValueError("Indica una ruta de salida distinta con extensión .html.")
        frame, sheet = load_data(args.entrada, args.hoja, args.separador, args.codificacion)
        content = report(frame, args.entrada, sheet)
        with NamedTemporaryFile(dir=args.salida.parent, delete=False, mode="w", encoding="utf-8") as target:
            temporary_path = Path(target.name)
            target.write(content)
        temporary_path.replace(args.salida)
    except (OSError, ValueError, ImportError, LookupError) as error:
        parser.exit(2, f"Error: {error}\n")
    finally:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)
    print(f"Informe generado: {args.salida}")


if __name__ == "__main__":
    main()
