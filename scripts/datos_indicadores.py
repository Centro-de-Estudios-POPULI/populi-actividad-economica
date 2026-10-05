"""
datos_indicadores.py — Monitor de Actividad Económica: indicadores adelantados del INE → data/ind_*.json

    python scripts/datos_indicadores.py            # lee datos/indicadores_adelantados/ y escribe data/
    python scripts/datos_indicadores.py --revisar  # sólo lee y verifica; no escribe

Los gráficos (embed/*.html) leen estos JSON con PM.cargar: ningún número se escribe a mano en un
gráfico. Usa SÓLO los Excel que baja scripts/scrape_ine.py (el resto de la carpeta está
congelado). Escribe un JSON por gráfico:

  data/ind_cemento.json      producción y consumo de cemento, mensual desde 1991 (t)
  data/ind_permisos.json     permisos de construcción: número y superficie, mensual desde 2008
  data/ind_manufactura.json  índice de producción y utilización de la capacidad, trimestral desde 2017
  data/ind_energia.json      índice de consumo de energía eléctrica por categoría, mensual desde 2017
  data/ind_transporte.json   índice general de transporte por modalidad, mensual desde 2017
  data/ind_ferroviario.json  carga ferroviaria por red, mensual desde 1999 (t)

Reglas (las mismas de datos_pib.py):
- Los cuadros se leen por RÓTULO: la fila de encabezado es la que dice «PERIODO» y cada columna se
  identifica por su texto; los años y los meses por su rótulo («2026(p)», «Mayo »). Nunca por fila
  o columna fija: el INE inserta filas.
- Antes de escribir se verifica: meses (o trimestres) consecutivos y sin repetidos; el total
  nacional = la suma de sus partes (departamentos, tipos de trámite); el valor anual del cuadro =
  la suma (flujos) o el promedio (índices) de sus meses; la variación interanual que publica el INE
  = la que sale de su propio índice.
- Si algo no cierra, aborta SIN tocar data/. También aborta si el último dato de una serie
  RETROCEDE respecto del JSON ya escrito (fuente congelada o archivo cambiado: no se publica un
  dato peor que el que ya está).
- La salida no lleva fecha de generación: dos corridas seguidas con los mismos Excel dejan los
  mismos bytes (y un archivo que no cambia no se reescribe).
"""
from __future__ import annotations

import json
import re
import sys
import unicodedata
from pathlib import Path

import openpyxl
import xlrd

BASE = Path(__file__).resolve().parent.parent
DATOS = BASE / "datos" / "indicadores_adelantados"
DATA = BASE / "data"

MESES = {"ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4, "MAYO": 5, "JUNIO": 6, "JULIO": 7,
         "AGOSTO": 8, "SEPTIEMBRE": 9, "SETIEMBRE": 9, "OCTUBRE": 10, "NOVIEMBRE": 11, "DICIEMBRE": 12}
ROMANO = {"I": 1, "II": 2, "III": 3, "IV": 4}
URL_INE = "https://www.ine.gob.bo/index.php/estadisticas-economicas/"


class Falla(Exception):
    pass


def norma(s) -> str:
    s = unicodedata.normalize("NFD", str(s if s is not None else "")).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s).strip().upper()


def num(v):
    if v is None or isinstance(v, str):
        if isinstance(v, str):
            try:
                return float(v.replace(",", "."))
            except ValueError:
                return None
        return None
    try:
        x = float(v)
    except (TypeError, ValueError):
        return None
    return None if x != x else x


def r(v, d):
    return None if v is None else round(v, d)


# ── Lectura de un libro (xlsx u xls viejo con extensión .xlsx) ────────────────
def hojas(ruta: Path) -> dict[str, list[list]]:
    with open(ruta, "rb") as fh:
        firma = fh.read(4)
    out = {}
    if firma == b"PK\x03\x04":
        wb = openpyxl.load_workbook(ruta, data_only=True, read_only=True)
        for ws in wb.worksheets:
            out[ws.title] = [list(f) for f in ws.iter_rows(values_only=True)]
        wb.close()
    else:                                   # los permisos vienen en formato .xls (OLE2) aunque digan .xlsx
        wb = xlrd.open_workbook(ruta)
        for sh in wb.sheets():
            out[sh.name] = [sh.row_values(i) for i in range(sh.nrows)]
    return out


def hoja(libro: dict, *claves: str) -> list[list]:
    """La primera hoja cuyo nombre normalizado contiene todas las claves (la primera hoja si no hay claves)."""
    if not claves:
        return next(iter(libro.values()))
    for nombre, filas in libro.items():
        n = norma(nombre)
        if all(c in n for c in claves):
            return filas
    raise Falla(f"no se encontró la hoja {claves} (hay: {list(libro)})")


def encabezado(filas: list[list], ruta: Path) -> tuple[int, int]:
    """(fila, columna) de la celda «PERIODO»."""
    for i, f in enumerate(filas):
        for c, v in enumerate(f):
            if norma(v) == "PERIODO":
                return i, c
    raise Falla(f"{ruta.name}: no se encontró el encabezado «PERIODO»")


def anio_de(v):
    """2025 · 2025.0 · '2025(p)' · '2026 (p)' → (2025, preliminar?) ; otra cosa → None."""
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        a = int(v)
        return (a, False) if 1980 <= a <= 2100 and a == v else None
    m = re.match(r"^\s*(\d{4})\s*(\(\s*P\s*\))?\s*$", norma(v))
    if m and 1980 <= int(m.group(1)) <= 2100:
        return int(m.group(1)), bool(m.group(2))
    return None


def columnas(filas, ih, ic, reglas, ruta, filas_extra=0):
    """{clave: columna} según el texto del encabezado (puede ocupar filas_extra filas más abajo).
    reglas: [(clave, función(texto_normalizado) → bool)]; cada clave tiene que aparecer una sola vez."""
    ancho = max(len(f) for f in filas[ih:ih + 1 + filas_extra])
    textos = []
    for c in range(ancho):
        partes = [norma(filas[i][c]) if c < len(filas[i]) else "" for i in range(ih, ih + 1 + filas_extra)]
        textos.append(" ".join(p for p in partes if p))
    out = {}
    for clave, prueba in reglas:
        hallado = [c for c in range(ancho) if c != ic and textos[c] and prueba(textos[c])]
        if len(hallado) != 1:
            raise Falla(f"{ruta.name}: la columna «{clave}» aparece {len(hallado)} veces en el encabezado {textos}")
        out[clave] = hallado[0]
    return out


def leer_mensual(filas, ruta, cols, ih, ic):
    """Cuadro INE «año / meses» → (meses ['1991-01',…], {clave: [valores]}, {año: {clave: valor anual}}, años (p))."""
    meses, valores, anual, prelim = [], {k: [] for k in cols}, {}, []
    anio = None
    for f in filas[ih + 1:]:
        v = f[ic] if ic < len(f) else None
        n = norma(v)
        if not n:
            continue
        if n.startswith("FUENTE") or n.startswith("(") or n.startswith("NOTA") or n.startswith("INSTITUTO"):
            break
        a = anio_de(v)
        if a:
            anio = a[0]
            if a[1]:
                prelim.append(anio)
            if anio in anual:
                raise Falla(f"{ruta.name}: el año {anio} aparece dos veces")
            anual[anio] = {k: num(f[c]) if c < len(f) else None for k, c in cols.items()}
            continue
        m = MESES.get(n)
        if m is None:
            raise Falla(f"{ruta.name}: rótulo de período desconocido «{v}»")
        if anio is None:
            raise Falla(f"{ruta.name}: un mes («{v}») antes del primer año")
        meses.append(f"{anio}-{m:02d}")
        for k, c in cols.items():
            valores[k].append(num(f[c]) if c < len(f) else None)
    consecutivos(meses, ruta, 12)
    return meses, valores, anual, prelim


def leer_trimestral(filas, ruta, cols, ih, ic):
    """Cuadro INE «año / I-IV Trimestre» → (claves ['2017-T1',…], {clave: [valores]}, {año: {…}}, años (p), avisos).
    Si un rótulo repite el trimestre anterior dentro del mismo año (el INE escribió «III» dos veces en
    T4 2025 de la capacidad instalada) se acepta como el trimestre siguiente SÓLO si el valor anual
    del cuadro (promedio de sus trimestres) lo confirma; queda anotado como aviso."""
    claves, valores, anual, prelim, avisos = [], {k: [] for k in cols}, {}, [], []
    anio, previo = None, 0
    for f in filas[ih + 1:]:
        v = f[ic] if ic < len(f) else None
        n = norma(v)
        if not n:
            continue
        if n.startswith("FUENTE") or n.startswith("(") or n.startswith("NOTA"):
            break
        a = anio_de(v)
        if a:
            anio, previo = a[0], 0
            if a[1]:
                prelim.append(anio)
            if anio in anual:
                raise Falla(f"{ruta.name}: el año {anio} aparece dos veces")
            anual[anio] = {k: num(f[c]) if c < len(f) else None for k, c in cols.items()}
            continue
        m = re.match(r"^(IV|III|II|I)\s+TRIM", n)
        if not m or anio is None:
            raise Falla(f"{ruta.name}: rótulo de período desconocido «{v}»")
        q = ROMANO[m.group(1)]
        if q == previo and q < 4:
            avisos.append((anio, q + 1, str(v).strip()))
            q = q + 1
        previo = q
        claves.append(f"{anio}-T{q}")
        for k, c in cols.items():
            valores[k].append(num(f[c]) if c < len(f) else None)
    consecutivos(claves, ruta, 4)
    return claves, valores, anual, prelim, avisos


def consecutivos(claves, ruta, frec):
    if not claves:
        raise Falla(f"{ruta.name}: no se leyó ningún período")
    def ordinal(k):
        a, p = k.split("-")
        return int(a) * frec + (int(p[1:]) - 1 if p.startswith("T") else int(p) - 1)
    for x, y in zip(claves, claves[1:]):
        if ordinal(y) != ordinal(x) + 1:
            raise Falla(f"{ruta.name}: períodos no consecutivos o repetidos ({x} → {y})")


def recortar(claves, valores):
    """Quita los períodos del final sin ningún dato (filas vacías que el INE deja preparadas)."""
    n = len(claves)
    while n and all(valores[k][n - 1] is None for k in valores):
        n -= 1
    return claves[:n], {k: v[:n] for k, v in valores.items()}


def ultimo_con_dato(claves, serie):
    for k, v in zip(reversed(claves), reversed(serie)):
        if v is not None:
            return k
    return None


# ── Verificaciones ───────────────────────────────────────────────────────────
def verificar_suma(nombre, claves, total, partes, tol, fallas):
    for i, k in enumerate(claves):
        t = total[i]
        ps = [p[i] for p in partes]
        if t is None or any(p is None for p in ps):
            if t is not None or any(p is not None for p in ps):
                fallas.append(f"{nombre} {k}: total o partes incompletos")
            continue
        if abs(t - sum(ps)) > tol:
            fallas.append(f"{nombre} {k}: total {t:,.2f} ≠ suma de partes {sum(ps):,.2f}")


def verificar_anual(nombre, claves, serie, anual, clave, modo, tol, fallas, frec=12, avisos=None):
    """El valor del año en el cuadro = suma ('suma') o promedio ('prom') de sus períodos (también el año parcial).
    Con `avisos`, un año INCOMPLETO que no cierra se avisa en vez de abortar: en la capacidad instalada el INE pone
    en la fila de 2026 (dos trimestres) el valor del último trimestre, no el promedio. Un año completo siempre aborta."""
    por_anio: dict[int, list] = {}
    for k, v in zip(claves, serie):
        por_anio.setdefault(int(k[:4]), []).append(v)
    for a, vs in por_anio.items():
        if a not in anual:
            fallas.append(f"{nombre}: el año {a} no tiene fila anual")
            continue
        esperado = anual[a].get(clave)
        if esperado is None or any(v is None for v in vs):
            continue
        calc = sum(vs) if modo == "suma" else sum(vs) / len(vs)
        if abs(calc - esperado) > tol * (max(1.0, abs(esperado)) if modo == "prom" else 1):
            if avisos is not None and len(vs) < frec:
                avisos.append(f"{nombre} {a} (año incompleto, {len(vs)} de {frec}): la fila anual del cuadro ({esperado:,.3f}) "
                              f"no es el {'total' if modo == 'suma' else 'promedio'} de sus períodos ({calc:,.3f}); se usan los períodos")
                continue
            fallas.append(f"{nombre} {a}: valor anual del cuadro {esperado:,.3f} ≠ {'suma' if modo == 'suma' else 'promedio'} de sus períodos {calc:,.3f}")


def verificar_interanual(nombre, claves, indice, claves_ine, var_ine, frec, fallas, avisos, prelim, tol=0.02):
    """La variación interanual publicada por el INE = la que sale de su propio índice (por FECHA).
    En los años preliminares (p) el INE revisa el índice y a veces no actualiza la hoja de variaciones:
    ahí una diferencia de hasta 1 pp es un AVISO (se publica la variación que sale del índice); más que
    eso, o en un año definitivo, es una falla (columna o período mal leídos)."""
    pos = {k: i for i, k in enumerate(claves)}
    n, revisados = 0, []
    for k, v in zip(claves_ine, var_ine):
        if v is None or k not in pos:
            continue
        a, p = k.split("-")
        previo = f"{int(a) - 1}-{p}"
        if previo not in pos:
            continue
        i0, i1 = pos[previo], pos[k]
        if indice[i0] in (None, 0) or indice[i1] is None:
            continue
        calc = (indice[i1] / indice[i0] - 1) * 100
        n += 1
        if abs(calc - v) > tol:
            if int(a) in prelim and abs(calc - v) <= 1.0:
                revisados.append(f"{k} ({v:.2f} vs {calc:.2f})")
            else:
                fallas.append(f"{nombre} {k}: variación del INE {v:.3f}% ≠ la de su índice {calc:.3f}%")
    if revisados:
        avisos.append(f"{nombre}: en {len(revisados)} mes(es) preliminares la hoja de variaciones del INE no sigue a su índice "
                      f"(¿índice revisado?); se usa la del índice: " + ", ".join(revisados[:4]) + ("…" if len(revisados) > 4 else ""))
    if n == 0:
        fallas.append(f"{nombre}: no se pudo comparar ninguna variación interanual")


def no_retrocede(nombre_json, nuevo, fallas):
    """El último dato de cada serie no puede ser anterior al del JSON ya publicado."""
    ruta = DATA / nombre_json
    if not ruta.exists():
        return
    try:
        viejo = json.loads(ruta.read_text(encoding="utf-8"))
    except (ValueError, OSError):
        return
    uv, un = viejo.get("metadata", {}).get("ultimo_por_serie", {}), nuevo["metadata"]["ultimo_por_serie"]
    for k, u in uv.items():
        if k not in un or un[k] is None:
            fallas.append(f"{nombre_json}: la serie «{k}» desapareció (antes llegaba a {u})")
        elif u is not None and un[k] < u:
            fallas.append(f"{nombre_json}: la serie «{k}» retrocede de {u} a {un[k]} (¿fuente congelada o archivo cambiado?)")


def meta(fuente, url, archivos, unidad, claves, series, prelim, notas, extra=None):
    ult = {k: ultimo_con_dato(claves, v) for k, v in series.items()}
    m = {"fuente": fuente, "url": url, "archivos": archivos, "unidad": unidad,
         "primero": claves[0], "ultimo": max(u for u in ult.values() if u), "ultimo_por_serie": ult,
         "preliminar_desde": min(prelim) if prelim else None, "notas": notas}
    if extra:
        m.update(extra)
    return m


# ── Cemento ──────────────────────────────────────────────────────────────────
def cemento(fallas):
    out_series, archivos, prelims = {}, [], []
    for clave, arch in (("produccion", "cemento/produccion_cemento_depto_1991_2026.xlsx"),
                        ("consumo", "cemento/consumo_cemento_depto_2000_2026.xlsx")):
        ruta = DATOS / arch
        filas = hoja(hojas(ruta))
        ih, ic = encabezado(filas, ruta)
        textos = [norma(v) for v in filas[ih]]
        ct = [c for c, t in enumerate(textos) if t == "TOTAL"]
        if len(ct) != 1:
            raise Falla(f"{ruta.name}: no hay una única columna TOTAL")
        dep = [c for c, t in enumerate(textos) if t and c not in (ic, ct[0])]
        if len(dep) < 6:
            raise Falla(f"{ruta.name}: se esperaban al menos 6 departamentos, hay {len(dep)}")
        cols = {"total": ct[0], **{f"d{c}": c for c in dep}}
        meses, val, anual, prelim = leer_mensual(filas, ruta, cols, ih, ic)
        meses, val = recortar(meses, val)
        verificar_suma(f"cemento {clave}", meses, val["total"], [val[f"d{c}"] for c in dep], 1.0, fallas)
        verificar_anual(f"cemento {clave}", meses, val["total"], anual, "total", "suma", 1.0, fallas)
        archivos.append(arch)
        prelims += prelim
        out_series[clave] = (meses, val["total"])
    # (las ventas totales del cuadro C2 coinciden con el consumo total recién desde marzo de 2016; antes
    # difieren hasta en 10%: no es una identidad y no se usan)
    claves = sorted(set(out_series["produccion"][0]) | set(out_series["consumo"][0]))
    consecutivos(claves, Path("cemento"), 12)
    series = {}
    for k, (ms, vs) in out_series.items():
        d = dict(zip(ms, vs))
        series[k] = [r(d.get(m), 2) for m in claves]
    return {
        "metadata": meta("INE con datos de las empresas productoras de cemento", URL_INE + "construccion/produccion-venta-y-consumo-de-cemento-cuadros-estadisticos/",
                         archivos, "toneladas métricas", claves, series, prelims,
                         ["El consumo es la cantidad de cemento demandada en el país, por departamento de destino.",
                          "El consumo se publica desde 2000; la producción, desde 1991.",
                          "Cifras preliminares (p) del INE en los años marcados, sujetas a revisión."],
                         {"nombres": {"produccion": "Producción", "consumo": "Consumo"}}),
        "meses": claves,
        "series": series,
    }


# ── Permisos de construcción ─────────────────────────────────────────────────
def permisos(fallas):
    TIPOS = [("total", lambda t: t == "TOTAL"),
             ("aprobacion", lambda t: t.startswith("APROBACION")),
             ("legalizacion", lambda t: t.startswith("LEGALIZACION")),
             ("otros", lambda t: t.startswith("OTROS"))]
    series, archivos, prelims, claves = {}, [], [], None
    for clave, arch, dec in (("numero", "construccion/numero_permisos_2008_2026.xlsx", 0),
                             ("superficie", "construccion/superficie_permisos_2008_2026.xlsx", 1)):
        ruta = DATOS / arch
        filas = hoja(hojas(ruta))
        ih, ic = encabezado(filas, ruta)
        cols = columnas(filas, ih, ic, TIPOS, ruta)
        meses, val, anual, prelim = leer_mensual(filas, ruta, cols, ih, ic)
        meses, val = recortar(meses, val)
        verificar_suma(f"permisos {clave}", meses, val["total"], [val["aprobacion"], val["legalizacion"], val["otros"]], 0.6, fallas)
        verificar_anual(f"permisos {clave}", meses, val["total"], anual, "total", "suma", 1.0, fallas)
        if claves is None:
            claves = meses
        elif meses != claves:
            fallas.append(f"permisos: número y superficie no cubren los mismos meses ({claves[-1]} / {meses[-1]})")
        series[clave] = [r(v, dec) for v in val["total"]]
        series[clave + "_aprobacion"] = [r(v, dec) for v in val["aprobacion"]]
        archivos.append(arch)
        prelims += prelim
    return {
        "metadata": meta("INE con datos de los Gobiernos Autónomos Municipales (GAM)", URL_INE + "construccion/permisos-de-construccion-cuadros-estadisticos/",
                         archivos, {"numero": "número de registros", "superficie": "metros cuadrados"}, claves,
                         {k: v for k, v in series.items() if k in ("numero", "superficie")}, prelims,
                         ["Cubre las ciudades capitales y las conurbaciones de La Paz, Cochabamba y Santa Cruz (no todo el país).",
                          "Incluye construcciones residenciales y no residenciales y todos los tipos de trámite: aprobación de planos (aprobaciones, ampliaciones, remodelaciones), legalización y regularización, y otros.",
                          "Cifras preliminares (p) del INE en los años marcados, sujetas a revisión."],
                         {"nombres": {"numero": "Permisos", "superficie": "Superficie"}}),
        "meses": claves,
        "series": series,
    }


# ── Manufactura (trimestral) ─────────────────────────────────────────────────
def manufactura(fallas, avisos):
    def reglas(pre):
        return [("total", lambda t: "MANUFACTU" in t and "ALIMENTOS" not in t and "OTRAS" not in t),
                ("alimentos", lambda t: "ALIMENTOS" in t),
                ("otras", lambda t: "OTRAS" in t)]
    out, archivos, prelims, claves_ref = {}, [], [], None
    for clave, arch, nombre_hoja in (("produccion", "manufactura/indice_produccion_manufactura_2017_2025.xlsx", ("INDICE",)),
                                     ("capacidad", "manufactura/utilizacion_capacidad_instalada_2017_2025.xlsx", ())):
        ruta = DATOS / arch
        libro = hojas(ruta)
        filas = hoja(libro, *nombre_hoja)
        ih, ic = encabezado(filas, ruta)
        cols = columnas(filas, ih, ic, reglas(clave), ruta)
        claves, val, anual, prelim, av = leer_trimestral(filas, ruta, cols, ih, ic)
        claves, val = recortar(claves, val)
        for k in cols:   # valor anual del cuadro = promedio de sus trimestres (confirma también el rótulo corregido)
            verificar_anual(f"manufactura {clave} {k}", claves, val[k], anual, k, "prom", 0.0005, fallas, frec=4, avisos=avisos)
        for a, q, rot in av:
            avisos.append(f"{ruta.name}: el rótulo «{rot}» de {a} repite el trimestre anterior; se leyó como T{q} (el promedio anual del cuadro lo confirma)")
        if clave == "produccion":   # la variación interanual del INE sale de su índice
            fi = hoja(libro, "INTERANUAL")
            jh, jc = encabezado(fi, ruta)
            ci = columnas(fi, jh, jc, reglas(clave), ruta)
            ck, cv, _, _, _ = leer_trimestral(fi, ruta, ci, jh, jc)
            for k in ci:
                verificar_interanual(f"manufactura {k}", claves, val[k], ck, cv[k], 4, fallas, avisos, set(prelim))
        if claves_ref is None:
            claves_ref = claves
        elif claves != claves_ref:
            fallas.append(f"manufactura: producción y capacidad no cubren los mismos trimestres ({claves_ref[-1]} / {claves[-1]})")
        out[clave] = {k: [r(v, 4) for v in val[k]] for k in cols}
        archivos.append(arch)
        prelims += prelim
    planas = {f"{a}_{b}": v for a, d in out.items() for b, v in d.items()}
    return {
        "metadata": meta("INE — Encuesta trimestral de la industria manufacturera", URL_INE + "industria-manufacturera-y-comercio/estadisticas-coyunturales-cuadros-estadisticos/",
                         archivos, {"produccion": "índice de volumen (2017 = 100)", "capacidad": "porcentaje de la capacidad instalada"},
                         claves_ref, planas, prelims,
                         ["Índice de volumen de la producción con año de referencia 2017 (base móvil): 2017 = 100.",
                          "La muestra corresponde principalmente a empresas grandes y medianas.",
                          "El índice no está desestacionalizado: el primer trimestre es siempre el más bajo.",
                          "Cifras preliminares (p) del INE en los años marcados, sujetas a revisión."],
                         {"nombres": {"total": "Industria manufacturera", "alimentos": "Alimentos y bebidas", "otras": "Otras industrias"}}),
        "trimestres": claves_ref,
        "series": out,
    }


# ── Índices mensuales con desglose (energía y transporte) ────────────────────
def indice_mensual(fallas, avisos, arch, reglas, nombre, hoja_ind, hoja_var, extra_enc=0):
    ruta = DATOS / arch
    libro = hojas(ruta)
    filas = hoja(libro, *hoja_ind)
    ih, ic = encabezado(filas, ruta)
    cols = columnas(filas, ih, ic, reglas, ruta, extra_enc)
    meses, val, anual, prelim = leer_mensual(filas, ruta, cols, ih, ic)
    meses, val = recortar(meses, val)
    for k in cols:
        verificar_anual(f"{nombre} {k}", meses, val[k], anual, k, "prom", 0.0005, fallas)
    fv = hoja(libro, *hoja_var)
    jh, jc = encabezado(fv, ruta)
    cv = columnas(fv, jh, jc, reglas, ruta, extra_enc)
    vm, vv, _, _ = leer_mensual(fv, ruta, cv, jh, jc)
    for k in cols:
        verificar_interanual(f"{nombre} {k}", meses, val[k], vm, vv[k], 12, fallas, avisos, set(prelim))
    return meses, {k: [r(v, 4) for v in val[k]] for k in cols}, prelim


def energia(fallas, avisos):
    arch = "energia/indice_consumo_energia_electrica_2017_2026.xlsx"
    reglas = [("general", lambda t: t.startswith("INDICE GENERAL")),
              ("domestico", lambda t: t.startswith("DOMESTICO")),
              ("comercial", lambda t: t == "GENERAL"),
              ("industria", lambda t: t.startswith("INDUSTRIA")),
              ("mineria", lambda t: t.startswith("MINERIA")),
              ("alumbrado", lambda t: t.startswith("ALUMBRADO")),
              ("otros", lambda t: t.startswith("OTROS"))]
    meses, series, prelim = indice_mensual(fallas, avisos, arch, reglas, "energía", ("INDICE",), ("12 MESES",))
    return {
        "metadata": meta("INE con datos de la Autoridad de Fiscalización de Electricidad y Tecnología Nuclear (AETN)", URL_INE + "servicios-basicos-cuadros-estadisticos/",
                         [arch], "índice de volumen (2017 = 100)", meses, series, prelim,
                         ["Índice de volumen del consumo de energía eléctrica con año de referencia 2017 (base móvil): 2017 = 100.",
                          "La categoría General corresponde a comercio, servicios, entidades sin fines de lucro y asociaciones civiles de mediana demanda.",
                          "Los consumidores no regulados se suman a la categoría según su actividad económica.",
                          "Cifras preliminares (p) del INE en los años marcados, sujetas a revisión."],
                         {"nombres": {"general": "Índice general", "domestico": "Doméstico", "comercial": "General (comercio y servicios)",
                                      "industria": "Industria", "mineria": "Minería", "alumbrado": "Alumbrado público", "otros": "Otros"}}),
        "meses": meses,
        "series": series,
    }


def transporte(fallas, avisos):
    arch = "transporte/indice_general_transporte_2017_2026.xlsx"
    reglas = [("general", lambda t: t.startswith("INDICE")),
              ("ferroviario", lambda t: t.startswith("FERROVIARIO")),
              ("carretero", lambda t: t.startswith("CARRETERO")),
              ("aereo", lambda t: t.startswith("AEREO")),
              ("ductos", lambda t: t.startswith("DUCTOS")),
              ("urbano", lambda t: t.startswith("URBANO"))]
    meses, series, prelim = indice_mensual(fallas, avisos, arch, reglas, "transporte", ("INDICE",), ("12 MESES",), extra_enc=1)
    return {
        "metadata": meta("INE — Índice general de transporte", URL_INE + "transportes/transporte-cuadros-estadisticos/",
                         [arch], "índice (2017 = 100)", meses, series, prelim,
                         ["Índice de volumen del transporte con año de referencia 2017 (base móvil): 2017 = 100.",
                          "Combina carga y pasajeros de cinco modalidades: ferroviario, carretero, aéreo, ductos (hidrocarburos) y urbano.",
                          "Cifras preliminares (p) del INE en los años marcados, sujetas a revisión."],
                         {"nombres": {"general": "Índice general", "ferroviario": "Ferroviario", "carretero": "Carretero",
                                      "aereo": "Aéreo", "ductos": "Ductos", "urbano": "Urbano"}}),
        "meses": meses,
        "series": series,
    }


# ── Carga ferroviaria ────────────────────────────────────────────────────────
def ferroviario(fallas):
    arch = "transporte/flujo_ferroviario_1999_2026.xlsx"
    ruta = DATOS / arch
    filas = hoja(hojas(ruta))
    ih, ic = encabezado(filas, ruta)
    # dos filas de encabezado: la red (celda combinada) y el tipo de servicio debajo
    red, cols = None, {}
    sub = filas[ih + 1]
    for c in range(max(len(filas[ih]), len(sub))):
        t = norma(filas[ih][c]) if c < len(filas[ih]) else ""
        if t.startswith("RED "):
            red = "andina" if "ANDINA" in t else "oriental" if "ORIENTAL" in t else None
            if red is None:
                raise Falla(f"{ruta.name}: red desconocida «{t}»")
        s = norma(sub[c]) if c < len(sub) else ""
        if red and s.startswith("CARGA"):
            if red in cols:
                raise Falla(f"{ruta.name}: dos columnas de carga para la red {red}")
            cols[red] = c
    if set(cols) != {"andina", "oriental"}:
        raise Falla(f"{ruta.name}: no se hallaron las columnas de carga de las dos redes ({cols})")
    meses, val, anual, prelim = leer_mensual(filas, ruta, cols, ih + 1, ic)
    meses, val = recortar(meses, val)
    for k in cols:
        verificar_anual(f"ferroviario {k}", meses, val[k], anual, k, "suma", 1.0, fallas)
    total = [None if a is None or o is None else a + o for a, o in zip(val["andina"], val["oriental"])]
    series = {"total": [r(v, 2) for v in total], "oriental": [r(v, 2) for v in val["oriental"]], "andina": [r(v, 2) for v in val["andina"]]}
    return {
        "metadata": meta("INE con datos de la Empresa Ferroviaria Oriental y la Empresa Ferroviaria Andina", URL_INE + "transportes/transporte-cuadros-estadisticos/",
                         [arch], "toneladas métricas", meses, series, prelim,
                         ["La carga total es la suma de la Red Oriental (Santa Cruz–Puerto Quijarro–Yacuiba) y la Red Andina (La Paz–Oruro–Potosí–Villazón y ramales).",
                          "Cifras preliminares (p) del INE en los años marcados, sujetas a revisión."],
                         {"nombres": {"total": "Total", "oriental": "Red Oriental", "andina": "Red Andina"}}),
        "meses": meses,
        "series": series,
    }


# ── Salida ───────────────────────────────────────────────────────────────────
def texto_json(obj) -> str:
    return json.dumps(obj, ensure_ascii=False, separators=(",", ":"))


def resumen(nombre, obj):
    claves = obj.get("meses") or obj.get("trimestres")
    m = obj["metadata"]
    filas = []
    for k, u in m["ultimo_por_serie"].items():
        filas.append(f"{k} hasta {u}")
    print(f"  {nombre:24} {claves[0]} → {m['ultimo']}  ({len(claves)} períodos) · " + " · ".join(filas))


def main() -> int:
    fallas, avisos = [], []
    try:
        salida = {
            "ind_cemento.json": cemento(fallas),
            "ind_permisos.json": permisos(fallas),
            "ind_manufactura.json": manufactura(fallas, avisos),
            "ind_energia.json": energia(fallas, avisos),
            "ind_transporte.json": transporte(fallas, avisos),
            "ind_ferroviario.json": ferroviario(fallas),
        }
    except Falla as e:
        print(f"✗ {e}\n  No se escribió nada.")
        return 1
    for nombre, obj in salida.items():
        no_retrocede(nombre, obj, fallas)
    if fallas:
        print(f"✗ {len(fallas)} verificación(es) no cierran; no se escribió nada:")
        for f in fallas[:30]:
            print("   ·", f)
        return 1
    print("Indicadores adelantados · identidades verificadas")
    for nombre, obj in salida.items():
        resumen(nombre, obj)
    for a in avisos:
        print("  ⚠", a)
    if "--revisar" in sys.argv:
        return 0
    DATA.mkdir(exist_ok=True)
    for nombre, obj in salida.items():
        ruta, txt = DATA / nombre, texto_json(obj)
        if ruta.exists() and ruta.read_text(encoding="utf-8") == txt:
            print(f"  = data/{nombre} sin cambios")
            continue
        ruta.write_text(txt, encoding="utf-8")
        print(f"  ✓ data/{nombre} ({ruta.stat().st_size // 1024} KB)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
