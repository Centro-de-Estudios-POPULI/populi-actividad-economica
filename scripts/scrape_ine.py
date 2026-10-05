"""
scrape_ine.py — Monitor de Actividad Económica: descarga los cuadros del INE (nube.ine.gob.bo).

    python scripts/scrape_ine.py           # descarga todo lo que cambió
    python scripts/scrape_ine.py --check   # sólo muestra los enlaces que resolvería (no baja ni escribe)

Dos grupos:
  - PIB trimestral, año de referencia 2017 (23 cuadros). Los enlaces se RESUELVEN cada vez por el título
    del cuadro en https://www.ine.gob.bo/referencia2017/pib_trimestral.html: cuando el INE publica un
    trimestre nuevo sube archivos nuevos («2017 - 2026») con otro enlace. Si un título no aparece, se usa
    el último enlace conocido y se avisa.
  - Indicadores adelantados (enlaces fijos de nube.ine.gob.bo).

Un archivo sólo se reemplaza si lo descargado ES un Excel (xlsx = zip, xls = OLE2): una página de error
que responde 200 no pisa el dato. Los datos que leen los gráficos (data/*.json) los escriben después
scripts/datos_pib.py y scripts/datos_indicadores.py; este script no los toca.
"""
from __future__ import annotations

import hashlib
import html
import json
import re
import sys
import time
import unicodedata
from pathlib import Path

import requests

BASE = Path(__file__).resolve().parent.parent
DATOS = BASE / "datos"
HASHES = BASE / "scripts" / ".file_hashes.json"
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
NUBE = "https://nube.ine.gob.bo/index.php/s/{}/download"
PAGINA_PIB = "https://www.ine.gob.bo/referencia2017/pib_trimestral.html"
_IMC = "https://www.ine.gob.bo/index.php/estadisticas-economicas/industria-manufacturera-y-comercio/"
PAGINA_MANUF = _IMC + "cuadros-estadisticos-i/"                              # encuesta trimestral (2017 – año en curso)
PAGINA_CEMENTO = _IMC + "estadisticas-coyunturales-cuadros-estadisticos/"

# (ruta en datos/, título del cuadro en la página del INE sin «BOLIVIA:» ni años, último enlace conocido)
PIB = [
    ("pib_trimestral_base2017/corrientes/01_pib_actividad_corrientes.xlsx",
     "PRODUCTO INTERNO BRUTO POR GRUPOS DE ACTIVIDAD ECONOMICA SEGUN TRIMESTRE", "mvRnhfidKJeJfsv"),
    ("pib_trimestral_base2017/corrientes/02_pib_actividad_estructura_corrientes.xlsx",
     "ESTRUCTURA DEL PRODUCTO INTERNO BRUTO POR GRUPOS DE ACTIVIDAD ECONOMICA SEGUN TRIMESTRE", "q51Rr0Dp2XC5Slo"),
    ("pib_trimestral_base2017/corrientes/03_vab_actividad_corrientes.xlsx",
     "VALOR AGREGADO BRUTO A PRECIOS CORRIENTES POR GRUPOS DE ACTIVIDAD ECONOMICA SEGUN TRIMESTRE", "17KwlvhwPtpkEPH"),
    ("pib_trimestral_base2017/corrientes/04_vab_actividad_estructura_corrientes.xlsx",
     "ESTRUCTURA DEL VALOR AGREGADO BRUTO POR GRUPOS DE ACTIVIDAD ECONOMICA SEGUN TRIMESTRE", "A5DxGVcnu9FM0xR"),
    ("pib_trimestral_base2017/corrientes/05_pib_gasto_corrientes.xlsx",
     "PRODUCTO INTERNO BRUTO POR TIPO DE GASTO SEGUN TRIMESTRE", "pkGnaCmO6tpRvgK"),
    ("pib_trimestral_base2017/corrientes/06_pib_gasto_estructura_corrientes.xlsx",
     "ESTRUCTURA DEL PRODUCTO INTERNO BRUTO A PRECIOS CORRIENTES POR TIPO DE GASTO SEGUN TRIMESTRE", "hewOHNG3DGoL5hL"),
    ("pib_trimestral_base2017/encadenadas/07_pib_actividad_encadenadas.xlsx",
     "MEDIDAS DE VOLUMEN ENCADENADAS DEL PRODUCTO INTERNO BRUTO POR GRUPOS DE ACTIVIDAD ECONOMICA SEGUN TRIMESTRE", "1RFsK6XS7ZuMXaH"),
    ("pib_trimestral_base2017/encadenadas/08_pib_actividad_indices_encadenados.xlsx",
     "INDICES DE VOLUMEN ENCADENADOS DEL PRODUCTO INTERNO BRUTO POR GRUPOS DE ACTIVIDAD ECONOMICA SEGUN TRIMESTRE", "JVqHZtWjpifKc1D"),
    ("pib_trimestral_base2017/encadenadas/09_pib_actividad_variacion_encadenadas.xlsx",
     "VARIACION DE LAS MEDIDAS DE VOLUMEN ENCADENADAS DEL PRODUCTO INTERNO BRUTO POR GRUPOS DE ACTIVIDAD ECONOMICA SEGUN TRIMESTRE", "elvT17FSHM0hveX"),
    ("pib_trimestral_base2017/encadenadas/10_pib_actividad_variacion_acum_encadenadas.xlsx",
     "VARIACION ACUMULADA DE LAS MEDIDAS DE VOLUMEN ENCADENADAS DEL PRODUCTO INTERNO BRUTO POR ACTIVIDAD ECONOMICA", "pooLtCIFKjl0cQi"),
    ("pib_trimestral_base2017/encadenadas/11_pib_actividad_contribucion_encadenadas.xlsx",
     "CONTRIBUCION A LA VARIACION DE LAS MEDIDAS DE VOLUMEN ENCADENADAS DEL PRODUCTO INTERNO BRUTO POR ACTIVIDAD ECONOMICA", "uIj0pGEwtIacBj4"),
    ("pib_trimestral_base2017/encadenadas/12_pib_actividad_contribucion_acum_encadenadas.xlsx",
     "CONTRIBUCION A LA VARIACION ACUMULADA DE LAS MEDIDAS DE VOLUMEN ENCADENADAS DEL PRODUCTO INTERNO BRUTO POR ACTIVIDAD ECONOMICA", "e6te3KQv7cT3kXQ"),
    ("pib_trimestral_base2017/encadenadas/13_vab_actividad_encadenadas.xlsx",
     "MEDIDAS DE VOLUMEN ENCADENADAS DEL VALOR AGREGADO BRUTO POR GRUPOS DE ACTIVIDAD ECONOMICA SEGUN TRIMESTRE", "fEz6MqEOjVy7kjT"),
    ("pib_trimestral_base2017/encadenadas/14_vab_actividad_indices_encadenados.xlsx",
     "INDICES DE VOLUMEN ENCADENADOS DEL VALOR AGREGADO BRUTO POR GRUPOS DE ACTIVIDAD ECONOMICA SEGUN TRIMESTRE", "YVYnnfGBARHDwMC"),
    ("pib_trimestral_base2017/encadenadas/15_vab_actividad_variacion_encadenadas.xlsx",
     "VARIACION DE LAS MEDIDAS DE VOLUMEN ENCADENADAS DEL VALOR AGREGADO BRUTO POR GRUPOS DE ACTIVIDAD ECONOMICA SEGUN TRIMESTRE", "ySix1fqKIdaBhmr"),
    ("pib_trimestral_base2017/encadenadas/16_vab_actividad_variacion_acum_encadenadas.xlsx",
     "VARIACION ACUMULADA DE LAS MEDIDAS DE VOLUMEN ENCADENADAS DEL VALOR AGREGADO BRUTO POR GRUPOS DE ACTIVIDAD ECONOMICA SEGUN TRIMESTRE", "yw0mIIhxr76V6ry"),
    ("pib_trimestral_base2017/encadenadas/17_vab_actividad_contribucion_encadenadas.xlsx",
     "CONTRIBUCION A LA VARIACION DE LAS MEDIDAS DE VOLUMEN ENCADENADAS DEL VALOR AGREGADO BRUTO A SIMILAR PERIODO POR GRUPOS DE ACTIVIDAD ECONOMICA SEGUN TRIMESTRE", "y2FnxztQzaPkNl9"),
    ("pib_trimestral_base2017/encadenadas/18_vab_actividad_contribucion_acum_encadenadas.xlsx",
     "CONTRIBUCION A LA VARIACION ACUMULADA DE LAS MEDIDAS DE VOLUMEN ENCADENADAS DEL VALOR AGREGADO BRUTO A SIMILAR PERIODO POR ACTIVIDAD ECONOMICA SEGUN TRIMESTRE", "ClW58hxkrUuufel"),
    ("pib_trimestral_base2017/encadenadas/19_pib_gasto_encadenadas.xlsx",
     "VOLUMEN ENCADENADO DEL PRODUCTO INTERNO BRUTO POR TIPO DE GASTO SEGUN TRIMESTRE", "o4mdb7rYkI8hVu9"),
    ("pib_trimestral_base2017/encadenadas/20_pib_gasto_crecimiento_encadenadas.xlsx",
     "CRECIMIENTO DEL PRODUCTO INTERNO BRUTO A SIMILAR PERIODO POR TIPO DE GASTO SEGUN TRIMESTRE", "Llh3q00LGhYDnjw"),
    ("pib_trimestral_base2017/encadenadas/21_pib_gasto_crecimiento_acum_encadenadas.xlsx",
     "CRECIMIENTO ACUMULADO DEL PRODUCTO INTERNO BRUTO POR TIPO DE GASTO SEGUN TRIMESTRE", "EreOh4HJ20zGYlh"),
    ("pib_trimestral_base2017/encadenadas/22_pib_gasto_contribucion_encadenadas.xlsx",
     "CONTRIBUCIONES AL CRECIMIENTO DEL PRODUCTO INTERNO BRUTO POR TIPO DE GASTO SEGUN TRIMESTRE", "vhm2SPcnONzcbRJ"),
    ("pib_trimestral_base2017/encadenadas/23_pib_gasto_contribucion_acum_encadenadas.xlsx",
     "CONTRIBUCIONES AL CRECIMIENTO ACUMULADO DEL PRODUCTO INTERNO BRUTO POR TIPO DE GASTO SEGUN TRIMESTRE", "VkO702VwvFmZYlv"),
]

# (ruta en datos/, enlace de nube.ine.gob.bo, y —si se conoce— la página y el título por los que se resuelve)
# ⛔ La manufactura se congeló en 2025: el INE subió «2017 – 2026» con OTROS enlaces y los viejos siguieron
# respondiendo 200 con el archivo de mayo. Por eso, donde hay página, el enlace se busca por título cada vez.
# (Los nombres de archivo locales «_2017_2025» quedaron por historia: el contenido es el que trae el INE.)
INDICADORES = [
    ("indicadores_adelantados/cemento/produccion_cemento_depto_1991_2026.xlsx", "KO5oEVNveSsP3Qm",
     PAGINA_CEMENTO, "PRODUCCION DE CEMENTO POR DEPARTAMENTO SEGUN ANO Y MES"),
    ("indicadores_adelantados/cemento/ventas_cemento_depto_1991_2026.xlsx", "KkklD3dMnq2McoL",
     PAGINA_CEMENTO, "VENTAS DE CEMENTO POR DEPARTAMENTO SEGUN ANO Y MES"),
    ("indicadores_adelantados/cemento/consumo_cemento_depto_2000_2026.xlsx", "8Zv5IDKUwWjvsDI"),
    ("indicadores_adelantados/construccion/superficie_permisos_2008_2026.xlsx", "Ys6tEPQpLWYR4ba"),
    ("indicadores_adelantados/construccion/numero_permisos_2008_2026.xlsx", "jLb5wFW0JvKCqtR"),
    ("indicadores_adelantados/manufactura/indice_produccion_manufactura_2017_2025.xlsx", "POWVT4ZqvMZUIBl",
     PAGINA_MANUF, "INDICE DE VOLUMEN DE PRODUCCION DE LA INDUSTRIA MANUFACTURERA SEGUN TRIMESTRE"),
    ("indicadores_adelantados/manufactura/utilizacion_capacidad_instalada_2017_2025.xlsx", "rmFjGFy0nomAE9U",
     PAGINA_MANUF, "PORCENTAJE DE UTILIZACION DE LA CAPACIDAD PRODUCTIVA INSTALADA DE LA INDUSTRIA MANUFACTURERA SEGUN TRIMESTRE"),
    ("indicadores_adelantados/energia/indice_consumo_energia_electrica_2017_2026.xlsx", "yIcR99mxZZATQsM"),
    ("indicadores_adelantados/transporte/indice_general_transporte_2017_2026.xlsx", "WrlzgqwMRFOA6py"),
    ("indicadores_adelantados/transporte/flujo_ferroviario_1999_2026.xlsx", "LoswDmbMeSpofJH"),
]


def norma(s: str) -> str:
    # las rayas largas del INE («2017 – 2026», «Bolivia – Producción») pasan a guion ANTES de quitar los acentos:
    # si no, se pierden y el año final no se reconoce
    s = re.sub(r"[‐-―]", "-", s)
    s = unicodedata.normalize("NFD", s).encode("ascii", "ignore").decode()
    s = re.sub(r"\s+", " ", s).strip().upper()
    s = re.sub(r"^BOLIVIA\s*[:-]?\s*", "", s)
    s = re.sub(r",?\s*\d{4}\s*-?\s*\d{4}\s*$", "", s)   # «, 2017 - 2025» · «trimestre 2017 – 2026»
    return s.strip(" ,")


def enlaces(pagina: str) -> dict:
    """título normalizado → id de nube.ine.gob.bo, leído de una página del INE ({} si no se pudo)."""
    try:
        r = requests.get(pagina, headers=UA, timeout=60)
        r.raise_for_status()
        t = r.content.decode("utf-8", errors="replace")
    except Exception as e:
        print(f"  ⚠ no se pudo leer {pagina} ({e}): se usan los enlaces conocidos")
        return {}
    out = {}
    for m in re.finditer(r'<a[^>]*href="([^"]+)"[^>]*>(.*?)</a>', t, re.S):
        u = re.search(r"nube\.ine\.gob\.bo/index\.php/s/([A-Za-z0-9]+)", m.group(1))
        if not u:
            continue
        titulo = norma(html.unescape(re.sub(r"<[^>]+>", " ", m.group(2))))
        out.setdefault(titulo, u.group(1))      # el primero gana (la página repite algunos)
    return out


def es_excel(b: bytes) -> bool:
    return b[:4] == b"PK\x03\x04" or b[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def bajar(ident: str) -> bytes | None:
    for intento in range(3):
        try:
            r = requests.get(NUBE.format(ident), headers=UA, timeout=120)
            r.raise_for_status()
            if not es_excel(r.content):
                raise ValueError(f"no es un Excel ({len(r.content)} bytes, empieza con {r.content[:15]!r})")
            return r.content
        except Exception as e:
            if intento < 2:
                print(f"    intento {intento + 1}/3 falló: {e}; reintento en {[5, 15][intento]} s")
                time.sleep([5, 15][intento])
            else:
                print(f"    ✗ {e}")
    return None


def md5(b: bytes) -> str:
    return hashlib.md5(b).hexdigest()


def main() -> int:
    solo_ver = "--check" in sys.argv
    hashes = json.loads(HASHES.read_text(encoding="utf-8")) if HASHES.exists() else {}
    paginas = {}
    def resolver(ruta, conocido, pagina, titulo):
        if pagina not in paginas:
            paginas[pagina] = enlaces(pagina)
        ident = paginas[pagina].get(titulo)
        if not ident:
            print(f"  ⚠ «{titulo[:70]}…» no está en la página: se usa el enlace conocido")
        elif ident != conocido:
            print(f"  ↻ enlace nuevo para {Path(ruta).name}: {ident} (antes {conocido}): actualizar el conocido")
        return ident or conocido
    lista = [(ruta, resolver(ruta, conocido, PAGINA_PIB, titulo)) for ruta, titulo, conocido in PIB]
    for e in INDICADORES:
        lista.append((e[0], resolver(*e) if len(e) == 4 else e[1]))
    if solo_ver:
        for ruta, ident in lista:
            print(f"  {ident}  {ruta}")
        return 0
    cambiados, fallidos = [], []
    for ruta, ident in lista:
        destino = DATOS / ruta
        b = bajar(ident)
        if b is None:
            fallidos.append(ruta)
            continue
        h = md5(b)
        clave = ruta.replace("indicadores_adelantados/", "")   # claves históricas del archivo de huellas
        if destino.exists() and md5(destino.read_bytes()) == h:
            print(f"  = {ruta}")
        else:
            destino.parent.mkdir(parents=True, exist_ok=True)
            destino.write_bytes(b)
            cambiados.append(ruta)
            print(f"  ↑ {ruta} ({len(b) // 1024} KB)")
        hashes[clave] = h
    HASHES.write_text(json.dumps(hashes, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"\n{len(cambiados)} archivo(s) nuevos de {len(lista)}; {len(fallidos)} no se pudieron bajar")
    for f in fallidos:
        print(f"  ✗ {f}")
    # Un fallo de red de algunos archivos no es grave (la próxima corrida los trae); que no baje NINGUNO, sí.
    return 1 if len(fallidos) == len(lista) else 0


if __name__ == "__main__":
    sys.exit(main())
