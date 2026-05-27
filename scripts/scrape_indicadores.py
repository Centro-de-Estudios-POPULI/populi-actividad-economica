"""
scrape_indicadores.py — POPULI Monitor de Actividad Económica
Descarga Excel mensuales/trimestrales del INE Bolivia (nube.ine.gob.bo).
Si hay datos nuevos, regenera los HTML con build_indicadores_charts.py.

Uso:
  python scripts/scrape_indicadores.py            # descarga todo
  python scripts/scrape_indicadores.py --check    # solo verifica si hay cambios
"""

import hashlib
import json
import subprocess
import sys
import time
from pathlib import Path

import requests

BASE = Path(__file__).resolve().parent.parent
DATOS = BASE / "datos" / "indicadores_adelantados"
HASHES_FILE = BASE / "scripts" / ".file_hashes.json"

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
HEADERS = {"User-Agent": UA}

# ── Mapa de archivos: (URL, ruta local relativa a DATOS) ───────────────────
FILES = [
    # Cemento
    (
        "https://nube.ine.gob.bo/index.php/s/KO5oEVNveSsP3Qm/download",
        "cemento/produccion_cemento_depto_1991_2026.xlsx",
    ),
    (
        "https://nube.ine.gob.bo/index.php/s/KkklD3dMnq2McoL/download",
        "cemento/ventas_cemento_depto_1991_2026.xlsx",
    ),
    (
        "https://nube.ine.gob.bo/index.php/s/8Zv5IDKUwWjvsDI/download",
        "cemento/consumo_cemento_depto_2000_2026.xlsx",
    ),
    # Permisos de construcción
    (
        "https://nube.ine.gob.bo/index.php/s/Ys6tEPQpLWYR4ba/download",
        "construccion/superficie_permisos_2008_2026.xlsx",
    ),
    (
        "https://nube.ine.gob.bo/index.php/s/jLb5wFW0JvKCqtR/download",
        "construccion/numero_permisos_2008_2026.xlsx",
    ),
    # Manufactura (trimestral)
    (
        "https://nube.ine.gob.bo/index.php/s/DUuqeeGN6vFTHMm/download",
        "manufactura/indice_produccion_manufactura_2017_2025.xlsx",
    ),
    (
        "https://nube.ine.gob.bo/index.php/s/edMA2JVcXwNxByO/download",
        "manufactura/utilizacion_capacidad_instalada_2017_2025.xlsx",
    ),
    # Energía eléctrica
    (
        "https://nube.ine.gob.bo/index.php/s/yIcR99mxZZATQsM/download",
        "energia/indice_consumo_energia_electrica_2017_2026.xlsx",
    ),
    # Transporte — índices base 2017
    (
        "https://nube.ine.gob.bo/index.php/s/WrlzgqwMRFOA6py/download",
        "transporte/indice_general_transporte_2017_2026.xlsx",
    ),
    # Transporte — flujo ferroviario
    (
        "https://nube.ine.gob.bo/index.php/s/LoswDmbMeSpofJH/download",
        "transporte/flujo_ferroviario_1999_2026.xlsx",
    ),
]


def md5(path: Path) -> str:
    if not path.exists():
        return ""
    return hashlib.md5(path.read_bytes()).hexdigest()


def load_hashes() -> dict:
    if HASHES_FILE.exists():
        return json.loads(HASHES_FILE.read_text(encoding="utf-8"))
    return {}


def save_hashes(h: dict):
    HASHES_FILE.write_text(json.dumps(h, indent=2), encoding="utf-8")


def download_file(url: str, dest: Path) -> bool:
    dest.parent.mkdir(parents=True, exist_ok=True)
    for attempt in range(3):
        try:
            r = requests.get(url, headers=HEADERS, timeout=120, stream=True)
            r.raise_for_status()
            dest.write_bytes(r.content)
            size_kb = len(r.content) // 1024
            print(f"  OK: {dest.name} ({size_kb} KB)")
            return True
        except Exception as e:
            if attempt < 2:
                delay = [5, 15, 30][attempt]
                print(f"  Intento {attempt + 1}/3 falló: {e} — reintentando en {delay}s")
                time.sleep(delay)
            else:
                print(f"  ERROR tras 3 intentos: {dest.name} — {e}")
                return False
    return False


def main():
    check_only = "--check" in sys.argv

    old_hashes = load_hashes()
    new_hashes = {}
    changed = []

    print("=== Descargando indicadores del INE ===\n")

    for url, rel_path in FILES:
        dest = DATOS / rel_path
        old_hash = old_hashes.get(rel_path, "")

        if check_only:
            cur_hash = md5(dest)
            new_hashes[rel_path] = cur_hash
            continue

        print(f"Descargando {rel_path}...")
        old_local_hash = md5(dest)

        ok = download_file(url, dest)
        if not ok:
            new_hashes[rel_path] = old_hash
            continue

        new_hash = md5(dest)
        new_hashes[rel_path] = new_hash

        if new_hash != old_local_hash:
            changed.append(rel_path)
            print(f"  ↑ ACTUALIZADO (hash cambió)")
        else:
            print(f"  = Sin cambios")

    save_hashes(new_hashes)

    if check_only:
        print("Modo check: no se descargó nada.")
        return

    print(f"\n{'=' * 50}")
    print(f"Archivos actualizados: {len(changed)}/{len(FILES)}")

    if changed:
        print("\nArchivos con cambios:")
        for f in changed:
            print(f"  • {f}")

        print("\n=== Regenerando HTML embeds ===\n")
        build_script = BASE / "scripts" / "build_indicadores_charts.py"
        if build_script.exists():
            result = subprocess.run(
                [sys.executable, str(build_script)],
                cwd=str(BASE),
                capture_output=True,
                text=True,
            )
            print(result.stdout)
            if result.returncode != 0:
                print(f"ERROR en build: {result.stderr}")
                sys.exit(1)
        else:
            print(f"WARN: {build_script} no encontrado, saltando rebuild")
    else:
        print("No hay cambios — nada que regenerar.")


if __name__ == "__main__":
    main()
