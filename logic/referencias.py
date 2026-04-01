import io
import re
import pathlib
import pandas as pd
from functools import lru_cache

BASE_DIR = pathlib.Path(__file__).resolve().parent.parent
TEMPLATE_PATH = BASE_DIR / "assets" / "template.xlsx"

def _cargar_bytes_template() -> bytes:
    """Lee el template desde disco. Falla con mensaje claro si no existe."""
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(
            f"No se encontró el template en {TEMPLATE_PATH}. "
            f"Archivos en assets/: {list((BASE_DIR / 'assets').iterdir()) if (BASE_DIR / 'assets').exists() else 'carpeta no existe'}"
        )
    return TEMPLATE_PATH.read_bytes()

@lru_cache(maxsize=1)
def cargar_referencias() -> dict:
    data = _cargar_bytes_template()
    df = pd.read_excel(io.BytesIO(data), sheet_name="Referencias", header=0, engine="openpyxl")
    refs = {}
    for col in df.columns:
        vals = [str(v) for v in df[col].dropna().tolist() if str(v).strip()]
        if vals:
            refs[col] = vals
    return refs

def get_template_bytes() -> bytes:
    return _cargar_bytes_template()

def get_opciones(campo: str) -> list:
    return cargar_referencias().get(campo, [])

def extraer_codigo(valor: str) -> str:
    if not valor:
        return ""
    m = re.match(r'^([^\s\-]+)', str(valor).strip())
    return m.group(1) if m else str(valor).strip()

def extraer_codigo_chofer(valor: str) -> str:
    m = re.search(r'\(([^)]+)\)', str(valor))
    return m.group(1) if m else valor

def buscar_codigo_socio(titular: str) -> str:
    if not titular:
        return ""
    opciones = get_opciones("Código socio")
    t = titular.lower().strip()
    for op in opciones:
        partes = op.split(" - ", 1)
        if len(partes) == 2:
            nombre = partes[1].lower()
            if t in nombre or nombre in t:
                return op
    for op in opciones:
        partes = op.split(" - ", 1)
        if len(partes) == 2:
            nombre = partes[1].lower()
            palabras = [p for p in t.split() if len(p) >= 3]
            if palabras and any(p in nombre for p in palabras):
                return op
    return ""

def buscar_codigo_especie(especie: str) -> str:
    MAPA = {
        "girasol comun": "14", "girasol común": "14",
        "girasol alto oleico": "07", "girasol confitero": "09",
        "soja": "04", "maiz": "02", "maíz": "02",
        "trigo": "03", "cebada": "13",
        "cebada forrajera": "13", "cebada cervecera": "13",
    }
    opciones = get_opciones("Código especie")
    e = especie.lower().strip() if especie else ""
    codigo = MAPA.get(e, "")
    if codigo:
        for op in opciones:
            if re.match(rf'^{codigo}\s*[-–]', op):
                return op
    for op in opciones:
        if e in op.lower():
            return op
    return ""

def buscar_codigo_campania(campania: str) -> str:
    opciones = get_opciones("Código campaña")
    for op in opciones:
        if campania and campania.strip() in op:
            return op
    return opciones[0] if opciones else ""
