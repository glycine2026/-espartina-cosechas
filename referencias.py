import pandas as pd
import re
from functools import lru_cache

TEMPLATE_PATH = "assets/ImportacionCosecha_template.xlsx"

@lru_cache(maxsize=1)
def cargar_referencias():
    df = pd.read_excel(TEMPLATE_PATH, sheet_name="Referencias", header=0)
    refs = {}
    for col in df.columns:
        vals = df[col].dropna().tolist()
        refs[col] = [str(v) for v in vals if str(v).strip()]
    return refs

def get_opciones(campo: str) -> list[str]:
    refs = cargar_referencias()
    return refs.get(campo, [])

def extraer_codigo(valor: str) -> str:
    """De '01 - Lartirigoyen y Cia. S.A.' extrae '01'"""
    if not valor:
        return ""
    match = re.match(r'^([^\s\-]+)', str(valor).strip())
    return match.group(1) if match else str(valor).strip()

def extraer_codigo_chofer(valor: str) -> str:
    """De 'Mauro Petti (20-29406866-0)' extrae '20-29406866-0'"""
    match = re.search(r'\(([^)]+)\)', str(valor))
    return match.group(1) if match else valor

def buscar_codigo_socio(titular: str) -> str:
    """Fuzzy match de titular de Monday hacia código socio de Referencias"""
    if not titular:
        return ""
    opciones = get_opciones("Código socio")
    titular_lower = titular.lower().strip()

    # Primero: match exacto por fragmento (titular contenido en opción o viceversa)
    for op in opciones:
        partes = op.split(" - ", 1)
        if len(partes) == 2:
            nombre = partes[1].lower()
            if titular_lower in nombre or nombre in titular_lower:
                return op

    # Segundo: match por palabras clave significativas (mínimo 3 chars)
    for op in opciones:
        partes = op.split(" - ", 1)
        if len(partes) == 2:
            nombre = partes[1].lower()
            palabras = [p for p in titular_lower.split() if len(p) >= 3]
            if palabras and any(p in nombre for p in palabras):
                return op
    return ""

def buscar_codigo_especie(especie_monday: str) -> str:
    """Mapea especie de Monday a código de especie"""
    MAPA = {
        "girasol comun": "14",
        "girasol común": "14",
        "girasol alto oleico": "07",
        "girasol confitero": "09",
        "soja": "04",
        "maiz": "02",
        "maíz": "02",
        "trigo": "03",
        "cebada": "13",
        "cebada forrajera": "13",
        "cebada cervecera": "13",
    }
    opciones = get_opciones("Código especie")
    especie_lower = especie_monday.lower().strip() if especie_monday else ""
    
    codigo = MAPA.get(especie_lower, "")
    if codigo:
        for op in opciones:
            if op.startswith(codigo + " -") or op.startswith(codigo + "-"):
                return op
    for op in opciones:
        if especie_lower in op.lower():
            return op
    return ""

def buscar_codigo_campania(campania: str) -> str:
    opciones = get_opciones("Código campaña")
    for op in opciones:
        if campania and campania.strip() in op:
            return op
    return opciones[0] if opciones else ""
