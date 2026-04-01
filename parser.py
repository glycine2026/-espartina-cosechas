import pandas as pd
from datetime import datetime
from logic.referencias import buscar_codigo_socio, buscar_codigo_especie, buscar_codigo_campania

COLUMNAS_ESPERADAS = [
    "Name", "Subelementos", "Estado", "Zona", "Creado por", "Prioridad",
    "Especie", "Campaña", "Establecimiento", "Localidad", "Localidad CP",
    "Titular", "Fecha carga *", "Fecha cupo *", "Contrato Albor",
    "Modelo CPE", "Destino", "Num CTG", "Observaciones",
]

def parsear_monday(uploaded_file) -> tuple[pd.DataFrame, dict]:
    raw = pd.read_excel(uploaded_file, header=None)

    # Encontrar la fila de headers (contiene "Name")
    header_row = None
    for i, row in raw.iterrows():
        if "Name" in row.values:
            header_row = i
            break
    if header_row is None:
        raise ValueError("No se encontró la fila de encabezados en el archivo.")

    df = pd.read_excel(uploaded_file, header=header_row)
    df = df[df["Num CTG"].notna() & (df["Num CTG"].astype(str).str.strip() != "")]
    df = df[df["Num CTG"].astype(str).str.strip() != "nan"].copy()
    df["Num CTG"] = df["Num CTG"].astype(str).str.strip()

    zonas = sorted(df["Zona"].dropna().unique().tolist())
    
    fechas = pd.to_datetime(df["Fecha carga *"], errors="coerce").dropna()
    fecha_min = fechas.min().strftime("%d/%m/%Y") if len(fechas) > 0 else "—"
    fecha_max = fechas.max().strftime("%d/%m/%Y") if len(fechas) > 0 else "—"

    meta = {
        "zonas": zonas,
        "fecha_min": fecha_min,
        "fecha_max": fecha_max,
        "total_ctg": len(df),
    }
    return df, meta

def precarga_ctg(row: pd.Series) -> dict:
    """Construye los datos pre-cargados para un CTG desde una fila de Monday."""
    titular = str(row.get("Titular", "") or "").strip()
    especie = str(row.get("Especie", "") or "").strip()
    campania = str(row.get("Campaña", "") or "").strip()
    
    fecha_raw = row.get("Fecha carga *")
    try:
        fecha = pd.to_datetime(fecha_raw).strftime("%Y-%m-%d")
    except Exception:
        fecha = ""

    cod_socio = buscar_codigo_socio(titular)
    cod_especie = buscar_codigo_especie(especie)
    cod_campania = buscar_codigo_campania(campania)

    return {
        "ctg": str(int(float(row.get("Num CTG", 0)))).strip() if str(row.get("Num CTG","")).replace(".","").isdigit() else str(row.get("Num CTG", "")).strip(),
        "cupo": str(row.get("Name", "")).strip(),
        "fecha": fecha,
        "titular_raw": titular,
        "especie_raw": especie,
        "establecimiento": str(row.get("Establecimiento", "") or "").strip(),
        "localidad": str(row.get("Localidad", "") or "").strip(),
        "zona": str(row.get("Zona", "") or "").strip(),
        "contrato_albor": str(row.get("Contrato Albor", "") or "").strip(),
        "destino_raw": str(row.get("Destino", "") or "").strip(),
        "modelo_cpe": str(row.get("Modelo CPE", "") or "").strip(),
        "precargado": {
            "Código socio": cod_socio,
            "Código campaña": cod_campania,
            "Código especie": cod_especie,
            "CTG": str(int(float(row.get("Num CTG", 0)))).strip() if str(row.get("Num CTG","")).replace(".","").isdigit() else str(row.get("Num CTG", "")).strip(),
            "Número comprobante contrato": str(row.get("Contrato Albor", "") or "").strip(),
            "Turno": str(row.get("Name", "")).strip(),
            "Tipo CPE": "E - Electrónica",
        }
    }
