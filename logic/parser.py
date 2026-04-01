import io
import pandas as pd
from logic.referencias import buscar_codigo_socio, buscar_codigo_especie, buscar_codigo_campania

def _to_bytes(uploaded_file) -> bytes:
    """Convierte cualquier tipo de input a bytes."""
    # Caso 1: ya son bytes crudos (lo que pasamos desde app.py)
    if isinstance(uploaded_file, (bytes, bytearray)):
        return bytes(uploaded_file)
    # Caso 2: BytesIO
    if isinstance(uploaded_file, io.BytesIO):
        uploaded_file.seek(0)
        data = uploaded_file.read()
        if not data:
            raise ValueError("BytesIO vacío")
        return data
    # Caso 3: objeto con getvalue (Streamlit UploadedFile)
    if hasattr(uploaded_file, "getvalue"):
        data = uploaded_file.getvalue()
        if not data:
            raise ValueError("getvalue() devolvió vacío")
        return data
    # Caso 4: file-like con read
    if hasattr(uploaded_file, "read"):
        try:
            uploaded_file.seek(0)
        except Exception:
            pass
        data = uploaded_file.read()
        if not data:
            raise ValueError("read() devolvió vacío")
        return data
    raise ValueError(f"Tipo no soportado: {type(uploaded_file)}")

def _ctg_str(val) -> str:
    s = str(val).strip()
    try:
        return str(int(float(s)))
    except Exception:
        return s

def parsear_monday(uploaded_file) -> tuple:
    raw_bytes = _to_bytes(uploaded_file)

    raw = pd.read_excel(io.BytesIO(raw_bytes), header=None, engine="openpyxl")

    header_row = None
    for i, row in raw.iterrows():
        if "Name" in row.values:
            header_row = i
            break
    if header_row is None:
        raise ValueError("No se encontró la fila de encabezados. Verificá que el archivo sea el exportado de Monday.")

    df = pd.read_excel(io.BytesIO(raw_bytes), header=header_row, engine="openpyxl")

    df = df[df["Num CTG"].notna()].copy()
    df = df[df["Num CTG"].astype(str).str.strip().str.replace(".", "", regex=False) != ""].copy()
    df = df[df["Num CTG"].astype(str).str.strip() != "nan"].copy()
    df["Num CTG"] = df["Num CTG"].astype(str).str.strip()

    zonas = sorted(df["Zona"].dropna().unique().tolist())
    fechas = pd.to_datetime(df["Fecha carga *"], errors="coerce").dropna()
    fecha_min = fechas.min().strftime("%d/%m/%Y") if len(fechas) > 0 else "—"
    fecha_max = fechas.max().strftime("%d/%m/%Y") if len(fechas) > 0 else "—"

    return df, {
        "zonas": zonas,
        "fecha_min": fecha_min,
        "fecha_max": fecha_max,
        "total_ctg": len(df),
    }

def precarga_ctg(row: pd.Series) -> dict:
    titular  = str(row.get("Titular",      "") or "").strip()
    especie  = str(row.get("Especie",       "") or "").strip()
    campania = str(row.get("Campaña",       "") or "").strip()

    try:
        fecha = pd.to_datetime(row.get("Fecha carga *")).strftime("%Y-%m-%d")
    except Exception:
        fecha = ""

    cod_socio   = buscar_codigo_socio(titular)
    cod_especie = buscar_codigo_especie(especie)
    cod_camp    = buscar_codigo_campania(campania)
    ctg_val     = _ctg_str(row.get("Num CTG", ""))

    return {
        "ctg":            ctg_val,
        "cupo":           str(row.get("Name",            "")).strip(),
        "fecha":          fecha,
        "titular_raw":    titular,
        "especie_raw":    especie,
        "establecimiento": str(row.get("Establecimiento", "") or "").strip(),
        "localidad":      str(row.get("Localidad",        "") or "").strip(),
        "zona":           str(row.get("Zona",             "") or "").strip(),
        "contrato_albor": str(row.get("Contrato Albor",   "") or "").strip(),
        "destino_raw":    str(row.get("Destino",          "") or "").strip(),
        "modelo_cpe":     str(row.get("Modelo CPE",       "") or "").strip(),
        "precargado": {
            "Código socio":               cod_socio,
            "Código campaña":             cod_camp,
            "Código especie":             cod_especie,
            "CTG":                        ctg_val,
            "Número comprobante contrato": str(row.get("Contrato Albor", "") or "").strip(),
            "Turno":                      str(row.get("Name", "")).strip(),
            "Tipo CPE":                   "E - Electrónica",
        }
    }
