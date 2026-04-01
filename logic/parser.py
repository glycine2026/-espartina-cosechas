import io
import pandas as pd
from logic.referencias import buscar_codigo_socio, buscar_codigo_especie, buscar_codigo_campania

def parsear_monday(uploaded_file) -> tuple[pd.DataFrame, dict]:
    # Leer todos los bytes de una vez para poder reutilizar el stream
    if hasattr(uploaded_file, "read"):
        contenido = uploaded_file.read()
    else:
        contenido = uploaded_file

    raw = pd.read_excel(io.BytesIO(contenido), header=None, engine="openpyxl")

    # Encontrar la fila de headers (contiene "Name")
    header_row = None
    for i, row in raw.iterrows():
        if "Name" in row.values:
            header_row = i
            break
    if header_row is None:
        raise ValueError("No se encontró la fila de encabezados en el archivo.")

    df = pd.read_excel(io.BytesIO(contenido), header=header_row, engine="openpyxl")
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


def _ctg_str(val) -> str:
    s = str(val).strip()
    if s.replace(".", "").isdigit():
        return str(int(float(s)))
    return s


def precarga_ctg(row: pd.Series) -> dict:
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
    ctg_val = _ctg_str(row.get("Num CTG", ""))

    return {
        "ctg": ctg_val,
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
            "CTG": ctg_val,
            "Número comprobante contrato": str(row.get("Contrato Albor", "") or "").strip(),
            "Turno": str(row.get("Name", "")).strip(),
            "Tipo CPE": "E - Electrónica",
        }
    }
