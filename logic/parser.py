import io
import pandas as pd
from logic.referencias import (
    buscar_codigo_socio, buscar_codigo_especie,
    buscar_codigo_campania, buscar_cuit_titular
)


def _to_bytes(f) -> bytes:
    if isinstance(f, (bytes, bytearray)):
        return bytes(f)
    if hasattr(f, "getvalue"):
        return f.getvalue()
    if hasattr(f, "read"):
        try: f.seek(0)
        except: pass
        return f.read()
    raise ValueError(f"Tipo no soportado: {type(f)}")


def _ctg_str(val) -> str:
    s = str(val).strip()
    try: return str(int(float(s)))
    except: return s


def parsear_monday(uploaded_file) -> tuple:
    raw = _to_bytes(uploaded_file)
    df_raw = pd.read_excel(io.BytesIO(raw), header=None, engine="openpyxl")

    header_row = None
    for i, row in df_raw.iterrows():
        if "Name" in row.values:
            header_row = i
            break
    if header_row is None:
        raise ValueError("No se encontró la fila de encabezados con 'Name'.")

    df = pd.read_excel(io.BytesIO(raw), header=header_row, engine="openpyxl")
    df = df[df["Num CTG"].notna()].copy()
    df = df[df["Num CTG"].astype(str).str.strip().isin(["", "nan"]) == False].copy()
    df["Num CTG"] = df["Num CTG"].astype(str).str.strip()

    zonas = sorted(df["Zona"].dropna().unique().tolist())
    fechas = pd.to_datetime(df["Fecha carga *"], errors="coerce").dropna()
    fecha_min = fechas.min().strftime("%d/%m/%Y") if len(fechas) else "—"
    fecha_max = fechas.max().strftime("%d/%m/%Y") if len(fechas) else "—"

    return df, {"zonas": zonas, "fecha_min": fecha_min,
                "fecha_max": fecha_max, "total_ctg": len(df)}


def precarga_ctg(row: pd.Series) -> dict:
    titular  = str(row.get("Titular",  "") or "").strip()
    especie  = str(row.get("Especie",  "") or "").strip()
    campania = str(row.get("Campaña",  "") or "").strip()
    ctg_val  = _ctg_str(row.get("Num CTG", ""))

    try:
        fecha = pd.to_datetime(row.get("Fecha carga *")).strftime("%Y-%m-%d")
    except:
        fecha = ""

    cod_socio = buscar_codigo_socio(titular)
    cuit      = buscar_cuit_titular(titular)

    return {
        "ctg":             ctg_val,
        "cupo":            str(row.get("Name", "")).strip(),
        "fecha":           fecha,
        "titular_raw":     titular,
        "especie_raw":     especie,
        "establecimiento": str(row.get("Establecimiento", "") or "").strip(),
        "localidad":       str(row.get("Localidad",       "") or "").strip(),
        "zona":            str(row.get("Zona",            "") or "").strip(),
        "contrato_albor":  str(row.get("Contrato Albor",  "") or "").strip(),
        "destino_raw":     str(row.get("Destino",         "") or "").strip(),
        "modelo_cpe":      str(row.get("Modelo CPE",      "") or "").strip(),
        "precargado": {
            "Código socio":                cod_socio,
            "Código campaña":              buscar_codigo_campania(campania),
            "Código especie":              buscar_codigo_especie(especie),
            "CTG":                         ctg_val,
            "Número comprobante contrato": str(row.get("Contrato Albor", "") or "").strip(),
            "Turno":                       str(row.get("Name", "")).strip(),
            "Tipo CPE":                    "E - Electrónica",
            "CUIT Pagador Flete":          cuit,
            "Pagador Flete CUIT display":  f"{titular} — {cuit}" if cuit else titular,
        }
    }
