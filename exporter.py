import pandas as pd
import io
from openpyxl import load_workbook
from logic.referencias import extraer_codigo, extraer_codigo_chofer

TEMPLATE_PATH = "assets/ImportacionCosecha_template.xlsx"

CAMPOS_CON_CODIGO = [
    "Código tipo de comprobante", "Tipo CPE", "Numerador CPE", "Código socio",
    "Código campaña", "Código especie", "Tipo de Grano", "Código cultivo",
    "Código depósito origen", "Código depósito destino", "Número comprobante contrato",
    "Código contratista", "Código embolsador", "Código transportista",
    "Código intermediario flete", "Código entregador", "Código remitente",
    "Código destinatario/mercado a termino", "Código destino", "Tipo de Flete",
    "Código CATAC", "Código corredor", "Código extractor", "Código Personal Chofer",
    "Código Camión Personal", "Pagador Flete",
]

MAPA_CAMPOS = {
    "Cupo": "Turno",
    "Fecha": "Fecha",
    "Código socio": "Código socio",
    "Código campaña": "Código campaña",
    "Código especie": "Código especie",
    "Código cultivo": "Código cultivo",
    "Código depósito origen": "Código depósito origen",
    "Código depósito destino": "Código depósito destino",
    "Número comprobante contrato": "Número comprobante contrato",
    "Pagador Flete": "Pagador Flete",
    "CUIT Pagador Flete": "CUIT Pagador Flete",
    "Peso Origen Bruto": "Peso Origen Bruto",
    "Peso Origen Tara": "Peso Origen Tara",
    "Peso Origen Neto": "Peso Origen Neto",
    "% Humedad Origen": "% Humedad Origen",
    "% Humedad Destino": "% Humedad Destino",
    "Peso Destino Bruto": "Peso Destino Bruto",
    "Peso Destino Tara": "Peso Destino Tara",
    "Peso Destino Neto": "Peso Destino Neto",
    "Código embolsador": "Código embolsador",
    "Precio Embolsador por TN": "Precio Embolsador por TN",
    "Chofer (CUIT)": "Chofer (CUIT)",
    "Código Personal Chofer": "Código Personal Chofer",
    "Código transportista": "Código transportista",
    "Código intermediario flete": "Código intermediario flete",
    "Tipo CPE": "Tipo CPE",
    "Numerador CPE": "Numerador CPE",
    "Carta de Porte": "Carta de Porte",
    "CTG": "CTG",
    "Catac": "Código CATAC",
    "Distancia Planta": "Distancia Planta",
    "Tarifa Catac": "Tarifa Catac",
    "Coeficiente": "Coeficiente",
    "Aforado": "Aforado",
    "Código destino": "Código destino",
    "Código extractor": "Código extractor",
    "Precio Extractor por TN": "Precio Extractor por TN",
}

def _limpiar_valor(campo: str, valor) -> str:
    if valor is None or str(valor).strip() in ("", "nan", "None"):
        return ""
    valor_str = str(valor).strip()
    if campo in CAMPOS_CON_CODIGO:
        return extraer_codigo(valor_str)
    if campo == "Chofer (CUIT)":
        return extraer_codigo_chofer(valor_str)
    return valor_str

def generar_excel(ctgs_data: list[dict]) -> bytes:
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb["Cosechas a importar"]

    header_row = list(ws.iter_rows(min_row=1, max_row=1, values_only=True))[0]
    col_index = {str(h).strip(): i+1 for i, h in enumerate(header_row) if h}

    CAMPO_FIJO = {
        "Código tipo de comprobante": "COSI",
    }

    for row_num, ctg in enumerate(ctgs_data, start=2):
        campos_formulario = ctg.get("formulario", {})
        precargado = ctg.get("precargado", {})
        todos = {**precargado, **campos_formulario}

        for campo_form, campo_template in MAPA_CAMPOS.items():
            valor = todos.get(campo_form) or todos.get(campo_template, "")
            limpio = _limpiar_valor(campo_template, valor)
            if campo_template in col_index:
                ws.cell(row=row_num, column=col_index[campo_template], value=limpio)

        for campo, valor in CAMPO_FIJO.items():
            if campo in col_index:
                ws.cell(row=row_num, column=col_index[campo], value=valor)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()
