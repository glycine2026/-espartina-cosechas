# Espartina — Importador de Cosechas

Plataforma web para que el Asistente de Producción complete y exporte el Excel de importación de cosechas a partir del archivo de cupos exportado de Monday.

## Flujo de uso

1. **Cargar archivo** — subir el Excel exportado de Monday (`AP_Detalle_de_cupos_*.xlsx`)
2. **Completar CTGs** — el sistema pre-carga datos de Monday y el usuario completa los campos faltantes con dropdowns de la hoja Referencias
3. **Exportar** — descarga el archivo `cosechas_importar.xlsx` con la estructura exacta del template de Albor

## Instalación local

```bash
git clone https://github.com/tu-usuario/espartina-cosechas.git
cd espartina-cosechas
pip install -r requirements.txt
streamlit run app.py
```

## Estructura de archivos

```
espartina-cosechas/
├── app.py                          # App principal Streamlit
├── requirements.txt
├── assets/
│   └── ImportacionCosecha_template.xlsx   # Template fijo (no modificar)
└── logic/
    ├── parser.py       # Lee y procesa el Excel de Monday
    ├── referencias.py  # Carga la hoja Referencias del template
    └── exporter.py     # Genera el Excel final de importación
```

## Deploy en Streamlit Community Cloud (gratuito)

1. Subir este repositorio a GitHub (puede ser privado)
2. Ir a [share.streamlit.io](https://share.streamlit.io)
3. Conectar el repo y configurar:
   - **Main file path:** `app.py`
   - **Python version:** 3.11
4. Click en **Deploy**

El archivo `ImportacionCosecha_template.xlsx` debe estar en la carpeta `assets/` del repositorio.

## Notas técnicas

- La hoja **Referencias** del template es la fuente de verdad para todos los dropdowns
- El mapeo de Titular → Código socio usa búsqueda por palabras clave
- El Excel final solo guarda los **códigos** (ej: `01`, no `01 - Lartirigoyen y Cia. S.A.`)
- Los campos calculados (Peso Neto = Bruto - Tara) se calculan automáticamente
