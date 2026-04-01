import streamlit as st
import pandas as pd
from logic.parser import parsear_monday, precarga_ctg
from logic.referencias import get_opciones, extraer_codigo
from logic.exporter import generar_excel

st.set_page_config(
    page_title="Espartina — Importador de Cosechas",
    page_icon="🌾",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&display=swap');

html, body, [class*="css"], .stApp { font-family: 'Inter', sans-serif !important; }

.stApp { background: #f5f4f0; }

/* Header */
.esp-header {
    background: #1b3a1b; padding: 14px 28px;
    display: flex; align-items: center; justify-content: space-between;
    margin: -1rem -1rem 1.5rem -1rem;
    border-bottom: 3px solid #4a8c1c;
}
.esp-logo { color: #fff; font-size: 1.2rem; font-weight: 600; letter-spacing: .3px; }
.esp-logo span { color: #7cb84a; }
.esp-badge { background: #2d5a2d; color: #a8d47a; font-size: 0.68rem;
    padding: 4px 12px; border-radius: 20px; font-weight: 500; letter-spacing: .5px; }

/* Stepper */
.stepper { display:flex; background:#fff; border-radius:10px;
    border:1px solid #dddbd5; overflow:hidden; margin-bottom:1.2rem; }
.step { flex:1; padding:11px 16px; font-size:0.8rem; font-weight:500;
    color:#888; border-right:1px solid #eee; display:flex; align-items:center; gap:8px; }
.step:last-child { border-right:none; }
.step.active { background:#1b3a1b; color:#fff; }
.step.done { background:#eaf3de; color:#3B6D11; }
.step-num { width:20px; height:20px; border-radius:50%; background:rgba(0,0,0,.1);
    display:inline-flex; align-items:center; justify-content:center;
    font-size:0.68rem; font-weight:700; flex-shrink:0; }
.step.active .step-num { background:rgba(255,255,255,.2); }
.step.done .step-num { background:#3B6D11; color:#fff; }

/* Chips precarga */
.precarga-card { background:#f0f7e8; border:1px solid #c5ddb5; border-radius:10px;
    padding:12px 16px; margin-bottom:14px; }
.precarga-title { font-size:0.62rem; font-weight:700; text-transform:uppercase;
    letter-spacing:1.5px; color:#3B6D11; margin-bottom:8px; }
.chip { display:inline-flex; align-items:center; gap:5px; background:#fff;
    border:1px solid #c5ddb5; border-radius:6px; padding:3px 10px; margin:2px;
    font-size:0.75rem; }
.chip-key { font-weight:600; color:#1b3a1b; }
.chip-val { color:#555; }

/* Section headers */
.sec-header { font-size:0.62rem; font-weight:700; text-transform:uppercase;
    letter-spacing:1.5px; color:#1b3a1b; border-bottom:2px solid #c5ddb5;
    padding-bottom:5px; margin:16px 0 10px; }

/* Custom field labels */
.field-label { font-size:0.72rem; font-weight:600; color:#444;
    margin-bottom:3px; display:flex; align-items:center; gap:5px; }
.badge-monday { background:#e8f2e0; color:#2d5a2d; font-size:0.58rem;
    font-weight:700; padding:1px 6px; border-radius:3px;
    text-transform:uppercase; letter-spacing:.5px; }
.badge-req { color:#c0392b; font-size:0.8rem; }

/* Alert */
.alert-ok { background:#eaf3de; border:1px solid #b5d99e; border-radius:8px;
    padding:10px 16px; font-size:0.82rem; color:#2d5a2d;
    display:flex; align-items:center; gap:10px; margin-bottom:1rem; }

/* Stats */
div[data-testid="metric-container"] {
    background:#fff !important; border:1px solid #dddbd5 !important;
    border-radius:10px !important; }

/* Streamlit widget cleanup */
.stSelectbox > div > div, .stTextInput > div > div, .stNumberInput > div > div {
    border-radius: 8px !important; border-color: #dddbd5 !important; }
.stButton > button { border-radius: 8px !important; font-weight: 500 !important; }
div[data-testid="stExpander"] { border:1px solid #dddbd5 !important;
    border-radius:10px !important; background:#fff !important; margin-bottom:8px; }
.stProgress > div > div { background: #4a8c1c !important; }
</style>
""", unsafe_allow_html=True)


# ── Helpers ───────────────────────────────────────
def field_label(texto, monday=False, req=False):
    req_mark = " *" if req else ""
    monday_mark = " ↩ Monday" if monday else ""
    st.markdown(f"**{texto}{req_mark}**{monday_mark}")

def sec(titulo):
    st.markdown(f'<div class="sec-header">{titulo}</div>', unsafe_allow_html=True)

def buscar_contrato_fuzzy(contrato_excel, opciones):
    if not contrato_excel or not opciones:
        return 0
    ce = contrato_excel.strip()
    for j, c in enumerate(opciones):
        if ce == c.split(" - ")[0] or ce == c:
            return j
    for j, c in enumerate(opciones):
        if ce in c:
            return j
    palabras = ce.upper().split()
    best_j, best_score = 0, 0
    for j, c in enumerate(opciones):
        if not c:
            continue
        partes = c.split(" - ")[0].upper().split()
        score = sum(2 if a == b else (1 if len(a) >= 4 and a[:4] == b[:4] else 0)
                    for a, b in zip(palabras, partes))
        if score > best_score:
            best_score, best_j = score, j
    return best_j if best_score >= 2 else 0


# ── Session state ─────────────────────────────────
for k, v in {"paso": 1, "df_monday": None, "meta": None, "ctgs": [], "formularios": {}}.items():
    if k not in st.session_state:
        st.session_state[k] = v


# ── Header ────────────────────────────────────────
st.markdown("""
<div class="esp-header">
  <div class="esp-logo">Espartina <span>·</span> Importador de cosechas</div>
  <div class="esp-badge">Asistente de producción</div>
</div>""", unsafe_allow_html=True)


# ── Stepper ───────────────────────────────────────
paso = st.session_state.paso
def cls(n): return "step active" if paso==n else ("step done" if paso>n else "step")
def ico(n): return "✓" if paso>n else str(n)

st.markdown(f"""
<div class="stepper">
  <div class="{cls(1)}"><span class="step-num">{ico(1)}</span> Cargar archivo</div>
  <div class="{cls(2)}"><span class="step-num">{ico(2)}</span> Completar CTGs</div>
  <div class="{cls(3)}"><span class="step-num">{ico(3)}</span> Revisar y exportar</div>
</div>""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════
#  PASO 1 — Cargar archivo
# ══════════════════════════════════════════════════
if paso == 1:
    st.markdown("""
    <div style="border:2px dashed #4a8c1c;border-radius:12px;padding:2.5rem;
                text-align:center;background:#f9fbf7;margin-bottom:1.5rem">
      <div style="font-size:2rem;margin-bottom:.5rem">🌾</div>
      <div style="font-size:1rem;font-weight:600;color:#1b3a1b;margin-bottom:.3rem">
        Cargá el archivo de cupos exportado de Monday</div>
      <div style="font-size:0.82rem;color:#888">
        El sistema detecta automáticamente zonas, fechas y CTGs válidos</div>
    </div>""", unsafe_allow_html=True)

    archivo = st.file_uploader("Archivo Excel de Monday", type=["xlsx","xls"],
                               label_visibility="collapsed")

    if archivo:
        with st.spinner("Procesando..."):
            try:
                raw_bytes = archivo.getvalue()
                if len(raw_bytes) < 100:
                    raise ValueError(f"Archivo vacío ({len(raw_bytes)} bytes)")
                df, meta = parsear_monday(raw_bytes)
                st.session_state.df_monday = df
                st.session_state.meta = meta
                ctgs = []
                for _, row in df.iterrows():
                    ctg_data = precarga_ctg(row)
                    ctg_data["completado"] = False
                    ctgs.append(ctg_data)
                st.session_state.ctgs = ctgs
                st.session_state.formularios = {c["ctg"]: {} for c in ctgs}

                zonas_html = "".join(
                    f'<span style="background:#2d5a2d;color:#a8d47a;font-size:0.68rem;'
                    f'padding:2px 9px;border-radius:20px;margin:0 3px;font-weight:600">{z}</span>'
                    for z in meta["zonas"]
                )
                st.markdown(f"""
                <div class="alert-ok">
                  <span style="font-size:1.2rem">✓</span>
                  <div>Archivo cargado · Zonas: {zonas_html} ·
                  Fechas: {meta["fecha_min"]} → {meta["fecha_max"]} ·
                  <strong>{meta["total_ctg"]} CTGs</strong> detectados</div>
                </div>""", unsafe_allow_html=True)

                _, _, col3 = st.columns([3,1,1])
                with col3:
                    if st.button("Continuar →", type="primary", use_container_width=True):
                        st.session_state.paso = 2
                        st.rerun()
            except Exception as e:
                import traceback
                st.error(f"Error: {e}")
                st.code(traceback.format_exc())


# ══════════════════════════════════════════════════
#  PASO 2 — Completar CTGs
# ══════════════════════════════════════════════════
elif paso == 2:
    ctgs       = st.session_state.ctgs
    meta       = st.session_state.meta
    formularios= st.session_state.formularios
    total      = len(ctgs)
    completados= sum(1 for c in ctgs if c.get("completado"))
    pct        = int(completados/total*100) if total else 0

    zonas_html = "".join(
        f'<span style="background:#2d5a2d;color:#a8d47a;font-size:0.68rem;'
        f'padding:2px 9px;border-radius:20px;margin:0 3px;font-weight:600">{z}</span>'
        for z in meta["zonas"]
    )
    st.markdown(f"""
    <div class="alert-ok">
      <span style="font-size:1.1rem">📋</span>
      <div>Zonas: {zonas_html} · Fechas: {meta["fecha_min"]} → {meta["fecha_max"]}</div>
    </div>""", unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total CTGs", total)
    c2.metric("Completados", completados)
    c3.metric("Pendientes", total - completados)
    c4.metric("Progreso", f"{pct}%")
    st.progress(pct/100)
    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    # ── Loop CTGs ──
    for i, ctg in enumerate(ctgs):
        ctg_id   = ctg["ctg"]
        pre      = ctg["precargado"]
        form_key = ctg_id
        completado = ctg.get("completado", False)

        ico_ctg = "✅" if completado else "📋"
        titulo  = f"{ico_ctg} {ctg_id} · {ctg['especie_raw']} — {ctg['establecimiento']}, {ctg['localidad']}"

        with st.expander(titulo, expanded=(i == 0 and not completado)):

            # Tarjeta precarga
            pre_items = [
                ("Cupo",     pre.get("Turno","")),
                ("Titular",  ctg["titular_raw"]),
                ("Socio",    pre.get("Código socio","")),
                ("Especie",  pre.get("Código especie","")),
                ("Campaña",  pre.get("Código campaña","")),
                ("Contrato", pre.get("Número comprobante contrato","")),
                ("Fecha",    ctg.get("fecha","")),
                ("CTG",      pre.get("CTG","")),
            ]
            chips = "".join(
                f'<span class="chip"><span class="chip-key">{k}</span>'
                f'<span class="chip-val">{v}</span></span>'
                for k,v in pre_items if v
            )
            st.markdown(f"""
            <div class="precarga-card">
              <div class="precarga-title">Datos precargados desde Monday</div>
              <div>{chips}</div>
            </div>""", unsafe_allow_html=True)

            # ── Comprobante ──
            sec("Comprobante e identificación")

            c1, c2, c3 = st.columns(3)
            with c1:
                field_label("Cupo / turno", monday=True)
                st.text_input("_cupo", value=pre.get("Turno",""),
                              key=f"{form_key}_cupo", disabled=True,
                              label_visibility="collapsed")
            with c2:
                field_label("Fecha de carga", monday=True)
                formularios[form_key]["Fecha"] = st.text_input(
                    "_fecha", value=ctg.get("fecha",""),
                    key=f"{form_key}_fecha", label_visibility="collapsed")
            with c3:
                field_label("CTG", monday=True)
                st.text_input("_ctg", value=ctg_id, key=f"{form_key}_ctg_val",
                              disabled=True, label_visibility="collapsed")
                formularios[form_key]["CTG"] = ctg_id

            c4, c5, c6 = st.columns(3)
            with c4:
                field_label("Código socio", monday=True, req=True)
                opts = [""] + get_opciones("Código socio")
                idx  = next((j for j,o in enumerate(opts) if pre.get("Código socio","") and pre.get("Código socio","") in o), 0)
                formularios[form_key]["Código socio"] = st.selectbox(
                    "_socio", opts, index=idx, key=f"{form_key}_socio",
                    label_visibility="collapsed")
            with c5:
                field_label("Código campaña", monday=True, req=True)
                opts = [""] + get_opciones("Código campaña")
                idx  = next((j for j,o in enumerate(opts) if pre.get("Código campaña","") and pre.get("Código campaña","") in o), 0)
                formularios[form_key]["Código campaña"] = st.selectbox(
                    "_camp", opts, index=idx, key=f"{form_key}_camp",
                    label_visibility="collapsed")
            with c6:
                field_label("Código especie", monday=True, req=True)
                opts = [""] + get_opciones("Código especie")
                idx  = next((j for j,o in enumerate(opts) if pre.get("Código especie","") and pre.get("Código especie","") in o), 0)
                formularios[form_key]["Código especie"] = st.selectbox(
                    "_esp", opts, index=idx, key=f"{form_key}_esp",
                    label_visibility="collapsed")

            c7, c8, c9 = st.columns(3)
            with c7:
                field_label("Código cultivo", req=True)
                opts = [""] + get_opciones("Código cultivo")
                formularios[form_key]["Código cultivo"] = st.selectbox(
                    "_cult", opts, key=f"{form_key}_cult",
                    label_visibility="collapsed")
            with c8:
                field_label("Nro. comprobante contrato", monday=True, req=True)
                contrato_pre = pre.get("Número comprobante contrato","")
                opts_c = [""] + get_opciones("Código contrato")
                idx_c  = buscar_contrato_fuzzy(contrato_pre, opts_c)
                formularios[form_key]["Número comprobante contrato"] = st.selectbox(
                    "_contrato", opts_c, index=idx_c, key=f"{form_key}_contrato",
                    label_visibility="collapsed")
                if idx_c > 0 and contrato_pre:
                    st.markdown(
                        f'<div style="font-size:0.68rem;color:#3B6D11;margin-top:2px">'
                        f'↑ Monday: <code>{contrato_pre}</code></div>',
                        unsafe_allow_html=True)
            with c9:
                st.write("")

            # ── Flete ──
            sec("Flete y pagador")
            c1, c2, c3 = st.columns(3)
            with c1:
                field_label("Pagador flete", monday=True, req=True)
                opts = [""] + get_opciones("Código socio")
                idx  = next((j for j,o in enumerate(opts) if pre.get("Código socio","") and pre.get("Código socio","") in o), 0)
                formularios[form_key]["Pagador Flete"] = st.selectbox(
                    "_pagador", opts, index=idx, key=f"{form_key}_pagador",
                    label_visibility="collapsed")
            with c2:
                field_label("CUIT pagador flete")
                formularios[form_key]["CUIT Pagador Flete"] = st.text_input(
                    "_cuit", key=f"{form_key}_cuit_pag", label_visibility="collapsed")
            with c3:
                field_label("Tipo de flete")
                opts_tf = get_opciones("Tipo de Flete")
                idx_tf  = next((j for j,o in enumerate(opts_tf) if "Tercero" in o), 0)
                formularios[form_key]["Tipo de Flete"] = st.selectbox(
                    "_flete", opts_tf, index=idx_tf, key=f"{form_key}_tipo_flete",
                    label_visibility="collapsed")

            # ── Pesos origen ──
            sec("Pesos y humedad — origen")
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                field_label("Peso bruto (kg)", req=True)
                formularios[form_key]["Peso Origen Bruto"] = st.number_input(
                    "_pb", min_value=0, value=45000, step=100,
                    key=f"{form_key}_pb", label_visibility="collapsed")
            with c2:
                field_label("Tara (kg)", req=True)
                formularios[form_key]["Peso Origen Tara"] = st.number_input(
                    "_tara", min_value=0, value=15000, step=100,
                    key=f"{form_key}_tara", label_visibility="collapsed")
            with c3:
                field_label("Peso neto (kg)")
                neto = (formularios[form_key]["Peso Origen Bruto"] or 0) - \
                       (formularios[form_key]["Peso Origen Tara"] or 0)
                formularios[form_key]["Peso Origen Neto"] = neto
                st.number_input("_neto", value=neto, disabled=True,
                                key=f"{form_key}_neto", label_visibility="collapsed")
            with c4:
                field_label("% Humedad origen", req=True)
                formularios[form_key]["% Humedad Origen"] = st.number_input(
                    "_hum_o", min_value=0.0, max_value=40.0, step=0.1, format="%.1f",
                    key=f"{form_key}_hum_o", label_visibility="collapsed")

            # ── Transporte ──
            sec("Transporte")
            c1, c2, c3 = st.columns(3)
            with c1:
                field_label("Transportista", req=True)
                opts = [""] + get_opciones("Código transportista")
                formularios[form_key]["Código transportista"] = st.selectbox(
                    "_trans", opts, key=f"{form_key}_trans", label_visibility="collapsed")
            with c2:
                field_label("Chofer (CUIT)", req=True)
                opts = [""] + get_opciones("Chofer (CUIT)")
                formularios[form_key]["Chofer (CUIT)"] = st.selectbox(
                    "_chofer", opts, key=f"{form_key}_chofer", label_visibility="collapsed")
            with c3:
                field_label("Intermediario flete")
                opts = [""] + get_opciones("Código intermediario flete")
                formularios[form_key]["Código intermediario flete"] = st.selectbox(
                    "_int", opts, key=f"{form_key}_int", label_visibility="collapsed")

            c4, c5 = st.columns(2)
            with c4:
                field_label("Carta de porte", req=True)
                formularios[form_key]["Carta de Porte"] = st.text_input(
                    "_cpe", key=f"{form_key}_cpe", label_visibility="collapsed")
            with c5:
                field_label("% Humedad destino")
                formularios[form_key]["% Humedad Destino"] = st.number_input(
                    "_hum_d", min_value=0.0, max_value=40.0, step=0.1, format="%.1f",
                    key=f"{form_key}_hum_d", label_visibility="collapsed")

            # ── CATAC ──
            sec("CATAC y distancias")
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                field_label("CATAC", req=True)
                opts = [""] + get_opciones("Código CATAC")
                formularios[form_key]["Catac"] = st.selectbox(
                    "_catac", opts, key=f"{form_key}_catac", label_visibility="collapsed")
            with c2:
                field_label("Distancia planta (km)", req=True)
                formularios[form_key]["Distancia Planta"] = st.number_input(
                    "_dist", min_value=0, step=1, key=f"{form_key}_dist",
                    label_visibility="collapsed")
            with c3:
                field_label("Tarifa CATAC")
                formularios[form_key]["Tarifa Catac"] = st.number_input(
                    "_tarifa", min_value=0.0, step=0.01, format="%.2f",
                    key=f"{form_key}_tarifa", label_visibility="collapsed")
            with c4:
                field_label("Coeficiente")
                formularios[form_key]["Coeficiente"] = st.number_input(
                    "_coef", min_value=0.0, step=0.001, format="%.3f",
                    key=f"{form_key}_coef", label_visibility="collapsed")

            c5, c6 = st.columns(2)
            with c5:
                field_label("Aforado")
                formularios[form_key]["Aforado"] = st.number_input(
                    "_aforado", min_value=0.0, step=0.1,
                    key=f"{form_key}_aforado", label_visibility="collapsed")
            with c6:
                field_label("Código destino", req=True)
                opts = [""] + get_opciones("Código destino")
                formularios[form_key]["Código destino"] = st.selectbox(
                    "_dest", opts, key=f"{form_key}_dest", label_visibility="collapsed")

            # ── CPE ──
            sec("CPE")
            c1, c2 = st.columns(2)
            with c1:
                field_label("Tipo CPE")
                opts  = [""] + get_opciones("Tipo CPE")
                idx   = next((j for j,o in enumerate(opts) if "Electrónica" in o), 0)
                formularios[form_key]["Tipo CPE"] = st.selectbox(
                    "_tipo_cpe", opts, index=idx, key=f"{form_key}_tipo_cpe",
                    label_visibility="collapsed")
            with c2:
                field_label("Numerador CPE")
                opts = [""] + get_opciones("Numerador CPE")
                formularios[form_key]["Numerador CPE"] = st.selectbox(
                    "_num_cpe", opts, key=f"{form_key}_num_cpe",
                    label_visibility="collapsed")

            # ── Opcionales ──
            st.markdown(
                '<div style="font-size:0.62rem;font-weight:700;text-transform:uppercase;'
                'letter-spacing:1.5px;color:#aaa;border-bottom:1px dashed #ddd;'
                'padding-bottom:5px;margin:16px 0 6px">Campos opcionales</div>',
                unsafe_allow_html=True)
            with st.expander("Depósitos · Embolsado · Extracción"):
                st.caption("Completar solo si corresponde al tipo de operación")
                cd1, cd2 = st.columns(2)
                with cd1:
                    field_label("Depósito origen (solo extracción)")
                    opts = [""] + get_opciones("Código depósito")
                    formularios[form_key]["Código depósito origen"] = st.selectbox(
                        "_dep_o", opts, key=f"{form_key}_dep_o", label_visibility="collapsed")
                with cd2:
                    field_label("Depósito destino (solo embolsado/rechazo)")
                    opts = [""] + get_opciones("Código depósito")
                    formularios[form_key]["Código depósito destino"] = st.selectbox(
                        "_dep_d", opts, key=f"{form_key}_dep_d", label_visibility="collapsed")
                st.divider()
                ce1, ce2, ce3 = st.columns(3)
                with ce1:
                    field_label("Embolsador")
                    opts = [""] + get_opciones("Código embolsador")
                    formularios[form_key]["Código embolsador"] = st.selectbox(
                        "_emb", opts, key=f"{form_key}_emb", label_visibility="collapsed")
                with ce2:
                    field_label("Precio embolsador / TN")
                    formularios[form_key]["Precio Embolsador por TN"] = st.number_input(
                        "_p_emb", min_value=0.0, step=0.01,
                        key=f"{form_key}_precio_emb", label_visibility="collapsed")
                with ce3:
                    field_label("Extractor")
                    opts = [""] + get_opciones("Código extractor")
                    formularios[form_key]["Código extractor"] = st.selectbox(
                        "_ext", opts, key=f"{form_key}_ext", label_visibility="collapsed")
                field_label("Precio extractor / TN")
                formularios[form_key]["Precio Extractor por TN"] = st.number_input(
                    "_p_ext", min_value=0.0, step=0.01,
                    key=f"{form_key}_precio_ext", label_visibility="collapsed")

            # ── Guardar ──
            st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)
            cs, ci = st.columns([1,3])
            with cs:
                if st.button("Guardar CTG", key=f"save_{form_key}",
                             type="primary", use_container_width=True):
                    req = [formularios[form_key].get(f) for f in
                           ["Código socio","Código campaña","Código especie",
                            "Código destino","Carta de Porte","Catac"]]
                    if all(req):
                        st.session_state.ctgs[i]["completado"] = True
                        st.session_state.formularios[form_key] = formularios[form_key]
                        st.success("CTG guardado.")
                        st.rerun()
                    else:
                        st.error("Completá los campos obligatorios (*) antes de guardar.")
            with ci:
                if completado:
                    st.success("Este CTG está completo y se incluirá en la exportación.")

    st.divider()
    cb, _, cn = st.columns([1,3,1])
    with cb:
        if st.button("← Volver", use_container_width=True):
            st.session_state.paso = 1
            st.rerun()
    with cn:
        if completados == 0:
            st.button("Continuar →", disabled=True, use_container_width=True)
        else:
            if st.button(f"Continuar → ({completados}/{total})", type="primary",
                         use_container_width=True):
                st.session_state.paso = 3
                st.rerun()


# ══════════════════════════════════════════════════
#  PASO 3 — Revisar y exportar
# ══════════════════════════════════════════════════
elif paso == 3:
    ctgs        = st.session_state.ctgs
    formularios = st.session_state.formularios
    listos      = [c for c in ctgs if c.get("completado")]
    pendientes  = [c for c in ctgs if not c.get("completado")]

    st.markdown(f"""
    <div class="alert-ok">
      <span style="font-size:1.2rem">✓</span>
      <div><strong>{len(listos)} CTGs listos para exportar.</strong>
      {"&nbsp;·&nbsp; " + str(len(pendientes)) + " CTGs incompletos no se incluirán." if pendientes else ""}
      </div>
    </div>""", unsafe_allow_html=True)

    if pendientes:
        with st.expander(f"⚠️ {len(pendientes)} CTGs que NO se exportarán"):
            for c in pendientes:
                st.write(f"- `{c['ctg']}` · {c['especie_raw']} — {c['establecimiento']}")

    st.markdown('<div class="sec-header">Resumen de CTGs a exportar</div>',
                unsafe_allow_html=True)
    for c in listos:
        fk  = c["ctg"]
        f   = formularios.get(fk, {})
        pre = c.get("precargado", {})
        with st.expander(f"✅ {c['ctg']} · {c['especie_raw']} — {c['establecimiento']}"):
            r1, r2, r3 = st.columns(3)
            r1.write(f"**Socio:** {pre.get('Código socio','—')}")
            r1.write(f"**Especie:** {pre.get('Código especie','—')}")
            r1.write(f"**Campaña:** {pre.get('Código campaña','—')}")
            r2.write(f"**Peso bruto:** {f.get('Peso Origen Bruto','—')} kg")
            r2.write(f"**Peso neto:** {f.get('Peso Origen Neto','—')} kg")
            r2.write(f"**Humedad:** {f.get('% Humedad Origen','—')}%")
            r3.write(f"**Transportista:** {extraer_codigo(f.get('Código transportista','—'))}")
            r3.write(f"**Carta de porte:** {f.get('Carta de Porte','—')}")
            r3.write(f"**CATAC:** {extraer_codigo(f.get('Catac','—'))}")

    st.divider()
    cb2, _, ce2 = st.columns([1,2,1])
    with cb2:
        if st.button("← Volver a editar", use_container_width=True):
            st.session_state.paso = 2
            st.rerun()
    with ce2:
        if st.button("Exportar Excel", type="primary", use_container_width=True):
            payload = [{"precargado": c.get("precargado",{}),
                        "formulario": formularios.get(c["ctg"],{})}
                       for c in listos]
            try:
                excel_bytes = generar_excel(payload)
                st.download_button(
                    label="Descargar cosechas_importar.xlsx",
                    data=excel_bytes,
                    file_name="cosechas_importar.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
                st.success("Archivo generado. Hacé clic en el botón para descargarlo.")
            except Exception as e:
                st.error(f"Error al generar el Excel: {e}")
