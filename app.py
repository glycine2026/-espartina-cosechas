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

# ─────────────────────────────────────────────
#  CSS — Identidad Espartina
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Serif+Display&display=swap');

:root {
    --verde-oscuro: #1b3a1b;
    --verde-medio: #2d5a2d;
    --verde-campo: #4a8c1c;
    --verde-claro: #7cb84a;
    --verde-bg: #f2f5ef;
    --verde-mint: #e8f2e0;
    --beige: #f7f5f0;
    --gris-borde: #dddbd5;
    --gris-texto: #6b6b6b;
    --blanco: #ffffff;
    --rojo-error: #c0392b;
    --amarillo-warn: #e67e22;
}

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif !important;
}

.stApp {
    background: var(--beige);
}

/* Header institucional */
.esp-header {
    background: var(--verde-oscuro);
    padding: 0 2rem;
    height: 60px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin: -1rem -1rem 2rem -1rem;
    border-bottom: 3px solid var(--verde-campo);
}
.esp-logo-text {
    font-family: 'DM Serif Display', serif;
    color: white;
    font-size: 1.4rem;
    letter-spacing: .5px;
}
.esp-logo-sub {
    color: var(--verde-claro);
    font-size: 0.75rem;
    letter-spacing: 1.5px;
    text-transform: uppercase;
    font-weight: 500;
}
.esp-header-badge {
    background: var(--verde-medio);
    color: var(--verde-claro);
    font-size: 0.7rem;
    padding: 4px 12px;
    border-radius: 20px;
    letter-spacing: .5px;
    font-weight: 500;
    text-transform: uppercase;
}

/* Stepper */
.stepper {
    display: flex;
    gap: 0;
    margin-bottom: 1.5rem;
    border-radius: 10px;
    overflow: hidden;
    border: 1px solid var(--gris-borde);
}
.step {
    flex: 1;
    padding: 12px 16px;
    font-size: 0.8rem;
    font-weight: 500;
    color: var(--gris-texto);
    background: white;
    border-right: 1px solid var(--gris-borde);
    display: flex;
    align-items: center;
    gap: 8px;
}
.step:last-child { border-right: none; }
.step.active { background: var(--verde-oscuro); color: white; }
.step.done { background: var(--verde-mint); color: var(--verde-campo); }
.step-num {
    width: 22px; height: 22px;
    border-radius: 50%;
    background: rgba(0,0,0,0.1);
    display: inline-flex;
    align-items: center;
    justify-content: center;
    font-size: 0.7rem;
    font-weight: 700;
    flex-shrink: 0;
}
.step.active .step-num { background: rgba(255,255,255,0.2); }
.step.done .step-num { background: var(--verde-campo); color: white; }

/* Cards */
.esp-card {
    background: white;
    border-radius: 12px;
    border: 1px solid var(--gris-borde);
    padding: 1.25rem 1.5rem;
    margin-bottom: 1rem;
}
.esp-card-header {
    font-size: 0.75rem;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 1px;
    color: var(--gris-texto);
    margin-bottom: 1rem;
    padding-bottom: 0.5rem;
    border-bottom: 1px solid var(--gris-borde);
}

/* Stats */
.stat-grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 10px;
    margin-bottom: 1.25rem;
}
.stat-box {
    background: var(--beige);
    border-radius: 10px;
    padding: 14px 16px;
    border: 1px solid var(--gris-borde);
}
.stat-label {
    font-size: 0.65rem;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 1px;
    color: var(--gris-texto);
    margin-bottom: 4px;
}
.stat-val { font-size: 1.6rem; font-weight: 300; color: var(--verde-oscuro); }
.stat-val.ok { color: var(--verde-campo); }
.stat-val.warn { color: var(--amarillo-warn); }
.stat-val.err { color: var(--rojo-error); }
.stat-sub { font-size: 0.7rem; color: var(--gris-texto); margin-top: 2px; }

/* Progress bar */
.prog-track {
    height: 4px;
    background: var(--gris-borde);
    border-radius: 2px;
    margin-top: 6px;
    overflow: hidden;
}
.prog-fill {
    height: 100%;
    background: var(--verde-campo);
    border-radius: 2px;
    transition: width .4s ease;
}

/* Alert banner */
.alert-success {
    background: var(--verde-mint);
    border: 1px solid #b5d99e;
    border-radius: 10px;
    padding: 10px 16px;
    font-size: 0.82rem;
    color: var(--verde-medio);
    margin-bottom: 1rem;
    display: flex;
    align-items: center;
    gap: 10px;
}
.alert-dot {
    width: 8px; height: 8px;
    background: var(--verde-campo);
    border-radius: 50%;
    flex-shrink: 0;
}
.zona-tag {
    display: inline-block;
    background: var(--verde-medio);
    color: var(--verde-claro);
    font-size: 0.68rem;
    padding: 2px 8px;
    border-radius: 20px;
    margin: 0 3px;
    font-weight: 600;
}

/* Upload zone */
.upload-zone {
    border: 2px dashed var(--verde-campo);
    border-radius: 12px;
    padding: 3rem;
    text-align: center;
    background: var(--verde-bg);
    margin-bottom: 1.5rem;
}
.upload-title {
    font-family: 'DM Serif Display', serif;
    font-size: 1.1rem;
    color: var(--verde-oscuro);
    margin-bottom: 0.5rem;
}
.upload-sub { font-size: 0.82rem; color: var(--gris-texto); }

/* CTG accordion */
.ctg-item {
    background: white;
    border: 1px solid var(--gris-borde);
    border-radius: 10px;
    margin-bottom: 8px;
    overflow: hidden;
}
.ctg-header-row {
    display: flex;
    align-items: center;
    gap: 12px;
    padding: 12px 16px;
    cursor: pointer;
}
.ctg-num {
    font-family: 'DM Mono', monospace;
    font-size: 0.82rem;
    font-weight: 600;
    color: var(--verde-oscuro);
    background: var(--verde-mint);
    padding: 3px 10px;
    border-radius: 6px;
    white-space: nowrap;
}
.ctg-desc {
    flex: 1;
    font-size: 0.85rem;
    color: #333;
}
.pill {
    font-size: 0.68rem;
    font-weight: 600;
    padding: 3px 10px;
    border-radius: 20px;
    white-space: nowrap;
}
.pill-ok { background: #eaf3de; color: #2d5a0e; }
.pill-warn { background: #fef3e2; color: #8c5a00; }
.pill-err { background: #fdecea; color: #8c1f1f; }
.pill-gray { background: #f1efe8; color: #5a5a5a; }

/* Form sections */
.form-section-title {
    font-size: 0.68rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 1.5px;
    color: var(--verde-oscuro);
    padding: 12px 0 6px;
    border-top: 1px solid var(--gris-borde);
    margin-top: 8px;
}
.precarga-box {
    background: var(--verde-mint);
    border-left: 3px solid var(--verde-campo);
    border-radius: 0 8px 8px 0;
    padding: 8px 12px;
    font-size: 0.78rem;
    color: var(--verde-medio);
    margin-bottom: 1rem;
    line-height: 1.7;
}
.precarga-box strong { color: var(--verde-oscuro); }

/* Bottom bar */
.bottom-bar {
    position: sticky;
    bottom: 0;
    background: white;
    border-top: 1px solid var(--gris-borde);
    padding: 12px 24px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    margin: 2rem -1rem -1rem -1rem;
    z-index: 100;
}

/* Override Streamlit elements */
.stButton > button {
    font-family: 'DM Sans', sans-serif !important;
    font-weight: 500 !important;
    border-radius: 8px !important;
    transition: all .15s !important;
}
.stSelectbox label, .stTextInput label, .stNumberInput label {
    font-size: 0.72rem !important;
    font-weight: 600 !important;
    text-transform: uppercase !important;
    letter-spacing: 0.8px !important;
    color: var(--gris-texto) !important;
}
div[data-testid="stExpander"] {
    border: 1px solid var(--gris-borde) !important;
    border-radius: 10px !important;
    background: white !important;
}
.streamlit-expanderHeader {
    font-family: 'DM Sans', sans-serif !important;
    font-weight: 500 !important;
}
div[data-testid="metric-container"] {
    background: var(--beige);
    border: 1px solid var(--gris-borde);
    border-radius: 10px;
    padding: 14px 16px;
}
.stProgress > div > div {
    background: var(--verde-campo) !important;
}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
#  Helper: fuzzy match contrato
# ─────────────────────────────────────────────
def buscar_contrato_fuzzy(contrato_excel: str, opciones: list) -> int:
    """Devuelve el índice (en opciones) del mejor match para el contrato del Excel."""
    if not contrato_excel or not opciones:
        return 0
    ce = contrato_excel.strip()
    # 1. Match exacto
    for j, c in enumerate(opciones):
        codigo = c.split(" - ")[0]
        if ce == codigo or ce == c:
            return j
    # 2. Contiene la cadena completa
    for j, c in enumerate(opciones):
        if ce in c:
            return j
    # 3. Match por palabras del prefijo del código
    palabras_ce = ce.upper().split()
    best_j, best_score = 0, 0
    for j, c in enumerate(opciones):
        if not c:
            continue
        codigo_c = c.split(" - ")[0].upper()
        palabras_c = codigo_c.split()
        score = 0
        for p_excel, p_ref in zip(palabras_ce, palabras_c):
            if p_excel == p_ref:
                score += 2
            elif len(p_excel) >= 4 and p_excel[:4] == p_ref[:4]:
                score += 1
            else:
                break
        if score > best_score:
            best_score, best_j = score, j
    return best_j if best_score >= 2 else 0


# ─────────────────────────────────────────────
#  Estado de sesión
# ─────────────────────────────────────────────
def init_state():
    defaults = {
        "paso": 1,
        "df_monday": None,
        "meta": None,
        "ctgs": [],
        "formularios": {},
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()


# ─────────────────────────────────────────────
#  Header
# ─────────────────────────────────────────────
st.markdown("""
<div class="esp-header">
  <div>
    <div class="esp-logo-text">Espartina</div>
    <div class="esp-logo-sub">Importador de cosechas</div>
  </div>
  <div class="esp-header-badge">Asistente de producción</div>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
#  Stepper
# ─────────────────────────────────────────────
paso = st.session_state.paso

def cls_step(n):
    if paso > n: return "step done"
    if paso == n: return "step active"
    return "step"

def icon_step(n):
    if paso > n: return "✓"
    return str(n)

st.markdown(f"""
<div class="stepper">
  <div class="{cls_step(1)}"><span class="step-num">{icon_step(1)}</span> Cargar archivo</div>
  <div class="{cls_step(2)}"><span class="step-num">{icon_step(2)}</span> Completar CTGs</div>
  <div class="{cls_step(3)}"><span class="step-num">{icon_step(3)}</span> Revisar y exportar</div>
</div>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════
#  PASO 1 — Cargar archivo
# ═══════════════════════════════════════════════
if paso == 1:
    st.markdown("""
    <div class="upload-zone">
      <div class="upload-title">🌾 Cargá el archivo de cupos de Monday</div>
      <div class="upload-sub">El sistema detecta automáticamente las zonas, fechas y CTGs válidos</div>
    </div>
    """, unsafe_allow_html=True)

    archivo = st.file_uploader(
        "Seleccioná el archivo Excel exportado de Monday",
        type=["xlsx", "xls"],
        label_visibility="collapsed",
    )

    if archivo:
        with st.spinner("Procesando archivo..."):
            try:
                # Leer bytes directamente del objeto Streamlit
                raw_bytes = archivo.getvalue()
                if len(raw_bytes) < 100:
                    raise ValueError(f"El archivo parece estar vacío ({len(raw_bytes)} bytes). Volvé a subirlo.")
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

                zonas_html = "".join(f'<span class="zona-tag">{z}</span>' for z in meta["zonas"])
                st.markdown(f"""
                <div class="alert-success">
                  <div class="alert-dot"></div>
                  <div>
                    <strong>Archivo cargado correctamente.</strong> &nbsp;
                    Zonas: {zonas_html} &nbsp;·&nbsp;
                    Fechas: {meta["fecha_min"]} → {meta["fecha_max"]} &nbsp;·&nbsp;
                    <strong>{meta["total_ctg"]} CTGs</strong> con número válido
                  </div>
                </div>
                """, unsafe_allow_html=True)

                col1, col2, col3 = st.columns([3, 1, 1])
                with col3:
                    if st.button("Continuar →", type="primary", use_container_width=True):
                        st.session_state.paso = 2
                        st.rerun()

            except Exception as e:
                import traceback
                st.error(f"Error: {e}")
                st.code(traceback.format_exc())


# ═══════════════════════════════════════════════
#  PASO 2 — Completar CTGs
# ═══════════════════════════════════════════════
elif paso == 2:
    ctgs = st.session_state.ctgs
    meta = st.session_state.meta
    formularios = st.session_state.formularios

    total = len(ctgs)
    completados = sum(1 for c in ctgs if c.get("completado"))
    pct = int(completados / total * 100) if total > 0 else 0

    # Summary stats
    zonas_html = "".join(f'<span class="zona-tag">{z}</span>' for z in meta["zonas"])
    st.markdown(f"""
    <div class="alert-success">
      <div class="alert-dot"></div>
      <div>Zonas: {zonas_html} &nbsp;·&nbsp; Fechas: {meta["fecha_min"]} → {meta["fecha_max"]}</div>
    </div>
    """, unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total CTGs", total)
    c2.metric("Completados", completados, delta=None)
    c3.metric("Pendientes", total - completados)
    c4.metric("Progreso", f"{pct}%")
    st.progress(pct / 100)

    st.markdown("---")

    # ── Loop de CTGs ──
    for i, ctg in enumerate(ctgs):
        ctg_id = ctg["ctg"]
        pre = ctg["precargado"]
        form_key = ctg_id

        completado = ctg.get("completado", False)
        pill = "🟢 Completo" if completado else "🟡 Pendiente"
        label = f"**{ctg_id}** &nbsp; {ctg['especie_raw']} — {ctg['establecimiento']}, {ctg['localidad']} &nbsp; `{pill}`"

        titulo_expander = (
            f"✅ {ctg_id} · {ctg['especie_raw']} — {ctg['establecimiento']}"
            if completado else
            f"📋 {ctg_id} · {ctg['especie_raw']} — {ctg['establecimiento']}, {ctg['localidad']}"
        )

        with st.expander(titulo_expander, expanded=(i == 0 and not completado)):

            # ── Banner: campos precargados del Excel (verde) vs manuales ──
            pre_items = [
                ("Cupo", pre.get("Turno", "—")),
                ("Titular", ctg["titular_raw"]),
                ("Cód. socio", pre.get("Código socio", "—")),
                ("Especie", pre.get("Código especie", "—")),
                ("Campaña", pre.get("Código campaña", "—")),
                ("Contrato", pre.get("Número comprobante contrato", "—")),
                ("Fecha", ctg.get("fecha", "—")),
                ("CTG", pre.get("CTG", "—")),
            ]
            pre_html = " &nbsp;·&nbsp; ".join(
                f'<span style="color:#1e3a1e"><strong>{k}:</strong></span> <span style="color:#2d5a2d">{v}</span>'
                for k, v in pre_items if v and v != "—"
            )
            st.markdown(f"""<div class="precarga-box">
            <span style="font-size:0.7rem;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:#3B6D11">
            ✦ Precargado automáticamente desde Monday</span><br/><br/>{pre_html}
            </div>""", unsafe_allow_html=True)

            # Helper para label con indicador de fuente
            def lbl(texto, fuente_monday=False, requerido=False):
                badge = ' <span style="font-size:0.6rem;background:#e8f2e0;color:#3B6D11;padding:1px 5px;border-radius:3px;font-weight:600">MONDAY</span>' if fuente_monday else ''
                req = ' *' if requerido else ''
                return texto + req + badge

            # ── SECCIÓN: Comprobante ──
            st.markdown('<div class="form-section-title">Comprobante e identificación</div>', unsafe_allow_html=True)
            col1, col2, col3 = st.columns(3)

            with col1:
                st.text_input(lbl("Cupo / turno", fuente_monday=True),
                              value=pre.get("Turno", ""), key=f"{form_key}_cupo", disabled=True)
            with col2:
                formularios[form_key]["Fecha"] = st.text_input(
                    lbl("Fecha", fuente_monday=True),
                    value=ctg.get("fecha", ""), key=f"{form_key}_fecha")
            with col3:
                st.text_input(lbl("CTG", fuente_monday=True),
                              value=ctg_id, key=f"{form_key}_ctg_val", disabled=True)
                formularios[form_key]["CTG"] = ctg_id

            col4, col5, col6 = st.columns(3)
            with col4:
                opts_socio = [""] + get_opciones("Código socio")
                idx_socio = next((j for j, o in enumerate(opts_socio) if pre.get("Código socio","") and pre.get("Código socio","") in o), 0)
                formularios[form_key]["Código socio"] = st.selectbox(
                    lbl("Código socio", fuente_monday=True, requerido=True),
                    opts_socio, index=idx_socio, key=f"{form_key}_socio")
            with col5:
                opts_camp = [""] + get_opciones("Código campaña")
                idx_camp = next((j for j, o in enumerate(opts_camp) if pre.get("Código campaña","") and pre.get("Código campaña","") in o), 0)
                formularios[form_key]["Código campaña"] = st.selectbox(
                    lbl("Código campaña", fuente_monday=True, requerido=True),
                    opts_camp, index=idx_camp, key=f"{form_key}_camp")
            with col6:
                opts_esp = [""] + get_opciones("Código especie")
                idx_esp = next((j for j, o in enumerate(opts_esp) if pre.get("Código especie","") and pre.get("Código especie","") in o), 0)
                formularios[form_key]["Código especie"] = st.selectbox(
                    lbl("Código especie", fuente_monday=True, requerido=True),
                    opts_esp, index=idx_esp, key=f"{form_key}_esp")

            col7, col8, col9 = st.columns(3)
            with col7:
                opts_cult = [""] + get_opciones("Código cultivo")
                formularios[form_key]["Código cultivo"] = st.selectbox(
                    "Código cultivo *", opts_cult, key=f"{form_key}_cult")
            with col8:
                # Contrato: fuzzy match entre valor del Excel y Referencias
                contrato_pre = pre.get("Número comprobante contrato", "")
                opts_contrato = [""] + get_opciones("Código contrato")
                idx_contrato = buscar_contrato_fuzzy(contrato_pre, opts_contrato)
                formularios[form_key]["Número comprobante contrato"] = st.selectbox(
                    lbl("Nro. comprobante contrato", fuente_monday=True, requerido=True),
                    opts_contrato, index=idx_contrato, key=f"{form_key}_contrato")
                if idx_contrato > 0 and contrato_pre:
                    st.caption(f"Excel: `{contrato_pre}` → match automático")
            with col9:
                st.write("")

            # ── SECCIÓN: Flete ──
            st.markdown('<div class="form-section-title">Flete y pagador</div>', unsafe_allow_html=True)
            col1, col2, col3 = st.columns(3)
            with col1:
                opts_pag = [""] + get_opciones("Código socio")
                idx_pag = next((j for j, o in enumerate(opts_pag) if pre.get("Código socio","") and pre.get("Código socio","") in o), 0)
                formularios[form_key]["Pagador Flete"] = st.selectbox(
                    lbl("Pagador flete", fuente_monday=True, requerido=True),
                    opts_pag, index=idx_pag, key=f"{form_key}_pagador")
            with col2:
                formularios[form_key]["CUIT Pagador Flete"] = st.text_input(
                    "CUIT pagador flete", key=f"{form_key}_cuit_pag")
            with col3:
                opts_tipo_flete = get_opciones("Tipo de Flete")
                # Default: T - Tercero
                idx_flete = next((j for j, o in enumerate(opts_tipo_flete) if "Tercero" in o or o.startswith("T")), 0)
                formularios[form_key]["Tipo de Flete"] = st.selectbox(
                    lbl("Tipo de flete", fuente_monday=False),
                    opts_tipo_flete, index=idx_flete, key=f"{form_key}_tipo_flete")

            # ── SECCIÓN: Pesos Origen ──
            st.markdown('<div class="form-section-title">Pesos y humedad — Origen</div>', unsafe_allow_html=True)
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                formularios[form_key]["Peso Origen Bruto"] = st.number_input(
                    "Peso bruto (kg) *", min_value=0, value=45000, step=100, key=f"{form_key}_pb")
            with col2:
                formularios[form_key]["Peso Origen Tara"] = st.number_input(
                    "Tara (kg) *", min_value=0, value=15000, step=100, key=f"{form_key}_tara")
            with col3:
                neto = (formularios[form_key]["Peso Origen Bruto"] or 0) - (formularios[form_key]["Peso Origen Tara"] or 0)
                formularios[form_key]["Peso Origen Neto"] = neto
                st.number_input("Peso neto (kg)", value=neto, disabled=True, key=f"{form_key}_neto")
            with col4:
                formularios[form_key]["% Humedad Origen"] = st.number_input(
                    "% Humedad origen *", min_value=0.0, max_value=40.0, step=0.1, format="%.1f",
                    key=f"{form_key}_hum_o")

            # ── SECCIÓN: Transporte ──
            st.markdown('<div class="form-section-title">Transporte</div>', unsafe_allow_html=True)
            col1, col2, col3 = st.columns(3)
            with col1:
                opts_trans = [""] + get_opciones("Código transportista")
                formularios[form_key]["Código transportista"] = st.selectbox(
                    "Transportista *", opts_trans, key=f"{form_key}_trans")
            with col2:
                opts_chofer = [""] + get_opciones("Chofer (CUIT)")
                formularios[form_key]["Chofer (CUIT)"] = st.selectbox(
                    "Chofer (CUIT) *", opts_chofer, key=f"{form_key}_chofer")
            with col3:
                opts_int = [""] + get_opciones("Código intermediario flete")
                formularios[form_key]["Código intermediario flete"] = st.selectbox(
                    "Intermediario flete", opts_int, key=f"{form_key}_int")

            col4, col5 = st.columns(2)
            with col4:
                formularios[form_key]["Carta de Porte"] = st.text_input(
                    "Carta de porte *", key=f"{form_key}_cpe")
            with col5:
                formularios[form_key]["% Humedad Destino"] = st.number_input(
                    "% Humedad destino", min_value=0.0, max_value=40.0, step=0.1, format="%.1f",
                    key=f"{form_key}_hum_d")

            # ── SECCIÓN: CATAC y distancias ──
            st.markdown('<div class="form-section-title">CATAC y distancias</div>', unsafe_allow_html=True)
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                opts_catac = [""] + get_opciones("Código CATAC")
                formularios[form_key]["Catac"] = st.selectbox(
                    "CATAC *", opts_catac, key=f"{form_key}_catac")
            with col2:
                formularios[form_key]["Distancia Planta"] = st.number_input(
                    "Distancia planta (km) *", min_value=0, step=1, key=f"{form_key}_dist")
            with col3:
                formularios[form_key]["Tarifa Catac"] = st.number_input(
                    "Tarifa CATAC", min_value=0.0, step=0.01, format="%.2f", key=f"{form_key}_tarifa")
            with col4:
                formularios[form_key]["Coeficiente"] = st.number_input(
                    "Coeficiente", min_value=0.0, step=0.001, format="%.3f", key=f"{form_key}_coef")

            col5, col6 = st.columns(2)
            with col5:
                formularios[form_key]["Aforado"] = st.number_input(
                    "Aforado", min_value=0.0, step=0.1, key=f"{form_key}_aforado")
            with col6:
                opts_dest = [""] + get_opciones("Código destino")
                formularios[form_key]["Código destino"] = st.selectbox(
                    "Código destino *", opts_dest, key=f"{form_key}_dest")

            # ── SECCIÓN: CPE ──
            st.markdown('<div class="form-section-title">CPE</div>', unsafe_allow_html=True)
            col1, col2 = st.columns(2)
            with col1:
                opts_tipo_cpe = [""] + get_opciones("Tipo CPE")
                idx_cpe = next((i for i, o in enumerate(opts_tipo_cpe) if "Electrónica" in o or "electronica" in o.lower()), 0)
                formularios[form_key]["Tipo CPE"] = st.selectbox(
                    "Tipo CPE", opts_tipo_cpe, index=idx_cpe, key=f"{form_key}_tipo_cpe")
            with col2:
                opts_num_cpe = [""] + get_opciones("Numerador CPE")
                formularios[form_key]["Numerador CPE"] = st.selectbox(
                    "Numerador CPE", opts_num_cpe, key=f"{form_key}_num_cpe")

            # ── SECCIÓN OPCIONAL: Embolsado / Extracción ──
            with st.expander("Opcionales: depósitos, embolsado, extracción, corredor"):
                st.markdown('<div style="font-size:0.75rem;color:#888;margin-bottom:8px">Solo completar si corresponde según el tipo de operación</div>', unsafe_allow_html=True)
                col_d1, col_d2 = st.columns(2)
                with col_d1:
                    opts_dep_o = [""] + get_opciones("Código depósito")
                    formularios[form_key]["Código depósito origen"] = st.selectbox(
                        "Código depósito origen (solo extracción)", opts_dep_o, key=f"{form_key}_dep_o")
                with col_d2:
                    opts_dep_d = [""] + get_opciones("Código depósito")
                    formularios[form_key]["Código depósito destino"] = st.selectbox(
                        "Código depósito destino (solo embolsado/rechazo)", opts_dep_d, key=f"{form_key}_dep_d")
                st.markdown("---")
                col1, col2, col3 = st.columns(3)
                with col1:
                    opts_emb = [""] + get_opciones("Código embolsador")
                    formularios[form_key]["Código embolsador"] = st.selectbox(
                        "Embolsador", opts_emb, key=f"{form_key}_emb")
                with col2:
                    formularios[form_key]["Precio Embolsador por TN"] = st.number_input(
                        "Precio embolsador / TN", min_value=0.0, step=0.01, key=f"{form_key}_precio_emb")
                with col3:
                    opts_ext = [""] + get_opciones("Código extractor")
                    formularios[form_key]["Código extractor"] = st.selectbox(
                        "Extractor", opts_ext, key=f"{form_key}_ext")
                formularios[form_key]["Precio Extractor por TN"] = st.number_input(
                    "Precio extractor / TN", min_value=0.0, step=0.01, key=f"{form_key}_precio_ext")

            # ── Botón guardar CTG ──
            st.markdown("---")
            col_save, col_info = st.columns([1, 3])
            with col_save:
                if st.button("✅ Guardar este CTG", key=f"save_{form_key}", type="primary", use_container_width=True):
                    campos_req = [
                        formularios[form_key].get("Código socio"),
                        formularios[form_key].get("Código campaña"),
                        formularios[form_key].get("Código especie"),
                        formularios[form_key].get("Código destino"),
                        formularios[form_key].get("Carta de Porte"),
                        formularios[form_key].get("Catac"),
                    ]
                    if all(c for c in campos_req):
                        st.session_state.ctgs[i]["completado"] = True
                        st.session_state.formularios[form_key] = formularios[form_key]
                        st.success("CTG guardado.")
                        st.rerun()
                    else:
                        st.error("Completá los campos obligatorios (*) antes de guardar.")
            with col_info:
                if completado:
                    st.success("Este CTG ya está completo y será incluido en la exportación.")

    # ── Botones de navegación ──
    st.markdown("---")
    col_back, col_mid, col_next = st.columns([1, 3, 1])
    with col_back:
        if st.button("← Volver", use_container_width=True):
            st.session_state.paso = 1
            st.rerun()
    with col_next:
        if completados == 0:
            st.button("Continuar →", disabled=True, use_container_width=True,
                      help="Completá al menos un CTG para continuar")
        else:
            if st.button(f"Continuar → ({completados}/{total})", type="primary", use_container_width=True):
                st.session_state.paso = 3
                st.rerun()


# ═══════════════════════════════════════════════
#  PASO 3 — Revisar y exportar
# ═══════════════════════════════════════════════
elif paso == 3:
    ctgs = st.session_state.ctgs
    formularios = st.session_state.formularios

    completados = [c for c in ctgs if c.get("completado")]
    pendientes = [c for c in ctgs if not c.get("completado")]

    st.markdown(f"""
    <div class="alert-success">
      <div class="alert-dot"></div>
      <div>
        <strong>{len(completados)} CTGs listos para exportar.</strong>
        {f"&nbsp;·&nbsp; {len(pendientes)} CTGs incompletos no serán incluidos." if pendientes else ""}
      </div>
    </div>
    """, unsafe_allow_html=True)

    if pendientes:
        with st.expander(f"⚠️ {len(pendientes)} CTGs que NO se exportarán"):
            for c in pendientes:
                st.write(f"- `{c['ctg']}` · {c['especie_raw']} — {c['establecimiento']}")

    st.markdown("### Resumen de CTGs a exportar")
    for c in completados:
        fk = c["ctg"]
        f = formularios.get(fk, {})
        pre = c.get("precargado", {})
        with st.expander(f"✅ {c['ctg']} · {c['especie_raw']} — {c['establecimiento']}"):
            col1, col2, col3 = st.columns(3)
            col1.write(f"**Socio:** {pre.get('Código socio', '—')}")
            col1.write(f"**Especie:** {pre.get('Código especie', '—')}")
            col1.write(f"**Campaña:** {pre.get('Código campaña', '—')}")
            col2.write(f"**Peso bruto:** {f.get('Peso Origen Bruto', '—')} kg")
            col2.write(f"**Peso neto:** {f.get('Peso Origen Neto', '—')} kg")
            col2.write(f"**Humedad:** {f.get('% Humedad Origen', '—')}%")
            col3.write(f"**Transportista:** {extraer_codigo(f.get('Código transportista', '—'))}")
            col3.write(f"**Carta de porte:** {f.get('Carta de Porte', '—')}")
            col3.write(f"**CATAC:** {extraer_codigo(f.get('Catac', '—'))}")

    st.markdown("---")
    col_back, col_mid, col_export = st.columns([1, 2, 1])
    with col_back:
        if st.button("← Volver a editar", use_container_width=True):
            st.session_state.paso = 2
            st.rerun()
    with col_export:
        if st.button("⬇️ Exportar Excel", type="primary", use_container_width=True):
            ctgs_para_exportar = []
            for c in completados:
                ctgs_para_exportar.append({
                    "precargado": c.get("precargado", {}),
                    "formulario": formularios.get(c["ctg"], {}),
                })
            try:
                excel_bytes = generar_excel(ctgs_para_exportar)
                st.download_button(
                    label="📥 Descargar archivo",
                    data=excel_bytes,
                    file_name="cosechas_importar.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
                st.success("¡Archivo generado! Hacé clic en 'Descargar archivo' para guardarlo.")
            except Exception as e:
                st.error(f"Error al generar el Excel: {e}")
