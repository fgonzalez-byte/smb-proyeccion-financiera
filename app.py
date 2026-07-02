"""
Dashboard de Proyección Financiera — Empresa de Factoring
Interfaz Streamlit con parámetros modificables, gráficos Plotly y exportación Excel.
"""
import os
import streamlit as st
import pandas as pd
from datetime import datetime

from model import (
    ProjectionParams, Override, PARAM_LABELS, PARAM_UNITS,
    generate_monthly_projection, aggregate_by_period,
)
from charts import (
    plot_portfolio_evolution, plot_income_vs_costs, plot_net_result,
    plot_margin_trend, plot_cost_breakdown, plot_waterfall_annual, plot_comparison,
)
from export import export_to_excel
from scenarios import list_scenarios, save_scenario, load_scenario, delete_scenario

EXCEL_FILE = "Balance 05-2026 v2.xlsx"

# ── Configuración de página ───────────────────────────────────────────────────

st.set_page_config(
    page_title="SMB · Proyección Financiera",
    page_icon="logo_smb.png",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Tema corporativo SMB: naranja #F47920 + gris ─────────────────────────────
st.markdown("""
<style>
/* ── Fondo global ───────────────────────────────────────────────────────── */
.stApp { background-color: #1C1C1C; }

/* ── Sidebar ────────────────────────────────────────────────────────────── */
section[data-testid="stSidebar"] > div:first-child {
    background: #242424;
    border-right: 2px solid #F47920;
}
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] span { color: #D0D0D0; }
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 { color: #F47920; }

/* número inputs y selects en sidebar */
section[data-testid="stSidebar"] input,
section[data-testid="stSidebar"] select {
    background-color: #2E2E2E !important;
    color: #F5F5F5 !important;
    border-color: #4A4A4A !important;
}

/* ── Separadores de sección en sidebar ──────────────────────────────────── */
.sb-section {
    color: #F47920;
    font-size: 0.75rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    border-bottom: 1px solid #F47920;
    padding-bottom: 3px;
    margin-top: 16px;
    margin-bottom: 6px;
}

/* ── KPI cards ──────────────────────────────────────────────────────────── */
.kpi-grid { display: flex; gap: 12px; flex-wrap: wrap; margin-bottom: 8px; }
.kpi-card {
    flex: 1; min-width: 160px;
    background: linear-gradient(135deg, #2E2E2E 0%, #222222 100%);
    border: 1px solid #F47920;
    border-radius: 10px;
    padding: 14px 18px;
}
.kpi-label   { color: #A8A8A8; font-size: 0.75rem; margin-bottom: 4px; }
.kpi-value   { color: #F5F5F5; font-size: 1.65rem; font-weight: 700; line-height: 1.1; }
.kpi-delta   { font-size: 0.78rem; margin-top: 4px; }
.kpi-pos     { color: #F47920; }
.kpi-neg     { color: #EF4444; }
.kpi-neutral { color: #A8A8A8; }

/* ── Override badge ─────────────────────────────────────────────────────── */
.ov-badge {
    background: #2A2A2A;
    border: 1px solid #F47920;
    border-radius: 6px;
    padding: 4px 8px;
    margin: 3px 0;
    font-size: 0.8rem;
    color: #D0D0D0;
}

/* ── Tabs ───────────────────────────────────────────────────────────────── */
button[data-baseweb="tab"] {
    font-size: 0.9rem;
    color: #A8A8A8 !important;
}
button[data-baseweb="tab"][aria-selected="true"] {
    color: #F47920 !important;
    border-bottom-color: #F47920 !important;
}

/* ── Botones ────────────────────────────────────────────────────────────── */
.stButton > button {
    border-radius: 6px;
    border-color: #F47920 !important;
    color: #F47920 !important;
}
.stButton > button:hover {
    background-color: #F47920 !important;
    color: #1C1C1C !important;
}

/* ── Tabla ──────────────────────────────────────────────────────────────── */
div[data-testid="stDataFrame"] { border-radius: 8px; overflow: hidden; }

/* ── Textos generales ───────────────────────────────────────────────────── */
h1, h2, h3 { color: #F47920 !important; }
p, li, span { color: #D0D0D0; }
</style>
""", unsafe_allow_html=True)


# ── Inicialización de estado ───────────────────────────────────────────────────

if "overrides" not in st.session_state:
    st.session_state.overrides: list = []
if "active_scenario" not in st.session_state:
    st.session_state.active_scenario: str = ""
if "_imported_data" not in st.session_state:
    st.session_state._imported_data: dict = {}

# Auto-cargar balance al inicio si aún no se ha importado y el archivo existe
if not st.session_state._imported_data and os.path.exists(EXCEL_FILE):
    try:
        from excel_loader import load_smb_params as _load_smb
        _auto = _load_smb(EXCEL_FILE)
        st.session_state._imported_data = _auto
        st.session_state["nb_portfolio"] = _auto["initial_portfolio"]
        st.session_state["nb_op_costs"]  = _auto["current_op_costs"]
        st.session_state["nb_rem"]       = _auto["current_remuneration"]
        st.session_state["nb_other"]     = _auto["other_expenses"]
        st.session_state["nb_rate"]      = _auto["placement_rate"]
        st.session_state["nb_fund"]      = _auto["funding_cost_rate"]
        st.session_state.active_scenario = "SMB Base 05-2026"
    except Exception:
        pass

# Valores por defecto de los widgets del sidebar (modificables por import)
_D = st.session_state._imported_data   # atajo


def _dv(key, fallback):
    """Retorna el valor importado si existe, si no el fallback."""
    return _D.get(key, fallback)


def kpi_card(label: str, value: str, delta: str = "", delta_type: str = "neutral") -> str:
    delta_html = f'<div class="kpi-delta kpi-{delta_type}">{delta}</div>' if delta else ""
    return f"""
    <div class="kpi-card">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{value}</div>
        {delta_html}
    </div>"""


def _param_bar(p, df_m) -> str:
    """Barra fija con parámetros clave; JS ajusta left al ancho real del sidebar."""
    m1  = df_m.iloc[0]["Resultado_neto"]
    m72 = df_m.iloc[-1]["Resultado_neto"]

    def _chip(label, value, color="#DEDEDE"):
        return (
            f'<div class="pbar-chip">'
            f'<div class="pbar-lbl">{label}</div>'
            f'<div class="pbar-val" style="color:{color};">{value}</div>'
            f'</div>'
        )

    chips = "".join([
        _chip("Cartera Fact.",  f"M${p.initial_portfolio:,.0f}"),
        _chip("Crec./año",      f"M${p.annual_portfolio_growth:,.0f}"),
        _chip("Tasa Coloc.",    f"{p.placement_rate:.2f}%"),
        _chip("Costo Fondo",    f"{p.funding_cost_rate:.2f}%"),
        _chip("Leasing",        f"{p.leasing_portfolio_uf:,.0f} UF"),
        _chip("Tasa Leas.",     f"{p.leasing_annual_rate:.1f}%/año"),
        _chip("IPC Proy.",      f"{p.monthly_ipc:.2f}%/mes"),
        _chip("Remuner.",       f"M${p.current_remuneration:,.0f}"),
        _chip("Costos Op.",     f"M${p.current_op_costs:,.0f}"),
        _chip("Prov. anual",    f"{p.provision_rate:.1f}%"),
        _chip("Imp. renta",     f"{p.tax_rate:.0f}%"),
        _chip("Res. Mes 1",  f"M${m1:+,.1f}",  "#F47920" if m1  >= 0 else "#EF4444"),
        _chip("Res. Mes 72", f"M${m72:+,.1f}", "#F47920" if m72 >= 0 else "#EF4444"),
    ])

    return f"""
<style>
.smb-pbar {{
    position: fixed;
    top: 3.75rem;
    left: 0; right: 0;        /* JS sobreescribe left al ancho real del sidebar */
    z-index: 200;
    background: rgba(20, 20, 20, 0.98);
    border-bottom: 2px solid #F47920;
    border-top: 1px solid #333;
    padding: 5px 18px 5px 16px;
    display: flex;
    align-items: center;
    overflow-x: auto;
    scrollbar-width: thin;
    scrollbar-color: #F47920 #1A1A1A;
    backdrop-filter: blur(10px);
    -webkit-backdrop-filter: blur(10px);
    box-shadow: 0 3px 16px rgba(0,0,0,0.70);
}}
.smb-pbar::-webkit-scrollbar {{ height: 4px; }}
.smb-pbar::-webkit-scrollbar-thumb {{ background: #F47920; border-radius: 2px; }}
.pbar-chip {{
    display: flex;
    flex-direction: column;
    align-items: center;
    min-width: 96px;
    padding: 3px 13px;
    border-right: 1px solid #303030;
    flex-shrink: 0;
}}
.pbar-chip:last-child {{ border-right: none; }}
.pbar-lbl {{
    font-size: 0.64rem;
    color: #6A6A6A;
    text-transform: uppercase;
    letter-spacing: 0.07em;
    white-space: nowrap;
    line-height: 1.4;
}}
.pbar-val {{
    font-size: 0.88rem;
    font-weight: 700;
    white-space: nowrap;
    line-height: 1.35;
}}
.main .block-container {{
    padding-top: 7rem !important;
}}
</style>
<div class="smb-pbar" id="smb-pbar-fixed">{chips}</div>
<script>
(function() {{
    function adjustBar() {{
        var bar = document.getElementById('smb-pbar-fixed');
        if (!bar) return;
        var sb = document.querySelector('[data-testid="stSidebar"]');
        if (sb) {{
            var r = sb.getBoundingClientRect().right;
            bar.style.left = (r > 20 ? r : 0) + 'px';
        }} else {{
            bar.style.left = '0px';
        }}
    }}
    adjustBar();
    [200, 600, 1200, 2500].forEach(function(ms) {{ setTimeout(adjustBar, ms); }});
    window.addEventListener('resize', adjustBar);
    var mo = new MutationObserver(adjustBar);
    mo.observe(document.documentElement, {{
        subtree: true, childList: true,
        attributes: true, attributeFilter: ['class','style','aria-expanded']
    }});
}})();
</script>
"""


# ── SIDEBAR ───────────────────────────────────────────────────────────────────

with st.sidebar:
    # Logo SMB
    if os.path.exists("logo_smb_blanco.png"):
        st.image("logo_smb_blanco.png", use_container_width=True)
    else:
        st.markdown("## **SMB** Servicios Financieros")
    st.markdown(
        '<div style="text-align:center;color:#A8A8A8;font-size:0.72rem;margin-top:-8px;'
        'margin-bottom:6px;letter-spacing:0.05em;">PROYECCIÓN FINANCIERA</div>',
        unsafe_allow_html=True,
    )
    if st.session_state.active_scenario:
        st.markdown(
            f'<div style="text-align:center;background:#2E2E2E;border:1px solid #F47920;'
            f'border-radius:6px;padding:4px 8px;font-size:0.75rem;color:#F47920;margin-bottom:4px;">'
            f'Escenario: <b>{st.session_state.active_scenario}</b></div>',
            unsafe_allow_html=True,
        )

    # ── Importar desde Excel ────────────────────────────────────────────────
    st.markdown('<p class="sb-section">📂 Importar desde Balance</p>', unsafe_allow_html=True)

    excel_exists = os.path.exists(EXCEL_FILE)
    if excel_exists:
        if st.button("📥 Reimportar Balance 05-2026", use_container_width=True, key="btn_import_excel"):
            try:
                from excel_loader import load_smb_params
                data = load_smb_params(EXCEL_FILE)
                st.session_state._imported_data = data
                _D = data
                st.session_state["nb_portfolio"]   = data["initial_portfolio"]
                st.session_state["nb_op_costs"]    = data["current_op_costs"]
                st.session_state["nb_rem"]         = data["current_remuneration"]
                st.session_state["nb_other"]       = data["other_expenses"]
                st.session_state["nb_rate"]        = data["placement_rate"]
                st.session_state["nb_fund"]        = data["funding_cost_rate"]
                st.session_state.active_scenario   = "SMB Base 05-2026"
                st.rerun()
            except Exception as e:
                st.error(f"Error al importar: {e}")

        if _D:
            with st.expander("Ver datos EERR", expanded=False):
                st.markdown(f"**EERR IFRS · {_D.get('_periodo','')}** (promedios 5 meses)")
                rows = [
                    ("Cartera factoring",      f"M$ {_D.get('_cartera_mm',0):,.0f}"),
                    ("---", "---"),
                    ("Ing. factoring (mayo)",  f"M$ {_D.get('_ing_fact_mayo_mm',0):,.1f}"),
                    ("Comisiones (mayo)",       f"M$ {_D.get('_com_fact_mayo_mm',0):,.1f}"),
                    ("Ingresos prom. fact.",    f"M$ {_D.get('_ing_fact_prom_mm',0):,.1f}/mes"),
                    ("Costo fondos (mayo)",    f"M$ {_D.get('_costo_fond_mayo_mm',0):,.1f}"),
                    ("Costo fondos prom.",      f"M$ {_D.get('_costo_fond_prom_mm',0):,.1f}/mes"),
                    ("---", "---"),
                    ("Tasa colocación",        f"{_D.get('placement_rate',0):.2f}% mensual"),
                    ("Costo fondo",            f"{_D.get('funding_cost_rate',0):.2f}% mensual"),
                    ("Spread neto",            f"{_D.get('_spread_pct',0):.2f}% mensual"),
                    ("---", "---"),
                    ("Remuneraciones prom.",   f"M$ {_D.get('_rem_prom_mm',0):,.1f}/mes"),
                    ("Costos op. recurrentes", f"M$ {_D.get('_op_rec_mm',0):,.1f}/mes"),
                    ("Otros no op. prom.",     f"M$ {_D.get('_otros_noop_mm',0):,.1f}/mes"),
                    ("Finiquitos prom.",       f"M$ {_D.get('_finiquitos_prom_mm',0):,.1f}/mes"),
                    ("---", "---"),
                    ("Resultado neto (mayo)",  f"M$ {_D.get('_res_neto_mayo_mm',0):,.1f}"),
                    ("Resultado neto (acum.)", f"M$ {_D.get('_res_neto_acum_mm',0):,.1f}"),
                ]
                for label, val in rows:
                    if label == "---":
                        st.markdown("---")
                    else:
                        st.markdown(
                            f'<div style="display:flex;justify-content:space-between;'
                            f'font-size:0.78rem;color:#94a3b8;margin:1px 0">'
                            f'<span>{label}</span><span style="color:#e2e8f0;font-weight:600">{val}</span></div>',
                            unsafe_allow_html=True,
                        )
    else:
        st.caption(f"'{EXCEL_FILE}' no encontrado en la carpeta del proyecto.")

    # ── Estado financiero base ──────────────────────────────────────────────
    st.markdown('<p class="sb-section">📁 Estado Financiero Base</p>', unsafe_allow_html=True)

    initial_portfolio = st.number_input(
        "Cartera inicial (M$)",
        min_value=0.0, value=_dv("initial_portfolio", 5018.0), step=100.0,
        help="Stock factoring inicial en millones de pesos CLP",
        key="nb_portfolio",
    )
    current_op_costs = st.number_input(
        "Costos operacionales base (M$/mes)",
        min_value=0.0, value=_dv("current_op_costs", 28.0), step=1.0,
        help="Costos recurrentes excl. remuneraciones y no operacionales",
        key="nb_op_costs",
    )
    current_remuneration = st.number_input(
        "Remuneraciones base (M$/mes)",
        min_value=0.0, value=_dv("current_remuneration", 57.0), step=1.0,
        key="nb_rem",
    )
    other_expenses = st.number_input(
        "Otros gastos no op. (M$/mes)",
        min_value=0.0, value=_dv("other_expenses", 11.0), step=1.0,
        help="Otros costos no operacionales (promedio mensual EERR)",
        key="nb_other",
    )

    # ── Parámetros de crecimiento ───────────────────────────────────────────
    st.markdown('<p class="sb-section">📈 Crecimiento de Cartera</p>', unsafe_allow_html=True)

    annual_portfolio_growth = st.number_input(
        "Crecimiento base (M$/año)",
        min_value=0.0, value=1000.0, step=100.0,
        help="Crecimiento mensual base = este valor ÷ 12. Los tramos de abajo lo reemplazan mes a mes.",
        key="nb_growth",
    )
    base_monthly = annual_portfolio_growth / 12.0
    st.caption(f"Base: **M${base_monthly:,.1f}/mes** — se reemplaza si agregas tramos abajo")

    # ── Plan de crecimiento por tramos ──────────────────────────────────────
    # Filtra los overrides de portfolio_growth para mostrar el plan actual
    _growth_ovs = [ov for ov in st.session_state.overrides if ov.get("param") == "portfolio_growth"]

    with st.expander(
        f"Tramos de crecimiento ({len(_growth_ovs)} definidos)" if _growth_ovs else "Definir tramos de crecimiento",
        expanded=bool(_growth_ovs),
    ):
        # Formulario para agregar tramo
        col_gm, col_gv = st.columns(2)
        with col_gm:
            tramo_mes = st.number_input("Desde mes #", min_value=1, max_value=72, value=1, step=1, key="tramo_mes")
        with col_gv:
            tramo_val = st.number_input("Crecer (M$/mes)", min_value=0.0, value=100.0, step=10.0, key="tramo_val")

        _yr_t = (int(tramo_mes) - 1) // 12 + 1
        _mo_t = (int(tramo_mes) - 1) % 12 + 1
        st.caption(f"Año {_yr_t}, Mes {_mo_t:02d} en adelante")

        if st.button("Agregar tramo", key="btn_tramo", use_container_width=True):
            st.session_state.overrides.append({
                "param":      "portfolio_growth",
                "from_month": int(tramo_mes),
                "value":      float(tramo_val),
                "note":       f"Tramo crecimiento M${tramo_val:.0f}/mes",
                "is_delta":   False,
            })
            st.rerun()

        # Mostrar tramos activos ordenados
        if _growth_ovs:
            st.markdown("**Tramos activos:**")
            tramos_sorted = sorted(_growth_ovs, key=lambda x: x["from_month"])
            for ov in tramos_sorted:
                yr_t = (ov["from_month"] - 1) // 12 + 1
                st.markdown(
                    f'<div style="display:flex;justify-content:space-between;'
                    f'background:#1e293b;border-radius:5px;padding:4px 8px;'
                    f'margin:2px 0;font-size:0.8rem;">'
                    f'<span style="color:#94a3b8;">Mes {ov["from_month"]} (Año {yr_t})</span>'
                    f'<span style="color:#F47920;font-weight:700;">M${ov["value"]:,.0f}/mes</span>'
                    f'</div>',
                    unsafe_allow_html=True,
                )
            if st.button("Limpiar tramos", key="btn_clear_tramos"):
                st.session_state.overrides = [
                    ov for ov in st.session_state.overrides
                    if ov.get("param") != "portfolio_growth"
                ]
                st.rerun()

        # Preview: tabla de crecimiento efectivo por año
        if _growth_ovs:
            st.markdown("**Preview crecimiento efectivo:**")
            tramos_sorted = sorted(_growth_ovs, key=lambda x: x["from_month"])
            rows_prev = []
            for yr in range(1, 7):
                m_mid = (yr - 1) * 12 + 6  # mes central del año
                # valor efectivo ese mes
                val = base_monthly
                for ov in tramos_sorted:
                    if ov["from_month"] <= m_mid:
                        val = ov["value"]
                rows_prev.append({"Año": f"Año {yr}", "Crecimiento/mes": f"M${val:,.0f}", "Crecimiento/año": f"M${val*12:,.0f}"})
            st.dataframe(rows_prev, hide_index=True, use_container_width=True)

    # ── Tasas financieras ───────────────────────────────────────────────────
    st.markdown('<p class="sb-section">💰 Tasas Financieras</p>', unsafe_allow_html=True)

    col_t1, col_t2 = st.columns(2)
    with col_t1:
        placement_rate = st.number_input(
            "Tasa coloc.\n(% mensual)",
            min_value=0.0, max_value=20.0,
            value=_dv("placement_rate", 2.0), step=0.01, format="%.2f",
            key="nb_rate",
        )
    with col_t2:
        funding_cost_rate = st.number_input(
            "Costo fondo\n(% mensual)",
            min_value=0.0, max_value=10.0,
            value=_dv("funding_cost_rate", 0.44), step=0.01, format="%.2f",
            key="nb_fund",
        )

    spread = placement_rate - funding_cost_rate
    st.caption(f"Spread neto factoring: **{spread:.2f}%** mensual · costo fondo sobre cartera total")

    # ── Leasing ─────────────────────────────────────────────────────────────
    st.markdown('<p class="sb-section">🏗️ Leasing (UF + IPC)</p>', unsafe_allow_html=True)

    leasing_portfolio_uf = st.number_input(
        "Stock inicial leasing (UF)",
        min_value=0.0, value=_dv("leasing_portfolio_uf", 69354.96), step=1000.0,
        format="%.2f", key="nb_leas_uf",
    )
    uf_value = st.number_input(
        "Valor UF al inicio (CLP)",
        min_value=0.0, value=_dv("uf_value", 39900.0), step=100.0,
        format="%.0f", key="nb_uf_val",
        help="Valor de la UF al inicio de la proyección (Mayo 2026)",
    )
    leasing_mm_ini = leasing_portfolio_uf * uf_value / 1_000_000
    st.caption(f"Stock leasing inicial ≈ **M${leasing_mm_ini:,.0f}**")

    col_l1, col_l2 = st.columns(2)
    with col_l1:
        leasing_annual_rate = st.number_input(
            "Tasa anual\nleasing (%)",
            min_value=0.0, max_value=50.0,
            value=_dv("leasing_annual_rate", 17.0), step=0.5, format="%.1f",
            key="nb_leas_rate",
        )
    with col_l2:
        monthly_ipc = st.number_input(
            "IPC mensual\nproyectado (%)",
            min_value=0.0, max_value=5.0,
            value=_dv("monthly_ipc", 0.35), step=0.05, format="%.2f",
            key="nb_ipc",
            help="Reajuste UF mensual (≈0.35% = 4.3% anual)",
        )
    leasing_annual_growth_mm = st.number_input(
        "Crecimiento anual leasing (M$)",
        min_value=0.0, value=_dv("leasing_annual_growth_mm", 200.0), step=50.0,
        key="nb_leas_growth",
        help="Nuevos contratos de leasing por año en M$ nominales",
    )
    leasing_avg_term = st.number_input(
        "Plazo promedio contratos (meses)",
        min_value=6, max_value=120, value=36, step=6,
        key="nb_leas_term",
        help="Plazo promedio de los contratos de leasing. SMB: 36 meses.",
    )
    # Calcula y muestra la tasa lineal efectiva resultante
    _r = leasing_annual_rate / 12 / 100
    _n = leasing_avg_term
    _cuota = _r / (1 - (1 + _r) ** (-_n)) if _r > 0 else 1/_n
    _total_int = (_cuota * _n - 1) * 100
    _sl_rate   = _total_int / _n
    st.caption(
        f"Tasa mensual: **{leasing_annual_rate/12:.3f}%** · "
        f"IPC anual equiv.: **{((1+monthly_ipc/100)**12-1)*100:.1f}%** · "
        f"Interés total contrato: **{_total_int:.1f}%** → lineal **{_sl_rate:.3f}%/mes**"
    )

    # ── Riesgo, provisiones e impuesto ─────────────────────────────────────
    st.markdown('<p class="sb-section">⚖️ Riesgo e Impuesto</p>', unsafe_allow_html=True)

    col_r1, col_r2 = st.columns(2)
    with col_r1:
        provision_rate = st.number_input(
            "Tasa prov.\n(% anual cartera)",
            min_value=0.0, max_value=10.0, value=1.5, step=0.1, format="%.1f",
            key="nb_prov",
            help="Gasto anual en provisiones por riesgo de crédito, sobre cartera total",
        )
    with col_r2:
        npl_rate = st.number_input(
            "NPL\n(% cartera morosa)",
            min_value=0.0, max_value=20.0, value=2.0, step=0.1, format="%.1f",
            key="nb_npl",
            help="% estimado de cartera en mora (informativo)",
        )
    initial_equity = st.number_input(
        "Patrimonio inicial (M$)",
        min_value=0.0, value=500.0, step=50.0,
        key="nb_equity",
        help="Capital propio de la empresa al inicio de la proyección",
    )
    tax_rate = st.number_input(
        "Impuesto renta (%)",
        min_value=0.0, max_value=40.0, value=27.0, step=1.0, format="%.0f",
        key="nb_tax",
        help="Tasa de impuesto corporativo Chile: 27%",
    )
    st.caption(
        f"Prov. mensual ≈ **M${initial_portfolio * provision_rate / 12 / 100:,.1f}** "
        f"· NPL estimado ≈ **M${initial_portfolio * npl_rate / 100:,.0f}**"
    )

    # ── Personal ────────────────────────────────────────────────────────────
    st.markdown('<p class="sb-section">👥 Personal y Costos</p>', unsafe_allow_html=True)

    annual_op_increment = st.number_input(
        "Incremento anual costos op. (M$/año)",
        min_value=0.0, value=5.0, step=1.0,
        key="nb_op_inc",
    )

    # ── Ajustes de personal ─────────────────────────────────────────────────
    _PERSONAL_TIPOS = {
        "Aumento de sueldo":  ("remuneration",      True),
        "Bono / gratificación": ("remuneration",    True),
        "Reajuste IPC/UF":    ("remuneration",      True),
        "Nuevo cargo":        ("remuneration",      True),
        "Aumento costos op.": ("operational_costs", True),
        "Otro gasto personal":("otros_gastos",      True),
    }
    _DURACION_OPTS = ["Permanente (desde ese mes)", "Solo ese mes", "Período específico"]

    _personal_ovs = [
        ov for ov in st.session_state.overrides
        if ov.get("param") in ("remuneration", "operational_costs", "otros_gastos")
        and ov.get("_seccion") == "personal"
    ]

    with st.expander(
        f"Ajustes de personal ({len(_personal_ovs)} activos)" if _personal_ovs else "Agregar ajuste de personal",
        expanded=False,
    ):
        p_tipo  = st.selectbox("Tipo", list(_PERSONAL_TIPOS.keys()), key="p_tipo_sel")
        p_desc  = st.text_input("Descripción", key="p_desc_inp", max_chars=60,
                                placeholder="ej: Reajuste sueldo base, Bono diciembre…")
        p_monto = st.number_input("Monto (M$/mes)", min_value=0.0, step=0.5,
                                  format="%.1f", key="p_monto_inp")
        p_desde = st.number_input("Desde mes #", min_value=1, max_value=72,
                                  value=1, step=1, key="p_desde_inp")

        p_dur   = st.radio("Duración", _DURACION_OPTS, key="p_dur_radio", horizontal=False)

        p_hasta = p_desde  # default para "solo ese mes"
        if p_dur == "Período específico":
            p_hasta = st.number_input("Hasta mes #", min_value=p_desde,
                                      max_value=72, value=min(p_desde+11, 72),
                                      step=1, key="p_hasta_inp")

        # Preview
        _yr_p  = (int(p_desde) - 1) // 12 + 1
        _mo_p  = (int(p_desde) - 1) % 12 + 1
        if p_dur == "Permanente (desde ese mes)":
            _dur_txt = f"desde Año {_yr_p} Mes {_mo_p:02d} en adelante"
            _to = 0
        elif p_dur == "Solo ese mes":
            _dur_txt = f"solo Año {_yr_p} Mes {_mo_p:02d}"
            _to = int(p_desde)
        else:
            _yr_h = (int(p_hasta) - 1) // 12 + 1
            _mo_h = (int(p_hasta) - 1) % 12 + 1
            _dur_txt = f"Mes {p_desde} → Mes {p_hasta} (Año {_yr_p}M{_mo_p:02d}–Año {_yr_h}M{_mo_h:02d})"
            _to = int(p_hasta)

        st.caption(f"M${p_monto:.1f}/mes · {_dur_txt}")

        if st.button("Agregar ajuste", key="btn_add_personal", use_container_width=True):
            _param, _delta = _PERSONAL_TIPOS[p_tipo]
            _note = p_tipo + (f": {p_desc.strip()}" if p_desc.strip() else "")
            st.session_state.overrides.append({
                "param":      _param,
                "from_month": int(p_desde),
                "to_month":   _to,
                "value":      float(p_monto),
                "note":       _note,
                "is_delta":   _delta,
                "_seccion":   "personal",
            })
            st.rerun()

        # Listado de ajustes activos de personal
        if _personal_ovs:
            st.markdown("**Ajustes activos:**")
            for idx_g, ov in enumerate(st.session_state.overrides):
                if ov.get("_seccion") != "personal":
                    continue
                yr_f = (ov["from_month"] - 1) // 12 + 1
                to_m = ov.get("to_month", 0)
                if to_m == 0:
                    rng = f"Mes {ov['from_month']} (Año {yr_f}) →"
                elif to_m == ov["from_month"]:
                    rng = f"Solo mes {ov['from_month']} (Año {yr_f})"
                else:
                    yr_t = (to_m - 1) // 12 + 1
                    rng = f"Mes {ov['from_month']}–{to_m} (Año {yr_f}–{yr_t})"
                col_pa, col_pb = st.columns([5, 1])
                with col_pa:
                    st.markdown(
                        f'<div class="ov-badge" style="border-color:#6ee7b7;">'
                        f'<b>{ov.get("note","")}</b> '
                        f'<span style="color:#6ee7b7;font-weight:700;">+M${ov["value"]:.1f}/mes</span>'
                        f'<br><span style="color:#64748b;font-size:0.75rem;">{rng}</span></div>',
                        unsafe_allow_html=True,
                    )
                with col_pb:
                    if st.button("✕", key=f"del_p_{idx_g}"):
                        st.session_state.overrides.pop(idx_g)
                        st.rerun()

    # ── Gastos proyectados (contrataciones y nuevos costos) ─────────────────
    st.markdown('<p class="sb-section">📅 Gastos Proyectados</p>', unsafe_allow_html=True)
    st.caption("Agrega contrataciones o gastos nuevos desde un mes específico.")

    with st.expander("➕ Nueva contratación o gasto"):
        gasto_tipo = st.selectbox(
            "Tipo de gasto",
            ["Nueva contratación", "Gasto operacional", "Otro gasto"],
            key="gasto_tipo_sel",
        )
        gasto_desc = st.text_input(
            "Descripción",
            key="gasto_desc_inp",
            max_chars=60,
            placeholder="ej: Analista de riesgo, Arriendo oficina…",
        )
        gasto_monto = st.number_input(
            "Monto (M$/mes)",
            min_value=0.0, step=0.5, format="%.1f",
            key="gasto_monto_inp",
        )
        gasto_mes = st.number_input(
            "Desde el mes #",
            min_value=1, max_value=72, value=13, step=1,
            key="gasto_mes_inp",
        )
        _yr_g = (int(gasto_mes) - 1) // 12 + 1
        _mo_g = (int(gasto_mes) - 1) % 12 + 1
        st.caption(f"→ Año {_yr_g}, Mes {_mo_g:02d}  ·  incremento sobre el valor base")

        _GASTO_PARAM = {
            "Nueva contratación": "remuneration",
            "Gasto operacional":  "operational_costs",
            "Otro gasto":         "otros_gastos",
        }
        if st.button("✅ Agregar gasto", key="btn_add_gasto", use_container_width=True):
            _note = f"{gasto_tipo}" + (f": {gasto_desc.strip()}" if gasto_desc.strip() else "")
            st.session_state.overrides.append({
                "param":      _GASTO_PARAM[gasto_tipo],
                "from_month": int(gasto_mes),
                "value":      float(gasto_monto),
                "note":       _note,
                "is_delta":   True,   # siempre incremental en este formulario
            })
            st.rerun()

    # ── Modificaciones programadas ──────────────────────────────────────────
    st.markdown('<p class="sb-section">⚡ Modificaciones Programadas</p>', unsafe_allow_html=True)
    st.caption("Cambia un parámetro desde un mes específico en adelante.")

    with st.expander("➕ Agregar modificación"):
        ov_param = st.selectbox(
            "Parámetro a cambiar",
            options=list(PARAM_LABELS.keys()),
            format_func=lambda k: PARAM_LABELS[k],
            key="ov_param_input",
        )
        ov_from = st.number_input(
            "Desde el mes #",
            min_value=1, max_value=72, value=13, step=1,
            key="ov_from_input",
        )
        yr_ = (ov_from - 1) // 12 + 1
        mo_ = (ov_from - 1) % 12 + 1
        st.caption(f"→ Año {yr_}, Mes {mo_:02d}  |  Unidad: {PARAM_UNITS.get(ov_param, '')}")

        ov_value = st.number_input(
            "Nuevo valor",
            value=0.0, step=0.1,
            key="ov_value_input",
            help="Para tasas usa %, para montos usa M$",
        )
        ov_note = st.text_input("Nota (opcional)", key="ov_note_input", max_chars=80)

        if st.button("✅ Agregar", use_container_width=True, key="btn_add_ov"):
            st.session_state.overrides.append({
                "param":      ov_param,
                "from_month": int(ov_from),
                "value":      float(ov_value),
                "note":       ov_note.strip(),
            })
            st.rerun()

    # Listado de overrides activos
    if st.session_state.overrides:
        st.markdown(f"**{len(st.session_state.overrides)} modificación(es) activa(s):**")
        for idx, ov in enumerate(st.session_state.overrides):
            col_a, col_b = st.columns([5, 1])
            label    = PARAM_LABELS.get(ov["param"], ov["param"])
            yr_ov    = (ov["from_month"] - 1) // 12 + 1
            unit     = PARAM_UNITS.get(ov["param"], "")
            is_delta = ov.get("is_delta", False)
            icon     = "➕" if is_delta else "📌"
            prefix   = "+" if is_delta else ""
            clr_val  = "#6ee7b7" if is_delta else "#F47920"
            with col_a:
                st.markdown(
                    f'<div class="ov-badge">{icon} <b>{label}</b>: '
                    f'<span style="color:{clr_val};font-weight:700;">'
                    f'{prefix}{ov["value"]} {unit}</span>'
                    f' desde mes {ov["from_month"]} (Año {yr_ov})'
                    + (f'<br><i style="color:#64748b">{ov["note"]}</i>' if ov.get("note") else "")
                    + "</div>",
                    unsafe_allow_html=True,
                )
            with col_b:
                if st.button("✕", key=f"del_ov_{idx}", help="Eliminar esta modificación"):
                    st.session_state.overrides.pop(idx)
                    st.rerun()

        if st.button("🗑 Limpiar todas las modificaciones", use_container_width=True):
            st.session_state.overrides = []
            st.rerun()
    else:
        st.caption("Sin modificaciones programadas.")

    # ── Escenarios ──────────────────────────────────────────────────────────
    st.markdown('<p class="sb-section">💾 Escenarios</p>', unsafe_allow_html=True)

    sc_name = st.text_input("Nombre del escenario", value="Escenario Base", key="sc_name_input")

    col_sv, col_ld = st.columns(2)
    with col_sv:
        if st.button("💾 Guardar", use_container_width=True, key="btn_save_sc"):
            # params se construye más abajo; necesitamos guardar estado actual
            # Usamos session state para comunicar la señal de guardado
            st.session_state["_save_trigger"] = sc_name
    with col_ld:
        existing = list_scenarios()
        if existing:
            sc_load_sel = st.selectbox("Escenario", options=existing, key="sc_load_sel", label_visibility="collapsed")
            if st.button("📂 Cargar", use_container_width=True, key="btn_load_sc"):
                loaded = load_scenario(sc_load_sel)
                if loaded:
                    st.session_state.overrides = [
                        {"param": o.param, "from_month": o.from_month,
                         "value": o.value, "note": o.note}
                        for o in loaded.overrides
                    ]
                    st.session_state.active_scenario = sc_load_sel
                    st.success(f"Cargado: {sc_load_sel}")
                    st.rerun()
        else:
            st.caption("Sin escenarios guardados.")


# ── Construcción de parámetros ─────────────────────────────────────────────────

params = ProjectionParams(
    initial_portfolio=initial_portfolio,
    annual_portfolio_growth=annual_portfolio_growth,
    placement_rate=placement_rate,
    leasing_portfolio_uf=leasing_portfolio_uf,
    leasing_annual_rate=leasing_annual_rate,
    leasing_annual_growth_mm=leasing_annual_growth_mm,
    uf_value=uf_value,
    monthly_ipc=monthly_ipc,
    funding_cost_rate=funding_cost_rate,
    current_op_costs=current_op_costs,
    current_remuneration=current_remuneration,
    other_expenses=other_expenses,
    annual_op_increment=annual_op_increment,
    leasing_avg_term_months=leasing_avg_term,
    provision_rate=provision_rate,
    npl_rate=npl_rate,
    initial_equity=initial_equity,
    tax_rate=tax_rate,
    overrides=[Override(**ov) for ov in st.session_state.overrides],
)

# Guardar escenario si fue solicitado
if st.session_state.get("_save_trigger"):
    name_to_save = st.session_state.pop("_save_trigger")
    if save_scenario(name_to_save, params):
        st.session_state.active_scenario = name_to_save
        st.sidebar.success(f"✅ Guardado: '{name_to_save}'")
    else:
        st.sidebar.error("Error al guardar escenario.")

# ── Generación de proyección ──────────────────────────────────────────────────

df_monthly = generate_monthly_projection(params)


# ── Encabezado y selector de periodicidad ─────────────────────────────────────

hdr_col, per_col = st.columns([4, 1])
with hdr_col:
    st.markdown("# 📊 Proyección Financiera — Factoring")
    st.markdown("Horizonte: **6 años (72 meses)** · Valores en millones de pesos (M$)")
with per_col:
    period = st.selectbox(
        "Periodicidad tabla",
        options=["Mensual", "Trimestral", "Semestral", "Anual"],
        index=3,
        key="period_sel",
    )

df_period = aggregate_by_period(df_monthly, period)

# ── Header principal ─────────────────────────────────────────────────────────

col_logo, col_title = st.columns([1, 5])
with col_logo:
    if os.path.exists("logo_smb.png"):
        st.image("logo_smb.png", use_container_width=True)
with col_title:
    st.markdown(
        '<h1 style="color:#F47920;margin-bottom:0;padding-bottom:0;">'
        'Proyección Financiera</h1>'
        '<p style="color:#A8A8A8;margin-top:2px;font-size:0.9rem;">'
        'SMB Servicios Financieros S.A. &nbsp;·&nbsp; Horizonte 6 años (72 meses)</p>',
        unsafe_allow_html=True,
    )
st.markdown('<hr style="border-color:#F47920;margin-top:6px;margin-bottom:16px;">', unsafe_allow_html=True)

# ── Barra fija de parámetros (sticky bajo el toolbar de Streamlit) ─────────────
st.markdown(_param_bar(params, df_monthly), unsafe_allow_html=True)

# ── Tabs principales ──────────────────────────────────────────────────────────

tab_dash, tab_table, tab_charts, tab_scenarios = st.tabs([
    "🎯  Dashboard KPIs",
    "📋  Tabla de Proyecciones",
    "📈  Gráficos Detallados",
    "💾  Escenarios",
])


# ════════════════════════════════════════════════════════════════════════════════
# TAB 1 — DASHBOARD KPIs
# ════════════════════════════════════════════════════════════════════════════════

with tab_dash:
    last   = df_monthly.iloc[-1]
    first  = df_monthly.iloc[0]

    total_income    = df_monthly["Ingresos_factoring"].sum()
    total_net       = df_monthly["Resultado_neto"].sum()
    avg_margin      = df_monthly["Margen_neto_pct"].mean()
    final_portfolio = last["Cartera"]
    growth_portfolio = final_portfolio - initial_portfolio
    final_roa       = last["ROA_pct"]
    final_eff       = last["Eficiencia_pct"]

    # KPI cards
    cards_html = '<div class="kpi-grid">'
    cards_html += kpi_card(
        "Cartera Final (Año 6)",
        f"${final_portfolio:,.0f} M",
        f"▲ ${growth_portfolio:,.0f} M vs inicio",
        "pos",
    )
    cards_html += kpi_card(
        "Ingresos Totales (6 años)",
        f"${total_income:,.0f} M",
        f"Promedio: ${total_income/72:,.0f} M/mes",
        "neutral",
    )
    cards_html += kpi_card(
        "Resultado Neto Total",
        f"${total_net:,.0f} M",
        f"Margen prom: {avg_margin:.1f}%",
        "pos" if total_net >= 0 else "neg",
    )
    cards_html += kpi_card(
        "ROA Mensual (Mes 72)",
        f"{final_roa:.2f}%",
        f"Eficiencia: {final_eff:.1f}%",
        "pos" if final_roa >= 0 else "neg",
    )
    final_roe       = last["ROE_pct"]
    final_patrimonio = last["Patrimonio"]
    cards_html += kpi_card(
        "ROE Anualizado (Mes 72)",
        f"{final_roe:.1f}%",
        f"Patrimonio: M${final_patrimonio:,.0f}",
        "pos" if final_roe >= 0 else "neg",
    )
    cards_html += "</div>"
    st.markdown(cards_html, unsafe_allow_html=True)

    st.markdown("---")

    # Tabla resumen anual
    st.subheader("Resumen Anual")

    df_annual = aggregate_by_period(df_monthly, "Anual")
    display_cols = ["Periodo", "Cartera", "Ingresos_factoring", "Margen_financiero",
                    "Provision_expense", "Total_costos", "EBIT", "Impuesto",
                    "Resultado_neto", "Patrimonio", "Margen_neto_pct", "ROA_pct",
                    "ROE_pct", "Eficiencia_pct"]
    display_cols = [c for c in display_cols if c in df_annual.columns]

    rename_map = {
        "Periodo": "Período", "Mes": "Mes", "Año": "Año",
        "Cartera": "Cartera (M$)", "Cartera_total": "Cartera Total (M$)",
        "Ingresos_factoring": "Ingresos (M$)", "Ingresos_total": "Ing. Total (M$)",
        "Costo_fondo": "Costo Fondo (M$)",
        "Margen_financiero": "Margen Fin. (M$)", "Margen_financiero_neto": "Margen Neto Fin. (M$)",
        "Provision_expense": "Provisiones (M$)",
        "Costos_operacionales": "Costos Op. (M$)", "Remuneraciones": "Remuner. (M$)",
        "Otros_gastos": "Otros Gastos (M$)",
        "Total_costos": "Total Costos (M$)",
        "EBIT": "EBIT (M$)", "Diferencia_cambio": "Dif. Cambio (M$)",
        "Resultado_antes_imp": "Res. Antes Imp. (M$)", "Impuesto": "Impuesto (M$)",
        "Resultado_neto": "Resultado Neto (M$)",
        "Patrimonio": "Patrimonio (M$)",
        "Margen_neto_pct": "Margen Neto (%)", "ROA_pct": "ROA (%)",
        "ROE_pct": "ROE (%)", "Eficiencia_pct": "Eficiencia (%)",
    }

    df_ann_disp = df_annual[display_cols].rename(columns=rename_map)

    # Función de formateo inline
    def _fmt(df_in: pd.DataFrame):
        fmt_dict = {}
        for col in df_in.columns:
            if col == "Período":
                continue
            elif "(%)" in col:
                fmt_dict[col] = "{:.1f}%"
            elif "Ejec." in col:
                fmt_dict[col] = "{:.0f}"
            else:
                fmt_dict[col] = "${:,.1f}"

        def _color_net(val):
            if isinstance(val, (int, float)):
                return "color: #F47920" if val >= 0 else "color: #EF4444"
            return ""

        net_col = ["Resultado Neto (M$)"] if "Resultado Neto (M$)" in df_in.columns else []
        styled = df_in.style.format(fmt_dict, na_rep="—")
        if net_col:
            styled = styled.map(_color_net, subset=net_col)
        return styled.set_properties(**{"background-color": "#0a1628", "color": "#e2e8f0"})

    st.dataframe(_fmt(df_ann_disp), use_container_width=True, hide_index=True)

    # Gráficos rápidos en dashboard
    st.markdown("---")
    ch1, ch2 = st.columns(2)
    with ch1:
        st.plotly_chart(plot_portfolio_evolution(df_monthly), use_container_width=True, key="dash_portfolio")
    with ch2:
        st.plotly_chart(plot_net_result(df_monthly), use_container_width=True, key="dash_net")

    # Waterfall anual
    st.plotly_chart(plot_waterfall_annual(df_annual), use_container_width=True, key="dash_wf")


# ════════════════════════════════════════════════════════════════════════════════
# TAB 2 — TABLA DE PROYECCIONES
# ════════════════════════════════════════════════════════════════════════════════

with tab_table:
    st.subheader(f"Proyección {period}")

    # Columnas a mostrar según periodicidad
    if period == "Mensual":
        show_cols = ["Mes", "Año", "Cartera", "Ingresos_factoring", "Costo_fondo",
                     "Margen_financiero", "Provision_expense", "Costos_operacionales",
                     "Remuneraciones", "Otros_gastos", "Total_costos",
                     "EBIT", "Diferencia_cambio", "Resultado_antes_imp",
                     "Impuesto", "Resultado_neto", "Patrimonio",
                     "Margen_neto_pct", "ROE_pct", "Eficiencia_pct", "ROA_pct"]
    else:
        show_cols = ["Periodo", "Cartera", "Ingresos_factoring", "Costo_fondo",
                     "Margen_financiero", "Provision_expense", "Costos_operacionales",
                     "Remuneraciones", "Otros_gastos", "Total_costos",
                     "EBIT", "Diferencia_cambio", "Resultado_antes_imp",
                     "Impuesto", "Resultado_neto", "Patrimonio",
                     "Margen_neto_pct", "ROE_pct", "Eficiencia_pct", "ROA_pct"]

    show_cols = [c for c in show_cols if c in df_period.columns]
    df_tbl    = df_period[show_cols].rename(columns=rename_map)

    # Filtro de columnas opcional
    with st.expander("⚙️ Seleccionar columnas visibles"):
        all_cols   = df_tbl.columns.tolist()
        sel_cols   = st.multiselect("Columnas", options=all_cols, default=all_cols, key="col_sel")
        if sel_cols:
            df_tbl = df_tbl[sel_cols]

    st.dataframe(
        _fmt(df_tbl) if len(df_tbl.columns) > 0 else df_tbl,
        use_container_width=True, hide_index=True, height=520,
    )

    st.markdown("---")
    exp_col, info_col = st.columns([2, 3])
    with exp_col:
        excel_bytes = export_to_excel(df_monthly, params)
        st.download_button(
            label="📥 Exportar a Excel (todas las periodicidades)",
            data=excel_bytes,
            file_name=f"proyeccion_factoring_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with info_col:
        st.info(
            "El archivo Excel incluye 5 hojas: **Resumen Ejecutivo**, "
            "**Mensual** (72 filas), **Trimestral**, **Semestral** y **Anual**."
        )


# ════════════════════════════════════════════════════════════════════════════════
# TAB 3 — GRÁFICOS DETALLADOS
# ════════════════════════════════════════════════════════════════════════════════

with tab_charts:
    # Selector de gráficos
    chart_options = [
        "Evolución de Cartera",
        "Ingresos vs Costos",
        "Resultado Neto Mensual",
        "Margen y Eficiencia",
        "Composición de Costos",
        "Waterfall Anual",
    ]
    selected_charts = st.multiselect(
        "Gráficos a mostrar",
        options=chart_options,
        default=chart_options,
        key="chart_sel",
    )

    if "Evolución de Cartera" in selected_charts:
        st.plotly_chart(plot_portfolio_evolution(df_monthly), use_container_width=True, key="ch_port")

    if "Ingresos vs Costos" in selected_charts or "Resultado Neto Mensual" in selected_charts:
        ch_col1, ch_col2 = st.columns(2)
        if "Ingresos vs Costos" in selected_charts:
            with ch_col1:
                st.plotly_chart(plot_income_vs_costs(df_monthly), use_container_width=True, key="ch_inc")
        if "Resultado Neto Mensual" in selected_charts:
            with ch_col2:
                st.plotly_chart(plot_net_result(df_monthly), use_container_width=True, key="ch_net")

    if "Margen y Eficiencia" in selected_charts:
        st.plotly_chart(plot_margin_trend(df_monthly), use_container_width=True, key="ch_marg")

    if "Composición de Costos" in selected_charts:
        st.plotly_chart(plot_cost_breakdown(df_monthly), use_container_width=True, key="ch_cost")

    if "Waterfall Anual" in selected_charts:
        df_annual_ch = aggregate_by_period(df_monthly, "Anual")
        st.plotly_chart(plot_waterfall_annual(df_annual_ch), use_container_width=True, key="ch_wf")


# ════════════════════════════════════════════════════════════════════════════════
# TAB 4 — GESTIÓN DE ESCENARIOS
# ════════════════════════════════════════════════════════════════════════════════

with tab_scenarios:
    st.subheader("Gestión de Escenarios")
    st.markdown(
        "Guarda distintas configuraciones para comparar proyecciones alternativas. "
        "Los escenarios se almacenan como archivos JSON en la carpeta `scenarios/`."
    )

    sc_list = list_scenarios()

    # ── Guardar desde el tab ────────────────────────────────────────────────
    with st.expander("💾 Guardar escenario actual", expanded=not bool(sc_list)):
        new_sc_name = st.text_input("Nombre", value="Mi Escenario", key="tab_sc_name")
        if st.button("Guardar escenario", key="tab_btn_save"):
            if save_scenario(new_sc_name, params):
                st.session_state.active_scenario = new_sc_name
                st.success(f"✅ Guardado como **{new_sc_name}**")
                st.rerun()
            else:
                st.error("Error al guardar. Verifica el nombre.")

    st.markdown("---")

    # ── Listado ──────────────────────────────────────────────────────────────
    sc_list = list_scenarios()
    if not sc_list:
        st.info("No hay escenarios guardados todavía.")
    else:
        st.markdown(f"**{len(sc_list)} escenario(s) disponible(s):**")
        for sc in sc_list:
            col_name, col_load, col_del = st.columns([5, 1, 1])
            is_active = (sc == st.session_state.active_scenario)
            with col_name:
                label = f"{'✅ ' if is_active else '📌 '}{sc}"
                st.write(label)
            with col_load:
                if st.button("Cargar", key=f"load_{sc}", use_container_width=True):
                    loaded = load_scenario(sc)
                    if loaded:
                        st.session_state.overrides = [
                            {"param": o.param, "from_month": o.from_month,
                             "value": o.value, "note": o.note}
                            for o in loaded.overrides
                        ]
                        st.session_state.active_scenario = sc
                        st.rerun()
            with col_del:
                if st.button("🗑", key=f"del_{sc}", use_container_width=True, help="Eliminar"):
                    delete_scenario(sc)
                    if st.session_state.active_scenario == sc:
                        st.session_state.active_scenario = ""
                    st.rerun()

    # ── Comparación de escenarios ────────────────────────────────────────────
    st.markdown("---")
    st.subheader("Comparar Escenarios")

    sc_list_fresh = list_scenarios()
    if len(sc_list_fresh) < 1:
        st.info("Guarda al menos 1 escenario para comparar con el actual.")
    else:
        comp_sc = st.selectbox("Comparar escenario actual vs:", options=sc_list_fresh, key="comp_sel")
        if st.button("📊 Comparar", key="btn_compare"):
            loaded_comp = load_scenario(comp_sc)
            if loaded_comp:
                df_comp = generate_monthly_projection(loaded_comp)
                fig_comp = plot_comparison(df_monthly, df_comp, "Escenario Actual", comp_sc)
                st.plotly_chart(fig_comp, use_container_width=True, key="ch_comp")

                # Diferencias clave
                st.markdown("#### Diferencias clave (Mes 72)")
                diff_cols = ["Cartera", "Resultado_neto", "Margen_neto_pct", "Total_costos"]
                diff_data = {}
                for col in diff_cols:
                    v1 = df_monthly.iloc[-1][col]
                    v2 = df_comp.iloc[-1][col]
                    diff_data[col] = {"Actual": v1, comp_sc: v2, "Δ": v1 - v2}
                st.dataframe(pd.DataFrame(diff_data).T, use_container_width=True)
            else:
                st.error(f"No se pudo cargar el escenario '{comp_sc}'.")
