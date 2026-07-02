"""
Gráficos Plotly para el dashboard de proyección financiera.
"""
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import pandas as pd

# ── Paleta de colores SMB (naranja #F47920 + gris) ───────────────────────────
C = {
    "portfolio":    "#F47920",   # naranja SMB — factoring
    "leasing":      "#FFA94D",   # naranja claro — leasing
    "reajuste":     "#F0CFA0",   # crema — reajuste UF
    "income":       "#F47920",   # naranja
    "cost":         "#EF4444",   # rojo
    "margin":       "#C8C8C8",   # gris claro
    "net_pos":      "#F47920",   # naranja positivo
    "net_neg":      "#EF4444",   # rojo negativo
    "remuneration": "#D4651A",   # naranja oscuro
    "operational":  "#8B8B8B",   # gris medio
    "other":        "#5C5C5C",   # gris oscuro
    "bg":           "rgba(28,28,28,0.85)",
    "grid":         "#3D3D3D",
    "text":         "#A8A8A8",
}

_LEGEND = dict(
    bgcolor="rgba(0,0,0,0)",
    font=dict(size=10, color=C["text"]),
    orientation="h",
    yanchor="bottom", y=1.02,
    xanchor="left",   x=0.0,
)

LAYOUT = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor=C["bg"],
    font=dict(color=C["text"], size=11),
    margin=dict(l=45, r=20, t=40, b=40),
    legend=_LEGEND,
    xaxis=dict(gridcolor=C["grid"], showgrid=True, zeroline=False),
    yaxis=dict(gridcolor=C["grid"], showgrid=True, zeroline=False),
    hovermode="x unified",
)


def _add_year_markers(fig, n_years: int = 6, row=None, col=None):
    """Líneas verticales punteadas en cada fin de año."""
    for yr in range(1, n_years):
        kwargs = dict(
            x=yr * 12, line_dash="dot", line_color="#334155",
            annotation_text=f"Año {yr}", annotation_position="top right",
            annotation_font_size=9, annotation_font_color=C["text"],
        )
        if row:
            kwargs["row"] = row
            kwargs["col"] = col
        fig.add_vline(**kwargs)


# ── Gráfico 1: Evolución de cartera (factoring + leasing apilado) ────────────

def plot_portfolio_evolution(df: pd.DataFrame) -> go.Figure:
    fig = go.Figure()

    has_leasing = "Cartera_leasing" in df.columns and df["Cartera_leasing"].sum() > 0

    if has_leasing:
        fig.add_trace(go.Scatter(
            x=df["Mes"], y=df["Cartera_leasing"],
            name="Cartera Leasing",
            stackgroup="portfolio",
            line=dict(color="#a78bfa", width=1),
            fillcolor="rgba(167,139,250,0.25)",
            hovertemplate="Mes %{x}<br>Leasing: $%{y:,.0f}M<extra></extra>",
        ))
        fig.add_trace(go.Scatter(
            x=df["Mes"], y=df["Cartera"],
            name="Cartera Factoring",
            stackgroup="portfolio",
            line=dict(color=C["portfolio"], width=2),
            fillcolor="rgba(96,165,250,0.20)",
            hovertemplate="Mes %{x}<br>Factoring: $%{y:,.0f}M<extra></extra>",
        ))
        # Total como línea encima
        fig.add_trace(go.Scatter(
            x=df["Mes"], y=df["Cartera_total"],
            name="Total", mode="lines",
            line=dict(color="#e2e8f0", width=1.5, dash="dot"),
            hovertemplate="Mes %{x}<br>Total: $%{y:,.0f}M<extra></extra>",
        ))
    else:
        fig.add_trace(go.Scatter(
            x=df["Mes"], y=df["Cartera"],
            name="Cartera Factoring",
            fill="tozeroy", line=dict(color=C["portfolio"], width=2),
            fillcolor="rgba(96,165,250,0.12)",
            hovertemplate="Mes %{x}<br>Cartera: $%{y:,.0f}M<extra></extra>",
        ))


    _add_year_markers(fig)
    fig.update_layout(
        title=dict(text=""),
        xaxis_title="Mes", yaxis_title="M$", **LAYOUT,
    )
    return fig


# ── Gráfico 2: Ingresos vs Costos ────────────────────────────────────────────

def plot_income_vs_costs(df: pd.DataFrame) -> go.Figure:
    fig = go.Figure()

    has_leasing = "Ingresos_leasing" in df.columns and df["Ingresos_leasing"].sum() > 0

    if has_leasing:
        traces = [
            ("Ing. Factoring",   "Ingresos_factoring",  C["income"],       "solid"),
            ("Ing. Leasing",     "Ingresos_leasing",    "#a78bfa",         "solid"),
            ("Reajuste UF",      "Reajuste_leasing",    "#f0abfc",         "dot"),
            ("Costo de Fondo",   "Costo_fondo",          C["cost"],         "dash"),
            ("Margen Financiero","Margen_financiero",    C["margin"],       "solid"),
            ("Total Costos",     "Total_costos",         C["remuneration"], "dot"),
        ]
    else:
        traces = [
            ("Ingresos Factoring","Ingresos_factoring", C["income"],       "solid"),
            ("Costo de Fondo",    "Costo_fondo",         C["cost"],         "dash"),
            ("Margen Financiero", "Margen_financiero",   C["margin"],       "solid"),
            ("Total Costos",      "Total_costos",        C["remuneration"], "dot"),
        ]

    for name, col, color, dash in traces:
        if col in df.columns:
            fig.add_trace(go.Scatter(
                x=df["Mes"], y=df[col], name=name,
                line=dict(color=color, width=2, dash=dash),
                hovertemplate=f"Mes %{{x}}<br>{name}: $%{{y:,.1f}}M<extra></extra>",
            ))

    _add_year_markers(fig)
    fig.update_layout(
        title=dict(text=""), xaxis_title="Mes", yaxis_title="M$", **LAYOUT,
    )
    return fig


# ── Gráfico 3: Resultado Neto mensual ────────────────────────────────────────

def plot_net_result(df: pd.DataFrame) -> go.Figure:
    colors = [C["income"] if v >= 0 else C["cost"] for v in df["Resultado_neto"]]
    fig = go.Figure()

    fig.add_trace(go.Bar(
        x=df["Mes"], y=df["Resultado_neto"],
        marker_color=colors,
        name="Resultado Neto",
        hovertemplate="Mes %{x}<br><b>Resultado: $%{y:,.1f}M</b><extra></extra>",
    ))

    # Línea de breakeven
    fig.add_hline(y=0, line_dash="dot", line_color="#475569", line_width=1)

    _add_year_markers(fig)
    fig.update_layout(
        title=dict(text=""), xaxis_title="Mes", yaxis_title="M$",
        showlegend=False, **LAYOUT,
    )
    return fig


# ── Gráfico 4: Margen neto y eficiencia (eje dual) ───────────────────────────

def plot_margin_trend(df: pd.DataFrame) -> go.Figure:
    fig = make_subplots(specs=[[{"secondary_y": True}]])

    fig.add_trace(go.Scatter(
        x=df["Mes"], y=df["Margen_neto_pct"],
        name="Margen Neto (%)", line=dict(color=C["net_pos"], width=2),
        hovertemplate="Mes %{x}<br>Margen: %{y:.1f}%<extra></extra>",
    ), secondary_y=False)

    fig.add_trace(go.Scatter(
        x=df["Mes"], y=df["Eficiencia_pct"],
        name="Eficiencia (%)", line=dict(color=C["cost"], width=2, dash="dash"),
        hovertemplate="Mes %{x}<br>Eficiencia: %{y:.1f}%<extra></extra>",
    ), secondary_y=True)

    for yr in range(1, 6):
        fig.add_vline(x=yr * 12, line_dash="dot", line_color=C["grid"])

    fig.update_layout(
        title=dict(text=""),
        xaxis=dict(title="Mes", gridcolor=C["grid"]),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor=C["bg"],
        font=dict(color=C["text"], size=11),
        margin=dict(l=45, r=45, t=40, b=40),
        legend=_LEGEND,
        hovermode="x unified",
    )
    fig.update_yaxes(title_text="Margen Neto (%)", secondary_y=False, gridcolor=C["grid"])
    fig.update_yaxes(title_text="Eficiencia (%)",  secondary_y=True)
    return fig


# ── Gráfico 5: Composición de costos (área apilada) ──────────────────────────

def plot_cost_breakdown(df: pd.DataFrame) -> go.Figure:
    fig = go.Figure()

    cost_layers = [
        ("Costos Operacionales", "Costos_operacionales", C["operational"], "rgba(251,191,36,0.4)"),
        ("Remuneraciones",       "Remuneraciones",        C["remuneration"], "rgba(251,146,60,0.4)"),
    ]
    if df["Otros_gastos"].sum() > 0:
        cost_layers.append(("Otros Gastos", "Otros_gastos", C["other"], "rgba(148,163,184,0.3)"))

    for name, col, line_color, fill_color in cost_layers:
        fig.add_trace(go.Scatter(
            x=df["Mes"], y=df[col],
            name=name, stackgroup="costs",
            line=dict(color=line_color, width=1),
            fillcolor=fill_color,
            hovertemplate=f"Mes %{{x}}<br>{name}: $%{{y:,.1f}}M<extra></extra>",
        ))

    # Overlay margen financiero
    fig.add_trace(go.Scatter(
        x=df["Mes"], y=df["Margen_financiero"],
        name="Margen Financiero",
        line=dict(color=C["margin"], width=2.5, dash="dot"),
        hovertemplate="Mes %{x}<br>Margen: $%{y:,.1f}M<extra></extra>",
    ))

    _add_year_markers(fig)
    fig.update_layout(
        title=dict(text=""),
        xaxis_title="Mes", yaxis_title="M$", **LAYOUT,
    )
    return fig


# ── Gráfico 6: Waterfall anual ───────────────────────────────────────────────

def plot_waterfall_annual(df_annual: pd.DataFrame) -> go.Figure:
    """Waterfall de resultado neto acumulado año a año."""
    periods = df_annual["Periodo"].tolist() if "Periodo" in df_annual.columns else [f"Año {i+1}" for i in range(len(df_annual))]
    values  = df_annual["Resultado_neto"].tolist()

    measure = ["relative"] * len(values) + ["total"]
    x       = periods + ["Total 6 años"]
    y       = values + [sum(values)]

    fig = go.Figure(go.Waterfall(
        orientation="v",
        measure=measure,
        x=x, y=y,
        connector=dict(line=dict(color=C["grid"], width=1)),
        increasing=dict(marker_color=C["income"]),
        decreasing=dict(marker_color=C["cost"]),
        totals=dict(marker_color=C["portfolio"]),
        hovertemplate="%{x}<br><b>$%{y:,.0f}M</b><extra></extra>",
    ))

    fig.update_layout(
        title=dict(text=""),
        yaxis_title="M$", **LAYOUT,
    )
    return fig


# ── Gráfico 7: Comparación de escenarios ─────────────────────────────────────

def plot_comparison(df1: pd.DataFrame, df2: pd.DataFrame, name1: str, name2: str) -> go.Figure:
    """Comparación side-by-side de dos escenarios."""
    fig = make_subplots(
        rows=2, cols=2,
        subplot_titles=["Resultado Neto (M$)", "Cartera (M$)", "Margen Neto (%)", "Total Costos (M$)"],
    )

    specs = [
        (1, 1, "Resultado_neto",   "M$"),
        (1, 2, "Cartera",          "M$"),
        (2, 1, "Margen_neto_pct",  "%"),
        (2, 2, "Total_costos",     "M$"),
    ]

    colors1 = [C["net_pos"],   C["portfolio"], C["margin"], C["operational"]]
    colors2 = [C["remuneration"], C["cost"],  C["income"],  C["other"]]

    for (row, col, metric, unit), c1, c2 in zip(specs, colors1, colors2):
        show_legend_row = (row == 1 and col == 1)
        for df, name, color, dash in [(df1, name1, c1, "solid"), (df2, name2, c2, "dash")]:
            fig.add_trace(go.Scatter(
                x=df["Mes"], y=df[metric], name=name,
                line=dict(color=color, width=2, dash=dash),
                showlegend=show_legend_row,
                hovertemplate=f"Mes %{{x}}<br>{name}: %{{y:,.1f}}{unit}<extra></extra>",
            ), row=row, col=col)

    fig.update_layout(
        title=dict(text=""),
        height=640,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor=C["bg"],
        font=dict(color=C["text"], size=11),
        margin=dict(l=45, r=20, t=40, b=40),
        legend=_LEGEND,
        hovermode="x unified",
    )
    for axis in fig.layout:
        if axis.startswith("xaxis") or axis.startswith("yaxis"):
            fig.layout[axis].gridcolor = C["grid"]
    return fig
