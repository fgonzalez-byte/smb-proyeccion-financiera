"""
Motor de cálculo financiero para proyecciones de empresa de factoring + leasing.
Soporta overrides de parámetros por mes específico.
"""
from __future__ import annotations
from dataclasses import dataclass, field, fields, asdict
from typing import List
import pandas as pd


# ── Parámetros de override ────────────────────────────────────────────────────

@dataclass
class Override:
    """Cambio de parámetro desde un mes específico en adelante."""
    param: str        # clave del parámetro a sobreescribir
    from_month: int   # mes desde el cual aplica (1-72)
    value: float      # nuevo valor (absoluto) o incremento (si is_delta=True)
    note: str = ""    # descripción opcional
    is_delta: bool = False  # True → suma incremental al valor vigente; False → reemplaza
    to_month: int = 0       # 0 = permanente; >0 = aplica solo hasta ese mes (inclusive)

    def year_label(self) -> str:
        yr = (self.from_month - 1) // 12 + 1
        mo = (self.from_month - 1) % 12 + 1
        return f"Año {yr}, Mes {mo:02d}"


# ── Parámetros principales ────────────────────────────────────────────────────

PARAM_LABELS = {
    "portfolio_growth":      "Crecimiento cartera factoring (M$/mes)",
    "funding_cost":          "Costo de fondo (% mensual)",
    "base_funding_cost":     "Costo fijo de fondo (M$/mes)",
    "placement_rate":        "Tasa de colocación factoring (% mensual)",
    "operational_costs":     "Costos operacionales (M$/mes)",
    "remuneration":          "Remuneraciones base (M$/mes)",
    "otros_gastos":          "Otros gastos no operacionales (M$/mes)",
    "leasing_rate":          "Tasa anual leasing (%)",
    "leasing_growth":        "Crecimiento cartera leasing (M$/año)",
    "ipc_mensual":           "IPC mensual proyectado (%) — reajuste UF",
    "provision_rate":        "Tasa de provisión anual (% cartera total)",
    "tax_rate":              "Tasa impuesto corporativo (%)",
}

PARAM_UNITS = {
    "portfolio_growth":      "M$/mes",
    "funding_cost":          "%",
    "base_funding_cost":     "M$/mes",
    "placement_rate":        "%",
    "operational_costs":     "M$/mes",
    "remuneration":          "M$/mes",
    "otros_gastos":          "M$/mes",
    "leasing_rate":          "%",
    "leasing_growth":        "M$/año",
    "ipc_mensual":           "%",
    "provision_rate":        "%",
    "tax_rate":              "%",
}


@dataclass
class ProjectionParams:
    # ── Factoring ─────────────────────────────────────────────────────────────
    initial_portfolio: float = 4560.0        # Cartera factoring inicial (M$) — calibrado EERR May-26
    annual_portfolio_growth: float = 1000.0  # Crecimiento anual factoring (M$/año)
    placement_rate: float = 2.0              # Tasa colocación mensual (%)

    # ── Leasing (valores con reajuste UF) ─────────────────────────────────────
    leasing_portfolio_uf: float = 92295.0    # Stock inicial leasing en UF — calibrado EERR May-26
    leasing_annual_rate: float = 17.0        # Tasa anual leasing (%) con reajuste
    leasing_annual_growth_mm: float = 200.0  # Crecimiento anual leasing (M$ nominal)
    uf_value: float = 39900.0               # Valor UF en pesos al inicio proyección
    monthly_ipc: float = 0.35               # IPC mensual proyectado (%) para reajuste UF
    leasing_avg_term_months: int = 36        # Plazo promedio contratos leasing (meses)

    # ── Financiamiento ────────────────────────────────────────────────────────
    funding_cost_rate: float = 0.44          # Costo de fondo mensual (%) — sobre cartera total
    base_funding_cost_mm: float = 9.93       # Costo fijo de fondo (M$/mes) — overhead no proporcional a cartera

    # ── Costos operacionales ──────────────────────────────────────────────────
    current_op_costs: float = 28.0           # Costos op. recurrentes base (M$/mes)
    current_remuneration: float = 59.8       # Remuneraciones base (M$/mes) — calibrado EERR May-26
    other_expenses: float = 13.8             # Otros gastos no operacionales (M$/mes) — calibrado EERR May-26
    annual_op_increment: float = 5.0         # Incremento anual costos op (M$/año)


    # ── Riesgo y provisiones ──────────────────────────────────────────────────
    provision_rate: float = 1.5              # % anual gasto provisiones sobre cartera total
    npl_rate: float = 2.0                    # % cartera morosa (informativo, no afecta P&L directo)

    # ── Patrimonio e impuesto ─────────────────────────────────────────────────
    initial_equity: float = 500.0            # M$ patrimonio inicial de la empresa
    tax_rate: float = 27.0                   # % impuesto renta sobre resultado positivo (Chile)
    apply_tax: bool = True                   # Si False, impuesto = 0 y no aparece en reportes

    # ── Overrides programados ─────────────────────────────────────────────────
    overrides: List[Override] = field(default_factory=list)

    def monthly_growth(self) -> float:
        return self.annual_portfolio_growth / 12.0

    def get_param_at_month(self, param: str, month: int) -> float:
        """Valor efectivo del parámetro en el mes dado, aplicando overrides.

        Overrides absolutos (is_delta=False) reemplazan el valor vigente.
        Overrides delta (is_delta=True) se acumulan sobre el valor vigente.
        Se aplican en orden cronológico, por lo que un absoluto posterior
        reinicia la base sobre la cual siguen acumulando los deltas siguientes.
        """
        base_values = {
            "portfolio_growth":  self.annual_portfolio_growth / 12.0,
            "funding_cost":      self.funding_cost_rate,
            "base_funding_cost": self.base_funding_cost_mm,
            "placement_rate":    self.placement_rate,
            "operational_costs": self.current_op_costs,
            "remuneration":      self.current_remuneration,
            "otros_gastos":      self.other_expenses,
            "leasing_rate":      self.leasing_annual_rate,
            "leasing_growth":    self.leasing_annual_growth_mm,
            "ipc_mensual":       self.monthly_ipc,
            "provision_rate":    self.provision_rate,
            "tax_rate":          self.tax_rate,
        }
        value = base_values.get(param, 0.0)
        for ov in sorted(self.overrides, key=lambda x: x.from_month):
            in_range = ov.from_month <= month and (ov.to_month == 0 or month <= ov.to_month)
            if ov.param == param and in_range:
                if ov.is_delta:
                    value += ov.value   # acumulativo
                else:
                    value = ov.value    # reemplaza
        return value

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, d: dict) -> "ProjectionParams":
        d = d.copy()
        overrides_raw = d.pop("overrides", [])
        valid_fields = {f.name for f in fields(cls) if f.name != "overrides"}
        filtered = {k: v for k, v in d.items() if k in valid_fields}
        obj = cls(**filtered)
        obj.overrides = [Override(**o) for o in overrides_raw]
        return obj


# ── Generación de proyecciones ────────────────────────────────────────────────

def generate_monthly_projection(params: ProjectionParams, total_months: int = 72) -> pd.DataFrame:
    """Genera DataFrame mensual con 72 meses de proyección financiera consolidada."""
    rows = []
    factoring_portfolio = params.initial_portfolio
    leasing_uf          = params.leasing_portfolio_uf
    uf_current          = params.uf_value
    equity              = params.initial_equity   # patrimonio acumula mes a mes

    for m in range(1, total_months + 1):
        year          = (m - 1) // 12 + 1
        month_in_year = (m - 1) % 12 + 1

        # Parámetros efectivos este mes (incluye overrides programados)
        monthly_fact_growth  = params.get_param_at_month("portfolio_growth", m)
        funding_rate_pct     = params.get_param_at_month("funding_cost", m)
        placement_pct        = params.get_param_at_month("placement_rate", m)
        base_op              = params.get_param_at_month("operational_costs", m)
        base_rem             = params.get_param_at_month("remuneration", m)
        leasing_annual_rate  = params.get_param_at_month("leasing_rate", m)
        leasing_growth_mm_yr = params.get_param_at_month("leasing_growth", m)
        ipc_mensual          = params.get_param_at_month("ipc_mensual", m)

        # ── Cartera factoring ─────────────────────────────────────────────────
        factoring_portfolio += monthly_fact_growth

        # ── Cartera leasing: reajuste UF + nuevos contratos ───────────────────
        # 1) Guardar cartera UF en pesos ANTES de aplicar IPC (para calcular reajuste)
        uf_prev              = uf_current
        leasing_mm_pre_reaj  = leasing_uf * uf_prev / 1_000_000.0

        # 2) Actualizar valor UF con IPC mensual (puede variar por override)
        uf_current = uf_prev * (1.0 + ipc_mensual / 100.0)

        # 3) Reajuste = ganancia por alza UF sobre cartera existente
        leasing_mm_post_reaj = leasing_uf * uf_current / 1_000_000.0
        reajuste_leasing     = leasing_mm_post_reaj - leasing_mm_pre_reaj

        # 4) Nuevos contratos en UF (en base al nuevo valor UF)
        leasing_monthly_growth_mm = leasing_growth_mm_yr / 12.0
        leasing_uf_growth = leasing_monthly_growth_mm * 1_000_000.0 / uf_current
        leasing_uf        += leasing_uf_growth
        leasing_portfolio_mm = leasing_uf * uf_current / 1_000_000.0

        # ── Ingresos ──────────────────────────────────────────────────────────
        factoring_income  = factoring_portfolio * (placement_pct / 100.0)

        # Ingreso leasing: método lineal (total intereses / plazo), igual que SMB
        # Total intereses = cuota × plazo - principal; cuota = P × r/(1-(1+r)^-n)
        r_mes = leasing_annual_rate / 12.0 / 100.0
        n     = params.leasing_avg_term_months
        if r_mes > 0:
            cuota_unitaria    = r_mes / (1.0 - (1.0 + r_mes) ** (-n))
            total_int_factor  = cuota_unitaria * n - 1.0   # % interés sobre principal
        else:
            total_int_factor  = 0.0
        monthly_sl_rate   = total_int_factor / n            # tasa lineal mensual
        leasing_income    = leasing_portfolio_mm * monthly_sl_rate

        # Reajuste UF: diferencia de cambio (línea 135 EERR) — no es ingreso operacional
        total_income      = factoring_income + leasing_income

        # ── Costo de fondo (proporcional a cartera + componente fijo) ────────
        total_portfolio    = factoring_portfolio + leasing_portfolio_mm
        base_funding_fixed = params.get_param_at_month("base_funding_cost", m)
        funding_cost       = total_portfolio * (funding_rate_pct / 100.0) + base_funding_fixed

        # ── Margen financiero bruto ───────────────────────────────────────────
        financial_margin = total_income - funding_cost

        # ── Provisiones y cartera morosa ─────────────────────────────────────
        provision_rate_pct   = params.get_param_at_month("provision_rate", m)
        provision_expense    = total_portfolio * (provision_rate_pct / 12.0 / 100.0)
        npl_portfolio        = total_portfolio * (params.npl_rate / 100.0)
        net_financial_margin = financial_margin - provision_expense

        # ── Costos operacionales ──────────────────────────────────────────────
        annual_increment_accrued = params.annual_op_increment * (year - 1)
        total_op_costs = base_op + annual_increment_accrued / 12.0

        total_remuneration = base_rem
        other_exp          = params.get_param_at_month("otros_gastos", m)

        total_costs = total_op_costs + total_remuneration + other_exp

        # ── EBIT ─────────────────────────────────────────────────────────────
        ebit = net_financial_margin - total_costs

        # ── Diferencia de cambio (reajuste UF, línea 135 EERR) ───────────────
        # Se suma al EBIT antes del impuesto, igual que en el EERR real
        resultado_antes_impuesto = ebit + reajuste_leasing

        # ── Impuesto ─────────────────────────────────────────────────────────
        tax_rate_pct = params.get_param_at_month("tax_rate", m)
        if params.apply_tax:
            tax = max(0.0, resultado_antes_impuesto * (tax_rate_pct / 100.0))
        else:
            tax = 0.0
        net_result   = resultado_antes_impuesto - tax

        # ── Patrimonio ────────────────────────────────────────────────────────
        equity_start = equity
        equity      += net_result

        # ── KPIs ──────────────────────────────────────────────────────────────
        total_income_incl_reaj = total_income + reajuste_leasing
        net_margin_pct = (net_result  / total_income_incl_reaj * 100) if total_income_incl_reaj > 0 else 0.0
        efficiency_pct = (total_costs / financial_margin        * 100) if financial_margin        > 0 else 0.0
        roa_pct        = (net_result * 12 / total_portfolio     * 100) if total_portfolio         > 0 else 0.0
        roe_pct        = (net_result * 12 / equity_start        * 100) if equity_start            > 0 else 0.0

        rows.append({
            "Mes":                   m,
            "Año":                   year,
            "Mes_en_año":            month_in_year,
            # Factoring
            "Cartera":               round(factoring_portfolio, 2),
            "Crecimiento_mensual":   round(monthly_fact_growth, 2),
            "Tasa_colocacion_pct":   round(placement_pct, 3),
            "Ingresos_factoring":    round(factoring_income, 2),
            # Leasing
            "Cartera_leasing_uf":    round(leasing_uf, 2),
            "UF_valor":              round(uf_current, 2),
            "Cartera_leasing":       round(leasing_portfolio_mm, 2),
            "Ingresos_leasing":      round(leasing_income, 2),
            "Reajuste_leasing":      round(reajuste_leasing, 2),
            # Consolidado
            "Cartera_total":           round(total_portfolio, 2),
            "Ingresos_total":          round(total_income, 2),
            "Costo_fondo_pct":         round(funding_rate_pct, 3),
            "Costo_fondo":             round(funding_cost, 2),
            "Margen_financiero":       round(financial_margin, 2),
            # Provisiones y riesgo
            "Provision_expense":       round(provision_expense, 2),
            "NPL_portfolio":           round(npl_portfolio, 2),
            "Margen_financiero_neto":  round(net_financial_margin, 2),
            # Costos
            "Costos_operacionales":    round(total_op_costs, 2),
            "Remuneraciones":          round(total_remuneration, 2),
            "Otros_gastos":            round(other_exp, 2),
            "Total_costos":            round(total_costs, 2),
            # EBIT, Diferencia de cambio, Impuesto, Resultado
            "EBIT":                    round(ebit, 2),
            "Diferencia_cambio":       round(reajuste_leasing, 2),
            "Resultado_antes_imp":     round(resultado_antes_impuesto, 2),
            "Impuesto":                round(tax, 2),
            "Resultado_neto":          round(net_result, 2),
            # Patrimonio
            "Patrimonio":              round(equity, 2),
            # KPIs
            "Margen_neto_pct":         round(net_margin_pct, 2),
            "Eficiencia_pct":          round(efficiency_pct, 2),
            "ROA_pct":                 round(roa_pct, 2),
            "ROE_pct":                 round(roe_pct, 2),
        })

    return pd.DataFrame(rows)


def aggregate_by_period(df: pd.DataFrame, period: str) -> pd.DataFrame:
    """Agrega proyección mensual al período seleccionado."""
    if period == "Mensual":
        return df.copy()

    df = df.copy()

    if period == "Trimestral":
        df["Grupo"]   = (df["Mes"] - 1) // 3
        df["Periodo"] = df["Grupo"].apply(lambda g: f"T{g + 1}  (Año {g // 4 + 1})")
    elif period == "Semestral":
        df["Grupo"]   = (df["Mes"] - 1) // 6
        df["Periodo"] = df["Grupo"].apply(lambda g: f"S{g + 1}  (Año {g // 2 + 1})")
    elif period == "Anual":
        df["Grupo"]   = df["Año"] - 1
        df["Periodo"] = df["Año"].apply(lambda y: f"Año {y}")
    else:
        return df

    sum_cols = [
        "Ingresos_factoring", "Ingresos_leasing", "Reajuste_leasing", "Ingresos_total",
        "Costo_fondo", "Margen_financiero",
        "Provision_expense", "Margen_financiero_neto",
        "Costos_operacionales", "Remuneraciones", "Otros_gastos",
        "Total_costos", "EBIT", "Diferencia_cambio", "Resultado_antes_imp",
        "Impuesto", "Resultado_neto", "Crecimiento_mensual",
    ]
    last_cols = [
        "Cartera", "Cartera_leasing", "Cartera_total",
        "Cartera_leasing_uf", "UF_valor",
        "NPL_portfolio", "Patrimonio",
    ]
    agg_dict = {c: "sum" for c in sum_cols if c in df.columns}
    agg_dict.update({c: "last" for c in last_cols if c in df.columns})
    agg_dict["Periodo"] = "first"

    result = df.groupby("Grupo").agg(agg_dict).reset_index(drop=True)

    with pd.option_context("mode.chained_assignment", None):
        denom_income  = result["Ingresos_total"].replace(0, float("nan"))
        denom_margin  = result["Margen_financiero"].replace(0, float("nan"))
        denom_cart    = result["Cartera_total"].replace(0, float("nan"))
        denom_patr    = result["Patrimonio"].replace(0, float("nan")) if "Patrimonio" in result.columns else float("nan")
        n_months      = 12 if period == "Anual" else (6 if period == "Semestral" else (3 if period == "Trimestral" else 1))
        result["Margen_neto_pct"] = (result["Resultado_neto"] / denom_income  * 100).round(2).fillna(0)
        result["Eficiencia_pct"]  = (result["Total_costos"]   / denom_margin  * 100).round(2).fillna(0)
        result["ROA_pct"]         = (result["Resultado_neto"] / denom_cart    * n_months * 100).round(2).fillna(0)
        if "Patrimonio" in result.columns:
            result["ROE_pct"]     = (result["Resultado_neto"] / denom_patr    * n_months * 100).round(2).fillna(0)

    return result
