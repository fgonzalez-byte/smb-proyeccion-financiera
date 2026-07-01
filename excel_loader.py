"""
Carga parámetros financieros reales desde el Excel de SMB Factoring.
Extrae tasas y costos actuales del EERR IFRS para inicializar el sistema.

Notas:
  - Todos los valores del EERR están en M$ (miles de pesos).
  - Se convierten a M$ millones dividiendo por K=1000.
  - La cartera factoring ($5.018.356.208) se convierte a M$ millones directamente.
  - Las tasas se calculan contra la cartera factoring declarada.
  - Los costos cubren la operación total de la empresa (factoring + leasing).
"""
from __future__ import annotations
import pandas as pd
import warnings
warnings.filterwarnings("ignore")

# Cartera factoring real al 31/05/2026 (en pesos, declarada por el usuario)
STOCK_FACTORING_PESOS = 5_018_356_208.0

# Índices de columna del EERR IFRS
C_ENE, C_FEB, C_MAR, C_ABR, C_MAY = 4, 5, 6, 7, 8
C_ACUM = 16
MESES = [C_ENE, C_FEB, C_MAR, C_ABR, C_MAY]
N_MESES = len(MESES)
K = 1_000.0   # miles → millones

# Índices de fila (0-based) del EERR IFRS
ROW_INT_FACT   = 6    # INTERESES FACTORING
ROW_COM_FACT   = 30   # COMISIONES FACTORING
ROW_COSTO_FOND = 46   # COSTO FINANCIAMIENTO OPERACIONAL
ROW_REM        = 55   # REMUNERACIONES
ROW_FINIQUITOS = 66   # FINIQUITOS (no recurrente)
ROW_OTROS_NOOP = 99   # OTROS COSTOS NO OPERACIONALES
ROW_GASTOS_OP  = 128  # GASTOS OPERACIONALES (total)
ROW_RES_NETO   = 154  # RESULTADO NETO


def load_smb_params(filepath: str) -> dict:
    """
    Lee hoja 'SMB EERR IFRS 05-2026' y retorna parámetros para ProjectionParams.

    Retorna dict con:
      - Claves directas para ProjectionParams (initial_portfolio, placement_rate, ...)
      - Claves "_meta_*" con datos de respaldo para mostrar al usuario
    """
    eerr = pd.ExcelFile(filepath).parse("SMB EERR IFRS 05-2026", header=None)

    def v(row: int, col: int = C_MAY) -> float:
        x = eerr.iloc[row, col]
        return float(x) if pd.notna(x) else 0.0

    def avg_meses(row: int) -> float:
        """Promedio absoluto de los 5 meses (valores pueden ser negativos en hoja)."""
        return sum(abs(v(row, c)) for c in MESES) / N_MESES

    # ── Cartera factoring inicial ────────────────────────────────────────────
    portfolio_mm = STOCK_FACTORING_PESOS / 1_000_000.0   # en M$ millones

    # ── Ingresos factoring (acumulado 5 meses, en M$ miles) ─────────────────
    int_fact_acum  = abs(v(ROW_INT_FACT,  C_ACUM))
    com_fact_acum  = abs(v(ROW_COM_FACT,  C_ACUM))
    ing_fact_mm    = (int_fact_acum + com_fact_acum) / N_MESES / K   # M$/mes

    # ── Costo de financiamiento (acumulado) ──────────────────────────────────
    costo_fond_acum = abs(v(ROW_COSTO_FOND, C_ACUM))
    costo_fond_mm   = costo_fond_acum / N_MESES / K   # M$/mes

    # ── Tasas ────────────────────────────────────────────────────────────────
    placement_rate    = 2.0   # tasa de colocación inicial definida
    funding_cost_rate = 0.44   # tasa real de fondeo, no derivada del EERR total

    # ── Remuneraciones (promedio mensual, acumulado excl. finiquitos) ────────
    rem_acum       = abs(v(ROW_REM,       C_ACUM))
    finiq_acum     = abs(v(ROW_FINIQUITOS, C_ACUM))   # no recurrente
    rem_mm         = rem_acum / N_MESES / K   # M$/mes

    # ── Costos operacionales recurrentes ─────────────────────────────────────
    # = Total GASTOS OP - Remuneraciones - Finiquitos (no-rec) - Otros no-op
    gastos_op_acum  = abs(v(ROW_GASTOS_OP, C_ACUM))
    otros_noop_acum = abs(v(ROW_OTROS_NOOP, C_ACUM))
    op_rec_mm = (gastos_op_acum - rem_acum - finiq_acum - otros_noop_acum) / N_MESES / K
    op_rec_mm = max(op_rec_mm, 0.0)

    # ── Otros gastos no operacionales (promedio mensual) ─────────────────────
    otros_mm = otros_noop_acum / N_MESES / K   # M$/mes

    # ── Resultados de respaldo ───────────────────────────────────────────────
    res_neto_may  = v(ROW_RES_NETO, C_MAY)
    res_neto_acum = v(ROW_RES_NETO, C_ACUM)

    return {
        # ── Parámetros para ProjectionParams ─────────────────────────────────
        "initial_portfolio":    round(portfolio_mm, 0),
        "placement_rate":       round(placement_rate, 2),
        "funding_cost_rate":    round(funding_cost_rate, 2),
        "current_remuneration": round(rem_mm, 0),
        "current_op_costs":     round(op_rec_mm, 0),
        "other_expenses":       round(otros_mm, 0),

        # ── Metadatos para mostrar al usuario ─────────────────────────────────
        "_periodo":              "Mayo 2026",
        "_cartera_mm":           round(portfolio_mm, 0),
        "_total_fondeo_mm":      round(costo_fond_mm, 1),
        "_ing_fact_mayo_mm":     round(abs(v(ROW_INT_FACT, C_MAY)) / K, 1),
        "_com_fact_mayo_mm":     round(abs(v(ROW_COM_FACT, C_MAY)) / K, 1),
        "_costo_fond_mayo_mm":   round(abs(v(ROW_COSTO_FOND, C_MAY)) / K, 1),
        "_ing_fact_prom_mm":     round(ing_fact_mm, 1),
        "_costo_fond_prom_mm":   round(costo_fond_mm, 1),
        "_spread_pct":           round(placement_rate - funding_cost_rate, 2),
        "_rem_mayo_mm":          round(abs(v(ROW_REM, C_MAY)) / K, 1),
        "_rem_prom_mm":          round(rem_mm, 1),
        "_op_rec_mm":            round(op_rec_mm, 1),
        "_otros_noop_mm":        round(otros_mm, 1),
        "_finiquitos_prom_mm":   round(finiq_acum / N_MESES / K, 1),
        "_res_neto_mayo_mm":     round(res_neto_may / K, 1),
        "_res_neto_acum_mm":     round(res_neto_acum / K, 1),
    }


def _safe_float(df: pd.DataFrame, row: int, col: int) -> float:
    try:
        vv = df.iloc[row, col]
        return float(vv) if pd.notna(vv) else 0.0
    except (IndexError, ValueError, TypeError):
        return 0.0
