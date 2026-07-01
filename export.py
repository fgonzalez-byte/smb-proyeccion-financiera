"""
Exportación a Excel con formato ejecutivo usando openpyxl.
Genera libro con 5 hojas: Resumen, Mensual, Trimestral, Semestral, Anual.
"""
import io
from datetime import datetime

from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, numbers as xl_numbers
)
from openpyxl.utils import get_column_letter
import pandas as pd

from model import ProjectionParams, aggregate_by_period

# ── Estilos ───────────────────────────────────────────────────────────────────

F_DARK       = PatternFill("solid", fgColor="0F2037")
F_NAVY       = PatternFill("solid", fgColor="1E3A5F")
F_ALT        = PatternFill("solid", fgColor="0A1628")
F_POS        = PatternFill("solid", fgColor="052E16")
F_NEG        = PatternFill("solid", fgColor="450A0A")

FONT_TITLE   = Font(name="Calibri", color="60A5FA", size=16, bold=True)
FONT_HDR     = Font(name="Calibri", color="FFFFFF", size=10, bold=True)
FONT_SUB     = Font(name="Calibri", color="60A5FA", size=10, bold=True)
FONT_NORMAL  = Font(name="Calibri", color="E2E8F0", size=10)
FONT_POS     = Font(name="Calibri", color="4ADE80", size=10, bold=True)
FONT_NEG     = Font(name="Calibri", color="F87171", size=10, bold=True)
FONT_LABEL   = Font(name="Calibri", color="94A3B8", size=10)

ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_RIGHT  = Alignment(horizontal="right",  vertical="center")
ALIGN_LEFT   = Alignment(horizontal="left",   vertical="center")

BORDER_THIN  = Border(
    bottom=Side(border_style="thin", color="1E3A5F"),
    top=Side(border_style="thin", color="1E3A5F"),
)

FMT_CURR = '"$"#,##0.0'
FMT_PCT  = '#,##0.00"%"'
FMT_INT  = '0'
FMT_NONE = '#,##0.0'

COL_FORMATS = {
    "Mes": FMT_INT, "Año": FMT_INT, "Mes_en_año": FMT_INT, "Ejecutivos_adicionales": FMT_INT,
    "Tasa_colocacion_pct": FMT_PCT, "Costo_fondo_pct": FMT_PCT,
    "Margen_neto_pct": FMT_PCT, "Eficiencia_pct": FMT_PCT, "ROA_pct": FMT_PCT,
}

COL_NAMES_ES = {
    "Mes": "Mes", "Año": "Año", "Periodo": "Período",
    "Cartera": "Cartera (M$)", "Crecimiento_mensual": "Crecimiento (M$)",
    "Tasa_colocacion_pct": "Tasa Coloc. (%)", "Costo_fondo_pct": "Costo Fdo. (%)",
    "Ingresos_factoring": "Ingresos (M$)", "Costo_fondo": "Costo Fondo (M$)",
    "Margen_financiero": "Margen Fin. (M$)", "Costos_operacionales": "Op. (M$)",
    "Remuneraciones": "Rem. (M$)", "Otros_gastos": "Otros (M$)",
    "Total_costos": "Total Costos (M$)", "Resultado_neto": "Resultado Neto (M$)",
    "Margen_neto_pct": "Margen Neto (%)", "Eficiencia_pct": "Eficiencia (%)",
    "ROA_pct": "ROA (%)", "Ejecutivos_adicionales": "Ejecutivos +",
}


def _write_df(ws, df: pd.DataFrame, start_row: int = 1, freeze: bool = True):
    """Escribe DataFrame con cabecera estilizada."""
    ws.sheet_view.showGridLines = False

    # Cabecera
    for ci, col in enumerate(df.columns, start=1):
        cell = ws.cell(row=start_row, column=ci, value=COL_NAMES_ES.get(col, col))
        cell.fill      = F_DARK
        cell.font      = FONT_HDR
        cell.alignment = ALIGN_CENTER
        cell.border    = BORDER_THIN

    # Datos
    for ri, row_data in enumerate(df.itertuples(index=False), start=start_row + 1):
        alt = (ri % 2 == 0)
        for ci, value in enumerate(row_data, start=1):
            cell = ws.cell(row=ri, column=ci, value=value)
            cell.font      = FONT_NORMAL
            cell.alignment = ALIGN_RIGHT

            col_name = df.columns[ci - 1]

            if alt:
                cell.fill = F_ALT

            # Formato numérico
            if col_name in COL_FORMATS:
                cell.number_format = COL_FORMATS[col_name]
            elif col_name in ("Periodo",):
                cell.alignment = ALIGN_LEFT
                cell.font = FONT_NORMAL
            elif isinstance(value, float):
                cell.number_format = FMT_CURR

            # Colorear resultado neto
            if col_name == "Resultado_neto" and isinstance(value, (int, float)):
                cell.fill = F_POS if value >= 0 else F_NEG
                cell.font = FONT_POS if value >= 0 else FONT_NEG

    # Congelar fila de cabecera
    if freeze:
        ws.freeze_panes = ws.cell(row=start_row + 1, column=1)

    # Ajustar anchos
    for ci, col in enumerate(df.columns, start=1):
        letter = get_column_letter(ci)
        header_len = len(COL_NAMES_ES.get(col, col))
        max_data   = max(
            (len(str(v)) for v in df[col] if v is not None),
            default=0,
        )
        ws.column_dimensions[letter].width = min(max(header_len, max_data) + 3, 22)

    # Altura de cabecera
    ws.row_dimensions[start_row].height = 32


def export_to_excel(df_monthly: pd.DataFrame, params: ProjectionParams) -> bytes:
    """Genera workbook Excel y retorna bytes para descarga."""
    wb = Workbook()

    # ── Hoja 1: Resumen Ejecutivo ─────────────────────────────────────────────
    ws_sum = wb.active
    ws_sum.title = "Resumen Ejecutivo"
    ws_sum.sheet_view.showGridLines = False
    ws_sum.column_dimensions["A"].width = 3
    ws_sum.column_dimensions["B"].width = 32
    ws_sum.column_dimensions["C"].width = 22

    ws_sum.merge_cells("B2:F2")
    ws_sum["B2"] = "PROYECCIÓN FINANCIERA — FACTORING"
    ws_sum["B2"].font      = FONT_TITLE
    ws_sum["B2"].fill      = F_DARK
    ws_sum["B2"].alignment = ALIGN_CENTER

    ws_sum["B3"] = f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}  |  Horizonte: 6 años (72 meses)"
    ws_sum["B3"].font = FONT_LABEL

    # Parámetros base
    ws_sum["B5"] = "PARÁMETROS BASE"
    ws_sum["B5"].font = FONT_SUB
    ws_sum["B5"].fill = F_NAVY

    param_rows = [
        ("Cartera Inicial",              f"${params.initial_portfolio:,.0f}M"),
        ("Crecimiento Anual de Cartera", f"${params.annual_portfolio_growth:,.0f}M"),
        ("Tasa de Colocación (mensual)", f"{params.placement_rate:.2f}%"),
        ("Costo de Fondo (mensual)",     f"{params.funding_cost_rate:.2f}%"),
        ("Costos Operacionales Base",    f"${params.current_op_costs:,.1f}M/mes"),
        ("Remuneraciones Base",          f"${params.current_remuneration:,.1f}M/mes"),
        ("Sueldo Nuevo Ejecutivo",       f"${params.new_executive_salary:,.1f}M/mes"),
        ("Contratación Ejecutivo Cada",  f"{params.executive_hire_every_n_years} años"),
        ("Incremento Anual Op.",         f"${params.annual_op_increment:,.1f}M/año"),
        ("Otros Gastos",                 f"${params.other_expenses:,.1f}M/mes"),
    ]
    if params.overrides:
        param_rows.append(("Modificaciones Programadas", f"{len(params.overrides)} override(s)"))

    for i, (label, value) in enumerate(param_rows, start=6):
        ws_sum[f"B{i}"] = label
        ws_sum[f"C{i}"] = value
        ws_sum[f"B{i}"].font = FONT_LABEL
        ws_sum[f"C{i}"].font = Font(name="Calibri", color="FFFFFF", size=10, bold=True)
        ws_sum[f"B{i}"].alignment = ALIGN_LEFT
        ws_sum[f"C{i}"].alignment = ALIGN_LEFT
        if i % 2 == 0:
            ws_sum[f"B{i}"].fill = F_ALT
            ws_sum[f"C{i}"].fill = F_ALT

    # KPIs finales
    last = df_monthly.iloc[-1]
    total_net  = df_monthly["Resultado_neto"].sum()
    avg_margin = df_monthly["Margen_neto_pct"].mean()
    final_portfolio = last["Cartera"]

    kpi_start = 18
    ws_sum[f"B{kpi_start}"] = "KPIs PROYECTADOS (6 AÑOS)"
    ws_sum[f"B{kpi_start}"].font = FONT_SUB
    ws_sum[f"B{kpi_start}"].fill = F_NAVY

    kpi_rows = [
        ("Cartera Final (Mes 72)",        f"${final_portfolio:,.0f}M"),
        ("Resultado Neto Total",           f"${total_net:,.0f}M"),
        ("Margen Neto Promedio",           f"{avg_margin:.1f}%"),
        ("ROA Final (Mes 72)",             f"{last['ROA_pct']:.2f}%"),
        ("Eficiencia Final (Mes 72)",      f"{last['Eficiencia_pct']:.1f}%"),
        ("Ejecutivos Adicionales (Año 6)", f"{int(last['Ejecutivos_adicionales'])}"),
    ]
    for i, (label, value) in enumerate(kpi_rows, start=kpi_start + 1):
        ws_sum[f"B{i}"] = label
        ws_sum[f"C{i}"] = value
        ws_sum[f"B{i}"].font = FONT_LABEL
        ws_sum[f"C{i}"].font = FONT_POS if "M" in str(value) or "%" in str(value) else FONT_NORMAL
        if i % 2 == 0:
            ws_sum[f"B{i}"].fill = F_ALT
            ws_sum[f"C{i}"].fill = F_ALT

    # Tabla anual en hoja resumen
    df_annual = aggregate_by_period(df_monthly, "Anual")
    sum_cols  = ["Periodo", "Cartera", "Ingresos_factoring", "Margen_financiero",
                 "Total_costos", "Resultado_neto", "Margen_neto_pct", "ROA_pct"]
    df_a_disp = df_annual[[c for c in sum_cols if c in df_annual.columns]]

    ws_sum[f"B{kpi_start + 9}"] = "RESUMEN ANUAL"
    ws_sum[f"B{kpi_start + 9}"].font = FONT_SUB
    ws_sum[f"B{kpi_start + 9}"].fill = F_NAVY

    _write_df(ws_sum, df_a_disp, start_row=kpi_start + 10, freeze=False)

    # ── Hojas de proyección ───────────────────────────────────────────────────
    periods = [
        ("Proyección Mensual",    "Mensual"),
        ("Trimestral",            "Trimestral"),
        ("Semestral",             "Semestral"),
        ("Anual",                 "Anual"),
    ]

    monthly_cols = ["Mes", "Año", "Cartera", "Ingresos_factoring", "Costo_fondo",
                    "Margen_financiero", "Costos_operacionales", "Remuneraciones",
                    "Otros_gastos", "Total_costos", "Resultado_neto",
                    "Margen_neto_pct", "Eficiencia_pct", "ROA_pct", "Ejecutivos_adicionales"]

    period_cols = ["Periodo", "Cartera", "Ingresos_factoring", "Costo_fondo",
                   "Margen_financiero", "Costos_operacionales", "Remuneraciones",
                   "Total_costos", "Resultado_neto", "Margen_neto_pct", "Eficiencia_pct", "ROA_pct"]

    for sheet_name, period in periods:
        ws = wb.create_sheet(sheet_name)
        df_p = aggregate_by_period(df_monthly, period)

        if period == "Mensual":
            cols = [c for c in monthly_cols if c in df_p.columns]
        else:
            cols = [c for c in period_cols if c in df_p.columns]

        _write_df(ws, df_p[cols])

    # ── Serializar a bytes ────────────────────────────────────────────────────
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()
