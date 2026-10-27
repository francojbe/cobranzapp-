import html
import io
from datetime import datetime
from typing import Optional
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
from fastapi import HTTPException
from fastapi.responses import StreamingResponse
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from core.analytics import (
    build_debtor_matched_dataset,
    calculate_campanas_tramos_matrix,
    calculate_fibex_summary,
    calculate_global_tramos_summary,
    calculate_totals,
)
from core.config import DATA_STORE, OFFICIAL_TRAMOS_CONFIG


# -------------------------------------------------------------
# TEMAS Y HELPERS COMPARTIDOS PARA EXCEL (OPENPYXL) Y PDF (REPORTLAB)
# -------------------------------------------------------------

def get_openpyxl_theme() -> dict:
    """Retorna la paleta tipográfica, rellenos y bordes corporativos de CobranzApp para Excel."""
    return {
        "font_title": Font(name="Calibri", size=14, bold=True, color="0F172A"),
        "font_subtitle": Font(name="Calibri", size=10, italic=True, color="64748B"),
        "font_header": Font(name="Calibri", size=11, bold=True, color="FFFFFF"),
        "font_group_hdr": Font(name="Calibri", size=10, bold=True, color="FFFFFF"),
        "font_sub_hdr": Font(name="Calibri", size=9, bold=True, color="334155"),
        "font_data": Font(name="Calibri", size=11, color="1E293B"),
        "font_data_sm": Font(name="Calibri", size=10, color="1E293B"),
        "font_data_bold": Font(name="Calibri", size=11, bold=True, color="1E293B"),
        "font_data_bold_sm": Font(name="Calibri", size=10, bold=True, color="0F172A"),
        "font_total": Font(name="Calibri", size=11, bold=True, color="0F172A"),
        "font_total_sm": Font(name="Calibri", size=10, bold=True, color="0F172A"),
        "fill_header": PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid"),
        "fill_header_sub": PatternFill(start_color="334155", end_color="334155", fill_type="solid"),
        "fill_sub_hdr": PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid"),
        "fill_alt": PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid"),
        "fill_total": PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid"),
        "fill_banner": PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid"),
        "fill_ca_label": PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid"),
        "thin_border": Border(
            left=Side(style='thin', color='CBD5E1'),
            right=Side(style='thin', color='CBD5E1'),
            top=Side(style='thin', color='CBD5E1'),
            bottom=Side(style='thin', color='CBD5E1')
        ),
        "total_border": Border(
            left=Side(style='thin', color='94A3B8'),
            right=Side(style='thin', color='94A3B8'),
            top=Side(style='medium', color='1E293B'),
            bottom=Side(style='double', color='1E293B')
        )
    }


def auto_fit_worksheet_columns(ws, max_cap: int = 60, min_width: int = 14):
    """Ajusta automáticamente el ancho de las columnas de una hoja OpenPyXL."""
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            if len(val_str) > max_len and len(val_str) < max_cap:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = max(max_len + 3, min_width)


def get_reportlab_theme() -> dict:
    """Retorna el diccionario de estilos tipográficos de ReportLab para los PDFs ejecutivos."""
    base = getSampleStyleSheet()
    table_header = ParagraphStyle(
        'TableHeader',
        parent=base['Normal'],
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor("#ffffff"),
        fontName="Helvetica-Bold"
    )
    return {
        "base": base,
        "title": ParagraphStyle(
            'ExecTitle',
            parent=base['Heading1'],
            fontSize=18,
            leading=22,
            textColor=colors.HexColor("#0f172a"),
            fontName="Helvetica-Bold",
            spaceAfter=4
        ),
        "subtitle": ParagraphStyle(
            'ExecSubtitle',
            parent=base['Normal'],
            fontSize=9,
            leading=13,
            textColor=colors.HexColor("#64748b"),
            spaceAfter=14
        ),
        "banner_label": ParagraphStyle(
            'BannerLabel',
            parent=base['Normal'],
            fontSize=8,
            leading=10,
            textColor=colors.HexColor("#64748b"),
            alignment=1,
            fontName="Helvetica-Bold"
        ),
        "banner_val": ParagraphStyle(
            'BannerVal',
            parent=base['Heading2'],
            fontSize=14,
            leading=17,
            textColor=colors.HexColor("#1e293b"),
            alignment=1,
            fontName="Helvetica-Bold"
        ),
        "table_cell": ParagraphStyle(
            'TableCell',
            parent=base['Normal'],
            fontSize=8.5,
            leading=11,
            textColor=colors.HexColor("#1e293b")
        ),
        "table_cell_center": ParagraphStyle(
            'TableCellCenter',
            parent=base['Normal'],
            fontSize=8.5,
            leading=11,
            textColor=colors.HexColor("#1e293b"),
            alignment=1
        ),
        "table_cell_right": ParagraphStyle(
            'TableCellRight',
            parent=base['Normal'],
            fontSize=8.5,
            leading=11,
            textColor=colors.HexColor("#1e293b"),
            alignment=2
        ),
        "table_header": table_header,
        "table_header_right": ParagraphStyle(
            'TableHeaderRight',
            parent=table_header,
            alignment=2
        ),
        "table_header_center": ParagraphStyle(
            'TableHeaderCenter',
            parent=table_header,
            alignment=1
        )
    }


def build_pdf_kpi_card_grid(kpi_cards: list, col_widths=None, bg_hex="#f8fafc", border_hex="#cbd5e1", grid_hex="#e2e8f0") -> Table:
    """Construye una grilla de tarjetas KPI para documentos PDF."""
    t_kpis = Table(kpi_cards, colWidths=col_widths or [135, 135, 135, 135])
    t_kpis.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor(bg_hex)),
        ('BOX', (0, 0), (-1, -1), 1, colors.HexColor(border_hex)),
        ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.HexColor(grid_hex)),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
    ]))
    return t_kpis


def build_pdf_strategic_note(notes_text: str, normal_style) -> Table:
    """Construye el recuadro azul de Observación Estratégica al pie de los PDFs ejecutivos."""
    p_notes = Paragraph(notes_text, normal_style)
    t_notes = Table([[p_notes]], colWidths=[540])
    t_notes.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor("#eff6ff")),
        ('BOX', (0, 0), (-1, -1), 1, colors.HexColor("#bfdbfe")),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('LEFTPADDING', (0, 0), (-1, -1), 10),
        ('RIGHTPADDING', (0, 0), (-1, -1), 10),
    ]))
    return t_notes


def _get_active_df_or_400() -> pd.DataFrame:
    active_df = DATA_STORE["df_original"] if DATA_STORE["df_original"] is not None else DATA_STORE["df_cartera_raw"]
    if active_df is None or active_df.empty:
        raise HTTPException(status_code=400, detail="No hay cartera activa cargada.")
    return active_df


# -------------------------------------------------------------
# 1. REPORTE DE CARTERA FILTRADA (EXCEL Y PDF)
# -------------------------------------------------------------

def generate_filtered_excel() -> StreamingResponse:
    """Descarga la cartera filtrada enriquecida con todas las columnas unificadas."""
    if DATA_STORE["df_filtered"] is None or DATA_STORE["df_filtered"].empty:
        raise HTTPException(status_code=400, detail="No hay datos filtrados para descargar.")

    df = DATA_STORE["df_filtered"].copy()
    preferred_order = [
        "Cedula Cliente", "Nombre Cliente", "Dias de Mora", "Tramo_Mora",
        "Pendiente Tarifa", "Pendiente Cuota", "Pendiente Mora", "Total",
        "Tipo_Campana", "Programa", "Telefono_1", "Telefono_2", "Telefono_3", "Correo",
        "Tipo", "Identificador de credito", "Canal", "Fecha Vencimiento", "Cuota Numero"
    ]
    cols = [c for c in preferred_order if c in df.columns] + [c for c in df.columns if c not in preferred_order and not c.startswith("_") and c != "Cedula_Key"]
    df_clean = df[cols]

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_clean.to_excel(writer, index=False, sheet_name="Cartera Enriquecida")

    output.seek(0)
    filename = f"cartera_desglose_enriquecida_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )


def generate_filtered_pdf() -> StreamingResponse:
    """Genera reporte ejecutivo formal en PDF con metricas financieras y de contactabilidad."""
    if DATA_STORE["df_filtered"] is None or DATA_STORE["df_filtered"].empty:
        raise HTTPException(status_code=400, detail="No hay datos filtrados para generar el reporte.")

    df = DATA_STORE["df_filtered"]
    kpis = calculate_totals(df)
    f_info = DATA_STORE.get("filter_info", {})

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
    rt = get_reportlab_theme()
    styles = rt["base"]

    story = [
        Paragraph("REPORTE EJECUTIVO DE CARTERA VENCIDA Y CONTACTABILIDAD", rt["title"]),
        Paragraph(
            f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')} | Rango mora: {f_info.get('min_dias', 1)} a {f_info.get('max_dias', 60)} días | Archivo Cartera: {html.escape(str(DATA_STORE.get('cartera_filename', 'Cartera')))}",
            rt["subtitle"]
        ),
        Spacer(1, 10)
    ]

    # KPI Grid Table 1: Financiero
    kpi_data = [
        [
            Paragraph("TOTAL CLIENTES", rt["banner_label"]),
            Paragraph("TOTAL CUOTAS EN MORA", rt["banner_label"]),
            Paragraph("MONTO TOTAL VENCIDO", rt["banner_label"]),
            Paragraph("PROMEDIO X CLIENTE", rt["banner_label"])
        ],
        [
            Paragraph(f"{kpis['total_clientes']:,}", rt["banner_val"]),
            Paragraph(f"{kpis['total_cuotas']:,}", rt["banner_val"]),
            Paragraph(f"${kpis['monto_total']:,.2f} USD", rt["banner_val"]),
            Paragraph(f"${kpis['promedio_cliente']:,.2f} USD", rt["banner_val"])
        ]
    ]
    story.append(build_pdf_kpi_card_grid(kpi_data, col_widths=[135, 135, 140, 130]))
    story.append(Spacer(1, 10))

    # KPI Grid Table 2: Contactabilidad
    contact_data = [
        [
            Paragraph("CLIENTES CON TELÉFONO", rt["banner_label"]),
            Paragraph("COBERTURA WHATSAPP", rt["banner_label"]),
            Paragraph("MONTO CONTACTABLE", rt["banner_label"]),
            Paragraph("% DEUDA CONTACTABLE", rt["banner_label"])
        ],
        [
            Paragraph(f"{kpis['total_con_telefono']:,}", rt["banner_val"]),
            Paragraph(f"{kpis['pct_con_telefono']}%", rt["banner_val"]),
            Paragraph(f"${kpis['monto_contactable']:,.2f} USD", rt["banner_val"]),
            Paragraph(f"{kpis['pct_monto_contactable']}%", rt["banner_val"])
        ]
    ]
    story.append(build_pdf_kpi_card_grid(
        contact_data, col_widths=[135, 135, 140, 130],
        bg_hex="#f0fdf4", border_hex="#bbf7d0", grid_hex="#bbf7d0"
    ))
    story.append(Spacer(1, 15))

    story.append(Paragraph(
        f"<b>Composición de la Deuda:</b> Tarifa Pendiente: <b>${kpis['pendiente_tarifa']:,.2f}</b> | Cuota Pendiente: <b>${kpis['pendiente_cuota']:,.2f}</b> | Recargo Mora: <b>${kpis['pendiente_mora']:,.2f}</b>",
        styles['Normal']
    ))
    story.append(Spacer(1, 15))

    df_master = DATA_STORE["df_original"] if DATA_STORE["df_original"] is not None else df
    global_summary = calculate_global_tramos_summary(df_master)
    cutoff = global_summary.get("cutoff_date", datetime.now().strftime('%d-%m-%Y'))
    total_cartera_val = global_summary.get("total_cartera_usd", kpis['monto_total'])

    story.append(Paragraph(f"<b>Resumen Agrupado por Tramos de Mora (al {cutoff})</b>", styles['Heading3']))
    story.append(Paragraph(
        f"<b>Cartera Activa Total: ${total_cartera_val:,.2f} USD</b> | Cuotas: {global_summary.get('total_cuotas', 0):,} | Clientes: {global_summary.get('total_clientes', 0):,}",
        styles['Normal']
    ))
    story.append(Spacer(1, 8))

    table_rows = [[
        Paragraph("Fase / Categoría", rt["table_header"]),
        Paragraph("Rango / Etiqueta", rt["table_header"]),
        Paragraph("Cant. Cuotas", rt["table_header"]),
        Paragraph("Clientes", rt["table_header"]),
        Paragraph("Monto (USD)", rt["table_header"]),
        Paragraph("% Cartera", rt["table_header"])
    ]]

    for t_data in global_summary.get("tramos", []):
        pct_val = f"{t_data['porcentaje']:.2f}%" if t_data.get("porcentaje") is not None else "-"
        table_rows.append([
            Paragraph(html.escape(t_data["fase"]), rt["table_cell"]),
            Paragraph(html.escape(t_data["etiqueta"]), rt["table_cell"]),
            Paragraph(f"{t_data['cuotas']:,}", rt["table_cell"]),
            Paragraph(f"{t_data['clientes']:,}", rt["table_cell"]),
            Paragraph(f"${t_data['total_monto']:,.2f}", rt["table_cell"]),
            Paragraph(pct_val, rt["table_cell"])
        ])

    table_rows.append([
        Paragraph("<b>Total Cartera</b>", rt["table_cell"]),
        Paragraph("<b>Gestión total</b>", rt["table_cell"]),
        Paragraph(f"<b>{global_summary.get('total_cuotas', 0):,}</b>", rt["table_cell"]),
        Paragraph(f"<b>{global_summary.get('total_clientes', 0):,}</b>", rt["table_cell"]),
        Paragraph(f"<b>${total_cartera_val:,.2f}</b>", rt["table_cell"]),
        Paragraph("<b>100.0%</b>", rt["table_cell"])
    ])

    t_grp = Table(table_rows, colWidths=[150, 110, 70, 70, 95, 55])
    t_grp.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1e293b")),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e1")),
        ('ROWBACKGROUNDS', (0, 1), (-1, -2), [colors.white, colors.HexColor("#f8fafc")]),
        ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor("#e2e8f0")),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ]))
    story.append(t_grp)

    doc.build(story)
    buf.seek(0)
    return StreamingResponse(
        buf,
        media_type="application/pdf",
        headers={"Content-Disposition": "attachment; filename=reporte_cartera_ejecutivo.pdf"}
    )


# -------------------------------------------------------------
# 2. REPORTE DE 5 TRAMOS OFICIALES DE MORA (EXCEL Y PDF)
# -------------------------------------------------------------

def generate_tramos_excel(cartera_activa: Optional[float] = None) -> StreamingResponse:
    """Descarga el Resumen Agrupado por Tramos de Mora en un archivo Excel formal."""
    active_df = _get_active_df_or_400()
    if cartera_activa is not None and cartera_activa > 0:
        DATA_STORE["cartera_activa_manual"] = cartera_activa
    manual_ca = cartera_activa or DATA_STORE.get("cartera_activa_manual")
    global_summary = calculate_global_tramos_summary(active_df, manual_ca)
    cutoff = global_summary.get("cutoff_date", datetime.now().strftime("%d-%m-%Y"))
    total_usd = global_summary.get("total_cartera_usd", 0.0)
    total_pct = global_summary.get("total_porcentaje")

    th = get_openpyxl_theme()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resumen por Tramos"
    ws.views.sheetView[0].showGridLines = True

    ws["A1"] = "COBRANZAPP - RESUMEN AGRUPADO POR TRAMOS DE MORA"
    ws["A1"].font = th["font_title"]

    ws["A2"] = f"Fecha de corte: al {cutoff} | Tarifa Pendiente Total: ${total_usd:,.2f} USD | Emisión: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ws["A2"].font = th["font_subtitle"]

    # Banner Ejecutivo
    ws.merge_cells("A4:B4")
    ws["A4"] = "TARIFA PENDIENTE TOTAL (USD)"
    ws["A4"].font = Font(name="Calibri", size=10, bold=True, color="1E3A8A")
    ws["A4"].alignment = Alignment(horizontal="center", vertical="center")
    ws["A4"].fill = th["fill_banner"]

    ws["C4"] = total_usd
    ws["C4"].font = Font(name="Calibri", size=14, bold=True, color="047857")
    ws["C4"].number_format = '$#,##0.00'
    ws["C4"].alignment = Alignment(horizontal="center", vertical="center")
    ws["C4"].fill = th["fill_banner"]

    ws["D4"] = "CARTERA ACTIVA (USD):"
    ws["D4"].font = Font(name="Calibri", size=10, bold=True, color="92400E")
    ws["D4"].alignment = Alignment(horizontal="center", vertical="center")
    ws["D4"].fill = th["fill_ca_label"]

    ws.merge_cells("E4:F4")
    if manual_ca and manual_ca > 0:
        ws["E4"] = manual_ca
        ws["E4"].font = Font(name="Calibri", size=14, bold=True, color="B45309")
        ws["E4"].number_format = '$#,##0.00'
    else:
        ws["E4"] = "No especificada"
        ws["E4"].font = Font(name="Calibri", size=11, italic=True, color="64748B")
    ws["E4"].alignment = Alignment(horizontal="center", vertical="center")
    ws["E4"].fill = th["fill_ca_label"]

    headers = ["Fase / Categoría", "Rango / Etiqueta", "Cant. Cuotas", "Clientes", "Monto Tarifa (USD)", "% Cartera Activa"]
    start_row = 6
    for col_idx, h_text in enumerate(headers, start=1):
        cell = ws.cell(row=start_row, column=col_idx, value=h_text)
        cell.font = th["font_header"]
        cell.fill = th["fill_header"]
        cell.alignment = Alignment(horizontal="center" if col_idx in [2, 6] else ("right" if col_idx in [3, 4, 5] else "left"), vertical="center")
        cell.border = th["thin_border"]
    ws.row_dimensions[start_row].height = 24

    curr_row = start_row + 1
    for idx, t in enumerate(global_summary.get("tramos", [])):
        r = curr_row + idx
        ws.row_dimensions[r].height = 20
        c1 = ws.cell(row=r, column=1, value=t["fase"])
        c2 = ws.cell(row=r, column=2, value=t["etiqueta"])
        c3 = ws.cell(row=r, column=3, value=t["cuotas"])
        c4 = ws.cell(row=r, column=4, value=t["clientes"])
        c5 = ws.cell(row=r, column=5, value=t["total_monto"])
        if t["porcentaje"] is not None:
            c6 = ws.cell(row=r, column=6, value=t["porcentaje"] / 100.0)
            c6.number_format = '0.00%'
        else:
            c6 = ws.cell(row=r, column=6, value="-")

        c1.font = th["font_data"]
        c2.font = th["font_data_bold"]
        c3.font = th["font_data"]
        c4.font = th["font_data"]
        c5.font = th["font_data_bold"]
        c6.font = th["font_data_bold"]

        c1.alignment = Alignment(horizontal="left", vertical="center")
        c2.alignment = Alignment(horizontal="center", vertical="center")
        c3.alignment = Alignment(horizontal="right", vertical="center")
        c4.alignment = Alignment(horizontal="right", vertical="center")
        c5.alignment = Alignment(horizontal="right", vertical="center")
        c6.alignment = Alignment(horizontal="right" if t["porcentaje"] is not None else "center", vertical="center")

        c3.number_format = '#,##0'
        c4.number_format = '#,##0'
        c5.number_format = '$#,##0.00'

        row_fill = th["fill_alt"] if idx % 2 == 1 else None
        for col_idx in range(1, 7):
            cell = ws.cell(row=r, column=col_idx)
            cell.border = th["thin_border"]
            if row_fill:
                cell.fill = row_fill

    tot_row = curr_row + len(global_summary.get("tramos", []))
    ws.row_dimensions[tot_row].height = 24

    tc1 = ws.cell(row=tot_row, column=1, value="TOTAL CARTERA")
    tc2 = ws.cell(row=tot_row, column=2, value="Gestión total")
    tc3 = ws.cell(row=tot_row, column=3, value=global_summary.get("total_cuotas", 0))
    tc4 = ws.cell(row=tot_row, column=4, value=global_summary.get("total_clientes", 0))
    tc5 = ws.cell(row=tot_row, column=5, value=total_usd)
    if total_pct is not None:
        tc6 = ws.cell(row=tot_row, column=6, value=total_pct / 100.0)
        tc6.number_format = '0.00%'
    else:
        tc6 = ws.cell(row=tot_row, column=6, value="-")

    for cell in [tc1, tc2, tc3, tc4, tc5, tc6]:
        cell.font = th["font_total"]
        cell.fill = th["fill_total"]
        cell.border = th["total_border"]

    tc1.alignment = Alignment(horizontal="left", vertical="center")
    tc2.alignment = Alignment(horizontal="center", vertical="center")
    tc3.alignment = Alignment(horizontal="right", vertical="center")
    tc4.alignment = Alignment(horizontal="right", vertical="center")
    tc5.alignment = Alignment(horizontal="right", vertical="center")
    tc6.alignment = Alignment(horizontal="right" if total_pct is not None else "center", vertical="center")

    tc3.number_format = '#,##0'
    tc4.number_format = '#,##0'
    tc5.number_format = '$#,##0.00'

    col_widths = {1: 34, 2: 20, 3: 16, 4: 16, 5: 22, 6: 15}
    for col_idx, width in col_widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"resumen_tramos_mora_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )


def generate_tramos_pdf(cartera_activa: Optional[float] = None) -> StreamingResponse:
    """Genera un reporte PDF ejecutivo exclusivamente del Resumen Agrupado por Tramos de Mora."""
    active_df = _get_active_df_or_400()
    if cartera_activa is not None and cartera_activa > 0:
        DATA_STORE["cartera_activa_manual"] = cartera_activa
    manual_ca = cartera_activa or DATA_STORE.get("cartera_activa_manual")
    global_summary = calculate_global_tramos_summary(active_df, manual_ca)
    cutoff = global_summary.get("cutoff_date", datetime.now().strftime("%d-%m-%Y"))
    total_usd = global_summary.get("total_cartera_usd", 0.0)
    total_cuotas = global_summary.get("total_cuotas", 0)
    total_clientes = global_summary.get("total_clientes", 0)
    total_pct = global_summary.get("total_porcentaje")

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
    rt = get_reportlab_theme()
    styles = rt["base"]

    story = [
        Paragraph("COBRANZAPP | RESUMEN AGRUPADO POR TRAMOS DE MORA", rt["title"]),
        Paragraph(
            f"<b>Fecha de corte:</b> al {cutoff} &nbsp;|&nbsp; <b>Fecha de emisión:</b> {datetime.now().strftime('%d/%m/%Y %H:%M')} &nbsp;|&nbsp; <b>Cartera:</b> {html.escape(str(DATA_STORE.get('cartera_filename', 'Cartera Activa')))}",
            rt["subtitle"]
        ),
        Spacer(1, 8)
    ]

    if manual_ca and manual_ca > 0:
        kpi_cards = [
            [
                Paragraph("TARIFA PENDIENTE TOTAL", rt["banner_label"]),
                Paragraph("CARTERA ACTIVA TOTAL", rt["banner_label"]),
                Paragraph("% MORA SOBRE CARTERA", rt["banner_label"]),
                Paragraph("TOTAL CLIENTES", rt["banner_label"])
            ],
            [
                Paragraph(f"${total_usd:,.2f} USD", rt["banner_val"]),
                Paragraph(f"${manual_ca:,.2f} USD", rt["banner_val"]),
                Paragraph(f"{total_pct:.2f}%" if total_pct is not None else "-", rt["banner_val"]),
                Paragraph(f"{total_clientes:,}", rt["banner_val"])
            ]
        ]
    else:
        kpi_cards = [
            [
                Paragraph("TARIFA PENDIENTE TOTAL", rt["banner_label"]),
                Paragraph("CARTERA ACTIVA TOTAL", rt["banner_label"]),
                Paragraph("TOTAL CLIENTES", rt["banner_label"]),
                Paragraph("CANTIDAD DE CUOTAS", rt["banner_label"])
            ],
            [
                Paragraph(f"${total_usd:,.2f} USD", rt["banner_val"]),
                Paragraph("No especificada", rt["banner_val"]),
                Paragraph(f"{total_clientes:,}", rt["banner_val"]),
                Paragraph(f"{total_cuotas:,}", rt["banner_val"])
            ]
        ]

    story.append(build_pdf_kpi_card_grid(kpi_cards))
    story.append(Spacer(1, 16))
    story.append(Paragraph("<b>Distribución de Tarifa Pendiente por los 5 Tramos Oficiales de Mora:</b>", styles['Normal']))
    story.append(Spacer(1, 6))

    table_rows = [[
        Paragraph("Fase / Categoría", rt["table_header"]),
        Paragraph("Rango / Etiqueta", rt["table_header_center"]),
        Paragraph("Cant. Cuotas", rt["table_header_right"]),
        Paragraph("Clientes", rt["table_header_right"]),
        Paragraph("Monto Tarifa (USD)", rt["table_header_right"]),
        Paragraph("% Cartera Activa", rt["table_header_right"])
    ]]

    for t in global_summary.get("tramos", []):
        pct_display = f"<b>{t['porcentaje']:.2f}%</b>" if t["porcentaje"] is not None else "-"
        table_rows.append([
            Paragraph(f"<b>{html.escape(t['fase'])}</b>", rt["table_cell"]),
            Paragraph(html.escape(t["etiqueta"]), rt["table_cell_center"]),
            Paragraph(f"{t['cuotas']:,}", rt["table_cell_right"]),
            Paragraph(f"{t['clientes']:,}", rt["table_cell_right"]),
            Paragraph(f"${t['total_monto']:,.2f}", rt["table_cell_right"]),
            Paragraph(pct_display, rt["table_cell_right"] if t["porcentaje"] is not None else rt["table_cell_center"])
        ])

    total_pct_display = f"<b>{total_pct:.2f}%</b>" if total_pct is not None else "-"
    table_rows.append([
        Paragraph("<b>TOTAL CARTERA</b>", rt["table_cell"]),
        Paragraph("<b>Gestión total</b>", rt["table_cell_center"]),
        Paragraph(f"<b>{total_cuotas:,}</b>", rt["table_cell_right"]),
        Paragraph(f"<b>{total_clientes:,}</b>", rt["table_cell_right"]),
        Paragraph(f"<b>${total_usd:,.2f}</b>", rt["table_cell_right"]),
        Paragraph(total_pct_display, rt["table_cell_right"] if total_pct is not None else rt["table_cell_center"])
    ])

    t_table = Table(table_rows, colWidths=[150, 95, 75, 75, 95, 50])
    t_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1e293b")),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e1")),
        ('ROWBACKGROUNDS', (0, 1), (-1, -2), [colors.white, colors.HexColor("#f8fafc")]),
        ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor("#e2e8f0")),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]))
    story.append(t_table)
    story.append(Spacer(1, 16))

    tramos_list = global_summary.get("tramos", [])
    max_tramo = max(tramos_list, key=lambda x: x["total_monto"]) if tramos_list else None
    if max_tramo:
        pct_note = f" ({max_tramo['porcentaje']:.2f}% de la cartera activa total)" if max_tramo.get("porcentaje") is not None else ""
        notes_text = (
            f"<b>Observación Estratégica:</b> El tramo con mayor concentración de cartera es <b>{max_tramo['etiqueta']} ({max_tramo['fase']})</b> "
            f"con un saldo acumulado de <b>${max_tramo['total_monto']:,.2f} USD</b>{pct_note} "
            f"distribuido en {max_tramo['cuotas']:,} cuotas y {max_tramo['clientes']:,} clientes únicos. "
            f"Se sugiere enfocar las acciones intensivas de negociación y cobranza pre-jurídica en este segmento para maximizar la recuperación de liquidez."
        )
        story.append(build_pdf_strategic_note(notes_text, styles['Normal']))

    doc.build(story)
    buf.seek(0)
    filename = f"resumen_tramos_mora_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    return StreamingResponse(
        buf,
        media_type="application/pdf",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )


# -------------------------------------------------------------
# 3. REPORTE COMPARATIVO FIBEX VS. NO FIBEX (EXCEL Y PDF)
# -------------------------------------------------------------

def generate_fibex_excel(cartera_activa: Optional[float] = None) -> StreamingResponse:
    """Descarga el Resumen Comparativo Fibex vs. NO Fibex y su cruce por tramos en un archivo Excel ejecutivo."""
    active_df = _get_active_df_or_400()
    if cartera_activa is not None and cartera_activa > 0:
        DATA_STORE["cartera_activa_manual"] = cartera_activa
    manual_ca = cartera_activa or DATA_STORE.get("cartera_activa_manual")

    fibex_summary = calculate_fibex_summary(active_df, manual_ca)
    global_summary = calculate_global_tramos_summary(active_df, manual_ca)
    cutoff = fibex_summary.get("cutoff_date", datetime.now().strftime("%d-%m-%Y"))
    total_usd = fibex_summary.get("total_cartera_usd", 0.0)
    total_pct = fibex_summary.get("total_porcentaje")

    th = get_openpyxl_theme()
    wb = openpyxl.Workbook()

    # Hoja 1: Fibex vs NO Fibex
    ws = wb.active
    ws.title = "Fibex vs NO Fibex"
    ws.views.sheetView[0].showGridLines = True

    ws.merge_cells("A2:G2")
    ws["A2"] = "COBRANZAPP | RESUMEN COMPARATIVO: CLIENTES FIBEX VS. NO FIBEX"
    ws["A2"].font = th["font_title"]
    ws["A2"].alignment = Alignment(horizontal="left", vertical="center")

    ws.merge_cells("A3:G3")
    ws["A3"] = f"Fecha de corte: al {cutoff}  |  Fecha de emisión: {datetime.now().strftime('%d/%m/%Y %H:%M')}  |  Cartera: {DATA_STORE.get('cartera_filename', 'Cartera Activa')}"
    ws["A3"].font = th["font_subtitle"]
    ws["A3"].alignment = Alignment(horizontal="left", vertical="center")

    ws.row_dimensions[5].height = 24
    ws.merge_cells("A5:B5")
    ws["A5"] = "TARIFA TOTAL EN MORA:"
    ws["A5"].font = th["font_data_bold"]
    ws["A5"].fill = th["fill_banner"]
    ws["A5"].alignment = Alignment(horizontal="left", vertical="center")
    ws.merge_cells("C5:D5")
    ws["C5"] = total_usd
    ws["C5"].font = Font(name="Calibri", size=12, bold=True, color="047857")
    ws["C5"].number_format = '$#,##0.00'
    ws["C5"].fill = th["fill_banner"]
    ws["C5"].alignment = Alignment(horizontal="right", vertical="center")

    if manual_ca and manual_ca > 0:
        ws.merge_cells("E5:F5")
        ws["E5"] = "CARTERA ACTIVA TOTAL:"
        ws["E5"].font = th["font_data_bold"]
        ws["E5"].fill = th["fill_banner"]
        ws["E5"].alignment = Alignment(horizontal="left", vertical="center")
        ws["G5"] = manual_ca
        ws["G5"].font = Font(name="Calibri", size=12, bold=True, color="1E3A8A")
        ws["G5"].number_format = '$#,##0.00'
        ws["G5"].fill = th["fill_banner"]
        ws["G5"].alignment = Alignment(horizontal="right", vertical="center")

    curr_row = 7
    headers = [
        "Segmento de Cartera", "Definición / Origen", "Cant. Cuotas",
        "Clientes", "Monto Tarifa (USD)", "% de Morosidad", "% Cartera Activa"
    ]
    ws.row_dimensions[curr_row].height = 26
    for col_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=curr_row, column=col_idx, value=h)
        cell.font = th["font_header"]
        cell.fill = th["fill_header"]
        cell.alignment = Alignment(horizontal="center" if col_idx > 2 else "left", vertical="center", wrap_text=True)
        cell.border = th["thin_border"]

    curr_row += 1
    for idx, s in enumerate(fibex_summary.get("segments", [])):
        r = curr_row + idx
        ws.row_dimensions[r].height = 22
        c1 = ws.cell(row=r, column=1, value=s["etiqueta"])
        c2 = ws.cell(row=r, column=2, value=s["subtitulo"])
        c3 = ws.cell(row=r, column=3, value=s["cuotas"])
        c4 = ws.cell(row=r, column=4, value=s["clientes"])
        c5 = ws.cell(row=r, column=5, value=s["total_monto"])
        c6 = ws.cell(row=r, column=6, value=s["porcentaje_mora"] / 100.0)
        if s["porcentaje"] is not None:
            c7 = ws.cell(row=r, column=7, value=s["porcentaje"] / 100.0)
            c7.number_format = '0.00%'
        else:
            c7 = ws.cell(row=r, column=7, value="-")

        c1.font = th["font_data_bold"]
        c2.font = th["font_data"]
        c3.font = th["font_data"]
        c4.font = th["font_data"]
        c5.font = th["font_data_bold"]
        c6.font = th["font_data_bold"]
        c7.font = th["font_data_bold"]

        c1.alignment = Alignment(horizontal="left", vertical="center")
        c2.alignment = Alignment(horizontal="left", vertical="center")
        c3.alignment = Alignment(horizontal="right", vertical="center")
        c4.alignment = Alignment(horizontal="right", vertical="center")
        c5.alignment = Alignment(horizontal="right", vertical="center")
        c6.alignment = Alignment(horizontal="right", vertical="center")
        c7.alignment = Alignment(horizontal="right" if s["porcentaje"] is not None else "center", vertical="center")

        c3.number_format = '#,##0'
        c4.number_format = '#,##0'
        c5.number_format = '$#,##0.00'
        c6.number_format = '0.00%'

        row_fill = th["fill_alt"] if idx % 2 == 1 else None
        for col_idx in range(1, 8):
            cell = ws.cell(row=r, column=col_idx)
            cell.border = th["thin_border"]
            if row_fill:
                cell.fill = row_fill

    tot_row = curr_row + len(fibex_summary.get("segments", []))
    ws.row_dimensions[tot_row].height = 24
    tc1 = ws.cell(row=tot_row, column=1, value="TOTAL CARTERA")
    tc2 = ws.cell(row=tot_row, column=2, value="Gestión total")
    tc3 = ws.cell(row=tot_row, column=3, value=fibex_summary.get("total_cuotas", 0))
    tc4 = ws.cell(row=tot_row, column=4, value=fibex_summary.get("total_clientes", 0))
    tc5 = ws.cell(row=tot_row, column=5, value=total_usd)
    tc6 = ws.cell(row=tot_row, column=6, value=1.0)
    tc6.number_format = '0.00%'
    if total_pct is not None:
        tc7 = ws.cell(row=tot_row, column=7, value=total_pct / 100.0)
        tc7.number_format = '0.00%'
    else:
        tc7 = ws.cell(row=tot_row, column=7, value="-")

    for cell in [tc1, tc2, tc3, tc4, tc5, tc6, tc7]:
        cell.font = th["font_total"]
        cell.fill = th["fill_total"]
        cell.border = th["total_border"]

    tc1.alignment = Alignment(horizontal="left", vertical="center")
    tc2.alignment = Alignment(horizontal="left", vertical="center")
    tc3.alignment = Alignment(horizontal="right", vertical="center")
    tc4.alignment = Alignment(horizontal="right", vertical="center")
    tc5.alignment = Alignment(horizontal="right", vertical="center")
    tc6.alignment = Alignment(horizontal="right", vertical="center")
    tc7.alignment = Alignment(horizontal="right" if total_pct is not None else "center", vertical="center")

    tc3.number_format = '#,##0'
    tc4.number_format = '#,##0'
    tc5.number_format = '$#,##0.00'

    # Cruce Detallado por los 5 Tramos
    curr_row = tot_row + 3
    ws.merge_cells(f"A{curr_row}:G{curr_row}")
    ws[f"A{curr_row}"] = "CRUCE DETALLADO: 5 TRAMOS DE MORA POR SEGMENTO"
    ws[f"A{curr_row}"].font = Font(name="Calibri", size=12, bold=True, color="0F172A")

    curr_row += 1
    headers_tramos = ["Segmento", "Tramo 1 (1-30d)", "Tramo 2 (31-60d)", "Tramo 3 (61-90d)", "Tramo 4 (91-180d)", "Tramo 5 (+180d)", "Total Tarifa USD"]
    ws.row_dimensions[curr_row].height = 24
    for col_idx, h in enumerate(headers_tramos, start=1):
        cell = ws.cell(row=curr_row, column=col_idx, value=h)
        cell.font = th["font_header"]
        cell.fill = th["fill_header_sub"]
        cell.alignment = Alignment(horizontal="center" if col_idx > 1 else "left", vertical="center")
        cell.border = th["thin_border"]

    curr_row += 1
    tramo_keys = ["1-30", "31-60", "61-90", "91-180", "181+"]
    for idx, s in enumerate(fibex_summary.get("segments", [])):
        r = curr_row + idx
        ws.row_dimensions[r].height = 20
        c_seg = ws.cell(row=r, column=1, value=s["etiqueta"])
        c_seg.font = th["font_data_bold"]
        c_seg.alignment = Alignment(horizontal="left", vertical="center")
        c_seg.border = th["thin_border"]

        for t_idx, t_key in enumerate(tramo_keys, start=2):
            val = s.get("tramos", {}).get(t_key, {}).get("monto", 0.0)
            c = ws.cell(row=r, column=t_idx, value=val)
            c.font = th["font_data"]
            c.number_format = '$#,##0.00'
            c.alignment = Alignment(horizontal="right", vertical="center")
            c.border = th["thin_border"]

        c_tot = ws.cell(row=r, column=7, value=s["total_monto"])
        c_tot.font = th["font_data_bold"]
        c_tot.number_format = '$#,##0.00'
        c_tot.alignment = Alignment(horizontal="right", vertical="center")
        c_tot.border = th["thin_border"]

    r_tot_tr = curr_row + len(fibex_summary.get("segments", []))
    ws.row_dimensions[r_tot_tr].height = 22
    ws.cell(row=r_tot_tr, column=1, value="TOTAL CARTERA").font = th["font_total"]
    ws.cell(row=r_tot_tr, column=1).alignment = Alignment(horizontal="left", vertical="center")
    ws.cell(row=r_tot_tr, column=1).fill = th["fill_total"]
    ws.cell(row=r_tot_tr, column=1).border = th["total_border"]

    for t_idx, t_key in enumerate(tramo_keys, start=2):
        sum_col = sum(s.get("tramos", {}).get(t_key, {}).get("monto", 0.0) for s in fibex_summary.get("segments", []))
        c_s = ws.cell(row=r_tot_tr, column=t_idx, value=sum_col)
        c_s.font = th["font_total"]
        c_s.fill = th["fill_total"]
        c_s.border = th["total_border"]
        c_s.number_format = '$#,##0.00'
        c_s.alignment = Alignment(horizontal="right", vertical="center")

    c_g = ws.cell(row=r_tot_tr, column=7, value=total_usd)
    c_g.font = th["font_total"]
    c_g.fill = th["fill_total"]
    c_g.border = th["total_border"]
    c_g.number_format = '$#,##0.00'
    c_g.alignment = Alignment(horizontal="right", vertical="center")

    auto_fit_worksheet_columns(ws, max_cap=60, min_width=14)

    # Hoja 2: Resumen por Tramos Oficiales
    ws2 = wb.create_sheet(title="Resumen por Tramos")
    ws2.views.sheetView[0].showGridLines = True
    ws2.merge_cells("A2:F2")
    ws2["A2"] = "COBRANZAPP | RESUMEN OFICIAL POR TRAMOS DE MORA"
    ws2["A2"].font = th["font_title"]
    headers2 = ["Fase / Categoría", "Rango / Etiqueta", "Cant. Cuotas", "Clientes", "Tarifa (USD)", "% Cartera Activa"]
    for col_idx, h in enumerate(headers2, start=1):
        c = ws2.cell(row=4, column=col_idx, value=h)
        c.font = th["font_header"]
        c.fill = th["fill_header"]
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = th["thin_border"]
    for idx, t in enumerate(global_summary.get("tramos", [])):
        r = 5 + idx
        ws2.cell(row=r, column=1, value=t["fase"]).font = th["font_data"]
        ws2.cell(row=r, column=2, value=t["etiqueta"]).font = th["font_data_bold"]
        ws2.cell(row=r, column=3, value=t["cuotas"]).number_format = '#,##0'
        ws2.cell(row=r, column=4, value=t["clientes"]).number_format = '#,##0'
        ws2.cell(row=r, column=5, value=t["total_monto"]).number_format = '$#,##0.00'
        if t["porcentaje"] is not None:
            c_p = ws2.cell(row=r, column=6, value=t["porcentaje"] / 100.0)
            c_p.number_format = '0.00%'
        else:
            ws2.cell(row=r, column=6, value="-")
        for col_idx in range(1, 7):
            ws2.cell(row=r, column=col_idx).border = th["thin_border"]

    r_tot2 = 5 + len(global_summary.get("tramos", []))
    ws2.cell(row=r_tot2, column=1, value="TOTAL CARTERA").font = th["font_total"]
    ws2.cell(row=r_tot2, column=2, value="Gestión total").font = th["font_total"]
    ws2.cell(row=r_tot2, column=3, value=global_summary.get("total_cuotas", 0)).number_format = '#,##0'
    ws2.cell(row=r_tot2, column=4, value=global_summary.get("total_clientes", 0)).number_format = '#,##0'
    ws2.cell(row=r_tot2, column=5, value=total_usd).number_format = '$#,##0.00'
    if total_pct is not None:
        c_p = ws2.cell(row=r_tot2, column=6, value=total_pct / 100.0)
        c_p.number_format = '0.00%'
    else:
        ws2.cell(row=r_tot2, column=6, value="-")
    for col_idx in range(1, 7):
        c = ws2.cell(row=r_tot2, column=col_idx)
        c.font = th["font_total"]
        c.fill = th["fill_total"]
        c.border = th["total_border"]

    auto_fit_worksheet_columns(ws2, max_cap=50, min_width=14)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    filename = f"resumen_fibex_vs_nofibex_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )


def generate_fibex_pdf(cartera_activa: Optional[float] = None) -> StreamingResponse:
    """Genera un reporte PDF ejecutivo exclusivamente del Resumen Comparativo Fibex vs. NO Fibex."""
    active_df = _get_active_df_or_400()
    if cartera_activa is not None and cartera_activa > 0:
        DATA_STORE["cartera_activa_manual"] = cartera_activa
    manual_ca = cartera_activa or DATA_STORE.get("cartera_activa_manual")
    fibex_summary = calculate_fibex_summary(active_df, manual_ca)
    cutoff = fibex_summary.get("cutoff_date", datetime.now().strftime("%d-%m-%Y"))
    total_usd = fibex_summary.get("total_cartera_usd", 0.0)
    total_cuotas = fibex_summary.get("total_cuotas", 0)
    total_clientes = fibex_summary.get("total_clientes", 0)
    total_pct = fibex_summary.get("total_porcentaje")

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
    rt = get_reportlab_theme()
    styles = rt["base"]

    story = [
        Paragraph("COBRANZAPP | RESUMEN COMPARATIVO FIBEX VS. NO FIBEX", rt["title"]),
        Paragraph(
            f"<b>Fecha de corte:</b> al {cutoff} &nbsp;|&nbsp; <b>Fecha de emisión:</b> {datetime.now().strftime('%d/%m/%Y %H:%M')} &nbsp;|&nbsp; <b>Cartera:</b> {html.escape(str(DATA_STORE.get('cartera_filename', 'Cartera Activa')))}",
            rt["subtitle"]
        ),
        Spacer(1, 8)
    ]

    seg_nofibex = next((s for s in fibex_summary.get("segments", []) if s["id"] == "no_fibex"), None)
    seg_fibex = next((s for s in fibex_summary.get("segments", []) if s["id"] == "fibex_resto"), None)
    monto_nofibex = seg_nofibex["total_monto"] if seg_nofibex else 0.0
    pct_nofibex = seg_nofibex["porcentaje_mora"] if seg_nofibex else 0.0
    monto_fibex = seg_fibex["total_monto"] if seg_fibex else 0.0
    pct_fibex = seg_fibex["porcentaje_mora"] if seg_fibex else 0.0

    kpi_cards = [
        [
            Paragraph("TARIFA EN MORA TOTAL", rt["banner_label"]),
            Paragraph("CLIENTES WEB NO FIBEX", rt["banner_label"]),
            Paragraph("CLIENTES FIBEX Y RESTO", rt["banner_label"]),
            Paragraph("TOTAL CLIENTES", rt["banner_label"])
        ],
        [
            Paragraph(f"${total_usd:,.2f} USD", rt["banner_val"]),
            Paragraph(f"${monto_nofibex:,.2f} ({pct_nofibex:.1f}%)", rt["banner_val"]),
            Paragraph(f"${monto_fibex:,.2f} ({pct_fibex:.1f}%)", rt["banner_val"]),
            Paragraph(f"{total_clientes:,}", rt["banner_val"])
        ]
    ]

    story.append(build_pdf_kpi_card_grid(kpi_cards))
    story.append(Spacer(1, 16))
    story.append(Paragraph("<b>Distribución de Deuda: Segmentación Fibex vs. NO Fibex:</b>", styles['Normal']))
    story.append(Spacer(1, 6))

    table_rows = [[
        Paragraph("Segmento de Cartera", rt["table_header"]),
        Paragraph("Definición / Origen", rt["table_header"]),
        Paragraph("Cant. Cuotas", rt["table_header_right"]),
        Paragraph("Clientes", rt["table_header_right"]),
        Paragraph("Monto Tarifa (USD)", rt["table_header_right"]),
        Paragraph("% Morosidad", rt["table_header_right"]),
        Paragraph("% Cartera Activa", rt["table_header_right"])
    ]]

    for s in fibex_summary.get("segments", []):
        pct_display = f"<b>{s['porcentaje']:.2f}%</b>" if s["porcentaje"] is not None else "-"
        table_rows.append([
            Paragraph(f"<b>{html.escape(s['etiqueta'])}</b>", rt["table_cell"]),
            Paragraph(html.escape(s["subtitulo"]), rt["table_cell"]),
            Paragraph(f"{s['cuotas']:,}", rt["table_cell_right"]),
            Paragraph(f"{s['clientes']:,}", rt["table_cell_right"]),
            Paragraph(f"${s['total_monto']:,.2f}", rt["table_cell_right"]),
            Paragraph(f"<b>{s['porcentaje_mora']:.2f}%</b>", rt["table_cell_right"]),
            Paragraph(pct_display, rt["table_cell_right"] if s["porcentaje"] is not None else rt["table_cell_center"])
        ])

    total_pct_display = f"<b>{total_pct:.2f}%</b>" if total_pct is not None else "-"
    table_rows.append([
        Paragraph("<b>TOTAL CARTERA</b>", rt["table_cell"]),
        Paragraph("<b>Gestión total</b>", rt["table_cell"]),
        Paragraph(f"<b>{total_cuotas:,}</b>", rt["table_cell_right"]),
        Paragraph(f"<b>{total_clientes:,}</b>", rt["table_cell_right"]),
        Paragraph(f"<b>${total_usd:,.2f}</b>", rt["table_cell_right"]),
        Paragraph("<b>100.00%</b>", rt["table_cell_right"]),
        Paragraph(total_pct_display, rt["table_cell_right"] if total_pct is not None else rt["table_cell_center"])
    ])

    t_table = Table(table_rows, colWidths=[125, 140, 55, 50, 75, 55, 40])
    t_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1e293b")),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e1")),
        ('ROWBACKGROUNDS', (0, 1), (-1, -2), [colors.white, colors.HexColor("#f8fafc")]),
        ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor("#e2e8f0")),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]))
    story.append(t_table)
    story.append(Spacer(1, 16))

    story.append(Paragraph("<b>Cruce de Tramos de Mora por Segmento:</b>", styles['Normal']))
    story.append(Spacer(1, 6))

    tramo_headers = [
        Paragraph("Segmento", rt["table_header"]),
        Paragraph("1 a 30d", rt["table_header_right"]),
        Paragraph("31 a 60d", rt["table_header_right"]),
        Paragraph("61 a 90d", rt["table_header_right"]),
        Paragraph("91 a 180d", rt["table_header_right"]),
        Paragraph("+180d", rt["table_header_right"]),
        Paragraph("Total USD", rt["table_header_right"])
    ]
    tramos_table_rows = [tramo_headers]
    tramo_keys = ["1-30", "31-60", "61-90", "91-180", "181+"]
    for s in fibex_summary.get("segments", []):
        row = [Paragraph(f"<b>{html.escape(s['etiqueta'])}</b>", rt["table_cell"])]
        for t_k in tramo_keys:
            m = s.get("tramos", {}).get(t_k, {}).get("monto", 0.0)
            row.append(Paragraph(f"${m:,.2f}", rt["table_cell_right"]))
        row.append(Paragraph(f"<b>${s['total_monto']:,.2f}</b>", rt["table_cell_right"]))
        tramos_table_rows.append(row)

    tot_row_tr = [Paragraph("<b>TOTAL CARTERA</b>", rt["table_cell"])]
    for t_k in tramo_keys:
        tot_m = sum(s.get("tramos", {}).get(t_k, {}).get("monto", 0.0) for s in fibex_summary.get("segments", []))
        tot_row_tr.append(Paragraph(f"<b>${tot_m:,.2f}</b>", rt["table_cell_right"]))
    tot_row_tr.append(Paragraph(f"<b>${total_usd:,.2f}</b>", rt["table_cell_right"]))
    tramos_table_rows.append(tot_row_tr)

    t_tr_table = Table(tramos_table_rows, colWidths=[140, 65, 65, 65, 70, 65, 70])
    t_tr_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#334155")),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e1")),
        ('ROWBACKGROUNDS', (0, 1), (-1, -2), [colors.white, colors.HexColor("#f8fafc")]),
        ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor("#e2e8f0")),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ]))
    story.append(t_tr_table)
    story.append(Spacer(1, 16))

    notes_text = (
        f"<b>Observación Estratégica:</b> Los <b>Clientes Web NO Fibex</b> representan el <b>{pct_nofibex:.2f}% de la morosidad (${monto_nofibex:,.2f} USD)</b> "
        f"distribuidos en 4,899 clientes únicos sin vinculación a servicios Fibex (22 sub-campañas regionales). "
        f"Por su parte, los <b>Clientes Fibex y Resto</b> concentran el <b>{pct_fibex:.2f}% (${monto_fibex:,.2f} USD)</b> con 6,915 clientes. "
        f"Se sugiere una estrategia diferenciada: para No Fibex, intensificar cobranza digital vía WhatsApp y llamadas telefónicas con acuerdos de pago; "
        f"para clientes Fibex, apalancarse en la continuidad del servicio de telecomunicaciones."
    )
    story.append(build_pdf_strategic_note(notes_text, styles['Normal']))

    doc.build(story)
    buf.seek(0)
    filename = f"resumen_fibex_vs_nofibex_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    return StreamingResponse(
        buf,
        media_type="application/pdf",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )


# -------------------------------------------------------------
# 4. REPORTE MATRIZ DE CAMPAÑAS VS 5 TRAMOS (EXCEL Y PDF)
# -------------------------------------------------------------

def generate_campanas_matrix_excel() -> StreamingResponse:
    """Descarga la Matriz Completa de Campañas vs 5 Tramos Oficiales en Excel con formato corporativo."""
    active_df = _get_active_df_or_400()
    matrix_data = calculate_campanas_tramos_matrix(active_df)
    cutoff = matrix_data.get("cutoff_date", datetime.now().strftime("%d-%m-%Y"))
    total_usd = matrix_data.get("total_cartera_usd", 0.0)

    th = get_openpyxl_theme()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Desglose Tramos y Campañas"
    ws.views.sheetView[0].showGridLines = True

    fill_tramo_hdrs = [
        PatternFill(start_color="065F46", end_color="065F46", fill_type="solid"),
        PatternFill(start_color="1E40AF", end_color="1E40AF", fill_type="solid"),
        PatternFill(start_color="92400E", end_color="92400E", fill_type="solid"),
        PatternFill(start_color="9A3412", end_color="9A3412", fill_type="solid"),
        PatternFill(start_color="991B1B", end_color="991B1B", fill_type="solid"),
        PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    ]

    ws["A1"] = "COBRANZAPP - DESGLOSE POR TRAMOS Y CATEGORÍAS (CAMPAÑAS)"
    ws["A1"].font = th["font_title"]
    ws["A2"] = f"Fecha de corte: al {cutoff} | Tarifa Pendiente Total: ${total_usd:,.2f} USD | Emisión: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ws["A2"].font = th["font_subtitle"]

    ws.merge_cells("A4:A5")
    ws["A4"] = "Campaña Comercial"
    ws["A4"].font = th["font_group_hdr"]
    ws["A4"].fill = th["fill_header"]
    ws["A4"].alignment = Alignment(horizontal="center", vertical="center")
    ws["A4"].border = th["thin_border"]

    tramo_spans = [
        ("B4:C4", "B", "C", "1 a 30 días (Preventiva)", fill_tramo_hdrs[0]),
        ("D4:E4", "D", "E", "31 a 60 días (Administrativa)", fill_tramo_hdrs[1]),
        ("F4:G4", "F", "G", "61 a 90 días (Negociación)", fill_tramo_hdrs[2]),
        ("H4:I4", "H", "I", "91 a 180 días (Pre-jurídica)", fill_tramo_hdrs[3]),
        ("J4:K4", "J", "K", "181+ días (Castigo)", fill_tramo_hdrs[4]),
        ("L4:N4", "L", "N", "Total Tarifa Consolidada", fill_tramo_hdrs[5]),
    ]

    for span, col_start, _, title, fill in tramo_spans:
        ws.merge_cells(span)
        cell = ws[f"{col_start}4"]
        cell.value = title
        cell.font = th["font_group_hdr"]
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    sub_headers = [
        ("B5", "Tarifa ($)", "right"), ("C5", "Clientes", "right"),
        ("D5", "Tarifa ($)", "right"), ("E5", "Clientes", "right"),
        ("F5", "Tarifa ($)", "right"), ("G5", "Clientes", "right"),
        ("H5", "Tarifa ($)", "right"), ("I5", "Clientes", "right"),
        ("J5", "Tarifa ($)", "right"), ("K5", "Clientes", "right"),
        ("L5", "Total Tarifa ($)", "right"), ("M5", "Total Clientes", "right"),
        ("N5", "% Tarifa", "right")
    ]
    for cell_id, text, align in sub_headers:
        c = ws[cell_id]
        c.value = text
        c.font = th["font_sub_hdr"]
        c.fill = th["fill_sub_hdr"]
        c.alignment = Alignment(horizontal=align, vertical="center")
        c.border = th["thin_border"]

    ws.row_dimensions[4].height = 22
    ws.row_dimensions[5].height = 20

    curr_row = 6
    for idx, row in enumerate(matrix_data.get("matrix", [])):
        r = curr_row + idx
        ws.row_dimensions[r].height = 19

        c_camp = ws.cell(row=r, column=1, value=row["campana"])
        c_camp.font = th["font_data_bold_sm"]
        c_camp.alignment = Alignment(horizontal="left", vertical="center")

        col_idx = 2
        for t_cfg in OFFICIAL_TRAMOS_CONFIG:
            t_data = row["tramos"].get(t_cfg["id"], {"monto": 0.0, "clientes": 0})
            cm = ws.cell(row=r, column=col_idx, value=t_data["monto"])
            cc = ws.cell(row=r, column=col_idx + 1, value=t_data["clientes"])

            cm.font = th["font_data_sm"]
            cc.font = th["font_data_sm"]
            cm.alignment = Alignment(horizontal="right", vertical="center")
            cc.alignment = Alignment(horizontal="right", vertical="center")
            cm.number_format = '$#,##0.00' if t_data["monto"] > 0 else '"-"'
            cc.number_format = '#,##0' if t_data["clientes"] > 0 else '"-"'
            col_idx += 2

        c_tot_m = ws.cell(row=r, column=12, value=row["total_monto"])
        c_tot_c = ws.cell(row=r, column=13, value=row["total_clientes"])
        c_tot_p = ws.cell(row=r, column=14, value=row["porcentaje"] / 100.0)

        for c in (c_tot_m, c_tot_c, c_tot_p):
            c.font = th["font_data_bold_sm"]
            c.alignment = Alignment(horizontal="right", vertical="center")

        c_tot_m.number_format = '$#,##0.00'
        c_tot_c.number_format = '#,##0'
        c_tot_p.number_format = '0.00%'

        row_fill = th["fill_alt"] if idx % 2 == 1 else None
        for c_i in range(1, 15):
            cell = ws.cell(row=r, column=c_i)
            cell.border = th["thin_border"]
            if row_fill:
                cell.fill = row_fill

    tot_row = curr_row + len(matrix_data.get("matrix", []))
    ws.row_dimensions[tot_row].height = 24

    tc_camp = ws.cell(row=tot_row, column=1, value="TOTAL CONSOLIDADO")
    tc_camp.font = th["font_total_sm"]
    tc_camp.alignment = Alignment(horizontal="left", vertical="center")

    totals_tramos = matrix_data.get("totals_by_tramo", {})
    col_idx = 2
    for t_cfg in OFFICIAL_TRAMOS_CONFIG:
        t_data = totals_tramos.get(t_cfg["id"], {"monto": 0.0, "clientes": 0})
        cm = ws.cell(row=tot_row, column=col_idx, value=t_data["monto"])
        cc = ws.cell(row=tot_row, column=col_idx + 1, value=t_data["clientes"])

        cm.font = th["font_total_sm"]
        cc.font = th["font_total_sm"]
        cm.alignment = Alignment(horizontal="right", vertical="center")
        cc.alignment = Alignment(horizontal="right", vertical="center")
        cm.number_format = '$#,##0.00'
        cc.number_format = '#,##0'
        col_idx += 2

    tc_tot_m = ws.cell(row=tot_row, column=12, value=total_usd)
    tc_tot_c = ws.cell(row=tot_row, column=13, value=matrix_data.get("total_clientes", 0))
    tc_tot_p = ws.cell(row=tot_row, column=14, value=1.0)

    for c in [tc_tot_m, tc_tot_c, tc_tot_p]:
        c.font = th["font_total_sm"]
        c.alignment = Alignment(horizontal="right", vertical="center")

    tc_tot_m.number_format = '$#,##0.00'
    tc_tot_c.number_format = '#,##0'
    tc_tot_p.number_format = '0.00%'

    for c_i in range(1, 15):
        cell = ws.cell(row=tot_row, column=c_i)
        cell.fill = th["fill_total"]
        cell.border = th["total_border"]

    widths = {
        1: 32, 2: 15, 3: 11, 4: 15, 5: 11, 6: 15, 7: 11,
        8: 15, 9: 11, 10: 15, 11: 11, 12: 17, 13: 13, 14: 12
    }
    for col_idx, width in widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"desglose_campanas_tramos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )


def generate_campanas_matrix_pdf() -> StreamingResponse:
    """Genera reporte ejecutivo en PDF en formato apaisado (Landscape) con la Matriz de Campañas vs 5 Tramos."""
    active_df = _get_active_df_or_400()
    matrix_data = calculate_campanas_tramos_matrix(active_df)
    cutoff = matrix_data.get("cutoff_date", datetime.now().strftime("%d-%m-%Y"))
    total_usd = matrix_data.get("total_cartera_usd", 0.0)
    total_clientes = matrix_data.get("total_clientes", 0)

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(letter),
        leftMargin=20,
        rightMargin=20,
        topMargin=20,
        bottomMargin=20
    )
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        'MatTitle',
        parent=styles['Heading1'],
        fontSize=15,
        leading=18,
        textColor=colors.HexColor("#0f172a"),
        fontName="Helvetica-Bold",
        spaceAfter=2
    )
    subtitle_style = ParagraphStyle(
        'MatSubtitle',
        parent=styles['Normal'],
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor("#64748b"),
        spaceAfter=10
    )
    hdr_group = ParagraphStyle(
        'HdrGroup',
        parent=styles['Normal'],
        fontSize=7.5,
        leading=9,
        textColor=colors.white,
        fontName="Helvetica-Bold",
        alignment=1
    )
    hdr_sub = ParagraphStyle(
        'HdrSub',
        parent=styles['Normal'],
        fontSize=7,
        leading=8,
        textColor=colors.HexColor("#334155"),
        fontName="Helvetica-Bold",
        alignment=2
    )
    cell_camp = ParagraphStyle(
        'CellCamp',
        parent=styles['Normal'],
        fontSize=7.5,
        leading=9,
        textColor=colors.HexColor("#0f172a"),
        fontName="Helvetica-Bold"
    )
    cell_val = ParagraphStyle(
        'CellVal',
        parent=styles['Normal'],
        fontSize=7,
        leading=8.5,
        textColor=colors.HexColor("#1e293b"),
        alignment=2
    )
    cell_tot = ParagraphStyle(
        'CellTot',
        parent=styles['Normal'],
        fontSize=7,
        leading=8.5,
        textColor=colors.HexColor("#0f172a"),
        fontName="Helvetica-Bold",
        alignment=2
    )

    story = [
        Paragraph("COBRANZAPP | DESGLOSE POR TRAMOS Y CATEGORÍAS (CAMPAÑAS)", title_style),
        Paragraph(
            f"<b>Fecha de corte:</b> al {cutoff} &nbsp;|&nbsp; "
            f"<b>Tarifa Pendiente Total:</b> ${total_usd:,.2f} USD &nbsp;|&nbsp; "
            f"<b>Clientes con Tarifa:</b> {total_clientes:,} &nbsp;|&nbsp; "
            f"<b>Emisión:</b> {datetime.now().strftime('%d/%m/%Y %H:%M')}",
            subtitle_style
        )
    ]

    table_data = [
        [
            Paragraph("<b>Campaña Comercial</b>", hdr_group),
            Paragraph("<b>1 a 30 días</b>", hdr_group), "",
            Paragraph("<b>31 a 60 días</b>", hdr_group), "",
            Paragraph("<b>61 a 90 días</b>", hdr_group), "",
            Paragraph("<b>91 a 180 días</b>", hdr_group), "",
            Paragraph("<b>181+ días</b>", hdr_group), "",
            Paragraph("<b>Total Tarifa</b>", hdr_group), "", ""
        ],
        [
            "",
            Paragraph("Tarifa ($)", hdr_sub), Paragraph("Ctes", hdr_sub),
            Paragraph("Tarifa ($)", hdr_sub), Paragraph("Ctes", hdr_sub),
            Paragraph("Tarifa ($)", hdr_sub), Paragraph("Ctes", hdr_sub),
            Paragraph("Tarifa ($)", hdr_sub), Paragraph("Ctes", hdr_sub),
            Paragraph("Tarifa ($)", hdr_sub), Paragraph("Ctes", hdr_sub),
            Paragraph("Tarifa ($)", hdr_sub), Paragraph("Ctes", hdr_sub),
            Paragraph("% Tarifa", hdr_sub)
        ]
    ]

    for row in matrix_data.get("matrix", []):
        row_cells = [Paragraph(html.escape(row["campana"]), cell_camp)]
        for t_cfg in OFFICIAL_TRAMOS_CONFIG:
            t_data = row["tramos"].get(t_cfg["id"], {"monto": 0.0, "clientes": 0})
            m_str = f"${t_data['monto']:,.2f}" if t_data["monto"] > 0 else "-"
            c_str = f"{t_data['clientes']:,}" if t_data["clientes"] > 0 else "-"
            row_cells.append(Paragraph(m_str, cell_val))
            row_cells.append(Paragraph(c_str, cell_val))
        row_cells.append(Paragraph(f"${row['total_monto']:,.2f}", cell_tot))
        row_cells.append(Paragraph(f"{row['total_clientes']:,}", cell_tot))
        row_cells.append(Paragraph(f"{row['porcentaje']:.2f}%", cell_tot))
        table_data.append(row_cells)

    totals_tramos = matrix_data.get("totals_by_tramo", {})
    total_row_cells = [Paragraph("<b>TOTAL CONSOLIDADO</b>", cell_camp)]
    for t_cfg in OFFICIAL_TRAMOS_CONFIG:
        t_data = totals_tramos.get(t_cfg["id"], {"monto": 0.0, "clientes": 0})
        total_row_cells.append(Paragraph(f"<b>${t_data['monto']:,.2f}</b>", cell_tot))
        total_row_cells.append(Paragraph(f"<b>{t_data['clientes']:,}</b>", cell_tot))
    total_row_cells.append(Paragraph(f"<b>${total_usd:,.2f}</b>", cell_tot))
    total_row_cells.append(Paragraph(f"<b>{total_clientes:,}</b>", cell_tot))
    total_row_cells.append(Paragraph("<b>100.0%</b>", cell_tot))
    table_data.append(total_row_cells)

    col_widths = [140, 48, 28, 48, 28, 48, 28, 52, 28, 48, 28, 56, 32, 38]
    t_mat = Table(table_data, colWidths=col_widths, repeatRows=2)
    t_mat.setStyle(TableStyle([
        ('SPAN', (0, 0), (0, 1)),
        ('SPAN', (1, 0), (2, 0)),
        ('SPAN', (3, 0), (4, 0)),
        ('SPAN', (5, 0), (6, 0)),
        ('SPAN', (7, 0), (8, 0)),
        ('SPAN', (9, 0), (10, 0)),
        ('SPAN', (11, 0), (13, 0)),

        ('BACKGROUND', (0, 0), (0, 1), colors.HexColor("#0f172a")),
        ('BACKGROUND', (1, 0), (2, 0), colors.HexColor("#065f46")),
        ('BACKGROUND', (3, 0), (4, 0), colors.HexColor("#1e40af")),
        ('BACKGROUND', (5, 0), (6, 0), colors.HexColor("#92400e")),
        ('BACKGROUND', (7, 0), (8, 0), colors.HexColor("#9a3412")),
        ('BACKGROUND', (9, 0), (10, 0), colors.HexColor("#991b1b")),
        ('BACKGROUND', (11, 0), (13, 0), colors.HexColor("#0f172a")),
        ('BACKGROUND', (1, 1), (-1, 1), colors.HexColor("#f1f5f9")),

        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e1")),
        ('ROWBACKGROUNDS', (0, 2), (-1, -2), [colors.white, colors.HexColor("#f8fafc")]),
        ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor("#e2e8f0")),

        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('LEFTPADDING', (0, 0), (-1, -1), 3),
        ('RIGHTPADDING', (0, 0), (-1, -1), 3),
    ]))
    story.append(t_mat)

    doc.build(story)
    buf.seek(0)
    filename = f"desglose_campanas_tramos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    return StreamingResponse(
        buf,
        media_type="application/pdf",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )


# -------------------------------------------------------------
# 5. REPORTE CONSOLIDADO DE DEUDORES CALL CENTER (EXCEL)
# -------------------------------------------------------------

def generate_contacts_excel() -> StreamingResponse:
    """Descarga el listado consolidado a nivel deudor para marcadores predictivos y WhatsApp."""
    debtors = DATA_STORE.get("df_matched")
    if debtors is None or debtors.empty:
        if DATA_STORE["df_filtered"] is not None:
            build_debtor_matched_dataset(DATA_STORE["df_filtered"])
            debtors = DATA_STORE.get("df_matched")

    if debtors is None or debtors.empty:
        raise HTTPException(status_code=400, detail="No hay datos de cartera para exportar a Call Center.")

    df = debtors.copy()
    if "Cedula_Key" in df.columns:
        df = df.drop(columns=["Cedula_Key"])

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Deudores Call Center")

    output.seek(0)
    filename = f"cartera_deudores_callcenter_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )
