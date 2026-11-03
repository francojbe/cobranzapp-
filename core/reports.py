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
    calculate_3_tramos_summary,
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


def generate_tramos_3_excel(cartera_activa: Optional[float] = None) -> StreamingResponse:
    """Descarga el Resumen Agrupado por 3 Tramos Contables de Mora (Irregular, Vencido Gestión, Vencido Contable) en un archivo Excel."""
    active_df = _get_active_df_or_400()
    if cartera_activa is not None and cartera_activa > 0:
        DATA_STORE["cartera_activa_manual"] = cartera_activa
    manual_ca = cartera_activa or DATA_STORE.get("cartera_activa_manual")
    summary_3 = calculate_3_tramos_summary(active_df, manual_ca)
    cutoff = summary_3.get("cutoff_date", datetime.now().strftime("%d-%m-%Y"))
    total_usd = summary_3.get("total_cartera_usd", 0.0)
    total_pct = summary_3.get("total_porcentaje")

    th = get_openpyxl_theme()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resumen 3 Tramos"
    ws.views.sheetView[0].showGridLines = True

    ws["A1"] = "COBRANZAPP - RESUMEN DE CARTERA EN MORA (3 TRAMOS CONTABLES)"
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

    headers = ["Categoría", "Tramo / Días de Mora", "Cant. Cuotas", "Clientes", "Monto (USD)", "% Cartera Activa"]
    start_row = 6
    for col_idx, h_text in enumerate(headers, start=1):
        cell = ws.cell(row=start_row, column=col_idx, value=h_text)
        cell.font = th["font_header"]
        cell.fill = th["fill_header"]
        cell.alignment = Alignment(horizontal="center" if col_idx in [2, 6] else ("right" if col_idx in [3, 4, 5] else "left"), vertical="center")
        cell.border = th["thin_border"]
    ws.row_dimensions[start_row].height = 24

    fill_yellow_pct = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")
    fill_red_pct = PatternFill(start_color="DC2626", end_color="DC2626", fill_type="solid")
    font_red_pct = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

    curr_row = start_row + 1
    for idx, t in enumerate(summary_3.get("tramos", [])):
        r = curr_row + idx
        ws.row_dimensions[r].height = 22
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

        c1.font = th["font_data_bold"]
        c2.font = th["font_data"]
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

        # Semáforo en la columna % según categoría contable
        if (t.get("id") or t.get("key")) in ("1-15", "16-30"):
            c6.fill = fill_yellow_pct
        elif (t.get("id") or t.get("key")) == "31+":
            c6.fill = fill_red_pct
            c6.font = font_red_pct

    tot_row = curr_row + len(summary_3.get("tramos", []))
    ws.row_dimensions[tot_row].height = 24

    tc1 = ws.cell(row=tot_row, column=1, value="Total")
    tc2 = ws.cell(row=tot_row, column=2, value="Gestión total")
    tc3 = ws.cell(row=tot_row, column=3, value=summary_3.get("total_cuotas", 0))
    tc4 = ws.cell(row=tot_row, column=4, value=summary_3.get("total_clientes", 0))
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

    col_widths = {1: 24, 2: 28, 3: 16, 4: 16, 5: 22, 6: 16}
    for col_idx, width in col_widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"resumen_3_tramos_mora_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )


def generate_tramos_3_pdf(cartera_activa: Optional[float] = None) -> StreamingResponse:
    """Genera un reporte PDF ejecutivo del Resumen Agrupado en 3 Tramos Contables de Mora."""
    active_df = _get_active_df_or_400()
    if cartera_activa is not None and cartera_activa > 0:
        DATA_STORE["cartera_activa_manual"] = cartera_activa
    manual_ca = cartera_activa or DATA_STORE.get("cartera_activa_manual")
    summary_3 = calculate_3_tramos_summary(active_df, manual_ca)
    cutoff = summary_3.get("cutoff_date", datetime.now().strftime("%d-%m-%Y"))
    total_usd = summary_3.get("total_cartera_usd", 0.0)
    total_cuotas = summary_3.get("total_cuotas", 0)
    total_clientes = summary_3.get("total_clientes", 0)
    total_pct = summary_3.get("total_porcentaje")

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
    rt = get_reportlab_theme()
    styles = rt["base"]

    story = [
        Paragraph("COBRANZAPP | RESUMEN DE CARTERA EN MORA (3 TRAMOS CONTABLES)", rt["title"]),
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
    story.append(Paragraph("<b>Distribución de Cartera en Mora por Clasificación Contable (3 Tramos):</b>", styles['Normal']))
    story.append(Spacer(1, 6))

    table_rows = [[
        Paragraph("Categoría", rt["table_header"]),
        Paragraph("Tramo / Días de Mora", rt["table_header_center"]),
        Paragraph("Cant. Cuotas", rt["table_header_right"]),
        Paragraph("Clientes", rt["table_header_right"]),
        Paragraph("Monto (USD)", rt["table_header_right"]),
        Paragraph("% Cartera Activa", rt["table_header_right"])
    ]]

    white_bold_right = ParagraphStyle(
        "TableCellWhiteBoldRight",
        parent=rt["table_cell_right"],
        textColor=colors.white,
    )

    tramos_list = summary_3.get("tramos", [])
    for t in tramos_list:
        pct_display = f"<b>{t['porcentaje']:.2f}%</b>" if t["porcentaje"] is not None else "-"
        pct_style = white_bold_right if ((t.get("id") or t.get("key")) == "31+" and t["porcentaje"] is not None) else (
            rt["table_cell_right"] if t["porcentaje"] is not None else rt["table_cell_center"]
        )
        table_rows.append([
            Paragraph(f"<b>{html.escape(t['fase'])}</b>", rt["table_cell"]),
            Paragraph(html.escape(t["etiqueta"]), rt["table_cell_center"]),
            Paragraph(f"{t['cuotas']:,}", rt["table_cell_right"]),
            Paragraph(f"{t['clientes']:,}", rt["table_cell_right"]),
            Paragraph(f"${t['total_monto']:,.2f}", rt["table_cell_right"]),
            Paragraph(pct_display, pct_style)
        ])

    total_pct_display = f"<b>{total_pct:.2f}%</b>" if total_pct is not None else "-"
    table_rows.append([
        Paragraph("<b>Total</b>", rt["table_cell"]),
        Paragraph("<b>Gestión total</b>", rt["table_cell_center"]),
        Paragraph(f"<b>{total_cuotas:,}</b>", rt["table_cell_right"]),
        Paragraph(f"<b>{total_clientes:,}</b>", rt["table_cell_right"]),
        Paragraph(f"<b>${total_usd:,.2f}</b>", rt["table_cell_right"]),
        Paragraph(total_pct_display, rt["table_cell_right"] if total_pct is not None else rt["table_cell_center"])
    ])

    t_table = Table(table_rows, colWidths=[115, 130, 75, 75, 95, 50])
    style_cmds = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1e293b")),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e1")),
        ('ROWBACKGROUNDS', (0, 1), (-1, -2), [colors.white, colors.HexColor("#f8fafc")]),
        ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor("#e2e8f0")),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]
    for idx, t in enumerate(tramos_list, start=1):
        if (t.get("id") or t.get("key")) in ("1-15", "16-30"):
            style_cmds.append(('BACKGROUND', (5, idx), (5, idx), colors.HexColor("#fef08a")))
        elif (t.get("id") or t.get("key")) == "31+":
            style_cmds.append(('BACKGROUND', (5, idx), (5, idx), colors.HexColor("#dc2626")))

    t_table.setStyle(TableStyle(style_cmds))
    story.append(t_table)
    story.append(Spacer(1, 16))

    vencido_contable = next((x for x in tramos_list if (x.get("id") or x.get("key")) == "31+"), None)
    if vencido_contable:
        pct_note = f" situándose en <b>{vencido_contable['porcentaje']:.2f}%</b> sobre la cartera activa" if vencido_contable.get("porcentaje") is not None else ""
        notes_text = (
            f"<b>Observación Contable:</b> El tramo <b>Vencido contable (Más de 30 días)</b> concentra "
            f"<b>${vencido_contable['total_monto']:,.2f} USD</b>{pct_note}, "
            f"con {vencido_contable['cuotas']:,} cuotas pendientes y {vencido_contable['clientes']:,} clientes únicos. "
            f"Este indicador refleja la morosidad contable mayor a 30 días sobre el total de la cartera activa."
        )
        story.append(build_pdf_strategic_note(notes_text, styles['Normal']))

    doc.build(story)
    buf.seek(0)
    filename = f"resumen_3_tramos_mora_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
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
# 4. REPORTE MATRIZ DE CAMPAÑAS VS TRAMOS (5 Y 3 TRAMOS + % CARTERA ACTIVA)
# -------------------------------------------------------------

def generate_campanas_matrix_excel(mode: str = "5tramos") -> StreamingResponse:
    """Descarga la Matriz Completa de Campañas vs 5 Tramos Oficiales o 3 Tramos Contables en Excel,
    incluyendo las 3 columnas de Morosidad contra Cartera Activa (Fin de Mes, Día Anterior y Variación)."""
    active_df = _get_active_df_or_400()
    matrix_data = calculate_campanas_tramos_matrix(active_df)
    cutoff = matrix_data.get("cutoff_date", datetime.now().strftime("%d-%m-%Y"))
    total_usd = matrix_data.get("total_cartera_usd", 0.0)
    lbl_fm = matrix_data.get("label_fin_mes", "Fin de Mes")
    lbl_da = matrix_data.get("label_dia_anterior", "Día Anterior")
    is_3 = (str(mode or "").lower() in ("3", "3tramos", "tramos3"))

    th = get_openpyxl_theme()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Matriz 3 Tramos" if is_3 else "Matriz 5 Tramos"
    ws.views.sheetView[0].showGridLines = True

    fill_orange_hdr = PatternFill(start_color="C2410C", end_color="C2410C", fill_type="solid")
    fill_orange_sub = PatternFill(start_color="FFEDD5", end_color="FFEDD5", fill_type="solid")
    fill_amber_cell = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    font_var_pos = Font(name="Calibri", size=10, bold=True, color="B91C1C")
    font_var_neg = Font(name="Calibri", size=10, bold=True, color="15803D")

    title_mode = "3 TRAMOS CONTABLES" if is_3 else "5 TRAMOS OPERATIVOS"
    ws["A1"] = f"COBRANZAPP - RESUMEN CARTERA EN MORA POR CAMPAÑA ({title_mode})"
    ws["A1"].font = th["font_title"]
    ws["A2"] = f"Fecha de corte: al {cutoff} | Mora Total: ${total_usd:,.2f} USD | Emisión: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ws["A2"].font = th["font_subtitle"]

    ws.merge_cells("A4:A5")
    ws["A4"] = "Campaña Comercial"
    ws["A4"].font = th["font_group_hdr"]
    ws["A4"].fill = th["fill_header"]
    ws["A4"].alignment = Alignment(horizontal="center", vertical="center")
    ws["A4"].border = th["thin_border"]

    if is_3:
        tramo_defs = [
            ("31+", "Vencido contable (+30 días)", PatternFill(start_color="991B1B", end_color="991B1B", fill_type="solid")),
            ("16-30", "Vencido Gestión (16 a 30 días)", PatternFill(start_color="B45309", end_color="B45309", fill_type="solid")),
            ("1-15", "Vencido Irregular (1 a 15 días)", PatternFill(start_color="1E40AF", end_color="1E40AF", fill_type="solid")),
        ]
    else:
        tramo_defs = [
            ("1-30", "1 a 30 días (Preventiva)", PatternFill(start_color="065F46", end_color="065F46", fill_type="solid")),
            ("31-60", "31 a 60 días (Administrativa)", PatternFill(start_color="1E40AF", end_color="1E40AF", fill_type="solid")),
            ("61-90", "61 a 90 días (Negociación)", PatternFill(start_color="92400E", end_color="92400E", fill_type="solid")),
            ("91-180", "91 a 180 días (Pre-jurídica)", PatternFill(start_color="9A3412", end_color="9A3412", fill_type="solid")),
            ("181+", "181+ días (Castigo)", PatternFill(start_color="991B1B", end_color="991B1B", fill_type="solid")),
        ]

    col_cursor = 2
    for _, t_title, t_fill in tramo_defs:
        c_start = get_column_letter(col_cursor)
        c_end = get_column_letter(col_cursor + 1)
        ws.merge_cells(f"{c_start}4:{c_end}4")
        cell = ws[f"{c_start}4"]
        cell.value = t_title
        cell.font = th["font_group_hdr"]
        cell.fill = t_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

        for sub_idx, sub_txt in enumerate(["Monto (USD)", "Nro. Clientes"]):
            sc = ws.cell(row=5, column=col_cursor + sub_idx, value=sub_txt)
            sc.font = th["font_sub_hdr"]
            sc.fill = th["fill_sub_hdr"]
            sc.alignment = Alignment(horizontal="right", vertical="center")
            sc.border = th["thin_border"]
        col_cursor += 2

    # Bloque Total Mora Consolidada (3 columnas: Total Monto, Total Clientes, % Mora)
    tot_start_col = col_cursor
    c_tot_s = get_column_letter(tot_start_col)
    c_tot_e = get_column_letter(tot_start_col + 2)
    ws.merge_cells(f"{c_tot_s}4:{c_tot_e}4")
    c_tot_hdr = ws[f"{c_tot_s}4"]
    c_tot_hdr.value = "Total Mora Consolidada"
    c_tot_hdr.font = th["font_group_hdr"]
    c_tot_hdr.fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    c_tot_hdr.alignment = Alignment(horizontal="center", vertical="center")

    for sub_idx, sub_txt in enumerate(["Total Monto (USD)", "Total Clientes", "% del Total Mora"]):
        sc = ws.cell(row=5, column=tot_start_col + sub_idx, value=sub_txt)
        sc.font = th["font_sub_hdr"]
        sc.fill = th["fill_sub_hdr"]
        sc.alignment = Alignment(horizontal="right", vertical="center")
        sc.border = th["thin_border"]
    col_cursor += 3

    # Bloque Naranja Final: Morosidad vs. Cartera Activa (Fin de Mes, Día Anterior y Variación)
    ca_start_col = col_cursor
    c_ca_s = get_column_letter(ca_start_col)
    c_ca_e = get_column_letter(ca_start_col + 2)
    ws.merge_cells(f"{c_ca_s}4:{c_ca_e}4")
    c_ca_hdr = ws[f"{c_ca_s}4"]
    c_ca_hdr.value = "% Morosidad sobre Cartera Activa"
    c_ca_hdr.font = th["font_group_hdr"]
    c_ca_hdr.fill = fill_orange_hdr
    c_ca_hdr.alignment = Alignment(horizontal="center", vertical="center")

    for sub_idx, sub_txt in enumerate([f"Morosidad al {lbl_fm}", f"Morosidad al {lbl_da}", "Variación"]):
        sc = ws.cell(row=5, column=ca_start_col + sub_idx, value=sub_txt)
        sc.font = Font(name="Calibri", size=9, bold=True, color="7C2D12")
        sc.fill = fill_orange_sub
        sc.alignment = Alignment(horizontal="right", vertical="center")
        sc.border = th["thin_border"]
    total_cols = ca_start_col + 2

    ws.row_dimensions[4].height = 22
    ws.row_dimensions[5].height = 20

    curr_row = 6
    row_dict_key = "tramos_3" if is_3 else "tramos"
    for idx, row in enumerate(matrix_data.get("matrix", [])):
        r = curr_row + idx
        ws.row_dimensions[r].height = 19

        c_camp = ws.cell(row=r, column=1, value=row["campana"])
        c_camp.font = th["font_data_bold_sm"]
        c_camp.alignment = Alignment(horizontal="left", vertical="center")

        c_idx = 2
        row_tramos = row.get(row_dict_key, {})
        for t_key, _, _ in tramo_defs:
            t_data = row_tramos.get(t_key, {"monto": 0.0, "clientes": 0})
            cm = ws.cell(row=r, column=c_idx, value=t_data["monto"])
            cc = ws.cell(row=r, column=c_idx + 1, value=t_data["clientes"])

            cm.font = th["font_data_sm"]
            cc.font = th["font_data_sm"]
            cm.alignment = Alignment(horizontal="right", vertical="center")
            cc.alignment = Alignment(horizontal="right", vertical="center")
            cm.number_format = '$#,##0.00' if t_data["monto"] > 0 else '"-"'
            cc.number_format = '#,##0' if t_data["clientes"] > 0 else '"-"'
            c_idx += 2

        c_tot_m = ws.cell(row=r, column=tot_start_col, value=row["total_monto"])
        c_tot_c = ws.cell(row=r, column=tot_start_col + 1, value=row["total_clientes"])
        c_tot_p = ws.cell(row=r, column=tot_start_col + 2, value=row["porcentaje"] / 100.0)

        for c in (c_tot_m, c_tot_c, c_tot_p):
            c.font = th["font_data_bold_sm"]
            c.alignment = Alignment(horizontal="right", vertical="center")

        c_tot_m.number_format = '$#,##0.00'
        c_tot_c.number_format = '#,##0'
        c_tot_p.number_format = '0.00%'

        # 3 Columnas Naranjas: % Fin de Mes, % Día Anterior, Variación
        pfm = row.get("pct_fin_mes")
        pda = row.get("pct_dia_anterior")
        pvar = row.get("variacion_pct")

        c_fm = ws.cell(row=r, column=ca_start_col, value=(pfm / 100.0) if pfm is not None else "-")
        c_da = ws.cell(row=r, column=ca_start_col + 1, value=(pda / 100.0) if pda is not None else "-")
        c_var = ws.cell(row=r, column=ca_start_col + 2, value=(pvar / 100.0) if pvar is not None else "-")

        for c, val in [(c_fm, pfm), (c_da, pda), (c_var, pvar)]:
            c.font = th["font_data_bold_sm"]
            c.alignment = Alignment(horizontal="right" if val is not None else "center", vertical="center")
            if val is not None:
                c.number_format = '0.00%'

        if pvar is not None:
            c_var.font = font_var_neg if pvar <= 0 else font_var_pos

        row_fill = th["fill_alt"] if idx % 2 == 1 else None
        for c_i in range(1, total_cols + 1):
            cell = ws.cell(row=r, column=c_i)
            cell.border = th["thin_border"]
            if c_i >= ca_start_col:
                cell.fill = fill_amber_cell
            elif row_fill:
                cell.fill = row_fill

    tot_row = curr_row + len(matrix_data.get("matrix", []))
    ws.row_dimensions[tot_row].height = 24

    tc_camp = ws.cell(row=tot_row, column=1, value="TOTAL CONSOLIDADO")
    tc_camp.font = th["font_total_sm"]
    tc_camp.alignment = Alignment(horizontal="left", vertical="center")

    totals_tramos = matrix_data.get("totals_by_tramo_3" if is_3 else "totals_by_tramo", {})
    c_idx = 2
    for t_key, _, _ in tramo_defs:
        t_data = totals_tramos.get(t_key, {"monto": 0.0, "clientes": 0})
        cm = ws.cell(row=tot_row, column=c_idx, value=t_data["monto"])
        cc = ws.cell(row=tot_row, column=c_idx + 1, value=t_data["clientes"])

        cm.font = th["font_total_sm"]
        cc.font = th["font_total_sm"]
        cm.alignment = Alignment(horizontal="right", vertical="center")
        cc.alignment = Alignment(horizontal="right", vertical="center")
        cm.number_format = '$#,##0.00'
        cc.number_format = '#,##0'
        c_idx += 2

    tc_tot_m = ws.cell(row=tot_row, column=tot_start_col, value=total_usd)
    tc_tot_c = ws.cell(row=tot_row, column=tot_start_col + 1, value=matrix_data.get("total_clientes", 0))
    tc_tot_p = ws.cell(row=tot_row, column=tot_start_col + 2, value=1.0)

    for c in [tc_tot_m, tc_tot_c, tc_tot_p]:
        c.font = th["font_total_sm"]
        c.alignment = Alignment(horizontal="right", vertical="center")

    tc_tot_m.number_format = '$#,##0.00'
    tc_tot_c.number_format = '#,##0'
    tc_tot_p.number_format = '0.00%'

    t_pfm = matrix_data.get("total_pct_fin_mes")
    t_pda = matrix_data.get("total_pct_dia_anterior")
    t_pvar = matrix_data.get("total_variacion_pct")

    tc_fm = ws.cell(row=tot_row, column=ca_start_col, value=(t_pfm / 100.0) if t_pfm is not None else "-")
    tc_da = ws.cell(row=tot_row, column=ca_start_col + 1, value=(t_pda / 100.0) if t_pda is not None else "-")
    tc_var = ws.cell(row=tot_row, column=ca_start_col + 2, value=(t_pvar / 100.0) if t_pvar is not None else "-")

    for c, val in [(tc_fm, t_pfm), (tc_da, t_pda), (tc_var, t_pvar)]:
        c.font = th["font_total_sm"]
        c.alignment = Alignment(horizontal="right" if val is not None else "center", vertical="center")
        if val is not None:
            c.number_format = '0.00%'
    if t_pvar is not None:
        tc_var.font = font_var_neg if t_pvar <= 0 else font_var_pos

    for c_i in range(1, total_cols + 1):
        cell = ws.cell(row=tot_row, column=c_i)
        cell.fill = th["fill_total"]
        cell.border = th["total_border"]

    ws.column_dimensions["A"].width = 32
    for c_i in range(2, total_cols + 1):
        ws.column_dimensions[get_column_letter(c_i)].width = 14

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    suffix = "3_tramos" if is_3 else "5_tramos"
    filename = f"matriz_campanas_{suffix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )


def generate_campanas_matrix_pdf(mode: str = "5tramos") -> StreamingResponse:
    """Genera reporte ejecutivo en PDF apaisado (Landscape) con la Matriz de Campañas (5 o 3 Tramos + % Cartera Activa)."""
    active_df = _get_active_df_or_400()
    matrix_data = calculate_campanas_tramos_matrix(active_df)
    cutoff = matrix_data.get("cutoff_date", datetime.now().strftime("%d-%m-%Y"))
    total_usd = matrix_data.get("total_cartera_usd", 0.0)
    total_clientes = matrix_data.get("total_clientes", 0)
    lbl_fm = matrix_data.get("label_fin_mes", "Fin de Mes")
    lbl_da = matrix_data.get("label_dia_anterior", "Día Anterior")
    is_3 = (str(mode or "").lower() in ("3", "3tramos", "tramos3"))

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(letter),
        leftMargin=16,
        rightMargin=16,
        topMargin=18,
        bottomMargin=18
    )
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        'MatTitle',
        parent=styles['Heading1'],
        fontSize=13,
        leading=16,
        textColor=colors.HexColor("#0f172a"),
        fontName="Helvetica-Bold",
        spaceAfter=2
    )
    subtitle_style = ParagraphStyle(
        'MatSubtitle',
        parent=styles['Normal'],
        fontSize=8,
        leading=10,
        textColor=colors.HexColor("#64748b"),
        spaceAfter=8
    )
    hdr_group = ParagraphStyle(
        'HdrGroup',
        parent=styles['Normal'],
        fontSize=6.8,
        leading=8,
        textColor=colors.white,
        fontName="Helvetica-Bold",
        alignment=1
    )
    hdr_sub = ParagraphStyle(
        'HdrSub',
        parent=styles['Normal'],
        fontSize=6.3,
        leading=7.5,
        textColor=colors.HexColor("#334155"),
        fontName="Helvetica-Bold",
        alignment=2
    )
    cell_camp = ParagraphStyle(
        'CellCamp',
        parent=styles['Normal'],
        fontSize=6.8,
        leading=8,
        textColor=colors.HexColor("#0f172a"),
        fontName="Helvetica-Bold"
    )
    cell_val = ParagraphStyle(
        'CellVal',
        parent=styles['Normal'],
        fontSize=6.5,
        leading=8,
        textColor=colors.HexColor("#1e293b"),
        alignment=2
    )
    cell_tot = ParagraphStyle(
        'CellTot',
        parent=styles['Normal'],
        fontSize=6.5,
        leading=8,
        textColor=colors.HexColor("#0f172a"),
        fontName="Helvetica-Bold",
        alignment=2
    )

    title_mode = "3 TRAMOS CONTABLES" if is_3 else "5 TRAMOS OPERATIVOS"
    story = [
        Paragraph(f"COBRANZAPP | RESUMEN CARTERA EN MORA POR CAMPAÑA ({title_mode})", title_style),
        Paragraph(
            f"<b>Fecha de corte:</b> al {cutoff} &nbsp;|&nbsp; "
            f"<b>Mora Total:</b> ${total_usd:,.2f} USD &nbsp;|&nbsp; "
            f"<b>Clientes en Mora:</b> {total_clientes:,} &nbsp;|&nbsp; "
            f"<b>Emisión:</b> {datetime.now().strftime('%d/%m/%Y %H:%M')}",
            subtitle_style
        )
    ]

    if is_3:
        tramo_defs = [
            ("31+", "Vencido contable (+30d)", "#991b1b"),
            ("16-30", "Vencido Gestión (16-30d)", "#b45309"),
            ("1-15", "Vencido Irregular (1-15d)", "#1e40af"),
        ]
    else:
        tramo_defs = [
            ("1-30", "1 a 30 días", "#065f46"),
            ("31-60", "31 a 60 días", "#1e40af"),
            ("61-90", "61 a 90 días", "#92400e"),
            ("91-180", "91 a 180 días", "#9a3412"),
            ("181+", "181+ días", "#991b1b"),
        ]

    hdr_row_1 = [Paragraph("<b>Campaña</b>", hdr_group)]
    hdr_row_2 = [""]
    for _, t_label, _ in tramo_defs:
        hdr_row_1.extend([Paragraph(f"<b>{t_label}</b>", hdr_group), ""])
        hdr_row_2.extend([Paragraph("Monto ($)", hdr_sub), Paragraph("Ctes", hdr_sub)])

    hdr_row_1.extend([Paragraph("<b>Total Mora</b>", hdr_group), "", ""])
    hdr_row_2.extend([Paragraph("Monto ($)", hdr_sub), Paragraph("Ctes", hdr_sub), Paragraph("% Mora", hdr_sub)])

    hdr_row_1.extend([Paragraph("<b>% vs. Cartera Activa</b>", hdr_group), "", ""])
    hdr_row_2.extend([
        Paragraph(f"Al {html.escape(lbl_fm)}", hdr_sub),
        Paragraph(f"Al {html.escape(lbl_da)}", hdr_sub),
        Paragraph("Variación", hdr_sub),
    ])

    table_data = [hdr_row_1, hdr_row_2]
    row_dict_key = "tramos_3" if is_3 else "tramos"

    for row in matrix_data.get("matrix", []):
        row_cells = [Paragraph(html.escape(row["campana"]), cell_camp)]
        r_tramos = row.get(row_dict_key, {})
        for t_key, _, _ in tramo_defs:
            t_data = r_tramos.get(t_key, {"monto": 0.0, "clientes": 0})
            m_str = f"${t_data['monto']:,.2f}" if t_data["monto"] > 0 else "-"
            c_str = f"{t_data['clientes']:,}" if t_data["clientes"] > 0 else "-"
            row_cells.append(Paragraph(m_str, cell_val))
            row_cells.append(Paragraph(c_str, cell_val))
        row_cells.append(Paragraph(f"${row['total_monto']:,.2f}", cell_tot))
        row_cells.append(Paragraph(f"{row['total_clientes']:,}", cell_tot))
        row_cells.append(Paragraph(f"{row['porcentaje']:.2f}%", cell_tot))

        pfm = row.get("pct_fin_mes")
        pda = row.get("pct_dia_anterior")
        pvar = row.get("variacion_pct")
        row_cells.append(Paragraph(f"{pfm:.2f}%" if pfm is not None else "-", cell_tot))
        row_cells.append(Paragraph(f"{pda:.2f}%" if pda is not None else "-", cell_tot))
        row_cells.append(Paragraph(f"{pvar:+.2f}%" if pvar is not None else "-", cell_tot))
        table_data.append(row_cells)

    totals_tramos = matrix_data.get("totals_by_tramo_3" if is_3 else "totals_by_tramo", {})
    total_row_cells = [Paragraph("<b>TOTAL</b>", cell_camp)]
    for t_key, _, _ in tramo_defs:
        t_data = totals_tramos.get(t_key, {"monto": 0.0, "clientes": 0})
        total_row_cells.append(Paragraph(f"<b>${t_data['monto']:,.2f}</b>", cell_tot))
        total_row_cells.append(Paragraph(f"<b>{t_data['clientes']:,}</b>", cell_tot))
    total_row_cells.append(Paragraph(f"<b>${total_usd:,.2f}</b>", cell_tot))
    total_row_cells.append(Paragraph(f"<b>{total_clientes:,}</b>", cell_tot))
    total_row_cells.append(Paragraph("<b>100.0%</b>", cell_tot))

    t_pfm = matrix_data.get("total_pct_fin_mes")
    t_pda = matrix_data.get("total_pct_dia_anterior")
    t_pvar = matrix_data.get("total_variacion_pct")
    total_row_cells.append(Paragraph(f"<b>{t_pfm:.2f}%</b>" if t_pfm is not None else "-", cell_tot))
    total_row_cells.append(Paragraph(f"<b>{t_pda:.2f}%</b>" if t_pda is not None else "-", cell_tot))
    total_row_cells.append(Paragraph(f"<b>{t_pvar:+.2f}%</b>" if t_pvar is not None else "-", cell_tot))
    table_data.append(total_row_cells)

    if is_3:
        col_widths = [125, 58, 34, 58, 34, 58, 34, 62, 36, 38, 48, 48, 45]
    else:
        col_widths = [105, 42, 24, 42, 24, 42, 24, 44, 24, 42, 24, 48, 26, 32, 40, 40, 38]

    t_mat = Table(table_data, colWidths=col_widths, repeatRows=2)
    style_cmds = [
        ('SPAN', (0, 0), (0, 1)),
        ('BACKGROUND', (0, 0), (0, 1), colors.HexColor("#0f172a")),
    ]
    c_cursor = 1
    for _, _, hex_bg in tramo_defs:
        style_cmds.append(('SPAN', (c_cursor, 0), (c_cursor + 1, 0)))
        style_cmds.append(('BACKGROUND', (c_cursor, 0), (c_cursor + 1, 0), colors.HexColor(hex_bg)))
        c_cursor += 2

    style_cmds.append(('SPAN', (c_cursor, 0), (c_cursor + 2, 0)))
    style_cmds.append(('BACKGROUND', (c_cursor, 0), (c_cursor + 2, 0), colors.HexColor("#0f172a")))
    c_cursor += 3

    style_cmds.append(('SPAN', (c_cursor, 0), (c_cursor + 2, 0)))
    style_cmds.append(('BACKGROUND', (c_cursor, 0), (c_cursor + 2, 0), colors.HexColor("#c2410c")))
    style_cmds.append(('BACKGROUND', (c_cursor, 2), (c_cursor + 2, -2), colors.HexColor("#fef3c7")))

    style_cmds.extend([
        ('BACKGROUND', (1, 1), (-1, 1), colors.HexColor("#f1f5f9")),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#cbd5e1")),
        ('ROWBACKGROUNDS', (0, 2), (c_cursor - 1, -2), [colors.white, colors.HexColor("#f8fafc")]),
        ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor("#e2e8f0")),
        ('TOPPADDING', (0, 0), (-1, -1), 2.5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2.5),
        ('LEFTPADDING', (0, 0), (-1, -1), 2.5),
        ('RIGHTPADDING', (0, 0), (-1, -1), 2.5),
    ])

    t_mat.setStyle(TableStyle(style_cmds))
    story.append(t_mat)

    doc.build(story)
    buf.seek(0)
    suffix = "3_tramos" if is_3 else "5_tramos"
    filename = f"matriz_campanas_{suffix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
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
