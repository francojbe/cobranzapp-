import io
import os
from datetime import datetime
from typing import Optional
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

from core.config import BASE_DIR, DATA_STORE, OFFICIAL_TRAMOS_CONFIG, OFFICIAL_3_TRAMOS_CONFIG
from core.analytics import (
    calculate_totals,
    calculate_global_tramos_summary,
    calculate_3_tramos_summary,
    calculate_fibex_summary,
    calculate_campanas_tramos_matrix,
    resolve_active_debt_df,
    get_cutoff_date,
    detect_cartera_module
)

# -------------------------------------------------------------
# PALETA DE COLORES OFICIAL CREDIX & COBRANZAPP (16:9 PRESENTATION)
# -------------------------------------------------------------
C_MAGENTA = RGBColor(255, 0, 110)      # #ff006e - Magenta Credix
C_PURPLE = RGBColor(46, 0, 100)        # #2e0064 - Púrpura Corporativo Credix
C_TEAL = RGBColor(20, 184, 166)        # #14b8a6 - Acento Turquesa
C_DARK = RGBColor(15, 23, 42)          # #0f172a - Texto principal / Headers oscuros
C_NAVY = RGBColor(30, 58, 95)          # #1e3a5f - Azul marino tabla contable
C_GRAY = RGBColor(100, 116, 139)       # #64748b - Textos secundarios
C_LIGHT_GRAY = RGBColor(241, 245, 249) # #f1f5f9 - Filas alternas y cajas suaves
C_LIGHT_BG = RGBColor(248, 250, 252)   # #f8fafc - Fondo de tarjetas KPI
C_BORDER = RGBColor(226, 232, 240)     # #e2e8f0 - Bordes sutiles
C_WHITE = RGBColor(255, 255, 255)
C_EMERALD = RGBColor(16, 185, 129)     # #10b981 - Tramo preventiva
C_AMBER = RGBColor(245, 158, 11)       # #f59e0b - Tramo administrativo / irregular
C_ROSE = RGBColor(225, 29, 72)         # #e11d48 - Tramo vencido contable / alerta
C_YELLOW = RGBColor(254, 240, 138)     # #fef08a - Amarillo celda contable

FONT_HEADING = "Plus Jakarta Sans"
FONT_BODY = "Plus Jakarta Sans"

LOGO_PNG_PATH = os.path.join(BASE_DIR, "static", "img", "credix", "logo.png")
if not os.path.exists(LOGO_PNG_PATH):
    # Generar versión PNG si no existe a partir del webp
    try:
        from PIL import Image
        webp_path = os.path.join(BASE_DIR, "static", "img", "credix", "logo.webp")
        if os.path.exists(webp_path):
            im = Image.open(webp_path)
            im.save(LOGO_PNG_PATH, "PNG")
    except Exception:
        pass


def _add_running_header(slide, tag_text: str, title_text: str, subtitle_text: str):
    """Añade la cabecera ejecutiva con logo de Credix, tag superior y títulos."""
    # Logo Credix superior derecho
    if os.path.exists(LOGO_PNG_PATH):
        slide.shapes.add_picture(LOGO_PNG_PATH, Inches(11.2), Inches(0.4), height=Inches(0.48))
    else:
        tb_logo = slide.shapes.add_textbox(Inches(11.0), Inches(0.4), Inches(1.8), Inches(0.45))
        p_logo = tb_logo.text_frame.paragraphs[0]
        p_logo.text = "credix"
        p_logo.font.bold = True
        p_logo.font.size = Pt(22)
        p_logo.font.color.rgb = C_MAGENTA

    # Tag de sección con viñeta magenta
    tb_tag = slide.shapes.add_textbox(Inches(0.8), Inches(0.45), Inches(8.5), Inches(0.35))
    tf_tag = tb_tag.text_frame
    tf_tag.word_wrap = True
    p_tag = tf_tag.paragraphs[0]
    r_dot = p_tag.add_run()
    r_dot.text = "● "
    r_dot.font.color.rgb = C_MAGENTA
    r_dot.font.size = Pt(9)
    r_tag = p_tag.add_run()
    r_tag.text = tag_text.upper()
    r_tag.font.bold = True
    r_tag.font.size = Pt(9.5)
    r_tag.font.color.rgb = C_MAGENTA
    r_tag.font.name = FONT_HEADING

    # Título principal
    tb_title = slide.shapes.add_textbox(Inches(0.8), Inches(0.75), Inches(10.0), Inches(0.55))
    tf_title = tb_title.text_frame
    tf_title.word_wrap = True
    p_title = tf_title.paragraphs[0]
    p_title.text = title_text
    p_title.font.bold = True
    p_title.font.size = Pt(21)
    p_title.font.color.rgb = C_DARK
    p_title.font.name = FONT_HEADING

    # Subtítulo descriptivo
    tb_sub = slide.shapes.add_textbox(Inches(0.8), Inches(1.25), Inches(10.5), Inches(0.35))
    tf_sub = tb_sub.text_frame
    tf_sub.word_wrap = True
    p_sub = tf_sub.paragraphs[0]
    p_sub.text = subtitle_text
    p_sub.font.size = Pt(10.5)
    p_sub.font.color.rgb = C_GRAY
    p_sub.font.name = FONT_BODY


def _add_running_footer(slide, cutoff_date: str, page_num: int):
    """Añade el pie de página discreto institucional."""
    # Texto izquierda
    tb_l = slide.shapes.add_textbox(Inches(0.8), Inches(7.05), Inches(6.0), Inches(0.3))
    p_l = tb_l.text_frame.paragraphs[0]
    p_l.text = f"Credix | Informe de Cartera al {cutoff_date}"
    p_l.font.size = Pt(8.5)
    p_l.font.color.rgb = C_GRAY
    p_l.font.name = FONT_BODY

    # Número de página derecha
    tb_r = slide.shapes.add_textbox(Inches(11.2), Inches(7.05), Inches(1.3), Inches(0.3))
    p_r = tb_r.text_frame.paragraphs[0]
    p_r.text = f"Página {page_num}"
    p_r.font.size = Pt(8.5)
    p_r.alignment = PP_ALIGN.RIGHT
    p_r.font.color.rgb = C_GRAY
    p_r.font.name = FONT_BODY


def _add_kpi_card(slide, left, top, width, height, title, main_val, sub_val, highlight_sub=False):
    """Crea una tarjeta métrica KPI con borde suave y fondo limpio."""
    # Fondo de la tarjeta
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = C_LIGHT_BG
    shape.line.color.rgb = C_BORDER
    shape.line.width = Pt(1)

    tf = shape.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE

    # 1. Título pequeño
    p1 = tf.paragraphs[0]
    p1.text = title.upper()
    p1.font.size = Pt(8.5)
    p1.font.bold = True
    p1.font.color.rgb = C_GRAY
    p1.font.name = FONT_HEADING
    p1.space_after = Pt(4)

    # 2. Valor principal
    p2 = tf.add_paragraph()
    p2.text = str(main_val)
    p2.font.size = Pt(22)
    p2.font.bold = True
    p2.font.color.rgb = C_DARK if not highlight_sub else C_MAGENTA
    p2.font.name = FONT_HEADING
    p2.space_after = Pt(2)

    # 3. Subtítulo / Ratio
    p3 = tf.add_paragraph()
    p3.text = str(sub_val)
    p3.font.size = Pt(8.5)
    p3.font.bold = highlight_sub
    p3.font.color.rgb = C_MAGENTA if highlight_sub else C_GRAY
    p3.font.name = FONT_BODY


def _add_callout_box(slide, left, top, width, height, title: str, bullets: list):
    """Crea una caja de diagnóstico o recomendaciones con barra vertical izquierda en magenta."""
    # Caja contenedor blanca
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    box.fill.solid()
    box.fill.fore_color.rgb = C_WHITE
    box.line.color.rgb = C_BORDER
    box.line.width = Pt(1)

    # Barra vertical magenta de acento
    accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, Inches(0.08), height)
    accent.fill.solid()
    accent.fill.fore_color.rgb = C_MAGENTA
    accent.line.fill.background()

    # Contenido textual
    tb = slide.shapes.add_textbox(left + Inches(0.25), top + Inches(0.12), width - Inches(0.4), height - Inches(0.24))
    tf = tb.text_frame
    tf.word_wrap = True

    p_title = tf.paragraphs[0]
    p_title.text = title
    p_title.font.size = Pt(12)
    p_title.font.bold = True
    p_title.font.color.rgb = C_DARK
    p_title.font.name = FONT_HEADING
    p_title.space_after = Pt(8)

    for b in bullets:
        p_b = tf.add_paragraph()
        p_b.space_after = Pt(5)
        p_b.font.size = Pt(9.5)
        p_b.font.name = FONT_BODY

        # Si el bullet contiene ':' destacamos el prefijo en negrita
        if ":" in b:
            prefix, body = b.split(":", 1)
            r_pre = p_b.add_run()
            r_pre.text = f"• {prefix.strip()}: "
            r_pre.font.bold = True
            r_pre.font.color.rgb = C_DARK
            r_body = p_b.add_run()
            r_body.text = body.strip()
            r_body.font.color.rgb = C_DARK
        else:
            r = p_b.add_run()
            r.text = f"• {b.strip()}"
            r.font.color.rgb = C_DARK


def _style_table_header(table, cols, col_widths, bg_color=C_DARK):
    """Establece anchos y diseño del encabezado de tabla."""
    for idx, width in enumerate(col_widths):
        table.columns[idx].width = width

    for col_idx, col_name in enumerate(cols):
        cell = table.cell(0, col_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = bg_color
        cell.text_frame.word_wrap = True
        p = cell.text_frame.paragraphs[0]
        p.text = col_name
        p.font.size = Pt(9)
        p.font.bold = True
        p.font.color.rgb = C_WHITE
        p.font.name = FONT_HEADING
        if col_idx >= 2:
            p.alignment = PP_ALIGN.RIGHT


def _format_table_data_row(table, row_idx, data_vals, is_total=False, bg_color=None):
    """Llena y formatea una fila de datos en la tabla."""
    for col_idx, val in enumerate(data_vals):
        cell = table.cell(row_idx, col_idx)
        cell.fill.solid()
        if bg_color:
            cell.fill.fore_color.rgb = bg_color
        elif is_total:
            cell.fill.fore_color.rgb = C_LIGHT_GRAY
        elif row_idx % 2 == 0:
            cell.fill.fore_color.rgb = C_WHITE
        else:
            cell.fill.fore_color.rgb = C_LIGHT_GRAY

        cell.text_frame.word_wrap = True
        p = cell.text_frame.paragraphs[0]
        p.text = str(val)
        p.font.size = Pt(9)
        p.font.name = FONT_BODY
        if is_total:
            p.font.bold = True
            p.font.color.rgb = C_DARK
        else:
            p.font.color.rgb = C_DARK

        # Alineación: texto a la izquierda, números y montos a la derecha
        if col_idx >= 2 or (col_idx == 1 and str(val).isdigit()):
            p.alignment = PP_ALIGN.RIGHT
        else:
            p.alignment = PP_ALIGN.LEFT


# -------------------------------------------------------------
# CONSTRUCTOR PRINCIPAL DE LA PRESENTACIÓN POWERPOINT (.PPTX)
# -------------------------------------------------------------

def generate_executive_presentation_pptx() -> bytes:
    """Genera la presentación ejecutiva completa en PowerPoint (.pptx) con diseño 16:9 widescreen institucional."""
    active_df = DATA_STORE.get("df_original")
    if active_df is None or active_df.empty:
        active_df = DATA_STORE.get("df_cartera_raw")

    if active_df is None or active_df.empty:
        raise ValueError("No hay datos de cartera cargados para generar la presentación ejecutiva.")

    df_active = resolve_active_debt_df(active_df)
    filename = DATA_STORE.get("cartera_filename", "mora.csv")
    mod_id = detect_cartera_module(df_active, filename)
    mod_label = "Mora Ávila" if mod_id == "avila" else "Mora Ubi"
    cutoff_date = get_cutoff_date(df_active)

    # 1. Cálculos de métricas clave del sistema
    kpis = calculate_totals(active_df)
    g_tramos = calculate_global_tramos_summary(active_df)
    t3 = calculate_3_tramos_summary(active_df)
    fib = calculate_fibex_summary(active_df)
    camp_mat = calculate_campanas_tramos_matrix(active_df)

    cartera_activa_usd = t3.get("cartera_activa_usd") or g_tramos.get("cartera_activa_usd") or kpis.get("monto_total", 0.0)
    pct_mora_macro = round((kpis["monto_total"] / cartera_activa_usd * 100), 2) if cartera_activa_usd > 0 else 0.0

    # Inicializar PowerPoint con aspect ratio 16:9
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # =========================================================
    # DIAPOSITIVA 1: PORTADA EJECUTIVA
    # =========================================================
    slide_cover = prs.slides.add_slide(blank_layout)

    # Barra superior con degradado conceptual (Magenta y Turquesa)
    bar_top = slide_cover.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.12))
    bar_top.fill.solid()
    bar_top.fill.fore_color.rgb = C_MAGENTA
    bar_top.line.fill.background()

    # Logo Credix
    if os.path.exists(LOGO_PNG_PATH):
        slide_cover.shapes.add_picture(LOGO_PNG_PATH, Inches(10.8), Inches(0.8), height=Inches(0.65))

    # Tag institucional
    tb_cov_tag = slide_cover.shapes.add_textbox(Inches(1.0), Inches(2.3), Inches(8.0), Inches(0.4))
    p_c_tag = tb_cov_tag.text_frame.paragraphs[0]
    r_c_dot = p_c_tag.add_run()
    r_c_dot.text = "● "
    r_c_dot.font.color.rgb = C_MAGENTA
    r_c_tag = p_c_tag.add_run()
    r_c_tag.text = "CREDIX | COBRANZAPP"
    r_c_tag.font.bold = True
    r_c_tag.font.size = Pt(11)
    r_c_tag.font.color.rgb = C_MAGENTA
    r_c_tag.font.name = FONT_HEADING

    # Título de Portada
    tb_cov_title = slide_cover.shapes.add_textbox(Inches(1.0), Inches(2.75), Inches(11.0), Inches(1.2))
    tf_c_title = tb_cov_title.text_frame
    tf_c_title.word_wrap = True
    p_c_title = tf_c_title.paragraphs[0]
    p_c_title.text = f"Informe Ejecutivo de Análisis de Cartera {mod_label.replace('Mora ', '')}"
    p_c_title.font.bold = True
    p_c_title.font.size = Pt(32)
    p_c_title.font.color.rgb = C_DARK
    p_c_title.font.name = FONT_HEADING

    # Subtítulo de Portada
    tb_cov_sub = slide_cover.shapes.add_textbox(Inches(1.0), Inches(4.05), Inches(11.0), Inches(0.6))
    tf_c_sub = tb_cov_sub.text_frame
    tf_c_sub.word_wrap = True
    p_c_sub = tf_c_sub.paragraphs[0]
    p_c_sub.text = f"Cierre al {cutoff_date} — Estado de la Morosidad, Segmentación Fibex y Análisis de Tramos Contables"
    p_c_sub.font.size = Pt(13)
    p_c_sub.font.color.rgb = C_GRAY
    p_c_sub.font.name = FONT_BODY

    # Línea divisoria
    sep_line = slide_cover.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.0), Inches(5.1), Inches(11.333), Inches(0.02))
    sep_line.fill.solid()
    sep_line.fill.fore_color.rgb = C_BORDER
    sep_line.line.fill.background()

    # Metadata inferior (3 columnas)
    meta_cols = [
        ("FECHA DE CORTE", cutoff_date),
        ("EMISIÓN DE REPORTE", f"{datetime.now().strftime('%d/%m/%Y')} — {datetime.now().strftime('%H:%M')} hrs"),
        ("UNIDAD RESPONSABLE", "Coordinación de Cobranzas & Cartera")
    ]
    for idx, (m_lbl, m_val) in enumerate(meta_cols):
        left_pos = Inches(1.0 + idx * 3.8)
        tb_m = slide_cover.shapes.add_textbox(left_pos, Inches(5.35), Inches(3.5), Inches(0.9))
        tf_m = tb_m.text_frame
        p_m_lbl = tf_m.paragraphs[0]
        p_m_lbl.text = m_lbl
        p_m_lbl.font.size = Pt(9)
        p_m_lbl.font.bold = True
        p_m_lbl.font.color.rgb = C_GRAY
        p_m_lbl.font.name = FONT_HEADING
        p_m_val = tf_m.add_paragraph()
        p_m_val.text = m_val
        p_m_val.font.size = Pt(11)
        p_m_val.font.bold = True
        p_m_val.font.color.rgb = C_DARK
        p_m_val.font.name = FONT_BODY

    # =========================================================
    # DIAPOSITIVA 2: PANORAMA GLOBAL DE CARTERA EN MORA
    # =========================================================
    slide_global = prs.slides.add_slide(blank_layout)
    _add_running_header(
        slide_global,
        "RESUMEN EJECUTIVO",
        "Panorama Global de Cartera en Mora",
        f"Magnitudes financieras principales y volumen general evaluado al {cutoff_date}"
    )

    # 4 Tarjetas KPI superiores
    vencido_contable_monto = 0.0
    for t_item in t3.get("tramos", []):
        if t_item.get("id") == "31+":
            vencido_contable_monto = t_item.get("total_monto", 0.0)

    pct_vc_activa = round((vencido_contable_monto / cartera_activa_usd * 100), 2) if cartera_activa_usd > 0 else 0.0

    kpi_card_w = Inches(2.76)
    kpi_card_h = Inches(1.35)
    kpi_card_top = Inches(1.85)

    _add_kpi_card(
        slide_global, Inches(0.8), kpi_card_top, kpi_card_w, kpi_card_h,
        "CARTERA ACTIVA TOTAL", f"${cartera_activa_usd:,.0f}", "Base de cartera evaluada"
    )
    _add_kpi_card(
        slide_global, Inches(3.72), kpi_card_top, kpi_card_w, kpi_card_h,
        "TARIFA EN MORA TOTAL", f"${kpis['monto_total']:,.2f}", f"{pct_mora_macro:.2f}% % Mora s/ Cartera Activa",
        highlight_sub=True
    )
    _add_kpi_card(
        slide_global, Inches(6.64), kpi_card_top, kpi_card_w, kpi_card_h,
        "CLIENTES EN MORA", f"{kpis['total_clientes']:,}", f"{kpis['total_cuotas']:,} cuotas pendientes"
    )
    _add_kpi_card(
        slide_global, Inches(9.56), kpi_card_top, kpi_card_w, kpi_card_h,
        "VENCIDO CONTABLE (>30D)", f"${vencido_contable_monto:,.2f}", f"{pct_vc_activa:.2f}% de Cartera Activa"
    )

    # Diagnóstico Financiero Consolidado
    mora_temprana_monto = 0.0
    for t_item in t3.get("tramos", []):
        if t_item.get("id") in ("1-15", "16-30"):
            mora_temprana_monto += t_item.get("total_monto", 0.0)

    pct_temprana_mora = round((mora_temprana_monto / kpis['monto_total'] * 100), 2) if kpis['monto_total'] > 0 else 0.0

    diag_bullets = [
        f"Exposición Financiera: La morosidad total asciende a ${kpis['monto_total']:,.2f} USD distribuida en {kpis['total_clientes']:,} clientes únicos y {kpis['total_cuotas']:,} cuotas pendientes de pago.",
        f"Índice Macro de Morosidad (% Mora): El ratio general sobre la Cartera Activa Total (${cartera_activa_usd:,.2f} USD) se ubica en {pct_mora_macro:.2f}%.",
        f"Concentración del Riesgo: El {pct_temprana_mora:.2f}% de la mora (${mora_temprana_monto:,.2f} USD) se concentra dentro de los primeros 30 días de cobranza temprana, reportando ${vencido_contable_monto:,.2f} USD en mora contable superior a 30 días."
    ]
    _add_callout_box(slide_global, Inches(0.8), Inches(3.5), Inches(11.52), Inches(3.2), "Diagnóstico Financiero Consolidado", diag_bullets)
    _add_running_footer(slide_global, cutoff_date, 2)

    # =========================================================
    # DIAPOSITIVA 3: DISTRIBUCIÓN POR 5 TRAMOS OFICIALES DE MORA
    # =========================================================
    slide_5tramos = prs.slides.add_slide(blank_layout)
    _add_running_header(
        slide_5tramos,
        "AGING DE CARTERA",
        "Distribución por 5 Tramos Oficiales de Mora",
        "Envejecimiento de la tarifa pendiente y porcentaje de afectación a la cartera activa"
    )

    # Tabla 5 Tramos
    cols_5t = ["Fase / Categoría", "Rango de Mora", "Cant. Cuotas", "Clientes", "Monto Tarifa (USD)", "% Cartera Activa"]
    col_w_5t = [Inches(2.7), Inches(1.8), Inches(1.5), Inches(1.5), Inches(2.1), Inches(1.92)]
    tramos_list = g_tramos.get("tramos", [])
    num_rows_5t = len(tramos_list) + 2  # +1 header, +1 total

    t_shape_5t = slide_5tramos.shapes.add_table(num_rows_5t, len(cols_5t), Inches(0.8), Inches(1.85), Inches(11.52), Inches(3.3))
    table_5t = t_shape_5t.table
    _style_table_header(table_5t, cols_5t, col_w_5t)

    max_tramo_fase = "Cobranza temprana"
    max_tramo_monto = 0.0
    max_tramo_pct = 0.0

    for idx, t_info in enumerate(tramos_list):
        pct_str = f"{t_info.get('porcentaje'):.2f}%" if t_info.get("porcentaje") is not None else "-"
        monto_val = t_info.get("total_monto", 0.0)
        if monto_val > max_tramo_monto:
            max_tramo_monto = monto_val
            max_tramo_fase = t_info.get("fase", "")
            max_tramo_pct = t_info.get("porcentaje", 0.0) or 0.0

        row_vals = [
            t_info.get("fase", ""),
            t_info.get("etiqueta", ""),
            f"{t_info.get('cuotas', 0):,}",
            f"{t_info.get('clientes', 0):,}",
            f"${monto_val:,.2f}",
            pct_str
        ]
        _format_table_data_row(table_5t, idx + 1, row_vals)

    # Fila Total
    tot_row_vals = [
        "TOTAL CARTERA", "Gestión Total", f"{g_tramos.get('total_cuotas', 0):,}",
        f"{g_tramos.get('total_clientes', 0):,}", f"${g_tramos.get('total_cartera_usd', 0.0):,.2f}",
        f"{pct_mora_macro:.2f}%"
    ]
    _format_table_data_row(table_5t, num_rows_5t - 1, tot_row_vals, is_total=True)

    # Bloque de Comportamiento del Aging
    pct_max_sobre_mora = round((max_tramo_monto / kpis['monto_total'] * 100), 2) if kpis['monto_total'] > 0 else 0.0
    aging_bullets = [
        f"Punto Crítico de Antigüedad: El tramo '{max_tramo_fase}' concentra el mayor volumen con ${max_tramo_monto:,.2f} USD ({pct_max_sobre_mora:.2f}% de la mora total y {max_tramo_pct:.2f}% de la cartera activa). Se requiere intensificar el cobro oportuno para evitar el paso a tramos contables superiores."
    ]
    _add_callout_box(slide_5tramos, Inches(0.8), Inches(5.35), Inches(11.52), Inches(1.4), "Comportamiento del Aging de Cartera", aging_bullets)
    _add_running_footer(slide_5tramos, cutoff_date, 3)

    # =========================================================
    # DIAPOSITIVA 4: RESUMEN POR 3 TRAMOS CONTABLES DE GESTIÓN
    # =========================================================
    slide_3tramos = prs.slides.add_slide(blank_layout)
    _add_running_header(
        slide_3tramos,
        "ESTRUCTURA CONTABLE",
        "Resumen por 3 Tramos Contables de Gestión",
        "Clasificación metodológica directiva de la cartera en mora"
    )

    cols_3t = ["Categoría Contable", "Tramo / Días de Mora", "Cant. Cuotas", "Clientes", "Monto (USD)", "% Cartera Activa"]
    col_w_3t = [Inches(2.7), Inches(2.3), Inches(1.4), Inches(1.4), Inches(1.9), Inches(1.82)]
    tramos_3_list = t3.get("tramos", [])
    num_rows_3t = len(tramos_3_list) + 2

    t_shape_3t = slide_3tramos.shapes.add_table(num_rows_3t, len(cols_3t), Inches(0.8), Inches(1.85), Inches(11.52), Inches(2.7))
    table_3t = t_shape_3t.table
    _style_table_header(table_3t, cols_3t, col_w_3t)

    for idx, t_info in enumerate(tramos_3_list):
        pct_str = f"{t_info.get('porcentaje'):.2f}%" if t_info.get("porcentaje") is not None else "-"
        row_vals = [
            t_info.get("fase", ""),
            t_info.get("etiqueta", ""),
            f"{t_info.get('cuotas', 0):,}",
            f"{t_info.get('clientes', 0):,}",
            f"${t_info.get('total_monto', 0.0):,.2f}",
            pct_str
        ]
        _format_table_data_row(table_3t, idx + 1, row_vals)

    tot_3t_vals = [
        "TOTAL GESTIÓN", "Gestión Total", f"{t3.get('total_cuotas', 0):,}",
        f"{t3.get('total_clientes', 0):,}", f"${t3.get('total_cartera_usd', 0.0):,.2f}",
        f"{pct_mora_macro:.2f}%"
    ]
    _format_table_data_row(table_3t, num_rows_3t - 1, tot_3t_vals, is_total=True)

    obs_contable_bullets = [
        f"Observación Contable Directiva: El tramo Vencido Contable (>30 días) concentra ${vencido_contable_monto:,.2f} USD ({pct_vc_activa:.2f}% de la Cartera Activa Total). La mora temprana (Irregular + Vencido Gestión) representa ${mora_temprana_monto:,.2f} USD ({pct_temprana_mora:.2f}% del saldo adeudado), susceptible a contención inmediata mediante recordatorios automáticos y llamadas focalizadas."
    ]
    _add_callout_box(slide_3tramos, Inches(0.8), Inches(4.8), Inches(11.52), Inches(1.9), "Observación Contable Directiva", obs_contable_bullets)
    _add_running_footer(slide_3tramos, cutoff_date, 4)

    # =========================================================
    # DIAPOSITIVA 5: RESUMEN GRÁFICO OFICIAL 3 TRAMOS (COLOR MATRIX)
    # =========================================================
    slide_matrix3 = prs.slides.add_slide(blank_layout)
    _add_running_header(
        slide_matrix3,
        "ESTRUCTURA CONTABLE",
        "Resumen por 3 Tramos Contables de Gestión",
        "Clasificación metodológica directiva de la cartera en mora"
    )

    # Barra superior magenta y amarillo para Cartera Activa
    bar_m_left = Inches(3.6)
    bar_m_top = Inches(2.1)

    s_ca_lbl = slide_matrix3.shapes.add_shape(MSO_SHAPE.RECTANGLE, bar_m_left, bar_m_top, Inches(4.2), Inches(0.48))
    s_ca_lbl.fill.solid()
    s_ca_lbl.fill.fore_color.rgb = C_MAGENTA
    s_ca_lbl.line.color.rgb = C_DARK
    s_ca_lbl.line.width = Pt(1.5)
    p_calbl = s_ca_lbl.text_frame.paragraphs[0]
    p_calbl.text = "Cartera activa (USD)"
    p_calbl.font.bold = True
    p_calbl.font.size = Pt(12)
    p_calbl.font.color.rgb = C_DARK
    p_calbl.alignment = PP_ALIGN.CENTER

    s_ca_val = slide_matrix3.shapes.add_shape(MSO_SHAPE.RECTANGLE, bar_m_left + Inches(4.2), bar_m_top, Inches(2.5), Inches(0.48))
    s_ca_val.fill.solid()
    s_ca_val.fill.fore_color.rgb = C_YELLOW
    s_ca_val.line.color.rgb = C_DARK
    s_ca_val.line.width = Pt(1.5)
    p_caval = s_ca_val.text_frame.paragraphs[0]
    p_caval.text = f"{cartera_activa_usd:,.0f}"
    p_caval.font.bold = True
    p_caval.font.size = Pt(13)
    p_caval.font.color.rgb = C_DARK
    p_caval.alignment = PP_ALIGN.CENTER

    # Cuadro visual contable azul marino / celeste
    t_shape_vis = slide_matrix3.shapes.add_table(5, 5, Inches(1.8), Inches(3.0), Inches(9.8), Inches(2.6))
    table_vis = t_shape_vis.table
    w_vis = [Inches(2.2), Inches(2.7), Inches(1.5), Inches(1.8), Inches(1.6)]
    for i, w in enumerate(w_vis):
        table_vis.columns[i].width = w

    # Cabecera de la tabla contable
    table_vis.cell(0, 0).text = ""
    table_vis.cell(0, 1).text = ""
    for c_i, c_lbl in [(2, "Cant"), (3, "Monto\n(USD)"), (4, "%")]:
        cell = table_vis.cell(0, c_i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_NAVY
        p = cell.text_frame.paragraphs[0]
        p.text = c_lbl
        p.font.bold = True
        p.font.size = Pt(11)
        p.font.color.rgb = C_WHITE
        p.alignment = PP_ALIGN.CENTER

    # Llenado de filas visuales
    row_vis_data = [
        ("Irregular", "Desde el día 1 hasta el 15", "1-15", C_YELLOW),
        ("Vencido Gestión", "Desde el día 16 al 30", "16-30", C_YELLOW),
        ("Vencido contable", "Mas de 30 dias", "31+", C_ROSE),
        ("Total", "Gestión total", "TOTAL", C_YELLOW)
    ]

    for r_idx, (cat_name, r_lbl, t_id, pct_color) in enumerate(row_vis_data, start=1):
        if t_id == "TOTAL":
            t_obj = {"clientes": t3.get("total_clientes", 0), "total_monto": t3.get("total_cartera_usd", 0.0), "porcentaje": pct_mora_macro}
        else:
            t_obj = next((x for x in tramos_3_list if x.get("id") == t_id), {})

        # Celda 0 y 1 (Celeste suave)
        c0 = table_vis.cell(r_idx, 0)
        c0.fill.solid()
        c0.fill.fore_color.rgb = RGBColor(219, 234, 254)
        p0 = c0.text_frame.paragraphs[0]
        p0.text = cat_name
        p0.font.bold = True
        p0.font.size = Pt(11)
        p0.font.color.rgb = C_NAVY
        p0.alignment = PP_ALIGN.CENTER

        c1 = table_vis.cell(r_idx, 1)
        c1.fill.solid()
        c1.fill.fore_color.rgb = RGBColor(219, 234, 254)
        p1 = c1.text_frame.paragraphs[0]
        p1.text = r_lbl
        p1.font.bold = True
        p1.font.size = Pt(11)
        p1.font.color.rgb = C_NAVY
        p1.alignment = PP_ALIGN.CENTER

        # Cant
        c2 = table_vis.cell(r_idx, 2)
        c2.fill.solid()
        c2.fill.fore_color.rgb = C_WHITE
        p2 = c2.text_frame.paragraphs[0]
        p2.text = f"{t_obj.get('clientes', 0):,}"
        p2.font.bold = True
        p2.font.size = Pt(11)
        p2.font.color.rgb = C_DARK
        p2.alignment = PP_ALIGN.CENTER

        # Monto
        c3 = table_vis.cell(r_idx, 3)
        c3.fill.solid()
        c3.fill.fore_color.rgb = C_WHITE
        p3 = c3.text_frame.paragraphs[0]
        p3.text = f"{t_obj.get('total_monto', 0.0):,.0f}"
        p3.font.bold = True
        p3.font.size = Pt(11)
        p3.font.color.rgb = C_DARK
        p3.alignment = PP_ALIGN.CENTER

        # % con color de fondo
        c4 = table_vis.cell(r_idx, 4)
        c4.fill.solid()
        c4.fill.fore_color.rgb = pct_color
        p4 = c4.text_frame.paragraphs[0]
        pct_val = t_obj.get('porcentaje')
        p4.text = f"{pct_val:.2f}%" if pct_val is not None else "-"
        p4.font.bold = True
        p4.font.size = Pt(11)
        p4.font.color.rgb = C_WHITE if pct_color == C_ROSE else C_DARK
        p4.alignment = PP_ALIGN.CENTER

    _add_running_footer(slide_matrix3, cutoff_date, 5)

    # =========================================================
    # DIAPOSITIVA 6: ANÁLISIS COMPARATIVO FIBEX VS. NO FIBEX
    # =========================================================
    slide_fibex = prs.slides.add_slide(blank_layout)
    _add_running_header(
        slide_fibex,
        "SEGMENTACIÓN DE CARTERA",
        "Análisis Comparativo: Clientes Fibex vs. NO Fibex",
        "Distribución del saldo adeudado según vinculación al servicio de telecomunicaciones"
    )

    segments = fib.get("segments", [])
    seg_nofibex = next((s for s in segments if s.get("id") == "no_fibex"), {})
    seg_fibex = next((s for s in segments if s.get("id") == "fibex_resto"), {})

    def _add_fibex_card(slide, left, top, width, height, title, seg_data, strategy_text):
        monto = seg_data.get("total_monto", 0.0)
        pct_mora = seg_data.get("porcentaje_mora", 0.0)
        pct_activa = seg_data.get("porcentaje")
        pct_activa_str = f"{pct_activa:.2f}% de la Cartera Activa Total" if pct_activa is not None else ""
        clientes = seg_data.get("clientes", 0)
        cuotas = seg_data.get("cuotas", 0)
        tramos_dict = seg_data.get("tramos", {})

        box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        box.fill.solid()
        box.fill.fore_color.rgb = C_WHITE
        box.line.color.rgb = C_BORDER
        box.line.width = Pt(1)

        tb = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(0.2), width - Inches(0.4), height - Inches(0.4))
        tf = tb.text_frame
        tf.word_wrap = True

        # Header de tarjeta
        p_hdr = tf.paragraphs[0]
        p_hdr.text = title.upper()
        p_hdr.font.bold = True
        p_hdr.font.size = Pt(9.5)
        p_hdr.font.color.rgb = C_MAGENTA if "NO FIBEX" in title else RGBColor(59, 130, 246)
        p_hdr.space_after = Pt(4)

        # Monto grande y porcentaje
        p_monto = tf.add_paragraph()
        r_m = p_monto.add_run()
        r_m.text = f"${monto:,.2f} "
        r_m.font.bold = True
        r_m.font.size = Pt(22)
        r_m.font.color.rgb = C_DARK
        r_pct = p_monto.add_run()
        r_pct.text = f"({pct_mora:.2f}%)"
        r_pct.font.bold = True
        r_pct.font.size = Pt(13)
        r_pct.font.color.rgb = C_MAGENTA if "NO FIBEX" in title else RGBColor(59, 130, 246)
        p_monto.space_after = Pt(2)

        p_sub = tf.add_paragraph()
        p_sub.text = pct_activa_str
        p_sub.font.size = Pt(9)
        p_sub.font.color.rgb = C_GRAY
        p_sub.space_after = Pt(10)

        # Volúmenes
        p_vol = tf.add_paragraph()
        p_vol.text = f"• Volumen: {clientes:,} clientes | {cuotas:,} cuotas"
        p_vol.font.size = Pt(9.5)
        p_vol.font.bold = True
        p_vol.font.color.rgb = C_DARK

        p_orig = tf.add_paragraph()
        p_orig.text = f"• Origen: {seg_data.get('subtitulo', '')}"
        p_orig.font.size = Pt(9)
        p_orig.font.color.rgb = C_GRAY
        p_orig.space_after = Pt(8)

        # Desglose por tramos
        p_des = tf.add_paragraph()
        p_des.text = "• Desglose por Tramos:"
        p_des.font.size = Pt(9.5)
        p_des.font.bold = True
        p_des.font.color.rgb = C_DARK

        p_t1 = tf.add_paragraph()
        m_1_30 = tramos_dict.get("1-30", {}).get("monto", 0.0)
        m_31_60 = tramos_dict.get("31-60", {}).get("monto", 0.0)
        p_t1.text = f"   - 1-30d: ${m_1_30:,.2f} | 31-60d: ${m_31_60:,.2f}"
        p_t1.font.size = Pt(9)
        p_t1.font.color.rgb = C_GRAY

        p_t2 = tf.add_paragraph()
        m_61_90 = tramos_dict.get("61-90", {}).get("monto", 0.0)
        m_91_180 = tramos_dict.get("91-180", {}).get("monto", 0.0)
        m_181 = tramos_dict.get("181+", {}).get("monto", 0.0)
        p_t2.text = f"   - 61-90d: ${m_61_90:,.2f} | 91-180d: ${m_91_180:,.2f} | 181+d: ${m_181:,.2f}"
        p_t2.font.size = Pt(9)
        p_t2.font.color.rgb = C_GRAY
        p_t2.space_after = Pt(10)

        # Estrategia recomendada
        p_est = tf.add_paragraph()
        r_e_lbl = p_est.add_run()
        r_e_lbl.text = "Estrategia: "
        r_e_lbl.font.bold = True
        r_e_lbl.font.size = Pt(9.5)
        r_e_lbl.font.color.rgb = C_DARK
        r_e_txt = p_est.add_run()
        r_e_txt.text = strategy_text
        r_e_txt.font.size = Pt(9.5)
        r_e_txt.font.color.rgb = C_DARK

    card_w_fib = Inches(5.6)
    card_h_fib = Inches(4.8)
    _add_fibex_card(
        slide_fibex, Inches(0.8), Inches(1.85), card_w_fib, card_h_fib,
        "Clientes Web NO Fibex", seg_nofibex,
        "Cobranza digital directa vía WhatsApp y llamadas masivas preventivas con acuerdos de pago inmediatos."
    )
    _add_fibex_card(
        slide_fibex, Inches(6.72), Inches(1.85), card_w_fib, card_h_fib,
        "Clientes Fibex y Resto", seg_fibex,
        "Apalancamiento estratégico en la continuidad o suspensión preventiva del servicio de telecomunicaciones."
    )
    _add_running_footer(slide_fibex, cutoff_date, 6)

    # =========================================================
    # DIAPOSITIVA 7: CONCENTRACIÓN POR CAMPAÑAS COMERCIALES
    # =========================================================
    slide_camp = prs.slides.add_slide(blank_layout)
    _add_running_header(
        slide_camp,
        "PROGRAMAS COMERCIALES",
        "Concentración por Campañas Comerciales",
        f"Desglose de la cartera adeudada por origen comercial (Cierre al {cutoff_date})"
    )

    cols_camp = ["Campaña / Segmento Comercial", "Vencido Contable (>30d)", "Vencido Gestión (16-30d)", "Irregular (1-15d)", "Total Mora (USD)", "% Mora Total"]
    col_w_camp = [Inches(3.3), Inches(1.9), Inches(1.9), Inches(1.7), Inches(1.5), Inches(1.22)]

    matrix_rows = camp_mat.get("matrix", [])
    # Tomar las principales 7 campañas y agrupar el resto
    top_camps = matrix_rows[:7]
    other_camps = matrix_rows[7:]

    num_rows_c = len(top_camps) + (1 if other_camps else 0) + 2
    t_shape_c = slide_camp.shapes.add_table(num_rows_c, len(cols_camp), Inches(0.8), Inches(1.85), Inches(11.52), Inches(4.8))
    table_c = t_shape_c.table
    _style_table_header(table_c, cols_camp, col_w_camp)

    for idx, c_info in enumerate(top_camps):
        t3_c = c_info.get("tramos_3", {})
        m_vc = t3_c.get("31+", {}).get("monto", 0.0)
        cnt_vc = t3_c.get("31+", {}).get("clientes", 0)
        m_vg = t3_c.get("16-30", {}).get("monto", 0.0)
        cnt_vg = t3_c.get("16-30", {}).get("clientes", 0)
        m_irr = t3_c.get("1-15", {}).get("monto", 0.0)
        cnt_irr = t3_c.get("1-15", {}).get("clientes", 0)

        row_vals = [
            c_info.get("campana", ""),
            f"${m_vc:,.2f}" + (f" ({cnt_vc})" if cnt_vc > 0 else ""),
            f"${m_vg:,.2f}" + (f" ({cnt_vg})" if cnt_vg > 0 else ""),
            f"${m_irr:,.2f}" + (f" ({cnt_irr})" if cnt_irr > 0 else ""),
            f"${c_info.get('total_monto', 0.0):,.2f}",
            f"{c_info.get('porcentaje', 0.0):.2f}%"
        ]
        _format_table_data_row(table_c, idx + 1, row_vals)

    if other_camps:
        o_vc = sum(x.get("tramos_3", {}).get("31+", {}).get("monto", 0.0) for x in other_camps)
        o_vg = sum(x.get("tramos_3", {}).get("16-30", {}).get("monto", 0.0) for x in other_camps)
        o_irr = sum(x.get("tramos_3", {}).get("1-15", {}).get("monto", 0.0) for x in other_camps)
        o_tot = sum(x.get("total_monto", 0.0) for x in other_camps)
        o_pct = sum(x.get("porcentaje", 0.0) for x in other_camps)
        other_vals = [
            f"Resto de Campañas ({len(other_camps)} adicionales)",
            f"${o_vc:,.2f}",
            f"${o_vg:,.2f}",
            f"${o_irr:,.2f}",
            f"${o_tot:,.2f}",
            f"{o_pct:.2f}%"
        ]
        _format_table_data_row(table_c, len(top_camps) + 1, other_vals)

    # Fila Total
    tot_c_vals = [
        "TOTAL GENERAL",
        f"${vencido_contable_monto:,.2f}",
        f"${sum(t.get('total_monto', 0.0) for t in tramos_3_list if t.get('id') == '16-30'):,.2f}",
        f"${sum(t.get('total_monto', 0.0) for t in tramos_3_list if t.get('id') == '1-15'):,.2f}",
        f"${kpis['monto_total']:,.2f}",
        "100.00%"
    ]
    _format_table_data_row(table_c, num_rows_c - 1, tot_c_vals, is_total=True)
    _add_running_footer(slide_camp, cutoff_date, 7)

    # =========================================================
    # DIAPOSITIVA 8: LECTURA GERENCIAL Y PLAN DE ACCIÓN DIRECTIVO
    # =========================================================
    slide_action = prs.slides.add_slide(blank_layout)
    _add_running_header(
        slide_action,
        "ESTRATEGIA DE RECOBRO",
        "Lectura Gerencial y Plan de Acción Directivo",
        "Recomendaciones operativas prioritarias de recuperación de liquidez"
    )

    top_camp_name = matrix_rows[0].get("campana", "Principal") if matrix_rows else "General"
    top_camp_monto = matrix_rows[0].get("total_monto", 0.0) if matrix_rows else kpis["monto_total"]

    hallazgos_list = [
        f"Estructura de Salud Óptima: El porcentaje total de morosidad es de {pct_mora_macro:.2f}% (${kpis['monto_total']:,.2f} USD) sobre una cartera activa de ${cartera_activa_usd:,.0f} USD.",
        f"Concentración por Programa: La campaña '{top_camp_name}' lidera el saldo impago con ${top_camp_monto:,.2f} USD, representando el mayor foco de recuperación.",
        f"Foco Vencido Contable: La mora de más de 30 días asciende a ${vencido_contable_monto:,.2f} USD ({pct_vc_activa:.2f}% de la Cartera Activa Total), requiriendo acuerdos de pago acelerados."
    ]

    plan_list = [
        f"Barrido Preventivo Digital (1-15 días): Notificar de inmediato vía WhatsApp masivo a los clientes en tramo Irregular (${mora_temprana_monto:,.2f} USD) antes de que escalen a mora de gestión.",
        f"Cobranza Directa Focalizada: Contactar telefónicamente de forma prioritaria a los clientes del tramo Vencido Contable (${vencido_contable_monto:,.2f} USD) para suscribir promesas de pago formales.",
        "Apalancamiento de Servicio y Depuración: Coordinar con operaciones la gestión sobre servicios activos y verificar las asignaciones de campaña en CRM para registros sin categorizar."
    ]

    card_w_act = Inches(5.6)
    card_h_act = Inches(4.8)
    _add_callout_box(slide_action, Inches(0.8), Inches(1.85), card_w_act, card_h_act, "Hallazgos Clave", hallazgos_list)
    _add_callout_box(slide_action, Inches(6.72), Inches(1.85), card_w_act, card_h_act, "Plan de Acción Operativo", plan_list)
    _add_running_footer(slide_action, cutoff_date, 8)

    # Exportar a BytesIO
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.getvalue()
