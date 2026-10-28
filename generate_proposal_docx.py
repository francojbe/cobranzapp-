import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def set_table_borders(table, color="CBD5E1", sz="4", val="single"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:left w:val="none"/>'
        f'<w:right w:val="none"/>'
        f'<w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideV w:val="none"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

def make_callout(doc, title, text, bg_hex="F8FAFC", border_color="2563EB"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    tbl.columns[0].width = Inches(6.5)
    
    cell = tbl.cell(0, 0)
    set_cell_background(cell, bg_hex)
    set_cell_margins(cell, top=140, bottom=140, left=200, right=200)
    
    # Left border only
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top w:val="none"/>'
        f'<w:left w:val="single" w:sz="24" w:space="0" w:color="{border_color}"/>'
        f'<w:bottom w:val="none"/>'
        f'<w:right w:val="none"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(borders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(4)
    run_t = p.add_run(f"📌 {title}\n")
    run_t.font.name = "Calibri"
    run_t.font.size = Pt(11)
    run_t.font.bold = True
    run_t.font.color.rgb = RGBColor(15, 41, 74)
    
    run_b = p.add_run(text)
    run_b.font.name = "Calibri"
    run_b.font.size = Pt(10)
    run_b.font.color.rgb = RGBColor(51, 65, 85)
    
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

def build_docx(output_path):
    doc = docx.Document()
    
    # Page setup (Letter, 0.8 in margins)
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.85)
        section.right_margin = Inches(0.85)
        
        # Header & Footer
        header = section.header
        hp = header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hrun = hp.add_run("Efinnovation · Propuesta Técnica y Comercial | CobranzApp")
        hrun.font.name = "Calibri"
        hrun.font.size = Pt(8.5)
        hrun.font.color.rgb = RGBColor(148, 163, 184)
        
        footer = section.footer
        fp = footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        frun = fp.add_run("Efinnovation — Consultoría, Software a Medida & Ciberseguridad · efinnovation.cl")
        frun.font.name = "Calibri"
        frun.font.size = Pt(8.5)
        frun.font.color.rgb = RGBColor(148, 163, 184)

    # Palette
    NAVY = RGBColor(15, 41, 74)       # #0F294A
    BLUE = RGBColor(37, 99, 235)      # #2563EB
    DARK_TEXT = RGBColor(30, 41, 59)  # #1E293B
    MUTED = RGBColor(100, 116, 139)   # #64748B

    # Document Header / Brand Header
    logo_path = r"C:\Users\franc\OneDrive\Documentos\dev\efinnovation-redesign\assets\img\hero-logo.png"
    if os.path.exists(logo_path):
        header_table = doc.add_table(rows=1, cols=2)
        header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        header_table.columns[0].width = Inches(4.5)
        header_table.columns[1].width = Inches(2.2)
        
        c0 = header_table.cell(0, 0)
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_before = Pt(0)
        p0.paragraph_format.space_after = Pt(2)
        r0 = p0.add_run("EFINNOVATION")
        r0.font.name = "Calibri"
        r0.font.size = Pt(20)
        r0.font.bold = True
        r0.font.color.rgb = NAVY
        
        p0_sub = c0.add_paragraph()
        p0_sub.paragraph_format.space_before = Pt(0)
        p0_sub.paragraph_format.space_after = Pt(0)
        r0_sub = p0_sub.add_run("Consultoría, Software a Medida e Inteligencia de Datos")
        r0_sub.font.name = "Calibri"
        r0_sub.font.size = Pt(10)
        r0_sub.font.color.rgb = BLUE
        
        c1 = header_table.cell(0, 1)
        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p1.paragraph_format.space_before = Pt(0)
        p1.paragraph_format.space_after = Pt(0)
        try:
            p1.add_run().add_picture(logo_path, width=Inches(1.8))
        except Exception:
            pass
    else:
        p_brand = doc.add_paragraph()
        p_brand.paragraph_format.space_before = Pt(0)
        p_brand.paragraph_format.space_after = Pt(2)
        r_b = p_brand.add_run("EFINNOVATION")
        r_b.font.name = "Calibri"
        r_b.font.size = Pt(20)
        r_b.font.bold = True
        r_b.font.color.rgb = NAVY

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # Document Title
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(6)
    p_title.paragraph_format.space_after = Pt(2)
    r_title = p_title.add_run("PROPUESTA TÉCNICA Y COMERCIAL")
    r_title.font.name = "Calibri"
    r_title.font.size = Pt(22)
    r_title.font.bold = True
    r_title.font.color.rgb = NAVY

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(12)
    r_sub = p_sub.add_run("Sistema de Inteligencia y Gestión de Cobranza (CobranzApp v1.0)")
    r_sub.font.name = "Calibri"
    r_sub.font.size = Pt(13)
    r_sub.font.color.rgb = BLUE

    # Meta Table (Author, Client, Date)
    meta_table = doc.add_table(rows=4, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_table.columns[0].width = Inches(2.0)
    meta_table.columns[1].width = Inches(4.7)
    set_table_borders(meta_table, color="E2E8F0", sz="4")
    
    meta_data = [
        ("Preparado por:", "Franco Blanco — Fundador & Director Técnico, Efinnovation\nIngeniero en Informática · Magíster en Seguridad de la Información"),
        ("Destinatario:", "Dirección General & Gerencia de Operaciones y Cobranza"),
        ("Contacto Directo:", "contacto@efinnovation.cl | +56 9 7426 3408 | https://efinnovation.cl"),
        ("Fecha de Emisión:", "Octubre 2026 · Santiago de Chile")
    ]
    for i, (label, val) in enumerate(meta_data):
        c_lbl = meta_table.cell(i, 0)
        c_val = meta_table.cell(i, 1)
        set_cell_background(c_lbl, "F8FAFC")
        set_cell_background(c_val, "FFFFFF")
        set_cell_margins(c_lbl, 80, 80, 100, 100)
        set_cell_margins(c_val, 80, 80, 100, 100)
        
        p_lbl = c_lbl.paragraphs[0]
        p_lbl.paragraph_format.space_before = Pt(0)
        p_lbl.paragraph_format.space_after = Pt(0)
        r_l = p_lbl.add_run(label)
        r_l.font.name = "Calibri"
        r_l.font.size = Pt(9.5)
        r_l.font.bold = True
        r_l.font.color.rgb = NAVY
        
        p_val = c_val.paragraphs[0]
        p_val.paragraph_format.space_before = Pt(0)
        p_val.paragraph_format.space_after = Pt(0)
        r_v = p_val.add_run(val)
        r_v.font.name = "Calibri"
        r_v.font.size = Pt(9.5)
        r_v.font.color.rgb = DARK_TEXT

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    def add_h1(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(16)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = "Calibri"
        r.font.size = Pt(14)
        r.font.bold = True
        r.font.color.rgb = NAVY
        return p

    def add_h2(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        r = p.add_run(text)
        r.font.name = "Calibri"
        r.font.size = Pt(11.5)
        r.font.bold = True
        r.font.color.rgb = BLUE
        return p

    def add_p(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = 1.15
        r = p.add_run(text)
        r.font.name = "Calibri"
        r.font.size = Pt(10.5)
        r.font.color.rgb = DARK_TEXT
        return p

    # --- 1. RESUMEN EJECUTIVO ---
    add_h1("1. Resumen Ejecutivo")
    add_p(
        "En la gestión de carteras de telecomunicaciones con más de 150.000 suscriptores, procesar planillas "
        "Excel manualmente genera retrasos operativos, errores en el cruce de teléfonos y pérdida de "
        "oportunidades de recaudación en los tramos críticos de mora."
    )
    add_p(
        "Desde Efinnovation, hemos diseñado y construido CobranzApp: una plataforma web hecha a la medida de su "
        "operación real. A diferencia de un software genérico que obliga a la empresa a cambiar su forma de trabajar, "
        "CobranzApp se adapta exactamente a sus archivos de facturación, su base maestra de contactos y sus reglas de "
        "negocio (Fibex vs. NO Fibex, 5 tramos de antigüedad y 26 campañas comerciales activas)."
    )
    add_p(
        "El sistema ya se encuentra 100% operativo, refactorizado bajo una arquitectura modular de grado "
        "empresarial y validado con su propia data real (154.009 registros), listo para pasar a producción y servir "
        "como la base tecnológica para digitalizar futuros departamentos de la organización."
    )

    # --- 2. ALCANCE TÉCNICO ---
    add_h1("2. Alcance Técnico del Sistema Entregado")
    add_p("CobranzApp resuelve de punta a punta el ciclo operativo de cobranza mediante 6 módulos integrados:")

    scope_table = doc.add_table(rows=7, cols=3)
    scope_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    scope_table.columns[0].width = Inches(1.8)
    scope_table.columns[1].width = Inches(2.9)
    scope_table.columns[2].width = Inches(2.0)
    set_table_borders(scope_table, color="CBD5E1", sz="4")

    headers = ["Módulo del Sistema", "Capacidad Técnica Construida", "Impacto Operativo Directo"]
    for j, h in enumerate(headers):
        c = scope_table.cell(0, j)
        set_cell_background(c, "0F294A")
        set_cell_margins(c, 100, 100, 100, 100)
        p = c.paragraphs[0]
        r = p.add_run(h)
        r.font.name = "Calibri"
        r.font.size = Pt(9.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)

    modules_data = [
        ("1. Ingesta Inteligente y Normalización",
         "Lectura automática de reportes de facturación, limpieza de formatos contables y clasificación instantánea en 5 tramos de mora (0–30, 31–60, 61–90, 91–120 y +120 días).",
         "Elimina 3 a 4 horas diarias de trabajo manual con planillas Excel y riesgos de fórmulas rotas."),
        ("2. Motor de Cruce Ultrarrápido (SQLite)",
         "Base de datos local indexada con 154.009 contactos. Cruce automático por Cédula/RIF y Contrato con validación y formateo de teléfonos móviles reales en < 2 segundos.",
         "Resuelve el desfase histórico entre el reporte de saldos pendientes y los teléfonos de contacto."),
        ("3. Inteligencia Financiera y Comparador",
         "Dashboard interactivo con KPIs de deuda, distribución por tramos, comparativa entre cortes de fecha y el nuevo indicador estratégico Fibex vs. NO Fibex.",
         "Visibilidad inmediata para Gerencia sobre recuperación de cartera y concentración de riesgo."),
        ("4. Matriz de 26 Campañas Comerciales",
         "Segmentación automática de la cartera aplicando reglas específicas de exoneración, descuento o beneficio según tramo de mora y tipo de plan.",
         "Garantiza que cada cliente reciba exactamente la oferta comercial aprobada por la empresa."),
        ("5. Módulo Call Center & WhatsApp Directo",
         "Bandeja de gestión con búsqueda instantánea, filtros combinados y envío de WhatsApp con plantillas dinámicas autollenadas (nombre, monto, tramo, oferta y fecha).",
         "Acelera drásticamente la contactabilidad de los gestores sin costo de licencias por usuario."),
        ("6. Centro de Reportería Ejecutiva",
         "Generación en un clic de 9 reportes descargables en Excel (.xlsx con diseño profesional multipestaña) y PDF corporativos listos para directorio o auditoría.",
         "Estandariza la presentación formal de resultados ante jefaturas y dirección.")
    ]

    for i, row in enumerate(modules_data, start=1):
        bg = "FFFFFF" if i % 2 != 0 else "F8FAFC"
        for j, val in enumerate(row):
            c = scope_table.cell(i, j)
            set_cell_background(c, bg)
            set_cell_margins(c, 80, 80, 100, 100)
            p = c.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val)
            r.font.name = "Calibri"
            r.font.size = Pt(9)
            if j == 0:
                r.font.bold = True
                r.font.color.rgb = NAVY
            else:
                r.font.color.rgb = DARK_TEXT

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # --- 3. DIFERENCIAL EFINNOVATION ---
    add_h1("3. El Diferencial Efinnovation: Ciberseguridad y Arquitectura")
    add_p(
        "A diferencia de desarrollos informales o agencias que subcontratan, Efinnovation aplica estándares "
        "rigurosos de ingeniería:"
    )
    
    bullets = [
        ("Seguridad de la Información desde el Diseño: ",
         "La base de datos de 154.009 suscriptores y los saldos financieros son activos confidenciales. Implementamos principio de mínimo privilegio, control estricto de accesos y soberanía de datos, resguardando la información bajo la Ley 19.628 de Protección de Datos Personales."),
        ("Trato Senior Directo: ",
         "El proyecto es liderado de principio a fin por Franco Blanco (Ingeniero en Informática y Magíster en Seguridad de la Información), asegurando un interlocutor técnico del más alto nivel sin pérdida de tiempo ni intermediarios."),
        ("Arquitectura Modular Reutilizable (core/): ",
         "El código fue refactorizado en componentes desacoplados (config, database, parsers, analytics y reports). Esta decisión de ingeniería significa que el cliente ya tiene construida la base para futuros sistemas departamentales.")
    ]
    for b_title, b_desc in bullets:
        bp = doc.add_paragraph(style='List Bullet')
        bp.paragraph_format.space_before = Pt(2)
        bp.paragraph_format.space_after = Pt(4)
        bp.paragraph_format.line_spacing = 1.15
        r_bt = bp.add_run(b_title)
        r_bt.font.name = "Calibri"
        r_bt.font.size = Pt(10)
        r_bt.font.bold = True
        r_bt.font.color.rgb = NAVY
        r_bd = bp.add_run(b_desc)
        r_bd.font.name = "Calibri"
        r_bd.font.size = Pt(10)
        r_bd.font.color.rgb = DARK_TEXT

    # --- 4. PROPUESTA ECONÓMICA ---
    add_h1("4. Propuesta Económica e Inversión (CLP)")
    add_p(
        "La inversión contempla el valor por el desarrollo del software a medida (licencia perpetua de uso para su "
        "empresa) más la modalidad de infraestructura que mejor se adapte a sus políticas de TI:"
    )

    add_h2("4.1 Desarrollo del Software Base CobranzApp v1.0")
    add_p(
        "• Valor del Desarrollo Base: $2.000.000 CLP (Pago único)\n"
        "Incluye la entrega completa de los 6 módulos operativos, la base maestra SQLite optimizada con 154.009 clientes, "
        "el módulo Fibex vs. NO Fibex, la matriz de 26 campañas, los 9 reportes ejecutivos en Excel y PDF, y la capacitación operativa al equipo. "
        "Sin cobro de licencias mensuales por cantidad de usuarios."
    )

    add_h2("4.2 Modalidades de Despliegue e Infraestructura")

    add_p("El cliente puede seleccionar una de las siguientes dos modalidades según su preferencia operativa:")

    # Box Opción A
    add_p("OPCIÓN A — Despliegue en Servidores del Cliente (Modalidad On-Premise)")
    p_optA = doc.add_paragraph()
    p_optA.paragraph_format.left_indent = Inches(0.2)
    p_optA.paragraph_format.space_after = Pt(4)
    r_a = p_optA.add_run(
        "• Desarrollo Sistema Base: $2.000.000 CLP (Pago único)\n"
        "• Servicio de Configuración, Despliegue y Hardening en Servidor del Cliente: $480.000 CLP (Pago único)\n"
        "  - Preparación del entorno de ejecución (Linux/Windows Server o Docker).\n"
        "  - Configuración de certificados SSL/HTTPS, proxy inverso y reglas de seguridad.\n"
        "  - Migración e indexación de la base de datos cobranzapp.db.\n"
        "  - Script de respaldos automáticos locales.\n"
        "  - Traspaso técnico documentado al equipo de TI interno + 30 días de garantía.\n"
        "• Total Inversión Única (Llave en mano en su servidor): $2.480.000 CLP\n"
        "• Costo mensual posterior: $0 CLP (mantenido por el personal de TI del cliente)."
    )
    r_a.font.name = "Calibri"
    r_a.font.size = Pt(9.5)
    r_a.font.color.rgb = DARK_TEXT

    # Box Opción B
    add_p("OPCIÓN B — Alojamiento y Operación en Servidor Efinnovation (Modalidad Cloud Gestionada) ⭐ Recomendada")
    p_optB = doc.add_paragraph()
    p_optB.paragraph_format.left_indent = Inches(0.2)
    p_optB.paragraph_format.space_after = Pt(6)
    r_b = p_optB.add_run(
        "• Desarrollo Sistema Base: $2.000.000 CLP (Pago único)\n"
        "• Instalación y Configuración Inicial en Servidor Efinnovation: $0 CLP (100% Bonificado)\n"
        "• Plan Mensual de Servidor Privado, Seguridad y Soporte: $190.000 CLP / mes\n"
        "  - Servidor Cloud de alto rendimiento con disponibilidad 24/7.\n"
        "  - Certificado SSL/HTTPS y monitoreo perimetral de seguridad.\n"
        "  - Respaldos automáticos diarios de la base de datos con retención histórica.\n"
        "  - Soporte evolutivo de hasta 4 horas mensuales (cambios de textos de campañas, ajustes de tramos, recarga de teléfonos).\n"
        "  - Servidor listo para alojar futuros sistemas de otros departamentos sin costo de arriendo adicional.\n"
        "• Total Inversión Inicial: $2.000.000 CLP + $190.000 CLP/mes de operación."
    )
    r_b.font.name = "Calibri"
    r_b.font.size = Pt(9.5)
    r_b.font.color.rgb = DARK_TEXT

    # Callout aclaratorio IMPORTANTE
    make_callout(
        doc,
        title="ACLARACIÓN EXCLUSIVA SOBRE EL PLAN MENSUAL ($190.000/mes)",
        text="Los $190.000 CLP / mes corresponden exclusivamente al arriendo del servidor cloud, monitoreo de ciberseguridad, respaldos diarios y soporte operativo continuo de CobranzApp.\n"
             "Los nuevos sistemas que se desarrollen para otros departamentos (Finanzas, Ventas, Operaciones) se cotizan y cobran como proyectos de software independientes, pero con una ventaja económica clave: no tendrán que pagar arriendo de un nuevo servidor, ya que se hospedarán en esta misma infraestructura.",
        bg_hex="EFF6FF",
        border_color="2563EB"
    )

    # Tabla Comparativa
    add_h2("4.3 Tabla Comparativa Resumen")
    comp_table = doc.add_table(rows=8, cols=3)
    comp_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    comp_table.columns[0].width = Inches(3.1)
    comp_table.columns[1].width = Inches(1.8)
    comp_table.columns[2].width = Inches(1.8)
    set_table_borders(comp_table, color="CBD5E1", sz="4")

    c_headers = ["Concepto", "Opción A: Servidor Cliente", "Opción B: Servidor Efinnovation"]
    for j, h in enumerate(c_headers):
        c = comp_table.cell(0, j)
        set_cell_background(c, "0F294A")
        set_cell_margins(c, 90, 90, 90, 90)
        p = c.paragraphs[0]
        r = p.add_run(h)
        r.font.name = "Calibri"
        r.font.size = Pt(9)
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)

    comp_rows = [
        ("Desarrollo Software CobranzApp", "$2.000.000 CLP", "$2.000.000 CLP"),
        ("Configuración, Hardening y Puesta en Marcha", "$480.000 CLP (Pago único)", "$0 CLP (Bonificado)"),
        ("Servidor Cloud Dedicado y SSL 24/7", "A cargo de TI del cliente", "Incluido"),
        ("Respaldos Diarios Automáticos", "A cargo de TI del cliente", "Incluido"),
        ("Soporte evolutivo continuo (ajuste campañas)", "Tarifa por hora / bolsa adicional", "Incluido (hasta 4 hrs/mes)"),
        ("Inversión Inicial para Salir a Producción", "$2.480.000 CLP", "$2.000.000 CLP"),
        ("Costo Operativo Mensual Posterior", "$0 CLP / mes", "$190.000 CLP / mes")
    ]
    for i, row in enumerate(comp_rows, start=1):
        bg = "FFFFFF" if i % 2 != 0 else "F8FAFC"
        for j, val in enumerate(row):
            c = comp_table.cell(i, j)
            set_cell_background(c, bg)
            set_cell_margins(c, 70, 70, 90, 90)
            p = c.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(val)
            r.font.name = "Calibri"
            r.font.size = Pt(8.5)
            if j == 0:
                r.font.bold = True
                r.font.color.rgb = NAVY
            elif "2.480.000" in val or "2.000.000" in val or "190.000" in val:
                r.font.bold = True
                r.font.color.rgb = BLUE
            else:
                r.font.color.rgb = DARK_TEXT

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # --- 5. EXPANSIÓN A OTROS DEPARTAMENTOS ---
    add_h1("5. Hoja de Ruta Estratégica: Expansión a Otros Departamentos")
    add_p(
        "CobranzApp no fue programado como un script aislado, sino como una solución modular enterprise. "
        "Al contratar este proyecto, su empresa ya construyó los cimientos tecnológicos necesarios para "
        "interconectar otras áreas de la organización."
    )
    add_p("¿Cómo funciona comercial y técnicamente para los próximos sistemas?")

    points = [
        ("Desarrollos independientes cotizados a la medida: ",
         "Cada nuevo sistema que requiera la empresa (por ejemplo, Conciliación Bancaria en Finanzas o Seguimiento de Cuadrillas en Operaciones) se cotiza como un proyecto específico según sus requerimientos puntuales."),
        ("Ahorro preferencial del 25% al 35% en desarrollo: ",
         "Como ya no partimos de una hoja en blanco y podemos reutilizar el motor analítico, el lector de archivos, la base maestra de clientes y la interfaz gráfica, el costo y tiempo de desarrollo para su empresa disminuye drásticamente. Un software que en el mercado costaría $2.000.000 CLP, con nosotros se cotiza en torno a $1.400.000 – $1.500.000 CLP."),
        ("Cero costo de nuevos servidores (con Opción B): ",
         "No tendrán que arrendar servidores adicionales para cada departamento. La infraestructura cloud de Efinnovation ya tiene la capacidad para recibir los nuevos módulos sin pagar arriendos paralelos."),
        ("Base de datos unificada (sin islas de información): ",
         "Si Operaciones o Ventas actualiza el número telefónico de un cliente, esa corrección beneficia de forma inmediata al equipo de Cobranza.")
    ]
    for p_title, p_desc in points:
        bp = doc.add_paragraph(style='List Bullet')
        bp.paragraph_format.space_before = Pt(2)
        bp.paragraph_format.space_after = Pt(4)
        bp.paragraph_format.line_spacing = 1.15
        r_pt = bp.add_run(p_title)
        r_pt.font.name = "Calibri"
        r_pt.font.size = Pt(10)
        r_pt.font.bold = True
        r_pt.font.color.rgb = NAVY
        r_pd = bp.add_run(p_desc)
        r_pd.font.name = "Calibri"
        r_pd.font.size = Pt(10)
        r_pd.font.color.rgb = DARK_TEXT

    # --- 6. CONDICIONES COMERCIALES ---
    add_h1("6. Condiciones Comerciales y Próximos Pasos")
    conds = [
        ("Modalidad de Pago: ", "50% al aceptar la propuesta e iniciar el despliegue / 50% contra entrega final en producción y validación del equipo."),
        ("Tiempos de Puesta en Marcha: ", "Opción B (Servidor Efinnovation): 24 a 48 horas hábiles. Opción A (Servidor Cliente): 3 a 5 días hábiles tras recibir accesos."),
        ("Confidencialidad y Propiedad: ", "Todos los datos de cartera de clientes y montos pertenecen con exclusividad al cliente. Efinnovation suscribe acuerdo de confidencialidad (NDA).")
    ]
    for c_t, c_d in conds:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(4)
        r_t = p.add_run(f"• {c_t}")
        r_t.font.name = "Calibri"
        r_t.font.size = Pt(10)
        r_t.font.bold = True
        r_t.font.color.rgb = NAVY
        r_d = p.add_run(c_d)
        r_d.font.name = "Calibri"
        r_d.font.size = Pt(10)
        r_d.font.color.rgb = DARK_TEXT

    # Firma
    doc.add_paragraph().paragraph_format.space_after = Pt(20)
    p_sign = doc.add_paragraph()
    p_sign.paragraph_format.keep_with_next = True
    r_s1 = p_sign.add_run("Franco Blanco\n")
    r_s1.font.name = "Calibri"
    r_s1.font.size = Pt(12)
    r_s1.font.bold = True
    r_s1.font.color.rgb = NAVY

    r_s2 = p_sign.add_run(
        "Fundador & Director Técnico — Efinnovation\n"
        "Ingeniero en Informática · Magíster en Seguridad de la Información\n"
        "🌐 https://efinnovation.cl | 📱 +56 9 7426 3408 | ✉️ contacto@efinnovation.cl"
    )
    r_s2.font.name = "Calibri"
    r_s2.font.size = Pt(9.5)
    r_s2.font.color.rgb = MUTED

    doc.save(output_path)
    print(f"Propuesta generada exitosamente en: {output_path}")

if __name__ == "__main__":
    out_file = r"c:\Users\franc\OneDrive\Documentos\dev\ProFrank\Propuesta_Comercial_CobranzApp_Efinnovation.docx"
    build_docx(out_file)
