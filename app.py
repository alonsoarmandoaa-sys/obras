import sqlite3
import pandas as pd
import streamlit as st
from datetime import date
import io

# Librerías para generación de PDF
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# -------------------------------------------------------------------
# CONFIGURACIÓN DE PÁGINA Y ESTILOS CSS
# -------------------------------------------------------------------
st.set_page_config(
    page_title="Portal Ejecutivo - Control de Obras",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS personalizados para un entorno limpio y profesional
st.markdown("""
<style>
    /* Estilo general de la app */
    .stApp {
        background-color: #F8FAFC;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    }
    
    /* Banner principal de Obra */
    .header-card {
        background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
        color: #FFFFFF;
        padding: 24px;
        border-radius: 16px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        margin-bottom: 24px;
    }
    .header-title {
        font-size: 1.4rem;
        font-weight: 700;
        color: #F8FAFC;
        margin-bottom: 12px;
    }
    .badge {
        background-color: rgba(255, 255, 255, 0.12);
        border: 1px solid rgba(255, 255, 255, 0.2);
        padding: 6px 14px;
        border-radius: 8px;
        font-size: 0.88rem;
        display: inline-block;
        margin-right: 10px;
        margin-top: 4px;
    }
    
    /* Cards para métricas clave */
    .kpi-card {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 14px;
        padding: 18px 20px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        transition: transform 0.2s ease;
    }
    .kpi-label {
        font-size: 0.82rem;
        color: #64748B;
        text-transform: uppercase;
        font-weight: 600;
        letter-spacing: 0.5px;
    }
    .kpi-value {
        font-size: 1.5rem;
        font-weight: 700;
        color: #0F172A;
        margin-top: 4px;
    }
    .kpi-sub {
        font-size: 0.8rem;
        color: #2563EB;
        font-weight: 600;
        margin-top: 2px;
    }

    /* Ocultar elementos innecesarios */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

DB_NAME = "control_obra.db"

# -------------------------------------------------------------------
# FUNCIONES DE BASE DE DATOS Y AUXILIARES
# -------------------------------------------------------------------
def get_connection():
    return sqlite3.connect(DB_NAME, check_same_thread=False)

def limpiar_texto(val):
    if pd.isna(val) or val is None or str(val).strip().lower() in ["", "nan", "none", "general", "null"]:
        return "-"
    return str(val).strip()

def get_lista_puentes():
    conn = get_connection()
    cursor = conn.cursor()
    rows = []
    
    try:
        cursor.execute("SELECT id_puente, nombre_contrato FROM proyecto")
        rows = cursor.fetchall()
    except Exception:
        rows = []
    
    if not rows:
        try:
            cursor.execute("SELECT DISTINCT id_puente FROM catalogo")
            rows = [(r[0], f"Obra / Puente {r[0]}") for r in cursor.fetchall()]
        except Exception:
            rows = []
            
    conn.close()
    
    dict_obras = {}
    for row in rows:
        id_p = row[0]
        nombre = str(row[1]).strip() if len(row) > 1 and row[1] else f"Obra / Puente {id_p}"
        dict_obras[id_p] = nombre
    return dict_obras

def get_proyecto_info(id_puente):
    conn = get_connection()
    cursor = conn.cursor()
    row = None
    
    try:
        cursor.execute(
            "SELECT nombre_contrato, num_contrato, contratista FROM proyecto WHERE CAST(id_puente AS TEXT) = CAST(? AS TEXT)", 
            (str(id_puente),)
        )
        row = cursor.fetchone()
    except Exception:
        row = None
    finally:
        conn.close()
    
    if row:
        nombre = str(row[0]).strip() if row[0] and str(row[0]).strip() != "" else f"OBRA / PUENTE {id_puente}"
        num = str(row[1]).strip() if row[1] and str(row[1]).strip() != "" else "SIN ASIGNAR"
        contratista = str(row[2]).strip() if row[2] and str(row[2]).strip() != "" else "CONTRATISTA NO ESPECIFICADO"
        return {"nombre_contrato": nombre, "num_contrato": num, "contratista": contratista}
    
    return {
        "nombre_contrato": f"OBRA / PUENTE {id_puente}",
        "num_contrato": "SIN ASIGNAR",
        "contratista": "EMPRESA CONTRATISTA"
    }

# -------------------------------------------------------------------
# GENERACIÓN DE REPORTES (EXCEL Y PDF)
# -------------------------------------------------------------------
def generar_excel_bd(id_puente):
    buffer = io.BytesIO()
    conn = get_connection()
    
    query_general = '''
        SELECT 
            c.id_concepto AS "No.",
            c.partida AS "Partida",
            c.subpartida AS "Subpartida",
            c.norma AS "Norma / Espec.",
            c.descripcion AS "Descripción",
            c.unidad AS "Unidad",
            c.cantidad_contratada AS "Contratado",
            c.precio_unitario AS "Precio Unitario",
            (c.cantidad_contratada * c.precio_unitario) AS "Importe Contratado",
            COALESCE(SUM(r.cantidad), 0) AS "Total Ejecutado",
            (COALESCE(SUM(r.cantidad), 0) * c.precio_unitario) AS "Importe Ejecutado",
            COALESCE(SUM(CASE WHEN r.num_estimacion != 'Ejecutado (No estimado)' THEN r.cantidad ELSE 0 END), 0) AS "Total Estimado",
            (COALESCE(SUM(CASE WHEN r.num_estimacion != 'Ejecutado (No estimado)' THEN r.cantidad ELSE 0 END), 0) * c.precio_unitario) AS "Importe Estimado",
            (c.cantidad_contratada - COALESCE(SUM(r.cantidad), 0)) AS "Saldo Disponible",
            ROUND((COALESCE(SUM(r.cantidad), 0) / NULLIF(c.cantidad_contratada, 0)) * 100, 2) AS "% Avance Vol."
        FROM catalogo c
        LEFT JOIN consumos r ON c.id_concepto = r.id_concepto AND CAST(c.id_puente AS TEXT) = CAST(r.id_puente AS TEXT)
        WHERE CAST(c.id_puente AS TEXT) = CAST(? AS TEXT)
        GROUP BY c.id_concepto
        ORDER BY CAST(c.id_concepto AS INTEGER) ASC, c.id_concepto ASC
    '''
    df_general = pd.read_sql_query(query_general, conn, params=(str(id_puente),))

    query_consumos = '''
        SELECT 
            r.id AS "ID Registro",
            r.fecha AS "Fecha",
            r.id_concepto AS "No. Concepto",
            c.descripcion AS "Descripción Concepto",
            r.ubicacion AS "Ubicación",
            r.cantidad AS "Cantidad Registrada",
            c.unidad AS "Unidad",
            r.num_estimacion AS "Estimación / Estatus",
            r.remision AS "Folio Remisión"
        FROM consumos r
        JOIN catalogo c ON r.id_concepto = c.id_concepto AND CAST(r.id_puente AS TEXT) = CAST(c.id_puente AS TEXT)
        WHERE CAST(r.id_puente AS TEXT) = CAST(? AS TEXT)
        ORDER BY r.id DESC
    '''
    df_consumos = pd.read_sql_query(query_consumos, conn, params=(str(id_puente),))
    conn.close()

    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_general.to_excel(writer, sheet_name="Resumen General", index=False)
        df_consumos.to_excel(writer, sheet_name="Historial y Estimaciones", index=False)

    buffer.seek(0)
    return buffer

def generar_pdf_general(df, info_proj):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(letter), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    elements = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=13, leading=15, textColor=colors.HexColor('#1E3A8A'))
    cell_style = ParagraphStyle('CellStyle', parent=styles['Normal'], fontSize=6.5, leading=8)
    meta_style = ParagraphStyle('MetaStyle', parent=styles['Normal'], fontSize=7.5, leading=9.5)

    elements.append(Paragraph("Reporte General de Control de Volúmenes y Avance Financiero", title_style))
    elements.append(Spacer(1, 4))

    monto_contrato = (df["Contratado"] * df["P.U."]).sum() if "P.U." in df else 0
    monto_ejecutado = (df["Ejecutado"] * df["P.U."]).sum() if "P.U." in df else 0
    monto_estimado = (df["Estimado"] * df["P.U."]).sum() if "P.U." in df else 0
    pct_ejecutado = (monto_ejecutado / monto_contrato * 100) if monto_contrato > 0 else 0

    data_header = [
        [
            Paragraph(f"<b>Obra / Puente:</b> {info_proj['nombre_contrato']}", meta_style),
            Paragraph(f"<b>Nº Contrato:</b> {info_proj['num_contrato']}", meta_style),
            Paragraph(f"<b>Monto Contrato:</b> ${monto_contrato:,.2f}", meta_style)
        ],
        [
            Paragraph(f"<b>Empresa Contratista:</b> {info_proj['contratista']}", meta_style),
            Paragraph(f"<b>Fecha Emisión:</b> {date.today().strftime('%d/%m/%Y')}", meta_style),
            Paragraph(f"<b>% Avance Financiero:</b> {pct_ejecutado:.2f}%", meta_style)
        ]
    ]
    t_header = Table(data_header, colWidths=[310, 230, 210])
    t_header.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#94A3B8')),
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F1F5F9')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
    ]))
    elements.append(t_header)
    elements.append(Spacer(1, 8))

    data = [["No.", "Concepto", "Partida", "Unidad", "Contratado", "P.U.", "Importe Total", "Ejecutado", "Estimado", "Saldo", "% Av."]]

    for _, row in df.iterrows():
        desc_txt = str(row["Descripción"]).strip() if pd.notna(row["Descripción"]) else "-"
        pu_val = row.get("P.U.", 0)
        imp_total = row["Contratado"] * pu_val
        
        data.append([
            str(row["No."]),
            Paragraph(desc_txt, cell_style),
            Paragraph(limpiar_texto(row["Partida"]), cell_style),
            str(row["Unidad"]),
            f"{row['Contratado']:,.2f}",
            f"${pu_val:,.2f}",
            f"${imp_total:,.2f}",
            f"{row['Ejecutado']:,.2f}",
            f"{row['Estimado']:,.2f}",
            f"{row['Saldo Available']:,.2f}",
            f"{row['% Avance']:.1f}%"
        ])

    col_widths = [25, 180, 80, 35, 55, 50, 65, 55, 55, 55, 40]
    t = Table(data, repeatRows=1, colWidths=col_widths)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1E3A8A')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 6.5),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (0, 0), (0, -1), 'CENTER'),
        ('ALIGN', (3, 0), (-1, -1), 'CENTER'),
    ]))
    elements.append(t)
    doc.build(elements)
    buffer.seek(0)
    return buffer

def generar_pdf_concepto_individual(info_concepto, df_historia, info_proj):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    elements = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=13, leading=15, textColor=colors.HexColor('#1E3A8A'))
    cell_style = ParagraphStyle('CellStyle', parent=styles['Normal'], fontSize=7.5, leading=9.5)
    cell_style_bold = ParagraphStyle('CellStyleBold', parent=styles['Normal'], fontSize=8, leading=10, fontName='Helvetica-Bold')
    meta_style = ParagraphStyle('MetaStyle', parent=styles['Normal'], fontSize=8, leading=10)

    elements.append(Paragraph("<b>REPORTE DETALLADO DE CONCEPTO</b>", title_style))
    elements.append(Spacer(1, 6))

    data_header = [
        [
            Paragraph(f"<b>Obra / Puente:</b> {info_proj['nombre_contrato']}", meta_style),
            Paragraph(f"<b>Nº Contrato:</b> {info_proj['num_contrato']}", meta_style)
        ],
        [
            Paragraph(f"<b>Empresa Contratista:</b> {info_proj['contratista']}", meta_style),
            Paragraph(f"<b>Fecha de Emisión:</b> {date.today().strftime('%d/%m/%Y')}", meta_style)
        ]
    ]
    t_header = Table(data_header, colWidths=[330, 222])
    t_header.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#94A3B8')),
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F1F5F9')),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]))
    elements.append(t_header)
    elements.append(Spacer(1, 10))

    no_item = info_concepto['id_concepto']
    desc = info_concepto['descripcion']
    partida = limpiar_texto(info_concepto.get('partida', '-'))
    unidad = info_concepto['unidad']
    cant_contratada = info_concepto['cantidad_contratada']
    pu = info_concepto.get('precio_unitario', 0.0)
    imp_contratado = cant_contratada * pu
    
    cant_ejecutada = info_concepto.get('total_ejecutado', 0.0)
    saldo = cant_contratada - cant_ejecutada
    pct_avance = (cant_ejecutada / cant_contratada * 100) if cant_contratada > 0 else 0

    elements.append(Paragraph(f"<b>FICHA TÉCNICA DEL CONCEPTO NO. {no_item}</b>", ParagraphStyle('Sub', parent=styles['Heading2'], fontSize=10, textColor=colors.HexColor('#1E3A8A'))))
    elements.append(Spacer(1, 4))

    data_concepto = [
        [Paragraph("<b>Clave / No. Concepto:</b>", meta_style), Paragraph(str(no_item), meta_style), Paragraph("<b>Partida:</b>", meta_style), Paragraph(str(partida), meta_style)],
        [Paragraph("<b>Descripción:</b>", meta_style), Paragraph(desc, cell_style), Paragraph("<b>Unidad de Medida:</b>", meta_style), Paragraph(str(unidad), meta_style)],
        [Paragraph("<b>Volumen Contratado:</b>", meta_style), Paragraph(f"{cant_contratada:,.2f} {unidad}", meta_style), Paragraph("<b>Precio Unitario (P.U.):</b>", meta_style), Paragraph(f"${pu:,.2f}", meta_style)],
        [Paragraph("<b>Importe Contratado:</b>", meta_style), Paragraph(f"${imp_contratado:,.2f}", meta_style), Paragraph("<b>% Avance Físico:</b>", meta_style), Paragraph(f"{pct_avance:.2f}%", meta_style)],
        [Paragraph("<b>Volumen Ejecutado:</b>", meta_style), Paragraph(f"{cant_ejecutada:,.2f} {unidad}", meta_style), Paragraph("<b>Saldo Disponible:</b>", meta_style), Paragraph(f"{saldo:,.2f} {unidad}", meta_style)],
    ]
    t_concepto = Table(data_concepto, colWidths=[110, 166, 110, 166])
    t_concepto.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F8FAFC')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]))
    elements.append(t_concepto)
    elements.append(Spacer(1, 12))

    elements.append(Paragraph("<b>HISTORIAL DETALLADO DE CAPTURAS / REGISTROS DE CAMPO</b>", ParagraphStyle('Sub2', parent=styles['Heading2'], fontSize=10, textColor=colors.HexColor('#1E3A8A'))))
    elements.append(Spacer(1, 4))

    data_hist = [["ID", "Fecha", "Ubicación / Elemento", "Cantidad", "Remisión", "Estimación / Estatus"]]

    if df_historia.empty:
        data_hist.append(["-", "-", "Sin capturas registradas", "-", "-", "-"])
    else:
        tot_cant = 0.0
        for _, r in df_historia.iterrows():
            cant_val = float(r["Cantidad"])
            tot_cant += cant_val
            id_reg = r["ID"] if "ID" in r else r.get("ID Registro", "-")
            data_hist.append([
                str(id_reg),
                str(r["Fecha"]),
                Paragraph(str(r["Ubicación"]), cell_style),
                f"{cant_val:,.2f}",
                str(r["Remisión"] if pd.notna(r["Remisión"]) else "-"),
                str(r["Estimación"])
            ])
        data_hist.append([
            Paragraph("<b>TOTAL</b>", cell_style_bold),
            "",
            "",
            Paragraph(f"<b>{tot_cant:,.2f} {unidad}</b>", cell_style_bold),
            "",
            ""
        ])

    t_hist = Table(data_hist, repeatRows=1, colWidths=[35, 65, 180, 75, 80, 117])
    t_hist.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1E3A8A')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 7.5),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (0, 0), (1, -1), 'CENTER'),
        ('ALIGN', (3, 0), (3, -1), 'RIGHT'),
    ]))
    elements.append(t_hist)

    doc.build(elements)
    buffer.seek(0)
    return buffer

# -------------------------------------------------------------------
# PANEL LATERAL (SIDEBAR): SELECCIÓN Y CONFIGURACIÓN
# -------------------------------------------------------------------
dict_puentes = get_lista_puentes()

with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/4300/4300058.png", width=60)
    st.title("Control de Obras")
    st.caption("Sistema Ejecutivo de Consulta")
    st.divider()

    if not dict_puentes:
        st.error("No se encontraron obras registradas.")
        st.stop()

    opciones_obras = {f"{nombre} (ID: {id_p})": id_p for id_p, nombre in dict_puentes.items()}
    obra_seleccionada_label = st.selectbox(
        "🏗️ Seleccionar Obra / Proyecto:",
        options=list(opciones_obras.keys())
    )
    
    id_puente = opciones_obras[obra_seleccionada_label]
    info_proyecto = get_proyecto_info(id_puente)
    
    st.divider()
    st.caption("🔍 Utiliza las pestañas superiores para navegar entre el resumen ejecutivo, catálogo y fichas detalladas.")

# -------------------------------------------------------------------
# OBTENCIÓN Y CÁLCULO DE DATOS
# -------------------------------------------------------------------
conn = get_connection()
query = '''
    SELECT 
        c.id_concepto AS "No.",
        c.partida AS "Partida",
        c.subpartida AS "Subpartida",
        c.norma AS "Norma / Espec.",
        c.descripcion AS "Descripción",
        c.unidad AS "Unidad",
        c.cantidad_contratada AS "Contratado",
        c.precio_unitario AS "P.U.",
        COALESCE(SUM(r.cantidad), 0) AS "Ejecutado",
        COALESCE(SUM(CASE WHEN r.num_estimacion != 'Ejecutado (No estimado)' THEN r.cantidad ELSE 0 END), 0) AS "Estimado",
        (c.cantidad_contratada - COALESCE(SUM(r.cantidad), 0)) AS "Saldo Available",
        ROUND((COALESCE(SUM(r.cantidad), 0) / NULLIF(c.cantidad_contratada, 0)) * 100, 2) AS "% Avance"
    FROM catalogo c
    LEFT JOIN consumos r ON c.id_concepto = r.id_concepto AND CAST(c.id_puente AS TEXT) = CAST(r.id_puente AS TEXT)
    WHERE CAST(c.id_puente AS TEXT) = CAST(? AS TEXT)
    GROUP BY c.id_concepto
    ORDER BY CAST(c.id_concepto AS INTEGER) ASC, c.id_concepto ASC
'''
df = pd.read_sql_query(query, conn, params=(str(id_puente),))
conn.close()

if df.empty:
    st.warning("⚠️ Esta obra aún no cuenta con un catálogo de conceptos registrado.")
    st.stop()

# Cálculos globales
monto_contratado = (df["Contratado"] * df["P.U."]).sum()
monto_ejecutado = (df["Ejecutado"] * df["P.U."]).sum()
monto_estimado = (df["Estimado"] * df["P.U."]).sum()
pct_financiero_ejecutado = (monto_ejecutado / monto_contratado * 100) if monto_contratado > 0 else 0
pct_financiero_estimado = (monto_estimado / monto_contratado * 100) if monto_contratado > 0 else 0

# -------------------------------------------------------------------
# ENCABEZADO PRINCIPAL (BANNER EJECUTIVO)
# -------------------------------------------------------------------
st.markdown(f"""
<div class="header-card">
    <div class="header-title">📌 {info_proyecto['nombre_contrato']}</div>
    <div>
        <span class="badge">📄 <b>Contrato:</b> {info_proyecto['num_contrato']}</span>
        <span class="badge">🏗️ <b>Contratista:</b> {info_proyecto['contratista']}</span>
        <span class="badge">🔑 <b>ID Obra:</b> {id_puente}</span>
    </div>
</div>
""", unsafe_allow_html=True)

# -------------------------------------------------------------------
# NAVEGACIÓN POR PESTAÑAS (TABS)
# -------------------------------------------------------------------
tab_resumen, tab_catalogo, tab_concepto = st.tabs([
    "📊 Resumen Ejecutivo", 
    "📋 Catálogo y Avances", 
    "🔍 Detalle por Concepto"
])

# -------------------------------------------------------------------
# TAB 1: RESUMEN EJECUTIVO (SOLO PORCENTAJES DE AVANCE)
# -------------------------------------------------------------------
with tab_resumen:
    st.subheader("Indicadores Clave de Desempeño (% Avance)")
    
    # Cálculos porcentuales
    pct_pendiente_estimar = pct_financiero_ejecutado - pct_financiero_estimado
    pct_saldo_ejecutar = 100.0 - pct_financiero_ejecutado
    
    c1, c2, c3, c4 = st.columns(4)
    
    with c1:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">% Avance Físico</div>
            <div class="kpi-value">{pct_financiero_ejecutado:.2f}%</div>
            <div class="kpi-sub">Ejecutado en Campo</div>
        </div>
        """, unsafe_allow_html=True)
        
    with c2:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">% Avance Estimado</div>
            <div class="kpi-value">{pct_financiero_estimado:.2f}%</div>
            <div class="kpi-sub">Total Facturado</div>
        </div>
        """, unsafe_allow_html=True)

    with c3:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">% Pendiente de Estimar</div>
            <div class="kpi-value">{max(pct_pendiente_estimar, 0.0):.2f}%</div>
            <div class="kpi-sub">Ejecutado sin Estimar</div>
        </div>
        """, unsafe_allow_html=True)

    with c4:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">% Saldo por Ejecutar</div>
            <div class="kpi-value">{max(pct_saldo_ejecutar, 0.0):.2f}%</div>
            <div class="kpi-sub">Trabajo Pendiente</div>
        </div>
        """, unsafe_allow_html=True)

    st.write("")
    st.write("")
    
    # Barras de progreso visuales
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        st.write("📈 **Progreso Físico (Ejecutado):**")
        st.progress(min(pct_financiero_ejecutado / 100, 1.0))
    with col_p2:
        st.write("📝 **Progreso Administrativo (Estimado):**")
        st.progress(min(pct_financiero_estimado / 100, 1.0))
# -------------------------------------------------------------------
# TAB 2: CATÁLOGO Y TABLA GENERAL
# -------------------------------------------------------------------
with tab_catalogo:
    col_t1, col_t2 = st.columns(2)
    
    with col_t1:
        pdf_gen = generar_pdf_general(df, info_proyecto)
        st.download_button(
            "📄 Reporte PDF",
            data=pdf_gen,
            file_name=f"Reporte_General_{id_puente}_{date.today()}.pdf",
            mime="application/pdf",
            use_container_width=True
        )
        
    with col_t2:
        excel_bd = generar_excel_bd(id_puente)
        st.download_button(
            "📊 Excel Completo",
            data=excel_bd,
            file_name=f"BaseDatos_{id_puente}_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    st.write("")

    st.dataframe(
        df.style.format({
            "Contratado": "{:,.2f}",
            "P.U.": "${:,.2f}",
            "Ejecutado": "{:,.2f}",
            "Estimado": "{:,.2f}",
            "Saldo Available": "{:,.2f}",
            "% Avance": "{:.2f}%"
        }),
        use_container_width=True,
        height=450
    )

# -------------------------------------------------------------------
# TAB 3: CONSULTA DETALLADA POR CONCEPTO
# -------------------------------------------------------------------
with tab_concepto:
    lista_conceptos = df["No."].tolist()
    
    col_c1, col_c2 = st.columns([3, 1])
    
    with col_c1:
        concepto_elegido = st.selectbox(
            "Selecciona el concepto que deseas analizar:",
            options=lista_conceptos,
            format_func=lambda x: f"Concepto {x} - {df[df['No.'] == x]['Descripción'].values[0][:70]}..."
        )

    if concepto_elegido:
        row_ind = df[df["No."] == concepto_elegido].iloc[0]
        
        conn = get_connection()
        query_hist_ind = """
            SELECT r.id AS "ID", r.fecha AS "Fecha", r.ubicacion AS "Ubicación", 
                   r.cantidad AS "Cantidad", r.remision AS "Remisión", r.num_estimacion AS "Estimación"
            FROM consumos r
            WHERE CAST(r.id_puente AS TEXT) = CAST(? AS TEXT) AND r.id_concepto = ?
            ORDER BY r.id ASC
        """
        df_hist_ind = pd.read_sql_query(query_hist_ind, conn, params=(str(id_puente), concepto_elegido))
        conn.close()

        # Ajustamos el índice para que inicie en 1 en lugar de 0
        if not df_hist_ind.empty:
            df_hist_ind.index = range(1, len(df_hist_ind) + 1)

        info_c_dict = {
            'id_concepto': concepto_elegido,
            'partida': row_ind["Partida"],
            'subpartida': row_ind["Subpartida"],
            'norma': row_ind["Norma / Espec."],
            'descripcion': row_ind["Descripción"],
            'unidad': row_ind["Unidad"],
            'cantidad_contratada': float(row_ind["Contratado"]),
            'precio_unitario': float(row_ind["P.U."]),
            'total_ejecutado': float(row_ind["Ejecutado"]),
            'total_estimado': float(row_ind["Estimado"])
        }

        pdf_indiv_bytes = generar_pdf_concepto_individual(info_c_dict, df_hist_ind, info_proyecto)

        with col_c2:
            st.write("")
            st.write("")
            st.download_button(
                f"📄 Descargar PDF ({concepto_elegido})",
                data=pdf_indiv_bytes,
                file_name=f"Ficha_Concepto_{concepto_elegido}_{id_puente}_{date.today()}.pdf",
                mime="application/pdf",
                use_container_width=True
            )

        # -----------------------------------------------------------
        # CUADRO CONTENEDOR DETALLADO DEL CONCEPTO SELECCIONADO
        # -----------------------------------------------------------
        st.markdown(f"""
        <div style="background-color: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 14px; padding: 22px; margin-top: 15px; margin-bottom: 24px; box-shadow: 0 2px 5px rgba(0,0,0,0.04);">
            <div style="display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #2563EB; padding-bottom: 10px; margin-bottom: 16px;">
                <h4 style="margin: 0; color: #0F172A; font-size: 1.15rem; font-weight: 700;">
                    📌 Concepto No. {concepto_elegido}
                </h4>
                <span style="background-color: #EFF6FF; color: #1D4ED8; font-weight: 600; padding: 4px 12px; border-radius: 6px; font-size: 0.85rem;">
                    Partida: {limpiar_texto(row_ind['Partida'])}
                </span>
            </div>
            <p style="color: #334155; font-size: 0.95rem; line-height: 1.5; margin-bottom: 20px;">
                <b>Descripción:</b> {row_ind['Descripción']}
            </p>
            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 16px;">
                <div style="background-color: #F8FAFC; border-left: 4px solid #64748B; border-radius: 8px; padding: 12px 16px;">
                    <span style="font-size: 0.75rem; color: #64748B; text-transform: uppercase; font-weight: 700; letter-spacing: 0.5px;">Unidad</span>
                    <div style="font-size: 1.25rem; font-weight: 700; color: #0F172A; margin-top: 4px;">{row_ind['Unidad']}</div>
                </div>
                <div style="background-color: #F8FAFC; border-left: 4px solid #2563EB; border-radius: 8px; padding: 12px 16px;">
                    <span style="font-size: 0.75rem; color: #64748B; text-transform: uppercase; font-weight: 700; letter-spacing: 0.5px;">Cantidad Total</span>
                    <div style="font-size: 1.25rem; font-weight: 700; color: #0F172A; margin-top: 4px;">{row_ind['Contratado']:,.2f}</div>
                </div>
                <div style="background-color: #F8FAFC; border-left: 4px solid #16A34A; border-radius: 8px; padding: 12px 16px;">
                    <span style="font-size: 0.75rem; color: #64748B; text-transform: uppercase; font-weight: 700; letter-spacing: 0.5px;">Cantidad Realizada</span>
                    <div style="font-size: 1.25rem; font-weight: 700; color: #15803D; margin-top: 4px;">{row_ind['Ejecutado']:,.2f}</div>
                </div>
                <div style="background-color: #F8FAFC; border-left: 4px solid #D97706; border-radius: 8px; padding: 12px 16px;">
                    <span style="font-size: 0.75rem; color: #64748B; text-transform: uppercase; font-weight: 700; letter-spacing: 0.5px;">Saldo por Realizar</span>
                    <div style="font-size: 1.25rem; font-weight: 700; color: #B45309; margin-top: 4px;">{row_ind['Saldo Available']:,.2f}</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        st.write("##### 📋 Historial de Capturas de Campo")
        if not df_hist_ind.empty:
            st.dataframe(df_hist_ind, use_container_width=True, height=220)
        else:
            st.info("Sin registros de volumen capturados en campo para este concepto.")
