import sqlite3
import pandas as pd
import streamlit as st
from datetime import date
import io
import re

# Librerías para generación de PDF
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# 1. Configuración de la página web
st.set_page_config(
    page_title="Control de Materiales y Estimaciones - Puentes",
    page_icon="🏗️",
    layout="wide"
)

DB_NAME = "control_obra.db"

def get_connection():
    return sqlite3.connect(DB_NAME, check_same_thread=False)

def a_flotante(val):
    """Convierte de forma segura textos con $, comas o letras a valores numéricos (float)."""
    if pd.isna(val) or val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).strip()
    cleaned = re.sub(r'[^\d.-]', '', val_str.replace(',', ''))
    try:
        return float(cleaned) if cleaned not in ['', '.', '-'] else 0.0
    except ValueError:
        return 0.0

def limpiar_clave(val):
    """Limpia la clave o número de concepto eliminando terminaciones .0 de Excel."""
    if pd.isna(val) or val is None:
        return ""
    val_str = str(val).strip()
    if val_str.endswith('.0'):
        val_str = val_str[:-2]
    return val_str

def inicializar_bd():
    """Inicializa la BD y reestructura la tabla catalogo si proviene de una versión previa."""
    conn = get_connection()
    cursor = conn.cursor()
    
    # 1. Migración / Creación de la tabla Catalogo
    cursor.execute("PRAGMA table_info(catalogo)")
    cols_info = cursor.fetchall()
    
    pk_cols = [row[1] for row in cols_info if row[5] > 0]
    
    if len(cols_info) > 0 and len(pk_cols) < 2:
        cursor.execute("ALTER TABLE catalogo RENAME TO catalogo_old")
        cursor.execute('''
            CREATE TABLE catalogo (
                id_puente TEXT NOT NULL DEFAULT 'PACIOTLA',
                id_concepto TEXT NOT NULL,
                partida TEXT,
                subpartida TEXT,
                norma TEXT,
                descripcion TEXT NOT NULL,
                unidad TEXT NOT NULL,
                cantidad_contratada REAL NOT NULL,
                precio_unitario REAL DEFAULT 0,
                PRIMARY KEY (id_puente, id_concepto)
            )
        ''')
        cursor.execute('''
            INSERT OR IGNORE INTO catalogo (id_puente, id_concepto, partida, subpartida, norma, descripcion, unidad, cantidad_contratada, precio_unitario)
            SELECT COALESCE(id_puente, 'PACIOTLA'), id_concepto, partida, subpartida, norma, descripcion, unidad, cantidad_contratada, COALESCE(precio_unitario, 0)
            FROM catalogo_old
        ''')
        cursor.execute("DROP TABLE catalogo_old")
    elif len(cols_info) == 0:
        cursor.execute('''
            CREATE TABLE catalogo (
                id_puente TEXT NOT NULL DEFAULT 'PACIOTLA',
                id_concepto TEXT NOT NULL,
                partida TEXT,
                subpartida TEXT,
                norma TEXT,
                descripcion TEXT NOT NULL,
                unidad TEXT NOT NULL,
                cantidad_contratada REAL NOT NULL,
                precio_unitario REAL DEFAULT 0,
                PRIMARY KEY (id_puente, id_concepto)
            )
        ''')

    # 2. Tabla Consumos
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS consumos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            id_puente TEXT NOT NULL DEFAULT 'PACIOTLA',
            fecha TEXT NOT NULL,
            id_concepto TEXT NOT NULL,
            ubicacion TEXT NOT NULL,
            cantidad REAL NOT NULL,
            remision TEXT,
            num_estimacion TEXT DEFAULT 'Ejecutado (No estimado)'
        )
    ''')
    
    cursor.execute("PRAGMA table_info(consumos)")
    cols_con = [row[1] for row in cursor.fetchall()]
    if 'id_puente' not in cols_con:
        cursor.execute("ALTER TABLE consumos ADD COLUMN id_puente TEXT DEFAULT 'PACIOTLA'")
    if 'num_estimacion' not in cols_con:
        cursor.execute("ALTER TABLE consumos ADD COLUMN num_estimacion TEXT DEFAULT 'Ejecutado (No estimado)'")

    # 3. Tabla Proyecto / Puentes
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS proyecto (
            id_puente TEXT PRIMARY KEY,
            nombre_contrato TEXT,
            num_contrato TEXT,
            contratista TEXT
        )
    ''')

    puentes_default = {
        "PACIOTLA": (
            'RECONSTRUCCION DEL PUENTE "PACIOTLA II", CAMINO: PAHUATLAN-PACIOTLA, EN EL ESTADO DE PUEBLA',
            '2026-21-CF-D-568-W-00-2026',
            'EDIFICADORA Y URBANIZADORA CRAWLER S.A DE C.V.'
        ),
        "PUENTE_EL_SALTO": (
            'RECONSTRUCCION DEL PUENTE "EL SALTO KM 42+200"',
            '2026-21-CF-D-569-W-00-2026',
            'EDIFICADORA Y URBANIZADORA CRAWLER S.A DE C.V.'
        )
    }

    for p_id, (nom, num, cont) in puentes_default.items():
        cursor.execute("SELECT 1 FROM proyecto WHERE id_puente = ?", (p_id,))
        if not cursor.fetchone():
            cursor.execute(
                "INSERT INTO proyecto (id_puente, nombre_contrato, num_contrato, contratista) VALUES (?, ?, ?, ?)",
                (p_id, nom, num, cont)
            )

    cursor.execute("UPDATE catalogo SET id_puente = 'PACIOTLA' WHERE id_puente IS NULL OR id_puente = '' OR id_puente = 'None'")
    cursor.execute("UPDATE consumos SET id_puente = 'PACIOTLA' WHERE id_puente IS NULL OR id_puente = '' OR id_puente = 'None'")

    conn.commit()
    conn.close()

inicializar_bd()

# -------------------------------------------------------------------
# FUNCIONES AUXILIARES DE BASE DE DATOS Y PROYECTOS
# -------------------------------------------------------------------
def get_lista_puentes():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id_puente, nombre_contrato FROM proyecto")
    rows = cursor.fetchall()
    conn.close()
    return {row[0]: row[1] for row in rows}

def get_proyecto_info(id_puente):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT nombre_contrato, num_contrato, contratista FROM proyecto WHERE id_puente = ?", (id_puente,))
    row = cursor.fetchone()
    conn.close()
    if row:
        return {"nombre_contrato": str(row[0]), "num_contrato": str(row[1]), "contratista": str(row[2])}
    return {
        "nombre_contrato": f"OBRA / PUENTE {id_puente}",
        "num_contrato": "SIN ASIGNAR",
        "contratista": "EMPRESA CONTRATISTA"
    }

def update_proyecto_info(id_puente, nombre, num, contratista):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT 1 FROM proyecto WHERE id_puente = ?", (id_puente,))
    if cursor.fetchone():
        cursor.execute('''
            UPDATE proyecto SET nombre_contrato = ?, num_contrato = ?, contratista = ?
            WHERE id_puente = ?
        ''', (nombre, num, contratista, id_puente))
    else:
        cursor.execute('''
            INSERT INTO proyecto (id_puente, nombre_contrato, num_contrato, contratista)
            VALUES (?, ?, ?, ?)
        ''', (id_puente, nombre, num, contratista))
    conn.commit()
    conn.close()

def crear_nuevo_puente(id_puente, nombre, num, contratista):
    update_proyecto_info(id_puente.upper().replace(" ", "_"), nombre, num, contratista)

def limpiar_texto(val):
    if pd.isna(val) or val is None or str(val).strip().lower() in ["", "nan", "none", "general", "null"]:
        return "-"
    return str(val).strip()

def mostrar_mensaje_exito(mensaje):
    st.markdown(f"""
        <div style="background-color: #d1e7dd; color: #0f5132; padding: 16px; border-radius: 8px; margin-bottom: 20px; border: 1px solid #badbcc;">
            <h4 style="margin:0; padding:0;">✅ ¡Operación Exitosa!</h4>
            <p style="margin:5px 0 0 0; font-size: 16px;">{mensaje}</p>
        </div>
    """, unsafe_allow_html=True)

def desplegar_mensaje_flash():
    """Despliega y borra mensajes guardados en el session state."""
    if 'msg_exito' in st.session_state:
        mostrar_mensaje_exito(st.session_state['msg_exito'])
        del st.session_state['msg_exito']

# -------------------------------------------------------------------
# GENERACIÓN DE EXCEL Y PDF
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
        LEFT JOIN consumos r ON c.id_concepto = r.id_concepto AND c.id_puente = r.id_puente
        WHERE c.id_puente = ?
        GROUP BY c.id_concepto
        ORDER BY CAST(c.id_concepto AS INTEGER) ASC, c.id_concepto ASC
    '''
    df_general = pd.read_sql_query(query_general, conn, params=(id_puente,))

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
        JOIN catalogo c ON r.id_concepto = c.id_concepto AND r.id_puente = c.id_puente
        WHERE r.id_puente = ?
        ORDER BY r.id DESC
    '''
    df_consumos = pd.read_sql_query(query_consumos, conn, params=(id_puente,))
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
    pct_estimado = (monto_estimado / monto_contrato * 100) if monto_contrato > 0 else 0

    data_header = [
        [
            Paragraph(f"<b>Obra / Puente:</b> {info_proj['nombre_contrato']}", meta_style),
            Paragraph(f"<b>Nº Contrato:</b> {info_proj['num_contrato']}", meta_style),
            Paragraph(f"<b>Monto Contrato:</b> ${monto_contrato:,.2f}", meta_style)
        ],
        [
            Paragraph(f"<b>Empresa:</b> {info_proj['contratista']}", meta_style),
            Paragraph(f"<b>Fecha Emisión:</b> {date.today().strftime('%d/%m/%Y')}", meta_style),
            Paragraph(f"<b>% Avance Financiero Global:</b> {pct_ejecutado:.2f}% (Est: {pct_estimado:.2f}%)", meta_style)
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
    """Genera un reporte PDF detallado únicamente para un concepto específico."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    elements = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=13, leading=15, textColor=colors.HexColor('#1E3A8A'))
    cell_style = ParagraphStyle('CellStyle', parent=styles['Normal'], fontSize=7.5, leading=9.5)
    cell_style_bold = ParagraphStyle('CellStyleBold', parent=styles['Normal'], fontSize=8, leading=10, fontName='Helvetica-Bold')
    meta_style = ParagraphStyle('MetaStyle', parent=styles['Normal'], fontSize=8, leading=10)

    # 1. Título General
    elements.append(Paragraph("<b>REPORTE DETALLADO DE CONCEPTO</b>", title_style))
    elements.append(Spacer(1, 6))

    # 2. Encabezado de la Obra
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

    # 3. Datos del Concepto
    no_item = info_concepto['id_concepto']
    desc = info_concepto['descripcion']
    partida = limpiar_texto(info_concepto.get('partida', '-'))
    unidad = info_concepto['unidad']
    cant_contratada = info_concepto['cantidad_contratada']
    pu = info_concepto.get('precio_unitario', 0.0)
    imp_contratado = cant_contratada * pu
    
    cant_ejecutada = info_concepto.get('total_ejecutado', 0.0)
    cant_estimada = info_concepto.get('total_estimado', 0.0)
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

    # 4. Tabla de Registros de Campo
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
            data_hist.append([
                str(r["ID Registro"]),
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
# SISTEMA PRINCIPAL Y NAVEGACIÓN
# -------------------------------------------------------------------

if 'puente_seleccionado' not in st.session_state:
    st.session_state['puente_seleccionado'] = None

lista_puentes = get_lista_puentes()

# --- PANTALLA 1: SELECCIÓN / GESTIÓN DE PUENTES ---
if st.session_state['puente_seleccionado'] is None:
    st.title("🌉 Control de Materiales y Estimaciones de Obra")
    desplegar_mensaje_flash()
    
    st.subheader("Selecciona una obra / puente para administrar sus volúmenes y estimaciones:")

    cols = st.columns(2)
    idx = 0
    for p_id, p_nombre in lista_puentes.items():
        with cols[idx % 2]:
            with st.container(border=True):
                st.markdown(f"### **{p_nombre}**")
                info_p = get_proyecto_info(p_id)
                st.caption(f"**Contrato:** {info_p['num_contrato']} | **Contratista:** {info_p['contratista']}")
                
                if st.button(f"🚀 Entrar a Administrar", key=f"btn_{p_id}", use_container_width=True):
                    st.session_state['puente_seleccionado'] = p_id
                    st.rerun()
        idx += 1

    st.divider()
    with st.expander("➕ Registrar Nueva Obra / Puente"):
        with st.form("form_nuevo_puente"):
            c1, c2 = st.columns(2)
            n_clave = c1.text_input("Clave Corta / ID del Puente (Ej: PUENTE_RIO_SECO):")
            n_nombre = c2.text_input("Nombre Completo de la Obra / Puente:")
            
            c3, c4 = st.columns(2)
            n_num = c3.text_input("Número de Contrato:")
            n_cont = c4.text_input("Empresa Contratista:")
            
            if st.form_submit_button("Guardar y Registrar Puente"):
                if n_clave and n_nombre:
                    crear_nuevo_puente(n_clave, n_nombre, n_num, n_cont)
                    st.session_state['msg_exito'] = f"Obra '{n_nombre}' registrada exitosamente."
                    st.rerun()
                else:
                    st.error("Por favor completa al menos la clave y el nombre de la obra.")

# --- PANTALLA 2: ENTORNO DE TRABAJO DEL PUENTE SELECCIONADO ---
else:
    id_puente = st.session_state['puente_seleccionado']
    info_proyecto = get_proyecto_info(id_puente)

    with st.sidebar:
        st.markdown(f"### 🌉 **{info_proyecto['nombre_contrato'][:30]}...**")
        if st.button("⬅️ Volver a Lista de Puentes", use_container_width=True):
            st.session_state['puente_seleccionado'] = None
            st.rerun()
        
        with st.expander("⚙️ Editar Datos de Contrato / Empresa"):
            with st.form("form_editar_info_contrato"):
                e_nombre = st.text_area("Nombre de la Obra:", value=info_proyecto['nombre_contrato'], height=70)
                e_num = st.text_input("Número de Contrato:", value=info_proyecto['num_contrato'])
                e_cont = st.text_input("Empresa Contratista:", value=info_proyecto['contratista'])
                
                if st.form_submit_button("💾 Guardar Datos del Contrato"):
                    update_proyecto_info(id_puente, e_nombre, e_num, e_cont)
                    st.session_state['msg_exito'] = "Datos de contrato actualizados correctamente."
                    st.rerun()

        st.divider()
        opcion = st.radio("Menú de Trabajo:", ["Dashboard de Volúmenes", "Registrar Consumo / Estimación", "Gestión de Catálogo"])

    # -------------------------------------------------------------------
    # OPCIÓN 1: DASHBOARD DE CONTROL Y MÉTRICAS
    # -------------------------------------------------------------------
    if opcion == "Dashboard de Volúmenes":
        st.title(info_proyecto['nombre_contrato'])
        desplegar_mensaje_flash()
        
        with st.container(border=True):
            c_a, c_b, c_c = st.columns([2, 1, 1])
            c_a.markdown(f"**Obra / Proyecto:**\n{info_proyecto['nombre_contrato']}")
            c_b.markdown(f"**Nº Contrato:**\n{info_proyecto['num_contrato']}")
            c_c.markdown(f"**Contratista:**\n{info_proyecto['contratista']}")

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
            LEFT JOIN consumos r ON c.id_concepto = r.id_concepto AND c.id_puente = r.id_puente
            WHERE c.id_puente = ?
            GROUP BY c.id_concepto
            ORDER BY CAST(c.id_concepto AS INTEGER) ASC, c.id_concepto ASC
        '''
        df = pd.read_sql_query(query, conn, params=(id_puente,))
        conn.close()

        if not df.empty:
            monto_contratado = (df["Contratado"] * df["P.U."]).sum()
            monto_ejecutado = (df["Ejecutado"] * df["P.U."]).sum()
            monto_estimado = (df["Estimado"] * df["P.U."]).sum()
            pct_financiero_ejecutado = (monto_ejecutado / monto_contratado * 100) if monto_contratado > 0 else 0
            pct_financiero_estimado = (monto_estimado / monto_contratado * 100) if monto_contratado > 0 else 0

            st.subheader("📊 Métricas Financieras y Físicas Globales")
            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("Monto Contratado", f"${monto_contratado:,.2f}")
            m2.metric("Monto Ejecutado", f"${monto_ejecutado:,.2f}")
            m3.metric("% Avance Financiero", f"{pct_financiero_ejecutado:.2f}%")
            m4.metric("Monto Estimado", f"${monto_estimado:,.2f}")
            m5.metric("% Estimado Total", f"{pct_financiero_estimado:.2f}%")

            st.divider()

            c_head, c_btn_pdf, c_btn_excel = st.columns([2, 1, 1])
            with c_head:
                st.subheader("Estado General de Conceptos")
            with c_btn_pdf:
                pdf_gen = generar_pdf_general(df, info_proyecto)
                st.download_button("📄 Descargar PDF General", data=pdf_gen, file_name=f"Reporte_General_{id_puente}_{date.today()}.pdf", mime="application/pdf", use_container_width=True)
            with c_btn_excel:
                excel_bd = generar_excel_bd(id_puente)
                st.download_button("📊 Descargar Excel", data=excel_bd, file_name=f"BaseDatos_{id_puente}_{date.today()}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

            st.dataframe(
                df.style.format({
                    "Contratado": "{:,.2f}",
                    "P.U.": "${:,.2f}",
                    "Ejecutado": "{:,.2f}",
                    "Estimado": "{:,.2f}",
                    "Saldo Available": "{:,.2f}",
                    "% Avance": "{:.2f}%"
                }),
                width="stretch",
                height=350
            )

            st.divider()
            with st.expander("🔍 Consultar y Exportar PDF por Concepto Individual"):
                lista_conceptos = df["No."].tolist()
                c_indiv_sel = st.selectbox("Selecciona la clave del concepto:", lista_conceptos, key="sb_indiv_dash")
                
                if c_indiv_sel:
                    row_ind = df[df["No."] == c_indiv_sel].iloc[0]
                    
                    conn = get_connection()
                    query_hist_ind = """
                        SELECT r.id AS "ID Registro", r.fecha AS "Fecha", r.ubicacion AS "Ubicación", 
                               r.cantidad AS "Cantidad", r.remision AS "Remisión", r.num_estimacion AS "Estimación"
                        FROM consumos r
                        WHERE r.id_puente = ? AND r.id_concepto = ?
                        ORDER BY r.id ASC
                    """
                    df_hist_ind = pd.read_sql_query(query_hist_ind, conn, params=(id_puente, c_indiv_sel))
                    conn.close()

                    info_c_dict = {
                        'id_concepto': c_indiv_sel,
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

                    st.download_button(
                        f"📄 Descargar PDF del Concepto No. {c_indiv_sel}",
                        data=pdf_indiv_bytes,
                        file_name=f"Reporte_Concepto_{c_indiv_sel}_{id_puente}_{date.today()}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
        else:
            st.info("El catálogo de este puente está vacío. Importa el catálogo en la sección 'Gestión de Catálogo'.")

    # -------------------------------------------------------------------
    # OPCIÓN 2: REGISTRO DE CONSUMO / ESTIMACIÓN
    # -------------------------------------------------------------------
    elif opcion == "Registrar Consumo / Estimación":
        st.title("Gestión de Avances y Estimaciones")
        desplegar_mensaje_flash()

        tab_nuevo, tab_editar = st.tabs(["✍️ Capturar Nuevo Avance", "✏️ Editar / Eliminar Registros"])

        with tab_nuevo:
            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute("""
                SELECT 
                    c.id_concepto, c.partida, c.subpartida, c.norma, c.descripcion, c.unidad, c.cantidad_contratada, c.precio_unitario,
                    COALESCE(SUM(r.cantidad), 0) AS acumulado_ejecutado,
                    COALESCE(SUM(CASE WHEN r.num_estimacion != 'Ejecutado (No estimado)' THEN r.cantidad ELSE 0 END), 0) AS acumulado_estimado
                FROM catalogo c
                LEFT JOIN consumos r ON c.id_concepto = r.id_concepto AND c.id_puente = r.id_puente
                WHERE c.id_puente = ?
                GROUP BY c.id_concepto
                ORDER BY CAST(c.id_concepto AS INTEGER) ASC, c.id_concepto ASC
            """, (id_puente,))
            conceptos = cursor.fetchall()
            conn.close()

            if not conceptos:
                st.error("No hay conceptos cargados para este puente. Carga un catálogo primero.")
            else:
                dict_conceptos = {str(c[0]): c for c in conceptos}
                st.subheader("1. Selecciona el Concepto")
                numero_sel = st.selectbox("Número / Clave de Concepto:", list(dict_conceptos.keys()))

                datos_sel = dict_conceptos[numero_sel]
                no_item, partida, subpartida, norma, descripcion, unidad, contratado, pu, acumulado, acumulado_est = datos_sel
                saldo = contratado - acumulado

                with st.container(border=True):
                    st.markdown(f"### **Concepto No. {no_item}** - {limpiar_texto(partida)}")
                    st.info(f"**Descripción:** {descripcion}")
                    
                    c1, c2, c3 = st.columns(3)
                    c1.metric("Volumen Contratado", f"{contratado:,.2f} {unidad}")
                    c2.metric("Acumulado Ejecutado", f"{acumulado:,.2f} {unidad}")
                    c3.metric("Saldo Disponible", f"{saldo:,.2f} {unidad}")

                    # Botón PDF Individual
                    conn = get_connection()
                    query_hist_concepto = """
                        SELECT r.id AS "ID Registro", r.fecha AS "Fecha", r.ubicacion AS "Ubicación", 
                               r.cantidad AS "Cantidad", r.remision AS "Remisión", r.num_estimacion AS "Estimación"
                        FROM consumos r
                        WHERE r.id_puente = ? AND r.id_concepto = ?
                        ORDER BY r.id ASC
                    """
                    df_hist_concepto = pd.read_sql_query(query_hist_concepto, conn, params=(id_puente, numero_sel))
                    conn.close()

                    info_concepto_sel = {
                        'id_concepto': no_item,
                        'partida': partida,
                        'subpartida': subpartida,
                        'norma': norma,
                        'descripcion': descripcion,
                        'unidad': unidad,
                        'cantidad_contratada': contratado,
                        'precio_unitario': pu,
                        'total_ejecutado': acumulado,
                        'total_estimado': acumulado_est
                    }

                    pdf_concepto_bytes = generar_pdf_concepto_individual(info_concepto_sel, df_hist_concepto, info_proyecto)

                    st.download_button(
                        label=f"📄 Descargar PDF Detallado (Concepto No. {no_item})",
                        data=pdf_concepto_bytes,
                        file_name=f"Reporte_Concepto_{no_item}_{id_puente}_{date.today()}.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )

                st.subheader("2. Captura de Registro")
                with st.form("form_nuevo_registro"):
                    c_f1, c_f2 = st.columns(2)
                    fecha_in = c_f1.date_input("Fecha de Trabajo:", date.today())
                    cant_in = c_f2.number_input(f"Cantidad a Registrar ({unidad}):", min_value=0.01, step=1.0)

                    c_f3, c_f4 = st.columns(2)
                    ubicacion_in = c_f3.text_input("Ubicación / Elemento:", placeholder="Ej. Zapata Estribo 1")
                    remision_in = c_f4.text_input("Folio Remisión / Ticket:", placeholder="Ej. REM-1024")

                    st.markdown("**Asignación de Estimación:**")
                    num_est_in = st.text_input("Número / Folio de Estimación:", value="Estimación 1")
                    no_estimado_chk = st.checkbox("Marcar como solo 'Ejecutado' (Pendiente de estimar)", value=False)

                    if st.form_submit_button("Guardar Registro de Campo"):
                        est_final = "Ejecutado (No estimado)" if no_estimado_chk else (num_est_in.strip() if num_est_in.strip() else "Estimación 1")
                        
                        conn = get_connection()
                        cursor = conn.cursor()
                        cursor.execute(
                            "INSERT INTO consumos (id_puente, fecha, id_concepto, ubicacion, cantidad, remision, num_estimacion) VALUES (?, ?, ?, ?, ?, ?, ?)",
                            (id_puente, str(fecha_in), str(no_item), ubicacion_in, cant_in, remision_in, est_final)
                        )
                        conn.commit()
                        conn.close()

                        # Se almacena en la memoria y se fuerza la recarga
                        st.session_state['msg_exito'] = f"Se registraron correctamente {cant_in:,.2f} {unidad} al Concepto No. {no_item} bajo el estatus: '{est_final}'."
                        st.rerun()

        with tab_editar:
            st.subheader("Historial de Capturas de este Puente")
            conn = get_connection()
            query_hist = """
                SELECT 
                    r.id AS "ID Registro",
                    r.fecha AS "Fecha",
                    r.id_concepto AS "No. Concepto",
                    c.descripcion AS "Descripción",
                    r.ubicacion AS "Ubicación",
                    r.cantidad AS "Cantidad",
                    c.unidad AS "Unidad",
                    r.num_estimacion AS "Estimación",
                    r.remision AS "Remisión"
                FROM consumos r
                JOIN catalogo c ON r.id_concepto = c.id_concepto AND r.id_puente = c.id_puente
                WHERE r.id_puente = ?
                ORDER BY r.id DESC
            """
            df_hist = pd.read_sql_query(query_hist, conn, params=(id_puente,))
            conn.close()

            if df_hist.empty:
                st.info("Aún no existen capturas de campo grabadas.")
            else:
                st.dataframe(df_hist, width="stretch", height=250)
                st.divider()

                id_sel = st.selectbox("Selecciona ID de registro a editar o eliminar:", df_hist["ID Registro"].tolist())
                reg_sel = df_hist[df_hist["ID Registro"] == id_sel].iloc[0]

                with st.form("form_edit_reg"):
                    c_e1, c_e2 = st.columns(2)
                    nueva_f = c_e1.date_input("Fecha:", pd.to_datetime(reg_sel["Fecha"]).date())
                    nueva_c = c_e2.number_input(f"Cantidad ({reg_sel['Unidad']}):", min_value=0.01, value=float(reg_sel["Cantidad"]))

                    c_e3, c_e4 = st.columns(2)
                    nueva_u = c_e3.text_input("Ubicación:", value=str(reg_sel["Ubicación"]))
                    nueva_r = c_e4.text_input("Remisión:", value=str(reg_sel["Remisión"] if pd.notna(reg_sel["Remisión"]) else ""))

                    nueva_est = st.text_input("Número / Folio de Estimación:", value=str(reg_sel["Estimación"]))

                    b_act, b_eli = st.columns(2)
                    if b_act.form_submit_button("💾 Guardar Cambios"):
                        conn = get_connection()
                        cursor = conn.cursor()
                        cursor.execute("""
                            UPDATE consumos 
                            SET fecha = ?, ubicacion = ?, cantidad = ?, remision = ?, num_estimacion = ?
                            WHERE id = ?
                        """, (str(nueva_f), nueva_u, nueva_c, nueva_r, nueva_est, id_sel))
                        conn.commit()
                        conn.close()
                        st.session_state['msg_exito'] = f"Registro ID {id_sel} actualizado correctamente."
                        st.rerun()

                    if b_eli.form_submit_button("🗑️ Eliminar Registro"):
                        conn = get_connection()
                        cursor = conn.cursor()
                        cursor.execute("DELETE FROM consumos WHERE id = ?", (id_sel,))
                        conn.commit()
                        conn.close()
                        st.session_state['msg_exito'] = f"Registro ID {id_sel} eliminado correctamente."
                        st.rerun()

    # -------------------------------------------------------------------
    # OPCIÓN 3: GESTIÓN DE CATÁLOGO E IMPORTACIÓN
    # -------------------------------------------------------------------
    elif opcion == "Gestión de Catálogo":
        st.title(f"Catálogo de Conceptos - {info_proyecto['nombre_contrato']}")
        desplegar_mensaje_flash()

        st.subheader("📂 Cargar Presupuesto desde Excel")
        uploaded_file = st.file_uploader("Selecciona tu archivo (.xlsx o .csv)", type=["xlsx", "csv"])

        if uploaded_file is not None:
            try:
                dict_hojas = pd.read_excel(uploaded_file, sheet_name=None, header=None)
                hoja_sel = st.selectbox("Selecciona la hoja que contiene el catálogo:", list(dict_hojas.keys()))
                
                df_raw = dict_hojas[hoja_sel]

                # Detectar la fila de encabezados
                idx_header = 0
                for i, row in df_raw.iterrows():
                    row_str = " ".join([str(v).upper() for v in row.values if pd.notna(v)])
                    if ("CONCEPTO" in row_str or "DESCRIPCION" in row_str or "CLAVE" in row_str or "NO." in row_str or "UNIDAD" in row_str) and ("CANTIDAD" in row_str or "P.U" in row_str or "PRECIO" in row_str):
                        idx_header = i
                        break

                df_proc = pd.read_excel(uploaded_file, sheet_name=hoja_sel, skiprows=idx_header)
                df_proc.columns = [str(c).strip().upper() for c in df_proc.columns]

                st.markdown("---")
                st.subheader("🛠️ Mapeo Inteligente de Columnas")
                st.write("Verifica o ajusta la equivalencia entre tu Excel y la base de datos:")

                cols_excel = list(df_proc.columns)

                def buscar_col(patrones, excluir=None):
                    for p in patrones:
                        for c in cols_excel:
                            c_up = c.upper()
                            if excluir and any(e in c_up for e in excluir):
                                continue
                            if p in c_up:
                                return cols_excel.index(c)
                    return 0

                c_m1, c_m2 = st.columns(2)
                col_no_sel = c_m1.selectbox("Clave / No. Concepto (*):", cols_excel, index=buscar_col(["CLAVE", "NO", "ITEM", "NUM"]))
                col_desc_sel = c_m2.selectbox("Descripción (*):", cols_excel, index=buscar_col(["DESC", "CONCEPTO", "ESPEC"]))

                c_m3, c_m4 = st.columns(2)
                col_uni_sel = c_m3.selectbox("Unidad (*):", cols_excel, index=buscar_col(["UNIDAD", "UND", "MEDIDA"]))
                col_cant_sel = c_m4.selectbox("Cantidad Contratada (*):", cols_excel, index=buscar_col(["CANT", "VOLUMEN", "CANTIDAD"]))

                c_m5, c_m6 = st.columns(2)
                idx_pu = buscar_col(["P.U", "PRECIO UNITARIO", "UNITARIO", "PU", "PRECIO"], excluir=["LETRA", "TEXTO"])
                col_pu_sel = c_m5.selectbox("Precio Unitario (P.U.):", cols_excel, index=idx_pu)
                col_norma_sel = c_m6.selectbox("Norma / Especificación (Opcional):", ["-- Omitir --"] + cols_excel, index=0)

                c_m7, c_m8 = st.columns(2)
                col_partida_sel = c_m7.selectbox("Partida (Opcional):", ["-- Omitir --"] + cols_excel, index=0)
                col_sub_sel = c_m8.selectbox("Subpartida (Opcional):", ["-- Omitir --"] + cols_excel, index=0)

                st.markdown("---")
                st.warning("⚠️ **Modo de Importación:**")
                opcion_reemplazar = st.checkbox("Reemplazar catálogo actual completamente (Borra los conceptos anteriores de esta obra antes de cargar los nuevos)", value=True)

                if st.button("🚀 Procesar e Importar Catálogo", use_container_width=True):
                    conn = get_connection()
                    cursor = conn.cursor()

                    if opcion_reemplazar:
                        cursor.execute("DELETE FROM catalogo WHERE id_puente = ?", (id_puente,))

                    filas_insertadas = 0

                    for _, row in df_proc.iterrows():
                        no_item = limpiar_clave(row[col_no_sel])
                        desc = str(row[col_desc_sel]).strip() if pd.notna(row[col_desc_sel]) else ""
                        unidad = str(row[col_uni_sel]).strip() if pd.notna(row[col_uni_sel]) else ""

                        if not unidad or unidad.lower() == "nan" or not no_item or no_item.lower() == "nan":
                            continue

                        cant = a_flotante(row[col_cant_sel])
                        pu = a_flotante(row[col_pu_sel])
                        
                        norma_val = str(row[col_norma_sel]).strip() if col_norma_sel != "-- Omitir --" and pd.notna(row[col_norma_sel]) else "-"
                        partida_val = str(row[col_partida_sel]).strip() if col_partida_sel != "-- Omitir --" and pd.notna(row[col_partida_sel]) else "-"
                        subpartida_val = str(row[col_sub_sel]).strip() if col_sub_sel != "-- Omitir --" and pd.notna(row[col_sub_sel]) else "-"

                        cursor.execute('''
                            INSERT OR REPLACE INTO catalogo 
                            (id_puente, id_concepto, partida, subpartida, norma, descripcion, unidad, cantidad_contratada, precio_unitario)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (id_puente, no_item, partida_val, subpartida_val, norma_val, desc, unidad, cant, pu))
                        
                        filas_insertadas += 1

                    conn.commit()
                    conn.close()

                    st.session_state['msg_exito'] = f"¡Se importaron correctamente {filas_insertadas} conceptos al catálogo!"
                    st.rerun()

            except Exception as e:
                st.error(f"Error procesando el archivo: {e}")