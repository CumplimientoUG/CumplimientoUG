import streamlit as st
import sqlite3
import os
from datetime import datetime, date
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import pandas as pd
from pathlib import Path

# Configuración de la app
st.set_page_config(
    page_title="Cumplimiento UG - Gestión de Cartera",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Rutas
BASE_DIR = Path(__file__).parent
DB_PATH = BASE_DIR / "cartera.db"
CLIENTES_DIR = BASE_DIR / "clientes"
ASSETS_DIR = BASE_DIR / "assets"
LOGO_PATH = ASSETS_DIR / "logo.png"
FIRMA_PATH = ASSETS_DIR / "firma.png"

CLIENTES_DIR.mkdir(exist_ok=True)
ASSETS_DIR.mkdir(exist_ok=True)

# Inicializar base de datos
def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS clientes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL,
            nit TEXT,
            direccion TEXT,
            telefono TEXT,
            email TEXT,
            valor_mensual REAL DEFAULT 0,
            activo INTEGER DEFAULT 1,
            notas TEXT,
            fecha_creacion TEXT
        )
    ''')
    c.execute('''
        CREATE TABLE IF NOT EXISTS cuentas_cobro (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cliente_id INTEGER,
            numero TEXT,
            fecha TEXT,
            periodo TEXT,
            valor REAL,
            estado TEXT DEFAULT 'Pendiente',
            ruta_archivo TEXT,
            notas TEXT,
            FOREIGN KEY (cliente_id) REFERENCES clientes (id)
        )
    ''')
    conn.commit()
    conn.close()

init_db()

def get_conn():
    return sqlite3.connect(DB_PATH)

# Funciones de base de datos
def get_clientes(activos_only=True):
    conn = get_conn()
    query = "SELECT * FROM clientes"
    if activos_only:
        query += " WHERE activo = 1"
    query += " ORDER BY nombre"
    df = pd.read_sql_query(query, conn)
    conn.close()
    return df

def agregar_cliente(nombre, nit, direccion, telefono, email, valor, notas):
    conn = get_conn()
    c = conn.cursor()
    c.execute('''
        INSERT INTO clientes (nombre, nit, direccion, telefono, email, valor_mensual, notas, fecha_creacion)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    ''', (nombre, nit, direccion, telefono, email, valor, notas, datetime.now().isoformat()))
    conn.commit()
    # Crear carpeta del cliente
    cliente_dir = CLIENTES_DIR / nombre.replace("/", "_").replace("\\", "_")
    cliente_dir.mkdir(exist_ok=True)
    (cliente_dir / "cuentas").mkdir(exist_ok=True)
    conn.close()
    return cliente_dir

def generar_cuenta_cobro(cliente, numero, periodo, valor, fecha=None):
    if fecha is None:
        fecha = date.today()
    
    doc = Document()
    
    # Configurar márgenes
    sections = doc.sections
    for section in sections:
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin = Cm(1.5)
        section.right_margin = Cm(1.5)
    
    # Logo
    if LOGO_PATH.exists():
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(str(LOGO_PATH), width=Inches(2.2))
    
    # Encabezado
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("CUMPLIMIENTO_UG")
    run.bold = True
    run.font.size = Pt(16)
    run.font.color.rgb = RGBColor(31, 78, 121)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("JAIME RAFAEL ROMERO ACOSTA")
    run.bold = True
    run.font.size = Pt(12)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("NIT 98658918-7  ·  C.C. 98658918 de Envigado")
    run.font.size = Pt(10)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("GESTIÓN DE SISTEMAS DE CUMPLIMIENTO NORMATIVO")
    run.bold = True
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(46, 117, 182)
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("SARLAFT · PTEE · SAGRILAFT · RMM · SGSI · SDC · RNBD · GBG")
    run.italic = True
    run.font.size = Pt(9)
    
    doc.add_paragraph()
    
    # Título
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"CUENTA DE COBRO No. {numero}")
    run.bold = True
    run.font.size = Pt(16)
    run.font.color.rgb = RGBColor(31, 78, 121)
    
    doc.add_paragraph()
    
    # Cliente
    doc.add_paragraph("SEÑOR(ES):").runs[0].bold = True
    p = doc.add_paragraph()
    p.add_run(cliente['nombre']).bold = True
    if cliente['direccion']:
        doc.add_paragraph(cliente['direccion'])
    if cliente['nit']:
        doc.add_paragraph(f"NIT. {cliente['nit']}")
    
    doc.add_paragraph()
    doc.add_paragraph("DEBE A:").runs[0].bold = True
    p = doc.add_paragraph()
    p.add_run("JAIME RAFAEL ROMERO ACOSTA").bold = True
    doc.add_paragraph("C.C. 98658918 de Envigado - Antioquia")
    
    doc.add_paragraph()
    doc.add_paragraph("LA SUMA DE:").runs[0].bold = True
    # Valor en letras simplificado (puedes mejorar con num2words si se instala)
    doc.add_paragraph(f"La suma de {valor:,.0f} pesos colombianos".replace(",", "."))
    p = doc.add_paragraph()
    run = p.add_run(f"${valor:,.0f} COP".replace(",", "."))
    run.bold = True
    run.font.size = Pt(14)
    run.font.color.rgb = RGBColor(31, 78, 121)
    
    doc.add_paragraph()
    doc.add_paragraph(f"Por los conceptos asociados a CIIU N° 7020 Actividades de Consultoría de gestión, correspondientes al periodo {periodo}, así:")
    
    doc.add_paragraph("1. Relativo a las funciones como Oficial de Cumplimiento (SARLAFT y PTEE)").runs[0].bold = True
    for item in [
        "Representación ante la Superintendencia correspondiente.",
        "Monitoreo de relaciones vinculantes con las contrapartes.",
        "Consultas y Debidas Diligencias a las contrapartes en condición de ACTIVAS.",
        "Actualización de matrices SARLAFT y PTEE según normativa vigente."
    ]:
        doc.add_paragraph(f"• {item}", style='List Bullet')
    
    doc.add_paragraph()
    doc.add_paragraph("2. Relativo a las funciones de Oficial de Datos / SGSI (cuando aplique)").runs[0].bold = True
    for item in [
        "Representación ante la Superintendencia de Industria y Comercio.",
        "Gestión del Registro Nacional de Bases de Datos (RNBD).",
        "Acompañamiento en el Sistema de Gestión de Seguridad de la Información (SGSI).",
        "Acompañamiento a la responsabilidad de tratamiento de datos personales."
    ]:
        doc.add_paragraph(f"• {item}", style='List Bullet')
    
    doc.add_paragraph()
    p = doc.add_paragraph()
    p.add_run("Bajo gravedad de juramento me permito certificar que: ").bold = True
    p.add_run("Sobre los ingresos percibidos por rentas de trabajo NO tomaré costos o deducciones asociados a dichas rentas, objeto de mi contrato en la declaración de renta correspondiente al año gravable vigente, conforme a lo establecido en los artículos 1.2.4.1.6 y 1.2.4.1.17 del Decreto Único Reglamentario 1625 de 2016. De acuerdo con lo anterior, por favor aplicar la tarifa de retención establecida en el artículo 383 del Estatuto Tributario Nacional.")
    
    doc.add_paragraph()
    doc.add_paragraph("Anexos: Soporte de pago de seguridad social del mes correspondiente.").runs[0].bold = True
    
    doc.add_paragraph()
    meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio", 
             "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
    doc.add_paragraph(f"La presente se expide a los {fecha.day} días del mes de {meses[fecha.month-1]} del año {fecha.year}.")
    
    doc.add_paragraph()
    doc.add_paragraph("Cordialmente,")
    doc.add_paragraph()
    doc.add_paragraph()
    
    # Firma
    if FIRMA_PATH.exists():
        p = doc.add_paragraph()
        run = p.add_run()
        run.add_picture(str(FIRMA_PATH), width=Inches(2.0))
    
    p = doc.add_paragraph()
    p.add_run("Jaime Rafael Romero Acosta").bold = True
    doc.add_paragraph("C.C. 98658918")
    
    doc.add_paragraph()
    doc.add_paragraph("Teléfono: 322 632 9692")
    doc.add_paragraph("Dirección: Carrera 84 # 44-93 piso 1, Medellín – Antioquia")
    p = doc.add_paragraph()
    p.add_run("Cuenta bancaria: Ahorros Banco de Bogotá N° 510 301 302").bold = True
    
    # Guardar en carpeta del cliente
    cliente_folder = CLIENTES_DIR / cliente['nombre'].replace("/", "_").replace("\\", "_") / "cuentas"
    cliente_folder.mkdir(parents=True, exist_ok=True)
    filename = f"Cuenta_de_Cobro_{numero}_{cliente['nombre'][:30].replace(' ', '_')}.docx"
    filepath = cliente_folder / filename
    doc.save(str(filepath))
    
    return str(filepath)

# Interfaz
st.title("📋 Cumplimiento UG")
st.subheader("Gestión de Cartera y Cuentas de Cobro")
st.caption("Jaime Rafael Romero Acosta · Persona Natural · NIT 98658918-7 · CIIU 7020")

tab1, tab2, tab3, tab4 = st.tabs(["🏠 Dashboard", "👥 Clientes", "📄 Generar Cuentas", "📊 Históricos"])

with tab1:
    st.header("Dashboard")
    clientes = get_clientes()
    conn = get_conn()
    cuentas = pd.read_sql_query("SELECT * FROM cuentas_cobro", conn)
    conn.close()
    
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Clientes activos", len(clientes))
    col2.metric("Cuentas generadas", len(cuentas))
    pendientes = len(cuentas[cuentas['estado'] == 'Pendiente']) if not cuentas.empty else 0
    col3.metric("Pendientes de pago", pendientes)
    total = cuentas['valor'].sum() if not cuentas.empty else 0
    col4.metric("Total facturado", f"${total:,.0f}".replace(",", "."))
    
    if not clientes.empty:
        st.subheader("Clientes activos")
        st.dataframe(clientes[['nombre', 'nit', 'valor_mensual', 'telefono']], use_container_width=True)

with tab2:
    st.header("Gestión de Clientes")
    
    with st.form("nuevo_cliente"):
        st.subheader("Agregar nuevo cliente")
        col1, col2 = st.columns(2)
        nombre = col1.text_input("Nombre / Razón social *")
        nit = col2.text_input("NIT")
        direccion = st.text_input("Dirección")
        col3, col4 = st.columns(2)
        telefono = col3.text_input("Teléfono")
        email = col4.text_input("Email")
        valor = st.number_input("Valor mensual habitual (COP)", min_value=0, step=100000)
        notas = st.text_area("Notas")
        submitted = st.form_submit_button("Guardar cliente")
        if submitted and nombre:
            agregar_cliente(nombre, nit, direccion, telefono, email, valor, notas)
            st.success(f"Cliente '{nombre}' agregado y carpeta creada.")
            st.rerun()
    
    st.divider()
    st.subheader("Clientes registrados")
    df_clientes = get_clientes(activos_only=False)
    if not df_clientes.empty:
        st.dataframe(df_clientes, use_container_width=True)

with tab3:
    st.header("Generar Cuentas de Cobro")
    st.info("Las cuentas se generan automáticamente el día 25 de cada mes (puedes forzar la generación). Se organizan en carpetas por cliente.")
    
    hoy = date.today()
    es_dia_25 = hoy.day == 25
    if es_dia_25:
        st.success("✅ Hoy es día 25. Puedes generar las cuentas del mes.")
    else:
        st.warning(f"Hoy es día {hoy.day}. La generación automática está programada para el 25.")
    
    clientes_df = get_clientes()
    if clientes_df.empty:
        st.warning("Primero agrega clientes en la pestaña Clientes.")
    else:
        # Generación individual
        st.subheader("Generar cuenta individual")
        cliente_sel = st.selectbox("Seleccionar cliente", clientes_df['nombre'].tolist())
        cliente_row = clientes_df[clientes_df['nombre'] == cliente_sel].iloc[0]
        
        col1, col2 = st.columns(2)
        numero = col1.text_input("Número de cuenta", value=f"{len(pd.read_sql_query('SELECT * FROM cuentas_cobro', get_conn())) + 104:03d}_{hoy.strftime('%y')}")
        periodo = col2.text_input("Periodo", value=hoy.strftime("%B %Y").capitalize())
        valor = st.number_input("Valor (COP)", value=float(cliente_row['valor_mensual']), step=100000.0)
        
        if st.button("Generar Cuenta de Cobro"):
            cliente_dict = cliente_row.to_dict()
            ruta = generar_cuenta_cobro(cliente_dict, numero, periodo, valor, hoy)
            
            # Registrar en BD
            conn = get_conn()
            c = conn.cursor()
            c.execute('''
                INSERT INTO cuentas_cobro (cliente_id, numero, fecha, periodo, valor, ruta_archivo)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (cliente_row['id'], numero, hoy.isoformat(), periodo, valor, ruta))
            conn.commit()
            conn.close()
            
            st.success(f"Cuenta generada: {ruta}")
            with open(ruta, "rb") as f:
                st.download_button("Descargar Cuenta de Cobro", f, file_name=os.path.basename(ruta))
        
        st.divider()
        # Generación masiva (para el 25)
        st.subheader("Generación masiva (día 25 o forzar)")
        if st.button("Generar todas las cuentas del mes para clientes activos", type="primary"):
            generadas = 0
            for _, cli in clientes_df.iterrows():
                if cli['valor_mensual'] > 0:
                    num = f"{generadas + 104:03d}_{hoy.strftime('%y')}"
                    periodo_mes = hoy.strftime("%B %Y").capitalize()
                    ruta = generar_cuenta_cobro(cli.to_dict(), num, periodo_mes, cli['valor_mensual'], hoy)
                    conn = get_conn()
                    c = conn.cursor()
                    c.execute('''
                        INSERT INTO cuentas_cobro (cliente_id, numero, fecha, periodo, valor, ruta_archivo)
                        VALUES (?, ?, ?, ?, ?, ?)
                    ''', (cli['id'], num, hoy.isoformat(), periodo_mes, cli['valor_mensual'], ruta))
                    conn.commit()
                    conn.close()
                    generadas += 1
            st.success(f"Se generaron {generadas} cuentas de cobro en las carpetas de cada cliente.")

with tab4:
    st.header("Históricos de Cuentas de Cobro")
    conn = get_conn()
    df = pd.read_sql_query('''
        SELECT c.numero, cl.nombre as cliente, c.fecha, c.periodo, c.valor, c.estado, c.ruta_archivo
        FROM cuentas_cobro c
        JOIN clientes cl ON c.cliente_id = cl.id
        ORDER BY c.fecha DESC
    ''', conn)
    conn.close()
    if not df.empty:
        st.dataframe(df, use_container_width=True)
        # Permitir cambiar estado
        st.subheader("Actualizar estado de pago")
        if not df.empty:
            cuenta_id_sel = st.selectbox("Seleccionar cuenta", df['numero'].tolist())
            nuevo_estado = st.selectbox("Nuevo estado", ["Pendiente", "Pagada", "Vencida"])
            if st.button("Actualizar estado"):
                conn = get_conn()
                conn.execute("UPDATE cuentas_cobro SET estado = ? WHERE numero = ?", (nuevo_estado, cuenta_id_sel))
                conn.commit()
                conn.close()
                st.success("Estado actualizado")
                st.rerun()
    else:
        st.info("Aún no hay cuentas de cobro generadas.")

# Pie de página
st.sidebar.markdown("---")
st.sidebar.markdown("**Cumplimiento UG**")
st.sidebar.markdown("Gestión de Cartera")
st.sidebar.markdown("📱 Compatible con móvil y escritorio")
st.sidebar.markdown("---")
st.sidebar.info("Para generación automática el día 25, puedes programar una tarea en tu computador que ejecute la app o use el botón de generación masiva.")
