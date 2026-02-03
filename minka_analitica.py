import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# 1. CONFIGURACIÓN E IDENTIDAD
st.set_page_config(page_title="Minka-Data ANALÍTICA", page_icon="📈", layout="wide")

# Inicialización de estado para limpiar datos
if 'reset_key' not in st.session_state:
    st.session_state.reset_key = 0

def limpiar_datos():
    st.session_state.reset_key += 1
    st.rerun()

# --- BARRA LATERAL ESTILIZADA ---
with st.sidebar:
    # Logo UGEL Melgar (Usamos el oficial del MINEDU)
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/b/bb/Logo_del_Ministerio_de_Educaci%C3%B3n_del_Per%C3%BA.svg/1200px-Logo_del_Ministerio_de_Educaci%C3%B3n_del_Per%C3%BA.svg.png", width=200)
    
    st.markdown("### **UGEL MELGAR**")
    st.markdown("---")
    
    # Autoría en fuente menuda y estrecha
    st.markdown("""
    <div style="font-size: 11px; line-height: 1.2; color: #555;">
        <b>Autor:</b> Bernardo Bautista Gutiérrez<br>
        <b>Email:</b> bbautistag@ugelmelgar.edu.pe<br>
        <b>Cel:</b> 965 654 898
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    st.info("📊 Herramienta de Monitoreo Estratégico para el Liderazgo Pedagógico.")

# --- CUERPO PRINCIPAL ---
st.title("📈 MINKA-DATA: Inteligencia de Gestión Educativa")
st.markdown("#### 🏛️ Diagnóstico de Compromisos de Gestión Escolar (CGE 1 y 2)")

# --- FUNCIONES ---
def limpiar(t):
    return re.sub(r'\s+', ' ', str(t)).strip() if t else ""

def procesar_acta_universal(pdf_file):
    alumnos_acumulados = {}
    nombre_archivo = pdf_file.name
    anio_match = re.search(r'(202[3-6])', nombre_archivo)
    anio = anio_match.group(0) if anio_match else "2025"
    SITUACIONES_OFICIALES = ['PRO', 'RR', 'T', 'F', 'PER', 'R', 'PE', 'AE', 'PG']

    with pdfplumber.open(pdf_file) as pdf:
        for pagina in pdf.pages:
            tabla = pagina.extract_table()
            if not tabla: continue
            for fila in tabla:
                f_str = [limpiar(c) for c in fila]
                dni_raw = "".join([c for c in f_str if c.isdigit() and len(c) == 1])
                # Ajuste para capturar DNI en el flujo de la tabla
                dni_search = re.search(r'\d{8}', "".join(f_str))
                if dni_search:
                    dni = dni_search.group(0)
                    if dni not in alumnos_acumulados:
                        alumnos_acumulados[dni] = {"AÑO": anio, "NOTAS": [], "SIT": "N/A"}
                    for celda in f_str:
                        if celda in ['AD', 'A', 'B', 'C']:
                            alumnos_acumulados[dni]["NOTAS"].append(celda)
                    for celda in f_str:
                        if celda in SITUACIONES_OFICIALES:
                            alumnos_acumulados[dni]["SIT"] = celda
                            break
    return list(alumnos_acumulados.values())

# --- CARGA Y BOTONES ---
archivos = st.file_uploader("📂 Cargue actas PDF (2023, 2024, 2025)", type="pdf", accept_multiple_files=True, key=f"up_{st.session_state.reset_key}")

col1, col2 = st.columns([2, 1])
with col1:
    btn_generar = st.button("🚀 GENERAR REPORTE DE GESTIÓN EDUCATIVA", use_container_width=True)
with col2:
    btn_limpiar = st.button("♻️ LIMPIAR DATOS", on_click=limpiar_datos, use_container_width=True)

if archivos and btn_generar:
    data_total = []
    for f in archivos:
        data_total.extend(procesar_acta_universal(f))
    
    if data_total:
        df_base = pd.DataFrame(data_total)
        
        # CGE 1 - Logros
        notas_list = []
        for reg in data_total:
            for n in reg["NOTAS"]:
                notas_list.append({"AÑO": reg["AÑO"], "NIVEL": n})
        df_cge1 = pd.DataFrame(notas_list).groupby(['AÑO', 'NIVEL']).size().unstack(fill_value=0)
        orden = [c for c in ['AD', 'A', 'B', 'C'] if c in df_cge1.columns]
        df_cge1 = df_cge1[orden]
        df_cge1['TOTAL'] = df_cge1.sum(axis=1)
        df_cge1_pct = df_cge1.iloc[:, :-1].div(df_cge1['TOTAL'], axis=0) * 100

        # CGE 2 - Matrícula
        df_cge2 = df_base.groupby(['AÑO', 'SIT']).size().unstack(fill_value=0)
        df_cge2['TOTAL_MATR'] = df_cge2.sum(axis=1)
        df_cge2_pct = df_cge2.div(df_cge2['TOTAL_MATR'], axis=0) * 100

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_base.to_excel(writer, sheet_name='DATOS_CRUDOS', index=False)
            workbook = writer.book
            fmt_pct = workbook.add_format({'num_format': '0.00"%"'}) # Formato con 2 decimales
            
            # --- HOJA CGE 1 ---
            df_cge1.to_excel(writer, sheet_name='ANALISIS_CGE1', startrow=1)
            df_cge1_pct.to_excel(writer, sheet_name='ANALISIS_CGE1', startrow=len(df_cge1)+5)
            sh1 = writer.sheets['ANALISIS_CGE1']
            sh1.set_column('B:G', 12, fmt_pct)
            
            # Gráfico de Logros con Porcentajes
            chart1 = workbook.add_chart({'type': 'column'})
            colores = {'AD': '#0070C0', 'A': '#00B050', 'B': '#FFC000', 'C': '#FF0000'}
            for i, nivel in enumerate(orden):
                chart1.add_series({
                    'name':       ['ANALISIS_CGE1', 1, i+1],
                    'categories': ['ANALISIS_CGE1', 2, 0, len(df_cge1)+1, 0],
                    # Gráfico basado en la tabla de porcentajes
                    'values':     ['ANALISIS_CGE1', len(df_cge1)+6, i+1, len(df_cge1)*2+5, i+1],
                    'fill':       {'color': colores.get(nivel)},
                    'data_labels': {'value': True, 'num_format': '0.00"%"'}
                })
            chart1.set_title({'name': 'Niveles de Logro (%) - CGE 1'})
            sh1.insert_chart('J2', chart1)

            # --- HOJA CGE 2 ---
            df_cge2.to_excel(writer, sheet_name='ANALISIS_CGE2', startrow=1)
            df_cge2_pct.to_excel(writer, sheet_name='ANALISIS_CGE2', startrow=len(df_cge2)+5)
            sh2 = writer.sheets['ANALISIS_CGE2']
            sh2.set_column('B:M', 12, fmt_pct) # Formato porcentaje a la tabla de CGE2
            
            # Gráfico de Tendencia Histórica (Línea)
            chart_tend = workbook.add_chart({'type': 'line'})
            col_total = df_cge2.columns.get_loc('TOTAL_MATR') + 1
            chart_tend.add_series({
                'name': 'Tendencia Matrícula Total',
                'categories': ['ANALISIS_CGE2', 2, 0, len(df_cge2)+1, 0],
                'values':     ['ANALISIS_CGE2', 2, col_total, len(df_cge2)+1, col_total],
                'line':       {'color': '#FF5733', 'width': 2.5},
                'marker':     {'type': 'circle', 'size': 8},
                'data_labels': {'value': True}
            })
            chart_tend.set_title({'name': 'Histórico de Matrícula Escolar (Tendencia)'})
            sh2.insert_chart('J2', chart_tend)

        st.balloons()
        st.success("✅ Diamante 2 actualizado y listo.")
        st.download_button("📥 Descargar Reporte Final", data=output.getvalue(), file_name="Minka_Analisis_Final.xlsx")
