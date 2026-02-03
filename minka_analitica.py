import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# 1. CONFIGURACIÓN E IDENTIDAD
st.set_page_config(page_title="Minka-Data ANALÍTICA", page_icon="📊", layout="wide")

# Inicializar estado para el botón de limpiar
if 'reset_key' not in st.session_state:
    st.session_state.reset_key = 0

def limpiar_campos():
    st.session_state.reset_key += 1
    st.rerun()

# Barra lateral con Logo e Identidad
with st.sidebar:
    st.image("https://i.ibb.co/k2n2fHLZ/Logo-UGEL-Melgar-especial.png", width=200)
    st.title("UGEL - GESTIÓN PEI")
    st.markdown("---")
    if st.button("♻️ LIMPIAR TODO"):
        limpiar_campos()

st.title("📊 MINKA-DATA: Inteligencia Gerencial")
st.markdown("### 🏛️ Diagnóstico de Gestión Escolar (CGE 1 y 2)")

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
                digitos_idx = [i for i, c in enumerate(f_str) if c.isdigit() and len(c) == 1]
                dni_raw = "".join([f_str[i] for i in digitos_idx if 4 < i < 16])
                if len(dni_raw) == 8:
                    dni = dni_raw
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

# --- CARGA DE ARCHIVOS ---
archivos = st.file_uploader("📂 Suba sus actas PDF", type="pdf", accept_multiple_files=True, key=f"uploader_{st.session_state.reset_key}")

if archivos and st.button("🚀 GENERAR REPORTE GERENCIAL"):
    data_total = []
    for f in archivos:
        data_total.extend(procesar_acta_universal(f))
    
    if data_total:
        df_base = pd.DataFrame(data_total)
        
        # PROCESAMIENTO CGE 1
        notas_list = []
        for reg in data_total:
            for n in reg["NOTAS"]:
                notas_list.append({"AÑO": reg["AÑO"], "NIVEL": n})
        df_cge1 = pd.DataFrame(notas_list).groupby(['AÑO', 'NIVEL']).size().unstack(fill_value=0)
        orden_cge1 = [c for c in ['AD', 'A', 'B', 'C'] if c in df_cge1.columns]
        df_cge1 = df_cge1[orden_cge1]
        df_cge1['TOTAL'] = df_cge1.sum(axis=1)
        df_cge1_pct = df_cge1.iloc[:, :-1].div(df_cge1['TOTAL'], axis=0) * 100
        df_cge1_pct['TOTAL_%'] = 100.0

        # PROCESAMIENTO CGE 2
        df_cge2 = df_base.groupby(['AÑO', 'SIT']).size().unstack(fill_value=0)
        df_cge2['TOTAL_MATRICULA'] = df_cge2.sum(axis=1)
        df_cge2_pct = df_cge2.div(df_cge2['TOTAL_MATRICULA'], axis=0) * 100

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_base.to_excel(writer, sheet_name='DATOS_CRUDOS', index=False)
            
            workbook = writer.book
            fmt_pct = workbook.add_format({'num_format': '0.0"%"'})
            
            # --- PESTAÑA CGE 1 ---
            df_cge1.to_excel(writer, sheet_name='ANALISIS_CGE1', startrow=1)
            df_cge1_pct.to_excel(writer, sheet_name='ANALISIS_CGE1', startrow=len(df_cge1)+5)
            sheet1 = writer.sheets['ANALISIS_CGE1']
            sheet1.write('A1', 'TABLA 1: RECUENTO DE NIVELES')
            sheet1.write(f'A{len(df_cge1)+5}', 'TABLA 2: PORCENTAJES (%)')
            
            chart1 = workbook.add_chart({'type': 'column'})
            colores = {'AD': '#0070C0', 'A': '#00B050', 'B': '#FFC000', 'C': '#FF0000'}
            for i, nivel in enumerate(orden_cge1):
                chart1.add_series({
                    'name': ['ANALISIS_CGE1', 1, i+1],
                    'categories': ['ANALISIS_CGE1', 2, 0, len(df_cge1)+1, 0],
                    'values': ['ANALISIS_CGE1', 2, i+1, len(df_cge1)+1, i+1],
                    'fill': {'color': colores.get(nivel)},
                    'data_labels': {'value': True}
                })
            sheet1.insert_chart('J2', chart1)

            # --- PESTAÑA CGE 2 ---
            df_cge2.to_excel(writer, sheet_name='ANALISIS_CGE2', startrow=1)
            df_cge2_pct.to_excel(writer, sheet_name='ANALISIS_CGE2', startrow=len(df_cge2)+5)
            sheet2 = writer.sheets['ANALISIS_CGE2']
            sheet2.write('A1', 'TABLA 1: SITUACIÓN FINAL (RECUENTO)')
            sheet2.write(f'A{len(df_cge2)+5}', 'TABLA 2: SITUACIÓN FINAL (%)')
            
            # Gráfico de Matrícula Total
            chart_mat = workbook.add_chart({'type': 'column'})
            col_t = df_cge2.columns.get_loc('TOTAL_MATRICULA') + 1
            chart_mat.add_series({
                'name': 'Matrícula Total',
                'categories': ['ANALISIS_CGE2', 2, 0, len(df_cge2)+1, 0],
                'values': ['ANALISIS_CGE2', 2, col_t, len(df_cge2)+1, col_t],
                'fill': {'color': '#7030A0'},
                'data_labels': {'value': True}
            })
            sheet2.insert_chart('J2', chart_mat)

            # RESTAURACIÓN: Gráfico de Situaciones por Año
            chart_sit = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
            for i, sit in enumerate(df_cge2.columns[:-1]):
                chart_sit.add_series({
                    'name': ['ANALISIS_CGE2', 1, i+1],
                    'categories': ['ANALISIS_CGE2', 2, 0, len(df_cge2)+1, 0],
                    'values': ['ANALISIS_CGE2', 2, i+1, len(df_cge2)+1, i+1],
                    'data_labels': {'value': True}
                })
            chart_sit.set_title({'name': 'Distribución de Situación Final por Año'})
            sheet2.insert_chart('J18', chart_sit)

        st.balloons()
        st.success("✅ Reporte Generado con Identidad Institucional")
        st.download_button("📥 Descargar Reporte Final", data=output.getvalue(), file_name="Minka_Data_Final_UGEL.xlsx")
