import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import gc  # Garbage Collector para limpiar memoria

st.set_page_config(page_title="Minka-Data ANALÍTICA", page_icon="📈", layout="wide")

if 'reset_key' not in st.session_state:
    st.session_state.reset_key = 0

def limpiar_datos():
    st.session_state.reset_key += 1
    st.rerun()

# --- BARRA LATERAL ---
with st.sidebar:
    st.image("https://i.ibb.co/k2n2fHLZ/Logo-UGEL-Melgar-especial.png", width=180)
    st.markdown("### **UGEL MELGAR**")
    st.markdown("---")
    st.markdown(f"""
    <div style="font-size: 11px; line-height: 1.2; color: #555;">
        <b>Autor:</b> Bernardo Bautista Gutiérrez<br>
        <b>Email:</b> bbautistag@ugelmelgar.edu.pe<br>
        <b>Cel:</b> 965 654 898
    </div>
    """, unsafe_allow_html=True)
    st.info("📊 Optimizado para procesamiento masivo de I.E. grandes.")

st.title("📈 MINKA-DATA: Inteligencia de Gestión Educativa")
st.markdown("#### 🏛️ Diagnóstico de CGE 1 y 2 (Versión Alto Rendimiento)")

def limpiar(t):
    return re.sub(r'\s+', ' ', str(t)).strip() if t else ""

def procesar_acta_universal(pdf_file):
    alumnos_acumulados = {}
    nombre_archivo = pdf_file.name
    anio_match = re.search(r'(202[3-6])', nombre_archivo)
    anio = anio_match.group(0) if anio_match else "2025"
    siglas = ['PRO', 'RR', 'T', 'F', 'PER', 'R', 'PE', 'AE', 'PG']

    try:
        with pdfplumber.open(pdf_file) as pdf:
            for pagina in pdf.pages:
                tabla = pagina.extract_table()
                if not tabla: continue
                for fila in tabla:
                    f_str = [limpiar(c) for c in fila]
                    dni_search = re.search(r'\d{8}', "".join(f_str))
                    if dni_search:
                        dni = dni_search.group(0)
                        if dni not in alumnos_acumulados:
                            alumnos_acumulados[dni] = {"AÑO": anio, "NOTAS": [], "SIT": "N/A"}
                        for celda in f_str:
                            if celda in ['AD', 'A', 'B', 'C']:
                                alumnos_acumulados[dni]["NOTAS"].append(celda)
                            if celda in siglas:
                                alumnos_acumulados[dni]["SIT"] = celda
    except Exception as e:
        st.warning(f"Error en archivo {nombre_archivo}: {e}")
    return list(alumnos_acumulados.values())

archivos = st.file_uploader("📂 Cargue las 126 actas aquí", type="pdf", accept_multiple_files=True, key=f"up_{st.session_state.reset_key}")

if archivos and st.button("🚀 GENERAR REPORTE ESTRATÉGICO"):
    data_total = []
    progreso = st.progress(0)
    
    for i, f in enumerate(archivos):
        data_total.extend(procesar_acta_universal(f))
        progreso.progress((i + 1) / len(archivos))
        if i % 20 == 0: gc.collect() # Limpia memoria cada 20 archivos

    if data_total:
        with st.spinner("Construyendo tablas y gráficos..."):
            df_base = pd.DataFrame(data_total)
            
            # Procesamiento CGE 1
            notas_list = []
            for reg in data_total:
                for n in reg["NOTAS"]:
                    notas_list.append({"AÑO": reg["AÑO"], "NIVEL": n})
            
            df_cge1 = pd.DataFrame(notas_list).groupby(['AÑO', 'NIVEL']).size().unstack(fill_value=0)
            orden = [c for c in ['AD', 'A', 'B', 'C'] if c in df_cge1.columns]
            df_cge1 = df_cge1[orden]
            df_cge1['TOTAL'] = df_cge1.sum(axis=1)
            df_cge1_pct = df_cge1.iloc[:, :-1].div(df_cge1['TOTAL'], axis=0) * 100

            # Procesamiento CGE 2
            df_cge2 = df_base.groupby(['AÑO', 'SIT']).size().unstack(fill_value=0)
            df_cge2['TOTAL_MATR'] = df_cge2.sum(axis=1)
            df_cge2_pct = df_cge2.div(df_cge2['TOTAL_MATR'], axis=0) * 100

            output = io.BytesIO()
            # Usar 'xlsxwriter' es más eficiente para archivos grandes
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_base.to_excel(writer, sheet_name='DATOS_CRUDOS', index=False)
                workbook = writer.book
                fmt_pct = workbook.add_format({'num_format': '0.00"%"'})
                fmt_num = workbook.add_format({'num_format': '#,##0'})

                # PESTAÑA CGE 1
                df_cge1.to_excel(writer, sheet_name='ANALISIS_CGE1', startrow=1)
                df_cge1_pct.to_excel(writer, sheet_name='ANALISIS_CGE1', startrow=len(df_cge1)+5)
                sh1 = writer.sheets['ANALISIS_CGE1']
                sh1.write('A1', 'TABLA 1: RECUENTO DE NIVELES (CANTIDADES)')
                sh1.write(f'A{len(df_cge1)+5}', 'TABLA 2: PORCENTAJES (%)')
                sh1.set_column('B:G', 12, fmt_num)
                
                chart1 = workbook.add_chart({'type': 'column'})
                for i, nivel in enumerate(orden):
                    chart1.add_series({
                        'name': ['ANALISIS_CGE1', 1, i+1],
                        'categories': ['ANALISIS_CGE1', 2, 0, len(df_cge1)+1, 0],
                        'values': ['ANALISIS_CGE1', len(df_cge1)+6, i+1, len(df_cge1)*2+5, i+1],
                        'data_labels': {'value': True, 'num_format': '0.00"%"'}
                    })
                sh1.insert_chart('J2', chart1)

                # PESTAÑA CGE 2
                df_cge2.to_excel(writer, sheet_name='ANALISIS_CGE2', startrow=1)
                df_cge2_pct.to_excel(writer, sheet_name='ANALISIS_CGE2', startrow=len(df_cge2)+5)
                sh2 = writer.sheets['ANALISIS_CGE2']
                sh2.write('A1', 'TABLA 1: SITUACIÓN FINAL (CANTIDADES)')
                sh2.write(f'A{len(df_cge2)+5}', 'TABLA 2: PORCENTAJES (%)')
                
                # Gráfico Tendencia Matrícula
                chart_tend = workbook.add_chart({'type': 'line'})
                col_t = df_cge2.columns.get_loc('TOTAL_MATR') + 1
                chart_tend.add_series({
                    'name': 'Matrícula Total',
                    'categories': ['ANALISIS_CGE2', 2, 0, len(df_cge2)+1, 0],
                    'values': ['ANALISIS_CGE2', 2, col_t, len(df_cge2)+1, col_t],
                    'marker': {'type': 'circle'},
                    'data_labels': {'value': True}
                })
                sh2.insert_chart('J2', chart_tend)

            st.balloons()
            st.success("✅ ¡Procesamiento masivo completado con éxito!")
            st.download_button("📥 Descargar Reporte Final (Melgar)", data=output.getvalue(), file_name="Minka_Analisis_Masivo.xlsx")
