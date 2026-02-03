import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# 1. CONFIGURACIÓN DEL LABORATORIO DE ANALÍTICA
st.set_page_config(page_title="Minka-Data ANALÍTICA", page_icon="📊", layout="wide")

if 'reset_key' not in st.session_state:
    st.session_state.reset_key = 0

st.title("📊 MINKA-DATA: Módulo de Analítica PEI")
st.markdown("### 🏛️ Diagnóstico de Compromisos de Gestión (CGE 1 y 2)")
st.info("Sugerencia: Nombre sus archivos empezando por el año (Ej: '2024 - Acta.pdf') para ver el histórico.")

# --- FUNCIONES DE EXTRACCIÓN ---
def limpiar(t):
    return re.sub(r'\s+', ' ', str(t)).strip() if t else ""

def procesar_acta_universal(pdf_file):
    alumnos_acumulados = {}
    nombre_archivo = pdf_file.name
    
    # DETECTOR DE AÑO (Busca 2023, 2024, 2025 en el nombre)
    anio_match = re.search(r'(202[3-5])', nombre_archivo)
    anio = anio_match.group(0) if anio_match else "2025" # Por defecto 2025
    
    with pdfplumber.open(pdf_file) as pdf:
        for pagina in pdf.pages:
            tabla = pagina.extract_table()
            if not tabla: continue
            for fila in tabla:
                f_str = [limpiar(c) for c in fila]
                # Detectar DNI (8 dígitos)
                digitos_idx = [i for i, c in enumerate(f_str) if c.isdigit() and len(c) == 1]
                dni_raw = "".join([f_str[i] for i in digitos_idx if 4 < i < 16])
                
                if len(dni_raw) == 8:
                    dni = dni_raw
                    if dni not in alumnos_acumulados:
                        alumnos_acumulados[dni] = {"AÑO": anio, "NOTAS": [], "SIT": "N/A"}
                    
                    # Capturar Notas (AD, A, B, C)
                    for celda in f_str:
                        if celda in ['AD', 'A', 'B', 'C']:
                            alumnos_acumulados[dni]["NOTAS"].append(celda)
                    
                    # Situación Final (CGE 2)
                    sit_final = [c for c in f_str if c in ['PRO', 'PG', 'RR', 'R', 'F', 'PER']]
                    if sit_final: alumnos_acumulados[dni]["SIT"] = sit_final[0]

    return list(alumnos_acumulados.values())

# --- INTERFAZ ---
archivos = st.file_uploader("📂 Cargue actas de varios años", type="pdf", accept_multiple_files=True, key=f"an_{st.session_state.reset_key}")

if archivos and st.button("🚀 GENERAR DIAGNÓSTICO HISTÓRICO"):
    data_total = []
    for f in archivos:
        data_total.extend(procesar_acta_universal(f))
    
    if data_total:
        df_base = pd.DataFrame(data_total)
        
        # PROCESAR CGE 1 (Aprendizajes)
        notas_list = []
        for reg in data_total:
            for n in reg["NOTAS"]:
                notas_list.append({"AÑO": reg["AÑO"], "NIVEL": n})
        df_cge1 = pd.DataFrame(notas_list).groupby(['AÑO', 'NIVEL']).size().unstack(fill_value=0)
        
        # PROCESAR CGE 2 (Permanencia)
        df_cge2 = df_base.groupby(['AÑO', 'SIT']).size().unstack(fill_value=0)

        # GENERAR EXCEL
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_base.to_excel(writer, sheet_name='DATOS_CRUDOS', index=False)
            df_cge1.to_excel(writer, sheet_name='ANALISIS_CGE1')
            df_cge2.to_excel(writer, sheet_name='ANALISIS_CGE2')
            
            workbook = writer.book
            # Gráfico CGE 1
            if not df_cge1.empty:
                sheet1 = writer.sheets['ANALISIS_CGE1']
                chart1 = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
                for i, nivel in enumerate(df_cge1.columns):
                    chart1.add_series({
                        'name': ['ANALISIS_CGE1', 0, i+1],
                        'categories': ['ANALISIS_CGE1', 1, 0, len(df_cge1), 0],
                        'values': ['ANALISIS_CGE1', 1, i+1, len(df_cge1), i+1],
                    })
                chart1.set_title({'name': 'HISTÓRICO CGE 1: Niveles de Logro'})
                sheet1.insert_chart('G2', chart1)

        st.balloons()
        st.success("¡Diagnóstico PEI listo!")
        st.download_button("📥 Descargar Reporte de Gestión", data=output.getvalue(), file_name="Minka_Data_Analitica.xlsx")
