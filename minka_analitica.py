import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

st.set_page_config(page_title="Minka-Data ANALÍTICA", page_icon="📊", layout="wide")

st.title("📊 MINKA-DATA: Módulo de Analítica PEI")
st.markdown("### 🏛️ Diagnóstico de Gestión Escolar (CGE 1 y 2)")

def limpiar(t):
    return re.sub(r'\s+', ' ', str(t)).strip() if t else ""

def procesar_acta_universal(pdf_file):
    alumnos_acumulados = {}
    nombre_archivo = pdf_file.name
    anio_match = re.search(r'(202[3-6])', nombre_archivo)
    anio = anio_match.group(0) if anio_match else "2025"
    
    # Lista Maestra de Situaciones Finales
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
                    
                    # Capturar Notas AD, A, B, C
                    for celda in f_str:
                        if celda in ['AD', 'A', 'B', 'C']:
                            alumnos_acumulados[dni]["NOTAS"].append(celda)
                    
                    # Capturar Situación Final (Mapeo completo)
                    for celda in f_str:
                        if celda in SITUACIONES_OFICIALES:
                            alumnos_acumulados[dni]["SIT"] = celda
                            break # Encontrada la situación principal

    return list(alumnos_acumulados.values())

archivos = st.file_uploader("📂 Suba sus actas PDF", type="pdf", accept_multiple_files=True)

if archivos and st.button("🚀 GENERAR REPORTE GERENCIAL"):
    data_total = []
    for f in archivos:
        data_total.extend(procesar_acta_universal(f))
    
    if data_total:
        df_base = pd.DataFrame(data_total)
        
        # --- PROCESAMIENTO CGE 1 (Logros) ---
        notas_list = []
        for reg in data_total:
            for n in reg["NOTAS"]:
                notas_list.append({"AÑO": reg["AÑO"], "NIVEL": n})
        df_cge1 = pd.DataFrame(notas_list).groupby(['AÑO', 'NIVEL']).size().unstack(fill_value=0)
        
        # --- PROCESAMIENTO CGE 2 (Permanencia) ---
        df_cge2 = df_base.groupby(['AÑO', 'SIT']).size().unstack(fill_value=0)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_base.to_excel(writer, sheet_name='DATOS_CRUDOS', index=False)
            df_cge1.to_excel(writer, sheet_name='ANALISIS_CGE1')
            df_cge2.to_excel(writer, sheet_name='ANALISIS_CGE2')
            
            workbook = writer.book
            
            # --- CONFIGURAR GRÁFICO CGE 1 (Mapa de Calor) ---
            if not df_cge1.empty:
                sheet1 = writer.sheets['ANALISIS_CGE1']
                chart1 = workbook.add_chart({'type': 'column'})
                colores = {'AD': '#0070C0', 'A': '#00B050', 'B': '#FFC000', 'C': '#FF0000'}
                
                for i, nivel in enumerate(df_cge1.columns):
                    chart1.add_series({
                        'name': ['ANALISIS_CGE1', 0, i+1],
                        'categories': ['ANALISIS_CGE1', 1, 0, len(df_cge1), 0],
                        'values': ['ANALISIS_CGE1', 1, i+1, len(df_cge1), i+1],
                        'fill': {'color': colores.get(nivel, '#D3D3D3')},
                        'data_labels': {'value': True, 'position': 'outside_end'}
                    })
                chart1.set_title({'name': 'HISTÓRICO CGE 1: Niveles de Logro'})
                sheet1.insert_chart('G2', chart1)

            # --- CONFIGURAR GRÁFICO CGE 2 (Permanencia) ---
            if not df_cge2.empty:
                sheet2 = writer.sheets['ANALISIS_CGE2']
                chart2 = workbook.add_chart({'type': 'bar'})
                chart2.add_series({
                    'name': 'Situación Final',
                    'categories': ['ANALISIS_CGE2', 1, 0, len(df_cge2), 0],
                    'values': ['ANALISIS_CGE2', 1, 1, len(df_cge2), 1],
                    'data_labels': {'value': True}
                })
                chart2.set_title({'name': 'CGE 2: Situación Final de Alumnos'})
                sheet2.insert_chart('G2', chart2)

        st.balloons()
        st.success("✅ ¡Reporte Gerencial Finalizado!")
        st.download_button("📥 Descargar Reporte Completo", data=output.getvalue(), file_name="Minka_Data_Analitica_Final.xlsx")
