import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

st.set_page_config(page_title="Minka-Data ANALÍTICA", page_icon="📊", layout="wide")

st.title("📊 MINKA-DATA: Inteligencia Gerencial PEI")
st.markdown("### 🏛️ Diagnóstico Completo CGE 1 y CGE 2")

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

archivos = st.file_uploader("📂 Suba sus actas PDF", type="pdf", accept_multiple_files=True)

if archivos and st.button("🚀 GENERAR REPORTE INTELIGENTE"):
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
        # ORDEN ESTRICTO AD, A, B, C
        orden_cge1 = [c for c in ['AD', 'A', 'B', 'C'] if c in df_cge1.columns]
        df_cge1 = df_cge1[orden_cge1]
        df_cge1_pct = df_cge1.div(df_cge1.sum(axis=1), axis=0) * 100

        # --- PROCESAMIENTO CGE 2 (Permanencia) ---
        df_cge2 = df_base.groupby(['AÑO', 'SIT']).size().unstack(fill_value=0)
        df_cge2['TOTAL_MATRICULA'] = df_cge2.sum(axis=1) # Sumatoria columna E
        df_cge2_pct = df_cge2.div(df_cge2['TOTAL_MATRICULA'], axis=0) * 100

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_base.to_excel(writer, sheet_name='DATOS_CRUDOS', index=False)
            
            # HOJA CGE 1: LOGROS
            df_cge1.to_excel(writer, sheet_name='ANALISIS_CGE1', startrow=1)
            df_cge1_pct.to_excel(writer, sheet_name='ANALISIS_CGE1', startrow=len(df_cge1)+5)
            
            # HOJA CGE 2: PERMANENCIA
            df_cge2.to_excel(writer, sheet_name='ANALISIS_CGE2', startrow=1)
            df_cge2_pct.to_excel(writer, sheet_name='ANALISIS_CGE2', startrow=len(df_cge2)+5)
            
            workbook = writer.book
            colores = {'AD': '#0070C0', 'A': '#00B050', 'B': '#FFC000', 'C': '#FF0000'}

            # --- GRÁFICOS CGE 1 ---
            sheet1 = writer.sheets['ANALISIS_CGE1']
            sheet1.write('A1', 'TABLA 1: RECUENTO DE NIVELES')
            sheet1.write(f'A{len(df_cge1)+5}', 'TABLA 2: PORCENTAJES (%)')
            
            # Gráfico de Cantidades
            chart1 = workbook.add_chart({'type': 'column'})
            for i, nivel in enumerate(df_cge1.columns):
                chart1.add_series({
                    'name': ['ANALISIS_CGE1', 1, i+1],
                    'categories': ['ANALISIS_CGE1', 2, 0, len(df_cge1)+1, 0],
                    'values': ['ANALISIS_CGE1', 2, i+1, len(df_cge1)+1, i+1],
                    'fill': {'color': colores.get(nivel)},
                    'data_labels': {'value': True, 'position': 'outside_end'}
                })
            chart1.set_title({'name': 'CGE 1: Cantidad de Logros (AD-A-B-C)'})
            sheet1.insert_chart('J2', chart1)

            # --- GRÁFICOS CGE 2 ---
            sheet2 = writer.sheets['ANALISIS_CGE2']
            sheet2.write('A1', 'TABLA 1: RECUENTO DE MATRÍCULA Y SITUACIÓN')
            
            # Gráfico Histórico de Matrícula (TOTAL)
            chart_mat = workbook.add_chart({'type': 'column'})
            idx_total = df_cge2.columns.get_loc('TOTAL_MATRICULA') + 1
            chart_mat.add_series({
                'name': 'Matrícula Total',
                'categories': ['ANALISIS_CGE2', 2, 0, len(df_cge2)+1, 0],
                'values': ['ANALISIS_CGE2', 2, idx_total, len(df_cge2)+1, idx_total],
                'fill': {'color': '#7030A0'},
                'data_labels': {'value': True}
            })
            chart_mat.set_title({'name': 'Histórico de Matrícula Escolar'})
            sheet2.insert_chart('J2', chart_mat)

            # Gráfico Barras Verticales Situaciones Finales
            chart_sit = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
            for i, sit in enumerate(df_cge2.columns[:-1]): # Excluir el total
                chart_sit.add_series({
                    'name': ['ANALISIS_CGE2', 1, i+1],
                    'categories': ['ANALISIS_CGE2', 2, 0, len(df_cge2)+1, 0],
                    'values': ['ANALISIS_CGE2', 2, i+1, len(df_cge2)+1, i+1],
                    'data_labels': {'value': True}
                })
            chart_sit.set_title({'name': 'CGE 2: Situaciones Finales por Año'})
            sheet2.insert_chart('J18', chart_sit)

        st.balloons()
        st.success("✅ ¡Inteligencia Gerencial Completada!")
        st.download_button("📥 Descargar Reporte Estratégico", data=output.getvalue(), file_name="Minka_Data_Inteligencia.xlsx")
