import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# 1. CONFIGURACIÓN DE PÁGINA Y ESCUDO OFICIAL
st.set_page_config(page_title="Minka-Data Melgar", page_icon="💎", layout="wide")

# Inicializamos la llave de reseteo si no existe
if 'reset_key' not in st.session_state:
    st.session_state.reset_key = 0

col1, col2 = st.columns([1, 5])
with col1:
    # Tu escudo oficial de la UGEL Melgar (vía ImgBB o link directo)
    st.image("https://i.ibb.co/647ea5.jpg", width=110) # He usado el ID de tu última captura
with col2:
    st.title("💎 MINKA-DATA: Procesador Web de Actas")
    st.markdown("### 🏛️ UGEL Melgar - Innovación Tecnológica")

st.info("Bienvenido. Cargue las actas en PDF para consolidar la información en un solo archivo Excel.")

# --- FUNCIONES DE LIMPIEZA Y PROCESAMIENTO ---
def limpiar(t):
    return re.sub(r'\s+', ' ', str(t)).strip() if t else ""

def procesar_acta_universal(pdf_file):
    alumnos_acumulados = {}
    nombre_archivo = pdf_file.name
    
    # Nomenclatura del archivo
    partes = nombre_archivo.replace('.pdf', '').split(' - ')
    cod_modular = partes[0] if len(partes) > 0 else "N/A"
    nombre_ie = partes[1] if len(partes) > 1 else "IE DESCONOCIDA"
    resto = partes[2] if len(partes) > 2 else ""
    gra_match = re.search(r'(\d+)(ro|do|to|a)', resto.lower())
    grado_texto = gra_match.group(0) if gra_match else "N/A"
    sec_match = re.search(r'\s([A-Z])(?:\s|$)', resto.upper())
    seccion = sec_match.group(1) if sec_match else "N/A"

    siglas_leyenda = ['PRO', 'RR', 'T', 'F', 'PER', 'R', 'PE', 'AE', 'PG', 'PROMOVIDO', 'FALLECIDO', 'RETIRADO']

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
                        nombre = next((c for c in f_str if len(c) > 12 and not c.isdigit()), "N/A")
                        sexo_raw = next((c for c in f_str if c in ['H', 'M']), "N/A")
                        genero = "Hombre" if sexo_raw == "H" else "Mujer" if sexo_raw == "M" else "N/A"

                        alumnos_acumulados[dni] = {
                            "UGEL": "MELGAR", "COD_MOD": cod_modular, "IE": nombre_ie,
                            "MOD": "EBR", "GRA": grado_texto, "SEC": seccion,
                            "DNI": dni, "ESTUDIANTE": nombre, "SEXO": genero,
                            "NOTAS_LISTA": [], "SIT_FINAL": "N/A"
                        }

                    for i, celda in enumerate(f_str):
                        if i in digitos_idx or i < 5: continue
                        if celda in ['AD', 'A', 'B', 'C', 'T'] or (celda.isdigit() and 0 <= int(celda) <= 20):
                            alumnos_acumulados[dni]["NOTAS_LISTA"].append(celda)

                    sit_actual = [c for c in f_str if c in siglas_leyenda]
                    if sit_actual:
                        val = sit_actual[0]
                        if "FALLECIDO" in sit_actual or "F" in sit_actual: val = "F"
                        elif "RETIRADO" in sit_actual or "R" in sit_actual: val = "R"
                        alumnos_acumulados[dni]["SIT_FINAL"] = val

    return list(alumnos_acumulados.values())

# --- CARGADOR CON LLAVE DINÁMICA ---
archivos_cargados = st.file_uploader(
    "📂 Selecciona o arrastra las actas PDF aquí", 
    type="pdf", 
    accept_multiple_files=True,
    key=f"uploader_{st.session_state.reset_key}"
)

if archivos_cargados:
    col_btn1, col_btn2 = st.columns(2)
    
    with col_btn1:
        if st.button("🚀 INICIAR PROCESAMIENTO MASIVO"):
            lista_maestra = []
            barra = st.progress(0)
            
            for i, pdf_file in enumerate(archivos_cargados):
                datos = procesar_acta_universal(pdf_file)
                lista_maestra.extend(datos)
                barra.progress((i + 1) / len(archivos_cargados))

            if lista_maestra:
                df_base = pd.DataFrame(lista_maestra)
                df_notas = pd.DataFrame(df_base["NOTAS_LISTA"].tolist()).add_prefix('COMP_')
                df_final = pd.concat([df_base.drop(columns=["NOTAS_LISTA", "SIT_FINAL"]), df_notas, df_base["SIT_FINAL"]], axis=1)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False)
                
                st.balloons()
                st.success(f"📊 ¡Éxito! {len(lista_maestra)} alumnos consolidados.")
                st.download_button("📥 Descargar Excel Consolidado", data=output.getvalue(), file_name="Minka_Data_Melgar.xlsx")
            else:
                st.error("No se encontraron datos válidos.")

    with col_btn2:
        # BOTÓN DE LIMPIEZA ATÓMICA
        if st.button("♻️ LIMPIAR PARA NUEVA CARGA"):
            st.session_state.reset_key += 1
            st.rerun()
