import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# 1. CONFIGURACIÓN Y LOGO REAL
st.set_page_config(page_title="Minka-Data Melgar", page_icon="💎", layout="wide")

# Colocamos el logo oficial (Si tienes el link directo del logo de la UGEL, cámbialo aquí)
col1, col2 = st.columns([1, 5])
with col1:
    # He puesto un logo genérico de educación, pero puedes cambiarlo por la URL del logo de la UGEL Melgar
    st.image("https://drive.google.com/uc?export=view&id=1cGVMXE6RdzgsVXGbWIUGgVZDbfXReQhb", width=100)
with col2:
    st.title("💎 MINKA-DATA: Procesador Web de Actas")
    st.markdown("### 🏛️ UGEL Melgar - Innovación Tecnológica")

st.info("Sube tus archivos PDF para iniciar el procesamiento masivo.")

# --- TU LÓGICA DE EXTRACCIÓN (Mantenida intacta porque funciona) ---
def limpiar(t):
    return re.sub(r'\s+', ' ', str(t)).strip() if t else ""

def procesar_acta_universal(pdf_file):
    alumnos_acumulados = {}
    nombre_archivo = pdf_file.name
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

# --- INTERFAZ ---
archivos_cargados = st.file_uploader("📂 Selecciona o arrastra las actas PDF aquí", type="pdf", accept_multiple_files=True)

if archivos_cargados:
    col_acc1, col_acc2 = st.columns(2)
    
    with col_acc1:
        if st.button("🚀 INICIAR PROCESAMIENTO"):
            lista_maestra = []
            progreso = st.progress(0)
            for i, pdf_file in enumerate(archivos_cargados):
                datos = procesar_acta_universal(pdf_file)
                lista_maestra.extend(datos)
                progreso.progress((i + 1) / len(archivos_cargados))

            if lista_maestra:
                df_base = pd.DataFrame(lista_maestra)
                df_notas = pd.DataFrame(df_base["NOTAS_LISTA"].tolist()).add_prefix('COMP_')
                df_final = pd.concat([df_base.drop(columns=["NOTAS_LISTA", "SIT_FINAL"]), df_notas, df_base["SIT_FINAL"]], axis=1)
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, index=False)
                
                st.balloons()
                st.success(f"📊 {len(lista_maestra)} alumnos procesados.")
                st.download_button("📥 Descargar Excel Final", data=output.getvalue(), file_name="Data_Minka_Melgar.xlsx")
    
    with col_acc2:
        # 2. BOTÓN PARA LIMPIAR TODO
        if st.button("♻️ LIMPIAR PARA NUEVA CARGA"):
            st.rerun() # Esto refresca la página y borra los archivos subidos
