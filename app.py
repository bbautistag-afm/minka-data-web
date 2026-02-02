import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# Configuración de la página
st.set_page_config(page_title="Minka-Data Melgar", page_icon="💎", layout="wide")

# Títulos (Lo que ya ves en pantalla)
st.title("💎 MINKA-DATA: Procesador Web de Actas")
st.markdown("### 🏛️ UGEL Melgar - Innovación Tecnológica")
st.info("Bienvenido al sistema de procesamiento masivo. Esta herramienta extrae datos de actas PDF y los consolida en Excel.")

# EL MOTOR: Cuadro de carga de archivos
archivos_pdf = st.file_uploader("📂 Arrastre sus Actas en PDF aquí", type="pdf", accept_multiple_files=True)

if archivos_pdf:
    st.success(f"✅ {len(archivos_pdf)} archivos listos para procesar.")
    
    if st.button("🚀 INICIAR PROCESAMIENTO MASIVO"):
        datos_totales = []
        barra_progreso = st.progress(0)
        
        for i, archivo in enumerate(archivos_pdf):
            try:
                with pdfplumber.open(archivo) as pdf:
                    for pagina in pdf.pages:
                        texto = pagina.extract_text()
                        if texto:
                            # Aquí va tu lógica de extracción del Diamante Pulido
                            for linea in texto.split('\n'):
                                # Ejemplo de captura de DNI y Nombre (ajusta según tu lógica original)
                                match = re.search(r'(\d{8})\s+([A-ZÑÁÉÍÓÚ\s,]+)', linea)
                                if match:
                                    datos_totales.append({
                                        "DNI": match.group(1),
                                        "Estudiante": match.group(2).strip(),
                                        "Archivo": archivo.name
                                    })
            except Exception as e:
                st.error(f"Error en {archivo.name}: {e}")
            
            barra_progreso.progress((i + 1) / len(archivos_pdf))

        if datos_totales:
            df = pd.DataFrame(datos_totales)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False)
            
            st.balloons()
            st.success(f"📊 ¡Éxito! Se procesaron {len(datos_totales)} registros.")
            st.download_button("📥 Descargar Excel Consolidado", data=output.getvalue(), file_name="Data_Minka_Melgar.xlsx")
        else:
            st.warning("⚠️ No se encontraron datos. Verifica que los PDF sean actas oficiales.")
