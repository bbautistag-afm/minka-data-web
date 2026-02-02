import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

# Configuración de la página
st.set_page_config(page_title="Minka-Data Web", page_icon="💎", layout="wide")

st.title("💎 MINKA-DATA: Procesador Web de Actas")
st.markdown("""
    ### 🏛️ UGEL Melgar - Innovación Tecnológica
    Bienvenido al sistema de procesamiento masivo. Esta herramienta extrae datos de actas PDF y los consolida en Excel.
""")

# Selector de archivos
archivos = st.file_uploader("📂 Arrastre sus Actas en PDF aquí", type="pdf", accept_multiple_files=True)

if archivos:
    st.info(f"✅ {len(archivos)} archivos cargados listos para procesar.")
    
    if st.button("🚀 INICIAR PROCESAMIENTO MASIVO"):
        lista_resultados = []
        progreso = st.progress(0)
        
        for i, archivo in enumerate(archivos):
            try:
                with pdfplumber.open(archivo) as pdf:
                    for pagina in pdf.pages:
                        texto = pagina.extract_text()
                        if texto:
                            for fila in texto.split('\n'):
                                # EXPRESIÓN DIAMANTE: Busca DNI (8 dígitos) + Nombre + Notas
                                match = re.search(r'(\d{8})\s+([A-ZÁÉÍÓÚÑ\s,]+)\s+([A-D0-9\s]+)$', fila)
                                
                                if match:
                                    lista_resultados.append({
                                        "DNI": match.group(1),
                                        "Estudiante": match.group(2).strip(),
                                        "Calificaciones": match.group(3).strip(),
                                        "Archivo Origen": archivo.name
                                    })
            except Exception as e:
                st.warning(f"⚠️ Error al leer {archivo.name}")
            
            progreso.progress((i + 1) / len(archivos))
        
        if lista_resultados:
            df_final = pd.DataFrame(lista_resultados)
            
            # Crear archivo Excel en memoria
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False, sheet_name='DATA_EXTRAIDA')
            
            st.balloons()
            st.success(f"📊 ¡Proceso terminado! Se extrajeron {len(lista_resultados)} registros.")
            
            st.download_button(
                label="📥 Descargar Excel Consolidado",
                data=output.getvalue(),
                file_name="MINKA_DATA_CONSOLIDADO.xlsx",
                mime="application/vnd.ms-excel"
            )
        else:
            st.error("❌ No se encontraron datos válidos. Verifique el formato de las actas.")

st.sidebar.markdown("---")
st.sidebar.info("Desarrollado para la mejora educativa en Melgar.")
