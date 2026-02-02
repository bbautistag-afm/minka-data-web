import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

st.set_page_config(page_title="Minka-Data Web", page_icon="💎", layout="wide")

st.title("💎 MINKA-DATA: Procesador Web de Actas")
st.markdown("### 🏛️ UGEL Melgar - Innovación Tecnológica")

archivos = st.file_uploader("📂 Arrastre sus Actas en PDF aquí", type="pdf", accept_multiple_files=True)

if archivos:
    st.info(f"✅ {len(archivos)} archivos cargados.")
    
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
                                # Nueva lógica más flexible: Busca un DNI de 8 dígitos y captura lo que sigue
                                match = re.search(r'(\d{8})\s+([A-ZÑÁÉÍÓÚ\s,]+)\s+([\d\sA-D]+)$', fila)
                                
                                if match:
                                    lista_resultados.append({
                                        "DNI": match.group(1),
                                        "Estudiante": match.group(2).strip(),
                                        "Notas": match.group(3).strip(),
                                        "Institución": "MARIANO MELGAR",
                                        "Archivo": archivo.name
                                    })
            except Exception as e:
                st.error(f"Error en {archivo.name}")
            
            progreso.progress((i + 1) / len(archivos))
        
        if lista_resultados:
            df = pd.DataFrame(lista_resultados)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False)
            
            st.balloons()
            st.success(f"📊 Se extrajeron {len(lista_resultados)} registros con éxito.")
            st.download_button("📥 Descargar Excel", data=output.getvalue(), file_name="DATA_MELGAR.xlsx")
        else:
            st.error("❌ El formato de este PDF es distinto. ¡No te rindas! Vamos a ajustarlo.")
