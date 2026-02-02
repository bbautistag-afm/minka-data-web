import streamlit as st
import pdfplumber
import pandas as pd
import os
import re
import io

# --- FASE 1: CONEXIÓN ---

def limpiar(t):
    return re.sub(r'\s+', ' ', str(t)).strip() if t else ""

def procesar_acta_universal(ruta_pdf, nombre_archivo):
    alumnos_acumulados = {}

    # --- EXTRACCIÓN POR NOMENCLATURA (IE y GRADO) ---
    partes = nombre_archivo.replace('.pdf', '').split(' - ')
    cod_modular = partes[0] if len(partes) > 0 else "N/A"
    nombre_ie = partes[1] if len(partes) > 1 else "IE DESCONOCIDA"

    resto = partes[2] if len(partes) > 2 else ""
    gra_match = re.search(r'(\d+)(ro|do|to|a)', resto.lower())
    grado_texto = gra_match.group(0) if gra_match else "N/A"

    sec_match = re.search(r'\s([A-Z])(?:\s|$)', resto.upper())
    seccion = sec_match.group(1) if sec_match else "N/A"

    # --- LISTA DE SITUACIONES FINALES BLINDADA ---
    # Incluye las básicas y las nuevas: PER, R, PE, AE, PG
    siglas_leyenda = ['PRO', 'RR', 'T', 'F', 'PER', 'R', 'PE', 'AE', 'PG', 'PROMOVIDO', 'FALLECIDO', 'RETIRADO']

    with pdfplumber.open(ruta_pdf) as pdf:
        for pagina in pdf.pages:
            tabla = pagina.extract_table()
            if not tabla: continue

            for fila in tabla:
                f_str = [limpiar(c) for c in fila]

                # BUSCADOR DE DNI (Mantenemos la lógica que ya funciona)
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

                    # CAPTURA DE NOTAS (Blindada)
                    for i, celda in enumerate(f_str):
                        if i in digitos_idx or i < 5: continue
                        if celda in ['AD', 'A', 'B', 'C', 'T'] or (celda.isdigit() and 0 <= int(celda) <= 20):
                            alumnos_acumulados[dni]["NOTAS_LISTA"].append(celda)

                    # BÚSQUEDA DE SITUACIÓN FINAL (Ampliada con los nuevos casos)
                    sit_actual = [c for c in f_str if c in siglas_leyenda]
                    if sit_actual:
                        # Priorizamos 'F' o 'R' por seguridad
                        val = sit_actual[0]
                        if "FALLECIDO" in sit_actual or "F" in sit_actual: val = "F"
                        elif "RETIRADO" in sit_actual or "R" in sit_actual: val = "R"
                        alumnos_acumulados[dni]["SIT_FINAL"] = val

    return list(alumnos_acumulados.values())

# --- FASE 3: EJECUCIÓN DINÁMICA ---
# Cambia aquí el nombre de la carpeta según el nivel
NIVEL_A_PROCESAR = 'ACTA_INICIAL/'

CARPETA_RAIZ = '/content/drive/MyDrive/001_PROYECTO_PEI_ACTA_PDF/'
CARPETA_TRABAJO = os.path.join(CARPETA_RAIZ, NIVEL_A_PROCESAR)
NOMBRE_EXCEL = f"MINKA_DATA_MELGAR_{NIVEL_A_PROCESAR.replace('/', '')}.xlsx"
RUTA_SALIDA = os.path.join(CARPETA_RAIZ, NOMBRE_EXCEL)

if os.path.exists(CARPETA_TRABAJO):
    archivos = sorted([f for f in os.listdir(CARPETA_TRABAJO) if f.endswith('.pdf')])
    print(f"🚀 Iniciando MINKA-DATA para: {NIVEL_A_PROCESAR}")

    lista_maestra = []
    for a in archivos:
        print(f"📄 Extrayendo: {a}")
        lista_maestra.extend(procesar_acta_universal(os.path.join(CARPETA_TRABAJO, a), a))

    if lista_maestra:
        df_base = pd.DataFrame(lista_maestra)
        df_notas = pd.DataFrame(df_base["NOTAS_LISTA"].tolist()).add_prefix('COMP_')
        df_final = pd.concat([df_base.drop(columns=["NOTAS_LISTA", "SIT_FINAL"]), df_notas, df_base["SIT_FINAL"]], axis=1)

        df_final.to_excel(RUTA_SALIDA, index=False)
        print(f"\n🏆 ¡SISTEMA MINKA-DATA COMPLETADO!")
        print(f"📂 Archivo generado en la raíz: {NOMBRE_EXCEL}")
