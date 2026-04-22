# ===============================
# MERGE SECOP NO ESTRUCTURADO (VERSIÓN COMENTADA)
# ===============================
# Este script realiza:
# 1. Merge de contratos con URLs
# 2. Merge de procedimientos con URLs
# 3. Unión por URL
# 4. Fuzzy merge (LEFT JOIN) con estudios previos
# ===============================

import re
import sys
import unicodedata
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd

# -------------------------------
# CONFIGURACIÓN
# -------------------------------
FUZZY_THRESHOLD = 80.0  # 🔧 Ajusta aquí el nivel de exigencia del match

# -------------------------------
# FUNCIONES DE LIMPIEZA
# -------------------------------
def normalize_text(text):
    if text is None:
        return ""
    text = str(text)
    text = unicodedata.normalize("NFD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    text = text.lower()
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()

# -------------------------------
# SIMILITUD (usa rapidfuzz si existe)
# -------------------------------
def similarity(a, b):
    try:
        from rapidfuzz import fuzz
        return fuzz.token_set_ratio(a, b)
    except:
        from difflib import SequenceMatcher
        return SequenceMatcher(None, a, b).ratio() * 100

# -------------------------------
# CARGA DE ARCHIVOS
# -------------------------------
def cargar_archivos(ruta):
    ruta = Path(ruta)
    return {
        "url": pd.read_excel(ruta / "URLDocumento_NombresContratos_NombresSECOPProcedimiento.xlsx"),
        "contratos": pd.read_excel(ruta / "Contratos_Extraidos.xlsx"),
        "procedimientos": pd.read_excel(ruta / "secop_procedimiento_extraidos.xlsx"),
        "estudios": pd.read_excel(ruta / "estudios_previos_extraidos.xlsx"),
    }

# -------------------------------
# JOIN 1
# -------------------------------
def merge_contratos(contratos, url):
    contratos["key"] = contratos["archivo"].apply(lambda x: str(x).split("/")[-1])
    url["key"] = url["contratos_pdf_paths"].apply(lambda x: str(x).split("/")[-1])
    url = url.drop_duplicates("key")
    return contratos.merge(url[["key", "url"]], on="key", how="left")

# -------------------------------
# JOIN 2
# -------------------------------
def merge_procedimientos(proc, url):
    proc["key"] = proc["archivo_pdf"].apply(lambda x: str(x).split("/")[-1])
    url["key"] = url["Secop_procedimientos_pdf_path"].apply(lambda x: str(x).split("/")[-1])
    url = url.drop_duplicates("key")
    return proc.merge(url[["key", "url"]], on="key", how="left")

# -------------------------------
# JOIN 3
# -------------------------------
def merge_por_url(c1, c2):
    return c1.merge(c2, on="url", how="left")

# -------------------------------
# FUZZY LEFT JOIN
# -------------------------------
def fuzzy_merge(left, right, col_left, col_right):
    right_norm = right[col_right].fillna("").apply(normalize_text).tolist()

    resultados = []
    for _, row in left.iterrows():
        texto = normalize_text(row[col_left])
        mejor_score = 0
        mejor_idx = None

        for i, r in enumerate(right_norm):
            score = similarity(texto, r)
            if score > mejor_score:
                mejor_score = score
                mejor_idx = i

        nueva_fila = row.to_dict()
        nueva_fila["Proximidad_Objeto_descripcion"] = mejor_score

        if mejor_score >= FUZZY_THRESHOLD and mejor_idx is not None:
            for col in right.columns:
                nueva_fila[col] = right.iloc[mejor_idx][col]
        else:
            for col in right.columns:
                if col not in nueva_fila:
                    nueva_fila[col] = None

        resultados.append(nueva_fila)

    return pd.DataFrame(resultados)

# -------------------------------
# MAIN
# -------------------------------
def main():
    carpeta = input("📂 Ruta de inputs: ")
    salida = input("💾 Ruta de outputs: ")

    data = cargar_archivos(carpeta)

    print("🔄 Merge contratos...")
    contratos_url = merge_contratos(data["contratos"], data["url"])

    print("🔄 Merge procedimientos...")
    proc_url = merge_procedimientos(data["procedimientos"], data["url"])

    print("🔄 Merge por URL...")
    minutas = merge_por_url(contratos_url, proc_url)

    print("🔄 Fuzzy merge...")
    col_desc = "descripcion" if "descripcion" in minutas.columns else "descripción"
    final = fuzzy_merge(minutas, data["estudios"], col_desc, "OBJETO")

    salida = Path(salida)
    salida.mkdir(exist_ok=True)

    contratos_url.to_excel(salida / "Contratos_Extraidos_URL.xlsx", index=False)
    proc_url.to_excel(salida / "secop_procedimiento_extraidos_URL.xlsx", index=False)
    minutas.to_excel(salida / "MinutasYProcedimientosSECOP.xlsx", index=False)
    final.to_excel(salida / "SECOP_NoEstructurado.xlsx", index=False)

    print("✅ Proceso terminado")

if __name__ == "__main__":
    main()
