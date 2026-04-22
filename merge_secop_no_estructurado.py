# -*- coding: utf-8 -*-
"""
MERGE SECOP NO ESTRUCTURADO - VERSIÓN PRO
OUTPUT FINAL: DOS ARCHIVOS XLSX

AJUSTE SOLICITADO EN PASO 3:
- "Contratos_Extraidos_URL" y "secop_procedimiento_extraidos_URL" se unen por "url"
- LEFT JOIN con base principal = "Contratos_Extraidos_URL"
- "MinutasYProcedimientosSECOP" debe conservar EXACTAMENTE las filas de "Contratos_Extraidos_URL"
"""

import math
import os
import re
import sys
import unicodedata
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd

# =========================================================
# CONFIGURACIÓN
# =========================================================
INPUT_FILES = {
    "url_base": "URLDocumento_NombresContratos_NombresSECOPProcedimiento.xlsx",
    "contratos": "Contratos_Extraidos.xlsx",
    "procedimientos": "secop_procedimiento_extraidos.xlsx",
    "estudios": "estudios_previos_extraidos.xlsx",
}

FUZZY_THRESHOLD = 80.0
WEAK_MATCH_MIN = 70.0
WEAK_MATCH_MAX = 79.99
MAX_WORKERS = max(1, min(8, (os.cpu_count() or 4)))

OUTPUT_WORKBOOK_NAME = "SECOP_NoEstructurado_Consolidado.xlsx"
OUTPUT_FINAL_SHEET_NAME = "SECOP_NoEstructurado.xlsx"


# =========================================================
# UTILIDADES GENERALES
# =========================================================
def normalize_spaces(text: str) -> str:
    text = "" if text is None else str(text)
    text = text.replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\s*\n\s*", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def strip_accents(text: str) -> str:
    text = "" if text is None else str(text)
    return "".join(
        ch for ch in unicodedata.normalize("NFD", text)
        if unicodedata.category(ch) != "Mn"
    )


def normalize_for_match(text: str) -> str:
    text = normalize_spaces(text)
    text = strip_accents(text).lower()
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def basename_from_pathlike(value: str) -> str:
    value = "" if value is None else str(value)
    value = value.strip().strip('"').strip("'")
    value = value.replace("\\", "/")
    value = value.split("/")[-1]
    return value.strip()


def prompt_existing_folder(message: str) -> Path:
    while True:
        folder = input(message).strip().strip('"').strip("'")
        path = Path(folder)
        if path.exists() and path.is_dir():
            return path
        print(f"❌ La carpeta no existe o no es válida: {path}")


def prompt_output_folder(message: str) -> Path:
    while True:
        folder = input(message).strip().strip('"').strip("'")
        path = Path(folder)
        if not path.exists():
            create = input(f"La carpeta no existe. ¿Deseas crearla? [s/n]: ").strip().lower()
            if create in {"s", "si", "sí", "y", "yes"}:
                path.mkdir(parents=True, exist_ok=True)
                return path
            print("❌ Debes ingresar una carpeta válida.")
            continue
        if path.is_dir():
            return path
        print(f"❌ La ruta no corresponde a una carpeta: {path}")


def find_required_files(input_folder: Path) -> Dict[str, Path]:
    found = {}
    missing = []
    for key, filename in INPUT_FILES.items():
        full_path = input_folder / filename
        if full_path.exists():
            found[key] = full_path
        else:
            missing.append(filename)

    if missing:
        raise FileNotFoundError(
            "No se encontraron los siguientes archivos en la carpeta indicada:\n- "
            + "\n- ".join(missing)
        )
    return found


def explain_excel_permission_error(path: Path) -> str:
    return (
        f"No se pudo abrir o guardar el archivo:\n{path}\n\n"
        "Posibles causas:\n"
        "1. El archivo está abierto en Excel.\n"
        "2. OneDrive lo está sincronizando o bloqueando.\n"
        "3. No tienes permisos sobre esa carpeta.\n\n"
        "Prueba esto:\n"
        "- Cierra el archivo Excel.\n"
        "- Pausa OneDrive temporalmente.\n"
        "- Mueve los archivos a una carpeta local simple, por ejemplo: C:\\SECOP_Merge\\\n"
        "- Vuelve a ejecutar el script."
    )


def read_excel_file(path: Path) -> pd.DataFrame:
    try:
        return pd.read_excel(path)
    except PermissionError as e:
        raise PermissionError(explain_excel_permission_error(path)) from e
    except Exception as e:
        raise RuntimeError(f"Error leyendo el archivo Excel '{path}': {e}") from e


def validate_required_columns(df: pd.DataFrame, required_columns: List[str], df_name: str) -> None:
    missing = [col for col in required_columns if col not in df.columns]
    if missing:
        raise ValueError(
            f"La base '{df_name}' no contiene estas columnas requeridas: {missing}\n"
            f"Columnas disponibles: {list(df.columns)}"
        )


def detect_descripcion_column(df: pd.DataFrame) -> str:
    candidates = ["descripción", "descripcion", "Descripción", "Descripcion"]
    for c in candidates:
        if c in df.columns:
            return c
    raise ValueError(
        "No encontré la columna 'descripción' o 'descripcion' en MinutasYProcedimientosSECOP.\n"
        f"Columnas disponibles: {list(df.columns)}"
    )


def ensure_unique_right_columns(left_df: pd.DataFrame, right_df: pd.DataFrame, join_key: str) -> pd.DataFrame:
    right = right_df.copy()
    rename_map = {}
    for col in right.columns:
        if col != join_key and col in left_df.columns:
            rename_map[col] = f"{col}_estudios"
    if rename_map:
        right = right.rename(columns=rename_map)
    return right


# =========================================================
# FUZZY MATCH
# =========================================================
def _rapidfuzz_available() -> bool:
    try:
        import rapidfuzz  # noqa: F401
        return True
    except Exception:
        return False


def similarity_score(a: str, b: str) -> float:
    if not a or not b:
        return 0.0

    try:
        from rapidfuzz import fuzz
        ratio = float(fuzz.ratio(a, b))
        token_sort = float(fuzz.token_sort_ratio(a, b))
        token_set = float(fuzz.token_set_ratio(a, b))
        partial = float(fuzz.partial_ratio(a, b))
        score = (0.40 * ratio) + (0.25 * token_sort) + (0.25 * token_set) + (0.10 * partial)
        return round(score, 2)
    except Exception:
        from difflib import SequenceMatcher
        return round(SequenceMatcher(None, a, b).ratio() * 100, 2)


def best_match_one(query_norm: str, choices_norm: List[str]) -> Tuple[Optional[int], float]:
    if not query_norm:
        return None, 0.0

    try:
        from rapidfuzz import process, fuzz
        scored = process.extractOne(query_norm, choices_norm, scorer=fuzz.token_set_ratio)
        if scored is None:
            return None, 0.0
        matched_text, _, idx = scored
        score = similarity_score(query_norm, matched_text)
        return int(idx), score
    except Exception:
        best_idx = None
        best_score = 0.0
        for idx, choice in enumerate(choices_norm):
            score = similarity_score(query_norm, choice)
            if score > best_score:
                best_score = score
                best_idx = idx
        return best_idx, round(best_score, 2)


def _best_match_chunk(chunk_queries: List[str], choices_norm: List[str]) -> Dict[str, Tuple[Optional[int], float]]:
    result = {}
    for query_norm in chunk_queries:
        result[query_norm] = best_match_one(query_norm, choices_norm)
    return result


def build_best_match_map_parallel(unique_left_values: List[str], right_choices_norm: List[str]) -> Dict[str, Tuple[Optional[int], float]]:
    if not unique_left_values:
        return {}

    n = len(unique_left_values)
    workers = min(MAX_WORKERS, n)

    if workers <= 1 or n < 200:
        result = {}
        for i, query_norm in enumerate(unique_left_values, start=1):
            result[query_norm] = best_match_one(query_norm, right_choices_norm)
            if i % 250 == 0 or i == n:
                print(f"   Progreso fuzzy: {i}/{n}")
        return result

    chunk_size = math.ceil(n / workers)
    chunks = [unique_left_values[i:i + chunk_size] for i in range(0, n, chunk_size)]

    result: Dict[str, Tuple[Optional[int], float]] = {}
    done = 0
    with ThreadPoolExecutor(max_workers=workers) as executor:
        futures = [executor.submit(_best_match_chunk, chunk, right_choices_norm) for chunk in chunks]
        for future in as_completed(futures):
            partial = future.result()
            result.update(partial)
            done += len(partial)
            print(f"   Progreso fuzzy: {done}/{n}")

    return result


def fuzzy_left_merge_best_match(
    left_df: pd.DataFrame,
    right_df: pd.DataFrame,
    left_on: str,
    right_on: str,
    threshold: float = 80.0,
    proximity_col: str = "Proximidad_Objeto_descripcion",
) -> pd.DataFrame:
    validate_required_columns(left_df, [left_on], "MinutasYProcedimientosSECOP")
    validate_required_columns(right_df, [right_on], "estudios_previos_extraidos")

    left = left_df.copy()
    right = right_df.copy()

    right = ensure_unique_right_columns(left, right, join_key=right_on)
    right_key = right_on

    left["_left_match_norm"] = left[left_on].fillna("").astype(str).map(normalize_for_match)
    right["_right_match_norm"] = right[right_key].fillna("").astype(str).map(normalize_for_match)
    right_valid = right[right["_right_match_norm"] != ""].copy().reset_index(drop=True)

    if right_valid.empty:
        result = left.copy()
        result[proximity_col] = 0.0
        for col in right.columns:
            if col not in {right_on, "_right_match_norm"} and col not in result.columns:
                result[col] = pd.NA
        result.drop(columns=["_left_match_norm"], inplace=True, errors="ignore")
        return result

    unique_left_values = (
        left["_left_match_norm"]
        .fillna("")
        .astype(str)
        .drop_duplicates()
        .tolist()
    )
    right_choices_norm = right_valid["_right_match_norm"].tolist()

    print(f"🔎 Iniciando fuzzy match sobre {len(unique_left_values)} descripciones únicas...")
    print(f"   - rapidfuzz disponible: {_rapidfuzz_available()}")
    print(f"   - hilos usados: {min(MAX_WORKERS, max(1, len(unique_left_values)))}")

    best_match_map = build_best_match_map_parallel(unique_left_values, right_choices_norm)

    output_rows = []
    right_columns_to_add = [c for c in right_valid.columns if c != "_right_match_norm"]

    for _, row in left.iterrows():
        row_dict = row.to_dict()
        query_norm = row_dict.get("_left_match_norm", "")
        best_idx, best_score = best_match_map.get(query_norm, (None, 0.0))
        row_dict[proximity_col] = round(best_score, 2)

        if best_idx is not None and best_score >= threshold:
            matched = right_valid.iloc[best_idx].to_dict()
            matched.pop("_right_match_norm", None)
            for col, value in matched.items():
                if col not in row_dict:
                    row_dict[col] = value
                elif col == right_on:
                    row_dict[f"{col}_estudios"] = value
                else:
                    row_dict[col] = value
        else:
            for col in right_columns_to_add:
                if col not in row_dict:
                    row_dict[col] = pd.NA

        output_rows.append(row_dict)

    result = pd.DataFrame(output_rows)
    result.drop(columns=["_left_match_norm"], inplace=True, errors="ignore")

    if len(result) != len(left_df):
        raise RuntimeError(
            f"El fuzzy left join alteró el número de filas. "
            f"Esperadas: {len(left_df)}, obtenidas: {len(result)}"
        )

    return result


# =========================================================
# JOINS 1, 2 Y 3
# =========================================================
def merge_contratos_with_url(contratos_df: pd.DataFrame, url_df: pd.DataFrame) -> pd.DataFrame:
    validate_required_columns(contratos_df, ["archivo"], "Contratos_Extraidos")
    validate_required_columns(
        url_df,
        ["contratos_pdf_paths", "url"],
        "URLDocumento_NombresContratos_NombresSECOPProcedimiento",
    )

    left = contratos_df.copy()
    right = url_df[["contratos_pdf_paths", "url"]].copy()

    left["_join_key"] = left["archivo"].map(basename_from_pathlike)
    right["_join_key"] = right["contratos_pdf_paths"].map(basename_from_pathlike)
    right = right.drop_duplicates(subset=["_join_key"], keep="first")

    merged = left.merge(
        right[["_join_key", "url"]],
        on="_join_key",
        how="left",
        validate="m:1",
    )
    merged.drop(columns=["_join_key"], inplace=True, errors="ignore")

    if len(merged) != len(contratos_df):
        raise RuntimeError(
            f"El join de Contratos_Extraidos alteró el número de filas. "
            f"Esperadas: {len(contratos_df)}, obtenidas: {len(merged)}"
        )

    return merged


def merge_procedimientos_with_url(procedimientos_df: pd.DataFrame, url_df: pd.DataFrame) -> pd.DataFrame:
    validate_required_columns(procedimientos_df, ["archivo_pdf"], "secop_procedimiento_extraidos")
    validate_required_columns(
        url_df,
        ["Secop_procedimientos_pdf_path", "url"],
        "URLDocumento_NombresContratos_NombresSECOPProcedimiento",
    )

    left = procedimientos_df.copy()
    right = url_df[["Secop_procedimientos_pdf_path", "url"]].copy()

    left["_join_key"] = left["archivo_pdf"].map(basename_from_pathlike)
    right["_join_key"] = right["Secop_procedimientos_pdf_path"].map(basename_from_pathlike)
    right = right.drop_duplicates(subset=["_join_key"], keep="first")

    merged = left.merge(
        right[["_join_key", "url"]],
        on="_join_key",
        how="left",
        validate="m:1",
    )
    merged.drop(columns=["_join_key"], inplace=True, errors="ignore")

    if len(merged) != len(procedimientos_df):
        raise RuntimeError(
            f"El join de secop_procedimiento_extraidos alteró el número de filas. "
            f"Esperadas: {len(procedimientos_df)}, obtenidas: {len(merged)}"
        )

    return merged


def merge_minutas_and_procedimientos(contratos_url_df: pd.DataFrame, procedimientos_url_df: pd.DataFrame) -> pd.DataFrame:
    """
    PASO 3 AJUSTADO:
    LEFT JOIN con base principal = Contratos_Extraidos_URL
    La tabla resultante MinutasYProcedimientosSECOP debe tener exactamente
    las mismas filas que Contratos_Extraidos_URL.
    """
    validate_required_columns(contratos_url_df, ["url"], "Contratos_Extraidos_URL")
    validate_required_columns(procedimientos_url_df, ["url"], "secop_procedimiento_extraidos_URL")

    left = contratos_url_df.copy()
    right = procedimientos_url_df.copy()

    rename_map = {}
    for col in right.columns:
        if col != "url" and col in left.columns:
            rename_map[col] = f"{col}_procedimiento"
    if rename_map:
        right = right.rename(columns=rename_map)

    merged = left.merge(right, on="url", how="left")

    if len(merged) != len(contratos_url_df):
        raise RuntimeError(
            f"El paso 3 alteró el número de filas de Contratos_Extraidos_URL. "
            f"Esperadas: {len(contratos_url_df)}, obtenidas: {len(merged)}"
        )

    return merged


# =========================================================
# REPORTES
# =========================================================
def build_matching_reports(df: pd.DataFrame, proximity_col: str = "Proximidad_Objeto_descripcion") -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    prox = pd.to_numeric(df[proximity_col], errors="coerce").fillna(0)

    no_match = df.loc[prox < FUZZY_THRESHOLD].copy()
    weak_match = df.loc[(prox >= WEAK_MATCH_MIN) & (prox <= WEAK_MATCH_MAX)].copy()

    resumen = pd.DataFrame(
        {
            "Métrica": [
                "Filas MinutasYProcedimientosSECOP",
                "Filas SECOP_NoEstructurado",
                f"Matches aceptados (>= {FUZZY_THRESHOLD})",
                f"No match (< {FUZZY_THRESHOLD})",
                f"Matches débiles ({WEAK_MATCH_MIN} a {WEAK_MATCH_MAX})",
                "Score promedio",
                "Score máximo",
                "Score mínimo",
            ],
            "Valor": [
                len(df),
                len(df),
                int((prox >= FUZZY_THRESHOLD).sum()),
                int((prox < FUZZY_THRESHOLD).sum()),
                int(((prox >= WEAK_MATCH_MIN) & (prox <= WEAK_MATCH_MAX)).sum()),
                round(float(prox.mean()), 2) if len(prox) else 0,
                round(float(prox.max()), 2) if len(prox) else 0,
                round(float(prox.min()), 2) if len(prox) else 0,
            ],
        }
    )

    return no_match, weak_match, resumen


# =========================================================
# EXPORTACIÓN
# =========================================================
def export_single_workbook(sheets: Dict[str, pd.DataFrame], output_path: Path) -> Path:
    try:
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            for sheet_name, df in sheets.items():
                safe_sheet_name = str(sheet_name)[:31]
                df.to_excel(writer, sheet_name=safe_sheet_name, index=False)
        return output_path
    except PermissionError as e:
        raise PermissionError(explain_excel_permission_error(output_path)) from e


def export_single_sheet_excel(df: pd.DataFrame, output_path: Path, sheet_name: str = "SECOP_NoEstructurado") -> Path:
    try:
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
        return output_path
    except PermissionError as e:
        raise PermissionError(explain_excel_permission_error(output_path)) from e


# =========================================================
# MAIN
# =========================================================
def main():
    print("=" * 100)
    print("MERGE SECOP NO ESTRUCTURADO - VERSIÓN PRO (DOS ARCHIVOS XLSX)")
    print("=" * 100)

    if len(sys.argv) >= 2 and sys.argv[1].strip():
        input_folder = Path(sys.argv[1].strip().strip('"').strip("'"))
        if not input_folder.exists() or not input_folder.is_dir():
            raise FileNotFoundError(f"La carpeta de entrada no existe o no es válida: {input_folder}")
    else:
        input_folder = prompt_existing_folder(
            "📂 Ingresa la ruta de la carpeta donde están las 4 bases de datos: "
        )

    if len(sys.argv) >= 3 and sys.argv[2].strip():
        output_folder = Path(sys.argv[2].strip().strip('"').strip("'"))
        output_folder.mkdir(parents=True, exist_ok=True)
    else:
        output_folder = prompt_output_folder(
            "💾 Ingresa la ruta de la carpeta donde deseas guardar los outputs: "
        )

    print("\n🔎 Verificando archivos requeridos...")
    files = find_required_files(input_folder)
    for _, path in files.items():
        print(f"   - {path.name}")

    print("\n📥 Cargando bases...")
    df_url = read_excel_file(files["url_base"])
    df_contratos = read_excel_file(files["contratos"])
    df_procedimientos = read_excel_file(files["procedimientos"])
    df_estudios = read_excel_file(files["estudios"])

    print("✅ Bases cargadas correctamente.")
    print(f"   URL base: {df_url.shape}")
    print(f"   Contratos: {df_contratos.shape}")
    print(f"   Procedimientos: {df_procedimientos.shape}")
    print(f"   Estudios previos: {df_estudios.shape}")

    print("\n1️⃣ Uniendo Contratos_Extraidos + URLDocumento_NombresContratos_NombresSECOPProcedimiento...")
    contratos_url = merge_contratos_with_url(df_contratos, df_url)
    print(f"   Resultado Contratos_Extraidos_URL: {contratos_url.shape}")
    del df_contratos

    print("\n2️⃣ Uniendo secop_procedimiento_extraidos + URLDocumento_NombresContratos_NombresSECOPProcedimiento...")
    procedimientos_url = merge_procedimientos_with_url(df_procedimientos, df_url)
    print(f"   Resultado secop_procedimiento_extraidos_URL: {procedimientos_url.shape}")
    del df_procedimientos
    del df_url

    print("\n3️⃣ Uniendo Contratos_Extraidos_URL + secop_procedimiento_extraidos_URL por url (LEFT JOIN base = Contratos_Extraidos_URL)...")
    minutas_procedimientos = merge_minutas_and_procedimientos(contratos_url, procedimientos_url)
    print(f"   Resultado MinutasYProcedimientosSECOP: {minutas_procedimientos.shape}")
    print(f"   Filas esperadas según Contratos_Extraidos_URL: {len(contratos_url)}")
    print(f"   Filas obtenidas en MinutasYProcedimientosSECOP: {len(minutas_procedimientos)}")

    descripcion_col = detect_descripcion_column(minutas_procedimientos)

    print("\n4️⃣ Uniendo MinutasYProcedimientosSECOP + estudios_previos_extraidos con fuzzy LEFT JOIN...")
    print(f"   Columna usada en MinutasYProcedimientosSECOP: {descripcion_col}")
    print("   Columna usada en estudios_previos_extraidos: OBJETO")
    print(f"   Umbral de proximidad: {FUZZY_THRESHOLD}")

    secop_no_estructurado = fuzzy_left_merge_best_match(
        left_df=minutas_procedimientos,
        right_df=df_estudios,
        left_on=descripcion_col,
        right_on="OBJETO",
        threshold=FUZZY_THRESHOLD,
        proximity_col="Proximidad_Objeto_descripcion",
    )
    print(f"   Resultado SECOP_NoEstructurado: {secop_no_estructurado.shape}")

    if len(secop_no_estructurado) != len(minutas_procedimientos):
        raise RuntimeError(
            f"SECOP_NoEstructurado no conserva las filas de MinutasYProcedimientosSECOP. "
            f"Esperadas: {len(minutas_procedimientos)}, obtenidas: {len(secop_no_estructurado)}"
        )

    no_match_df, weak_match_df, resumen_df = build_matching_reports(secop_no_estructurado)

    print("\n5️⃣ Exportando los dos archivos finales...")
    consolidated_path = output_folder / OUTPUT_WORKBOOK_NAME
    final_sheet_path = output_folder / OUTPUT_FINAL_SHEET_NAME

    sheets = {
        "Contratos_Extraidos_URL": contratos_url,
        "secop_procedimiento_extraidos_URL": procedimientos_url,
        "MinutasYProcedimientosSECOP": minutas_procedimientos,
        "SECOP_NoEstructurado": secop_no_estructurado,
        "SECOP_NoMatch": no_match_df,
        "SECOP_MatchDebil": weak_match_df,
        "ResumenMatching": resumen_df,
    }

    export_single_workbook(sheets, consolidated_path)
    export_single_sheet_excel(secop_no_estructurado, final_sheet_path, sheet_name="SECOP_NoEstructurado")

    print("\n✅ Proceso terminado correctamente.")
    print("\nArchivos generados:")
    print(f"   - {consolidated_path}")
    print(f"   - {final_sheet_path}")

    print("\nHojas incluidas en el consolidado:")
    for name in sheets:
        print(f"   - {name}")

    print("\nParámetros usados:")
    print(f"   - FUZZY_THRESHOLD = {FUZZY_THRESHOLD}")
    print(f"   - WEAK_MATCH_MIN = {WEAK_MATCH_MIN}")
    print(f"   - WEAK_MATCH_MAX = {WEAK_MATCH_MAX}")
    print(f"   - rapidfuzz disponible = {_rapidfuzz_available()}")
    print(f"   - MAX_WORKERS = {MAX_WORKERS}")


if __name__ == "__main__":
    main()
