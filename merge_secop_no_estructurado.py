# -*- coding: utf-8 -*-
import math
import os
import re
import sys
import unicodedata
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import pandas as pd

INPUT_FILES = {
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
CONSOLIDATED_OBLIGATIONS_COLUMN = "obligaciones específicas consolidadas"
MATCH_MIN_YEAR = 2026

def normalize_spaces(text: str) -> str:
    text = "" if text is None else str(text)
    text = text.replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\s*\n\s*", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()

def strip_accents(text: str) -> str:
    text = "" if text is None else str(text)
    return "".join(ch for ch in unicodedata.normalize("NFD", text) if unicodedata.category(ch) != "Mn")

def normalize_for_match(text: str) -> str:
    text = normalize_spaces(text)
    text = strip_accents(text).lower()
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text

def is_general_or_empty_obligation(text: str) -> bool:
    normalized = normalize_for_match(text)
    if not normalized:
        return True
    general_patterns = [
        r".*\bprevistas\s+en\s+los\s+estudios\s+previos\b.*",
    ]
    return any(re.search(pattern, normalized) for pattern in general_patterns)

def build_consolidated_obligations_column(
    df: pd.DataFrame,
    source_col_a: str = "obligaciones_especificas",
    source_col_b: str = "OBLIGACIONES ESPECÍFICAS DEL CONTRATISTA",
    output_col: str = CONSOLIDATED_OBLIGATIONS_COLUMN,
) -> pd.DataFrame:
    validate_required_columns(df, [source_col_a, source_col_b], "SECOP_NoEstructurado")
    output = df.copy()
    values_a = output[source_col_a].fillna("").astype(str).map(normalize_spaces)
    values_b = output[source_col_b].fillna("").astype(str).map(normalize_spaces)
    consolidated = []
    for value_a, value_b in zip(values_a, values_b):
        empty_a = (value_a == "")
        empty_b = (value_b == "")
        if empty_a and empty_b:
            consolidated.append("")
        elif empty_a:
            consolidated.append(value_b)
        elif empty_b:
            consolidated.append(value_a)
        elif is_general_or_empty_obligation(value_a):
            consolidated.append(value_b)
        else:
            consolidated.append(value_a)
    output[output_col] = consolidated
    return output

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
    found, missing = {}, []
    for key, filename in INPUT_FILES.items():
        full_path = input_folder / filename
        if full_path.exists():
            found[key] = full_path
        else:
            missing.append(filename)
    if missing:
        raise FileNotFoundError("No se encontraron los siguientes archivos en la carpeta indicada:\n- " + "\n- ".join(missing))
    return found

def explain_excel_permission_error(path: Path) -> str:
    return (
        f"No se pudo abrir o guardar el archivo:\n{path}\n\n"
        "Posibles causas:\n1. El archivo está abierto en Excel.\n2. OneDrive lo está sincronizando o bloqueando.\n3. No tienes permisos sobre esa carpeta.\n\n"
        "Prueba esto:\n- Cierra el archivo Excel.\n- Pausa OneDrive temporalmente.\n- Mueve los archivos a una carpeta local simple, por ejemplo: C:\\SECOP_Merge\\\n- Vuelve a ejecutar el script."
    )

def read_excel_file(path: Path) -> pd.DataFrame:
    try:
        return pd.read_excel(path)
    except PermissionError as e:
        raise PermissionError(explain_excel_permission_error(path)) from e
    except Exception as e:
        raise RuntimeError(f"Error leyendo el archivo Excel '{path}': {e}") from e

def filter_estudios_by_min_year(df: pd.DataFrame, min_year: int = MATCH_MIN_YEAR) -> pd.DataFrame:
    validate_required_columns(df, ["AÑO"], "estudios_previos_extraidos")
    year_values = pd.to_numeric(df["AÑO"], errors="coerce")
    return df.loc[year_values >= min_year].copy()

def validate_required_columns(df: pd.DataFrame, required_columns: List[str], df_name: str) -> None:
    missing = [col for col in required_columns if col not in df.columns]
    if missing:
        raise ValueError(f"La base '{df_name}' no contiene estas columnas requeridas: {missing}\nColumnas disponibles: {list(df.columns)}")

def detect_descripcion_column(df: pd.DataFrame) -> str:
    for c in ["descripción", "descripcion", "Descripción", "Descripcion"]:
        if c in df.columns:
            return c
    raise ValueError(f"No encontré la columna descripción/descripcion. Columnas disponibles: {list(df.columns)}")

def ensure_unique_right_columns(left_df: pd.DataFrame, right_df: pd.DataFrame, join_key: str) -> pd.DataFrame:
    right = right_df.copy()
    rename_map = {c: f"{c}_estudios" for c in right.columns if c != join_key and c in left_df.columns}
    return right.rename(columns=rename_map) if rename_map else right

def normalize_contract_key(value: str) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = normalize_spaces(str(value)).upper()
    text = strip_accents(text)
    text = re.sub(r"[_/\\|]+", "-", text)
    text = re.sub(r"\s*-\s*", "-", text)
    text = re.sub(r"\s+", "-", text)
    text = re.sub(r"[^A-Z0-9-]", "-", text)
    text = re.sub(r"-{2,}", "-", text).strip("-")
    m = re.search(r"([A-Z]+)-?(\d{1,})-?(20\d{2})", text)
    if m:
        return f"{m.group(1)}-{m.group(2).zfill(3)}-{m.group(3)}"
    parts = [p for p in text.split("-") if p]
    if len(parts) >= 3:
        prefix = parts[0]
        number = re.sub(r"\D", "", parts[1])
        year = re.sub(r"\D", "", parts[2])
        if prefix and number and len(year) == 4:
            return f"{prefix}-{number.zfill(3)}-{year}"
    return text

def build_concatenado_id(df: pd.DataFrame, columns: List[str], df_name: str) -> pd.Series:
    validate_required_columns(df, columns, df_name)
    return (
        df[columns]
        .fillna("")
        .astype(str)
        .applymap(normalize_spaces)
        .agg(" ".join, axis=1)
        .map(normalize_spaces)
    )

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
        score = 0.40*float(fuzz.ratio(a,b)) + 0.25*float(fuzz.token_sort_ratio(a,b)) + 0.25*float(fuzz.token_set_ratio(a,b)) + 0.10*float(fuzz.partial_ratio(a,b))
        return round(score,2)
    except Exception:
        from difflib import SequenceMatcher
        return round(SequenceMatcher(None,a,b).ratio()*100,2)

def best_match_one(query_norm: str, choices_norm: List[str]) -> Tuple[Optional[int], float]:
    if not query_norm:
        return None, 0.0
    try:
        from rapidfuzz import process, fuzz
        scored = process.extractOne(query_norm, choices_norm, scorer=fuzz.token_set_ratio)
        if scored is None:
            return None, 0.0
        matched_text, _, idx = scored
        return int(idx), similarity_score(query_norm, matched_text)
    except Exception:
        best_idx, best_score = None, 0.0
        for idx, choice in enumerate(choices_norm):
            score = similarity_score(query_norm, choice)
            if score > best_score:
                best_score, best_idx = score, idx
        return best_idx, round(best_score,2)

def _best_match_chunk(chunk_queries: List[str], choices_norm: List[str]) -> Dict[str, Tuple[Optional[int], float]]:
    return {q: best_match_one(q, choices_norm) for q in chunk_queries}

def build_best_match_map_parallel(unique_left_values: List[str], right_choices_norm: List[str]) -> Dict[str, Tuple[Optional[int], float]]:
    if not unique_left_values:
        return {}
    n = len(unique_left_values)
    workers = min(MAX_WORKERS, n)
    if workers <= 1 or n < 200:
        result = {}
        for i, q in enumerate(unique_left_values, start=1):
            result[q] = best_match_one(q, right_choices_norm)
            if i % 250 == 0 or i == n:
                print(f"   Progreso fuzzy: {i}/{n}")
        return result
    chunk_size = math.ceil(n / workers)
    chunks = [unique_left_values[i:i + chunk_size] for i in range(0, n, chunk_size)]
    result, done = {}, 0
    with ThreadPoolExecutor(max_workers=workers) as executor:
        futures = [executor.submit(_best_match_chunk, chunk, right_choices_norm) for chunk in chunks]
        for future in as_completed(futures):
            partial = future.result(); result.update(partial); done += len(partial); print(f"   Progreso fuzzy: {done}/{n}")
    return result

def fuzzy_left_merge_best_match(
    left_df: pd.DataFrame,
    right_df: pd.DataFrame,
    left_on: str,
    right_on: str,
    threshold: float = 80.0,
    proximity_col: str = "Proximidad_Objeto_descripcion",
    left_year_col: str = "anio_proceso",
    right_year_col: str = "AÑO",
    min_year: int = MATCH_MIN_YEAR,
) -> pd.DataFrame:
    validate_required_columns(left_df, [left_on, left_year_col], "MinutasYProcedimientosSECOP")
    validate_required_columns(right_df, [right_on, right_year_col], "estudios_previos_extraidos")

    left = left_df.copy()
    left_year_values = pd.to_numeric(left[left_year_col], errors="coerce")
    left["_eligible_year_match"] = left_year_values >= min_year
    eligible_left_count = int(left["_eligible_year_match"].sum())

    right_year_values = pd.to_numeric(right_df[right_year_col], errors="coerce")
    right_filtered = right_df.loc[right_year_values >= min_year].copy()
    right = ensure_unique_right_columns(left_df, right_filtered, join_key=right_on); right_key = right_on

    left["_left_match_norm"] = left[left_on].fillna("").astype(str).map(normalize_for_match)
    right["_right_match_norm"] = right[right_key].fillna("").astype(str).map(normalize_for_match)
    right_valid = right[right["_right_match_norm"] != ""].copy().reset_index(drop=True)
    right_columns_to_add = [c for c in right.columns if c != "_right_match_norm"]

    if right_valid.empty or eligible_left_count == 0:
        result = left.copy(); result[proximity_col] = 0.0
        for col in right_columns_to_add:
            if col not in result.columns:
                result[col] = pd.NA
        return result.drop(columns=["_left_match_norm", "_eligible_year_match"], errors="ignore")

    eligible_left = left.loc[left["_eligible_year_match"]]
    unique_left_values = eligible_left["_left_match_norm"].fillna("").astype(str).drop_duplicates().tolist()
    right_choices_norm = right_valid["_right_match_norm"].tolist()
    print(f"🔎 Iniciando fuzzy match sobre {len(unique_left_values)} descripciones únicas con {left_year_col}>={min_year}...")
    print(f"   - registros MinutasYProcedimientosSECOP elegibles: {eligible_left_count}/{len(left)}")
    print(f"   - registros estudios_previos_extraidos elegibles: {len(right_valid)}/{len(right_df)}")
    print(f"   - rapidfuzz disponible: {_rapidfuzz_available()}")
    print(f"   - hilos usados: {min(MAX_WORKERS, max(1, len(unique_left_values)))}")
    best_match_map = build_best_match_map_parallel(unique_left_values, right_choices_norm)
    output_rows = []
    for _, row in left.iterrows():
        row_dict = row.to_dict(); query_norm = row_dict.get("_left_match_norm", "")
        is_year_eligible = bool(row_dict.get("_eligible_year_match", False))
        best_idx, best_score = best_match_map.get(query_norm, (None, 0.0)) if is_year_eligible else (None, 0.0)
        row_dict[proximity_col] = round(best_score, 2)
        if best_idx is not None and best_score >= threshold:
            matched = right_valid.iloc[best_idx].to_dict(); matched.pop("_right_match_norm", None)
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
    result = pd.DataFrame(output_rows).drop(columns=["_left_match_norm", "_eligible_year_match"], errors="ignore")
    if len(result) != len(left_df):
        raise RuntimeError(f"El fuzzy left join alteró el número de filas. Esperadas: {len(left_df)}, obtenidas: {len(result)}")
    return result

def merge_procedimientos_with_contratos(procedimientos_df: pd.DataFrame, contratos_df: pd.DataFrame) -> pd.DataFrame:
    validate_required_columns(procedimientos_df, ["numero_proceso"], "secop_procedimiento_extraidos")
    validate_required_columns(contratos_df, ["numero_contrato"], "Contratos_Extraidos")
    left = procedimientos_df.copy(); right = contratos_df.copy()
    left["_row_id"] = range(len(left))
    left["_join_key"] = left["numero_proceso"].map(normalize_contract_key)
    right["_join_key"] = right["numero_contrato"].map(normalize_contract_key)
    rename_map = {c: f"{c}_contrato" for c in right.columns if c != "_join_key" and c in left.columns}
    if rename_map: right = right.rename(columns=rename_map)
    right = right.drop_duplicates(subset=["_join_key"], keep="first")
    merged = left.merge(right, on="_join_key", how="left", validate="m:1").sort_values("_row_id").reset_index(drop=True)
    if len(merged) != len(procedimientos_df):
        raise RuntimeError(f"El paso 1 alteró el número de filas de secop_procedimiento_extraidos. Esperadas: {len(procedimientos_df)}, obtenidas: {len(merged)}")
    return merged

def build_matching_reports(df: pd.DataFrame, proximity_col: str = "Proximidad_Objeto_descripcion") -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    prox = pd.to_numeric(df[proximity_col], errors="coerce").fillna(0)
    no_match = df.loc[prox < FUZZY_THRESHOLD].copy(); weak_match = df.loc[(prox >= WEAK_MATCH_MIN) & (prox <= WEAK_MATCH_MAX)].copy()
    resumen = pd.DataFrame({"Métrica":["Filas MinutasYProcedimientosSECOP","Filas SECOP_NoEstructurado",f"Matches aceptados (>= {FUZZY_THRESHOLD})",f"No match (< {FUZZY_THRESHOLD})",f"Matches débiles ({WEAK_MATCH_MIN} a {WEAK_MATCH_MAX})","Score promedio","Score máximo","Score mínimo"],"Valor":[len(df),len(df),int((prox >= FUZZY_THRESHOLD).sum()),int((prox < FUZZY_THRESHOLD).sum()),int(((prox >= WEAK_MATCH_MIN) & (prox <= WEAK_MATCH_MAX)).sum()),round(float(prox.mean()),2) if len(prox) else 0,round(float(prox.max()),2) if len(prox) else 0,round(float(prox.min()),2) if len(prox) else 0]})
    return no_match, weak_match, resumen



def build_data_quality_report(df: pd.DataFrame) -> pd.DataFrame:
    total_rows = len(df)
    total_cols = len(df.columns)
    total_cells = total_rows * total_cols if total_rows and total_cols else 0

    non_null_cells = int(df.notna().sum().sum()) if total_cells else 0
    empty_after_trim = int(df.applymap(lambda x: isinstance(x, str) and normalize_spaces(x) == "").sum().sum()) if total_cells else 0
    completeness = round(((non_null_cells - empty_after_trim) / total_cells) * 100, 2) if total_cells else 0.0

    duplicated_rows = int(df.duplicated().sum()) if total_rows else 0
    uniqueness = round(((total_rows - duplicated_rows) / total_rows) * 100, 2) if total_rows else 0.0

    proximity = pd.to_numeric(df.get("Proximidad_Objeto_descripcion", pd.Series(dtype=float)), errors="coerce")
    exactitud = round(float((proximity >= FUZZY_THRESHOLD).mean() * 100), 2) if len(proximity) else 0.0
    precision = round(float(proximity.mean()), 2) if len(proximity) else 0.0
    confiabilidad = round(float((proximity >= WEAK_MATCH_MIN).mean() * 100), 2) if len(proximity) else 0.0

    normalized_rows = df.fillna("").astype(str).applymap(normalize_spaces) if total_rows else df
    consistency = round(float((normalized_rows.nunique(dropna=False) <= 1).mean() * 100), 2) if total_cols else 0.0

    key_columns = [
        "numero_proceso",
        "numero_contrato",
        "nombre_contratista",
        "objeto",
        CONSOLIDATED_OBLIGATIONS_COLUMN,
    ]
    existing_key_columns = [c for c in key_columns if c in df.columns]
    valid_rows = 0
    if existing_key_columns and total_rows:
        valid_rows = int(df[existing_key_columns].fillna("").astype(str).applymap(normalize_spaces).ne("").all(axis=1).sum())
    validez = round((valid_rows / total_rows) * 100, 2) if total_rows else 0.0

    integridad = round(((1 - (df.isna().any(axis=1).sum() / total_rows)) * 100), 2) if total_rows else 0.0

    quality_table = pd.DataFrame(
        {
            "Dimensión": [
                "Completitud",
                "Exactitud",
                "Consistencia",
                "Validez",
                "Unicidad",
                "Integridad",
                "Precisión",
                "Confiabilidad",
            ],
            "Indicador (%)": [
                completeness,
                exactitud,
                consistency,
                validez,
                uniqueness,
                integridad,
                precision,
                confiabilidad,
            ],
            "Descripción": [
                "Porcentaje de celdas diligenciadas (no nulas y no vacías).",
                f"Registros con score de proximidad >= {FUZZY_THRESHOLD}.",
                "Columnas con valores homogéneos por normalización textual.",
                "Registros con campos clave diligenciados.",
                "Registros no duplicados sobre el total.",
                "Registros sin valores nulos en ninguna columna.",
                "Promedio del score de proximidad del matching.",
                f"Registros con score de proximidad >= {WEAK_MATCH_MIN}.",
            ],
        }
    )

    meta = pd.DataFrame(
        {
            "Métrica": ["Total filas", "Total columnas", "Total celdas", "Registros duplicados"],
            "Valor": [total_rows, total_cols, total_cells, duplicated_rows],
        }
    )

    separator = pd.DataFrame({"Dimensión": [""], "Indicador (%)": [""], "Descripción": [""], "Métrica": [""], "Valor": [""]})
    quality_table_expanded = quality_table.copy()
    quality_table_expanded["Métrica"] = ""
    quality_table_expanded["Valor"] = ""
    meta_expanded = meta.copy()
    meta_expanded["Dimensión"] = ""
    meta_expanded["Indicador (%)"] = ""
    meta_expanded["Descripción"] = ""

    return pd.concat([quality_table_expanded, separator, meta_expanded], ignore_index=True)


def export_single_workbook(sheets: Dict[str, pd.DataFrame], output_path: Path) -> Path:
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=str(sheet_name)[:31], index=False)
    return output_path

def export_single_sheet_excel(df: pd.DataFrame, output_path: Path, sheet_name: str = "SECOP_NoEstructurado") -> Path:
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    return output_path

def main():
    print("=" * 100)
    print("MERGE SECOP NO ESTRUCTURADO - VERSIÓN PRO (DOS ARCHIVOS XLSX)")
    print("=" * 100)
    input_folder = Path(sys.argv[1].strip().strip('"').strip("'")) if len(sys.argv) >= 2 and sys.argv[1].strip() else prompt_existing_folder("📂 Ingresa la ruta de la carpeta donde están las 3 bases de datos: ")
    if not input_folder.exists() or not input_folder.is_dir():
        raise FileNotFoundError(f"La carpeta de entrada no existe o no es válida: {input_folder}")
    output_folder = Path(sys.argv[2].strip().strip('"').strip("'")) if len(sys.argv) >= 3 and sys.argv[2].strip() else prompt_output_folder("💾 Ingresa la ruta de la carpeta donde deseas guardar los outputs: ")
    output_folder.mkdir(parents=True, exist_ok=True)
    print("\n🔎 Verificando archivos requeridos...")
    files = find_required_files(input_folder)
    for _, path in files.items(): print(f"   - {path.name}")
    print("\n📥 Cargando bases...")
    df_contratos = read_excel_file(files["contratos"]); df_procedimientos = read_excel_file(files["procedimientos"]); df_estudios = read_excel_file(files["estudios"])
    print("✅ Bases cargadas correctamente.")
    print(f"   Contratos: {df_contratos.shape}")
    print(f"   Procedimientos: {df_procedimientos.shape}")
    print(f"   Estudios previos: {df_estudios.shape}")
    estudios_total_rows = len(df_estudios)
    df_estudios = filter_estudios_by_min_year(df_estudios, MATCH_MIN_YEAR)
    print(f"   Estudios previos filtrados AÑO>={MATCH_MIN_YEAR}: {df_estudios.shape} (de {estudios_total_rows} registros)")
    print("\n1️⃣ Uniendo secop_procedimiento_extraidos + Contratos_Extraidos por numero_proceso / numero_contrato...")
    minutas_procedimientos = merge_procedimientos_with_contratos(df_procedimientos, df_contratos)
    print(f"   Resultado MinutasYProcedimientosSECOP: {minutas_procedimientos.shape}")
    print("\n2️⃣ Uniendo MinutasYProcedimientosSECOP + estudios_previos_extraidos con fuzzy LEFT JOIN por descripcion / OBJETO...")
    secop_no_estructurado = fuzzy_left_merge_best_match(
        minutas_procedimientos,
        df_estudios,
        "descripcion",
        "OBJETO",
        FUZZY_THRESHOLD,
        "Proximidad_Objeto_descripcion",
    )
    secop_no_estructurado = build_consolidated_obligations_column(secop_no_estructurado)
    no_match_df, weak_match_df, resumen_df = build_matching_reports(secop_no_estructurado)
    calidad_df = build_data_quality_report(secop_no_estructurado)
    print("\n3️⃣ Exportando los dos archivos finales...")
    consolidated_path = output_folder / OUTPUT_WORKBOOK_NAME
    final_sheet_path = output_folder / OUTPUT_FINAL_SHEET_NAME
    sheets = {"Contratos_Extraidos_URL": df_contratos, "secop_procedimiento_extraidos_URL": df_procedimientos, "MinutasYProcedimientosSECOP": minutas_procedimientos, "SECOP_NoEstructurado": secop_no_estructurado, "SECOP_NoMatch": no_match_df, "SECOP_MatchDebil": weak_match_df, "ResumenMatching": resumen_df, "CalidadInformacion": calidad_df}
    export_single_workbook(sheets, consolidated_path)
    export_single_sheet_excel(secop_no_estructurado, final_sheet_path, sheet_name="SECOP_NoEstructurado")
    print("\n✅ Proceso terminado correctamente.")
    print(f"\nArchivos generados:\n   - {consolidated_path}\n   - {final_sheet_path}")

if __name__ == "__main__":
    main()
