import re
import sys
import unicodedata
from pathlib import Path
from typing import Dict, List, Tuple, Optional

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

FUZZY_THRESHOLD = 80  # puedes cambiarlo
JOIN_ON_URL_HOW = "inner"  # recomendado: "inner"
OUTPUT_FILES = {
    "contratos_url": "Contratos_Extraidos_URL.xlsx",
    "procedimientos_url": "secop_procedimiento_extraidos_URL.xlsx",
    "minutas_procedimientos": "MinutasYProcedimientosSECOP.xlsx",
    "secop_no_estructurado": "SECOP_NoEstructurado.xlsx",
}


# =========================================================
# UTILIDADES GENERALES
# =========================================================
def normalize_spaces(text: str) -> str:
    text = "" if text is None else str(text)
    text = text.replace("\xa0", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n+", " ", text)
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


def read_excel_file(path: Path) -> pd.DataFrame:
    return pd.read_excel(path)


def validate_required_columns(df: pd.DataFrame, required_columns: List[str], df_name: str) -> None:
    missing = [col for col in required_columns if col not in df.columns]
    if missing:
        raise ValueError(
            f"La base '{df_name}' no contiene estas columnas requeridas: {missing}\n"
            f"Columnas disponibles: {list(df.columns)}"
        )


def make_unique_columns(columns: List[str]) -> List[str]:
    counts = {}
    result = []
    for col in columns:
        if col not in counts:
            counts[col] = 0
            result.append(col)
        else:
            counts[col] += 1
            result.append(f"{col}_{counts[col]}")
    return result


# =========================================================
# FUZZY MATCH
# =========================================================
def similarity_ratio(a: str, b: str) -> float:
    try:
        from rapidfuzz import fuzz
        ratio = float(fuzz.ratio(a, b))
        token_sort = float(fuzz.token_sort_ratio(a, b))
        token_set = float(fuzz.token_set_ratio(a, b))
        partial = float(fuzz.partial_ratio(a, b))

        # combinación conservadora para que sea rigurosa
        score = (0.35 * ratio) + (0.25 * token_sort) + (0.25 * token_set) + (0.15 * partial)
        return round(score, 2)
    except Exception:
        from difflib import SequenceMatcher
        ratio = SequenceMatcher(None, a, b).ratio() * 100
        return round(ratio, 2)


def build_candidate_index(estudios_df: pd.DataFrame, object_col: str) -> Dict[str, List[int]]:
    index = {}
    for i, text in estudios_df[object_col].fillna("").astype(str).items():
        norm = normalize_for_match(text)
        tokens = [t for t in norm.split() if len(t) >= 4]
        key_tokens = tokens[:4] if tokens else []
        for token in key_tokens:
            index.setdefault(token, []).append(i)
    return index


def get_candidate_rows(
    text: str,
    candidate_index: Dict[str, List[int]],
    total_indices: List[int]
) -> List[int]:
    norm = normalize_for_match(text)
    tokens = [t for t in norm.split() if len(t) >= 4]

    candidates = set()
    for token in tokens[:8]:
        for idx in candidate_index.get(token, []):
            candidates.add(idx)

    # si no encontró candidatos por bloqueo, compara contra todos
    if not candidates:
        return total_indices

    return sorted(candidates)


def fuzzy_left_merge(
    left_df: pd.DataFrame,
    right_df: pd.DataFrame,
    left_on: str,
    right_on: str,
    threshold: float = 80.0,
    proximity_col: str = "Proximidad_Objeto_descripcion",
) -> pd.DataFrame:
    left = left_df.copy()
    right = right_df.copy()

    validate_required_columns(left, [left_on], "MinutasYProcedimientosSECOP")
    validate_required_columns(right, [right_on], "estudios_previos_extraidos")

    # preparar columnas del lado derecho para evitar duplicados
    right_cols_original = list(right.columns)
    right_cols_renamed = []
    for col in right_cols_original:
        if col == right_on:
            right_cols_renamed.append(col)
        elif col in left.columns:
            right_cols_renamed.append(f"{col}_estudios")
        else:
            right_cols_renamed.append(col)

    right = right.copy()
    right.columns = right_cols_renamed

    right_match_col = right_on
    if right_on not in right.columns:
        # seguridad extra, aunque no debería pasar
        for c in right.columns:
            if c.startswith(right_on):
                right_match_col = c
                break

    left["_match_text_norm"] = left[left_on].fillna("").astype(str).map(normalize_for_match)
    right["_match_text_norm"] = right[right_match_col].fillna("").astype(str).map(normalize_for_match)

    candidate_index = build_candidate_index(right, right_match_col)
    all_right_indices = list(right.index)

    best_rows = []
    for _, left_row in left.iterrows():
        left_text_raw = left_row[left_on]
        left_text_norm = left_row["_match_text_norm"]

        best_score = 0.0
        best_idx = None

        if left_text_norm:
            candidate_rows = get_candidate_rows(left_text_norm, candidate_index, all_right_indices)

            for idx in candidate_rows:
                right_text_norm = right.at[idx, "_match_text_norm"]
                if not right_text_norm:
                    continue
                score = similarity_ratio(left_text_norm, right_text_norm)
                if score > best_score:
                    best_score = score
                    best_idx = idx

        merged_row = left_row.to_dict()
        merged_row[proximity_col] = round(best_score, 2)

        if best_idx is not None and best_score >= threshold:
            right_data = right.loc[best_idx].to_dict()
            right_data.pop("_match_text_norm", None)
            for key, value in right_data.items():
                if key not in merged_row:
                    merged_row[key] = value
                elif key == right_match_col:
                    merged_row[f"{key}_estudios"] = value
                else:
                    merged_row[key] = value
        else:
            for col in right.columns:
                if col != "_match_text_norm" and col not in merged_row:
                    merged_row[col] = pd.NA

        best_rows.append(merged_row)

    result = pd.DataFrame(best_rows)
    result.drop(columns=["_match_text_norm"], errors="ignore", inplace=True)
    return result


# =========================================================
# JOINS SOLICITADOS
# =========================================================
def merge_contratos_with_url(
    contratos_df: pd.DataFrame,
    url_df: pd.DataFrame,
) -> pd.DataFrame:
    validate_required_columns(contratos_df, ["archivo"], "Contratos_Extraidos")
    validate_required_columns(url_df, ["contratos_pdf_paths", "url"], "URLDocumento_NombresContratos_NombresSECOPProcedimiento")

    left = contratos_df.copy()
    right = url_df[["contratos_pdf_paths", "url"]].copy()

    left["_key_archivo"] = left["archivo"].map(basename_from_pathlike)
    right["_key_archivo"] = right["contratos_pdf_paths"].map(basename_from_pathlike)

    # por si hay duplicados en la tabla de URLs, conservar el primero
    right = right.drop_duplicates(subset=["_key_archivo"], keep="first")

    merged = left.merge(
        right[["_key_archivo", "url"]],
        on="_key_archivo",
        how="left",
        validate="m:1",
    )

    merged.drop(columns=["_key_archivo"], inplace=True, errors="ignore")
    return merged


def merge_procedimientos_with_url(
    procedimientos_df: pd.DataFrame,
    url_df: pd.DataFrame,
) -> pd.DataFrame:
    validate_required_columns(procedimientos_df, ["archivo_pdf"], "secop_procedimiento_extraidos")
    validate_required_columns(url_df, ["Secop_procedimientos_pdf_path", "url"], "URLDocumento_NombresContratos_NombresSECOPProcedimiento")

    left = procedimientos_df.copy()
    right = url_df[["Secop_procedimientos_pdf_path", "url"]].copy()

    left["_key_archivo"] = left["archivo_pdf"].map(basename_from_pathlike)
    right["_key_archivo"] = right["Secop_procedimientos_pdf_path"].map(basename_from_pathlike)

    right = right.drop_duplicates(subset=["_key_archivo"], keep="first")

    merged = left.merge(
        right[["_key_archivo", "url"]],
        on="_key_archivo",
        how="left",
        validate="m:1",
    )

    merged.drop(columns=["_key_archivo"], inplace=True, errors="ignore")
    return merged


def merge_minutas_and_procedimientos(
    contratos_url_df: pd.DataFrame,
    procedimientos_url_df: pd.DataFrame,
    how: str = "inner",
) -> pd.DataFrame:
    validate_required_columns(contratos_url_df, ["url"], "Contratos_Extraidos_URL")
    validate_required_columns(procedimientos_url_df, ["url"], "secop_procedimiento_extraidos_URL")

    left = contratos_url_df.copy()
    right = procedimientos_url_df.copy()

    # renombrar solo columnas que colisionen, excepto url
    rename_map = {}
    for col in right.columns:
        if col != "url" and col in left.columns:
            rename_map[col] = f"{col}_procedimiento"
    right = right.rename(columns=rename_map)

    merged = left.merge(
        right,
        on="url",
        how=how,
    )
    return merged


def export_outputs(outputs: Dict[str, pd.DataFrame], output_folder: Path) -> Dict[str, Path]:
    exported_paths = {}
    for key, df in outputs.items():
        output_path = output_folder / OUTPUT_FILES[key]
        df.to_excel(output_path, index=False)
        exported_paths[key] = output_path
    return exported_paths


# =========================================================
# MAIN
# =========================================================
def main():
    print("=" * 80)
    print("MERGE DE BASES SECOP NO ESTRUCTURADAS")
    print("=" * 80)

    # Permitir pasar rutas por argumentos o pedirlas por consola
    if len(sys.argv) >= 2 and sys.argv[1].strip():
        input_folder = Path(sys.argv[1].strip().strip('"').strip("'"))
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
    for key, path in files.items():
        print(f"   - {key}: {path.name}")

    print("\n📥 Cargando bases de datos...")
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

    print("\n2️⃣ Uniendo secop_procedimiento_extraidos + URLDocumento_NombresContratos_NombresSECOPProcedimiento...")
    procedimientos_url = merge_procedimientos_with_url(df_procedimientos, df_url)
    print(f"   Resultado secop_procedimiento_extraidos_URL: {procedimientos_url.shape}")

    print("\n3️⃣ Uniendo Contratos_Extraidos_URL + secop_procedimiento_extraidos_URL por url...")
    minutas_procedimientos = merge_minutas_and_procedimientos(
        contratos_url,
        procedimientos_url,
        how=JOIN_ON_URL_HOW,
    )
    print(f"   Resultado MinutasYProcedimientosSECOP: {minutas_procedimientos.shape}")

    print("\n4️⃣ Uniendo MinutasYProcedimientosSECOP + estudios_previos_extraidos con fuzzy merge...")
    # La columna real en procedimientos extraídos es 'descripcion'
    left_key = "descripcion"
    right_key = "OBJETO"

    secop_no_estructurado = fuzzy_left_merge(
        left_df=minutas_procedimientos,
        right_df=df_estudios,
        left_on=left_key,
        right_on=right_key,
        threshold=FUZZY_THRESHOLD,
        proximity_col="Proximidad_Objeto_descripcion",
    )
    print(f"   Resultado SECOP_NoEstructurado: {secop_no_estructurado.shape}")

    print("\n5️⃣ Exportando archivos...")
    exported = export_outputs(
        outputs={
            "contratos_url": contratos_url,
            "procedimientos_url": procedimientos_url,
            "minutas_procedimientos": minutas_procedimientos,
            "secop_no_estructurado": secop_no_estructurado,
        },
        output_folder=output_folder,
    )

    print("\n✅ Proceso terminado correctamente.")
    print("\nArchivos exportados:")
    for _, path in exported.items():
        print(f"   - {path}")

    print("\nParámetros usados:")
    print(f"   - FUZZY_THRESHOLD = {FUZZY_THRESHOLD}")
    print(f"   - JOIN_ON_URL_HOW = '{JOIN_ON_URL_HOW}'")


if __name__ == "__main__":
    main()
