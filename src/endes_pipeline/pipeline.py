import io
import json
import logging
import re
import tempfile
import zipfile
from collections import OrderedDict
from dataclasses import dataclass
from datetime import datetime, UTC
from pathlib import Path
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd
from dbfread import DBF

from .analysis import build_analysis_settings, run_analysis


MISSING_TOKENS = {"", " ", "NA", "N/A", "NAN", "NONE", "*"}
CODED_MISSING_TOKENS = {"8", "9"}

PHQ_COLS = [f"QS700{letter}" for letter in "ABCDEFGHI"]
BP_COLS = ["QS903S", "QS903D", "QS905S", "QS905D"]
SYSTOLIC_COLS = ["QS903S", "QS905S"]
DIASTOLIC_COLS = ["QS903D", "QS905D"]
ANTHRO_COLS = ["QS900", "QS901", "QS907"]
ALCOHOL_SCREEN_COLS = ["QS713", "QS714", "QS715", "QS716", "QS717", "QS719", "QS720"]
DIET_UNIT_COLS = ["QS213U", "QS214U", "QS215U", "QS216U", "QS217U", "QS218U", "QS219U", "QS220U"]
DIET_MAGNITUDE_COLS = ["QS213C", "QS214C", "QS215C", "QS216C", "QS217C", "QS218C", "QS219C", "QS220CV", "QS220CC"]
DIET_COLS = [
    "QS213U",
    "QS213C",
    "QS214U",
    "QS214C",
    "QS215U",
    "QS215C",
    "QS216U",
    "QS216C",
    "QS217U",
    "QS217C",
    "QS218U",
    "QS218C",
    "QS219U",
    "QS219C",
    "QS220U",
    "QS220CV",
    "QS220CC",
]
DIET_SENTINEL_COLS = ["QS214C", "QS216C", "QS218C", "QS220CV"]
HOUSEHOLD_COLS = ["HV001", "HV022", "HV023", "HV024", "HV025", "HV005"]

CSALUD_BASE_COLS = [
    "HHID",
    "QSNUMERO",
    "QHCLUSTER",
    "QSRESULT",
    "QS23",
    "QSSEXO",
    "QS25N",
    "QS900",
    "QS901",
    "QS907",
    "QS102",
    "QS103U",
    "QS103C",
    "QS104",
    "QS106",
    "QS109",
    "QS201",
    "QS209",
    "QS702",
    "QS707",
    "QS709",
    "QS710",
    "QS711",
    *ALCOHOL_SCREEN_COLS,
    *DIET_COLS,
    *PHQ_COLS,
    *BP_COLS,
]

# Altitud y trazabilidad geografica desde RECH0. HV040 es la altitud del
# cluster/hogar en metros (fuente analitica primaria). UBIGEO/CODCCPP/NOMCCPP se
# conservan solo para trazabilidad y auditoria externa por UBIGEO+CODCCPP contra
# INEI 2017. LATITUDY/LONGITUDX se cargan solo para QC (en 2019 vienen enmascaradas).
ALTITUDE_TRACE_COLS = ["HV040", "UBIGEO", "CODCCPP", "NOMCCPP", "LATITUDY", "LONGITUDX"]
RECH0_COLS = ["HHID", *HOUSEHOLD_COLS, *ALTITUDE_TRACE_COLS]
RECH23_COLS = ["HHID", "HV270"]
REC42_COLS = ["CASEID", "V454"]
RE223132_COLS = ["CASEID", "V222"]

EXPLICIT_CODED_MISSING_COLS = sorted(
    set(
        PHQ_COLS
        + [
            "QS102",
            "QS104",
            "QS106",
            "QS109",
            "QS201",
            "QS209",
            "QS702",
            "QS707",
            "QS709",
            "QS710",
            "QS711",
            *DIET_UNIT_COLS,
            *ALCOHOL_SCREEN_COLS,
        ]
    )
)

NUMERIC_PARSE_COLS = sorted(
    set(
        [
            "QHCLUSTER",
            "QSRESULT",
            "QS23",
            "QSSEXO",
            "QS25N",
            "QS900",
            "QS901",
            "QS907",
            "QS102",
            "QS103U",
            "QS103C",
            "QS104",
            "QS106",
            "QS109",
            "QS201",
            "QS209",
            "QS702",
            "QS707",
            "QS709",
            "QS710",
            "QS711",
            *DIET_COLS,
            "PESO15_AMA",
            "PESO15_AMAS",
            "HV001",
            "HV022",
            "HV023",
            "HV024",
            "HV025",
            "HV005",
            "HV270",
            "HV040",
            "V454",
            "V222",
            *PHQ_COLS,
            *BP_COLS,
            *ALCOHOL_SCREEN_COLS,
        ]
    )
)

INTEGER_CODE_COLS = sorted(
    set(
        [
            "QHCLUSTER",
            "QSRESULT",
            "QS23",
            "QSSEXO",
            "QS25N",
            "QS102",
            "QS103U",
            "QS104",
            "QS106",
            "QS109",
            "QS201",
            "QS209",
            "QS702",
            "QS707",
            "QS709",
            "QS710",
            "QS711",
            *DIET_UNIT_COLS,
            "QS220CC",
            "HV001",
            "HV022",
            "HV023",
            "HV024",
            "HV025",
            "HV270",
            "V454",
            "V222",
            *PHQ_COLS,
            *ALCOHOL_SCREEN_COLS,
        ]
    )
)


@dataclass
class PipelineConfig:
    raw_archives_dir: Path
    interim_dir: Path
    output_dir: Path
    qc_dir: Path
    log_file: Path
    years: List[int]
    combined_year_divisor: float
    dataset_base_name: str
    write_csv: bool
    write_parquet: bool
    run_mice: bool
    run_vif: bool
    run_r_bridge: bool
    analysis_settings: Dict[str, object]


@dataclass
class YearArtifacts:
    final_df: pd.DataFrame
    flow_counts: Dict[str, int]
    source_map: Dict[str, str]
    missing_column_rows: List[Dict[str, str]]
    reproductive_summary: Dict[str, object]
    bp_summary: Dict[str, object]
    antro_summary: Dict[str, object]
    phq_summary: Dict[str, object]
    sample_summary: Dict[str, object]


def _csalud_cols_for_year(year: int) -> List[str]:
    weight_col = "PESO15_AMA" if year == 2019 else "PESO15_AMAS"
    return [*CSALUD_BASE_COLS, weight_col]


def _load_config(config_path: Path) -> PipelineConfig:
    config_data = json.loads(config_path.read_text(encoding="utf-8-sig"))

    paths = config_data["paths"]
    years = [int(year) for year in config_data["years"]]
    analysis = config_data.get("analysis", {})
    phases = config_data.get("phases", {})
    output = config_data["output"]

    return PipelineConfig(
        raw_archives_dir=Path(paths["raw_archives_dir"]),
        interim_dir=Path(paths["interim_dir"]),
        output_dir=Path(paths["output_dir"]),
        qc_dir=Path(paths["qc_dir"]),
        log_file=Path(paths["log_file"]),
        years=years,
        combined_year_divisor=float(analysis.get("combined_year_divisor", len(years))),
        dataset_base_name=output["dataset_base_name"],
        write_csv=bool(output["write_csv"]),
        write_parquet=bool(output["write_parquet"]),
        run_mice=bool(phases.get("run_mice", False)),
        run_vif=bool(phases.get("run_vif", False)),
        run_r_bridge=bool(phases.get("run_r_bridge", False)),
        analysis_settings=analysis,
    )


def _setup_logging(log_file: Path) -> None:
    log_file.parent.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[
            logging.FileHandler(log_file, mode="w", encoding="utf-8"),
            logging.StreamHandler(),
        ],
        force=True,
    )


def _log_release_metadata(config_path: Path) -> None:
    pipeline_path = Path(__file__).resolve()
    analysis_path = Path(run_analysis.__code__.co_filename).resolve()
    logging.info(
        "Release pipeline | config=%s | pipeline_mtime_utc=%s | analysis_mtime_utc=%s",
        config_path,
        datetime.fromtimestamp(pipeline_path.stat().st_mtime, tz=UTC).isoformat(),
        datetime.fromtimestamp(analysis_path.stat().st_mtime, tz=UTC).isoformat(),
    )


def _assert_release_invariants(csalud: pd.DataFrame, merged: pd.DataFrame, harmonized: pd.DataFrame, year: int) -> None:
    if len(merged) != len(csalud):
        raise AssertionError(
            f"El merge left del anio {year} altero el numero de filas: base={len(csalud)} merged={len(merged)}."
        )

    repro_outside_mef = (
        harmonized["FLAG_MEF"].ne(1)
        & (harmonized["FLAG_EXCL_GESTANTE"].eq(1) | harmonized["FLAG_EXCL_PUERPERIO"].eq(1))
    )
    invalid_count = int(repro_outside_mef.sum())
    if invalid_count:
        raise AssertionError(
            f"Se detectaron {invalid_count} exclusiones reproductivas fuera de MEF en el anio {year}."
        )


def _detect_csv_sep(header_line: str) -> str:
    return ";" if header_line.count(";") > header_line.count(",") else ","


def _read_csv_bytes(raw_bytes: bytes) -> pd.DataFrame:
    first_line = raw_bytes.splitlines()[0].decode("utf-8-sig", errors="ignore")
    sep = _detect_csv_sep(first_line)

    try:
        return pd.read_csv(
            io.BytesIO(raw_bytes),
            sep=sep,
            dtype=str,
            encoding="utf-8-sig",
            low_memory=False,
        )
    except UnicodeDecodeError:
        return pd.read_csv(
            io.BytesIO(raw_bytes),
            sep=sep,
            dtype=str,
            encoding="latin-1",
            low_memory=False,
        )


def _read_dbf_bytes(raw_bytes: bytes) -> pd.DataFrame:
    with tempfile.NamedTemporaryFile(suffix=".dbf", delete=False) as tmp_file:
        tmp_file.write(raw_bytes)
        tmp_path = Path(tmp_file.name)

    try:
        table = DBF(str(tmp_path), load=True, char_decode_errors="ignore")
        records = [dict(record) for record in table]
        return pd.DataFrame(records)
    finally:
        if tmp_path.exists():
            tmp_path.unlink()


def _clean_strings(df: pd.DataFrame) -> pd.DataFrame:
    for col in df.columns:
        series = df[col].astype("string").str.strip()
        series = series.str.replace(r"^(-?\d+)\.0+$", r"\1", regex=True)
        series = series.replace(list(MISSING_TOKENS), pd.NA)
        df[col] = series
    return df


def _ensure_columns(df: pd.DataFrame, needed_columns: List[str]) -> Tuple[pd.DataFrame, List[str]]:
    missing_columns = [col for col in needed_columns if col not in df.columns]
    for col in missing_columns:
        df[col] = pd.NA
    return df[needed_columns].copy(), missing_columns


def _normalize_caseid(raw_series: pd.Series) -> pd.Series:
    return (
        raw_series.astype("string")
        .fillna("")
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
        .replace("", pd.NA)
    )


def _build_caseid_key(hhid: pd.Series, qsnumero: pd.Series) -> pd.Series:
    line_number = (
        qsnumero.astype("string")
        .fillna("")
        .str.strip()
        .str.split(".", n=1, regex=False)
        .str[0]
    )
    return _normalize_caseid(hhid.astype("string").fillna("") + " " + line_number)


def _parse_year_from_name(file_name: str) -> int:
    match = re.match(r"^(\d{4})", file_name)
    if match is None:
        raise ValueError(f"No se pudo extraer el anio desde {file_name}")
    return int(match.group(1))


def _find_outer_archives(raw_archives_dir: Path, years: List[int]) -> Dict[int, Path]:
    archives: Dict[int, Path] = {}

    for zip_path in sorted(raw_archives_dir.glob("*.zip")):
        year = _parse_year_from_name(zip_path.name)
        if year in years:
            archives[year] = zip_path

    missing_years = [year for year in years if year not in archives]
    if missing_years:
        missing = ", ".join(str(year) for year in missing_years)
        raise FileNotFoundError(f"No se encontraron archivos .zip para los anios: {missing}")

    return archives


def _is_target_data_file(file_name: str, module_name: str) -> bool:
    inner_path = Path(file_name)
    stem_upper = inner_path.stem.upper()
    suffix = inner_path.suffix.lower()
    is_target = stem_upper == module_name or stem_upper.startswith(f"{module_name}_")
    return is_target and suffix in {".csv", ".dbf"}


def _parse_module_bytes(raw_dataset: bytes, file_name: str) -> pd.DataFrame:
    suffix = Path(file_name).suffix.lower()
    df = _read_csv_bytes(raw_dataset) if suffix == ".csv" else _read_dbf_bytes(raw_dataset)
    df.columns = [str(col).upper().strip() for col in df.columns]
    return _clean_strings(df)


def _load_module_from_nested_zip(outer_zip_path: Path, module_name: str) -> Tuple[pd.DataFrame, str]:
    with zipfile.ZipFile(outer_zip_path) as outer_zip:
        inner_zip_names = [name for name in outer_zip.namelist() if name.lower().endswith(".zip")]

        # Formato clasico (2019-2024): ZIP externo con ZIPs de modulo internos.
        for inner_zip_name in inner_zip_names:
            inner_bytes = outer_zip.read(inner_zip_name)
            with zipfile.ZipFile(io.BytesIO(inner_bytes)) as inner_zip:
                for inner_file in inner_zip.namelist():
                    if inner_file.endswith("/"):
                        continue
                    if not _is_target_data_file(inner_file, module_name):
                        continue

                    raw_dataset = inner_zip.read(inner_file)
                    df = _parse_module_bytes(raw_dataset, inner_file)
                    source_name = f"{outer_zip_path.name}::{Path(inner_file).name}"
                    return df, source_name

        # Formato 2025+: ZIP externo con CSV directos (p. ej. 'SOLO CVS/CSALUD01_2025.csv').
        for direct_file in outer_zip.namelist():
            if direct_file.endswith("/") or direct_file.lower().endswith(".zip"):
                continue
            if not _is_target_data_file(direct_file, module_name):
                continue

            raw_dataset = outer_zip.read(direct_file)
            df = _parse_module_bytes(raw_dataset, direct_file)
            source_name = f"{outer_zip_path.name}::{Path(direct_file).name}"
            return df, source_name

    raise FileNotFoundError(f"No se encontro el modulo {module_name} dentro de {outer_zip_path.name}")


def _load_required_module(
    outer_zip_path: Path,
    module_name: str,
    needed_columns: List[str],
    year: int,
) -> Tuple[pd.DataFrame, str, List[Dict[str, str]]]:
    df, source_name = _load_module_from_nested_zip(outer_zip_path, module_name)
    df, missing_columns = _ensure_columns(df, needed_columns)

    missing_rows = [
        {"year": str(year), "module": module_name, "missing_column": column}
        for column in missing_columns
    ]
    return df, source_name, missing_rows


def _replace_explicit_missing_codes(df: pd.DataFrame) -> pd.DataFrame:
    for col in EXPLICIT_CODED_MISSING_COLS:
        if col not in df.columns:
            continue
        df[col] = df[col].where(~df[col].isin(CODED_MISSING_TOKENS), pd.NA)
    return df


def _parse_numeric_columns(df: pd.DataFrame, columns: List[str]) -> pd.Series:
    decimal_comma_flag = pd.Series(False, index=df.index)

    for col in columns:
        if col not in df.columns:
            df[col] = np.nan
            continue

        raw = df[col].astype("string")
        decimal_comma_flag |= raw.str.contains(",", regex=False, na=False)
        normalized = raw.str.replace(",", ".", regex=False)
        df[col] = pd.to_numeric(normalized, errors="coerce")

    return decimal_comma_flag


def _cast_nullable_int(series: pd.Series) -> pd.Series:
    non_null = series.dropna()
    if non_null.empty:
        return series.astype("Int64")
    values = non_null.to_numpy(dtype=float)
    if not np.allclose(values, np.round(values)):
        return series
    return series.round().astype("Int64")


def _derive_phq_variables(work: pd.DataFrame) -> None:
    phq_items = work[PHQ_COLS].copy()
    work["PHQ9_ITEMS_FALTANTES"] = phq_items.isna().sum(axis=1).astype("Int64")

    observed_sum = phq_items.sum(axis=1, min_count=1)
    observed_mean = phq_items.mean(axis=1)
    phq_total = pd.Series(np.nan, index=work.index, dtype=float)

    no_missing = work["PHQ9_ITEMS_FALTANTES"].eq(0)
    prorate_mask = work["PHQ9_ITEMS_FALTANTES"].isin([1, 2])

    phq_total.loc[no_missing] = observed_sum.loc[no_missing]
    phq_total.loc[prorate_mask] = (
        observed_sum.loc[prorate_mask]
        + observed_mean.loc[prorate_mask] * work.loc[prorate_mask, "PHQ9_ITEMS_FALTANTES"].astype(float)
    )

    work["PHQ9_PRORRATEADO"] = phq_total
    work["PHQ9_TOTAL"] = phq_total
    work["SEVERIDAD_DEPRESIVA"] = pd.cut(
        work["PHQ9_TOTAL"],
        bins=[-0.001, 4, 9, 14, 19, 27],
        labels=["Minima", "Leve", "Moderada", "Mod_Severa", "Severa"],
        include_lowest=True,
    ).astype("string")


def _derive_bp_variables(work: pd.DataFrame) -> None:
    bp_clean = work[BP_COLS].copy()
    sentinel_mask = bp_clean.isin([999, 999.9])
    work["FLAG_CENTINELA_BP"] = sentinel_mask.any(axis=1).astype("Int64")
    bp_clean = bp_clean.mask(sentinel_mask)

    sys_invalid = pd.DataFrame(False, index=bp_clean.index, columns=SYSTOLIC_COLS)
    dia_invalid = pd.DataFrame(False, index=bp_clean.index, columns=DIASTOLIC_COLS)
    for col in SYSTOLIC_COLS:
        sys_invalid[col] = bp_clean[col].lt(70) | bp_clean[col].gt(270)
    for col in DIASTOLIC_COLS:
        dia_invalid[col] = bp_clean[col].lt(30) | bp_clean[col].gt(150)

    work["FLAG_BP_FUERA_RANGO"] = pd.concat([sys_invalid, dia_invalid], axis=1).any(axis=1).astype("Int64")

    for col in SYSTOLIC_COLS:
        bp_clean[col] = bp_clean[col].mask(sys_invalid[col])
    for col in DIASTOLIC_COLS:
        bp_clean[col] = bp_clean[col].mask(dia_invalid[col])

    inversion_mask = (
        (bp_clean["QS903S"].notna() & bp_clean["QS903D"].notna() & bp_clean["QS903S"].le(bp_clean["QS903D"]))
        | (bp_clean["QS905S"].notna() & bp_clean["QS905D"].notna() & bp_clean["QS905S"].le(bp_clean["QS905D"]))
    )
    work["FLAG_INVERSION_TENSIONAL"] = inversion_mask.astype("Int64")

    bp_valid = bp_clean.notna().all(axis=1) & ~inversion_mask
    work["FLAG_BP_VALIDA"] = bp_valid.astype("Int64")
    work[BP_COLS] = bp_clean

    work["PAS_PROM"] = np.where(bp_valid, bp_clean[SYSTOLIC_COLS].mean(axis=1), np.nan)
    work["PAD_PROM"] = np.where(bp_valid, bp_clean[DIASTOLIC_COLS].mean(axis=1), np.nan)

    outcome = pd.Series(pd.NA, index=work.index, dtype="Int64")
    elevated = work["PAS_PROM"].ge(140) | work["PAD_PROM"].ge(90)
    outcome.loc[bp_valid] = elevated.loc[bp_valid].astype(int).astype("Int64")
    work["PRESION_ARTERIAL_ELEVADA"] = outcome
    work["DESCONTROL_PA"] = outcome.copy()

    # Sensibilidad S1: umbral estilo ACC/AHA 2017 (>=130/80; Whelton et al.,
    # Hypertension 2018, PMID 29133356) sobre el mismo promedio de dos tomas y la
    # misma mascara de validez que el desenlace principal (140/90).
    outcome_130_80 = pd.Series(pd.NA, index=work.index, dtype="Int64")
    elevated_130_80 = work["PAS_PROM"].ge(130) | work["PAD_PROM"].ge(80)
    outcome_130_80.loc[bp_valid] = elevated_130_80.loc[bp_valid].astype(int).astype("Int64")
    work["PRESION_ARTERIAL_ELEVADA_130_80"] = outcome_130_80

    second_take_valid = (
        bp_clean["QS905S"].notna()
        & bp_clean["QS905D"].notna()
        & bp_clean["QS905S"].gt(bp_clean["QS905D"])
    )
    second_outcome = pd.Series(pd.NA, index=work.index, dtype="Int64")
    second_elevated = bp_clean["QS905S"].ge(140) | bp_clean["QS905D"].ge(90)
    second_outcome.loc[second_take_valid] = second_elevated.loc[second_take_valid].astype(int).astype("Int64")
    work["PRESION_ARTERIAL_ELEVADA_SEGUNDA_TOMA"] = second_outcome


def _derive_anthropometry_variables(work: pd.DataFrame) -> None:
    anthro_clean = work[ANTHRO_COLS].copy()
    sentinel_mask = anthro_clean.isin([999, 999.9])
    work["FLAG_CENTINELA_ANTRO"] = sentinel_mask.any(axis=1).astype("Int64")
    anthro_clean = anthro_clean.mask(sentinel_mask)

    peso_invalid = anthro_clean["QS900"].lt(25) | anthro_clean["QS900"].gt(300)
    talla_invalid = anthro_clean["QS901"].lt(100) | anthro_clean["QS901"].gt(250)
    work["FLAG_ANTRO_FUERA_RANGO"] = (peso_invalid | talla_invalid).astype("Int64")

    anthro_clean["QS900"] = anthro_clean["QS900"].mask(peso_invalid)
    anthro_clean["QS901"] = anthro_clean["QS901"].mask(talla_invalid)
    work[ANTHRO_COLS] = anthro_clean

    anthro_valid = anthro_clean["QS900"].notna() & anthro_clean["QS901"].notna()
    work["FLAG_ANTRO_VALIDA"] = anthro_valid.astype("Int64")
    work["IMC"] = np.where(
        anthro_valid,
        anthro_clean["QS900"] / ((anthro_clean["QS901"] / 100.0) ** 2),
        np.nan,
    )


def _derive_diet_quality_variables(work: pd.DataFrame) -> None:
    for col in DIET_SENTINEL_COLS:
        work[col] = work[col].mask(work[col].eq(9.9))

    fruta_entera = pd.Series(np.nan, index=work.index, dtype=float)
    fruta_entera.loc[work["QS213U"].eq(3)] = 0.0
    mask_fruta = work["QS213U"].eq(1) & work["QS214U"].eq(1)
    fruta_entera.loc[mask_fruta] = (
        work.loc[mask_fruta, "QS213C"] * work.loc[mask_fruta, "QS214C"]
    ) / 7.0

    jugo_fruta = pd.Series(np.nan, index=work.index, dtype=float)
    jugo_fruta.loc[work["QS215U"].eq(3)] = 0.0
    mask_jugo = work["QS215U"].eq(1) & work["QS216U"].eq(1)
    jugo_fruta.loc[mask_jugo] = (
        work.loc[mask_jugo, "QS215C"] * work.loc[mask_jugo, "QS216C"]
    ) / 7.0

    ensalada_fruta = pd.Series(np.nan, index=work.index, dtype=float)
    ensalada_fruta.loc[work["QS217U"].eq(3)] = 0.0
    mask_ensalada_fruta = work["QS217U"].eq(1) & work["QS218U"].eq(1)
    ensalada_fruta.loc[mask_ensalada_fruta] = (
        work.loc[mask_ensalada_fruta, "QS217C"] * work.loc[mask_ensalada_fruta, "QS218C"]
    ) / 7.0

    ensalada_verdura = pd.Series(np.nan, index=work.index, dtype=float)
    ensalada_verdura.loc[work["QS219U"].eq(3)] = 0.0
    mask_ensalada_verdura = work["QS219U"].eq(1)
    mask_porciones = mask_ensalada_verdura & work["QS220U"].eq(1)
    mask_cucharadas = mask_ensalada_verdura & work["QS220U"].eq(2)
    ensalada_verdura.loc[mask_porciones] = (
        work.loc[mask_porciones, "QS219C"] * work.loc[mask_porciones, "QS220CV"]
    ) / 7.0
    ensalada_verdura.loc[mask_cucharadas] = (
        work.loc[mask_cucharadas, "QS219C"] * (work.loc[mask_cucharadas, "QS220CC"] / 4.0)
    ) / 7.0

    diet_components = pd.DataFrame(
        {
            "FRUTA_ENTERA": fruta_entera,
            "JUGO_FRUTA": jugo_fruta,
            "ENSALADA_FRUTA": ensalada_fruta,
            "ENSALADA_VERDURA": ensalada_verdura,
        },
        index=work.index,
    )
    consumo_diario = diet_components.sum(axis=1, skipna=True, min_count=1)

    calidad = pd.Series(pd.NA, index=work.index, dtype="Int64")
    calidad.loc[consumo_diario.ge(5)] = 1
    componentes_completos = diet_components.notna().all(axis=1)
    calidad.loc[consumo_diario.lt(5) & componentes_completos] = 0

    work["FRUTA_ENTERA"] = fruta_entera
    work["JUGO_FRUTA"] = jugo_fruta
    work["ENSALADA_FRUTA"] = ensalada_fruta
    work["ENSALADA_VERDURA"] = ensalada_verdura
    work["CONSUMO_DIARIO"] = consumo_diario
    work["CALIDAD_DIETA"] = calidad


def _derive_secondary_variables(work: pd.DataFrame) -> None:
    work["DX_HTA_PREVIO"] = work["QS102"].map({1: 1, 2: 0}).astype("Int64")
    work["DIAGNOSTICO_HTA"] = work["DX_HTA_PREVIO"].copy()
    work["TRATAMIENTO_SALUD_MENTAL"] = work["QS707"].map({1: 1, 2: 0}).astype("Int64")

    no_adherencia = pd.Series(pd.NA, index=work.index, dtype="Int64")
    dx_mask = work["DX_HTA_PREVIO"].eq(1)
    no_adherencia.loc[dx_mask & (work["QS104"].eq(2) | work["QS106"].eq(2))] = 1
    no_adherencia.loc[dx_mask & work["QS104"].eq(1) & work["QS106"].eq(1)] = 0
    work["NO_ADHERENCIA_HTA"] = no_adherencia
    work["ADHERENCIA_MEDICACION"] = work["NO_ADHERENCIA_HTA"].map({0: 1, 1: 0}).astype("Int64")

    violencia = pd.Series(pd.NA, index=work.index, dtype="Int64")
    con_pareja = work["QS709"].eq(1)
    sin_pareja = work["QS709"].eq(2)
    violencia.loc[sin_pareja] = 2
    violencia.loc[con_pareja & work["QS710"].eq(1) & work["QS711"].eq(1)] = 0
    violencia_mask = (
        con_pareja
        & (work["QS710"].ge(2).fillna(False) | work["QS711"].ge(2).fillna(False))
    )
    violencia.loc[violencia_mask] = 1
    work["VIOLENCIA_PAREJA"] = violencia

    alcohol = pd.Series(pd.NA, index=work.index, dtype="Int64")
    alcohol.loc[work["QS209"].eq(2)] = 0
    screen_yes = work[ALCOHOL_SCREEN_COLS].eq(1).any(axis=1)
    screen_complete = work[ALCOHOL_SCREEN_COLS].notna().all(axis=1)
    alcohol.loc[work["QS209"].eq(1) & screen_yes] = 1
    alcohol.loc[work["QS209"].eq(1) & ~screen_yes & screen_complete] = 0
    work["ALCOHOL_PROBLEMATICO"] = alcohol

    _derive_diet_quality_variables(work)

    work["IMPACTO_FUNCIONAL_DEP"] = work["QS702"].map({3: 0, 2: 1, 1: 2}).astype("Int64")


def _final_column_order() -> List[str]:
    return [
        "ANIO",
        "HHID",
        "QSNUMERO",
        "CASEID_KEY",
        "QHCLUSTER",
        "HV001",
        "HV022",
        "HV023",
        "HV024",
        "HV025",
        "HV270",
        "PESO15_AMA",
        "PESO15_AMAS",
        "PESO_SALUD_RAW",
        "PESO_ANALISIS",
        "PESO_FINAL",
        "PESO_MUESTRAL",
        "PESO_HOGAR_AUX",
        "HV005",
        "QSRESULT",
        "QS23",
        "EDAD",
        "QSSEXO",
        "SEXO_CAT",
        "QS25N",
        "AREA_RESIDENCIA",
        "DOMINIO_GEO",
        "REGION_NATURAL",
        "INDICE_RIQUEZA",
        "UBIGEO",
        "CODCCPP",
        "NOMCCPP",
        "HV040",
        "ALTITUD_MSNM",
        "ALTITUD_KM",
        "ALTITUD_CAT3",
        "ALTITUD_ALTA_2500",
        "V454",
        "V222",
        "FLAG_MEF",
        "FLAG_EXCL_GESTANTE",
        "FLAG_EXCL_PUERPERIO",
        "FLAG_REPRO_INFO_INCOMPLETA",
        "FLAG_DECIMAL_COMA",
        "FLAG_CENTINELA_BP",
        "FLAG_CENTINELA_ANTRO",
        "FLAG_BP_FUERA_RANGO",
        "FLAG_ANTRO_FUERA_RANGO",
        "FLAG_INVERSION_TENSIONAL",
        "FLAG_BP_VALIDA",
        "FLAG_ANTRO_VALIDA",
        "QS900",
        "QS901",
        "QS907",
        "IMC",
        "QS102",
        "DX_HTA_PREVIO",
        "DIAGNOSTICO_HTA",
        "QS104",
        "QS106",
        "NO_ADHERENCIA_HTA",
        "ADHERENCIA_MEDICACION",
        "QS109",
        "QS201",
        "QS209",
        *ALCOHOL_SCREEN_COLS,
        "ALCOHOL_PROBLEMATICO",
        "QS707",
        "TRATAMIENTO_SALUD_MENTAL",
        "QS709",
        "QS710",
        "QS711",
        "VIOLENCIA_PAREJA",
        "QS702",
        "IMPACTO_FUNCIONAL_DEP",
        *DIET_COLS,
        "FRUTA_ENTERA",
        "JUGO_FRUTA",
        "ENSALADA_FRUTA",
        "ENSALADA_VERDURA",
        "CONSUMO_DIARIO",
        "CALIDAD_DIETA",
        *PHQ_COLS,
        "PHQ9_ITEMS_FALTANTES",
        "PHQ9_PRORRATEADO",
        "PHQ9_TOTAL",
        "SEVERIDAD_DEPRESIVA",
        *BP_COLS,
        "PAS_PROM",
        "PAD_PROM",
        "PRESION_ARTERIAL_ELEVADA",
        "PRESION_ARTERIAL_ELEVADA_130_80",
        "DESCONTROL_PA",
        "PRESION_ARTERIAL_ELEVADA_SEGUNDA_TOMA",
    ]


def _derive_altitude_variables(work: pd.DataFrame) -> None:
    """Deriva variables de altitud desde HV040 (altitud del cluster/hogar en m).

    HV040 es la fuente analitica primaria de altitud en ENDES. No se imputa: las
    derivadas quedan NA donde HV040 es NA. UBIGEO/CODCCPP/NOMCCPP se conservan
    como texto solo para trazabilidad y auditoria externa por UBIGEO+CODCCPP.
    """
    altitud = pd.to_numeric(work.get("HV040"), errors="coerce")
    work["ALTITUD_MSNM"] = altitud
    work["ALTITUD_KM"] = altitud / 1000.0
    work["ALTITUD_CAT3"] = pd.cut(
        altitud,
        bins=[-np.inf, 1499.999999, 2499.999999, np.inf],
        labels=["<1500", "1500-2499", ">=2500"],
    ).astype("string")
    alta = altitud.ge(2500)
    work["ALTITUD_ALTA_2500"] = alta.where(altitud.notna()).astype("Int64")

    for trace_col in ("UBIGEO", "CODCCPP", "NOMCCPP"):
        if trace_col in work.columns:
            work[trace_col] = work[trace_col].astype("string")


def _harmonize_dataset(merged: pd.DataFrame, year: int, combined_year_divisor: float) -> pd.DataFrame:
    work = merged.copy()
    work["ANIO"] = year

    if "PESO15_AMA" not in work.columns:
        work["PESO15_AMA"] = pd.NA
    if "PESO15_AMAS" not in work.columns:
        work["PESO15_AMAS"] = pd.NA

    work = _replace_explicit_missing_codes(work)
    decimal_comma_flag = _parse_numeric_columns(work, NUMERIC_PARSE_COLS)
    work["FLAG_DECIMAL_COMA"] = decimal_comma_flag.astype("Int64")

    for col in INTEGER_CODE_COLS:
        if col in work.columns:
            work[col] = _cast_nullable_int(work[col])

    weight_col = "PESO15_AMA" if year == 2019 else "PESO15_AMAS"
    work["PESO_SALUD_RAW"] = pd.to_numeric(work[weight_col], errors="coerce")
    work["PESO_ANALISIS"] = work["PESO_SALUD_RAW"] / 1_000_000.0
    work["PESO_FINAL"] = work["PESO_ANALISIS"] / combined_year_divisor
    work["PESO_MUESTRAL"] = work["PESO_FINAL"]
    work["PESO_HOGAR_AUX"] = pd.to_numeric(work["HV005"], errors="coerce") / 1_000_000.0

    work["EDAD"] = pd.to_numeric(work["QS23"], errors="coerce")
    work["SEXO_CAT"] = work["QSSEXO"].map({1: "Hombre", 2: "Mujer"}).astype("string")
    work["AREA_RESIDENCIA"] = work["HV025"]
    work["DOMINIO_GEO"] = work["HV023"]
    work["REGION_NATURAL"] = work["HV024"]
    work["INDICE_RIQUEZA"] = work["HV270"]

    _derive_altitude_variables(work)

    flag_mef = work["QSSEXO"].eq(2) & work["EDAD"].between(15, 49, inclusive="both")
    work["FLAG_MEF"] = flag_mef.astype("Int64")
    work["FLAG_EXCL_GESTANTE"] = (flag_mef & work["V454"].eq(1)).astype("Int64")
    work["FLAG_EXCL_PUERPERIO"] = (flag_mef & work["V222"].lt(2).fillna(False)).astype("Int64")
    work["FLAG_REPRO_INFO_INCOMPLETA"] = (
        flag_mef & work["V454"].isna() & work["V222"].isna()
    ).astype("Int64")

    _derive_phq_variables(work)
    _derive_bp_variables(work)
    _derive_anthropometry_variables(work)
    _derive_secondary_variables(work)

    ordered_columns = [column for column in _final_column_order() if column in work.columns]
    return work[ordered_columns].copy()


def _build_flow_masks(df: pd.DataFrame) -> "OrderedDict[str, pd.Series]":
    rech0_available = df[["HV023", "HV024", "HV025"]].notna().all(axis=1)
    design_available = df[["HV001", "HV022"]].notna().all(axis=1)

    return OrderedDict(
        [
            ("qsresult_ok", df["QSRESULT"].eq(1)),
            ("adult_18_plus", df["EDAD"].ge(18)),
            ("valid_sex", df["QSSEXO"].isin([1, 2])),
            ("rech0_available", rech0_available),
            ("design_available", design_available),
            ("rech23_available", df["HV270"].notna()),
            ("positive_health_weight", df["PESO_SALUD_RAW"].gt(0)),
            ("not_pregnant_or_not_mef", df["FLAG_EXCL_GESTANTE"].eq(0)),
            ("not_puerperal_or_not_mef", df["FLAG_EXCL_PUERPERIO"].eq(0)),
            ("phq_available", df["PHQ9_TOTAL"].notna()),
            ("bp_available", df["FLAG_BP_VALIDA"].eq(1)),
        ]
    )


def _compute_flow_counts(df: pd.DataFrame, masks: "OrderedDict[str, pd.Series]", year: int) -> Dict[str, int]:
    results: Dict[str, int] = {"year": year, "n_start": int(len(df))}
    running_mask = pd.Series(True, index=df.index)

    for step, mask in masks.items():
        step_mask = mask.fillna(False)
        before = int(running_mask.sum())
        updated = running_mask & step_mask
        after = int(updated.sum())
        results[f"before_{step}"] = before
        results[f"drop_{step}"] = before - after
        results[f"after_{step}"] = after
        running_mask = updated

    results["n_final"] = int(running_mask.sum())
    return results


def _running_mask_until(masks: "OrderedDict[str, pd.Series]", stop_step: str) -> pd.Series:
    running_mask = pd.Series(True, index=next(iter(masks.values())).index)
    for step, mask in masks.items():
        running_mask &= mask.fillna(False)
        if step == stop_step:
            return running_mask
    raise KeyError(stop_step)


def _summarize_reproductive(df: pd.DataFrame, masks: "OrderedDict[str, pd.Series]", year: int) -> Dict[str, object]:
    pre_repro_mask = _running_mask_until(masks, "positive_health_weight")
    structural_mask = _running_mask_until(masks, "not_puerperal_or_not_mef")
    subset = df.loc[pre_repro_mask]

    return {
        "year": year,
        "n_base_pre_repro": int(pre_repro_mask.sum()),
        "n_mef": int(subset["FLAG_MEF"].eq(1).sum()),
        "n_excl_gestante": int(subset["FLAG_EXCL_GESTANTE"].eq(1).sum()),
        "n_excl_puerperio": int(subset["FLAG_EXCL_PUERPERIO"].eq(1).sum()),
        "n_repro_info_incompleta": int(subset["FLAG_REPRO_INFO_INCOMPLETA"].eq(1).sum()),
        "n_structural_base": int(structural_mask.sum()),
    }


def _summarize_bp(df: pd.DataFrame, masks: "OrderedDict[str, pd.Series]", year: int) -> Dict[str, object]:
    structural_mask = _running_mask_until(masks, "not_puerperal_or_not_mef")
    subset = df.loc[structural_mask]

    return {
        "year": year,
        "n_structural_base": int(len(subset)),
        "n_flag_decimal_coma": int(subset["FLAG_DECIMAL_COMA"].eq(1).sum()),
        "n_centinela_bp": int(subset["FLAG_CENTINELA_BP"].eq(1).sum()),
        "n_fuera_rango_bp": int(subset["FLAG_BP_FUERA_RANGO"].eq(1).sum()),
        "n_inversion_tensional": int(subset["FLAG_INVERSION_TENSIONAL"].eq(1).sum()),
        "n_bp_valida": int(subset["FLAG_BP_VALIDA"].eq(1).sum()),
        "n_bp_invalida": int(subset["FLAG_BP_VALIDA"].ne(1).sum()),
    }


def _summarize_antro(df: pd.DataFrame, masks: "OrderedDict[str, pd.Series]", year: int) -> Dict[str, object]:
    structural_mask = _running_mask_until(masks, "not_puerperal_or_not_mef")
    subset = df.loc[structural_mask]

    return {
        "year": year,
        "n_structural_base": int(len(subset)),
        "n_flag_decimal_coma": int(subset["FLAG_DECIMAL_COMA"].eq(1).sum()),
        "n_centinela_antro": int(subset["FLAG_CENTINELA_ANTRO"].eq(1).sum()),
        "n_fuera_rango_antro": int(subset["FLAG_ANTRO_FUERA_RANGO"].eq(1).sum()),
        "n_antro_valida": int(subset["FLAG_ANTRO_VALIDA"].eq(1).sum()),
        "n_antro_invalida": int(subset["FLAG_ANTRO_VALIDA"].ne(1).sum()),
    }


def _summarize_phq(df: pd.DataFrame, masks: "OrderedDict[str, pd.Series]", year: int) -> Dict[str, object]:
    structural_mask = _running_mask_until(masks, "not_puerperal_or_not_mef")
    subset = df.loc[structural_mask]

    return {
        "year": year,
        "n_structural_base": int(len(subset)),
        "n_phq_sin_faltantes": int(subset["PHQ9_ITEMS_FALTANTES"].eq(0).sum()),
        "n_phq_prorrateado_1_2": int(subset["PHQ9_ITEMS_FALTANTES"].isin([1, 2]).sum()),
        "n_phq_excluido_mas_de_2": int(subset["PHQ9_ITEMS_FALTANTES"].gt(2).sum()),
        "n_phq_total_disponible": int(subset["PHQ9_TOTAL"].notna().sum()),
    }


def _summarize_samples(df: pd.DataFrame, masks: "OrderedDict[str, pd.Series]", year: int) -> Dict[str, object]:
    structural_mask = _running_mask_until(masks, "not_puerperal_or_not_mef")
    principal_mask = _running_mask_until(masks, "bp_available")
    principal_subset = df.loc[principal_mask]
    dx_domain = principal_subset["DX_HTA_PREVIO"].eq(1)

    return {
        "year": year,
        "n_cohorte_estructural": int(structural_mask.sum()),
        "n_cohorte_principal": int(len(principal_subset)),
        "n_sensibilidad_sin_2020": int((principal_subset["ANIO"] != 2020).sum()),
        "n_dx_hta_previo_dominio": int(dx_domain.sum()),
        "n_subanalisis_no_adherencia": int((dx_domain & principal_subset["NO_ADHERENCIA_HTA"].notna()).sum()),
        "n_subanalisis_descontrol_dxhta": int(dx_domain.sum()),
        "n_phq_prorrateado_en_final": int(principal_subset["PHQ9_ITEMS_FALTANTES"].isin([1, 2]).sum()),
    }


def _append_total_row(df: pd.DataFrame, label: str = "TOTAL") -> pd.DataFrame:
    if df.empty:
        return df

    numeric_cols = [
        col
        for col in df.columns
        if col not in {"year", "inclusion_pct"} and pd.api.types.is_numeric_dtype(df[col])
    ]
    total_row = {col: df[col].sum() for col in numeric_cols}
    total_row["year"] = label
    if "inclusion_pct" in df.columns and {"n_final", "n_start"}.issubset(total_row):
        total_row["inclusion_pct"] = round((total_row["n_final"] / total_row["n_start"]) * 100, 2)
    ordered = ["year", *[col for col in df.columns if col != "year"]]
    return pd.concat([df, pd.DataFrame([total_row])[ordered]], ignore_index=True)


def _build_population_summary(flow_df: pd.DataFrame) -> pd.DataFrame:
    summary = flow_df[
        [
            "year",
            "n_start",
            "after_not_puerperal_or_not_mef",
            "drop_phq_available",
            "drop_bp_available",
            "n_final",
        ]
    ].copy()
    summary = summary.rename(
        columns={
            "after_not_puerperal_or_not_mef": "n_cohorte_estructural",
            "drop_phq_available": "n_excluidos_phq",
            "drop_bp_available": "n_excluidos_bp",
        }
    )
    summary["n_excluidos_estructurales"] = summary["n_start"] - summary["n_cohorte_estructural"]
    summary["n_excluidos_totales"] = summary["n_start"] - summary["n_final"]
    summary["inclusion_pct"] = ((summary["n_final"] / summary["n_start"]) * 100).round(2)
    return summary


def _build_exclusions_total(flow_df: pd.DataFrame) -> pd.DataFrame:
    total_start = int(flow_df["n_start"].sum())
    criteria = [
        "drop_qsresult_ok",
        "drop_adult_18_plus",
        "drop_valid_sex",
        "drop_rech0_available",
        "drop_design_available",
        "drop_rech23_available",
        "drop_positive_health_weight",
        "drop_not_pregnant_or_not_mef",
        "drop_not_puerperal_or_not_mef",
        "drop_phq_available",
        "drop_bp_available",
    ]

    rows = []
    for criterion in criteria:
        excluded = int(flow_df[criterion].sum())
        rows.append(
            {
                "criterio": criterion,
                "n_excluidos": excluded,
                "pct_sobre_inicio": round((excluded / total_start) * 100, 4) if total_start else 0.0,
            }
        )

    return pd.DataFrame(rows)


def _build_included_by_sex(unified: pd.DataFrame) -> pd.DataFrame:
    result = (
        unified.groupby("SEXO_CAT", dropna=False, as_index=False)
        .size()
        .rename(columns={"size": "n_incluidos"})
        .sort_values("SEXO_CAT")
    )
    total = result["n_incluidos"].sum()
    result["pct"] = ((result["n_incluidos"] / total) * 100).round(2) if total else 0.0
    return result


def _build_model_sample_sizes(sample_rows: List[Dict[str, object]]) -> pd.DataFrame:
    sample_df = pd.DataFrame(sample_rows)
    totals = {
        "n_cohorte_estructural": int(sample_df["n_cohorte_estructural"].sum()),
        "n_cohorte_principal": int(sample_df["n_cohorte_principal"].sum()),
        "n_sensibilidad_sin_2020": int(sample_df["n_sensibilidad_sin_2020"].sum()),
        "n_dx_hta_previo_dominio": int(sample_df["n_dx_hta_previo_dominio"].sum()),
        "n_subanalisis_no_adherencia": int(sample_df["n_subanalisis_no_adherencia"].sum()),
        "n_subanalisis_descontrol_dxhta": int(sample_df["n_subanalisis_descontrol_dxhta"].sum()),
        "n_phq_prorrateado_en_final": int(sample_df["n_phq_prorrateado_en_final"].sum()),
    }
    return pd.DataFrame(
        [{"metrica": metric, "valor": value} for metric, value in totals.items()]
    )


def _process_year(year: int, outer_zip_path: Path, combined_year_divisor: float) -> YearArtifacts:
    logging.info("Procesando anio %s desde %s", year, outer_zip_path.name)

    missing_column_rows: List[Dict[str, str]] = []

    csalud, src_csalud, missing = _load_required_module(
        outer_zip_path,
        "CSALUD01",
        _csalud_cols_for_year(year),
        year,
    )
    missing_column_rows.extend(missing)

    rech0, src_rech0, missing = _load_required_module(outer_zip_path, "RECH0", RECH0_COLS, year)
    missing_column_rows.extend(missing)

    rech23, src_rech23, missing = _load_required_module(outer_zip_path, "RECH23", RECH23_COLS, year)
    missing_column_rows.extend(missing)

    rec42, src_rec42, missing = _load_required_module(outer_zip_path, "REC42", REC42_COLS, year)
    missing_column_rows.extend(missing)

    re223132, src_re223132, missing = _load_required_module(
        outer_zip_path,
        "RE223132",
        RE223132_COLS,
        year,
    )
    missing_column_rows.extend(missing)

    csalud["HHID"] = csalud["HHID"].astype("string").str.strip()
    rech0["HHID"] = rech0["HHID"].astype("string").str.strip()
    rech23["HHID"] = rech23["HHID"].astype("string").str.strip()

    csalud["CASEID_KEY"] = _build_caseid_key(csalud["HHID"], csalud["QSNUMERO"])
    rec42["CASEID_KEY"] = _normalize_caseid(rec42["CASEID"])
    re223132["CASEID_KEY"] = _normalize_caseid(re223132["CASEID"])

    rech0 = rech0.drop_duplicates(subset=["HHID"], keep="first")
    rech23 = rech23.drop_duplicates(subset=["HHID"], keep="first")
    rec42 = rec42.drop_duplicates(subset=["CASEID_KEY"], keep="first")
    re223132 = re223132.drop_duplicates(subset=["CASEID_KEY"], keep="first")

    merged = csalud.merge(rech0, on="HHID", how="left", validate="many_to_one")
    merged = merged.merge(rech23, on="HHID", how="left", validate="many_to_one")
    merged = merged.merge(rec42[["CASEID_KEY", "V454"]], on="CASEID_KEY", how="left", validate="many_to_one")
    merged = merged.merge(
        re223132[["CASEID_KEY", "V222"]],
        on="CASEID_KEY",
        how="left",
        validate="many_to_one",
    )

    harmonized = _harmonize_dataset(merged, year, combined_year_divisor)
    _assert_release_invariants(csalud, merged, harmonized, year)
    masks = _build_flow_masks(harmonized)
    flow_counts = _compute_flow_counts(harmonized, masks, year)

    final_mask = pd.Series(True, index=harmonized.index)
    for mask in masks.values():
        final_mask &= mask.fillna(False)

    final_df = harmonized.loc[final_mask].copy()

    source_map = {
        "year": str(year),
        "archive": outer_zip_path.name,
        "CSALUD01": src_csalud,
        "RECH0": src_rech0,
        "RECH23": src_rech23,
        "REC42": src_rec42,
        "RE223132": src_re223132,
    }

    if missing_column_rows:
        logging.warning(
            "Anio %s con columnas faltantes en modulos: %s",
            year,
            len(missing_column_rows),
        )

    logging.info("Anio %s finalizado: n_final=%s", year, flow_counts["n_final"])
    return YearArtifacts(
        final_df=final_df,
        flow_counts=flow_counts,
        source_map=source_map,
        missing_column_rows=missing_column_rows,
        reproductive_summary=_summarize_reproductive(harmonized, masks, year),
        bp_summary=_summarize_bp(harmonized, masks, year),
        antro_summary=_summarize_antro(harmonized, masks, year),
        phq_summary=_summarize_phq(harmonized, masks, year),
        sample_summary=_summarize_samples(harmonized, masks, year),
    )


def _write_missing_by_phq9_severity(unified: pd.DataFrame, config: PipelineConfig) -> None:
    """H3: % de datos faltantes en covariables imputadas (targets MICE) segun la
    severidad depresiva (PHQ-9). Sustenta el supuesto MAR de la imputacion multiple:
    un patron diferencial (mas faltantes a mayor severidad) refuerza la necesidad de MICE.
    """
    target_cols = ["QS25N", "IMC", "QS907", "CALIDAD_DIETA", "QS109", "VIOLENCIA_PAREJA"]
    severity_order = ["Minima", "Leve", "Moderada", "Mod_Severa", "Severa"]
    sev = unified["SEVERIDAD_DEPRESIVA"].astype("string")

    rows: List[Dict[str, object]] = []
    for var in target_cols:
        if var not in unified.columns:
            continue
        is_missing = unified[var].isna()
        row: Dict[str, object] = {"variable": var}
        pct_by_cat: Dict[str, float] = {}
        for cat in severity_order:
            mask = sev.eq(cat)
            n_cat = int(mask.sum())
            n_miss = int((mask & is_missing).sum())
            pct = round(100.0 * n_miss / n_cat, 3) if n_cat else np.nan
            pct_by_cat[cat] = pct
            row[f"n_{cat}"] = n_cat
            row[f"pct_missing_{cat}"] = pct
        row["pct_missing_global"] = round(float(is_missing.mean()) * 100, 3)
        pmin = pct_by_cat.get("Minima")
        psev = pct_by_cat.get("Severa")
        row["ratio_severa_vs_minima"] = (
            round(psev / pmin, 3) if pmin not in (None, 0) and not pd.isna(pmin) and psev is not None else np.nan
        )
        row["diferencial_gt2x"] = bool(row["ratio_severa_vs_minima"] > 2) if not pd.isna(row["ratio_severa_vs_minima"]) else False
        rows.append(row)

    pd.DataFrame(rows).to_csv(
        config.qc_dir / "missing_por_severidad_phq9.csv", index=False, encoding="utf-8-sig"
    )
    logging.info("QC H3 escrito: missing_por_severidad_phq9.csv (%s covariables)", len(rows))


def _write_altitude_qc(unified: pd.DataFrame, config: PipelineConfig) -> None:
    """Escribe los tres QC de altitud. INEI 2017 se usa solo para auditoria
    externa por UBIGEO+CODCCPP; nunca como llave por HV001 ni para completar la base."""
    alt = pd.to_numeric(unified["ALTITUD_MSNM"], errors="coerce")

    # 1) Faltantes y rango de ALTITUD_MSNM por anio.
    missing_rows: List[Dict[str, object]] = []
    for year, idx in unified.groupby("ANIO").groups.items():
        sub = alt.loc[idx]
        has = sub.notna()
        missing_rows.append(
            {
                "year": int(year),
                "n": int(len(sub)),
                "n_altitud_disponible": int(has.sum()),
                "n_altitud_faltante": int((~has).sum()),
                "pct_faltante": round(float((~has).mean()) * 100, 4),
                "altitud_min": float(sub.min()) if has.any() else np.nan,
                "altitud_mediana": float(sub.median()) if has.any() else np.nan,
                "altitud_max": float(sub.max()) if has.any() else np.nan,
                "n_lt1500": int((sub < 1500).sum()),
                "n_1500_2499": int(((sub >= 1500) & (sub < 2500)).sum()),
                "n_ge2500": int((sub >= 2500).sum()),
            }
        )
    total = {
        "year": "TOTAL",
        "n": int(len(alt)),
        "n_altitud_disponible": int(alt.notna().sum()),
        "n_altitud_faltante": int(alt.isna().sum()),
        "pct_faltante": round(float(alt.isna().mean()) * 100, 4),
        "altitud_min": float(alt.min()) if alt.notna().any() else np.nan,
        "altitud_mediana": float(alt.median()) if alt.notna().any() else np.nan,
        "altitud_max": float(alt.max()) if alt.notna().any() else np.nan,
        "n_lt1500": int((alt < 1500).sum()),
        "n_1500_2499": int(((alt >= 1500) & (alt < 2500)).sum()),
        "n_ge2500": int((alt >= 2500).sum()),
    }
    pd.DataFrame([*missing_rows, total]).to_csv(
        config.qc_dir / "altitude_missing_by_year.csv", index=False, encoding="utf-8-sig"
    )

    # 2) Consistencia de altitud dentro de cada cluster ANIO + HV001.
    consist_rows: List[Dict[str, object]] = []
    for year, sub in unified.groupby("ANIO"):
        grp = pd.to_numeric(sub["ALTITUD_MSNM"], errors="coerce").groupby(sub["HV001"])
        nun = grp.nunique(dropna=True)
        spread = grp.apply(lambda s: (s.max() - s.min()) if s.notna().any() else np.nan)
        consist_rows.append(
            {
                "year": int(year),
                "n_clusters": int(nun.shape[0]),
                "n_clusters_altitud_unica": int((nun <= 1).sum()),
                "n_clusters_altitud_variable": int((nun > 1).sum()),
                "max_rango_intra_cluster_m": float(spread.max()) if spread.notna().any() else np.nan,
                "mediana_rango_intra_cluster_m": float(spread.dropna().median()) if spread.notna().any() else np.nan,
            }
        )
    pd.DataFrame(consist_rows).to_csv(
        config.qc_dir / "altitude_cluster_consistency.csv", index=False, encoding="utf-8-sig"
    )

    # 3) Auditoria externa contra INEI 2017 por UBIGEO + CODCCPP (NO por HV001).
    inei_csv = (
        config.raw_archives_dir.parent
        / "inei_centros_poblados_2017"
        / "centros_poblados_inei_2017_consolidado.csv"
    )
    if not inei_csv.exists():
        logging.warning("INEI 2017 no encontrado en %s; se omite altitude_inei_match_audit.csv", inei_csv)
        return

    try:
        inei = pd.read_csv(inei_csv, dtype=str)
        inei_cols = {c.upper(): c for c in inei.columns}
        ubigeo_c = inei_cols.get("UBIGEO")
        codccpp_c = inei_cols.get("CODCCPP")
        alt_c = inei_cols.get("ALTITUD_MSNM")
        inei_key = (
            inei[ubigeo_c].astype("string").str.strip().str.zfill(6)
            + inei[codccpp_c].astype("string").str.strip().str.zfill(4)
        )
        inei_alt = pd.to_numeric(inei[alt_c].astype("string").str.replace(",", "."), errors="coerce")
        inei_lookup = pd.DataFrame({"KEY": inei_key, "INEI_ALTITUD_MSNM": inei_alt}).dropna(subset=["KEY"])
        inei_lookup = inei_lookup.drop_duplicates(subset=["KEY"], keep="first").set_index("KEY")["INEI_ALTITUD_MSNM"]

        end = unified[["UBIGEO", "CODCCPP", "ALTITUD_MSNM"]].copy()
        end_key = (
            end["UBIGEO"].astype("string").str.strip().str.zfill(6)
            + end["CODCCPP"].astype("string").str.strip().str.zfill(4)
        )
        end_alt = pd.to_numeric(end["ALTITUD_MSNM"], errors="coerce")
        uniq = pd.DataFrame({"KEY": end_key, "ENDES_ALTITUD_MSNM": end_alt}).dropna(subset=["KEY"])
        uniq = uniq.groupby("KEY", as_index=False)["ENDES_ALTITUD_MSNM"].median()
        uniq["INEI_ALTITUD_MSNM"] = uniq["KEY"].map(inei_lookup)
        matched = uniq["INEI_ALTITUD_MSNM"].notna()
        diff = (uniq.loc[matched, "ENDES_ALTITUD_MSNM"] - uniq.loc[matched, "INEI_ALTITUD_MSNM"]).abs()

        audit = pd.DataFrame(
            [
                {"metrica": "n_claves_endes_ubigeo_codccpp", "valor": int(uniq.shape[0])},
                {"metrica": "n_claves_con_match_inei", "valor": int(matched.sum())},
                {"metrica": "pct_match_inei", "valor": round(float(matched.mean()) * 100, 2) if uniq.shape[0] else 0.0},
                {"metrica": "dif_abs_mediana_m", "valor": round(float(diff.median()), 2) if matched.any() else np.nan},
                {"metrica": "dif_abs_media_m", "valor": round(float(diff.mean()), 2) if matched.any() else np.nan},
                {"metrica": "dif_abs_p95_m", "valor": round(float(diff.quantile(0.95)), 2) if matched.any() else np.nan},
                {"metrica": "pct_dentro_100m", "valor": round(float((diff <= 100).mean()) * 100, 2) if matched.any() else np.nan},
                {"metrica": "pct_dentro_250m", "valor": round(float((diff <= 250).mean()) * 100, 2) if matched.any() else np.nan},
            ]
        )
        audit.to_csv(config.qc_dir / "altitude_inei_match_audit.csv", index=False, encoding="utf-8-sig")
        logging.info(
            "Auditoria INEI altitud: %s/%s claves con match (%.1f%%)",
            int(matched.sum()),
            int(uniq.shape[0]),
            float(matched.mean()) * 100 if uniq.shape[0] else 0.0,
        )
    except Exception as exc:  # pragma: no cover - auditoria externa best-effort
        logging.warning("Fallo la auditoria INEI de altitud (no critico): %s", exc)


def run_pipeline(config_path: str = "config/pipeline_config.json") -> None:
    config_path_obj = Path(config_path)
    config = _load_config(config_path_obj)

    config.interim_dir.mkdir(parents=True, exist_ok=True)
    config.output_dir.mkdir(parents=True, exist_ok=True)
    config.qc_dir.mkdir(parents=True, exist_ok=True)

    _setup_logging(config.log_file)
    logging.info("Iniciando pipeline ENDES reimplementado")
    _log_release_metadata(config_path_obj.resolve())

    archives = _find_outer_archives(config.raw_archives_dir, config.years)

    final_frames: List[pd.DataFrame] = []
    flow_rows: List[Dict[str, int]] = []
    source_rows: List[Dict[str, str]] = []
    missing_column_rows: List[Dict[str, str]] = []
    reproductive_rows: List[Dict[str, object]] = []
    bp_rows: List[Dict[str, object]] = []
    antro_rows: List[Dict[str, object]] = []
    phq_rows: List[Dict[str, object]] = []
    sample_rows: List[Dict[str, object]] = []

    for year in config.years:
        artifacts = _process_year(year, archives[year], config.combined_year_divisor)
        final_frames.append(artifacts.final_df)
        flow_rows.append(artifacts.flow_counts)
        source_rows.append(artifacts.source_map)
        missing_column_rows.extend(artifacts.missing_column_rows)
        reproductive_rows.append(artifacts.reproductive_summary)
        bp_rows.append(artifacts.bp_summary)
        antro_rows.append(artifacts.antro_summary)
        phq_rows.append(artifacts.phq_summary)
        sample_rows.append(artifacts.sample_summary)

    logging.info(
        "TIEMPO_DX_HTA_MESES excluida formalmente del estudio: QS103U/QS103C vacios o no utilizables en los datos crudos."
    )

    unified = pd.concat(final_frames, axis=0, ignore_index=True).sort_values(
        ["ANIO", "HHID", "QSNUMERO"],
        kind="stable",
    )
    unified = unified.reset_index(drop=True)

    dataset_base = config.output_dir / config.dataset_base_name
    if config.write_parquet:
        parquet_path = dataset_base.with_suffix(".parquet")
        unified.to_parquet(parquet_path, index=False)
        logging.info("Dataset parquet guardado en %s", parquet_path)

    if config.write_csv:
        csv_path = dataset_base.with_suffix(".csv")
        unified.to_csv(csv_path, index=False, encoding="utf-8-sig")
        logging.info("Dataset csv guardado en %s", csv_path)

    flow_df = pd.DataFrame(flow_rows).sort_values("year")
    flow_df.to_csv(config.qc_dir / "filter_flow_by_year.csv", index=False, encoding="utf-8-sig")

    counts_by_year = (
        unified.groupby("ANIO", as_index=False)
        .size()
        .rename(columns={"size": "n_final"})
        .sort_values("ANIO")
    )
    counts_by_year.to_csv(config.qc_dir / "final_counts_by_year.csv", index=False, encoding="utf-8-sig")

    _write_altitude_qc(unified, config)
    _write_missing_by_phq9_severity(unified, config)

    counts_by_year_sex = (
        unified.groupby(["ANIO", "SEXO_CAT"], as_index=False)
        .size()
        .rename(columns={"size": "n"})
        .sort_values(["ANIO", "SEXO_CAT"])
    )
    counts_by_year_sex.to_csv(
        config.qc_dir / "final_counts_by_year_sex.csv",
        index=False,
        encoding="utf-8-sig",
    )

    pd.DataFrame(source_rows).sort_values("year").to_csv(
        config.qc_dir / "source_traceability.csv",
        index=False,
        encoding="utf-8-sig",
    )

    pd.DataFrame(missing_column_rows, columns=["year", "module", "missing_column"]).to_csv(
        config.qc_dir / "analysis_plan_missing_columns.csv",
        index=False,
        encoding="utf-8-sig",
    )

    population_summary = _append_total_row(_build_population_summary(flow_df))
    population_summary.to_csv(
        config.qc_dir / "analysis_plan_population_summary.csv",
        index=False,
        encoding="utf-8-sig",
    )

    exclusions_total = _build_exclusions_total(flow_df)
    exclusions_total.to_csv(
        config.qc_dir / "analysis_plan_exclusions_total.csv",
        index=False,
        encoding="utf-8-sig",
    )

    included_by_sex = _build_included_by_sex(unified)
    included_by_sex.to_csv(
        config.qc_dir / "analysis_plan_included_by_sex.csv",
        index=False,
        encoding="utf-8-sig",
    )

    _append_total_row(pd.DataFrame(reproductive_rows).sort_values("year")).to_csv(
        config.qc_dir / "analysis_plan_reproductive_summary.csv",
        index=False,
        encoding="utf-8-sig",
    )
    _append_total_row(pd.DataFrame(bp_rows).sort_values("year")).to_csv(
        config.qc_dir / "analysis_plan_bp_cleaning_summary.csv",
        index=False,
        encoding="utf-8-sig",
    )
    _append_total_row(pd.DataFrame(antro_rows).sort_values("year")).to_csv(
        config.qc_dir / "analysis_plan_antro_cleaning_summary.csv",
        index=False,
        encoding="utf-8-sig",
    )
    _append_total_row(pd.DataFrame(phq_rows).sort_values("year")).to_csv(
        config.qc_dir / "analysis_plan_phq_proration_summary.csv",
        index=False,
        encoding="utf-8-sig",
    )
    _build_model_sample_sizes(sample_rows).to_csv(
        config.qc_dir / "analysis_plan_model_sample_sizes.csv",
        index=False,
        encoding="utf-8-sig",
    )

    if config.run_mice or config.run_vif or config.run_r_bridge:
        analysis_settings = build_analysis_settings(config.output_dir, config.analysis_settings)
        run_analysis(
            unified,
            analysis_settings,
            run_mice=config.run_mice,
            run_vif=config.run_vif,
            run_r_bridge=config.run_r_bridge,
        )

    logging.info("Pipeline finalizado. Filas finales=%s", len(unified))


if __name__ == "__main__":
    run_pipeline()
