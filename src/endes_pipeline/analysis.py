from __future__ import annotations

import json
import logging
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

import numpy as np
import pandas as pd
from scipy import stats
from sklearn.experimental import enable_iterative_imputer  # noqa: F401
from sklearn.impute import IterativeImputer
from statsmodels.stats.outliers_influence import variance_inflation_factor
from statsmodels.tools.tools import add_constant


MODEL2_COVARS = ["QS23", "QSSEXO", "QS25N", "HV025", "HV270", "VIOLENCIA_PAREJA"]
MODEL3_EXTRA_COVARS = ["IMC", "QS907", "QS201", "ALCOHOL_PROBLEMATICO", "CALIDAD_DIETA", "QS109"]
NON_IMPUTED_SKIP_COVARS = ["QS201", "ALCOHOL_PROBLEMATICO"]
MICE_TARGETS = MODEL2_COVARS + [col for col in MODEL3_EXTRA_COVARS if col not in NON_IMPUTED_SKIP_COVARS]
CONTINUOUS_IMPUTE_COLS = {"IMC", "QS907"}
CATEGORICAL_IMPUTE_COLS = [col for col in MICE_TARGETS if col not in CONTINUOUS_IMPUTE_COLS]
MICE_PREDICTORS = ["PHQ9_TOTAL", "PRESION_ARTERIAL_ELEVADA", "ANIO", "HV001", "HV022", "PESO_FINAL"]
VIF_PREDICTORS = ["PHQ9_TOTAL", *MODEL2_COVARS, *MODEL3_EXTRA_COVARS, "ALTITUD_CAT3"]

TABLE1_CONTINUOUS = ["QS23", "PHQ9_TOTAL", "IMC", "QS907"]
TABLE1_CATEGORICAL = [
    "SEVERIDAD_DEPRESIVA",
    "QSSEXO",
    "QS25N",
    "HV025",
    "HV270",
    "VIOLENCIA_PAREJA",
    "QS201",
    "ALCOHOL_PROBLEMATICO",
    "CALIDAD_DIETA",
    "QS109",
    "DX_HTA_PREVIO",
    "ALTITUD_CAT3",
]
TABLE1_GROUPS: Tuple[Tuple[str, str], ...] = (
    ("overall", ""),
    ("by_outcome", "PRESION_ARTERIAL_ELEVADA"),
    ("by_sex", "QSSEXO"),
    ("by_severity", "SEVERIDAD_DEPRESIVA"),
)
RAO_SCOTT_VARS = [
    "SEVERIDAD_DEPRESIVA",
    "QSSEXO",
    "QS25N",
    "HV025",
    "HV270",
    "VIOLENCIA_PAREJA",
    "QS201",
    "ALCOHOL_PROBLEMATICO",
    "CALIDAD_DIETA",
    "QS109",
    "ALTITUD_CAT3",
]
BIVARIATE_TEST_VARS = RAO_SCOTT_VARS

# Conjunto estructural preespecificado (Modelo 2 original, SIN altitud).
_STRUCT_COVARS = (
    "QS23 + factor(QSSEXO) + factor(QS25N) + factor(HV025) + factor(HV270) + "
    "factor(VIOLENCIA_PAREJA) + factor(ANIO)"
)
# Enmienda 2026-06-01: la altitud (factor(ALTITUD_CAT3)) se incorpora al conjunto de
# ajuste ESTANDAR tras identificarse como confusor geografico ausente del plan original.
# Justificacion: en Peru la altitud se asocia con MAS depresion (Hernandez-Vasquez,
# J Affect Disord 2022; PMID 34942223) y MENOS hipertension (Mendoza-Quispe, J Hypertens
# 2023; PMID 37071440), lo que induce la asociacion inversa espuria PHQ-9<->PAE observada.
# Forma categorica (umbral, no lineal). NO se imputa; HV040 es completa.
_ALT = "factor(ALTITUD_CAT3)"
_MODEL2_NOALT = f"PRESION_ARTERIAL_ELEVADA ~ PHQ9_TOTAL + {_STRUCT_COVARS}"

MAIN_MODELS = {
    "model_1": "PRESION_ARTERIAL_ELEVADA ~ PHQ9_TOTAL",
    "model_2": f"{_MODEL2_NOALT} + {_ALT}",
    "model_3": (
        f"PRESION_ARTERIAL_ELEVADA ~ PHQ9_TOTAL + {_STRUCT_COVARS} + IMC + QS907 + "
        f"factor(QS201) + factor(ALCOHOL_PROBLEMATICO) + factor(CALIDAD_DIETA) + factor(QS109) + {_ALT}"
    ),
    "submodel_adherence": f"NO_ADHERENCIA_HTA ~ PHQ9_TOTAL + {_STRUCT_COVARS} + {_ALT}",
    "submodel_domain_bp": f"{_MODEL2_NOALT} + {_ALT}",
    "sensitivity_no_2020": f"{_MODEL2_NOALT} + {_ALT}",
    "sensitivity_second_bp_measure": (
        f"PRESION_ARTERIAL_ELEVADA_SEGUNDA_TOMA ~ PHQ9_TOTAL + {_STRUCT_COVARS} + {_ALT}"
    ),
    # Sensibilidad S1: desenlace con umbral estilo ACC/AHA 2017 (>=130/80) en lugar
    # de >=140/90; misma estructura de ajuste que model_2 (incluye altitud).
    "sensitivity_outcome_130_80": (
        f"PRESION_ARTERIAL_ELEVADA_130_80 ~ PHQ9_TOTAL + {_STRUCT_COVARS} + {_ALT}"
    ),
    # --- Panel de modificacion de efecto (interaccion PHQ-9 x modificador) ---
    "interaction_sex": (
        f"PRESION_ARTERIAL_ELEVADA ~ PHQ9_TOTAL * factor(QSSEXO) + QS23 + factor(QS25N) + "
        f"factor(HV025) + factor(HV270) + factor(VIOLENCIA_PAREJA) + factor(ANIO) + {_ALT}"
    ),
    "interaction_year": (
        f"PRESION_ARTERIAL_ELEVADA ~ PHQ9_TOTAL * factor(ANIO) + QS23 + factor(QSSEXO) + "
        f"factor(QS25N) + factor(HV025) + factor(HV270) + factor(VIOLENCIA_PAREJA) + {_ALT}"
    ),
    "interaction_area": (
        f"PRESION_ARTERIAL_ELEVADA ~ PHQ9_TOTAL * factor(HV025) + QS23 + factor(QSSEXO) + "
        f"factor(QS25N) + factor(HV270) + factor(VIOLENCIA_PAREJA) + factor(ANIO) + {_ALT}"
    ),
    "interaction_riqueza": (
        f"PRESION_ARTERIAL_ELEVADA ~ PHQ9_TOTAL * factor(HV270) + QS23 + factor(QSSEXO) + "
        f"factor(QS25N) + factor(HV025) + factor(VIOLENCIA_PAREJA) + factor(ANIO) + {_ALT}"
    ),
    "interaction_dxhta": (
        f"PRESION_ARTERIAL_ELEVADA ~ PHQ9_TOTAL * factor(DX_HTA_PREVIO) + {_STRUCT_COVARS} + {_ALT}"
    ),
    "interaction_altitud": (
        f"PRESION_ARTERIAL_ELEVADA ~ PHQ9_TOTAL * {_ALT} + {_STRUCT_COVARS}"
    ),
}

MODEL_SUBSETS = {
    "model_1": "",
    "model_2": "",
    "model_3": "",
    "submodel_adherence": "DX_HTA_PREVIO == 1",
    "submodel_domain_bp": "DX_HTA_PREVIO == 1",
    "sensitivity_no_2020": "ANIO != 2020",
    "sensitivity_second_bp_measure": "PRESION_ARTERIAL_ELEVADA_SEGUNDA_TOMA %in% c(0, 1)",
    "sensitivity_outcome_130_80": "PRESION_ARTERIAL_ELEVADA_130_80 %in% c(0, 1)",
    "interaction_sex": "",
    "interaction_year": "",
    "interaction_area": "",
    "interaction_riqueza": "",
    "interaction_dxhta": "",
    "interaction_altitud": "",
}

# Descomposicion jerarquica del PR de PHQ-9 (aporte por bloque). h2 = Modelo 2
# preespecificado SIN altitud (sirve de modelo suplementario de trazabilidad de la enmienda);
# h3 = Modelo 2 con altitud (= MAIN_MODELS["model_2"]).
HIERARCHICAL_MODELS = {
    "h0_crudo": "PRESION_ARTERIAL_ELEVADA ~ PHQ9_TOTAL",
    "h1_demografia": "PRESION_ARTERIAL_ELEVADA ~ PHQ9_TOTAL + QS23 + factor(QSSEXO)",
    "h2_estructural_sin_altitud": _MODEL2_NOALT,
    "h3_estructural_con_altitud": f"{_MODEL2_NOALT} + {_ALT}",
}

LOGISTIC_SENSITIVITY_MODELS = {
    "model_2_logistic_sensitivity": MAIN_MODELS["model_2"],
}
MAIN_MODEL_ORDER = ["model_1", "model_2", "model_3"]
CASCADE_MODEL_ORDER = ["submodel_adherence", "submodel_domain_bp"]
SENSITIVITY_MODEL_ORDER = [
    "sensitivity_no_2020",
    "sensitivity_second_bp_measure",
    "sensitivity_outcome_130_80",
]
# Panel de modificacion de efecto. Preespecificados: sexo, anio. Exploratorios (con
# correccion por multiplicidad): area, riqueza, DX_HTA previo, altitud.
EFFECT_MOD_MODELS = [
    "interaction_sex",
    "interaction_year",
    "interaction_area",
    "interaction_riqueza",
    "interaction_dxhta",
    "interaction_altitud",
]
EFFECT_MOD_PRESPECIFIED = {"interaction_sex", "interaction_year"}
# Compatibilidad con interactions_and_sensitivity_models.csv: sensibilidades + interacciones.
INTERACTION_MODEL_ORDER = SENSITIVITY_MODEL_ORDER + EFFECT_MOD_MODELS
R_SECTION_CHOICES = ("tables", "models", "figures")

# --- Subplan A: altitud ---
# Robustez de forma funcional: altitud continua (km) en lugar de la categorica del
# ajuste estandar (model_2). Usa el conjunto estructural SIN la categorica de altitud.
ALTITUDE_ADJUSTED_MODELS = {
    "model_2_altitud_km_robustez": f"{_MODEL2_NOALT} + ALTITUD_KM",
}
# Estratificacion de model_2 por altitud. IMPORTANTE: dentro de un estrato la altitud
# es (casi) constante, por lo que se usa el Modelo 2 SIN el termino de altitud
# (_MODEL2_NOALT); incluir factor(ALTITUD_CAT3) seria rank-deficiente. HV040 completo => sin NA.
ALTITUDE_STRATIFIED_MODELS = {
    "model_2_estrato_lt2500": ("ALTITUD_ALTA_2500 == 0", _MODEL2_NOALT),
    "model_2_estrato_ge2500": ("ALTITUD_ALTA_2500 == 1", _MODEL2_NOALT),
    "model_2_estrato_cat3_lt1500": ("ALTITUD_CAT3 == '<1500'", _MODEL2_NOALT),
    "model_2_estrato_cat3_1500_2499": ("ALTITUD_CAT3 == '1500-2499'", _MODEL2_NOALT),
    "model_2_estrato_cat3_ge2500": ("ALTITUD_CAT3 == '>=2500'", _MODEL2_NOALT),
}
# La modificacion de efecto por altitud se evalua via interaction_altitud (MAIN_MODELS),
# integrada al panel de modificacion de efecto (EFFECT_MOD_MODELS).
# Umbral minimo de adecuacion de estrato antes de interpretar heterogeneidad.
ALTITUDE_MIN_OBS = 500
ALTITUDE_MIN_EVENTS = 100

# Items/analisis omitidos formalmente del estudio (documentados en
# analysis_skipped_items.csv). No son bugs: son decisiones por datos insuficientes.
ANALYSIS_SKIPPED_ITEMS = [
    {
        "item": "TIEMPO_DX_HTA_MESES",
        "reason": (
            "Excluida formalmente del estudio: QS103U/QS103C vacios o no utilizables "
            "en los datos crudos."
        ),
    },
    {
        "item": "interaccion_paradoja_del_cuidado",
        "reason": (
            "No ejecutable por datos insuficientes: el analisis temporal de la 'paradoja "
            "del cuidado' (modificacion del efecto PHQ-9 -> PA elevada segun tiempo desde "
            "el diagnostico y control del tratamiento antihipertensivo) requiere "
            "TIEMPO_DX_HTA_MESES, ausente en RAW (QS103U/QS103C vacios). Los submodelos de "
            "cascada (submodel_adherence, submodel_domain_bp sobre DX_HTA_PREVIO==1) son un "
            "proxy transversal; la heterogeneidad temporal por deteccion/cuidado no es estimable."
        ),
    },
]

SPLINE_ADJUSTERS = (
    "QS23 + factor(QSSEXO) + factor(QS25N) + factor(HV025) + factor(HV270) + factor(VIOLENCIA_PAREJA) + factor(ANIO)"
)


@dataclass
class AnalysisSettings:
    analysis_output_dir: Path
    imputed_dir: Path
    tables_dir: Path
    models_dir: Path
    figures_dir: Path
    mice_num_imputations: int
    random_seed: int
    write_imputed_csv: bool
    r_home: Optional[str]
    r_lib_user: Optional[str]
    r_sections: Tuple[str, ...]


def build_analysis_settings(output_dir: Path, analysis_settings: Dict[str, object]) -> AnalysisSettings:
    analysis_root = output_dir / "analysis"
    root = Path(str(analysis_settings.get("analysis_output_dir", analysis_root)))
    imputed = Path(str(analysis_settings.get("imputed_dir", root / "imputed")))
    tables = Path(str(analysis_settings.get("tables_dir", root / "tables")))
    models = Path(str(analysis_settings.get("models_dir", root / "models")))
    figures = Path(str(analysis_settings.get("figures_dir", root / "figures")))

    return AnalysisSettings(
        analysis_output_dir=root,
        imputed_dir=imputed,
        tables_dir=tables,
        models_dir=models,
        figures_dir=figures,
        mice_num_imputations=int(analysis_settings.get("mice_num_imputations", 20)),
        random_seed=int(analysis_settings.get("random_seed", 20260313)),
        write_imputed_csv=bool(analysis_settings.get("write_imputed_csv", False)),
        r_home=_maybe_str(analysis_settings.get("r_home")),
        r_lib_user=_maybe_str(analysis_settings.get("r_lib_user")),
        r_sections=_normalize_r_sections(analysis_settings.get("r_sections")),
    )


def run_analysis(
    unified: pd.DataFrame,
    settings: AnalysisSettings,
    run_mice: bool,
    run_vif: bool,
    run_r_bridge: bool,
) -> None:
    _ensure_dirs(settings)

    if run_mice:
        imputed_paths = _run_mice(unified, settings)
    else:
        imputed_paths = sorted(settings.imputed_dir.glob("imputation_*.parquet"))
        if not imputed_paths:
            raise FileNotFoundError(
                f"No se encontraron imputaciones en {settings.imputed_dir}; active run_mice o regenere los datasets imputados."
            )

    first_imputed = pd.read_parquet(imputed_paths[0])
    if run_vif:
        _compute_vif(first_imputed, settings)

    if run_r_bridge:
        _run_r_bridge(imputed_paths, settings)


def _ensure_dirs(settings: AnalysisSettings) -> None:
    for path in [
        settings.analysis_output_dir,
        settings.imputed_dir,
        settings.tables_dir,
        settings.models_dir,
        settings.figures_dir,
    ]:
        path.mkdir(parents=True, exist_ok=True)


def _normalize_r_sections(raw_sections: object) -> Tuple[str, ...]:
    if raw_sections is None:
        return R_SECTION_CHOICES

    if isinstance(raw_sections, str):
        items = [item.strip().lower() for item in raw_sections.split(",") if item.strip()]
    else:
        items = [str(item).strip().lower() for item in raw_sections if str(item).strip()]

    if not items or "all" in items:
        return R_SECTION_CHOICES

    invalid = sorted({item for item in items if item not in R_SECTION_CHOICES})
    if invalid:
        allowed = ", ".join(R_SECTION_CHOICES)
        raise ValueError(f"Secciones R invalidas: {', '.join(invalid)}. Use solo: {allowed}.")

    return tuple(dict.fromkeys(items))


def _run_mice(unified: pd.DataFrame, settings: AnalysisSettings) -> List[Path]:
    logging.info(
        "Iniciando MICE con %s imputaciones. Targets=%s. Excluidas por skip pattern=%s",
        settings.mice_num_imputations,
        ", ".join(MICE_TARGETS),
        ", ".join(NON_IMPUTED_SKIP_COVARS),
    )

    matrix_cols = list(dict.fromkeys([*MICE_TARGETS, *MICE_PREDICTORS]))
    matrix = unified[matrix_cols].copy()
    for col in matrix_cols:
        matrix[col] = pd.to_numeric(matrix[col], errors="coerce")

    missing_before = matrix[MICE_TARGETS].isna().sum()
    summary_rows: List[Dict[str, object]] = []
    imputed_paths: List[Path] = []

    observed_ranges = {
        col: (
            float(matrix[col].dropna().min()) if matrix[col].notna().any() else np.nan,
            float(matrix[col].dropna().max()) if matrix[col].notna().any() else np.nan,
        )
        for col in CONTINUOUS_IMPUTE_COLS
    }
    allowed_categories = {
        col: np.sort(matrix[col].dropna().unique().astype(float))
        for col in CATEGORICAL_IMPUTE_COLS
    }

    for imputation_id in range(1, settings.mice_num_imputations + 1):
        random_state = settings.random_seed + imputation_id
        imputer = IterativeImputer(
            random_state=random_state,
            sample_posterior=True,
            max_iter=20,
            initial_strategy="most_frequent",
            skip_complete=True,
        )
        imputed_matrix = pd.DataFrame(
            imputer.fit_transform(matrix),
            columns=matrix_cols,
            index=matrix.index,
        )

        completed = unified.copy()
        for col in CONTINUOUS_IMPUTE_COLS:
            missing_mask = completed[col].isna()
            if not missing_mask.any():
                continue
            low, high = observed_ranges[col]
            values = imputed_matrix.loc[missing_mask, col]
            completed.loc[missing_mask, col] = values.clip(lower=low, upper=high)

        for col in CATEGORICAL_IMPUTE_COLS:
            missing_mask = completed[col].isna()
            if not missing_mask.any():
                continue
            allowed = allowed_categories[col]
            if allowed.size == 0:
                continue
            completed.loc[missing_mask, col] = _nearest_allowed(
                imputed_matrix.loc[missing_mask, col].to_numpy(dtype=float),
                allowed,
            )

        for col in CATEGORICAL_IMPUTE_COLS:
            completed[col] = pd.to_numeric(completed[col], errors="coerce").round().astype("Int64")

        output_path = settings.imputed_dir / f"imputation_{imputation_id:02d}.parquet"
        completed.to_parquet(output_path, index=False)
        if settings.write_imputed_csv:
            completed.to_csv(
                settings.imputed_dir / f"imputation_{imputation_id:02d}.csv",
                index=False,
                encoding="utf-8-sig",
            )

        imputed_paths.append(output_path)
        missing_after = completed[MICE_TARGETS].isna().sum()
        for variable in MICE_TARGETS:
            summary_rows.append(
                {
                    "imputation_id": imputation_id,
                    "variable": variable,
                    "missing_before": int(missing_before[variable]),
                    "missing_after": int(missing_after[variable]),
                    "random_seed": random_state,
                }
            )

    pd.DataFrame(summary_rows).to_csv(
        settings.analysis_output_dir / "mice_missingness_summary.csv",
        index=False,
        encoding="utf-8-sig",
    )

    manifest = {
        "mice_num_imputations": settings.mice_num_imputations,
        "targets": MICE_TARGETS,
        "predictors": MICE_PREDICTORS,
        "seed_base": settings.random_seed,
        "imputed_paths": [str(path) for path in imputed_paths],
    }
    (settings.analysis_output_dir / "mice_manifest.json").write_text(
        json.dumps(manifest, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )

    logging.info("MICE finalizado: %s datasets imputados", len(imputed_paths))
    return imputed_paths


def _compute_vif(imputed_df: pd.DataFrame, settings: AnalysisSettings) -> None:
    logging.info("Calculando VIF sobre la primera imputacion")

    vif_df = imputed_df[VIF_PREDICTORS].copy()
    continuous_cols = {"PHQ9_TOTAL", "QS23", "IMC", "QS907"}
    categorical_cols = [col for col in VIF_PREDICTORS if col not in continuous_cols]

    for col in categorical_cols:
        vif_df[col] = vif_df[col].astype("string")

    design_matrix = pd.get_dummies(vif_df, columns=categorical_cols, drop_first=True, dtype=float)
    design_matrix = design_matrix.loc[:, design_matrix.nunique(dropna=False) > 1]
    design_matrix = design_matrix.dropna(axis=0)
    design_matrix = add_constant(design_matrix, has_constant="add")

    if design_matrix.empty:
        raise ValueError("La matriz VIF quedo vacia tras eliminar filas con NA.")

    logging.info("VIF calculado sobre %s filas completas", len(design_matrix))

    rows = []
    values = design_matrix.to_numpy(dtype=float)
    columns = list(design_matrix.columns)
    for idx, column in enumerate(columns):
        if column == "const":
            continue
        rows.append({"predictor": column, "vif": float(variance_inflation_factor(values, idx))})

    pd.DataFrame(rows).sort_values("vif", ascending=False).to_csv(
        settings.analysis_output_dir / "vif_first_imputation.csv",
        index=False,
        encoding="utf-8-sig",
    )


def _run_r_bridge(imputed_paths: Sequence[Path], settings: AnalysisSettings) -> None:
    sections = set(settings.r_sections)
    run_tables = "tables" in sections
    run_models = "models" in sections
    run_figures = "figures" in sections

    logging.info(
        "Iniciando puente analitico con R para secciones: %s",
        ", ".join(settings.r_sections),
    )

    ro, pandas2ri, localconverter = _load_r_bridge(settings)
    ro.r(_R_HELPERS)

    model_results: List[pd.DataFrame] = []
    model_meta_rows: List[Dict[str, object]] = []
    logistic_results: List[pd.DataFrame] = []
    logistic_meta_rows: List[Dict[str, object]] = []
    table1_parts: List[pd.DataFrame] = []
    deff_parts: List[pd.DataFrame] = []
    rao_scott_parts: List[pd.DataFrame] = []
    bivariate_term_fits: List[Dict[str, object]] = []
    spline_curves: List[pd.DataFrame] = []
    spline_meta: List[Dict[str, object]] = []
    altitude_adjusted: List[pd.DataFrame] = []
    altitude_adjusted_meta: List[Dict[str, object]] = []
    altitude_stratified: List[pd.DataFrame] = []
    altitude_strat_meta: List[Dict[str, object]] = []
    effect_mod_fits: Dict[str, List[Dict[str, object]]] = {m: [] for m in EFFECT_MOD_MODELS}
    hierarchical_results: List[pd.DataFrame] = []
    skipped_rows: List[Dict[str, object]] = []
    if run_models:
        for skipped_item in ANALYSIS_SKIPPED_ITEMS:
            skipped_rows.append(
                {
                    "imputation_id": pd.NA,
                    "item": skipped_item["item"],
                    "reason": skipped_item["reason"],
                }
            )

    for imputation_id, path in enumerate(imputed_paths, start=1):
        logging.info(
            "Puente R | imputacion %s/%s | archivo=%s | secciones=%s",
            imputation_id,
            len(imputed_paths),
            path.name,
            ", ".join(settings.r_sections),
        )
        df = pd.read_parquet(path)
        r_df = _to_r_dataframe(df, pandas2ri, localconverter)

        if run_models:
            for model_name, formula in MAIN_MODELS.items():
                fit = ro.globalenv["fit_svyglm_model"](r_df, formula, MODEL_SUBSETS[model_name])
                model_results.append(_model_fit_to_frame(fit, imputation_id, model_name))
                model_meta_rows.append(_model_meta_to_row(fit, imputation_id, model_name))
                # Panel de modificacion de efecto: capturar terminos de interaccion para Wald conjunto.
                if model_name in effect_mod_fits:
                    effect_mod_fits[model_name].append(
                        _interaction_term_fit(fit, imputation_id, model_name)
                    )

            for model_name, formula in LOGISTIC_SENSITIVITY_MODELS.items():
                fit = ro.globalenv["fit_logistic_model"](r_df, formula)
                logistic_results.append(_model_fit_to_frame(fit, imputation_id, model_name))
                logistic_meta_rows.append(_logistic_meta_to_row(fit, imputation_id, model_name))

            # Descomposicion jerarquica del PR de PHQ-9 (aporte por bloque; h2 = preespecif. sin altitud).
            for model_name, formula in HIERARCHICAL_MODELS.items():
                fit = ro.globalenv["fit_svyglm_model"](r_df, formula, "")
                hierarchical_results.append(_model_fit_to_frame(fit, imputation_id, model_name))

            # Subplan A: altitud — robustez (forma continua) y estratificacion.
            for model_name, formula in ALTITUDE_ADJUSTED_MODELS.items():
                fit = ro.globalenv["fit_svyglm_model"](r_df, formula, "")
                altitude_adjusted.append(_model_fit_to_frame(fit, imputation_id, model_name))
                altitude_adjusted_meta.append(
                    _altitude_meta_to_row(fit, imputation_id, model_name, stratum="overall")
                )

            for model_name, (subset_expr, formula) in ALTITUDE_STRATIFIED_MODELS.items():
                fit = ro.globalenv["fit_svyglm_model"](r_df, formula, subset_expr)
                altitude_stratified.append(_model_fit_to_frame(fit, imputation_id, model_name))
                altitude_strat_meta.append(
                    _altitude_meta_to_row(fit, imputation_id, model_name, stratum=model_name)
                )
            logging.info("Puente R | imputacion %s/%s | modelos listos", imputation_id, len(imputed_paths))

        if run_tables:
            for summary_type, group_var in TABLE1_GROUPS:
                summary = ro.globalenv["summarize_table1_bundle"](
                    r_df,
                    ro.StrVector(TABLE1_CONTINUOUS),
                    ro.StrVector(TABLE1_CATEGORICAL),
                    group_var,
                )
                table1_parts.append(
                    _summary_to_frame(
                        summary,
                        imputation_id,
                        pandas2ri=pandas2ri,
                        localconverter=localconverter,
                        summary_type=summary_type,
                    )
                )

            deff_parts.append(
                _summary_to_frame(
                    ro.globalenv["compute_deff_metrics"](r_df),
                    imputation_id,
                    pandas2ri=pandas2ri,
                    localconverter=localconverter,
                    summary_type="deff",
                )
            )

            rao_scott_parts.append(
                _summary_to_frame(
                    ro.globalenv["rao_scott_bundle"](r_df, ro.StrVector(RAO_SCOTT_VARS)),
                    imputation_id,
                    pandas2ri=pandas2ri,
                    localconverter=localconverter,
                )
            )

            bivariate_term_fits.extend(
                _term_bundle_to_dicts(
                    ro.globalenv["fit_bivariate_term_bundle"](r_df, ro.StrVector(BIVARIATE_TEST_VARS)),
                    imputation_id,
                )
            )
            logging.info("Puente R | imputacion %s/%s | tablas listas", imputation_id, len(imputed_paths))

        if run_figures:
            knots = _select_spline_knots(df["PHQ9_TOTAL"])
            reference_row = _build_spline_reference(df)
            r_ref = _to_r_dataframe(reference_row, pandas2ri, localconverter)
            spline_fit = ro.globalenv["fit_spline_model"](r_df, ro.FloatVector(knots), SPLINE_ADJUSTERS, r_ref)
            spline_curves.append(_spline_curve_to_frame(spline_fit, imputation_id))
            spline_meta.append(_spline_meta_to_row(spline_fit, imputation_id, knots))
            logging.info("Puente R | imputacion %s/%s | figuras listas", imputation_id, len(imputed_paths))

    if run_models and model_results:
        pooled_models = _pool_model_results(pd.concat(model_results, ignore_index=True))
        pooled_models.to_csv(
            settings.models_dir / "pooled_model_results.csv",
            index=False,
            encoding="utf-8-sig",
        )
        pooled_models.loc[pooled_models["model"].isin(MAIN_MODEL_ORDER)].to_csv(
            settings.models_dir / "table3_main_models.csv",
            index=False,
            encoding="utf-8-sig",
        )
        pooled_models.loc[pooled_models["model"].isin(CASCADE_MODEL_ORDER)].to_csv(
            settings.models_dir / "table4_cascade_models.csv",
            index=False,
            encoding="utf-8-sig",
        )
        pooled_models.loc[pooled_models["model"].isin(INTERACTION_MODEL_ORDER)].to_csv(
            settings.models_dir / "interactions_and_sensitivity_models.csv",
            index=False,
            encoding="utf-8-sig",
        )

        model_meta_df = pd.DataFrame(model_meta_rows).sort_values(["model", "imputation_id"])
        model_meta_df.to_csv(
            settings.models_dir / "model_diagnostics_per_imputation.csv",
            index=False,
            encoding="utf-8-sig",
        )
        _pool_model_meta(model_meta_df).to_csv(
            settings.models_dir / "model_diagnostics_summary.csv",
            index=False,
            encoding="utf-8-sig",
        )

        pd.DataFrame(skipped_rows, columns=["imputation_id", "item", "reason"]).drop_duplicates(
            subset=["item", "reason"]
        ).to_csv(
            settings.models_dir / "analysis_skipped_items.csv",
            index=False,
            encoding="utf-8-sig",
        )

    if run_models and logistic_results:
        pooled_logistic = _pool_model_results(pd.concat(logistic_results, ignore_index=True))
        pooled_logistic.to_csv(
            settings.models_dir / "logistic_sensitivity_results.csv",
            index=False,
            encoding="utf-8-sig",
        )
    if run_models and logistic_meta_rows:
        logistic_meta_df = pd.DataFrame(logistic_meta_rows).sort_values(["model", "imputation_id"])
        logistic_meta_df.to_csv(
            settings.models_dir / "logistic_sensitivity_diagnostics_per_imputation.csv",
            index=False,
            encoding="utf-8-sig",
        )
        _pool_logistic_meta(logistic_meta_df).to_csv(
            settings.models_dir / "logistic_sensitivity_diagnostics_summary.csv",
            index=False,
            encoding="utf-8-sig",
        )

    # --- Subplan A: outputs de altitud ---
    if run_models and altitude_adjusted:
        _pool_model_results(pd.concat(altitude_adjusted, ignore_index=True)).to_csv(
            settings.models_dir / "altitude_adjusted_models.csv",
            index=False,
            encoding="utf-8-sig",
        )
        pd.DataFrame(altitude_adjusted_meta).sort_values(["model", "imputation_id"]).to_csv(
            settings.models_dir / "altitude_adjusted_diagnostics_per_imputation.csv",
            index=False,
            encoding="utf-8-sig",
        )

    if run_models and altitude_stratified:
        _pool_model_results(pd.concat(altitude_stratified, ignore_index=True)).to_csv(
            settings.models_dir / "altitude_stratified_models.csv",
            index=False,
            encoding="utf-8-sig",
        )
        _build_altitude_strata_adequacy(altitude_strat_meta).to_csv(
            settings.models_dir / "altitude_strata_adequacy.csv",
            index=False,
            encoding="utf-8-sig",
        )

    if run_models and any(effect_mod_fits.values()):
        _build_effect_modification_panel(effect_mod_fits).to_csv(
            settings.models_dir / "effect_modification_panel.csv",
            index=False,
            encoding="utf-8-sig",
        )

    if run_models and hierarchical_results:
        _pool_model_results(pd.concat(hierarchical_results, ignore_index=True)).to_csv(
            settings.models_dir / "hierarchical_decomposition.csv",
            index=False,
            encoding="utf-8-sig",
        )

    if run_tables and table1_parts:
        pooled_table1 = _pool_summary_results(pd.concat(table1_parts, ignore_index=True))
        pooled_table1.to_csv(
            settings.tables_dir / "table1_weighted_summary.csv",
            index=False,
            encoding="utf-8-sig",
        )
        deff_df = pd.concat(deff_parts, ignore_index=True)
        deff_df.to_csv(
            settings.tables_dir / "table1_deff_per_imputation.csv",
            index=False,
            encoding="utf-8-sig",
        )
        _build_deff_summary(pooled_table1, deff_df).to_csv(
            settings.tables_dir / "table1_deff_summary.csv",
            index=False,
            encoding="utf-8-sig",
        )

        rao_scott_df = pd.concat(rao_scott_parts, ignore_index=True).sort_values(["variable", "imputation_id"])
        rao_scott_df.to_csv(
            settings.tables_dir / "rao_scott_per_imputation.csv",
            index=False,
            encoding="utf-8-sig",
        )
        _summarize_rao_scott(rao_scott_df).to_csv(
            settings.tables_dir / "rao_scott_summary.csv",
            index=False,
            encoding="utf-8-sig",
        )
        _pool_bivariate_term_tests(bivariate_term_fits, rao_scott_df).to_csv(
            settings.tables_dir / "table2_bivariate_pooled.csv",
            index=False,
            encoding="utf-8-sig",
        )

    if run_figures and spline_curves:
        spline_curve_parts = pd.concat(spline_curves, ignore_index=True)
        spline_curve_parts.to_csv(
            settings.figures_dir / "spline_per_imputation.csv",
            index=False,
            encoding="utf-8-sig",
        )

        spline_curve_df = _pool_summary_results(spline_curve_parts)
        spline_curve_df.to_csv(
            settings.figures_dir / "spline_curve_pooled.csv",
            index=False,
            encoding="utf-8-sig",
        )

        spline_meta_df = pd.DataFrame(spline_meta)
        spline_meta_df.to_csv(
            settings.figures_dir / "spline_nonlinearity_per_imputation.csv",
            index=False,
            encoding="utf-8-sig",
        )
        _pool_spline_nonlinearity(spline_meta_df).to_csv(
            settings.figures_dir / "spline_nonlinearity_summary.csv",
            index=False,
            encoding="utf-8-sig",
        )

    logging.info("Puente analitico con R finalizado")


def _load_r_bridge(settings: AnalysisSettings):
    r_home = _resolve_r_home(settings.r_home)
    os.environ["R_HOME"] = r_home

    path_parts = [str(Path(r_home) / "bin" / "x64"), str(Path(r_home) / "bin")]
    current_path = os.environ.get("PATH", "")
    for part in reversed(path_parts):
        if part not in current_path:
            current_path = part + os.pathsep + current_path
    os.environ["PATH"] = current_path

    if settings.r_lib_user:
        os.environ["R_LIBS_USER"] = settings.r_lib_user
    else:
        os.environ.setdefault(
            "R_LIBS_USER",
            str(Path.home() / "Documents" / "R" / "win-library" / "4.5"),
        )

    import rpy2.robjects as ro
    from rpy2.robjects import pandas2ri
    from rpy2.robjects.conversion import localconverter
    import rpy2.robjects.packages as rpackages

    rpackages.importr("survey")
    rpackages.importr("rms")
    rpackages.importr("jsonlite")
    return ro, pandas2ri, localconverter


def _to_r_dataframe(df: pd.DataFrame, pandas2ri, localconverter):
    prepared = df.copy()
    for col in prepared.columns:
        if pd.api.types.is_integer_dtype(prepared[col].dtype):
            prepared[col] = prepared[col].astype("float64") if prepared[col].isna().any() else prepared[col].astype("int64")
        elif pd.api.types.is_string_dtype(prepared[col].dtype):
            prepared[col] = prepared[col].astype(object)
    with localconverter(pandas2ri.converter):
        return pandas2ri.py2rpy(prepared)


def _to_pandas_frame(r_frame, pandas2ri, localconverter) -> pd.DataFrame:
    with localconverter(pandas2ri.converter):
        return pandas2ri.rpy2py(r_frame)


def _model_fit_to_frame(fit, imputation_id: int, model_name: str) -> pd.DataFrame:
    coefs = np.asarray(fit.rx2("coef"), dtype=float)
    coef_names = list(fit.rx2("coef").names)
    vcov = np.asarray(fit.rx2("vcov"), dtype=float)
    n_obs = int(np.asarray(fit.rx2("n_obs"), dtype=int)[0])

    rows = []
    for idx, term in enumerate(coef_names):
        rows.append(
            {
                "imputation_id": imputation_id,
                "model": model_name,
                "term": term,
                "estimate": coefs[idx],
                "variance": vcov[idx, idx],
                "n_obs": n_obs,
            }
        )
    return pd.DataFrame(rows)


def _model_meta_to_row(fit, imputation_id: int, model_name: str) -> Dict[str, object]:
    return {
        "imputation_id": imputation_id,
        "model": model_name,
        "family": _extract_r_text(fit.rx2("family")),
        "link": _extract_r_text(fit.rx2("link")),
        "n_obs": int(np.asarray(fit.rx2("n_obs"), dtype=int)[0]),
        "df_resid": _extract_r_scalar(fit.rx2("df_resid")),
        "scale_deviance": _extract_r_scalar(fit.rx2("scale_deviance")),
    }


def _logistic_meta_to_row(fit, imputation_id: int, model_name: str) -> Dict[str, object]:
    row = _model_meta_to_row(fit, imputation_id, model_name)
    row.update(
        {
            "gof_method": _extract_r_text(fit.rx2("gof_method")),
            "gof_statistic": _extract_r_scalar(fit.rx2("gof_statistic")),
            "gof_df": _extract_r_scalar(fit.rx2("gof_df")),
            "gof_df_den": _extract_r_scalar(fit.rx2("gof_df_den")),
            "gof_n_groups": _extract_r_scalar(fit.rx2("gof_n_groups")),
            "gof_p_value": _extract_r_scalar(fit.rx2("gof_p_value")),
        }
    )
    return row


def _altitude_meta_to_row(fit, imputation_id: int, model_name: str, stratum: str) -> Dict[str, object]:
    row = _model_meta_to_row(fit, imputation_id, model_name)
    n_events = np.nan
    try:
        n_events = _extract_r_scalar(fit.rx2("n_events"))
    except Exception:  # pragma: no cover - n_events solo en fit_svyglm_model
        n_events = np.nan
    row["stratum"] = stratum
    row["n_events"] = n_events
    return row


def _interaction_term_fit(fit, imputation_id: int, model_name: str) -> Dict[str, object]:
    """Extrae los terminos de interaccion (que contienen ':') con su submatriz de
    covarianza para el test de Wald conjunto pooled via _pool_joint_wald."""
    coefs = np.asarray(fit.rx2("coef"), dtype=float)
    names = list(fit.rx2("coef").names)
    vcov = np.asarray(fit.rx2("vcov"), dtype=float)
    n_obs = int(np.asarray(fit.rx2("n_obs"), dtype=int)[0])

    idx = [i for i, name in enumerate(names) if ":" in name]
    terms = [names[i] for i in idx]
    sub_coef = coefs[idx] if idx else np.zeros(0, dtype=float)
    sub_vcov = vcov[np.ix_(idx, idx)] if idx else np.zeros((0, 0), dtype=float)
    return {
        "imputation_id": imputation_id,
        "variable": model_name,
        "terms": terms,
        "coef": sub_coef,
        "vcov": sub_vcov,
        "n_obs": n_obs,
    }


def _summary_to_frame(
    summary_obj,
    imputation_id: int,
    pandas2ri,
    localconverter,
    variable: Optional[str] = None,
    summary_type: Optional[str] = None,
) -> pd.DataFrame:
    frame = _to_pandas_frame(summary_obj, pandas2ri, localconverter)
    if frame.empty:
        return frame
    frame["imputation_id"] = imputation_id
    if variable is not None:
        frame["variable"] = variable
    if summary_type is not None:
        frame["summary_type"] = summary_type
    return frame


def _rao_scott_to_row(test_obj, imputation_id: int, variable: str) -> Dict[str, object]:
    return {
        "imputation_id": imputation_id,
        "variable": variable,
        "statistic": float(np.asarray(test_obj.rx2("statistic"), dtype=float)[0]),
        "df_num": float(np.asarray(test_obj.rx2("df_num"), dtype=float)[0]),
        "df_den": float(np.asarray(test_obj.rx2("df_den"), dtype=float)[0]),
        "p_value": float(np.asarray(test_obj.rx2("p_value"), dtype=float)[0]),
    }


def _term_fit_to_dict(term_fit, imputation_id: int, variable: str) -> Dict[str, object]:
    coef = np.asarray(term_fit.rx2("coef"), dtype=float)
    term_names = [str(value) for value in list(term_fit.rx2("terms"))]
    vcov = np.asarray(term_fit.rx2("vcov"), dtype=float)
    if coef.size == 0:
        vcov = np.zeros((0, 0), dtype=float)
    elif vcov.ndim == 0:
        vcov = np.asarray([[float(vcov)]], dtype=float)

    return {
        "imputation_id": imputation_id,
        "variable": variable,
        "terms": term_names,
        "coef": coef,
        "vcov": vcov,
        "n_obs": int(np.asarray(term_fit.rx2("n_obs"), dtype=int)[0]),
    }


def _term_bundle_to_dicts(bundle_json, imputation_id: int) -> List[Dict[str, object]]:
    raw = "".join(str(part) for part in list(bundle_json))
    payload = json.loads(raw) if raw else []
    if isinstance(payload, dict):
        payload = [payload]

    rows: List[Dict[str, object]] = []
    for item in payload:
        terms = item.get("terms", [])
        if isinstance(terms, str):
            terms = [terms]
        coef_raw = item.get("coef", [])
        if isinstance(coef_raw, (int, float)):
            coef_raw = [coef_raw]
        coef = np.asarray(coef_raw, dtype=float)

        vcov_raw = item.get("vcov", [])
        if isinstance(vcov_raw, dict):
            vcov_raw = [vcov_raw[key] for key in sorted(vcov_raw, key=lambda value: int(value) if str(value).isdigit() else str(value))]
        elif isinstance(vcov_raw, (int, float)):
            vcov_raw = [[vcov_raw]]
        vcov = np.asarray(vcov_raw, dtype=float)
        if coef.size == 0:
            vcov = np.zeros((0, 0), dtype=float)
        elif vcov.ndim == 1:
            vcov = np.atleast_2d(vcov)

        rows.append(
            {
                "imputation_id": imputation_id,
                "variable": str(item.get("variable", "")),
                "terms": [str(value) for value in terms],
                "coef": coef,
                "vcov": vcov,
                "n_obs": int(item.get("n_obs", 0)),
            }
        )
    return rows


def _spline_curve_to_frame(spline_fit, imputation_id: int) -> pd.DataFrame:
    return pd.DataFrame(
        {
            "imputation_id": imputation_id,
            "summary_type": "spline_curve",
            "variable": "PHQ9_TOTAL",
            "group": "overall",
            "label": pd.Series(np.asarray(spline_fit.rx2("grid"), dtype=float)).round(6).astype(str),
            "estimate": np.asarray(spline_fit.rx2("pred"), dtype=float),
            "pred_se": np.asarray(spline_fit.rx2("pred_se"), dtype=float),
            "lower_ci": np.asarray(spline_fit.rx2("lower_ci"), dtype=float),
            "upper_ci": np.asarray(spline_fit.rx2("upper_ci"), dtype=float),
            "variance": np.asarray(spline_fit.rx2("pred_se"), dtype=float) ** 2,
        }
    )


def _spline_meta_to_row(spline_fit, imputation_id: int, knots: Sequence[float]) -> Dict[str, object]:
    wald_stat = _extract_r_scalar(spline_fit.rx2("wald_stat"))
    df_num = _extract_r_scalar(spline_fit.rx2("df_num"))
    p_value = _extract_r_scalar(spline_fit.rx2("p_value"))
    return {
        "imputation_id": imputation_id,
        "statistic": wald_stat,
        "df_num": int(df_num) if pd.notna(df_num) else pd.NA,
        "p_value": p_value,
        "knots": ",".join(f"{value:.6g}" for value in knots),
    }


def _pool_model_results(model_results: pd.DataFrame) -> pd.DataFrame:
    rows: List[Dict[str, object]] = []
    grouped = model_results.groupby(["model", "term"], dropna=False)

    for (model, term), group in grouped:
        pooled = _pool_scalar(group["estimate"].to_numpy(), group["variance"].to_numpy())
        rr = np.exp(pooled["estimate"])
        rr_low = np.exp(pooled["ci_low"])
        rr_high = np.exp(pooled["ci_high"])
        rows.append(
            {
                "model": model,
                "term": term,
                "estimate": pooled["estimate"],
                "std_error": pooled["std_error"],
                "ci_low": pooled["ci_low"],
                "ci_high": pooled["ci_high"],
                "p_value": pooled["p_value"],
                "df": pooled["df"],
                "exp_estimate": rr,
                "exp_ci_low": rr_low,
                "exp_ci_high": rr_high,
                "mean_n_obs": round(group["n_obs"].mean(), 2),
            }
        )

    return pd.DataFrame(rows).sort_values(["model", "term"]).reset_index(drop=True)


def _pool_model_meta(model_meta_df: pd.DataFrame) -> pd.DataFrame:
    if model_meta_df.empty:
        return model_meta_df

    return (
        model_meta_df.groupby(["model", "family", "link"], as_index=False, dropna=False)
        .agg(
            imputations=("imputation_id", "nunique"),
            mean_n_obs=("n_obs", "mean"),
            mean_df_resid=("df_resid", "mean"),
            mean_scale_deviance=("scale_deviance", "mean"),
            median_scale_deviance=("scale_deviance", "median"),
            max_scale_deviance=("scale_deviance", "max"),
        )
        .sort_values("model")
        .reset_index(drop=True)
    )


def _build_altitude_strata_adequacy(meta_rows: List[Dict[str, object]]) -> pd.DataFrame:
    """Resume n y eventos por estrato de altitud y marca si el estrato tiene
    tamano suficiente para interpretar heterogeneidad (Subplan A / validacion)."""
    if not meta_rows:
        return pd.DataFrame(
            columns=["model", "stratum", "imputations", "mean_n_obs", "mean_n_events", "estrato_suficiente"]
        )
    df = pd.DataFrame(meta_rows)
    agg = (
        df.groupby(["model", "stratum"], as_index=False)
        .agg(
            imputations=("imputation_id", "nunique"),
            mean_n_obs=("n_obs", "mean"),
            mean_n_events=("n_events", "mean"),
        )
        .sort_values(["model", "stratum"])
        .reset_index(drop=True)
    )
    agg["estrato_suficiente"] = (
        (agg["mean_n_obs"] >= ALTITUDE_MIN_OBS) & (agg["mean_n_events"] >= ALTITUDE_MIN_EVENTS)
    )
    return agg


def _build_effect_modification_panel(effect_mod_fits: Dict[str, List[Dict[str, object]]]) -> pd.DataFrame:
    """Panel de modificacion de efecto: Wald conjunto pooled (Rubin D1) por modificador.
    Marca preespecificados vs exploratorios y aplica correccion de multiplicidad de Holm
    sobre los exploratorios (no se penaliza a los preespecificados sexo/anio)."""
    rows: List[Dict[str, object]] = []
    for model_name, fits in effect_mod_fits.items():
        if not fits:
            continue
        modifier = model_name.replace("interaction_", "")
        pooled = _pool_joint_wald(fits)
        rows.append(
            {
                "modificador": modifier,
                "model": model_name,
                "tipo": "preespecificado" if model_name in EFFECT_MOD_PRESPECIFIED else "exploratorio",
                "pool_method": "rubin_d1_chisq_approx",
                "terms": pooled["terms"],
                "df_num": pooled["df_num"],
                "statistic": pooled["statistic"],
                "p_value": pooled["p_value"],
                "imputations_used": pooled["imputations_used"],
                "mean_n_obs": pooled["mean_n_obs"],
            }
        )
    panel = pd.DataFrame(rows)
    if panel.empty:
        return panel

    # Holm sobre los exploratorios.
    panel["p_holm_exploratorios"] = pd.NA
    explor = panel["tipo"] == "exploratorio"
    sub = panel.loc[explor & panel["p_value"].notna()].sort_values("p_value")
    k = len(sub)
    holm_vals: Dict[int, float] = {}
    running_max = 0.0
    for rank, (idx, row) in enumerate(sub.iterrows()):
        adj = min((k - rank) * float(row["p_value"]), 1.0)
        running_max = max(running_max, adj)
        holm_vals[idx] = running_max
    for idx, val in holm_vals.items():
        panel.at[idx, "p_holm_exploratorios"] = val

    orden = {m: i for i, m in enumerate(EFFECT_MOD_MODELS)}
    panel["_o"] = panel["model"].map(orden).fillna(99)
    return panel.sort_values("_o").drop(columns="_o").reset_index(drop=True)


def _pool_logistic_meta(logistic_meta_df: pd.DataFrame) -> pd.DataFrame:
    if logistic_meta_df.empty:
        return logistic_meta_df

    return (
        logistic_meta_df.groupby(["model", "family", "link", "gof_method"], as_index=False, dropna=False)
        .agg(
            imputations=("imputation_id", "nunique"),
            mean_n_obs=("n_obs", "mean"),
            mean_df_resid=("df_resid", "mean"),
            mean_scale_deviance=("scale_deviance", "mean"),
            mean_gof_statistic=("gof_statistic", "mean"),
            mean_gof_df=("gof_df", "mean"),
            mean_gof_df_den=("gof_df_den", "mean"),
            mean_gof_n_groups=("gof_n_groups", "mean"),
            median_gof_p_value=("gof_p_value", "median"),
            mean_gof_p_value=("gof_p_value", "mean"),
            min_gof_p_value=("gof_p_value", "min"),
            max_gof_p_value=("gof_p_value", "max"),
        )
        .sort_values("model")
        .reset_index(drop=True)
    )


def _pool_summary_results(summary_results: pd.DataFrame) -> pd.DataFrame:
    if summary_results.empty:
        return summary_results

    rows: List[Dict[str, object]] = []
    group_cols = [col for col in ["summary_type", "kind", "variable", "group", "label"] if col in summary_results.columns]
    grouped = summary_results.groupby(group_cols, dropna=False)
    for group_key, chunk in grouped:
        pooled = _pool_scalar(chunk["estimate"].to_numpy(), chunk["variance"].to_numpy())
        row = dict(zip(group_cols, group_key if isinstance(group_key, tuple) else (group_key,)))
        row.update(
            {
                "estimate": pooled["estimate"],
                "std_error": pooled["std_error"],
                "ci_low": pooled["ci_low"],
                "ci_high": pooled["ci_high"],
                "p_value": pooled["p_value"],
                "df": pooled["df"],
            }
        )
        if "deff" in chunk.columns:
            row["mean_deff"] = float(chunk["deff"].dropna().mean()) if chunk["deff"].notna().any() else np.nan
        if "n_unweighted" in chunk.columns:
            row["mean_n_unweighted"] = float(chunk["n_unweighted"].mean())
        rows.append(row)

    sort_cols = [col for col in ["summary_type", "kind", "variable", "group", "label"] if col in rows[0]]
    return pd.DataFrame(rows).sort_values(sort_cols).reset_index(drop=True)


def _build_deff_summary(pooled_table1: pd.DataFrame, deff_df: pd.DataFrame) -> pd.DataFrame:
    rows: List[Dict[str, object]] = []

    if not pooled_table1.empty and "mean_deff" in pooled_table1.columns:
        overall = pooled_table1.loc[
            (pooled_table1["summary_type"] == "overall") & pooled_table1["mean_deff"].notna(),
            "mean_deff",
        ]
        rows.append(
            {
                "metric": "table1_overall_mean_deff",
                "value": float(overall.mean()) if not overall.empty else np.nan,
            }
        )

    if not deff_df.empty:
        grouped = (
            deff_df.groupby("metric", as_index=False, dropna=False)
            .agg(
                imputations=("imputation_id", "nunique"),
                value=("value", "mean"),
            )
            .sort_values("metric")
        )
        rows.extend(grouped.to_dict("records"))

    return pd.DataFrame(rows)


def _summarize_rao_scott(rao_scott_df: pd.DataFrame) -> pd.DataFrame:
    return (
        rao_scott_df.groupby("variable", as_index=False)
        .agg(
            imputations=("imputation_id", "nunique"),
            mean_statistic=("statistic", "mean"),
            median_p_value=("p_value", "median"),
            mean_p_value=("p_value", "mean"),
            min_p_value=("p_value", "min"),
            max_p_value=("p_value", "max"),
        )
        .sort_values("variable")
    )


def _pool_bivariate_term_tests(term_fits: List[Dict[str, object]], rao_scott_df: pd.DataFrame) -> pd.DataFrame:
    grouped: Dict[str, List[Dict[str, object]]] = {}
    for fit in term_fits:
        grouped.setdefault(str(fit["variable"]), []).append(fit)

    rao_lookup = {}
    if not rao_scott_df.empty:
        rao_summary = _summarize_rao_scott(rao_scott_df)
        rao_lookup = {str(row["variable"]): row for row in rao_summary.to_dict("records")}

    rows: List[Dict[str, object]] = []
    for variable, fits in grouped.items():
        pooled = _pool_joint_wald(fits)
        row = {
            "variable": variable,
            "pool_method": "rubin_d1_chisq_approx",
            **pooled,
        }
        if variable in rao_lookup:
            lookup = rao_lookup[variable]
            row["mean_rao_scott_statistic"] = lookup.get("mean_statistic")
            row["median_rao_scott_p_value"] = lookup.get("median_p_value")
            row["mean_rao_scott_p_value"] = lookup.get("mean_p_value")
        rows.append(row)

    if not rows:
        return pd.DataFrame(
            columns=[
                "variable",
                "pool_method",
                "terms",
                "df_num",
                "statistic",
                "p_value",
                "imputations_used",
                "mean_n_obs",
                "mean_rao_scott_statistic",
                "median_rao_scott_p_value",
                "mean_rao_scott_p_value",
            ]
        )

    return pd.DataFrame(rows).sort_values("variable").reset_index(drop=True)


def _pool_joint_wald(fits: List[Dict[str, object]]) -> Dict[str, object]:
    valid = [fit for fit in fits if fit["terms"]]
    if not valid:
        return {
            "terms": "",
            "df_num": 0,
            "statistic": np.nan,
            "p_value": np.nan,
            "imputations_used": 0,
            "mean_n_obs": np.nan,
        }

    common_terms = sorted(set(valid[0]["terms"]).intersection(*(set(fit["terms"]) for fit in valid[1:])))
    common_terms = [
        term
        for term in common_terms
        if all(
            np.isfinite(np.asarray(fit["coef"], dtype=float)[fit["terms"].index(term)])
            and np.isfinite(np.asarray(fit["vcov"], dtype=float)[fit["terms"].index(term), fit["terms"].index(term)])
            for fit in valid
        )
    ]
    if not common_terms:
        return {
            "terms": "",
            "df_num": 0,
            "statistic": np.nan,
            "p_value": np.nan,
            "imputations_used": 0,
            "mean_n_obs": np.nan,
        }

    coef_matrix = []
    vcov_mats = []
    mean_n_obs = []
    for fit in valid:
        index = {term: idx for idx, term in enumerate(fit["terms"])}
        selection = [index[term] for term in common_terms]
        coef_matrix.append(np.asarray(fit["coef"], dtype=float)[selection])
        vcov = np.asarray(fit["vcov"], dtype=float)
        vcov_mats.append(np.nan_to_num(vcov[np.ix_(selection, selection)], nan=0.0, posinf=0.0, neginf=0.0))
        mean_n_obs.append(float(fit["n_obs"]))

    q = len(common_terms)
    m = len(coef_matrix)
    q_matrix = np.vstack(coef_matrix)
    q_bar = q_matrix.mean(axis=0)
    u_bar = np.mean(np.stack(vcov_mats), axis=0)
    if m > 1:
        b_mat = np.atleast_2d(np.cov(q_matrix, rowvar=False, ddof=1))
        if q == 1:
            b_mat = np.asarray([[float(b_mat.squeeze())]], dtype=float)
    else:
        b_mat = np.zeros((q, q), dtype=float)
    total = u_bar + (1.0 + (1.0 / m)) * b_mat
    total = np.nan_to_num(total, nan=0.0, posinf=0.0, neginf=0.0)
    total = np.clip(total, -1e12, 1e12)
    if not np.isfinite(total).all():
        statistic = np.nan
        p_value = np.nan
    else:
        ridge = max(float(np.max(np.abs(np.diag(total)))) * 1e-8, 1e-8)
        total_reg = total + np.eye(q, dtype=float) * ridge
        with np.errstate(all="ignore"):
            try:
                solution = np.linalg.solve(total_reg, q_bar)
            except np.linalg.LinAlgError:
                solution = np.linalg.pinv(total_reg, rcond=1e-10) @ q_bar
            statistic = float(q_bar.T @ solution)
        p_value = float(1 - stats.chi2.cdf(statistic, df=q)) if np.isfinite(statistic) else np.nan

    return {
        "terms": ",".join(common_terms),
        "df_num": q,
        "statistic": statistic,
        "p_value": p_value,
        "imputations_used": m,
        "mean_n_obs": round(float(np.mean(mean_n_obs)), 2),
    }


def _pool_spline_nonlinearity(spline_meta_df: pd.DataFrame) -> pd.DataFrame:
    valid = spline_meta_df.loc[spline_meta_df["statistic"].notna()].copy()
    if valid.empty:
        return pd.DataFrame(
            [
                {
                    "mean_wald_statistic": np.nan,
                    "median_p_value": np.nan,
                    "mean_p_value": np.nan,
                    "min_p_value": np.nan,
                    "max_p_value": np.nan,
                    "rubin_style_mean_statistic": np.nan,
                    "n_imputations_with_test": 0,
                }
            ]
        )

    pooled = _pool_scalar(valid["statistic"].to_numpy(), np.zeros(len(valid)))
    return pd.DataFrame(
        [
            {
                "mean_wald_statistic": valid["statistic"].mean(),
                "median_p_value": valid["p_value"].median(),
                "mean_p_value": valid["p_value"].mean(),
                "min_p_value": valid["p_value"].min(),
                "max_p_value": valid["p_value"].max(),
                "rubin_style_mean_statistic": pooled["estimate"],
                "n_imputations_with_test": int(len(valid)),
            }
        ]
    )


def _pool_scalar(estimates: np.ndarray, variances: np.ndarray) -> Dict[str, float]:
    estimates = np.asarray(estimates, dtype=float)
    variances = np.asarray(variances, dtype=float)
    m = len(estimates)
    qbar = float(estimates.mean())
    ubar = float(variances.mean()) if variances.size else 0.0
    b = float(estimates.var(ddof=1)) if m > 1 else 0.0
    total_variance = ubar + (1.0 + (1.0 / m)) * b
    std_error = float(np.sqrt(max(total_variance, 0.0)))

    if std_error == 0:
        return {
            "estimate": qbar,
            "std_error": 0.0,
            "ci_low": qbar,
            "ci_high": qbar,
            "p_value": 0.0,
            "df": np.inf,
        }

    if b == 0:
        df = np.inf
        crit = stats.norm.ppf(0.975)
        p_value = 2 * (1 - stats.norm.cdf(abs(qbar / std_error)))
    else:
        lambda_value = ((1.0 + (1.0 / m)) * b) / total_variance
        df = (m - 1) / (lambda_value**2) if lambda_value > 0 else np.inf
        crit = stats.t.ppf(0.975, df) if np.isfinite(df) else stats.norm.ppf(0.975)
        p_value = (
            2 * (1 - stats.t.cdf(abs(qbar / std_error), df))
            if np.isfinite(df)
            else 2 * (1 - stats.norm.cdf(abs(qbar / std_error)))
        )

    return {
        "estimate": qbar,
        "std_error": std_error,
        "ci_low": qbar - crit * std_error,
        "ci_high": qbar + crit * std_error,
        "p_value": float(p_value),
        "df": float(df),
    }


def _nearest_allowed(values: np.ndarray, allowed: np.ndarray) -> np.ndarray:
    distances = np.abs(values[:, None] - allowed[None, :])
    indices = distances.argmin(axis=1)
    return allowed[indices]


def _select_spline_knots(series: pd.Series) -> List[float]:
    values = series.dropna().to_numpy(dtype=float)
    quantile_knots = np.unique(np.quantile(values, [0.05, 0.35, 0.65, 0.95]))
    if quantile_knots.size >= 4:
        return quantile_knots[:4].tolist()
    fallback = np.array([0.0, 4.0, 9.0, 14.0], dtype=float)
    return np.unique(fallback).tolist()


def _build_spline_reference(df: pd.DataFrame) -> pd.DataFrame:
    reference: Dict[str, object] = {}
    columns = ["QS23", "QSSEXO", "QS25N", "HV025", "HV270", "VIOLENCIA_PAREJA", "ANIO", "HV001", "HV022", "PESO_FINAL"]
    for col in columns:
        series = df[col]
        if pd.api.types.is_numeric_dtype(series):
            if col == "PESO_FINAL":
                reference[col] = float(series.median())
            else:
                reference[col] = float(series.dropna().mode().iloc[0]) if series.dropna().nunique() < 10 else float(series.median())
        else:
            reference[col] = series.dropna().mode().iloc[0]
    return pd.DataFrame([reference])


def _resolve_r_home(configured_r_home: Optional[str]) -> str:
    candidates: List[Path] = []
    if configured_r_home:
        candidates.append(Path(configured_r_home))
    env_r_home = os.environ.get("R_HOME")
    if env_r_home:
        candidates.append(Path(env_r_home))
    candidates.extend(sorted(Path("C:/Program Files/R").glob("R-*"), reverse=True))

    for candidate in candidates:
        if candidate.exists() and (candidate / "bin").exists():
            return str(candidate)

    raise RuntimeError(
        "No se pudo resolver R_HOME. Instale R o declare analysis.r_home en la configuracion antes de ejecutar run_r_bridge."
    )


def _maybe_str(value: object) -> Optional[str]:
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def _extract_r_scalar(r_obj) -> float:
    try:
        values = np.asarray(r_obj, dtype=float)
    except (TypeError, ValueError):
        return np.nan

    if values.size == 0:
        return np.nan

    value = values.reshape(-1)[0]
    if pd.isna(value):
        return np.nan
    return float(value)


def _extract_r_text(r_obj) -> str:
    try:
        values = list(r_obj)
    except TypeError:
        return ""

    if not values:
        return ""
    return str(values[0])


_R_HELPERS = r"""
.build_design <- function(df, subset_expr="") {
  options(survey.lonely.psu="adjust")
  design <- survey::svydesign(id=~HV001, strata=~HV022, weights=~PESO_FINAL, data=df, nest=TRUE)
  if (!is.null(subset_expr) && nzchar(subset_expr)) {
    design <- subset(design, eval(parse(text=subset_expr)))
  }
  design
}

.safe_deff <- function(stat_obj) {
  tryCatch(as.numeric(survey::deff(stat_obj)), error=function(e) NA_real_)
}

fit_svyglm_model <- function(df, formula_text, subset_expr="") {
  design <- .build_design(df, subset_expr)
  fit <- suppressWarnings(survey::svyglm(as.formula(formula_text), design=design, family=quasipoisson(link="log")))
  model_frame <- model.frame(fit)
  df_resid <- suppressWarnings(as.numeric(df.residual(fit)))
  scale_dev <- tryCatch(
    if (is.finite(df_resid) && df_resid > 0) as.numeric(stats::deviance(fit) / df_resid) else NA_real_,
    error=function(e) NA_real_
  )
  n_events <- tryCatch(sum(as.numeric(model_frame[[1]]) == 1, na.rm=TRUE), error=function(e) NA_real_)
  list(
    coef = coef(fit),
    vcov = vcov(fit),
    n_obs = nrow(model_frame),
    n_events = n_events,
    df_resid = df_resid,
    scale_deviance = scale_dev,
    family = fit$family$family,
    link = fit$family$link
  )
}

fit_logistic_model <- function(df, formula_text) {
  design <- .build_design(df)
  fit <- suppressWarnings(survey::svyglm(as.formula(formula_text), design=design, family=quasibinomial(link="logit")))
  model_frame <- model.frame(fit)
  df_resid <- suppressWarnings(as.numeric(df.residual(fit)))
  scale_dev <- tryCatch(
    if (is.finite(df_resid) && df_resid > 0) as.numeric(stats::deviance(fit) / df_resid) else NA_real_,
    error=function(e) NA_real_
  )

  # Goodness-of-fit compatible con diseno muestral complejo: prueba F-adjusted
  # mean residual de Archer-Lemeshow (Archer KJ, Lemeshow S, Hosmer DW. Goodness-of-fit
  # tests for logistic regression models when data are collected using a complex
  # sampling design. Stat Med 2007;26:1885-1901). Se implementa con el mismo
  # svydesign: residuos (y - p_hat) agrupados por decil de probabilidad predicha y
  # un test de Wald F basado en diseno (survey::regTermTest) sobre los indicadores
  # de grupo. NO es el proxy ponderado de deciles HL anterior.
  gof_stat <- NA_real_
  gof_df <- NA_real_
  gof_df_den <- NA_real_
  gof_p <- NA_real_
  gof_n_groups <- NA_real_
  gof_method <- "archer_lemeshow_f_adjusted_survey"

  vars_used <- all.vars(as.formula(formula_text))
  outcome_var <- vars_used[1]
  cc <- tryCatch(stats::complete.cases(df[, vars_used, drop=FALSE]), error=function(e) logical(0))
  pred <- suppressWarnings(as.numeric(fitted(fit)))

  if (length(cc) == nrow(df) && sum(cc) == length(pred) &&
      sum(is.finite(pred)) > 0 && length(unique(pred[is.finite(pred)])) > 1) {
    breaks <- unique(stats::quantile(pred, probs=seq(0, 1, 0.1), na.rm=TRUE, names=FALSE, type=8))
    if (length(breaks) >= 3) {
      grp_used <- cut(pred, breaks=breaks, include.lowest=TRUE, labels=FALSE)
      yv <- suppressWarnings(as.numeric(df[[outcome_var]][cc]))
      resid_full <- rep(NA_real_, nrow(df))
      grp_full <- rep(NA_integer_, nrow(df))
      resid_full[cc] <- yv - pred
      grp_full[cc] <- grp_used

      al_design <- update(design, .al_resid = resid_full, .al_grp = factor(grp_full))
      al_fit <- tryCatch(
        suppressWarnings(survey::svyglm(.al_resid ~ .al_grp, design = al_design)),
        error=function(e) NULL
      )
      if (!is.null(al_fit)) {
        rt <- tryCatch(
          survey::regTermTest(al_fit, ~.al_grp, df = survey::degf(al_design), method = "Wald"),
          error=function(e) NULL
        )
        if (!is.null(rt)) {
          gof_stat <- as.numeric(if (!is.null(rt$Ftest)) rt$Ftest else rt$chisq)
          gof_df <- as.numeric(rt$df)
          gof_df_den <- as.numeric(if (!is.null(rt$ddf)) rt$ddf else NA_real_)
          gof_p <- as.numeric(rt$p)
          gof_n_groups <- as.numeric(length(unique(grp_used[is.finite(grp_used)])))
        }
      }
    }
  }

  list(
    coef = coef(fit),
    vcov = vcov(fit),
    n_obs = nrow(model_frame),
    df_resid = df_resid,
    scale_deviance = scale_dev,
    family = fit$family$family,
    link = fit$family$link,
    gof_method = gof_method,
    gof_statistic = gof_stat,
    gof_df = gof_df,
    gof_df_den = gof_df_den,
    gof_n_groups = gof_n_groups,
    gof_p_value = gof_p
  )
}

summarize_table1_bundle <- function(df, continuous_vars, categorical_vars, group_var="") {
  design <- .build_design(df)
  rows <- list()

  for (variable in continuous_vars) {
    design_var <- update(design, .tmp_cont = as.numeric(df[[variable]]))
    if (is.null(group_var) || !nzchar(group_var)) {
      stat <- suppressWarnings(survey::svymean(~.tmp_cont, design_var, na.rm=TRUE, deff="replace"))
      rows[[length(rows) + 1]] <- data.frame(
        kind = "continuous",
        variable = variable,
        group = "overall",
        label = variable,
        estimate = unname(coef(stat)[1]),
        variance = unname(vcov(stat)[1, 1]),
        deff = .safe_deff(stat)[1],
        n_unweighted = sum(!is.na(df[[variable]]))
      )
    } else {
      design_grp <- update(design_var, .tmp_group = factor(df[[group_var]]))
      stat <- suppressWarnings(survey::svyby(~.tmp_cont, ~.tmp_group, design_grp, survey::svymean, na.rm=TRUE, vartype="se"))
      rows[[length(rows) + 1]] <- data.frame(
        kind = "continuous",
        variable = variable,
        group = as.character(stat$.tmp_group),
        label = variable,
        estimate = stat$.tmp_cont,
        variance = stat$se^2,
        deff = NA_real_,
        n_unweighted = NA_real_
      )
    }
  }

  for (variable in categorical_vars) {
    design_var <- update(design, .tmp_cat = factor(df[[variable]]))
    if (is.null(group_var) || !nzchar(group_var)) {
      stat <- suppressWarnings(survey::svymean(~.tmp_cat, design_var, na.rm=TRUE, deff="replace"))
      coef_vals <- unname(coef(stat))
      label_vals <- gsub("^\\.tmp_cat", "", names(coef(stat)))
      deff_vals <- .safe_deff(stat)
      if (length(deff_vals) != length(coef_vals)) {
        deff_vals <- rep(NA_real_, length(coef_vals))
      }
      rows[[length(rows) + 1]] <- data.frame(
        kind = "categorical",
        variable = variable,
        group = "overall",
        label = label_vals,
        estimate = coef_vals,
        variance = diag(vcov(stat)),
        deff = deff_vals,
        n_unweighted = sum(!is.na(df[[variable]]))
      )
    } else {
      design_grp <- update(design_var, .tmp_group = factor(df[[group_var]]))
      stat <- suppressWarnings(survey::svyby(~.tmp_cat, ~.tmp_group, design_grp, survey::svymean, na.rm=TRUE, vartype="se"))
      estimate_cols <- grep("^\\.tmp_cat", names(stat), value=TRUE)
      for (col_name in estimate_cols) {
        se_name <- paste0("se.", col_name)
        label_val <- gsub("^\\.tmp_cat", "", col_name)
        rows[[length(rows) + 1]] <- data.frame(
          kind = "categorical",
          variable = variable,
          group = as.character(stat$.tmp_group),
          label = label_val,
          estimate = stat[[col_name]],
          variance = stat[[se_name]]^2,
          deff = NA_real_,
          n_unweighted = NA_real_
        )
      }
    }
  }

  if (!length(rows)) {
    return(data.frame())
  }
  do.call(rbind, rows)
}

compute_deff_metrics <- function(df) {
  design <- .build_design(df)
  bp_stat <- suppressWarnings(survey::svymean(~factor(PRESION_ARTERIAL_ELEVADA), design, na.rm=TRUE, deff="replace"))
  bp_deff <- .safe_deff(bp_stat)
  bp_names <- names(coef(bp_stat))
  bp_index <- which(grepl("1$", bp_names))
  if (!length(bp_index)) {
    bp_index <- length(bp_deff)
  }
  phq_stat <- suppressWarnings(survey::svymean(~PHQ9_TOTAL, design, na.rm=TRUE, deff="replace"))
  data.frame(
    metric = c("bp_prevalence_deff", "phq9_mean_deff"),
    value = c(bp_deff[bp_index[1]], .safe_deff(phq_stat)[1])
  )
}

rao_scott_bundle <- function(df, variables) {
  design <- .build_design(df)
  rows <- list()
  for (variable in variables) {
    design_var <- update(
      design,
      .tmp_pred = factor(df[[variable]]),
      .tmp_outcome = factor(df[["PRESION_ARTERIAL_ELEVADA"]])
    )
    test <- tryCatch(
      suppressWarnings(survey::svychisq(~.tmp_pred + .tmp_outcome, design_var, statistic="F")),
      error=function(e) NULL
    )
    if (is.null(test)) {
      rows[[length(rows) + 1]] <- data.frame(
        variable = variable,
        statistic = NA_real_,
        df_num = NA_real_,
        df_den = NA_real_,
        p_value = NA_real_
      )
    } else {
      rows[[length(rows) + 1]] <- data.frame(
        variable = variable,
        statistic = unname(test$statistic),
        df_num = unname(test$parameter[1]),
        df_den = unname(test$parameter[2]),
        p_value = unname(test$p.value)
      )
    }
  }
  do.call(rbind, rows)
}

fit_bivariate_term_model <- function(df, variable) {
  design <- .build_design(df)
  design <- update(design, .tmp_pred = factor(df[[variable]]))
  fit <- tryCatch(
    suppressWarnings(survey::svyglm(PRESION_ARTERIAL_ELEVADA ~ .tmp_pred, design=design, family=quasibinomial(link="logit"))),
    error=function(e) NULL
  )
  if (is.null(fit)) {
    return(list(coef=numeric(), terms=character(), vcov=matrix(numeric(), nrow=0, ncol=0), n_obs=0))
  }
  model_frame <- model.frame(fit)
  coef_vals <- coef(fit)
  keep <- names(coef_vals) != "(Intercept)"
  term_names <- names(coef_vals)[keep]
  if (!length(term_names)) {
    return(list(coef=numeric(), terms=character(), vcov=matrix(numeric(), nrow=0, ncol=0), n_obs=nrow(model_frame)))
  }
  idx <- which(keep)
  list(
    coef = unname(coef_vals[idx]),
    terms = term_names,
    vcov = vcov(fit)[idx, idx, drop=FALSE],
    n_obs = nrow(model_frame)
  )
}

fit_bivariate_term_bundle <- function(df, variables) {
  design <- .build_design(df)
  rows <- list()
  for (variable in variables) {
    design_var <- update(design, .tmp_pred = factor(df[[variable]]))
    fit <- tryCatch(
      suppressWarnings(survey::svyglm(PRESION_ARTERIAL_ELEVADA ~ .tmp_pred, design=design_var, family=quasibinomial(link="logit"))),
      error=function(e) NULL
    )
    if (is.null(fit)) {
      rows[[length(rows) + 1]] <- list(variable=variable, terms=character(), coef=numeric(), vcov=list(), n_obs=0)
    } else {
      model_frame <- model.frame(fit)
      coef_vals <- coef(fit)
      keep <- names(coef_vals) != "(Intercept)"
      term_names <- names(coef_vals)[keep]
      if (!length(term_names)) {
        rows[[length(rows) + 1]] <- list(variable=variable, terms=character(), coef=numeric(), vcov=list(), n_obs=nrow(model_frame))
      } else {
        idx <- which(keep)
        vcov_mat <- vcov(fit)[idx, idx, drop=FALSE]
        rows[[length(rows) + 1]] <- list(
          variable=variable,
          terms=unname(term_names),
          coef=unname(as.numeric(coef_vals[idx])),
          vcov=split(as.numeric(vcov_mat), row(vcov_mat)),
          n_obs=nrow(model_frame)
        )
      }
    }
  }
  jsonlite::toJSON(rows, auto_unbox=TRUE, digits=NA, null="null")
}

fit_spline_model <- function(df, knots, adjuster_text, reference_df) {
  basis <- as.data.frame(Hmisc::rcspline.eval(df$PHQ9_TOTAL, knots=knots, inclx=TRUE))
  basis_names <- paste0("PHQ_SPLINE_", seq_len(ncol(basis)))
  names(basis) <- basis_names
  df2 <- cbind(df, basis)

  rhs <- paste(c(basis_names, adjuster_text), collapse=" + ")
  design <- .build_design(df2)
  fit <- suppressWarnings(survey::svyglm(as.formula(paste("PRESION_ARTERIAL_ELEVADA ~", rhs)), design=design, family=quasipoisson(link="log")))

  nonlinear_terms <- basis_names[-1]
  wald <- survey::regTermTest(fit, as.formula(paste("~", paste(nonlinear_terms, collapse=" + "))))
  wald_candidates <- c(unname(wald$chisq), unname(wald$Ftest), unname(wald$F))
  wald_candidates <- suppressWarnings(as.numeric(wald_candidates))
  wald_stat <- if (length(wald_candidates) && any(is.finite(wald_candidates))) {
    wald_candidates[which(is.finite(wald_candidates))[1]]
  } else {
    NA_real_
  }
  p_value <- suppressWarnings(as.numeric(unname(wald$p)))
  if (!length(p_value) || !is.finite(p_value[1])) {
    p_value <- NA_real_
  } else {
    p_value <- p_value[1]
  }

  grid <- data.frame(PHQ9_TOTAL=seq(min(df$PHQ9_TOTAL, na.rm=TRUE), max(df$PHQ9_TOTAL, na.rm=TRUE), by=1))
  grid_basis <- as.data.frame(Hmisc::rcspline.eval(grid$PHQ9_TOTAL, knots=knots, inclx=TRUE))
  names(grid_basis) <- basis_names
  reference_expanded <- reference_df[rep(1, nrow(grid)), , drop=FALSE]
  newdata <- cbind(reference_expanded, grid_basis)
  pred_link <- suppressWarnings(predict(fit, newdata=newdata, type="link", se=TRUE))
  pred_eta <- as.numeric(coef(pred_link))
  pred_eta_se <- as.numeric(survey::SE(pred_link))
  pred_fit <- exp(pred_eta)
  pred_se <- pred_fit * pred_eta_se
  lower_ci <- exp(pred_eta - 1.96 * pred_eta_se)
  upper_ci <- exp(pred_eta + 1.96 * pred_eta_se)

  list(
    grid = grid$PHQ9_TOTAL,
    pred = pred_fit,
    pred_se = pred_se,
    lower_ci = lower_ci,
    upper_ci = upper_ci,
    wald_stat = wald_stat,
    df_num = length(nonlinear_terms),
    p_value = p_value
  )
}
"""
