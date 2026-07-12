from __future__ import annotations

import importlib.util
import shutil
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parents[1]
OUTPUT_DIR = ROOT / "data" / "output"
SEND_DIR = OUTPUT_DIR / "Para Enviar"
REVIEW_SOURCE_DIR = OUTPUT_DIR / "Rodrigo Revisar RRRR"
CURRENT_SOURCE_DIR = SEND_DIR
PUBLIC_DIR = SEND_DIR / "Para Publicar"
MAIN_DIR = PUBLIC_DIR / "Principal"
SUPP_DIR = PUBLIC_DIR / "Suplementario"
TECH_DIR = PUBLIC_DIR / "Revision_Tecnica"
FALLBACK_PUBLIC_DIR = SEND_DIR / "Para Publicar Redisenado"

VARIABLE_LABELS = {
    "ALCOHOL_PROBLEMATICO": "Consumo problemático de alcohol",
    "CALIDAD_DIETA": "Calidad de la dieta",
    "DX_HTA_PREVIO": "Diagnóstico previo de hipertensión",
    "HV025": "Área de residencia",
    "HV270": "Índice de riqueza",
    "IMC": "Índice de masa corporal",
    "PHQ9_TOTAL": "PHQ-9 total",
    "QS109": "Diagnóstico de diabetes",
    "QS201": "Consumo de tabaco en los últimos 30 días",
    "QS23": "Edad (años)",
    "QS25N": "Nivel educativo",
    "QS907": "Perímetro abdominal",
    "QSSEXO": "Sexo",
    "SEVERIDAD_DEPRESIVA": "Severidad depresiva",
    "VIOLENCIA_PAREJA": "Violencia de pareja",
}

GROUP_LABELS = {
    "by_outcome": {"0": "Sin presión arterial elevada", "1": "Con presión arterial elevada"},
    "by_sex": {"1": "Hombre", "2": "Mujer"},
    "by_severity": {
        "Minima": "Mínima",
        "Leve": "Leve",
        "Moderada": "Moderada",
        "Mod_Severa": "Moderada-severa",
        "Severa": "Severa",
    },
}

VALUE_LABELS = {
    "ALCOHOL_PROBLEMATICO": {"0": "No", "1": "Sí"},
    "CALIDAD_DIETA": {"0": "No adecuada", "1": "Adecuada"},
    "DX_HTA_PREVIO": {"0": "No", "1": "Sí"},
    "QS109": {"0": "No", "1": "Sí"},
    "QS201": {"0": "No", "1": "Sí"},
    "QSSEXO": {"1": "Hombre", "2": "Mujer"},
}

MODEL_LABELS = {
    "model_1": "Modelo 1",
    "model_2": "Modelo 2 (principal)",
    "model_3": "Modelo 3 (exploratorio)",
    "submodel_adherence": "No adherencia terapéutica",
    "submodel_domain_bp": "Descontrol tensional en HTA previa",
    "interaction_sex": "Interacción PHQ-9 × sexo",
    "interaction_year": "Interacción PHQ-9 × año",
    "sensitivity_no_2020": "Sensibilidad excluyendo 2020",
    "sensitivity_second_bp_measure": "Sensibilidad con segunda toma de PA",
    "model_2_logistic_sensitivity": "Sensibilidad logística del Modelo 2",
}


def _run_build_submission_package() -> None:
    module_path = ROOT / "scripts" / "build_submission_package.py"
    spec = importlib.util.spec_from_file_location("build_submission_package", module_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"No se pudo cargar {module_path}.")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    module.main()


def _run_build_guide_display_audit() -> None:
    module_path = ROOT / "scripts" / "build_guide_display_audit.py"
    spec = importlib.util.spec_from_file_location("build_guide_display_audit", module_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"No se pudo cargar {module_path}.")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    module.build_audit(PUBLIC_DIR)


def _fmt_decimal(value: object, decimals: int = 2) -> str:
    if pd.isna(value):
        return ""
    return f"{float(value):.{decimals}f}".replace(".", ",")


def _fmt_p(value: object) -> str:
    if pd.isna(value):
        return ""
    value = float(value)
    if value < 0.001:
        return "<0,001"
    return _fmt_decimal(value, 3)


def _fmt_pct(value: object) -> str:
    if pd.isna(value):
        return ""
    return _fmt_decimal(float(value) * 100.0, 1)


def _fmt_ci(low: object, high: object, decimals: int = 2) -> str:
    if pd.isna(low) or pd.isna(high):
        return ""
    return f"{_fmt_decimal(low, decimals)} a {_fmt_decimal(high, decimals)}"


def _fmt_pct_ci(low: object, high: object) -> str:
    if pd.isna(low) or pd.isna(high):
        return ""
    return f"{_fmt_pct(low)} a {_fmt_pct(high)}"


def _format_group(summary_type: str, value: object) -> str:
    text = "" if pd.isna(value) else str(value)
    return GROUP_LABELS.get(summary_type, {}).get(text, text)


def _format_label(variable: str, value: object) -> str:
    text = "" if pd.isna(value) else str(value)
    return VALUE_LABELS.get(variable, {}).get(text, text)


def _format_table1_panel(df: pd.DataFrame, summary_type: str) -> pd.DataFrame:
    panel = df.loc[df["summary_type"] == summary_type].copy()
    if panel.empty:
        return panel

    panel["Variable"] = panel["variable"].map(lambda x: VARIABLE_LABELS.get(str(x), str(x)))
    panel["Grupo"] = panel["group"].map(lambda x: _format_group(summary_type, x))
    panel["Categoría"] = panel.apply(lambda row: _format_label(str(row["variable"]), row["label"]), axis=1)

    rows = []
    for _, row in panel.iterrows():
        categorical = row.get("kind") == "categorical"
        rows.append(
            {
                "Variable": row["Variable"],
                "Grupo": row["Grupo"],
                "Categoría": row["Categoría"],
                "Estimación": _fmt_pct(row["estimate"]) if categorical else _fmt_decimal(row["estimate"], 2),
                "EE": _fmt_pct(row["std_error"]) if categorical else _fmt_decimal(row["std_error"], 2),
                "IC 95%": _fmt_pct_ci(row["ci_low"], row["ci_high"]) if categorical else _fmt_ci(row["ci_low"], row["ci_high"], 2),
                "p": _fmt_p(row["p_value"]),
                "DEFF promedio": _fmt_decimal(row["mean_deff"], 2),
                "n no ponderado": "" if pd.isna(row.get("mean_n_unweighted")) else str(int(round(float(row["mean_n_unweighted"])))),
            }
        )
    return pd.DataFrame(rows)


def _format_table2(df: pd.DataFrame) -> pd.DataFrame:
    frame = df.copy()
    frame["Variable"] = frame["variable"].map(lambda x: VARIABLE_LABELS.get(str(x), str(x)))
    return pd.DataFrame(
        {
            "Variable": frame["Variable"],
            "Método de pooling": frame["pool_method"],
            "Términos": frame["terms"],
            "gl": frame["df_num"].map(lambda x: "" if pd.isna(x) else str(int(x))),
            "Estadístico pooled": frame["statistic"].map(lambda x: _fmt_decimal(x, 2)),
            "p pooled": frame["p_value"].map(_fmt_p),
            "Imputaciones": frame["imputations_used"].map(lambda x: "" if pd.isna(x) else str(int(x))),
            "n medio": frame["mean_n_obs"].map(lambda x: "" if pd.isna(x) else str(int(round(float(x))))),
            "Rao-Scott medio": frame["mean_rao_scott_statistic"].map(lambda x: _fmt_decimal(x, 2)),
            "p Rao-Scott mediana": frame["median_rao_scott_p_value"].map(_fmt_p),
        }
    )


def _format_table2_main(df: pd.DataFrame) -> pd.DataFrame:
    frame = df.copy()
    frame["Variable"] = frame["variable"].map(lambda x: VARIABLE_LABELS.get(str(x), str(x)))
    comparison_map = {
        "ALCOHOL_PROBLEMATICO": "Comparacion global de categorias observadas",
        "CALIDAD_DIETA": "Adecuada vs no adecuada",
        "HV025": "Urbano vs rural",
        "HV270": "Comparacion global por quintiles de riqueza",
        "QS109": "Si vs no",
        "QS201": "Comparacion global incluyendo faltantes observados",
        "QS25N": "Comparacion global por nivel educativo",
        "QSSEXO": "Mujer vs hombre",
        "SEVERIDAD_DEPRESIVA": "Comparacion global por severidad depresiva",
        "VIOLENCIA_PAREJA": "Comparacion global de categorias",
    }
    return pd.DataFrame(
        {
            "Variable": frame["Variable"],
            "Comparacion": frame["variable"].map(lambda x: comparison_map.get(str(x), "Comparacion global")),
            "Estadistico pooled": frame["statistic"].map(lambda x: _fmt_decimal(x, 2)),
            "p pooled": frame["p_value"].map(_fmt_p),
            "Imputaciones": frame["imputations_used"].map(lambda x: "" if pd.isna(x) else str(int(x))),
            "n medio": frame["mean_n_obs"].map(lambda x: "" if pd.isna(x) else str(int(round(float(x))))),
        }
    )


def _format_term(term: str) -> str:
    if term == "PHQ9_TOTAL":
        return "PHQ-9 total"
    if term == "QS23":
        return "Edad (años)"
    if term == "factor(QSSEXO)2":
        return "Sexo: mujer"
    if term.startswith("PHQ9_TOTAL:factor(QSSEXO)"):
        return "Interacción PHQ-9 × sexo"
    if term.startswith("PHQ9_TOTAL:factor(ANIO)"):
        return "Interacción PHQ-9 × año"
    if term.startswith("factor(QS25N)"):
        return f"Nivel educativo ({term.split(')')[-1] or term[-1]})"
    if term.startswith("factor(HV025)"):
        return f"Área de residencia ({term.split(')')[-1] or term[-1]})"
    if term.startswith("factor(HV270)"):
        return f"Índice de riqueza ({term.split(')')[-1] or term[-1]})"
    if term.startswith("factor(VIOLENCIA_PAREJA)"):
        return f"Violencia de pareja ({term.split(')')[-1] or term[-1]})"
    if term.startswith("factor(QS201)"):
        return f"Tabaquismo ({term.split(')')[-1] or term[-1]})"
    if term.startswith("factor(ALCOHOL_PROBLEMATICO)"):
        return f"Alcohol problemático ({term.split(')')[-1] or term[-1]})"
    if term.startswith("factor(CALIDAD_DIETA)"):
        return f"Calidad de la dieta ({term.split(')')[-1] or term[-1]})"
    if term.startswith("factor(QS109)"):
        return f"Diagnóstico de diabetes ({term.split(')')[-1] or term[-1]})"
    return term


def _format_model_table(df: pd.DataFrame) -> pd.DataFrame:
    frame = df.loc[df["term"] != "(Intercept)"].copy()
    return pd.DataFrame(
        {
            "Modelo": frame["model"].map(lambda x: MODEL_LABELS.get(str(x), str(x))),
            "Término": frame["term"].map(lambda x: _format_term(str(x))),
            "RP": frame["exp_estimate"].map(lambda x: _fmt_decimal(x, 2)),
            "IC 95%": frame.apply(lambda row: _fmt_ci(row["exp_ci_low"], row["exp_ci_high"], 2), axis=1),
            "p": frame["p_value"].map(_fmt_p),
            "n medio": frame["mean_n_obs"].map(lambda x: "" if pd.isna(x) else str(int(round(float(x))))),
        }
    )


def _format_model_diagnostics(df: pd.DataFrame) -> pd.DataFrame:
    frame = df.copy()
    return pd.DataFrame(
        {
            "Modelo": frame["model"].map(lambda x: MODEL_LABELS.get(str(x), str(x))),
            "Familia": frame["family"],
            "Enlace": frame["link"],
            "Imputaciones": frame["imputations"].map(lambda x: "" if pd.isna(x) else str(int(x))),
            "n medio": frame["mean_n_obs"].map(lambda x: "" if pd.isna(x) else str(int(round(float(x))))),
            "gl residuales medios": frame["mean_df_resid"].map(lambda x: _fmt_decimal(x, 2)),
            "Devianza escalada media": frame["mean_scale_deviance"].map(lambda x: _fmt_decimal(x, 2)),
            "Devianza escalada mediana": frame["median_scale_deviance"].map(lambda x: _fmt_decimal(x, 2)),
        }
    )


def _format_deff(df: pd.DataFrame) -> pd.DataFrame:
    frame = df.copy()
    frame["Indicador"] = frame["metric"].replace(
        {
            "table1_overall_mean_deff": "DEFF promedio global de Tabla 1",
            "bp_prevalence_deff": "DEFF prevalencia de presión arterial elevada",
            "phq9_mean_deff": "DEFF media de PHQ-9",
        }
    )
    return pd.DataFrame(
        {
            "Indicador": frame["Indicador"],
            "Valor": frame["value"].map(lambda x: _fmt_decimal(x, 2)),
            "Imputaciones": frame["imputations"].map(lambda x: "" if pd.isna(x) else str(int(x))),
        }
    )


def _build_robustness_table(
    main_models: pd.DataFrame,
    sensitivity_models: pd.DataFrame,
    spline_meta: pd.DataFrame,
) -> pd.DataFrame:
    model_2 = main_models.loc[
        (main_models["model"] == "model_2") & (main_models["term"] == "PHQ9_TOTAL")
    ].iloc[0]
    second_bp = sensitivity_models.loc[
        (sensitivity_models["model"] == "sensitivity_second_bp_measure")
        & (sensitivity_models["term"] == "PHQ9_TOTAL")
    ].iloc[0]
    spline = spline_meta.iloc[0]

    rows = [
        {
            "Escenario": "Definicion principal del desenlace",
            "Modelo base": "Modelo 2 (principal)",
            "Parametro": "RP por 1 punto adicional en PHQ-9",
            "Estimacion": _fmt_decimal(model_2["exp_estimate"], 2),
            "IC 95%": _fmt_ci(model_2["exp_ci_low"], model_2["exp_ci_high"], 2),
            "p": _fmt_p(model_2["p_value"]),
            "n medio": str(int(round(float(model_2["mean_n_obs"])))),
            "Imputaciones": "20",
            "Lectura": "La asociacion principal se mantiene.",
            "Nota": "Desenlace principal de presion arterial elevada segun la definicion preespecificada.",
        },
        {
            "Escenario": "Sensibilidad con segunda toma",
            "Modelo base": "Sensibilidad con segunda toma de PA",
            "Parametro": "RP por 1 punto adicional en PHQ-9",
            "Estimacion": _fmt_decimal(second_bp["exp_estimate"], 2),
            "IC 95%": _fmt_ci(second_bp["exp_ci_low"], second_bp["exp_ci_high"], 2),
            "p": _fmt_p(second_bp["p_value"]),
            "n medio": str(int(round(float(second_bp["mean_n_obs"])))),
            "Imputaciones": "20",
            "Lectura": "La direccion y la magnitud del efecto son consistentes con el analisis principal.",
            "Nota": "Sensibilidad usando solo la segunda medicion de presion arterial.",
        },
        {
            "Escenario": "Spline corregida",
            "Modelo base": "Curva spline cubica restringida",
            "Parametro": "Resumen de no linealidad",
            "Estimacion": "",
            "IC 95%": "",
            "p": "",
            "n medio": "",
            "Imputaciones": str(int(spline["n_imputations_with_test"])),
            "Lectura": "No hubo evidencia consistente de no linealidad entre imputaciones.",
            "Nota": (
                "Resumen entre imputaciones: p media de no linealidad = "
                f"{_fmt_p(spline['mean_p_value'])}; "
                f"mediana = {_fmt_p(spline['median_p_value'])}; "
                f"rango = {_fmt_p(spline['min_p_value'])} a {_fmt_p(spline['max_p_value'])}."
            ),
        },
    ]
    return pd.DataFrame(rows)


def _autofit_workbook(path: Path) -> None:
    wb = load_workbook(path)
    for ws in wb.worksheets:
        for column_cells in ws.columns:
            values = [str(cell.value) for cell in column_cells if cell.value is not None]
            if not values:
                continue
            width = min(max(len(value) for value in values) + 2, 48)
            ws.column_dimensions[column_cells[0].column_letter].width = width
    wb.save(path)


def _write_workbook(path: Path, sheets: dict[str, pd.DataFrame]) -> None:
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for sheet_name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    _autofit_workbook(path)


def _switch_public_dir(path: Path) -> None:
    global PUBLIC_DIR, MAIN_DIR, SUPP_DIR, TECH_DIR
    PUBLIC_DIR = path
    MAIN_DIR = PUBLIC_DIR / "Principal"
    SUPP_DIR = PUBLIC_DIR / "Suplementario"
    TECH_DIR = PUBLIC_DIR / "Revision_Tecnica"


def _ensure_dirs() -> None:
    target_dir = PUBLIC_DIR
    if target_dir.exists():
        try:
            shutil.rmtree(target_dir)
        except PermissionError:
            target_dir = FALLBACK_PUBLIC_DIR
            _switch_public_dir(target_dir)
            if target_dir.exists():
                shutil.rmtree(target_dir)
    MAIN_DIR.mkdir(parents=True, exist_ok=True)
    SUPP_DIR.mkdir(parents=True, exist_ok=True)
    TECH_DIR.mkdir(parents=True, exist_ok=True)


def _build_technical_review_materials() -> None:
    shutil.copy2(ROOT / "src" / "endes_pipeline" / "pipeline.py", TECH_DIR / "01_script_principal_pipeline.py")
    shutil.copy2(ROOT / "src" / "endes_pipeline" / "analysis.py", TECH_DIR / "01a_modulo_analisis_mice.py")
    shutil.copy2(ROOT / "logs" / "pipeline.log", TECH_DIR / "02_log_recodificacion_imputacion.log")
    resume_log = ROOT / "logs" / "run_analysis_resume.stderr.log"
    if resume_log.exists() and resume_log.stat().st_size > 0:
        shutil.copy2(resume_log, TECH_DIR / "02a_log_reanudacion_puente_r.log")

    manifest_path = OUTPUT_DIR / "analysis" / "mice_manifest.json"
    note_lines = [
        "Material tecnico para revision del pipeline",
        "",
        "Incluye:",
        "- script principal del pipeline (pipeline.py)",
        "- modulo de analisis/imputacion MICE (analysis.py), porque ahi vive la logica de MICE",
        "- log de recodificacion/imputacion",
    ]

    if resume_log.exists() and resume_log.stat().st_size > 0:
        note_lines.append("- log de reanudacion del puente R para completar tablas, modelos y figuras")

    if manifest_path.exists():
        shutil.copy2(manifest_path, TECH_DIR / "03_mice_manifest.json")
        note_lines.extend(
            [
                "- manifiesto de variables de MICE (targets, predictores, semilla e imputaciones generadas)",
                "",
                "Nota:",
                "- no existe un archivo persistido aparte con la matriz fila-por-fila usada por MICE;",
                "- el repo guarda el manifiesto de variables en mice_manifest.json, que es el artefacto disponible para revision.",
            ]
        )
    else:
        note_lines.extend(
            [
                "",
                "Nota:",
                "- no se encontro un artefacto persistido con la matriz o el manifiesto de variables de MICE.",
            ]
        )

    (TECH_DIR / "00_NOTA_REVISION_TECNICA.txt").write_text("\n".join(note_lines) + "\n", encoding="utf-8-sig")


def _build_review_manifest() -> pd.DataFrame:
    rows = [
        {
            "source_file": "00_LEEME.txt",
            "decision": "no_incluir",
            "destination": "",
            "reason_rpmesp": "Archivo interno; no corresponde a tabla, figura ni material suplementario para publicación.",
        },
        {
            "source_file": "01_base_analitica_final.csv",
            "decision": "no_incluir",
            "destination": "",
            "reason_rpmesp": "La base analítica puede considerarse base de datos suplementaria externa, pero no es necesaria en el paquete editorial principal.",
        },
        {
            "source_file": "01_base_analitica_final.parquet",
            "decision": "no_incluir",
            "destination": "",
            "reason_rpmesp": "Formato técnico interno; no es un formato de tabla o figura para envío editorial.",
        },
        {
            "source_file": "02_tabla_strobe_flujo.csv",
            "decision": "suplementario",
            "destination": "Suplementario/Tabla_S1_flujo_STROBE.xlsx",
            "reason_rpmesp": "Se prioriza Figura 1 en el cuerpo principal para respetar el máximo de seis tablas/figuras.",
        },
        {
            "source_file": "03_tabla_1_caracteristicas_basales.csv",
            "decision": "principal",
            "destination": "Principal/Tabla_1.xlsx",
            "reason_rpmesp": "Corresponde a una tabla principal del artículo original.",
        },
        {
            "source_file": "03a_tabla_1_deff_resumen.csv",
            "decision": "principal_apoyo",
            "destination": "Principal/Tabla_1.xlsx y Suplementario/Tabla_S2_soporte_tablas.xlsx",
            "reason_rpmesp": "El DEFF debe reportarse como nota de Tabla 1 y conservarse en soporte editable.",
        },
        {
            "source_file": "04_tabla_2_bivariado.csv",
            "decision": "principal",
            "destination": "Principal/Tabla_2.xlsx",
            "reason_rpmesp": "Corresponde a una tabla principal del artículo original.",
        },
        {
            "source_file": "05_tabla_3_modelos_principales.csv",
            "decision": "principal",
            "destination": "Principal/Tabla_3.xlsx",
            "reason_rpmesp": "Corresponde a una tabla principal del artículo original.",
        },
        {
            "source_file": "05a_diagnosticos_modelos.csv",
            "decision": "suplementario",
            "destination": "Suplementario/Tabla_S2_soporte_tablas.xlsx",
            "reason_rpmesp": "Diagnóstico de modelos complementa los hallazgos y puede ir como material suplementario.",
        },
        {
            "source_file": "06_tabla_4_cascada_cuidado.csv",
            "decision": "principal",
            "destination": "Principal/Tabla_4.xlsx",
            "reason_rpmesp": "Corresponde a una tabla principal del artículo original.",
        },
        {
            "source_file": "07_figura_1_strobe.png",
            "decision": "principal_reemplazado",
            "destination": "Principal/Figura_1_STROBE.(png/pdf/svg)",
            "reason_rpmesp": "La versión en Rodrigo Revisar RRRR fue sustituida por la versión corregida en Para Enviar.",
        },
        {
            "source_file": "07_figura_1_strobe.pdf",
            "decision": "principal_reemplazado",
            "destination": "Principal/Figura_1_STROBE.(png/pdf/svg)",
            "reason_rpmesp": "La versión en Rodrigo Revisar RRRR fue sustituida por la versión corregida en Para Enviar.",
        },
        {
            "source_file": "07_figura_1_strobe_datos.csv",
            "decision": "suplementario",
            "destination": "Suplementario/Tabla_S4_datos_figuras.xlsx",
            "reason_rpmesp": "Los datos fuente de la figura se preservan como material suplementario editable.",
        },
        {
            "source_file": "08_figura_2_spline_phq9_pa.png",
            "decision": "suplementario_reemplazado",
            "destination": "Suplementario/Figura_S1_Spline.(png/pdf/svg)",
            "reason_rpmesp": "Se conserva como figura suplementaria porque funciona mejor como chequeo metodológico que como hallazgo principal.",
        },
        {
            "source_file": "08_figura_2_spline_phq9_pa.pdf",
            "decision": "suplementario_reemplazado",
            "destination": "Suplementario/Figura_S1_Spline.(png/pdf/svg)",
            "reason_rpmesp": "Se conserva como figura suplementaria porque funciona mejor como chequeo metodológico que como hallazgo principal.",
        },
        {
            "source_file": "08_figura_2_spline_phq9_pa_datos.csv",
            "decision": "suplementario",
            "destination": "Suplementario/Tabla_S4_datos_figuras.xlsx",
            "reason_rpmesp": "Los datos fuente de la figura se preservan como material suplementario editable.",
        },
        {
            "source_file": "08a_spline_nolinealidad_resumen.csv",
            "decision": "suplementario",
            "destination": "Suplementario/Tabla_S4_datos_figuras.xlsx",
            "reason_rpmesp": "Resumen metodológico de la curva spline; no es necesario en el cuerpo principal.",
        },
        {
            "source_file": "09_analisis_sensibilidad_interacciones.csv",
            "decision": "suplementario",
            "destination": "Suplementario/Tabla_S3_sensibilidad_interacciones.xlsx",
            "reason_rpmesp": "Los análisis de sensibilidad pueden publicarse como material suplementario.",
        },
        {
            "source_file": "09a_sensibilidad_logistica.csv",
            "decision": "suplementario",
            "destination": "Suplementario/Tabla_S3_sensibilidad_interacciones.xlsx",
            "reason_rpmesp": "Análisis complementario de sensibilidad; corresponde a material suplementario.",
        },
        {
            "source_file": "09b_items_omitidos.csv",
            "decision": "no_incluir",
            "destination": "",
            "reason_rpmesp": "Nota operativa interna, no corresponde a tabla o figura publicable.",
        },
        {
            "source_file": "GENERADO_desde_tabla3_sensibilidad_y_spline",
            "decision": "principal",
            "destination": "Principal/Tabla_5.xlsx",
            "reason_rpmesp": "Sintetiza la estabilidad del hallazgo principal y reemplaza mejor a una segunda figura principal que no agrega un hallazgo nuevo.",
        },
    ]
    return pd.DataFrame(rows)


def _write_review_and_rules() -> None:
    review = _build_review_manifest()
    review.to_csv(PUBLIC_DIR / "00_revision_entregables_rrrr.csv", index=False, encoding="utf-8-sig")

    rules = """Criterios RPMESP aplicados para este paquete

Fuente exclusiva:
- Instrucciones para autores de la Revista Peruana de Medicina Experimental y Salud Pública.

Reglas aplicadas:
- Artículo original: máximo 6 tablas o figuras en el cuerpo principal.
- Las tablas deben estar en formato editable (Word o Excel).
- Los gráficos y diagramas deben enviarse en formato editable.
- Las figuras se envían aparte y también deben poder incorporarse al archivo principal del manuscrito.
- El material suplementario debe enviarse como archivo separado.
- p: 3 decimales.
- Estimadores y medidas de asociación: 2 decimales.
- Porcentajes: 1 decimal.
- En español se usa coma decimal.

Decisión editorial para este paquete:
- Principal: 5 tablas + 1 figura = 6 elementos.
- Suplementario: flujo STROBE tabular, soportes de Tabla 1 y Tabla 2, diagnósticos, sensibilidad/interacciones, figura spline y datos fuente de figuras.
"""
    (PUBLIC_DIR / "00_criterios_rpmesp.txt").write_text(rules, encoding="utf-8-sig")


def _build_main_tables() -> None:
    table1 = pd.read_csv(CURRENT_SOURCE_DIR / "03_tabla_1_caracteristicas_basales.csv")
    deff = pd.read_csv(CURRENT_SOURCE_DIR / "03a_tabla_1_deff_resumen.csv")
    table2 = pd.read_csv(CURRENT_SOURCE_DIR / "04_tabla_2_bivariado.csv")
    table3 = pd.read_csv(CURRENT_SOURCE_DIR / "05_tabla_3_modelos_principales.csv")
    table4 = pd.read_csv(CURRENT_SOURCE_DIR / "06_tabla_4_cascada_cuidado.csv")
    sensitivity = pd.read_csv(CURRENT_SOURCE_DIR / "09_analisis_sensibilidad_interacciones.csv")
    spline_meta = pd.read_csv(CURRENT_SOURCE_DIR / "08a_spline_nolinealidad_resumen.csv")

    _write_workbook(
        MAIN_DIR / "Tabla_1.xlsx",
        {
            "Resumen_general": _format_table1_panel(table1, "overall"),
            "Por_desenlace": _format_table1_panel(table1, "by_outcome"),
            "DEFF": _format_deff(deff),
        },
    )
    _write_workbook(MAIN_DIR / "Tabla_2.xlsx", {"Tabla_2": _format_table2_main(table2)})
    _write_workbook(MAIN_DIR / "Tabla_3.xlsx", {"Tabla_3": _format_model_table(table3)})
    _write_workbook(MAIN_DIR / "Tabla_4.xlsx", {"Tabla_4": _format_model_table(table4)})
    _write_workbook(
        MAIN_DIR / "Tabla_5.xlsx",
        {"Tabla_5": _build_robustness_table(table3, sensitivity, spline_meta)},
    )


def _build_supplementary_files() -> None:
    table3 = pd.read_csv(CURRENT_SOURCE_DIR / "05_tabla_3_modelos_principales.csv")
    table1 = pd.read_csv(CURRENT_SOURCE_DIR / "03_tabla_1_caracteristicas_basales.csv")
    table2 = pd.read_csv(CURRENT_SOURCE_DIR / "04_tabla_2_bivariado.csv")
    strobe = pd.read_csv(CURRENT_SOURCE_DIR / "02_tabla_strobe_flujo.csv")
    deff = pd.read_csv(CURRENT_SOURCE_DIR / "03a_tabla_1_deff_resumen.csv")
    diagnostics = pd.read_csv(CURRENT_SOURCE_DIR / "05a_diagnosticos_modelos.csv")
    sensitivity = pd.read_csv(CURRENT_SOURCE_DIR / "09_analisis_sensibilidad_interacciones.csv")
    logistic = pd.read_csv(CURRENT_SOURCE_DIR / "09a_sensibilidad_logistica.csv")
    fig1_data = pd.read_csv(CURRENT_SOURCE_DIR / "07_figura_1_strobe_datos.csv")
    fig2_data = pd.read_csv(CURRENT_SOURCE_DIR / "08_figura_2_spline_phq9_pa_datos.csv")
    fig2_meta = pd.read_csv(CURRENT_SOURCE_DIR / "08a_spline_nolinealidad_resumen.csv")

    strobe_fmt = strobe.copy()
    for col in ["pct_excluded_from_before", "pct_remaining_from_start"]:
        strobe_fmt[col] = strobe_fmt[col].map(lambda x: _fmt_decimal(x, 1))

    _write_workbook(
        SUPP_DIR / "Tabla_S1_flujo_STROBE.xlsx",
        {"STROBE": strobe_fmt},
    )
    _write_workbook(
        SUPP_DIR / "Tabla_S2_soporte_tablas.xlsx",
        {
            "Tabla_1_por_sexo": _format_table1_panel(table1, "by_sex"),
            "Tabla_1_por_severidad": _format_table1_panel(table1, "by_severity"),
            "Tabla_2_detalle": _format_table2(table2),
            "DEFF": _format_deff(deff),
            "Diagnosticos_modelos": _format_model_diagnostics(diagnostics),
        },
    )
    _write_workbook(
        SUPP_DIR / "Tabla_S3_sensibilidad_interacciones.xlsx",
        {
            "Sensibilidad_interacciones": _format_model_table(sensitivity),
            "Sensibilidad_logistica": _format_model_table(logistic),
        },
    )

    fig1_fmt = fig1_data.copy()
    fig1_fmt["pct_excluded_from_before"] = fig1_fmt["pct_excluded_from_before"].map(lambda x: _fmt_decimal(x, 1))
    fig1_fmt["pct_remaining_from_start"] = fig1_fmt["pct_remaining_from_start"].map(lambda x: _fmt_decimal(x, 1))

    fig2_fmt = fig2_data.copy()
    fig2_fmt["estimate_pct"] = fig2_fmt["estimate"].map(_fmt_pct)
    fig2_fmt["ci_low_pct"] = fig2_fmt["ci_low"].map(_fmt_pct)
    fig2_fmt["ci_high_pct"] = fig2_fmt["ci_high"].map(_fmt_pct)

    fig2_meta_fmt = fig2_meta.copy()
    for col in ["mean_wald_statistic", "rubin_style_mean_statistic"]:
        fig2_meta_fmt[col] = fig2_meta_fmt[col].map(lambda x: _fmt_decimal(x, 2))
    for col in ["median_p_value", "mean_p_value", "min_p_value", "max_p_value"]:
        fig2_meta_fmt[col] = fig2_meta_fmt[col].map(_fmt_p)

    _write_workbook(
        SUPP_DIR / "Tabla_S4_datos_figuras.xlsx",
        {
            "Figura_1_STROBE": fig1_fmt,
            "Figura_2_spline": fig2_fmt,
            "Figura_2_resumen": fig2_meta_fmt,
        },
    )


def _copy_main_figures() -> None:
    for src_name, dest_stem in [
        ("07_figura_1_strobe", "Figura_1_STROBE"),
    ]:
        for ext in [".png", ".pdf", ".svg"]:
            shutil.copy2(SEND_DIR / f"{src_name}{ext}", MAIN_DIR / f"{dest_stem}{ext}")


def _copy_supplementary_figures() -> None:
    for src_name, dest_stem in [
        ("08_figura_2_spline_phq9_pa", "Figura_S1_Spline"),
    ]:
        for ext in [".png", ".pdf", ".svg"]:
            shutil.copy2(SEND_DIR / f"{src_name}{ext}", SUPP_DIR / f"{dest_stem}{ext}")


def _write_figure_legends() -> None:
    legends = """Leyendas de figuras para RPMESP

Figura 1. Diagrama de flujo STROBE para la selección muestral del análisis multianual ENDES 2019-2024.
STROBE: Strengthening the Reporting of Observational Studies in Epidemiology; ENDES: Encuesta Demográfica y de Salud Familiar; MEF: mujeres en edad fértil; PHQ-9: Patient Health Questionnaire-9; PA: presión arterial.
"""
    (MAIN_DIR / "Leyendas_figuras_RPMESP.txt").write_text(legends, encoding="utf-8-sig")
    supp_legends = """Leyendas de figuras suplementarias

Figura S1. Curva spline cúbica restringida de la probabilidad predicha de presión arterial elevada según el puntaje PHQ-9 en 20 imputaciones.
La banda sombreada representa el IC 95%. La figura se conserva como verificación de forma funcional, no como hallazgo principal. PHQ-9: Patient Health Questionnaire-9; IC 95%: intervalo de confianza del 95%.
"""
    (SUPP_DIR / "Leyendas_figuras_suplementarias.txt").write_text(supp_legends, encoding="utf-8-sig")


def main() -> None:
    _run_build_submission_package()
    _ensure_dirs()
    _write_review_and_rules()
    _build_main_tables()
    _build_supplementary_files()
    _build_technical_review_materials()
    _copy_main_figures()
    _copy_supplementary_figures()
    _write_figure_legends()
    _run_build_guide_display_audit()
    print(f"Paquete RPMESP listo en: {PUBLIC_DIR}")


if __name__ == "__main__":
    main()
