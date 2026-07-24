"""Reajusta solo los seis modelos de interacción y conserva insumos D1 a 15 dígitos.

El rerun completo del 23-07-2026 serializó Qhat/Uhat con la precisión por
defecto de pandas (10 dígitos). Esa pérdida no afecta el panel calculado en
memoria, pero impide una reproducción a tolerancia estricta desde el JSON.
Este script evita repetir tablas y modelos no relacionados.
"""

from __future__ import annotations

import json
import logging
import sys
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parents[2]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from endes_pipeline import analysis as a  # noqa: E402


def main() -> None:
    config_path = ROOT / "config" / "pipeline_config_2025.json"
    config = json.loads(config_path.read_text(encoding="utf-8-sig"))
    settings = a.build_analysis_settings(
        Path(config["paths"]["output_dir"]),
        dict(config.get("analysis", {})),
    )
    paths = sorted(settings.imputed_dir.glob("imputation_*.parquet"))
    if len(paths) != 20:
        raise RuntimeError(f"Se esperaban 20 imputaciones; encontradas: {len(paths)}")

    ro, pandas2ri, localconverter = a._load_r_bridge(settings)
    ro.r(a._R_HELPERS)
    fits = {model: [] for model in a.EFFECT_MOD_MODELS}
    sex_slopes = []

    for imputation_id, path in enumerate(paths, start=1):
        logging.info("Refit interacción %s/20: %s", imputation_id, path.name)
        frame = pd.read_parquet(path)
        r_frame = a._to_r_dataframe(frame, pandas2ri, localconverter)
        for model_name in a.EFFECT_MOD_MODELS:
            fit = ro.globalenv["fit_svyglm_model"](
                r_frame,
                a.MAIN_MODELS[model_name],
                a.MODEL_SUBSETS[model_name],
            )
            fits[model_name].append(
                a._interaction_term_fit(fit, imputation_id, model_name)
            )
            if model_name == "interaction_sex":
                sex_slopes.extend(a._sex_specific_slope_fit(fit, imputation_id))

    panel = a._build_effect_modification_panel(fits)
    panel.to_csv(
        settings.models_dir / "effect_modification_panel.csv",
        index=False,
        encoding="utf-8-sig",
    )
    a._serialize_effect_modification_d1_inputs(fits).to_json(
        settings.models_dir / "effect_modification_d1_inputs.json",
        orient="records",
        indent=2,
        force_ascii=False,
        double_precision=15,
    )
    a._pool_sex_specific_slopes(sex_slopes).to_csv(
        settings.models_dir / "effect_modification_sex_stratified.csv",
        index=False,
        encoding="utf-8-sig",
    )
    logging.info("Refit de interacciones finalizado.")


if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        force=True,
    )
    main()
