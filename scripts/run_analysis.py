import argparse
import json
import logging
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from endes_pipeline.analysis import build_analysis_settings, run_analysis


def _load_config(config_path: Path) -> dict:
    return json.loads(config_path.read_text(encoding="utf-8-sig"))


def _resolve_dataset_path(config_data: dict, explicit_path: str | None) -> Path:
    if explicit_path:
        return Path(explicit_path)

    output_dir = Path(config_data["paths"]["output_dir"])
    dataset_base = output_dir / config_data["output"]["dataset_base_name"]
    parquet_path = dataset_base.with_suffix(".parquet")
    csv_path = dataset_base.with_suffix(".csv")

    if parquet_path.exists():
        return parquet_path
    if csv_path.exists():
        return csv_path
    return parquet_path


def _load_unified_dataset(dataset_path: Path) -> pd.DataFrame:
    if not dataset_path.exists():
        raise FileNotFoundError(f"No se encontro la base analitica en {dataset_path}.")

    suffix = dataset_path.suffix.lower()
    if suffix == ".parquet":
        return pd.read_parquet(dataset_path)
    if suffix == ".csv":
        return pd.read_csv(dataset_path, low_memory=False)

    raise ValueError(f"Formato de dataset no soportado para {dataset_path}. Use .parquet o .csv.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Ejecuta fases analiticas parciales sin reconstruir la base ENDES.")
    parser.add_argument(
        "--config",
        default=str(ROOT / "config" / "pipeline_config.json"),
        help="Ruta al archivo de configuracion JSON.",
    )
    parser.add_argument(
        "--dataset",
        help="Ruta a la base analitica ya unificada (.parquet o .csv). Solo se necesita si se ejecuta MICE.",
    )
    parser.add_argument("--run-mice", action="store_true", help="Regenera imputaciones MICE desde la base ya unificada.")
    parser.add_argument("--run-vif", action="store_true", help="Recalcula VIF usando la primera imputacion disponible.")
    parser.add_argument(
        "--run-r-bridge",
        action="store_true",
        help="Ejecuta el puente analitico con R usando las imputaciones ya disponibles.",
    )
    parser.add_argument(
        "--r-sections",
        help="Secciones del puente R separadas por coma: tables,models,figures o all.",
    )
    args = parser.parse_args()

    if not (args.run_mice or args.run_vif or args.run_r_bridge):
        parser.error("Active al menos una fase: --run-mice, --run-vif o --run-r-bridge.")

    logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s", force=True)

    config_path = Path(args.config)
    config_data = _load_config(config_path)
    analysis_config = dict(config_data.get("analysis", {}))
    if args.r_sections:
        analysis_config["r_sections"] = args.r_sections

    output_dir = Path(config_data["paths"]["output_dir"])
    settings = build_analysis_settings(output_dir, analysis_config)

    if args.run_mice:
        dataset_path = _resolve_dataset_path(config_data, args.dataset)
        unified = _load_unified_dataset(dataset_path)
        logging.info("Base analitica cargada desde %s con %s filas", dataset_path, len(unified))
    else:
        unified = pd.DataFrame()

    run_analysis(
        unified,
        settings,
        run_mice=args.run_mice,
        run_vif=args.run_vif,
        run_r_bridge=args.run_r_bridge,
    )
