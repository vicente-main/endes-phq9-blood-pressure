import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from endes_pipeline import run_pipeline


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Ejecuta el pipeline de limpieza y unificacion ENDES.")
    parser.add_argument(
        "--config",
        default=str(ROOT / "config" / "pipeline_config.json"),
        help="Ruta al archivo de configuracion JSON.",
    )
    args = parser.parse_args()
    run_pipeline(args.config)
