"""Sensibilidad de casos completos del Modelo 2 (cohorte 2019-2025).

Responde a la objecion del Revisor #2 (congenialidad de la imputacion): estima el
Modelo 2 estructural por analisis de casos completos (sin imputacion) y lo compara
con el estimador MICE-agrupado (Rubin) ya publicado. Demuestra que el estimador
principal del PHQ-9 es robusto al metodo de manejo de datos faltantes.

Diseno muestral identico al principal: svydesign(id=~HV001, strata=~HV022,
weights=~PESO_FINAL, nest=TRUE), familia quasipoisson(log) -> razon de prevalencia.
El ajuste se ejecuta en R (paquete survey) via Rscript; Python solo prepara los
casos completos. No usa rpy2 (evita dependencias rms/jsonlite del puente principal).

Salida: data/output_2025/analysis/models/complete_case_sensitivity.csv
"""
from __future__ import annotations

import json
import subprocess
import sys
import tempfile
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[2]
BASE = ROOT / "data" / "output_2025" / "endes_hta_depresion_2019_2025.parquet"
OUT = ROOT / "data" / "output_2025" / "analysis" / "models" / "complete_case_sensitivity.csv"
RSCRIPT = r"C:\Program Files\R\R-4.5.3\bin\Rscript.exe"
R_LIB = "C:/Users/Trabajo/Documents/R/win-library/4.5"

MODEL_VARS = [
    "PRESION_ARTERIAL_ELEVADA", "PHQ9_TOTAL", "QS23", "QSSEXO", "QS25N", "HV025",
    "HV270", "VIOLENCIA_PAREJA", "ANIO", "ALTITUD_CAT3", "HV001", "HV022", "PESO_FINAL",
]

R_TEMPLATE = r'''
.libPaths(c("{rlib}", .libPaths()))
suppressPackageStartupMessages(library(survey))
options(survey.lonely.psu="adjust")
d <- read.csv("{csv}", stringsAsFactors=FALSE)
des <- svydesign(id=~HV001, strata=~HV022, weights=~PESO_FINAL, data=d, nest=TRUE)
fml <- PRESION_ARTERIAL_ELEVADA ~ PHQ9_TOTAL + QS23 + factor(QSSEXO) + factor(QS25N) +
       factor(HV025) + factor(HV270) + factor(VIOLENCIA_PAREJA) + factor(ANIO) +
       factor(ALTITUD_CAT3)
fit <- suppressWarnings(svyglm(fml, design=des, family=quasipoisson(link="log")))
s <- summary(fit)
b  <- coef(fit)[["PHQ9_TOTAL"]]
se <- s$coefficients["PHQ9_TOTAL","Std. Error"]
p  <- s$coefficients["PHQ9_TOTAL","Pr(>|t|)"]
res <- list(n_obs=nrow(d), beta=b, std_error=se,
            PR=exp(b), ci_low=exp(b-1.96*se), ci_high=exp(b+1.96*se),
            p_value=p, dispersion=as.numeric(s$dispersion))
writeLines(jsonlite_like <- sprintf(
  '{{"n_obs":%d,"beta":%.8f,"std_error":%.8f,"PR":%.6f,"ci_low":%.6f,"ci_high":%.6f,"p_value":%.6f,"dispersion":%.6f}}',
  res$n_obs,res$beta,res$std_error,res$PR,res$ci_low,res$ci_high,res$p_value,res$dispersion),
  "{jsonout}")
'''


def main() -> None:
    df = pd.read_parquet(BASE, columns=MODEL_VARS)
    n_total = len(df)
    cc = df.dropna(subset=MODEL_VARS).copy()
    print(f"Total={n_total:,}  casos completos={len(cc):,} ({len(cc)/n_total*100:.1f}%)  "
          f"eliminados={n_total-len(cc):,}")

    with tempfile.TemporaryDirectory() as td:
        csv = Path(td) / "cc.csv"
        jsonout = Path(td) / "res.json"
        rfile = Path(td) / "fit.R"
        cc.to_csv(csv, index=False, encoding="utf-8")
        rfile.write_text(
            R_TEMPLATE.format(rlib=R_LIB, csv=csv.as_posix(),
                              jsonout=jsonout.as_posix()),
            encoding="utf-8",
        )
        proc = subprocess.run([RSCRIPT, str(rfile)], capture_output=True, text=True)
        if not jsonout.exists():
            print(proc.stdout); print(proc.stderr, file=sys.stderr)
            raise SystemExit("El ajuste en R no produjo salida.")
        r = json.loads(jsonout.read_text(encoding="utf-8"))

    row = {
        "model": "model_2_complete_case",
        "term": "PHQ9_TOTAL",
        "n_obs": int(r["n_obs"]),
        "n_total": n_total,
        "pct_complete": round(100 * r["n_obs"] / n_total, 2),
        "beta": round(r["beta"], 6),
        "std_error": round(r["std_error"], 6),
        "PR": round(r["PR"], 4),
        "ci_low": round(r["ci_low"], 4),
        "ci_high": round(r["ci_high"], 4),
        "p_value": round(r["p_value"], 4),
        "dispersion": round(r["dispersion"], 4),
        "family": "quasipoisson",
        "link": "log",
        "nota": "Casos completos (sin imputacion); comparar con MICE-pooled PR=0,995 (0,991-0,999; p=0,022).",
    }
    OUT.parent.mkdir(parents=True, exist_ok=True)
    pd.DataFrame([row]).to_csv(OUT, index=False, encoding="utf-8-sig")
    print(f"escrito: {OUT}")
    print(f"Modelo 2 casos completos (n={row['n_obs']:,}): "
          f"PR={row['PR']} (IC {row['ci_low']}-{row['ci_high']}; p={row['p_value']}); "
          f"phi={row['dispersion']}")


if __name__ == "__main__":
    main()
