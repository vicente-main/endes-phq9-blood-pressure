"""Tabla 2 (bivariado) — fuente de DATOS OBSERVADOS (Opción B).

Calcula, por variable categórica, la prueba F de Rao-Scott (svychisq) de asociación
con PA elevada sobre los DATOS OBSERVADOS (complete-case, un solo dataset),
NO pooled sobre imputaciones. Escribe rao_scott_observado.csv que consume
build_post_auditoria.build_tabla_2.
"""
from __future__ import annotations

import subprocess
import tempfile
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[2]
BASE_PARQUET = ROOT / "data" / "output" / "endes_hta_depresion_2019_2024.parquet"
OUT_CSV = ROOT / "data" / "output" / "analysis" / "tables" / "rao_scott_observado.csv"
RSCRIPT_EXE = Path("C:/Program Files/R/R-4.5.3/bin/Rscript.exe")
R_LIB_USER = Path("C:/Users/Trabajo/Documents/R/win-library/4.5")

# Variables categóricas de la Tabla 2 (incluye sexo, que en Tabla 1 es panel)
CAT_VARS = ["SEVERIDAD_DEPRESIVA", "QSSEXO", "QS25N", "HV025", "HV270",
            "VIOLENCIA_PAREJA", "ALCOHOL_PROBLEMATICO", "QS201",
            "QS109", "CALIDAD_DIETA", "ALTITUD_CAT3"]

_RSCRIPT = r"""
.libPaths(c('{r_lib_user}', .libPaths()))
suppressPackageStartupMessages({{ library(survey) }})
options(survey.lonely.psu="adjust")
args <- commandArgs(trailingOnly = TRUE)
csv_in <- args[1]; csv_vars <- args[2]; csv_out <- args[3]
df <- read.csv(csv_in, stringsAsFactors = FALSE, fileEncoding = "UTF-8")
vars <- read.csv(csv_vars, stringsAsFactors = FALSE, fileEncoding = "UTF-8")$variable
res <- data.frame(variable=character(), statistic=numeric(), p_value=numeric(),
                  n_complete=integer(), stringsAsFactors=FALSE)
for (v in vars) {{
  sub <- df[!is.na(df[[v]]), , drop=FALSE]
  sub[[v]] <- as.factor(sub[[v]])
  sub$PRESION_ARTERIAL_ELEVADA <- as.factor(sub$PRESION_ARTERIAL_ELEVADA)
  d <- survey::svydesign(id=~HV001, strata=~HV022, weights=~PESO_FINAL, data=sub, nest=TRUE)
  f <- as.formula(paste0("~`", v, "` + PRESION_ARTERIAL_ELEVADA"))
  r <- tryCatch(suppressWarnings(survey::svychisq(f, design=d, statistic="F")), error=function(e) NULL)
  Fst <- if (is.null(r)) NA_real_ else as.numeric(r$statistic)
  pv  <- if (is.null(r)) NA_real_ else as.numeric(r$p.value)
  res <- rbind(res, data.frame(variable=v, statistic=Fst, p_value=pv,
                               n_complete=nrow(sub), stringsAsFactors=FALSE))
  cat(sprintf("%s: F=%.4g p=%.5g n=%d\n", v, Fst, pv, nrow(sub)))
}}
write.csv(res, csv_out, row.names=FALSE, fileEncoding="UTF-8")
cat("OK\n")
"""


def main() -> None:
    print("Tabla 2 (bivariado, datos observados):")
    df = pd.read_parquet(BASE_PARQUET)
    needed = ["HV001", "HV022", "PESO_FINAL", "PRESION_ARTERIAL_ELEVADA"] + CAT_VARS
    sub = df[needed].copy()
    for c in sub.columns:
        if pd.api.types.is_integer_dtype(sub[c].dtype):
            sub[c] = sub[c].astype("float64")
    with tempfile.TemporaryDirectory() as tmp:
        t = Path(tmp)
        ci, cv, co, rs = t / "data.csv", t / "vars.csv", t / "out.csv", t / "p.R"
        sub.to_csv(ci, index=False, encoding="utf-8")
        pd.DataFrame({"variable": CAT_VARS}).to_csv(cv, index=False, encoding="utf-8")
        rs.write_text(_RSCRIPT.format(r_lib_user=str(R_LIB_USER).replace("\\", "/")), encoding="utf-8")
        print("  ejecutando Rscript (Rao-Scott F sobre datos observados)...")
        proc = subprocess.run([str(RSCRIPT_EXE), "--vanilla", str(rs), str(ci), str(cv), str(co)],
                              capture_output=True, text=True, encoding="utf-8", errors="replace")
        if proc.returncode != 0:
            print("STDOUT:", proc.stdout); print("STDERR:", proc.stderr)
            raise RuntimeError("Rscript falló")
        for line in [l for l in proc.stdout.splitlines() if l.strip()][-20:]:
            print(f"    R| {line}")
        res = pd.read_csv(co)
    OUT_CSV.parent.mkdir(parents=True, exist_ok=True)
    res.to_csv(OUT_CSV, index=False, encoding="utf-8-sig")
    print(f"  -> {OUT_CSV.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
