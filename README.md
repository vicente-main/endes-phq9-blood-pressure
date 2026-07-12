# Depressive symptomatology (PHQ-9) and elevated blood pressure — ENDES 2019–2025

Reproducible analysis pipeline for a cross-sectional study of the association between depressive
symptomatology (PHQ-9) and elevated blood pressure in Peruvian adults, using Peru's Demographic and
Family Health Survey (ENDES, 2019–2025; n = 191 757).

This repository contains the **analysis code**. ENDES microdata are **not** redistributed here; they
are publicly available from INEI (see [Data](#data)).

## Overview

Two-stage pipeline:

1. **Stage 1 — build the analytic base** (`src/endes_pipeline/pipeline.py`): reads the raw ENDES
   module archives, merges household/individual modules, harmonizes variables, derives the PHQ-9
   score and the blood-pressure outcome, applies the reproductive exclusions, and produces the
   STROBE filter flow.
2. **Stage 2 — analysis** (`src/endes_pipeline/analysis.py`): multiple imputation
   (scikit-learn `IterativeImputer`, 20 imputations), VIF, and survey-weighted models in R (`survey`)
   via an `rpy2` bridge — prevalence ratios from quasi-Poisson models, Rao–Scott tests, restricted
   cubic splines, and a logistic sensitivity model. Estimates are pooled across imputations with
   Rubin's rules.

## Environment

| Component | Version |
|---|---|
| Python | 3.13 |
| R | 4.5.3 |
| R · `survey` | 4.5 |

Python dependencies are pinned in [`requirements.txt`](requirements.txt) (pandas 2.2, numpy 2.2,
scipy 1.15, scikit-learn 1.8.0, statsmodels 0.14, pyarrow 19, rpy2 3.6, matplotlib 3.10).

## Reproducibility

- **Multiple imputation:** MICE-style chained equations via `sklearn.impute.IterativeImputer`
  (`sample_posterior=True`, `max_iter=20`), 20 imputations. Continuous targets are constrained to the
  observed range; categorical/ordinal targets are reassigned to the nearest observed category.
- **Random seed:** base `20260313`; imputation *k* uses `random_state = 20260313 + k` (k = 1..20).
- **Survey design (always):** `svydesign(id = ~HV001, strata = ~HV022, weights = ~PESO_FINAL,
  nest = TRUE)` with `options(survey.lonely.psu = "adjust")`.
- **Primary model** (`model_2`, total effect):
  `PRESION_ARTERIAL_ELEVADA ~ PHQ9_TOTAL + QS23 + factor(QSSEXO) + factor(QS25N) + factor(HV025) +
  factor(HV270) + factor(VIOLENCIA_PAREJA) + factor(ANIO) + factor(ALTITUD_CAT3)`,
  family `quasipoisson(link = "log")` → prevalence ratio.

## Data

ENDES microdata (INEI, Peru) are freely available at <http://iinei.inei.gob.pe/microdatos/>.
Modules used: `CSALUD01` (individual health), `RECH0` (household; includes altitude `HV040`),
`RECH23` (wealth index `HV270`), `REC42` (pregnancy `V454`), `RE223132` (postpartum `V222`), for
survey years 2019–2025. Place the raw yearly archives in `data/input/raw_endes/` named `{YEAR}-*.zip`
before running.

## Usage

```powershell
# 1. Create the virtual environment and install dependencies
.\scripts\setup_env.ps1

# 2. Full pipeline for the 2019-2025 cohort (weight divisor = 7)
.\.venv\Scripts\python.exe .\scripts\run_pipeline.py --config config\pipeline_config_2025.json

# 3. (Optional) re-run analytic phases on the already-built base
.\.venv\Scripts\python.exe .\scripts\run_analysis.py --config config\pipeline_config_2025.json --run-r-bridge --r-sections tables,models,figures
```

Outputs are written under `data/output_2025/` (analytic base, `qc/`,
`analysis/{tables,models,figures,imputed}/`).

## Repository layout

```
src/endes_pipeline/   pipeline (Stage 1) and analysis (Stage 2) package
scripts/              orchestrators (run_pipeline, run_analysis, setup_env) and table/figure builders
config/               pipeline configuration (2019-2024 and 2019-2025 cohorts)
requirements.txt      pinned Python dependencies
```

> Configuration files reference the analyst's local R library path (`r_lib_user` in
> `config/*.json`); adjust it to your machine. Code identifiers and log messages are in Spanish.

## Citation

If you use this code, please cite the associated article (citation to be added upon publication).

## License

Released under the [MIT License](LICENSE).
