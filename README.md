# Depressive symptomatology (PHQ-9) and elevated blood pressure — ENDES 2019–2025

[![DOI](https://zenodo.org/badge/1298665481.svg)](https://doi.org/10.5281/zenodo.21328300)

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
   cubic splines, effect-modification tests, and a logistic sensitivity model. Scalar estimates are
   pooled with Rubin's rules, spline non-linearity is combined using D2, and multivariate interaction
   tests use the Li–Rubin D1 F procedure.

## Environment

| Component | Version |
|---|---|
| Python | 3.13 |
| R | 4.5.3 |
| R · `survey` | 4.5 |
| R · `mitml` (independent D1 validation) | 0.4-5 |

Python dependencies are pinned in [`requirements.txt`](requirements.txt) (pandas 2.2, numpy 2.2,
scipy 1.15, scikit-learn 1.8.0, statsmodels 0.14, pyarrow 19, rpy2 3.6, matplotlib 3.10).

## Reproducibility

- **Multiple imputation:** parametric chained-equations imputation inspired by MICE, implemented
  with `sklearn.impute.IterativeImputer` (`sample_posterior=True`, `max_iter=20`), 20 imputations.
  This is not the R `mice` package and does not use predictive mean matching. Nine targets are
  declared; five contain values that are replaced. Education is not imputed. Continuous targets are
  constrained to the observed range and categorical targets are reassigned to an allowed category.
- **Random seed:** base `20260313`; imputation *k* uses `random_state = 20260313 + k` (k = 1..20).
- **Survey design (always):** `svydesign(id = ~HV001, strata = ~HV022, weights = ~PESO_FINAL,
  nest = TRUE)` with `options(survey.lonely.psu = "adjust")`.
- **Main association model** (`model_2`):
  `PRESION_ARTERIAL_ELEVADA ~ PHQ9_TOTAL + QS23 + factor(QSSEXO) + factor(QS25N) + factor(HV025) +
  factor(HV270) + factor(VIOLENCIA_PAREJA) + factor(ANIO) + factor(ALTITUD_CAT3)`,
  family `quasipoisson(link = "log")` → prevalence ratio.
- **Effect modification:** sex and survey year are theory-informed; area, wealth, prior hypertension
  diagnosis, and altitude are exploratory. Holm correction is applied to the four exploratory tests.
  The prior-diagnosis interaction excludes missing diagnosis status and is a binary 1-df contrast.
- **Independent D1 check:** `scripts/audit/validate_effect_modification_d1.R` reconstructs the exact
  coefficient/covariance inputs and reproduces all six tests with `mitml:::.D1`.

## Data

ENDES microdata (INEI, Peru) are freely available at
<https://proyectos.inei.gob.pe/microdatos/>.
Modules used: `CSALUD01` (individual health), `RECH0` (household; includes altitude `HV040`),
`RECH23` (wealth index `HV270`), `REC42` (pregnancy `V454`), `RE223132` (postpartum `V222`), for
survey years 2019–2025. Place the raw yearly archives in `data/input/raw_endes/` named `{YEAR}-*.zip`
before running.

## Usage

```powershell
# 1. Create the virtual environment and install dependencies
.\scripts\setup_env.ps1

# 2. Full pipeline for the pooled 2019-2025 surveys (weight divisor = 7)
.\.venv\Scripts\python.exe .\scripts\run_pipeline.py --config config\pipeline_config_2025.json

# 3. (Optional) re-run analytic phases on the already-built base
.\.venv\Scripts\python.exe .\scripts\run_analysis.py --config config\pipeline_config_2025.json --run-r-bridge --r-sections tables,models,figures
```

Outputs are written under `data/output_2025/` (analytic base, `qc/`,
`analysis/{tables,models,figures,imputed}/`).

## Frozen validation outputs for v1.2.0

The Zenodo release candidate includes the small, non-microdata artifacts needed to audit the
revision:

- `effect_modification_panel.csv`;
- `effect_modification_d1_inputs.json`;
- `effect_modification_d1_mitml_validation.csv`;
- `effect_modification_sex_stratified.csv`;
- `table2_bivariate_observed_applicable.csv`;
- `mice_observed_imputed_diagnostics.csv` (legacy filename; the procedure is described above);
- `ANALYTIC_CHANGE_AUDIT_2026-07-23.{csv,md}`.

No ENDES microdata or imputed person-level datasets are redistributed.

## Repository layout

```
src/endes_pipeline/   pipeline (Stage 1) and analysis (Stage 2) package
scripts/              orchestrators (run_pipeline, run_analysis, setup_env) and table/figure builders
config/               pipeline configuration (2019-2024 and 2019-2025 cohorts)
requirements.txt      pinned Python dependencies
```

For the `v1.2.0` analytical revision, the canonical workflow is the core pipeline plus the dated
scripts in `scripts/audit/*2026_07_23*` listed in the release notes. Earlier undated audit/build
scripts are retained in the Git history for provenance; some encode superseded labels or numerical
results and must not be used to regenerate the current manuscript. The Zenodo `v1.2.0` archive
therefore excludes those historical helpers.

> R is discovered from the environment by default. If auto-detection fails, set `r_home` and/or
> `r_lib_user` in `config/*.json` to paths valid on your machine. Code identifiers and log messages
> are in Spanish.

## Citation

If you use this code, please cite the associated article (citation to be added upon publication).

## License

Released under the [MIT License](LICENSE).
