#!/usr/bin/env Rscript

# Exporta la salida canónica que corresponde exactamente a la Tabla 2 del
# manuscrito: pruebas bivariadas sobre los datos observados y, para alcohol y
# tabaco, únicamente sobre quienes recibieron la pregunta aplicable. No usa las
# imputaciones múltiples.

suppressPackageStartupMessages(library(survey))
options(survey.lonely.psu = "adjust")

args <- commandArgs(trailingOnly = TRUE)
if (length(args) != 2L) {
  stop(
    paste(
      "Uso: export_table2_observed_applicable_2026_07_23.R",
      "<base_analitica_csv> <salida_csv>"
    )
  )
}

input_path <- args[[1]]
output_path <- args[[2]]

df <- read.csv(input_path, stringsAsFactors = FALSE, check.names = FALSE)

required <- c(
  "HV001", "HV022", "PESO_FINAL", "PRESION_ARTERIAL_ELEVADA",
  "SEVERIDAD_DEPRESIVA", "QSSEXO", "QS25N", "HV025", "HV270",
  "VIOLENCIA_PAREJA", "ALCOHOL_PROBLEMATICO", "QS201", "QS109",
  "CALIDAD_DIETA", "ALTITUD_CAT3"
)
missing_columns <- setdiff(required, names(df))
if (length(missing_columns)) {
  stop("Faltan columnas: ", paste(missing_columns, collapse = ", "))
}

specifications <- data.frame(
  variable = c(
    "SEVERIDAD_DEPRESIVA", "QSSEXO", "QS25N", "HV025", "HV270",
    "VIOLENCIA_PAREJA", "ALCOHOL_PROBLEMATICO", "QS201", "QS109",
    "CALIDAD_DIETA", "ALTITUD_CAT3"
  ),
  label = c(
    "PHQ-9 severity (5 levels)", "Sex", "Educational level",
    "Area of residence", "Wealth quintile", "Intimate partner violence",
    "Problematic alcohol use", "Tobacco use (last 30 days)",
    "Diabetes diagnosis", "Diet quality", "Altitude (m a.s.l.)"
  ),
  denominator_rule = c(
    rep("observed", 6), "observed_applicable", "observed_applicable",
    rep("observed", 3)
  ),
  stringsAsFactors = FALSE
)

build_design <- function(data) {
  svydesign(
    id = ~HV001,
    strata = ~HV022,
    weights = ~PESO_FINAL,
    data = data,
    nest = TRUE
  )
}

test_one <- function(variable, label, denominator_rule) {
  keep <- !is.na(df[[variable]]) & !is.na(df$PRESION_ARTERIAL_ELEVADA)
  sub <- df[keep, , drop = FALSE]
  sub$.table2_predictor <- factor(sub[[variable]])
  sub$.table2_outcome <- factor(sub$PRESION_ARTERIAL_ELEVADA)
  design <- build_design(sub)
  test <- suppressWarnings(
    svychisq(
      ~.table2_predictor + .table2_outcome,
      design = design,
      statistic = "F"
    )
  )
  data.frame(
    variable = variable,
    label = label,
    analysis = "observed_data",
    denominator_rule = denominator_rule,
    n_unweighted = nrow(sub),
    levels_tested = nlevels(sub$.table2_predictor),
    statistic = unname(test$statistic),
    df_num = unname(test$parameter[[1]]),
    df_den = unname(test$parameter[[2]]),
    p_value = unname(test$p.value),
    survey_specification = paste0(
      "svydesign(id=~HV001, strata=~HV022, weights=~PESO_FINAL, ",
      "nest=TRUE); survey.lonely.psu='adjust'"
    ),
    stringsAsFactors = FALSE
  )
}

rows <- lapply(
  seq_len(nrow(specifications)),
  function(i) {
    test_one(
      specifications$variable[[i]],
      specifications$label[[i]],
      specifications$denominator_rule[[i]]
    )
  }
)
result <- do.call(rbind, rows)

expected_n <- c(
  SEVERIDAD_DEPRESIVA = 191757,
  QSSEXO = 191757,
  QS25N = 191754,
  HV025 = 191757,
  HV270 = 191757,
  VIOLENCIA_PAREJA = 191728,
  ALCOHOL_PROBLEMATICO = 112673,
  QS201 = 31232,
  QS109 = 191550,
  CALIDAD_DIETA = 183327,
  ALTITUD_CAT3 = 191757
)
observed_n <- setNames(result$n_unweighted, result$variable)
if (!all(as.numeric(observed_n[names(expected_n)]) == as.numeric(expected_n))) {
  stop(
    "Los denominadores no coinciden con la muestra congelada. Observados: ",
    paste(names(observed_n), observed_n, sep = "=", collapse = ", ")
  )
}

dir.create(dirname(output_path), recursive = TRUE, showWarnings = FALSE)
write.csv(result, output_path, row.names = FALSE, fileEncoding = "UTF-8")
message("Tabla 2 observada/aplicable exportada: ", output_path)
