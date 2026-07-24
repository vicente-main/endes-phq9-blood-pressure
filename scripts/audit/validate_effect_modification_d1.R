#!/usr/bin/env Rscript

# Valida el pooling D1 de Python contra la implementación de referencia de
# mitml:::.D1, utilizada por mice::D1 para pruebas Wald multivariadas.

suppressPackageStartupMessages({
  library(jsonlite)
  library(mitml)
})

args <- commandArgs(trailingOnly = TRUE)
if (length(args) != 3L) {
  stop(
    "Uso: validate_effect_modification_d1.R ",
    "<effect_modification_d1_inputs.json> ",
    "<effect_modification_panel.csv> <salida_validacion.csv>"
  )
}

input_json <- normalizePath(args[[1]], mustWork = TRUE)
panel_csv <- normalizePath(args[[2]], mustWork = TRUE)
output_csv <- args[[3]]

records <- fromJSON(input_json, simplifyVector = FALSE)
panel <- read.csv(panel_csv, check.names = FALSE, fileEncoding = "UTF-8-BOM")

models <- unique(vapply(records, function(x) x$model, character(1)))
rows <- vector("list", length(models))

for (model_index in seq_along(models)) {
  model_name <- models[[model_index]]
  model_records <- records[
    vapply(records, function(x) identical(x$model, model_name), logical(1))
  ]
  ord <- order(vapply(model_records, function(x) as.integer(x$imputation_id), integer(1)))
  model_records <- model_records[ord]

  common_terms <- Reduce(
    intersect,
    lapply(model_records, function(x) unlist(x$terms, use.names = FALSE))
  )
  if (length(common_terms) == 0L) {
    stop("No hay términos comunes para ", model_name)
  }

  k <- length(common_terms)
  m <- length(model_records)
  qhat <- matrix(NA_real_, nrow = k, ncol = m, dimnames = list(common_terms, NULL))
  uhat <- array(NA_real_, dim = c(k, k, m))
  df_resid <- numeric(m)

  for (j in seq_along(model_records)) {
    item <- model_records[[j]]
    terms <- unlist(item$terms, use.names = FALSE)
    coef <- as.numeric(unlist(item$coef, use.names = FALSE))
    vcov <- do.call(rbind, lapply(item$vcov, function(x) as.numeric(unlist(x))))
    selection <- match(common_terms, terms)
    if (anyNA(selection)) {
      stop("Términos inconsistentes en ", model_name, ", imputación ", j)
    }
    qhat[, j] <- coef[selection]
    uhat[, , j] <- vcov[selection, selection, drop = FALSE]
    df_resid[[j]] <- as.numeric(item$df_resid)
  }

  df_com <- mean(df_resid)
  reference <- mitml:::.D1(Qhat = qhat, Uhat = uhat, df.com = df_com)
  reference_p <- pf(reference$F, df1 = reference$k, df2 = reference$v, lower.tail = FALSE)
  python_row <- panel[panel$model == model_name, , drop = FALSE]
  if (nrow(python_row) != 1L) {
    stop("El panel no contiene una fila única para ", model_name)
  }

  rows[[model_index]] <- data.frame(
    model = model_name,
    imputations = m,
    terms = paste(common_terms, collapse = ","),
    df_num = k,
    df_com = df_com,
    python_F = python_row$statistic,
    mitml_F = as.numeric(reference$F),
    abs_delta_F = abs(python_row$statistic - as.numeric(reference$F)),
    python_df_den = python_row$df_den,
    mitml_df_den = as.numeric(reference$v),
    abs_delta_df_den = abs(python_row$df_den - as.numeric(reference$v)),
    python_riv = python_row$riv,
    mitml_riv = as.numeric(reference$r),
    abs_delta_riv = abs(python_row$riv - as.numeric(reference$r)),
    python_p = python_row$p_value,
    mitml_p = reference_p,
    abs_delta_p = abs(python_row$p_value - reference_p),
    match_1e_10 = (
      abs(python_row$statistic - as.numeric(reference$F)) < 1e-10 &&
      abs(python_row$df_den - as.numeric(reference$v)) < 1e-8 &&
      abs(python_row$riv - as.numeric(reference$r)) < 1e-10 &&
      abs(python_row$p_value - reference_p) < 1e-10
    ),
    stringsAsFactors = FALSE
  )
}

result <- do.call(rbind, rows)
dir.create(dirname(output_csv), recursive = TRUE, showWarnings = FALSE)
write.csv(result, output_csv, row.names = FALSE, fileEncoding = "UTF-8")

if (!all(result$match_1e_10)) {
  print(result)
  stop("La implementación Python no coincide con mitml:::.D1")
}

message(
  "Validación D1 completada: ",
  nrow(result),
  " modelos coinciden con mitml:::.D1."
)
