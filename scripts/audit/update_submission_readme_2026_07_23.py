"""Actualiza el mapa del paquete de envío tras ejecutar la guía del 23-07-2026."""

from __future__ import annotations

import hashlib
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parents[2]
SUBMISSION = ROOT / "ENVIO_PLOS_ONE_2019-2025"
REVIEWERS = SUBMISSION / "Revisores"
MANUSCRIPT = REVIEWERS / "MANUSCRITO_EN_PLOS_ONE.docx"
README = SUBMISSION / "README_manifest.txt"
MODELS = ROOT / "data" / "output_2025" / "analysis" / "models"
FIGURES = ROOT / "data" / "output_2025" / "analysis" / "figures"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def fmt_p(value: float) -> str:
    return "<0.001" if value < 0.001 else f"{value:.3f}"


def main() -> None:
    panel = pd.read_csv(MODELS / "effect_modification_panel.csv").set_index("modificador")
    main_models = pd.read_csv(MODELS / "table3_main_models.csv")
    main = main_models.loc[
        (main_models["model"] == "model_2")
        & (main_models["term"] == "PHQ9_TOTAL")
    ].iloc[0]
    spline = pd.read_csv(FIGURES / "spline_nonlinearity_summary.csv").iloc[0]

    lines = [
        "PAQUETE DE ENVÍO — PLOS ONE",
        "Estudio: Severidad de síntomas depresivos y presión arterial elevada (ENDES 2019–2025)",
        "Muestra analítica principal: n = 191 757 adultos",
        "Revisión ejecutada: 2026-07-23",
        "",
        "===================================================================",
        "ARTEFACTO ACTIVO",
        "",
        "El manuscrito vigente y corregido es:",
        "  Revisores/MANUSCRITO_EN_PLOS_ONE.docx",
        f"  Tamaño: {MANUSCRIPT.stat().st_size:,} bytes",
        f"  SHA-256: {sha256(MANUSCRIPT)}",
        "",
        "No usar para el envío:",
        "  MANUSCRITO_EN_PLOS_ONE.docx (copia anterior del paquete principal).",
        "  ENVIO_PLOS_ONE.zip (archivo histórico del 2026-07-16; no contiene esta revisión).",
        "",
        "===================================================================",
        "RESULTADOS CANÓNICOS SINCRONIZADOS",
        "",
        (
            f"  Modelo 2 principal .... PR {float(main['exp_estimate']):.7f}; "
            f"IC95% {float(main['exp_ci_low']):.7f}–{float(main['exp_ci_high']):.7f}; "
            f"p {float(main['p_value']):.7f}; n {int(round(float(main['mean_n_obs']))):,}"
        ),
        (
            f"  Spline no linealidad .. p D2 {float(spline['d2_p_value']):.7f}"
        ),
        "",
        "  Modificación de efecto — D1 F de Li–Rubin:",
    ]
    labels = {
        "sex": "sexo",
        "year": "año",
        "area": "área",
        "wealth": "riqueza",
        "dxhta": "diagnóstico HTA previo",
        "altitude": "altitud",
    }
    for modifier in ["sex", "year", "area", "wealth", "dxhta", "altitude"]:
        item = panel.loc[modifier]
        holm = item.get("p_holm_exploratorios")
        suffix = ""
        if pd.notna(holm):
            suffix = f"; Holm {fmt_p(float(holm))}"
        lines.append(
            f"    {labels[modifier]:24s} F({int(item['df_num'])}, "
            f"{float(item['df_den']):.1f}) = {float(item['statistic']):.3f}; "
            f"p {fmt_p(float(item['p_value']))}{suffix}"
        )
    lines.extend(
        [
            "",
            "La interacción PHQ-9 × diagnóstico previo excluye 192 estados faltantes,",
            "usa un contraste binario de 1 df y n = 191 565. Los seis D1 fueron",
            "reproducidos con mitml 0.4-5 en R 4.5.3 a tolerancia estricta.",
            "",
            "===================================================================",
            "ARCHIVOS ACTIVOS DEL ENVÍO",
            "",
            "  Manuscrito: Revisores/MANUSCRITO_EN_PLOS_ONE.docx",
            "  Carta: COVER_LETTER.docx",
            "  Figuras principales: Fig1.tif, Fig2.tif",
            "  Figuras suplementarias: S1_Fig.tif, S2_Fig.tif",
            "  Tablas suplementarias: S1_Table.xlsx … S11_Table.xlsx",
            "  Checklists: S1_Checklist.docx (STROBE), S2_Checklist.docx (RECORD complementario)",
            "  Campos de envío: PARA_EDITORIAL_MANAGER.md",
            "",
            "Las cuatro figuras TIFF están en RGB y a 300 dpi. Las Tablas 1–5 están",
            "incrustadas en el manuscrito.",
            "",
            "===================================================================",
            "BLOQUEOS EXTERNOS ANTES DEL ENVÍO",
            "",
            "1. Ética (P0-03): incorporar una determinación institucional verificable",
            "   con comité, tipo de decisión, número/fecha de resolución y tratamiento",
            "   del consentimiento secundario. Véase:",
            "   Revisores/SOLICITUD_CIEI_P0-03_2026-07-23.md",
            "",
            "2. Zenodo: publicar una nueva versión bajo el DOI de concepto",
            "   10.5281/zenodo.21328300 que contenga P0-01/P0-02, el validador D1 y",
            "   las salidas congeladas. v1.1.0 antecede esta revisión y la API oficial",
            "   reporta acceso de archivos restringido. El candidato local preparado es:",
            "   Revisores/ZENODO_RELEASE_CANDIDATE_v1.2.0.zip",
            (
                "   SHA-256: "
                + sha256(
                    REVIEWERS / "ZENODO_RELEASE_CANDIDATE_v1.2.0.zip"
                )
            ),
            "   Debe publicarse mediante New version con acceso público y verificarse",
            "   mediante una descarga anónima.",
            "",
            "No enviar mientras el manuscrito conserve la marca EDITORIAL HOLD.",
            "",
            "Fuente de verdad analítica: data/output_2025/analysis/.",
        ]
    )
    README.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"README actualizado: {README.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
