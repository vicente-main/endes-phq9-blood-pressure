from __future__ import annotations

import shutil
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch
from matplotlib.ticker import MultipleLocator
import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
OUTPUT_DIR = ROOT / "data" / "output"
QC_DIR = OUTPUT_DIR / "qc"
ANALYSIS_DIR = OUTPUT_DIR / "analysis"
PACKAGE_DIR = OUTPUT_DIR / "Para Enviar"
PUBLICATION_DIR = PACKAGE_DIR / "Para Publicar"

FLOW_STEP_LABELS = [
    ("qsresult_ok", "Cuestionario de salud valido (QSRESULT == 1)"),
    ("adult_18_plus", "Adultos de 18 anos o mas"),
    ("valid_sex", "Sexo biologico valido"),
    ("rech0_available", "Datos de hogar disponibles (RECH0)"),
    ("design_available", "Variables de diseno disponibles"),
    ("rech23_available", "Indice de riqueza disponible (RECH23)"),
    ("positive_health_weight", "Peso de salud positivo"),
    ("not_pregnant_or_not_mef", "No gestante o fuera de MEF"),
    ("not_puerperal_or_not_mef", "No puerperio temprano o fuera de MEF"),
    ("phq_available", "PHQ-9 disponible"),
    ("bp_available", "Presion arterial valida disponible"),
]


def _copy_file(src: Path, dest_name: str) -> None:
    target = PACKAGE_DIR / dest_name
    shutil.copy2(src, target)


def _build_strobe_table() -> pd.DataFrame:
    flow_df = pd.read_csv(QC_DIR / "filter_flow_by_year.csv")
    totals = flow_df.drop(columns=["year"]).sum(numeric_only=True)

    rows = [
        {
            "order": 0,
            "step": "start",
            "stage": "Muestra inicial ENDES 2019-2024",
            "n_before": int(totals["n_start"]),
            "n_excluded": 0,
            "n_after": int(totals["n_start"]),
            "pct_excluded_from_before": 0.0,
            "pct_remaining_from_start": 100.0,
        }
    ]

    n_start = int(totals["n_start"])
    for order, (step, label) in enumerate(FLOW_STEP_LABELS, start=1):
        before = int(totals[f"before_{step}"])
        dropped = int(totals[f"drop_{step}"])
        after = int(totals[f"after_{step}"])
        rows.append(
            {
                "order": order,
                "step": step,
                "stage": label,
                "n_before": before,
                "n_excluded": dropped,
                "n_after": after,
                "pct_excluded_from_before": round((dropped / before) * 100, 4) if before else 0.0,
                "pct_remaining_from_start": round((after / n_start) * 100, 4) if n_start else 0.0,
            }
        )

    return pd.DataFrame(rows)


def _draw_box(ax, x: float, y: float, w: float, h: float, text: str, facecolor: str) -> None:
    box = FancyBboxPatch(
        (x, y),
        w,
        h,
        boxstyle="round,pad=0.012,rounding_size=0.02",
        linewidth=1.4,
        edgecolor="#264653",
        facecolor=facecolor,
    )
    ax.add_patch(box)
    ax.text(x + w / 2, y + h / 2, text, ha="center", va="center", fontsize=10, color="#102a43", wrap=True)


def _draw_arrow(ax, x1: float, y1: float, x2: float, y2: float) -> None:
    arrow = FancyArrowPatch((x1, y1), (x2, y2), arrowstyle="-|>", mutation_scale=18, linewidth=1.4, color="#264653")
    ax.add_patch(arrow)


def _draw_side_connector(ax, main_x: float, main_y: float, main_w: float, side_x: float, side_y: float, side_h: float) -> None:
    start = (main_x + main_w, main_y + 0.05)
    end = (side_x, side_y + side_h / 2)
    connector = FancyArrowPatch(
        start,
        end,
        arrowstyle="-|>",
        mutation_scale=16,
        linewidth=1.2,
        color="#7a8b99",
        connectionstyle="angle3,angleA=0,angleB=90",
    )
    ax.add_patch(connector)


def _plot_strobe_figure(strobe_df: pd.DataFrame, out_png: Path, out_pdf: Path, out_svg: Path | None = None) -> None:
    n_start = int(strobe_df.loc[strobe_df["step"] == "start", "n_after"].iloc[0])
    after_qsresult = int(strobe_df.loc[strobe_df["step"] == "qsresult_ok", "n_after"].iloc[0])
    drop_qsresult = int(strobe_df.loc[strobe_df["step"] == "qsresult_ok", "n_excluded"].iloc[0])
    after_adult = int(strobe_df.loc[strobe_df["step"] == "adult_18_plus", "n_after"].iloc[0])
    drop_adult = int(strobe_df.loc[strobe_df["step"] == "adult_18_plus", "n_excluded"].iloc[0])
    after_structural = int(strobe_df.loc[strobe_df["step"] == "positive_health_weight", "n_after"].iloc[0])
    drop_gest = int(strobe_df.loc[strobe_df["step"] == "not_pregnant_or_not_mef", "n_excluded"].iloc[0])
    drop_puer = int(strobe_df.loc[strobe_df["step"] == "not_puerperal_or_not_mef", "n_excluded"].iloc[0])
    after_repro = int(strobe_df.loc[strobe_df["step"] == "not_puerperal_or_not_mef", "n_after"].iloc[0])
    drop_bp = int(strobe_df.loc[strobe_df["step"] == "bp_available", "n_excluded"].iloc[0])
    n_final = int(strobe_df.loc[strobe_df["step"] == "bp_available", "n_after"].iloc[0])

    fig, ax = plt.subplots(figsize=(13, 10))
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.axis("off")

    ax.text(
        0.5,
        0.97,
        "Figura 1. Diagrama de flujo STROBE para la selección muestral",
        ha="center",
        va="top",
        fontsize=15,
        fontweight="bold",
        color="#102a43",
    )

    main_x, main_w, main_h = 0.09, 0.5, 0.1
    side_x, side_w = 0.68, 0.24
    y_positions = [0.82, 0.64, 0.46, 0.28, 0.1]

    main_texts = [
        f"ENDES 2019-2024\nMicrodatos CSALUD01\nN = {n_start:,}",
        f"Cuestionarios de salud válidos\n(QSRESULT == 1)\nN = {after_qsresult:,}",
        f"Adultos de 18 años o más\nN = {after_adult:,}",
        f"Cohorte estructural elegible\ncon diseño, módulos y peso\nN = {after_structural:,}",
        f"Muestra analítica final\nPHQ-9 disponible y PA válida\nN = {n_final:,}",
    ]
    main_colors = ["#dceefb", "#edf6f9", "#edf6f9", "#edf6f9", "#d8f3dc"]

    for y, text, color in zip(y_positions, main_texts, main_colors):
        _draw_box(ax, main_x, y, main_w, main_h, text, color)

    for upper_y, lower_y in zip(y_positions, y_positions[1:]):
        _draw_arrow(ax, main_x + main_w / 2, upper_y, main_x + main_w / 2, lower_y + main_h)

    side_texts = [
        (0.73, f"Excluidos por\nQSRESULT != 1\nn = {drop_qsresult:,}", y_positions[0]),
        (0.55, f"Excluidos por\nedad < 18 años\nn = {drop_adult:,}", y_positions[1]),
        (
            0.37,
            f"Excluidas por gestación\no puerperio temprano en MEF\nn = {drop_gest + drop_puer:,}\n(gestación = {drop_gest:,}; puerperio = {drop_puer:,})",
            y_positions[3],
        ),
        (0.19, f"Excluidos por PA no válida\no medición faltante\nn = {drop_bp:,}", y_positions[4]),
    ]

    for y, text, main_y in side_texts:
        _draw_box(ax, side_x, y, side_w, 0.095, text, "#fef3c7")
        _draw_side_connector(ax, main_x, main_y, main_w, side_x, y, 0.095)

    ax.text(
        0.09,
        0.025,
        f"Nota: cohorte estructural con sexo válido, RECH0, diseño complejo, RECH23 y peso positivo. "
        f"Tras el filtro reproductivo quedaron {after_repro:,} observaciones antes del control final de PA.",
        fontsize=9,
        color="#334e68",
    )

    fig.tight_layout()
    fig.savefig(out_png, dpi=300, bbox_inches="tight")
    fig.savefig(out_pdf, bbox_inches="tight")
    if out_svg is not None:
        fig.savefig(out_svg, bbox_inches="tight")
    plt.close(fig)


def _plot_spline_figure(out_png: Path, out_pdf: Path, out_svg: Path | None = None) -> None:
    curve_df = pd.read_csv(ANALYSIS_DIR / "figures" / "spline_curve_pooled.csv")
    meta_df = pd.read_csv(ANALYSIS_DIR / "figures" / "spline_nonlinearity_summary.csv")

    curve_df = curve_df.copy()
    curve_df["phq9_total"] = pd.to_numeric(curve_df["label"], errors="coerce")
    curve_df = curve_df.sort_values("phq9_total").reset_index(drop=True)

    meta_row = meta_df.iloc[0]
    pooled_p = (
        float(meta_row["d2_p_value"])
        if "d2_p_value" in meta_row.index and pd.notna(meta_row["d2_p_value"])
        else float(meta_row["median_p_value"])
    )
    knots = "0, 4, 9, 14"

    x = curve_df["phq9_total"]
    y = curve_df["estimate"] * 100
    y_low = curve_df["ci_low"] * 100
    y_high = curve_df["ci_high"] * 100

    fig, ax = plt.subplots(figsize=(11, 6.5))
    ax.fill_between(x, y_low, y_high, color="#ffd6a5", alpha=0.35, label="IC 95%")
    ax.plot(x, y, color="#bb3e03", linewidth=2.4, label="Estimación pooled")
    ax.scatter(x, y, color="#9b2226", s=18, zorder=3)

    ax.set_title("Figura 2. Curva spline entre PHQ-9 y presión arterial elevada", fontsize=14, fontweight="bold")
    ax.set_xlabel("Puntaje PHQ-9")
    ax.set_ylabel("Prevalencia marginal estandarizada de presión arterial elevada (%)")
    ax.grid(axis="y", linestyle="--", alpha=0.35)
    ax.set_xlim(float(x.min()), float(x.max()))
    ax.set_ylim(5.5, 7.0)
    ax.yaxis.set_major_locator(MultipleLocator(0.25))

    subtitle = (
        f"RCS con 4 nodos ({knots}) sobre 20 imputaciones. "
        f"No se observó evidencia de no linealidad "
        f"(prueba D2 pooled, p = {pooled_p:.3f})."
    )
    ax.text(0.02, 0.96, subtitle, transform=ax.transAxes, ha="left", va="top", fontsize=10, color="#3d405b")
    ax.legend(frameon=False, loc="upper right")

    fig.tight_layout()
    fig.savefig(out_png, dpi=300, bbox_inches="tight")
    fig.savefig(out_pdf, bbox_inches="tight")
    if out_svg is not None:
        fig.savefig(out_svg, bbox_inches="tight")
    plt.close(fig)


def _write_manifest() -> None:
    manifest = """Entregables finales preparados para envio

Contenido principal:
- 01_base_analitica_final.csv
- 01_base_analitica_final.parquet
- 02_tabla_strobe_flujo.csv
- 03_tabla_1_caracteristicas_basales.csv
- 03a_tabla_1_deff_resumen.csv
- 04_tabla_2_bivariado.csv
- 05_tabla_3_modelos_principales.csv
- 05a_diagnosticos_modelos.csv
- 06_tabla_4_cascada_cuidado.csv
- 07_figura_1_strobe.png / .pdf
- 07_figura_1_strobe_datos.csv
- 08_figura_2_spline_phq9_pa.png / .pdf
- 08_figura_2_spline_phq9_pa_datos.csv
- 08a_spline_nolinealidad_resumen.csv
- 09_analisis_sensibilidad_interacciones.csv
- 09a_sensibilidad_logistica.csv
- 09b_items_omitidos.csv

Nota:
- TIEMPO_DX_HTA_MESES fue excluida formalmente del estudio porque QS103U/QS103C estuvieron vacios o no utilizables en el crudo.
"""
    (PACKAGE_DIR / "00_LEEME.txt").write_text(manifest, encoding="utf-8-sig")


def _copy_publication_graphics() -> None:
    PUBLICATION_DIR.mkdir(parents=True, exist_ok=True)
    for filename in [
        "07_figura_1_strobe.png",
        "07_figura_1_strobe.pdf",
        "07_figura_1_strobe.svg",
        "07_figura_1_strobe_datos.csv",
        "08_figura_2_spline_phq9_pa.png",
        "08_figura_2_spline_phq9_pa.pdf",
        "08_figura_2_spline_phq9_pa.svg",
        "08_figura_2_spline_phq9_pa_datos.csv",
        "08a_spline_nolinealidad_resumen.csv",
    ]:
        shutil.copy2(PACKAGE_DIR / filename, PUBLICATION_DIR / filename)

    note = """Archivos gráficos listos para revisión editorial.

Incluye:
- Figura 1 STROBE con texto corregido y conectores laterales.
- Figura 2 spline con título y subtítulo metodológicamente neutrales.
"""
    (PUBLICATION_DIR / "00_LEEME_GRAFICOS.txt").write_text(note, encoding="utf-8-sig")


def main() -> None:
    PACKAGE_DIR.mkdir(parents=True, exist_ok=True)

    strobe_table = _build_strobe_table()
    strobe_table.to_csv(PACKAGE_DIR / "02_tabla_strobe_flujo.csv", index=False, encoding="utf-8-sig")
    strobe_table.to_csv(PACKAGE_DIR / "07_figura_1_strobe_datos.csv", index=False, encoding="utf-8-sig")

    _plot_strobe_figure(
        strobe_table,
        PACKAGE_DIR / "07_figura_1_strobe.png",
        PACKAGE_DIR / "07_figura_1_strobe.pdf",
        PACKAGE_DIR / "07_figura_1_strobe.svg",
    )

    _copy_file(OUTPUT_DIR / "endes_hta_depresion_2019_2024.csv", "01_base_analitica_final.csv")
    _copy_file(OUTPUT_DIR / "endes_hta_depresion_2019_2024.parquet", "01_base_analitica_final.parquet")
    _copy_file(ANALYSIS_DIR / "tables" / "table1_weighted_summary.csv", "03_tabla_1_caracteristicas_basales.csv")
    _copy_file(ANALYSIS_DIR / "tables" / "table1_deff_summary.csv", "03a_tabla_1_deff_resumen.csv")
    _copy_file(ANALYSIS_DIR / "tables" / "table2_bivariate_pooled.csv", "04_tabla_2_bivariado.csv")
    _copy_file(ANALYSIS_DIR / "models" / "table3_main_models.csv", "05_tabla_3_modelos_principales.csv")
    _copy_file(ANALYSIS_DIR / "models" / "model_diagnostics_summary.csv", "05a_diagnosticos_modelos.csv")
    _copy_file(ANALYSIS_DIR / "models" / "table4_cascade_models.csv", "06_tabla_4_cascada_cuidado.csv")
    _copy_file(ANALYSIS_DIR / "figures" / "spline_curve_pooled.csv", "08_figura_2_spline_phq9_pa_datos.csv")
    _copy_file(ANALYSIS_DIR / "figures" / "spline_nonlinearity_summary.csv", "08a_spline_nolinealidad_resumen.csv")
    _copy_file(
        ANALYSIS_DIR / "models" / "interactions_and_sensitivity_models.csv",
        "09_analisis_sensibilidad_interacciones.csv",
    )
    _copy_file(ANALYSIS_DIR / "models" / "logistic_sensitivity_results.csv", "09a_sensibilidad_logistica.csv")
    _copy_file(ANALYSIS_DIR / "models" / "analysis_skipped_items.csv", "09b_items_omitidos.csv")

    _plot_spline_figure(
        PACKAGE_DIR / "08_figura_2_spline_phq9_pa.png",
        PACKAGE_DIR / "08_figura_2_spline_phq9_pa.pdf",
        PACKAGE_DIR / "08_figura_2_spline_phq9_pa.svg",
    )

    _write_manifest()
    _copy_publication_graphics()
    print(f"Paquete listo en: {PACKAGE_DIR}")


if __name__ == "__main__":
    main()
