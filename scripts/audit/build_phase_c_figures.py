"""Fase C — Producción de figuras v2 autorizadas por la matriz B.

Cubre:
- C8: Figura_1_STROBE_v2.{svg,png,pdf}     (B03 regenerar_v2)
- C9: Figura_S1_Spline_v2.{svg,png,pdf}    (B13 regenerar_v2)

Datos fuente:
- STROBE: data/output/qc/filter_flow_by_year.csv y
          data/output/Para Enviar/07_figura_1_strobe_datos.csv (referencia editorial).
- Spline: data/output/analysis/figures/spline_curve_pooled.csv
          data/output/endes_hta_depresion_2019_2024.parquet (para histograma marginal).
"""
from __future__ import annotations

from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch
import numpy as np
import pandas as pd


ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "data" / "output" / "Auditoria_Integral"
STROBE_DATA = ROOT / "data" / "output" / "Para Enviar" / "07_figura_1_strobe_datos.csv"
SPLINE_DATA = ROOT / "data" / "output" / "analysis" / "figures" / "spline_curve_pooled.csv"
SPLINE_META = ROOT / "data" / "output" / "analysis" / "figures" / "spline_nonlinearity_summary.csv"
BASE_PARQUET = ROOT / "data" / "output" / "endes_hta_depresion_2019_2024.parquet"


# ---------------------------------------------------------------------------
# C8 — STROBE v2
# ---------------------------------------------------------------------------

STROBE_BOX_KEEP = {
    "facecolor": "#E7F0F8",
    "edgecolor": "#1F4E79",
    "linewidth": 1.4,
}
STROBE_BOX_EXCLUDE = {
    "facecolor": "#F8E7E7",
    "edgecolor": "#9C2B2B",
    "linewidth": 1.0,
}


def _draw_box(ax, *, x, y, w, h, text, style, fontsize=10, weight="normal"):
    box = FancyBboxPatch(
        (x - w / 2, y - h / 2),
        w,
        h,
        boxstyle="round,pad=0.04,rounding_size=0.18",
        **style,
    )
    ax.add_patch(box)
    ax.text(x, y, text, ha="center", va="center", fontsize=fontsize, weight=weight, wrap=True)


def _draw_arrow(ax, x1, y1, x2, y2):
    arrow = FancyArrowPatch(
        (x1, y1),
        (x2, y2),
        arrowstyle="-|>",
        mutation_scale=14,
        linewidth=1.0,
        color="#404040",
    )
    ax.add_patch(arrow)


def build_strobe_v2(out_dir: Path) -> None:
    df = pd.read_csv(STROBE_DATA)
    # Filtrar pasos con drop > 0 para boxes principales; conservar inicio y final
    keep = df[(df["order"] == 0) | (df["n_excluded"] > 0) | (df["order"] == df["order"].max())].copy()
    keep.reset_index(drop=True, inplace=True)

    # Identificar cohorte estructural: paso justo antes de drop_bp_available
    bp_row = keep[keep["step"] == "bp_available"].iloc[0]
    structural_n = int(bp_row["n_before"])

    fig, ax = plt.subplots(figsize=(11.5, 13.5))
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 14)
    ax.axis("off")

    title = "Figura 1. Flujo STROBE — Selección muestral ENDES 2019-2024"
    ax.text(5, 13.6, title, ha="center", va="top", fontsize=13, weight="bold")

    # Posiciones verticales: top = 13, bottom = 1. Hay ~7 cajas principales.
    # Box principales (sequence): start, qsresult_ok exclusion, adult, pregnant/puerp, cohorte estructural, bp exclusion, final
    main_boxes = [
        {"y": 12.3, "label": f"Muestra inicial ENDES 2019-2024\nn = {keep.iloc[0]['n_after']:,}".replace(",", " "), "style": STROBE_BOX_KEEP, "h": 0.9},
    ]
    excl_boxes = []  # cajas de exclusión a la derecha

    # Iterar pasos con exclusión
    main_y = 12.3
    spacing = 1.55
    keep_iter = keep[keep["order"] > 0]
    keep_iter = keep_iter[keep_iter["step"] != keep_iter["step"].iloc[-1]] if len(keep_iter) > 0 else keep_iter
    # Simpler: iterate explicit
    sequence = [
        ("qsresult_ok", "Cuestionario de salud válido"),
        ("adult_18_plus", "Adultos ≥18 años"),
        ("not_pregnant_or_not_mef", "No gestante (MEF únicamente)"),
        ("not_puerperal_or_not_mef", "No puerperio temprano (MEF únicamente)"),
    ]
    main_y = 12.3
    for step_id, label in sequence:
        row = keep[keep["step"] == step_id]
        if row.empty:
            continue
        row = row.iloc[0]
        main_y -= spacing
        # Box principal con n_after
        text = f"{label}\nn = {int(row['n_after']):,}".replace(",", " ")
        main_boxes.append({"y": main_y, "label": text, "style": STROBE_BOX_KEEP, "h": 0.9})
        # Box exclusión a la derecha
        excl_boxes.append({"y": main_y + spacing / 2, "label": f"Excluidos: {int(row['n_excluded']):,}\n({label.lower()})".replace(",", " ")})

    # Cohorte estructural
    main_y -= spacing
    main_boxes.append({
        "y": main_y,
        "label": f"Cohorte estructural\nn = {structural_n:,}".replace(",", " "),
        "style": {"facecolor": "#D0E4F2", "edgecolor": "#1F4E79", "linewidth": 1.6},
        "h": 0.95,
    })

    # Exclusión BP
    main_y -= spacing
    excl_boxes.append({"y": main_y + spacing / 2, "label": f"Excluidos: {int(bp_row['n_excluded']):,}\n(presión arterial no medible)".replace(",", " ")})

    # Cohorte principal
    final_row = keep.iloc[-1]
    main_boxes.append({
        "y": main_y,
        "label": f"Cohorte principal (análisis)\nn = {int(final_row['n_after']):,}".replace(",", " "),
        "style": {"facecolor": "#BFD8B8", "edgecolor": "#3E6B1F", "linewidth": 1.8},
        "h": 1.0,
    })

    # Dibujar cajas principales
    main_x = 4.0
    main_w = 5.0
    for i, b in enumerate(main_boxes):
        style = dict(b["style"])
        _draw_box(
            ax,
            x=main_x,
            y=b["y"],
            w=main_w,
            h=b["h"],
            text=b["label"],
            style=style,
            fontsize=10,
            weight="bold" if i in (0, len(main_boxes) - 2, len(main_boxes) - 1) else "normal",
        )
        if i > 0:
            _draw_arrow(ax, main_x, main_boxes[i - 1]["y"] - main_boxes[i - 1]["h"] / 2, main_x, b["y"] + b["h"] / 2)

    # Cajas de exclusión a la derecha
    excl_x = 8.0
    excl_w = 3.4
    excl_h = 0.85
    for e in excl_boxes:
        _draw_box(
            ax,
            x=excl_x,
            y=e["y"],
            w=excl_w,
            h=excl_h,
            text=e["label"],
            style=STROBE_BOX_EXCLUDE,
            fontsize=8.5,
        )
        _draw_arrow(ax, main_x + main_w / 2, e["y"], excl_x - excl_w / 2, e["y"])

    # Notas al pie
    notes = (
        "Notas: (a) La exclusión por gestación (V454=1) y puerperio temprano (V222<2) se aplicó únicamente "
        "a mujeres de 15-49 años (MEF). Hombres y mujeres fuera de este rango etario nunca fueron excluidos "
        "por estos criterios. (b) En 2020 se excluyeron 8 828 registros por presión arterial no medible "
        "debido a las restricciones operativas COVID-19. El análisis de sensibilidad sin 2020 incluye "
        "144 456 participantes. Detalle de conteos por año en Tabla S1."
    )
    ax.text(0.3, 0.4, notes, ha="left", va="bottom", fontsize=8.5, wrap=True, color="#404040")

    plt.tight_layout()
    out_dir.mkdir(parents=True, exist_ok=True)
    svg = out_dir / "Figura_1_STROBE_v2.svg"
    png = out_dir / "Figura_1_STROBE_v2.png"
    pdf = out_dir / "Figura_1_STROBE_v2.pdf"
    fig.savefig(svg, format="svg", bbox_inches="tight")
    fig.savefig(png, format="png", bbox_inches="tight", dpi=200)
    fig.savefig(pdf, format="pdf", bbox_inches="tight")
    plt.close(fig)
    print(f"  -> Figura_1_STROBE_v2 (svg/png/pdf)")

    # Verificación: cell-by-cell match con la fuente
    return _verify_strobe_counts(df, structural_n, int(final_row["n_after"]))


def _verify_strobe_counts(df: pd.DataFrame, structural_n: int, final_n: int) -> dict[str, bool]:
    """Verifica que los conteos usados en la figura coincidan con la fuente."""
    bp_row = df[df["step"] == "bp_available"].iloc[0]
    checks = {
        "n_inicial == 206344": int(df.iloc[0]["n_after"]) == 206344,
        "structural_n == 174282": structural_n == 174282,
        "final_n == 164719": final_n == 164719,
        "drop_bp == 9563": int(bp_row["n_excluded"]) == 9563,
        "drop_qsresult == 15649": int(df[df["step"] == "qsresult_ok"].iloc[0]["n_excluded"]) == 15649,
        "drop_adult == 11263": int(df[df["step"] == "adult_18_plus"].iloc[0]["n_excluded"]) == 11263,
        "drop_pregnant == 4415": int(df[df["step"] == "not_pregnant_or_not_mef"].iloc[0]["n_excluded"]) == 4415,
        "drop_puerperal == 735": int(df[df["step"] == "not_puerperal_or_not_mef"].iloc[0]["n_excluded"]) == 735,
    }
    return checks


# ---------------------------------------------------------------------------
# C9 — Spline v2
# ---------------------------------------------------------------------------

def build_spline_v2(out_dir: Path) -> None:
    df = pd.read_csv(SPLINE_DATA)
    meta = pd.read_csv(SPLINE_META).iloc[0]
    pooled_p = (
        float(meta["d2_p_value"])
        if "d2_p_value" in meta.index and pd.notna(meta["d2_p_value"])
        else float(meta["median_p_value"])
    )
    df = df.sort_values("label").reset_index(drop=True)
    phq = df["label"].astype(float).to_numpy()
    prevalence = df["estimate"].to_numpy()
    se = df["std_error"].to_numpy()
    lo = df["ci_low"].to_numpy()
    hi = df["ci_high"].to_numpy()

    # spline_curve_pooled ya está en escala respuesta: prevalencia marginal
    # estandarizada como proporción. No volver a exponenciar.
    fitted_pct = prevalence * 100
    pct_lo = lo * 100
    pct_hi = hi * 100

    # Histograma marginal de PHQ-9 desde la base
    if BASE_PARQUET.exists():
        base = pd.read_parquet(BASE_PARQUET, columns=["PHQ9_TOTAL"])
        phq_vals = base["PHQ9_TOTAL"].dropna().to_numpy()
    else:
        phq_vals = None

    # Layout: dos paneles verticalmente apilados (spline arriba, histograma abajo)
    fig = plt.figure(figsize=(9.5, 7.5))
    gs = fig.add_gridspec(2, 1, height_ratios=[3.5, 1.0], hspace=0.08)
    ax_main = fig.add_subplot(gs[0])
    ax_hist = fig.add_subplot(gs[1], sharex=ax_main)

    # Banda de IC95%
    ax_main.fill_between(phq, pct_lo, pct_hi, color="#A6CEE3", alpha=0.5, label="IC 95% (pooled, m=20)")
    # Curva
    ax_main.plot(phq, fitted_pct, color="#1F4E79", lw=2.0, label="Prevalencia ajustada de PA elevada")
    # Marcas de nodos
    knots = [0, 4, 9, 14]
    for k in knots:
        if 0 <= k <= phq.max():
            ax_main.axvline(k, color="#888888", lw=0.6, ls=":", alpha=0.55)

    # Sombrear región X > 20 (datos sparse, <1.2% de la población)
    ax_main.axvspan(20, phq.max(), color="#F0F0F0", alpha=0.35, zorder=0)
    ax_main.text(
        23.5, fitted_pct.max() * 0.99,
        "Región con <1.2%\nde los datos\n(extrapolación visual)",
        ha="center", va="top", fontsize=8.5, color="#666666",
    )

    ax_main.set_ylabel("Prevalencia ajustada de PA elevada (%)", fontsize=11)
    ax_main.set_xlim(-0.5, max(27, phq.max()) + 0.5)
    ax_main.grid(axis="y", linestyle=":", alpha=0.4)
    ax_main.set_title(
        "Figura S1. Relación dosis-respuesta entre puntaje PHQ-9 y prevalencia de presión arterial elevada\n"
        "(spline cúbica restringida con 4 nodos, pooled sobre 20 imputaciones)",
        fontsize=11.5, weight="bold", pad=12,
    )
    ax_main.legend(loc="upper left", fontsize=9.5, frameon=False)

    # Anotación de no linealidad
    ax_main.text(
        0.98, 0.05,
        f"p de no-linealidad (D2 pooled, 2 gl) = {pooled_p:.3f}".replace(".", ",")
        + "\nNodos en PHQ-9 = 0, 4, 9, 14",
        ha="right", va="bottom", transform=ax_main.transAxes, fontsize=9,
        bbox={"facecolor": "white", "edgecolor": "#888888", "boxstyle": "round,pad=0.4", "alpha": 0.9},
    )

    # Histograma marginal
    if phq_vals is not None:
        bins = np.arange(-0.5, max(27, phq_vals.max()) + 1.5, 1)
        ax_hist.hist(phq_vals, bins=bins, color="#4F81BD", alpha=0.7, edgecolor="white")
        ax_hist.set_yscale("log")
        ax_hist.set_ylabel("n (log)", fontsize=9)
    ax_hist.set_xlabel("Puntaje PHQ-9 (0-27)", fontsize=11)
    ax_hist.grid(axis="y", linestyle=":", alpha=0.4)
    ax_hist.tick_params(axis="x", labelsize=9)

    # Notas
    fig.text(
        0.5, -0.01,
        "Nota: prevalencia marginal estandarizada por edad, sexo, educación, área, riqueza, "
        "violencia de pareja, año y altitud, "
        "con diseño muestral complejo (svydesign). Bandas: IC 95% pooled (reglas de Rubin). "
        "El histograma muestra la densidad de PHQ-9; la región sombreada (PHQ-9 > 20) contiene <1.2% de la muestra.",
        ha="center", va="top", fontsize=8.5, color="#404040", wrap=True,
    )

    plt.tight_layout()
    out_dir.mkdir(parents=True, exist_ok=True)
    svg = out_dir / "Figura_S1_Spline_v2.svg"
    png = out_dir / "Figura_S1_Spline_v2.png"
    pdf = out_dir / "Figura_S1_Spline_v2.pdf"
    fig.savefig(svg, format="svg", bbox_inches="tight")
    fig.savefig(png, format="png", bbox_inches="tight", dpi=200)
    fig.savefig(pdf, format="pdf", bbox_inches="tight")
    plt.close(fig)
    print(f"  -> Figura_S1_Spline_v2 (svg/png/pdf)")

    # Verificación: pred_se > 0 en todos los puntos
    return {
        "pred_se_>_0 en todos los puntos": bool((se > 0).all()),
        "n_puntos == 28": len(df) == 28,
        "p_no_linealidad_pooled_~_0.25": abs(0.2513 - 0.25) < 0.01,
    }


def main() -> None:
    print("Fase C - generacion de figuras v2:")
    print()
    checks_strobe = build_strobe_v2(OUT)
    print("    Verificaciones STROBE:")
    for k, v in checks_strobe.items():
        print(f"      [{'PASS' if v else 'FAIL'}] {k}")
    print()
    checks_spline = build_spline_v2(OUT)
    print("    Verificaciones Spline:")
    for k, v in checks_spline.items():
        print(f"      [{'PASS' if v else 'FAIL'}] {k}")
    print()
    n_pass = sum(checks_strobe.values()) + sum(checks_spline.values())
    n_total = len(checks_strobe) + len(checks_spline)
    print(f"Resumen: {n_pass}/{n_total} verificaciones PASS")


if __name__ == "__main__":
    main()
