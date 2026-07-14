"""Figura 1 — Flujo STROBE detallado (cascada) en INGLES para ENDES 2019-2025.

Replica el diseno de la cascada detallada de build_phase_c_figures.py (Figura_1_STROBE_v2)
pero: (a) en ingles, (b) con los conteos 2019-2025 tomados del pipeline
(data/output_2025/qc/filter_flow_by_year.csv), y (c) exporta a ENVIO como
Fig1_STROBE.{pdf,png,tif} (TIFF 600 dpi para PLOS ONE).

Sustituye a la version simplificada de 3 cajas que estaba en el paquete.
"""
from __future__ import annotations
from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch

ROOT = Path(__file__).resolve().parents[2]
FLOW = ROOT / "data" / "output_2025" / "qc" / "filter_flow_by_year.csv"
OUT = ROOT / "ENVIO_PLOS_ONE_2019-2025"

KEEP = {"facecolor": "#E7F0F8", "edgecolor": "#1F4E79", "linewidth": 1.4}
EXCL = {"facecolor": "#F8E7E7", "edgecolor": "#9C2B2B", "linewidth": 1.0}


def _box(ax, *, x, y, w, h, text, style, fontsize=10, weight="normal"):
    ax.add_patch(FancyBboxPatch((x - w / 2, y - h / 2), w, h,
                                boxstyle="round,pad=0.04,rounding_size=0.18", **style))
    ax.text(x, y, text, ha="center", va="center", fontsize=fontsize, weight=weight, wrap=True)


def _arrow(ax, x1, y1, x2, y2):
    ax.add_patch(FancyArrowPatch((x1, y1), (x2, y2), arrowstyle="-|>",
                                 mutation_scale=14, linewidth=1.0, color="#404040"))


def main():
    d = pd.read_csv(FLOW)
    d = d[d["year"].astype(str) != "TOTAL"]
    s = d.sum(numeric_only=True)
    g = lambda c: int(s[c])

    initial = g("n_start")
    # secuencia de pasos con exclusion (etiqueta_conservada, razon_exclusion, col_drop, col_after)
    seq = [
        ("Complete health interview", "incomplete interview", "drop_qsresult_ok", "after_qsresult_ok"),
        ("Adults ≥ 18 years", "age < 18 years", "drop_adult_18_plus", "after_adult_18_plus"),
        ("Not pregnant (WRA only)", "pregnant (WRA)", "drop_not_pregnant_or_not_mef", "after_not_pregnant_or_not_mef"),
        ("Not early postpartum (WRA only)", "early postpartum (WRA)", "drop_not_puerperal_or_not_mef", "after_not_puerperal_or_not_mef"),
    ]
    structural = g("after_not_puerperal_or_not_mef")  # = before BP = cohorte estructural (PHQ drop=0)
    bp_excl = g("drop_bp_available")
    final = g("n_final")

    # nota: 2020 y sin-2020
    row2020 = d[d["year"].astype(str) == "2020"]
    bp_2020 = int(row2020["drop_bp_available"].iloc[0]) if not row2020.empty else 0
    no2020 = final - int(row2020["n_final"].iloc[0]) if not row2020.empty else final

    fig, ax = plt.subplots(figsize=(11.5, 13.5))
    ax.set_xlim(0, 10); ax.set_ylim(0, 14); ax.axis("off")
    ax.text(5, 13.6, "Figure 1. STROBE flow diagram — sample selection, ENDES 2019-2025",
            ha="center", va="top", fontsize=13, weight="bold")

    def fmt(n): return f"{n:,}".replace(",", " ")

    main_boxes = [{"y": 12.3, "label": f"Initial sample, ENDES 2019-2025\nn = {fmt(initial)}", "style": KEEP, "h": 0.9}]
    excl_boxes = []
    spacing = 1.55
    main_y = 12.3
    for kept_label, reason, cdrop, cafter in seq:
        main_y -= spacing
        main_boxes.append({"y": main_y, "label": f"{kept_label}\nn = {fmt(g(cafter))}", "style": KEEP, "h": 0.9})
        excl_boxes.append({"y": main_y + spacing / 2, "label": f"Excluded: {fmt(g(cdrop))}\n({reason})"})

    main_y -= spacing
    main_boxes.append({"y": main_y, "label": f"Structural cohort\nn = {fmt(structural)}",
                       "style": {"facecolor": "#D0E4F2", "edgecolor": "#1F4E79", "linewidth": 1.6}, "h": 0.95})
    main_y -= spacing
    excl_boxes.append({"y": main_y + spacing / 2, "label": f"Excluded: {fmt(bp_excl)}\n(blood pressure not measurable)"})
    main_boxes.append({"y": main_y, "label": f"Main cohort (analysis)\nn = {fmt(final)}",
                       "style": {"facecolor": "#BFD8B8", "edgecolor": "#3E6B1F", "linewidth": 1.8}, "h": 1.0})

    main_x, main_w = 4.0, 5.0
    for i, b in enumerate(main_boxes):
        _box(ax, x=main_x, y=b["y"], w=main_w, h=b["h"], text=b["label"], style=dict(b["style"]),
             fontsize=10, weight="bold" if i in (0, len(main_boxes) - 2, len(main_boxes) - 1) else "normal")
        if i > 0:
            _arrow(ax, main_x, main_boxes[i - 1]["y"] - main_boxes[i - 1]["h"] / 2, main_x, b["y"] + b["h"] / 2)

    excl_x, excl_w, excl_h = 8.0, 3.4, 0.85
    for e in excl_boxes:
        _box(ax, x=excl_x, y=e["y"], w=excl_w, h=excl_h, text=e["label"], style=EXCL, fontsize=8.5)
        _arrow(ax, main_x + main_w / 2, e["y"], excl_x - excl_w / 2, e["y"])

    notes = (
        "Notes: (a) Exclusions for pregnancy (V454=1) and early postpartum (V222<2) were applied only to "
        "women of reproductive age (15-49 y, WRA); men and women outside this age range were never excluded "
        f"by these criteria. (b) In 2020, {fmt(bp_2020)} records were excluded for non-measurable blood "
        "pressure due to COVID-19 operational restrictions; the sensitivity analysis excluding 2020 includes "
        f"{fmt(no2020)} participants. Per-year counts in S1 Table."
    )
    ax.text(0.3, 0.4, notes, ha="left", va="bottom", fontsize=8.5, wrap=True, color="#404040")

    plt.tight_layout()
    OUT.mkdir(parents=True, exist_ok=True)
    png = OUT / "Fig1_STROBE.png"; pdf = OUT / "Fig1_STROBE.pdf"; tif = OUT / "Fig1_STROBE.tif"
    fig.savefig(pdf, format="pdf", bbox_inches="tight")
    fig.savefig(png, format="png", bbox_inches="tight", dpi=300)
    fig.savefig(tif, format="tiff", bbox_inches="tight", dpi=600, pil_kwargs={"compression": "tiff_lzw"})
    plt.close(fig)
    print(f"Initial={initial:,} structural={structural:,} final={final:,} bp_excl={bp_excl:,}")
    print(f"2020 bp_excl={bp_2020:,}  sin-2020 n={no2020:,}")
    for f in (png, pdf, tif):
        print("  ->", f.relative_to(ROOT))


if __name__ == "__main__":
    main()
