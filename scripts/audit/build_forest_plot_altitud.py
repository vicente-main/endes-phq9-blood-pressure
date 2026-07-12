"""Figura S3 — Forest plot: razon de prevalencia de PHQ-9 (por punto) sobre PA
elevada, global (Modelo 2 ajustado por altitud) y por estrato de altitud.

Visualiza la heterogeneidad por piso altitudinal (interaccion pooled p = 0,017).
Lee de los outputs canonicos del pipeline (no recalcula).
"""
from __future__ import annotations

from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd

ROOT = Path(__file__).resolve().parents[2]
MODELS = ROOT / "data" / "output" / "analysis" / "models"
OUTDIR = ROOT / "data" / "output" / "Post_Auditoria" / "Suplementario"

C_PT = "#1F4E79"
C_REF = "#B0B0B0"


def _pr(df, model, term="PHQ9_TOTAL"):
    r = df[(df["model"] == model) & (df["term"] == term)].iloc[0]
    return float(r["exp_estimate"]), float(r["exp_ci_low"]), float(r["exp_ci_high"]), float(r["p_value"]), float(r["mean_n_obs"])


def main() -> None:
    main_m = pd.read_csv(MODELS / "table3_main_models.csv")
    strat = pd.read_csv(MODELS / "altitude_stratified_models.csv")

    # Filas (de arriba hacia abajo): overall + 3 categorias de altitud.
    rows = []
    est, lo, hi, p, n = _pr(main_m, "model_2")
    rows.append(("Global (Modelo 2, ajustado por altitud)", est, lo, hi, p, n, True))
    for label, model in [
        ("Altitud < 1 500 m", "model_2_estrato_cat3_lt1500"),
        ("Altitud 1 500-2 499 m", "model_2_estrato_cat3_1500_2499"),
        ("Altitud ≥ 2 500 m", "model_2_estrato_cat3_ge2500"),
    ]:
        est, lo, hi, p, n = _pr(strat, model)
        rows.append((label, est, lo, hi, p, n, False))

    rows = rows[::-1]  # forest: primera fila arriba
    y = list(range(len(rows)))

    fig, ax = plt.subplots(figsize=(9.5, 4.2))
    ax.axvline(1.0, color=C_REF, linestyle="--", linewidth=1.0, zorder=1)

    for yi, (label, est, lo, hi, p, n, is_overall) in zip(y, rows):
        color = "#111111" if is_overall else C_PT
        ax.plot([lo, hi], [yi, yi], color=color, linewidth=1.8, zorder=2)
        ax.plot([lo, lo], [yi - 0.12, yi + 0.12], color=color, linewidth=1.8)
        ax.plot([hi, hi], [yi - 0.12, yi + 0.12], color=color, linewidth=1.8)
        marker = "D" if is_overall else "s"
        ax.scatter([est], [yi], s=70 if is_overall else 55, color=color, marker=marker, zorder=3)
        ptxt = "< 0,001" if p < 0.001 else f"{p:.3f}".replace(".", ",")
        est_txt = f"{est:.3f} ({lo:.3f}–{hi:.3f})".replace(".", ",")
        ax.text(1.022, yi, f"{est_txt}   p = {ptxt}", va="center", ha="left", fontsize=8.6,
                fontweight="bold" if is_overall else "normal")

    ax.set_yticks(y)
    ax.set_yticklabels(
        [f"{label}\n(n = {int(n):,})".replace(",", " ") for (label, *_rest, n, _o) in
         [(r[0], r[5], r[6]) for r in rows]],
        fontsize=8.8,
    )
    ax.set_xlim(0.965, 1.075)
    ax.set_xlabel("Razón de prevalencia de PHQ-9 por punto (IC 95 %)  —  PA elevada", fontsize=9.5)
    ax.set_title(
        "Figura S3. Asociación PHQ-9 → presión arterial elevada por estrato de altitud\n"
        "(ENDES 2019-2024; interacción PHQ-9 × altitud, p = 0,017)",
        fontsize=11, fontweight="bold",
    )
    ax.spines[["top", "right"]].set_visible(False)
    ax.tick_params(axis="y", length=0)
    fig.text(0.012, -0.02,
             "Cada estimado es el PR de PHQ-9 (por punto) dentro del estrato, ajustado por edad, sexo, educación, área, riqueza, "
             "violencia de pareja y año (Modelo 2 sin término de altitud dentro del estrato). svydesign + pooling de Rubin (20 imputaciones MICE). "
             "PR < 1 a la izquierda de la línea = asociación inversa. La señal inversa se concentra a 1 500-2 499 m y es nula ≥ 2 500 m.",
             fontsize=6.8, ha="left", va="top", wrap=True)

    OUTDIR.mkdir(parents=True, exist_ok=True)
    for ext in ("svg", "png", "pdf"):
        fig.savefig(OUTDIR / f"Figura_S3_forest_altitud.{ext}", bbox_inches="tight", dpi=200)
        print(f"  -> Suplementario/Figura_S3_forest_altitud.{ext}")
    plt.close(fig)


if __name__ == "__main__":
    main()
