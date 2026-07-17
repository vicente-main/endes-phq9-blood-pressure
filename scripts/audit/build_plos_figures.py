"""Genera las 4 figuras del envio CUMPLIENDO el estandar tecnico de PLOS ONE:
TIFF LZW, RGB (sin canal alfa), <= 2250 px de ancho y <= 2625 px de alto,
300 dpi, fuente Arial, SIN titulo/numero/leyenda incrustados (van en el .docx),
y nombres Fig1.tif (STROBE) / Fig2.tif (spline) / S1_Fig.tif (DAG) / S2_Fig.tif (forest).

Fuentes de datos: data/output_2025/ (run 2019-2025).
"""
from __future__ import annotations
import io
from pathlib import Path
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
matplotlib.rcParams["font.family"] = "Arial"
matplotlib.rcParams["axes.unicode_minus"] = False
import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch, FancyBboxPatch, Patch
from PIL import Image

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "ENVIO_PLOS_ONE_2019-2025"
FIG = ROOT / "data" / "output_2025" / "analysis" / "figures"
MODELS = ROOT / "data" / "output_2025" / "analysis" / "models"
QC = ROOT / "data" / "output_2025" / "qc"
MAX_W, MAX_H = 2250, 2625

COLOR = {
    "confounder": {"fill": "#D6EAD8", "edge": "#1B5E20", "arrow": "#4C8C4A"},
    "exposure":   {"fill": "#90CAF9", "edge": "#0D47A1", "arrow": "#0D47A1"},
    "outcome":    {"fill": "#F5B7B1", "edge": "#922B21", "arrow": "#922B21"},
    "mediator":   {"fill": "#FFE0B2", "edge": "#E65100", "arrow": "#E65100"},
    "unmeasured": {"fill": "#EDE1F3", "edge": "#6A1B9A", "arrow": "#6A1B9A"},
}
# Muted-but-visible green for the background confounder edges (fixes the previous
# alpha=0.15 "invisible back lines" problem while keeping them behind the nodes).
CONF_BG = "#7FB08A"


def plos_save(fig, name, render_dpi=600, final_dpi=300):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=render_dpi, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    buf.seek(0)
    im = Image.open(buf).convert("RGB")
    w, h = im.size
    scale = min(MAX_W / w, MAX_H / h)
    if scale < 1.0:
        im = im.resize((round(w * scale), round(h * scale)), Image.LANCZOS)
    im.save(OUT / name, format="TIFF", compression="tiff_lzw", dpi=(final_dpi, final_dpi))
    print(f"  {name:12s} {im.size[0]}x{im.size[1]}px RGB ({im.size[0]/final_dpi:.2f}x{im.size[1]/final_dpi:.2f} in)")


# ---------------------------------------------------------------- Fig 2 spline (principal)
def build_spline():
    d = pd.read_csv(FIG / "spline_curve_pooled.csv")
    d = d[d["summary_type"] == "spline_curve"].copy()
    d["x"] = d["label"].astype(float)
    d = d.sort_values("x")
    x, y = d["x"].to_numpy(), d["estimate"].to_numpy() * 100
    lo, hi = d["ci_low"].to_numpy() * 100, d["ci_high"].to_numpy() * 100
    fig, ax = plt.subplots(figsize=(6.5, 4.2))
    ax.fill_between(x, lo, hi, color="#A9C7E0", alpha=0.55, label="95% CI (Rubin)")
    ax.plot(x, y, color="#1F4E79", linewidth=2.4)
    for k in (0, 4, 9, 14):
        ax.axvline(k, color="#888888", linewidth=0.8, linestyle=":")
    ax.set_xlabel("PHQ-9 score", fontsize=11)
    # Auditoria 2026-07-16 (hallazgo #5): la curva es ahora la prevalencia marginal
    # estandarizada (g-computation ponderada), no una prediccion sobre un perfil modal.
    ax.set_ylabel("Marginal standardized prevalence\nof elevated BP (%)", fontsize=11)
    ax.tick_params(labelsize=10)
    ax.legend(loc="upper center", fontsize=10, frameon=True)
    ax.margins(x=0.01)
    fig.tight_layout()
    plos_save(fig, "Fig2.tif")   # spline = figura principal (auditoria 2026-07-16)


# ---------------------------------------------------------------- S3 forest
def build_forest():
    main_m = pd.read_csv(MODELS / "table3_main_models.csv")
    strat = pd.read_csv(MODELS / "altitude_stratified_models.csv")

    def pr(df, model):
        r = df[(df["model"] == model) & (df["term"] == "PHQ9_TOTAL")].iloc[0]
        return r["exp_estimate"], r["exp_ci_low"], r["exp_ci_high"]

    rows = [("Overall (Model 2)", pr(main_m, "model_2")),
            ("< 1,500 m", pr(strat, "model_2_estrato_cat3_lt1500")),
            ("1,500-2,499 m", pr(strat, "model_2_estrato_cat3_1500_2499")),
            ("≥ 2,500 m", pr(strat, "model_2_estrato_cat3_ge2500"))]
    labels = [r[0] for r in rows]
    ys = list(range(len(rows)))[::-1]
    fig, ax = plt.subplots(figsize=(6.6, 3.1))
    ax.axvline(1.0, color="#C0392B", linewidth=1.0, linestyle="--")
    for y, (_, (est, l, h)) in zip(ys, rows):
        ax.plot([l, h], [y, y], color="#1F4E79", linewidth=1.6)
        ax.plot(est, y, marker="s", markersize=8, color="#1F4E79")
    ax.set_yticks(ys); ax.set_yticklabels(labels, fontsize=10)
    ax.set_ylim(-0.6, len(rows) - 0.4)
    ax.set_xlabel("Prevalence ratio (PHQ-9, per point)", fontsize=11)
    ax.tick_params(axis="x", labelsize=10)
    ax.grid(axis="x", color="#EEEEEE", linewidth=0.6); ax.set_axisbelow(True)
    fig.tight_layout()
    plos_save(fig, "S2_Fig.tif")  # forest = suplementaria (renumerada)  # forest = suplementaria (renumerada)


# ---------------------------------------------------------------- Fig1 STROBE
def _sbox(ax, x, y, w, h, text, style, fs, weight="normal"):
    ax.add_patch(FancyBboxPatch((x - w / 2, y - h / 2), w, h,
                 boxstyle="round,pad=0.04,rounding_size=0.16", **style))
    ax.text(x, y, text, ha="center", va="center", fontsize=fs, weight=weight)


def _sarrow(ax, x1, y1, x2, y2):
    ax.add_patch(FancyArrowPatch((x1, y1), (x2, y2), arrowstyle="-|>",
                 mutation_scale=11, linewidth=0.9, color="#404040"))


def build_strobe():
    KEEP = {"facecolor": "#E7F0F8", "edgecolor": "#1F4E79", "linewidth": 1.3}
    EXCL = {"facecolor": "#F8E7E7", "edgecolor": "#9C2B2B", "linewidth": 0.9}
    d = pd.read_csv(QC / "filter_flow_by_year.csv")
    d = d[d["year"].astype(str) != "TOTAL"]; s = d.sum(numeric_only=True)
    g = lambda c: int(s[c])
    fmt = lambda n: f"{n:,}".replace(",", " ")
    seq = [("Complete health interview", "incomplete interview", "drop_qsresult_ok", "after_qsresult_ok"),
           ("Adults ≥ 18 years", "age < 18 years", "drop_adult_18_plus", "after_adult_18_plus"),
           ("Not pregnant (WRA only)", "pregnant (WRA)", "drop_not_pregnant_or_not_mef", "after_not_pregnant_or_not_mef"),
           ("Not early postpartum (WRA only)", "early postpartum (WRA)", "drop_not_puerperal_or_not_mef", "after_not_puerperal_or_not_mef")]
    structural, bp_excl, final = g("after_not_puerperal_or_not_mef"), g("drop_bp_available"), g("n_final")

    fig, ax = plt.subplots(figsize=(6.6, 8.5))
    ax.set_xlim(0, 10); ax.axis("off")
    FS, FSX = 9, 8
    spacing, main_x, main_w = 1.55, 3.9, 5.2
    main = [{"y": 12.3, "t": f"Initial sample, ENDES 2019-2025\nn = {fmt(g('n_start'))}", "st": KEEP, "h": 0.95, "b": True}]
    excl = []
    y = 12.3
    for kept, reason, cdrop, cafter in seq:
        y -= spacing
        main.append({"y": y, "t": f"{kept}\nn = {fmt(g(cafter))}", "st": KEEP, "h": 0.95, "b": False})
        excl.append({"y": y + spacing / 2, "t": f"Excluded: {fmt(g(cdrop))}\n({reason})"})
    y -= spacing
    main.append({"y": y, "t": f"Structural cohort\nn = {fmt(structural)}",
                 "st": {"facecolor": "#D0E4F2", "edgecolor": "#1F4E79", "linewidth": 1.5}, "h": 1.0, "b": True})
    y -= spacing
    excl.append({"y": y + spacing / 2, "t": f"Excluded: {fmt(bp_excl)}\n(blood pressure not measurable)"})
    main.append({"y": y, "t": f"Main cohort (analysis)\nn = {fmt(final)}",
                 "st": {"facecolor": "#BFD8B8", "edgecolor": "#3E6B1F", "linewidth": 1.7}, "h": 1.05, "b": True})
    ax.set_ylim(y - 0.9, 13.0)
    for i, b in enumerate(main):
        _sbox(ax, main_x, b["y"], main_w, b["h"], b["t"], dict(b["st"]), FS, "bold" if b["b"] else "normal")
        if i > 0:
            _sarrow(ax, main_x, main[i - 1]["y"] - main[i - 1]["h"] / 2, main_x, b["y"] + b["h"] / 2)
    ex_x, ex_w, ex_h = 8.15, 3.5, 0.9
    for e in excl:
        _sbox(ax, ex_x, e["y"], ex_w, ex_h, e["t"], EXCL, FSX)
        _sarrow(ax, main_x + main_w / 2, e["y"], ex_x - ex_w / 2, e["y"])
    fig.tight_layout()
    plos_save(fig, "Fig1.tif")


# ---------------------------------------------------------------- S2 DAG
def _dbox(ax, x, y, label, role, w, h, fs, weight="normal", dashed=False):
    st = COLOR[role]
    ax.add_patch(FancyBboxPatch((x - w / 2, y - h / 2), w, h,
                 boxstyle="round,pad=0.02,rounding_size=0.09",
                 facecolor=st["fill"], edgecolor=st["edge"], linewidth=0.9,
                 linestyle="--" if dashed else "-", zorder=3))
    ax.text(x, y, label, ha="center", va="center", fontsize=fs, weight=weight,
            color=st["edge"], zorder=4)
    return (x, y, w, h)


def _ep(node, d):
    x, y, w, h = node
    return {"bottom": (x, y - h / 2), "top": (x, y + h / 2),
            "left": (x - w / 2, y), "right": (x + w / 2, y)}[d]


def _darrow(ax, src, dst, role, cs="arc3,rad=0", lw=0.7, a=0.55, z=2, ls="-", color=None):
    ax.add_patch(FancyArrowPatch(src, dst, arrowstyle="-|>", mutation_scale=9,
                 linewidth=lw, color=color or COLOR[role]["arrow"], alpha=a,
                 connectionstyle=cs, zorder=z, linestyle=ls, capstyle="round"))


def _dcorr(ax, src, dst, rad=-0.10):
    ax.add_patch(FancyArrowPatch(src, dst, arrowstyle="<|-|>", mutation_scale=7,
                 linewidth=1.0, color=COLOR["confounder"]["edge"], alpha=0.6,
                 connectionstyle=f"arc3,rad={rad}", linestyle=(0, (4, 3)), zorder=4))


def build_dag():
    # Rediseño legibilidad (2026-07-14): aristas confusor->exp/desenlace curvadas con
    # jerarquia visual (cercana solida y visible; lejana mas fina y tenue pero VISIBLE,
    # bordeando el corredor central en vez de cruzarlo). Antes las lejanas iban a
    # alpha=0.15 (casi invisibles) y en recta cruzando toda la banda de mediadores.
    FN, FC, FM, FS_, FU, FL = 10.5, 8.0, 8.5, 9.5, 8.0, 8.5
    fig, ax = plt.subplots(figsize=(9.2, 7.2))
    ax.set_xlim(-0.6, 16.6); ax.set_ylim(-0.4, 12.2); ax.axis("off")
    ax.set_aspect("equal", adjustable="datalim")

    CY = 5.6
    phq = _dbox(ax, 1.35, CY, "Depressive\nsymptoms\n(PHQ-9)", "exposure", 2.15, 1.45, FN, "bold")
    pae = _dbox(ax, 15.0, CY, "Elevated\nblood\npressure", "outcome", 2.15, 1.45, FN, "bold")

    ax.text(8.2, 11.75, "Structural confounders — adjusted in Model 2 (main)",
            ha="center", va="center", fontsize=FS_, weight="bold", color=COLOR["confounder"]["edge"])
    confounders = [("EDAD", "Age", 2.4), ("SEXO", "Sex", 4.1), ("EDUC", "Education", 5.8),
                   ("AREA", "Area\n(urban/rural)", 7.5), ("ALTITUD", "Altitude", 9.2),
                   ("RIQUEZA", "Wealth\nquintile", 10.9), ("VPAR", "Intimate-\npartner\nviolence", 12.6),
                   ("ANIO", "Survey\nyear", 14.3)]
    cn = {}
    CBY = 10.05
    for k, lab, x in confounders:
        cn[k] = _dbox(ax, x, CBY, lab, "confounder", 1.62, 0.98, FC)
    _dcorr(ax, (cn["EDUC"][0], cn["EDUC"][1] + 0.49), (cn["RIQUEZA"][0], cn["RIQUEZA"][1] + 0.49), rad=-0.10)
    _dcorr(ax, (cn["EDAD"][0], cn["EDAD"][1] + 0.49), (cn["ANIO"][0], cn["ANIO"][1] + 0.49), rad=-0.05)
    ax.text(8.2, 11.28, "↔  correlations among confounders", ha="center", va="center",
            fontsize=FC - 0.5, style="italic", color=COLOR["confounder"]["edge"])

    ax.add_patch(FancyArrowPatch(_ep(phq, "right"), _ep(pae, "left"), arrowstyle="-|>",
                 mutation_scale=20, linewidth=2.6, color="#1F2937", alpha=0.97, zorder=7))
    px, py, pw, ph = phq; qx, qy, qw, qh = pae; mid = (px + qx) / 2
    for k, _l, x in confounders:
        src = _ep(cn[k], "bottom")
        near_phq = x < mid
        dst_phq = (px + pw / 2 - 0.15, py + ph / 2 - 0.12)
        dst_pae = (qx - qw / 2 + 0.15, qy + qh / 2 - 0.12)
        if near_phq:
            _darrow(ax, src, dst_phq, "confounder", cs="arc3,rad=0.16", lw=1.15, a=0.80, color=CONF_BG)
            _darrow(ax, src, dst_pae, "confounder", cs="arc3,rad=0.42", lw=0.8, a=0.45, color=CONF_BG)
        else:
            _darrow(ax, src, dst_pae, "confounder", cs="arc3,rad=-0.16", lw=1.15, a=0.80, color=CONF_BG)
            _darrow(ax, src, dst_phq, "confounder", cs="arc3,rad=-0.42", lw=0.8, a=0.45, color=CONF_BG)

    med_t = [("IMC", "BMI", 5.6), ("TABACO", "Tobacco use\n(last 30 d)", 8.2), ("DIETA", "Diet quality", 10.8)]
    med_b = [("CINTURA", "Waist\ncircumference", 5.6), ("ALC", "Problematic\nalcohol use", 8.2), ("DIAB", "Diabetes\ndiagnosis", 10.8)]
    mn = {}
    for k, lab, x in med_t:
        mn[k] = _dbox(ax, x, 7.35, lab, "mediator", 2.25, 0.92, FM)
    for k, lab, x in med_b:
        mn[k] = _dbox(ax, x, 3.85, lab, "mediator", 2.25, 0.92, FM)
    for k, _l, x in med_t:
        _darrow(ax, _ep(phq, "top"), _ep(mn[k], "left"), "mediator", cs="arc3,rad=-0.05", lw=1.15, a=0.85)
        _darrow(ax, _ep(mn[k], "right"), _ep(pae, "top"), "mediator", cs="arc3,rad=-0.05", lw=1.15, a=0.85)
    for k, _l, x in med_b:
        _darrow(ax, _ep(phq, "bottom"), _ep(mn[k], "left"), "mediator", cs="arc3,rad=0.05", lw=1.15, a=0.85)
        _darrow(ax, _ep(mn[k], "right"), _ep(pae, "bottom"), "mediator", cs="arc3,rad=0.05", lw=1.15, a=0.85)
    for sk, dk in [("EDAD", "IMC"), ("SEXO", "CINTURA"), ("RIQUEZA", "DIETA"), ("EDUC", "DIETA"), ("EDAD", "DIAB")]:
        _darrow(ax, _ep(cn[sk], "bottom"), _ep(mn[dk], "top"), "confounder", lw=0.7, a=0.4, ls=(0, (3, 2)), color=CONF_BG)

    ax.text(8.2, 2.75, "Potential mediators — added only in Model 3 (exploratory)",
            ha="center", va="center", fontsize=FS_, weight="bold", color=COLOR["mediator"]["edge"])
    u = _dbox(ax, 8.2, 1.35, "U: unmeasured factors\n(chronic stress, sleep, comorbidity,\nprior adherence, detection bias)",
              "unmeasured", 4.8, 1.05, FU, dashed=True)
    # Convex outward arcs (bow toward the bottom-outer corners) so they mirror the
    # mediator lines and stay clear of the "Potential mediators" label above U.
    _darrow(ax, _ep(u, "left"), _ep(phq, "bottom"), "unmeasured", cs="arc3,rad=-0.32", lw=1.1, a=0.7, ls="--")
    _darrow(ax, _ep(u, "right"), _ep(pae, "bottom"), "unmeasured", cs="arc3,rad=0.32", lw=1.1, a=0.7, ls="--")

    legend = [("Confounder (Model 2)", "confounder"), ("Exposure (PHQ-9)", "exposure"),
              ("Outcome (elevated BP)", "outcome"), ("Mediator (Model 3)", "mediator"),
              ("U: unmeasured (hypothetical)", "unmeasured")]
    handles = [Patch(facecolor=COLOR[r]["fill"], edgecolor=COLOR[r]["edge"],
                     linestyle="--" if r == "unmeasured" else "-", label=t) for t, r in legend]
    ax.legend(handles=handles, loc="upper center", bbox_to_anchor=(0.5, 0.02),
              ncol=3, fontsize=FL, frameon=False, handlelength=1.5,
              columnspacing=1.8, handletextpad=0.5)
    plos_save(fig, "S1_Fig.tif")  # DAG = suplementaria (renumerada)


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    print("Figuras PLOS (RGB, <=2250x2625, Arial, sin titulo):")
    build_spline()
    build_forest()
    build_strobe()
    build_dag()


if __name__ == "__main__":
    main()
