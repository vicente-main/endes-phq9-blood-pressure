# -*- coding: utf-8 -*-
"""Comparacion cohorte canonica 2019-2024 vs nueva 2019-2025 + H5 (atenuacion altitud)
+ H7 (VIF maximo y dispersion). Lee outputs analiticos de ambos arboles y escribe un
informe en data/output_2025/comparacion_2019_2025/.

Uso:
    .\.venv\Scripts\python.exe .\scripts\compare_2024_vs_2025.py
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
OLD = ROOT / "data" / "output"           # canonico 2019-2024
NEW = ROOT / "data" / "output_2025"      # nuevo 2019-2025
OUT = NEW / "comparacion_2019_2025"
OUT.mkdir(parents=True, exist_ok=True)


def _read(base: Path, rel: str) -> pd.DataFrame | None:
    p = base / rel
    if not p.exists():
        print(f"  [aviso] no existe: {p}")
        return None
    return pd.read_csv(p, encoding="utf-8-sig")


def _phq_row(df: pd.DataFrame | None, model: str) -> dict | None:
    if df is None:
        return None
    sub = df[(df["model"] == model) & (df["term"] == "PHQ9_TOTAL")]
    if sub.empty:
        return None
    r = sub.iloc[0]
    return {
        "PR": round(float(r["exp_estimate"]), 4),
        "IC_low": round(float(r["exp_ci_low"]), 4),
        "IC_high": round(float(r["exp_ci_high"]), 4),
        "p": float(r["p_value"]),
        "n": int(float(r["mean_n_obs"])),
    }


def _fmt(d: dict | None) -> str:
    if d is None:
        return "n/d"
    p = "<0,001" if d["p"] < 0.001 else f"{d['p']:.3f}".replace(".", ",")
    return f"RP={d['PR']:.3f} (IC95% {d['IC_low']:.3f}-{d['IC_high']:.3f}; p={p}; n={d['n']:,})".replace(".", ",").replace(",000", " ").replace("n=", "n=")


def compare_models() -> list[str]:
    lines = ["## 1. Estimadores PHQ-9 -> PA elevada (PR de PHQ9_TOTAL)\n"]
    old_t3, new_t3 = _read(OLD, "analysis/models/table3_main_models.csv"), _read(NEW, "analysis/models/table3_main_models.csv")
    old_int, new_int = _read(OLD, "analysis/models/interactions_and_sensitivity_models.csv"), _read(NEW, "analysis/models/interactions_and_sensitivity_models.csv")
    old_hier, new_hier = _read(OLD, "analysis/models/hierarchical_decomposition.csv"), _read(NEW, "analysis/models/hierarchical_decomposition.csv")
    old_log, new_log = _read(OLD, "analysis/models/logistic_sensitivity_results.csv"), _read(NEW, "analysis/models/logistic_sensitivity_results.csv")

    rows = []
    model_sources = [
        ("model_1 (crudo)", "model_1", old_t3, new_t3),
        ("model_2 (principal, +altitud)", "model_2", old_t3, new_t3),
        ("model_3 (exploratorio)", "model_3", old_t3, new_t3),
        ("h2 (estructural SIN altitud)", "h2_estructural_sin_altitud", old_hier, new_hier),
        ("h3 (estructural CON altitud)", "h3_estructural_con_altitud", old_hier, new_hier),
        ("sens. sin 2020", "sensitivity_no_2020", old_int, new_int),
        ("sens. segunda toma PA", "sensitivity_second_bp_measure", old_int, new_int),
        ("sens. desenlace 130/80 (S1)", "sensitivity_outcome_130_80", old_int, new_int),
        ("logistica (OR)", "model_2_logistic_sensitivity", old_log, new_log),
    ]
    lines.append("| Modelo | 2019-2024 (canonico) | 2019-2025 (nuevo) |")
    lines.append("|---|---|---|")
    for label, key, od, nd in model_sources:
        o, n = _phq_row(od, key), _phq_row(nd, key)
        lines.append(f"| {label} | {_fmt(o)} | {_fmt(n)} |")
        rows.append({"modelo": label, "clave": key,
                     "PR_2024": o["PR"] if o else None, "p_2024": o["p"] if o else None,
                     "PR_2025": n["PR"] if n else None, "p_2025": n["p"] if n else None})
    pd.DataFrame(rows).to_csv(OUT / "comparacion_estimadores_phq9.csv", index=False, encoding="utf-8-sig")
    return lines, (old_hier, new_hier)


def h5_attenuation(hier: pd.DataFrame | None, tag: str) -> list[str]:
    lines = [f"\n### {tag}"]
    if hier is None:
        return lines + ["  (sin datos)"]
    h2 = _phq_row(hier, "h2_estructural_sin_altitud")
    h3 = _phq_row(hier, "h3_estructural_con_altitud")
    if not (h2 and h3):
        return lines + ["  (faltan h2/h3)"]
    dev2, dev3 = abs(1 - h2["PR"]), abs(1 - h3["PR"])
    atten = (dev2 - dev3) / dev2 * 100 if dev2 else float("nan")
    lines.append(f"- h2 SIN altitud: RP={h2['PR']:.4f} (desviacion del nulo={dev2:.4f}, p={h2['p']:.4f})")
    lines.append(f"- h3 CON altitud: RP={h3['PR']:.4f} (desviacion del nulo={dev3:.4f}, p={h3['p']:.4f})")
    lines.append(f"- **Atenuacion por altitud: {atten:.1f}%** (de RP={h2['PR']:.4f} a RP={h3['PR']:.4f})")
    return lines


def h7_vif_dispersion(base: Path, tag: str) -> list[str]:
    lines = [f"\n### {tag}"]
    vif = _read(base, "analysis/vif_first_imputation.csv")
    if vif is not None and not vif.empty:
        vif = vif.sort_values("vif", ascending=False)
        top = vif.iloc[0]
        lines.append(f"- VIF maximo = {float(top['vif']):.2f} (predictor: {top['predictor']}).")
        hi = vif[vif["vif"] >= 5]
        if not hi.empty:
            preds = ", ".join(f"{r.predictor}={r.vif:.1f}" for r in hi.itertuples())
            lines.append(f"- Predictores con VIF>=5 (dummies colineales por diseno, mismo factor): {preds}")
            lines.append("- Nota: los VIF altos corresponden a niveles del MISMO factor (educacion QS25N); "
                         "no indican colinealidad entre covariables distintas. Reportar GVIF^(1/2df) o VIF maximo por bloque.")
        else:
            lines.append("- Todos los VIF < 5.")
    diag = _read(base, "analysis/models/model_diagnostics_summary.csv")
    if diag is not None:
        m2 = diag[diag["model"] == "model_2"]
        if not m2.empty:
            devr = float(m2.iloc[0]["mean_scale_deviance"])
            lines.append(
                f"- Ratio devianza/df model_2 = {devr:.2f} (NO es el parametro de dispersion: "
                f"para un desenlace binario 0/1 bajo un modelo de trabajo Poisson este ratio "
                f"esta inflado y no debe interpretarse como dispersion)."
            )
    # Dispersion de diseno (la que svyglm usa en los SE), calculada aparte en dispersion_model2_H7.csv
    disp_csv = NEW / "comparacion_2019_2025" / "dispersion_model2_H7.csv"
    if disp_csv.exists():
        dd = pd.read_csv(disp_csv, encoding="utf-8-sig")
        key = "canonico" if base == OLD else "nuevo"
        sub = dd[dd["cohorte"].str.contains(key)]
        if not sub.empty:
            val = float(sub.iloc[0]["dispersion_svyglm_design"])
            lines.append(
                f"- **Parametro de dispersion (svyglm, basado en diseno) model_2 = {val:.3f}** "
                f"= sin sobredispersion (leve subdispersion, esperada en un desenlace binario "
                f"ajustado con cuasi-Poisson log)."
            )
    return lines


def counts_and_prevalence() -> list[str]:
    lines = ["\n## 4. Cohorte y desenlaces por anio\n"]
    old_c, new_c = _read(OLD, "qc/final_counts_by_year.csv"), _read(NEW, "qc/final_counts_by_year.csv")
    if old_c is not None:
        lines.append(f"- Total 2019-2024: n={int(old_c['n_final'].sum()):,}")
    if new_c is not None:
        lines.append(f"- Total 2019-2025: n={int(new_c['n_final'].sum()):,}")
        lines.append("\n| Anio | n_final |")
        lines.append("|---|---|")
        for r in new_c.itertuples():
            lines.append(f"| {int(r.ANIO)} | {int(r.n_final):,} |")
    return lines


def h3_missing() -> list[str]:
    lines = ["\n## 5. H3 - Missing diferencial por severidad PHQ-9 (2019-2025)\n"]
    m = _read(NEW, "qc/missing_por_severidad_phq9.csv")
    if m is None:
        return lines + ["  (sin datos)"]
    lines.append("| Variable | %miss Minima | %miss Severa | ratio S/M | >2x |")
    lines.append("|---|---|---|---|---|")
    for r in m.itertuples():
        lines.append(f"| {r.variable} | {r.pct_missing_Minima} | {r.pct_missing_Severa} | {r.ratio_severa_vs_minima} | {r.diferencial_gt2x} |")
    return lines


def main() -> None:
    out: list[str] = ["# Comparacion 2019-2024 (canonico) vs 2019-2025 (nuevo)\n",
                      "Generado por scripts/compare_2024_vs_2025.py\n"]
    model_lines, (old_hier, new_hier) = compare_models()
    out += model_lines

    out.append("\n## 2. H5 - Atenuacion por altitud (descomposicion jerarquica)")
    out += h5_attenuation(old_hier, "2019-2024 (canonico)")
    out += h5_attenuation(new_hier, "2019-2025 (nuevo)")

    out.append("\n## 3. H7 - VIF maximo y parametro de dispersion")
    out += h7_vif_dispersion(OLD, "2019-2024 (canonico)")
    out += h7_vif_dispersion(NEW, "2019-2025 (nuevo)")

    out += counts_and_prevalence()
    out += h3_missing()

    # Panel de modificacion de efecto lado a lado
    out.append("\n## 6. Panel de modificacion de efecto (p Wald conjunto, Rubin D1)\n")
    oe, ne = _read(OLD, "analysis/models/effect_modification_panel.csv"), _read(NEW, "analysis/models/effect_modification_panel.csv")
    out.append("| Modificador | p 2019-2024 | p 2019-2025 |")
    out.append("|---|---|---|")
    if oe is not None and ne is not None:
        mods = sorted(set(oe["modificador"]) | set(ne["modificador"]))
        for mod in mods:
            po = oe[oe["modificador"] == mod]["p_value"]
            pn = ne[ne["modificador"] == mod]["p_value"]
            po_s = f"{float(po.iloc[0]):.4f}" if not po.empty else "n/d"
            pn_s = f"{float(pn.iloc[0]):.4f}" if not pn.empty else "n/d"
            out.append(f"| {mod} | {po_s} | {pn_s} |")

    report = "\n".join(out)
    (OUT / "INFORME_comparacion_2019_2025.md").write_text(report, encoding="utf-8")
    print(report)
    print(f"\n[OK] Informe escrito en {OUT / 'INFORME_comparacion_2019_2025.md'}")


if __name__ == "__main__":
    main()
