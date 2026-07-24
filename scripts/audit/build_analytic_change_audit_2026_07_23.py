"""Documenta qué cambió analíticamente con P0-01/P0-02."""

from __future__ import annotations

import math
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parents[2]
ANALYSIS = ROOT / "data" / "output_2025" / "analysis"
BEFORE = ANALYSIS / "pre_d1_fix_2026-07-23" / "models"
AFTER = ANALYSIS / "models"
REVIEWERS = ROOT / "ENVIO_PLOS_ONE_2019-2025" / "Revisores"
OUT_CSV = REVIEWERS / "ANALYTIC_CHANGE_AUDIT_2026-07-23.csv"
OUT_MD = REVIEWERS / "ANALYTIC_CHANGE_AUDIT_2026-07-23.md"
MODIFIER_NAMES = {"riqueza": "wealth", "altitud": "altitude"}


def max_numeric_delta(before: pd.DataFrame, after: pd.DataFrame) -> float:
    keys = [column for column in ["model", "term"] if column in before.columns]
    merged = before.merge(after, on=keys, suffixes=("_before", "_after"), how="outer", indicator=True)
    if not merged["_merge"].eq("both").all():
        return math.inf
    deltas = []
    for column in before.select_dtypes(include="number").columns:
        if column in after.columns:
            deltas.append(
                (
                    pd.to_numeric(merged[f"{column}_before"], errors="coerce")
                    - pd.to_numeric(merged[f"{column}_after"], errors="coerce")
                ).abs().max()
            )
    return float(pd.Series(deltas).max()) if deltas else 0.0


def main() -> None:
    stable_files = [
        "table3_main_models.csv",
        "table4_cascade_models.csv",
        "hierarchical_decomposition.csv",
    ]
    stability: dict[str, float] = {}
    for filename in stable_files:
        before_path = BEFORE / filename
        after_path = AFTER / filename
        if before_path.exists() and after_path.exists():
            stability[filename] = max_numeric_delta(
                pd.read_csv(before_path),
                pd.read_csv(after_path),
            )

    if stability.get("table3_main_models.csv", math.inf) > 1e-12:
        raise RuntimeError("El Modelo principal cambió durante la corrección secundaria.")

    before = pd.read_csv(BEFORE / "effect_modification_panel.csv")
    after = pd.read_csv(AFTER / "effect_modification_panel.csv")
    before["modificador"] = before["modificador"].replace(MODIFIER_NAMES)
    after["modificador"] = after["modificador"].replace(MODIFIER_NAMES)
    before = before.set_index("modificador")
    after = after.set_index("modificador")
    rows = []
    for modifier in after.index:
        old = before.loc[modifier]
        new = after.loc[modifier]
        rows.append(
            {
                "modifier": modifier,
                "old_method": old["pool_method"],
                "new_method": new["pool_method"],
                "old_df1": old["df_num"],
                "new_df1": new["df_num"],
                "old_df2": old.get("df_den", float("nan")),
                "new_df2": new["df_den"],
                "old_F_or_approx": old["statistic"],
                "new_D1_F": new["statistic"],
                "old_p": old["p_value"],
                "new_p": new["p_value"],
                "old_Holm": old.get("p_holm_exploratorios", float("nan")),
                "new_Holm": new.get("p_holm_exploratorios", float("nan")),
                "old_n": old["mean_n_obs"],
                "new_n": new["mean_n_obs"],
                "old_terms": old["terms"],
                "new_terms": new["terms"],
            }
        )
    comparison = pd.DataFrame(rows)
    comparison.to_csv(OUT_CSV, index=False, encoding="utf-8-sig")

    dx = after.loc["dxhta"]
    if (
        int(dx["df_num"]) != 1
        or int(round(float(dx["mean_n_obs"]))) != 191565
        or "nan" in str(dx["terms"]).lower()
    ):
        raise RuntimeError("P0-01 no quedó cerrado en el panel posterior.")

    lines = [
        "# Auditoría del cambio analítico — 2026-07-23",
        "",
        "P0-01 excluyó los 192 diagnósticos previos faltantes y dejó el contraste "
        "PHQ-9 × diagnóstico previo como binario de un grado de libertad. P0-02 "
        "sustituyó la aproximación chi-cuadrado por D1 F con df denominador finito.",
        "",
        "## Invariantes",
        "",
    ]
    for filename, delta in stability.items():
        lines.append(f"- `{filename}`: diferencia numérica máxima `{delta:.3g}`.")
    lines.extend(
        [
            "",
            "## Panel antes/después",
            "",
            "| Modificador | df1 antes→después | n antes→después | p antes→después | Holm antes→después |",
            "|---|---:|---:|---:|---:|",
        ]
    )
    for _, item in comparison.iterrows():
        old_holm = "—" if pd.isna(item["old_Holm"]) else f"{item['old_Holm']:.6g}"
        new_holm = "—" if pd.isna(item["new_Holm"]) else f"{item['new_Holm']:.6g}"
        lines.append(
            f"| {item['modifier']} | {item['old_df1']:.0f}→{item['new_df1']:.0f} | "
            f"{item['old_n']:.0f}→{item['new_n']:.0f} | "
            f"{item['old_p']:.6g}→{item['new_p']:.6g} | {old_holm}→{new_holm} |"
        )
    OUT_MD.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"Auditoría analítica: {OUT_MD.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
