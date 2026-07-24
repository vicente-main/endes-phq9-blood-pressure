"""Genera el manifiesto de hashes de los artefactos activos de la revisión."""

from __future__ import annotations

import csv
import hashlib
import subprocess
from datetime import datetime
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
SUBMISSION = ROOT / "ENVIO_PLOS_ONE_2019-2025"
REVIEWERS = SUBMISSION / "Revisores"
OUT_CSV = REVIEWERS / "MANIFEST_REVISION_2026-07-23.csv"
OUT_MD = REVIEWERS / "MANIFEST_REVISION_2026-07-23.md"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest().upper()


def git(*args: str) -> str:
    return subprocess.check_output(
        ["git", *args],
        cwd=ROOT,
        text=True,
        encoding="utf-8",
    ).strip()


def main() -> None:
    paths = [
        REVIEWERS / "MANUSCRITO_EN_PLOS_ONE.docx",
        SUBMISSION / "COVER_LETTER.docx",
        SUBMISSION / "Fig1.tif",
        SUBMISSION / "Fig2.tif",
        SUBMISSION / "S1_Fig.tif",
        SUBMISSION / "S2_Fig.tif",
        *[SUBMISSION / f"S{number}_Table.xlsx" for number in range(1, 12)],
        SUBMISSION / "S1_Checklist.docx",
        SUBMISSION / "S2_Checklist.docx",
        SUBMISSION / "PARA_EDITORIAL_MANAGER.md",
        SUBMISSION / "README_manifest.txt",
        REVIEWERS / "GUIA_CORREGIDA_REVISION_Y_EJECUCION_2026-07-23.md",
        REVIEWERS / "REFERENCE_AUDIT_2026-07-23.csv",
        REVIEWERS / "REFERENCE_AUDIT_2026-07-23.md",
        REVIEWERS / "ANALYTIC_CHANGE_AUDIT_2026-07-23.csv",
        REVIEWERS / "ANALYTIC_CHANGE_AUDIT_2026-07-23.md",
        REVIEWERS / "VALIDACION_GUIA_2026-07-23.json",
        REVIEWERS / "VALIDACION_GUIA_2026-07-23.md",
        REVIEWERS / "SOLICITUD_CIEI_P0-03_2026-07-23.md",
        REVIEWERS / "ZENODO_RELEASE_NOTES_v1.2.0.md",
        REVIEWERS / "ZENODO_METADATA_DRAFT_v1.2.0.json",
        REVIEWERS / "ZENODO_RELEASE_CANDIDATE_v1.2.0.zip",
        REVIEWERS / "ZENODO_RELEASE_CANDIDATE_v1.2.0.sha256",
        REVIEWERS / "ZENODO_RELEASE_CANDIDATE_v1.2.0.md",
        REVIEWERS / "ZENODO_RELEASE_VERIFICATION_v1.2.0.md",
        ROOT / "README.md",
        ROOT / "CITATION.cff",
        ROOT / "requirements.txt",
        ROOT / "config" / "pipeline_config.json",
        ROOT / "config" / "pipeline_config_2025.json",
        ROOT / "src" / "endes_pipeline" / "analysis.py",
        ROOT / "scripts" / "audit" / "export_table2_observed_applicable_2026_07_23.R",
        ROOT / "scripts" / "audit" / "validate_effect_modification_d1.R",
        ROOT / "scripts" / "audit" / "refit_effect_modification_exact_inputs_2026_07_23.py",
        ROOT / "scripts" / "audit" / "build_revision_supporting_artifacts_2026_07_23.py",
        ROOT / "scripts" / "audit" / "build_zenodo_release_candidate_2026_07_23.py",
        ROOT / "scripts" / "audit" / "update_cover_letter_2026_07_23.py",
        ROOT / "scripts" / "audit" / "update_checklists_2026_07_23.py",
        ROOT / "scripts" / "audit" / "update_remaining_supporting_text_2026_07_23.py",
        ROOT / "scripts" / "audit" / "update_submission_readme_2026_07_23.py",
        ROOT / "scripts" / "audit" / "update_plos_manuscript_2026_07_23.py",
        ROOT / "scripts" / "audit" / "audit_manuscript_references_2026_07_23.py",
        ROOT / "scripts" / "audit" / "build_analytic_change_audit_2026_07_23.py",
        ROOT / "scripts" / "audit" / "validate_guide_completion_2026_07_23.py",
        ROOT / "scripts" / "audit" / "build_revision_manifest_2026_07_23.py",
        ROOT / "data" / "output_2025" / "analysis" / "models" / "effect_modification_panel.csv",
        ROOT / "data" / "output_2025" / "analysis" / "models" / "effect_modification_sex_stratified.csv",
        ROOT / "data" / "output_2025" / "analysis" / "models" / "effect_modification_d1_inputs.json",
        ROOT / "data" / "output_2025" / "analysis" / "models" / "effect_modification_d1_mitml_validation.csv",
        ROOT / "data" / "output_2025" / "analysis" / "tables" / "table2_bivariate_observed_applicable.csv",
        ROOT / "data" / "output_2025" / "analysis" / "mice_observed_imputed_diagnostics.csv",
    ]
    missing = [str(path.relative_to(ROOT)) for path in paths if not path.exists()]
    if missing:
        raise FileNotFoundError("Faltan artefactos del manifiesto:\n" + "\n".join(missing))

    rows = []
    for path in paths:
        stat = path.stat()
        rows.append(
            {
                "path": path.relative_to(ROOT).as_posix(),
                "size_bytes": stat.st_size,
                "modified_local": datetime.fromtimestamp(stat.st_mtime).isoformat(timespec="seconds"),
                "sha256": sha256(path),
            }
        )
    with OUT_CSV.open("w", newline="", encoding="utf-8-sig") as stream:
        writer = csv.DictWriter(stream, fieldnames=list(rows[0]))
        writer.writeheader()
        writer.writerows(rows)

    head = git("rev-parse", "HEAD")
    description = git("describe", "--tags", "--always", "--dirty")
    lines = [
        "# Manifiesto de la revisión — 2026-07-23",
        "",
        f"- Git base: `{head}`",
        f"- Descripción del árbol: `{description}`",
        "- Los archivos `pre_guia` y el ZIP histórico no son artefactos activos de envío.",
        "",
        "| Archivo | Bytes | SHA-256 |",
        "|---|---:|---|",
    ]
    for row in rows:
        lines.append(f"| `{row['path']}` | {row['size_bytes']:,} | `{row['sha256']}` |")
    OUT_MD.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"Manifiesto: {OUT_CSV.relative_to(ROOT)}")
    print(f"Resumen: {OUT_MD.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
