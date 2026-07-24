"""Construye un candidato reproducible y sin microdatos para Zenodo v1.2.0."""

from __future__ import annotations

import hashlib
import json
import subprocess
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
REVIEWERS = ROOT / "ENVIO_PLOS_ONE_2019-2025" / "Revisores"
VERSION = "v1.2.0"
PREFIX = f"endes-phq9-blood-pressure-{VERSION}"
ARCHIVE = REVIEWERS / f"ZENODO_RELEASE_CANDIDATE_{VERSION}.zip"
CHECKSUM = REVIEWERS / f"ZENODO_RELEASE_CANDIDATE_{VERSION}.sha256"
REPORT = REVIEWERS / f"ZENODO_RELEASE_CANDIDATE_{VERSION}.md"
FIXED_TIME = (2026, 7, 23, 0, 0, 0)

MODELS = ROOT / "data" / "output_2025" / "analysis" / "models"
TABLES = ROOT / "data" / "output_2025" / "analysis" / "tables"
FIGURES = ROOT / "data" / "output_2025" / "analysis" / "figures"
ANALYSIS = ROOT / "data" / "output_2025" / "analysis"

CANONICAL_FILES = [
    ".gitignore",
    "CITATION.cff",
    "LICENSE",
    "README.md",
    "requirements.txt",
    "config/pipeline_config.json",
    "config/pipeline_config_2025.json",
    "scripts/run_analysis.py",
    "scripts/run_pipeline.py",
    "scripts/setup_env.ps1",
    "src/endes_pipeline/__init__.py",
    "src/endes_pipeline/analysis.py",
    "src/endes_pipeline/pipeline.py",
]

EXTRA_CODE = [
    "scripts/audit/audit_manuscript_references_2026_07_23.py",
    "scripts/audit/build_analytic_change_audit_2026_07_23.py",
    "scripts/audit/build_revision_manifest_2026_07_23.py",
    "scripts/audit/build_revision_supporting_artifacts_2026_07_23.py",
    "scripts/audit/build_zenodo_release_candidate_2026_07_23.py",
    "scripts/audit/export_table2_observed_applicable_2026_07_23.R",
    "scripts/audit/refit_effect_modification_exact_inputs_2026_07_23.py",
    "scripts/audit/update_checklists_2026_07_23.py",
    "scripts/audit/update_cover_letter_2026_07_23.py",
    "scripts/audit/update_plos_manuscript_2026_07_23.py",
    "scripts/audit/update_remaining_supporting_text_2026_07_23.py",
    "scripts/audit/update_submission_readme_2026_07_23.py",
    "scripts/audit/validate_effect_modification_d1.R",
    "scripts/audit/validate_guide_completion_2026_07_23.py",
]

FROZEN_OUTPUTS = {
    MODELS / "table3_main_models.csv": "frozen_outputs/analysis/models/table3_main_models.csv",
    MODELS / "table4_cascade_models.csv": "frozen_outputs/analysis/models/table4_cascade_models.csv",
    MODELS / "hierarchical_decomposition.csv": "frozen_outputs/analysis/models/hierarchical_decomposition.csv",
    MODELS / "complete_case_sensitivity.csv": "frozen_outputs/analysis/models/complete_case_sensitivity.csv",
    MODELS / "interactions_and_sensitivity_models.csv": (
        "frozen_outputs/analysis/models/interactions_and_sensitivity_models.csv"
    ),
    MODELS / "effect_modification_panel.csv": (
        "frozen_outputs/analysis/models/effect_modification_panel.csv"
    ),
    MODELS / "effect_modification_sex_stratified.csv": (
        "frozen_outputs/analysis/models/effect_modification_sex_stratified.csv"
    ),
    MODELS / "effect_modification_d1_inputs.json": (
        "frozen_outputs/analysis/models/effect_modification_d1_inputs.json"
    ),
    MODELS / "effect_modification_d1_mitml_validation.csv": (
        "frozen_outputs/analysis/models/effect_modification_d1_mitml_validation.csv"
    ),
    TABLES / "table2_bivariate_observed_applicable.csv": (
        "frozen_outputs/analysis/tables/table2_bivariate_observed_applicable.csv"
    ),
    FIGURES / "spline_nonlinearity_summary.csv": (
        "frozen_outputs/analysis/figures/spline_nonlinearity_summary.csv"
    ),
    ANALYSIS / "mice_observed_imputed_diagnostics.csv": (
        "frozen_outputs/analysis/mice_observed_imputed_diagnostics.csv"
    ),
}

VALIDATION_FILES = {
    REVIEWERS / "ANALYTIC_CHANGE_AUDIT_2026-07-23.csv": (
        "validation/ANALYTIC_CHANGE_AUDIT_2026-07-23.csv"
    ),
    REVIEWERS / "ANALYTIC_CHANGE_AUDIT_2026-07-23.md": (
        "validation/ANALYTIC_CHANGE_AUDIT_2026-07-23.md"
    ),
    REVIEWERS / "ZENODO_RELEASE_NOTES_v1.2.0.md": "ZENODO_RELEASE_NOTES_v1.2.0.md",
    REVIEWERS / "ZENODO_METADATA_DRAFT_v1.2.0.json": "ZENODO_METADATA_DRAFT_v1.2.0.json",
}

FORBIDDEN_SUFFIXES = {
    ".docx",
    ".xlsx",
    ".pdf",
    ".parquet",
    ".dta",
    ".sav",
    ".dbf",
    ".tif",
    ".tiff",
    ".png",
    ".zip",
}
FORBIDDEN_FRAGMENTS = {
    ".venv",
    "__pycache__",
    "data/input",
    "/imputed/",
    "solicitud_ciei",
    "agents.md",
    "claude.md",
    "system prompt",
}


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest().upper()


def sha256_file(path: Path) -> str:
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


def add_entry(
    entries: dict[str, bytes],
    source: Path,
    relative_destination: str,
) -> None:
    if not source.exists():
        raise FileNotFoundError(source)
    destination = relative_destination.replace("\\", "/").lstrip("/")
    entries[destination] = source.read_bytes()


def provenance() -> bytes:
    payload = {
        "release_candidate": VERSION,
        "built_on": "2026-07-23",
        "git_head": git("rev-parse", "HEAD"),
        "git_description": git("describe", "--tags", "--always", "--dirty"),
        "concept_doi": "10.5281/zenodo.21328300",
        "previous_version": {
            "version": "v1.1.0",
            "doi": "10.5281/zenodo.21403130",
            "access_observed": "restricted",
        },
        "target_access": "public",
        "analytic_validation": {
            "d1_tests_reproduced_with_mitml": 6,
            "d1_tolerance": "1e-10",
            "main_model_pr": 0.9949479044154493,
            "main_model_n": 191757,
        },
        "contains_person_level_data": False,
        "prepublication_note": (
            "Rebuild after committing and tagging v1.2.0; publishing to Zenodo is an "
            "external action not performed by this builder."
        ),
    }
    return (json.dumps(payload, indent=2, ensure_ascii=False) + "\n").encode("utf-8")


def build_entries() -> dict[str, bytes]:
    entries: dict[str, bytes] = {}
    for relative in CANONICAL_FILES:
        add_entry(entries, ROOT / relative, relative)
    for relative in EXTRA_CODE:
        add_entry(entries, ROOT / relative, relative)
    for source, destination in FROZEN_OUTPUTS.items():
        add_entry(entries, source, destination)
    for source, destination in VALIDATION_FILES.items():
        add_entry(entries, source, destination)
    entries["RELEASE_PROVENANCE.json"] = provenance()

    normalized = {path.lower() for path in entries}
    if len(normalized) != len(entries):
        raise RuntimeError("El candidato contiene rutas duplicadas por diferencias de mayúsculas.")

    violations = []
    for path in entries:
        lower = path.lower()
        if Path(lower).suffix in FORBIDDEN_SUFFIXES:
            violations.append(path)
        if any(fragment in lower for fragment in FORBIDDEN_FRAGMENTS):
            violations.append(path)
    if violations:
        raise RuntimeError(
            "El candidato contiene material prohibido:\n" + "\n".join(sorted(set(violations)))
        )
    return entries


def checksums_text(entries: dict[str, bytes]) -> bytes:
    lines = [
        f"{sha256_bytes(data)}  {path}"
        for path, data in sorted(entries.items())
    ]
    return ("\n".join(lines) + "\n").encode("utf-8")


def write_archive(entries: dict[str, bytes]) -> None:
    entries = dict(entries)
    entries["SHA256SUMS.txt"] = checksums_text(entries)
    temporary = ARCHIVE.with_suffix(".zip.tmp")
    with zipfile.ZipFile(
        temporary,
        mode="w",
        compression=zipfile.ZIP_DEFLATED,
        compresslevel=9,
    ) as archive:
        for relative, data in sorted(entries.items()):
            info = zipfile.ZipInfo(f"{PREFIX}/{relative}", date_time=FIXED_TIME)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o100644 << 16
            info.create_system = 3
            archive.writestr(info, data, compress_type=zipfile.ZIP_DEFLATED, compresslevel=9)
    temporary.replace(ARCHIVE)


def verify_archive(expected_entries: dict[str, bytes]) -> tuple[int, int]:
    with zipfile.ZipFile(ARCHIVE) as archive:
        members = [item for item in archive.infolist() if not item.is_dir()]
        expected_names = {
            f"{PREFIX}/{path}" for path in expected_entries
        } | {f"{PREFIX}/SHA256SUMS.txt"}
        observed_names = {item.filename for item in members}
        if observed_names != expected_names:
            raise RuntimeError(
                "Las rutas del ZIP no coinciden con el conjunto esperado: "
                f"faltan={sorted(expected_names - observed_names)}, "
                f"sobran={sorted(observed_names - expected_names)}"
            )

        for path, data in expected_entries.items():
            archived = archive.read(f"{PREFIX}/{path}")
            if archived != data:
                raise RuntimeError(f"Contenido discrepante dentro del ZIP: {path}")

        checksum_rows = archive.read(f"{PREFIX}/SHA256SUMS.txt").decode("utf-8").splitlines()
        parsed: dict[str, str] = {}
        for row in checksum_rows:
            digest, path = row.split("  ", 1)
            parsed[path] = digest
        if set(parsed) != set(expected_entries):
            raise RuntimeError("SHA256SUMS.txt no cubre exactamente los archivos de contenido.")
        for path, data in expected_entries.items():
            if parsed[path] != sha256_bytes(data):
                raise RuntimeError(f"Checksum interno discrepante: {path}")
        return len(members), sum(item.file_size for item in members)


def write_external_report(file_count: int, uncompressed_bytes: int) -> None:
    archive_hash = sha256_file(ARCHIVE)
    CHECKSUM.write_text(
        f"{archive_hash}  {ARCHIVE.name}\n",
        encoding="ascii",
    )
    lines = [
        "# Candidato de publicación Zenodo v1.2.0",
        "",
        f"- Archivo: `{ARCHIVE.name}`",
        f"- Tamaño: {ARCHIVE.stat().st_size:,} bytes",
        f"- SHA-256: `{archive_hash}`",
        f"- Archivos internos: {file_count}",
        f"- Bytes internos sin comprimir: {uncompressed_bytes:,}",
        f"- Prefijo interno: `{PREFIX}/`",
        "- Microdatos o imputaciones a nivel individual: **no incluidos**",
        "- Estado: **candidato local; no publicado**",
        "",
        "La publicación debe hacerse mediante **New version** desde `v1.1.0`, con acceso de archivos "
        "**público**. El ZIP debe reconstruirse después de comprometer y etiquetar el código como "
        "`v1.2.0`, porque el estado actual del árbol es `v1.1.0-dirty`.",
        "",
        "El archivo `.sha256` contiguo permite comprobar el ZIP antes y después de subirlo.",
    ]
    REPORT.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> None:
    entries = build_entries()
    write_archive(entries)
    file_count, uncompressed_bytes = verify_archive(entries)
    write_external_report(file_count, uncompressed_bytes)
    print(f"Archivo: {ARCHIVE.relative_to(ROOT)}")
    print(f"SHA-256: {sha256_file(ARCHIVE)}")
    print(f"Archivos internos: {file_count}")
    print(f"Reporte: {REPORT.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
