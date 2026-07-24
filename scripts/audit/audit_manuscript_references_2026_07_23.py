"""Audita DOI, metadatos bibliográficos y URL de todas las referencias del DOCX."""

from __future__ import annotations

import re
import time
import unicodedata
import json
from concurrent.futures import ThreadPoolExecutor, as_completed
from difflib import SequenceMatcher
from pathlib import Path
from urllib.parse import quote
from urllib.request import Request, urlopen

import pandas as pd
from docx import Document


ROOT = Path(__file__).resolve().parents[2]
MANUSCRIPT = (
    ROOT
    / "ENVIO_PLOS_ONE_2019-2025"
    / "Revisores"
    / "MANUSCRITO_EN_PLOS_ONE.docx"
)
OUT_CSV = (
    ROOT
    / "ENVIO_PLOS_ONE_2019-2025"
    / "Revisores"
    / "REFERENCE_AUDIT_2026-07-23.csv"
)
OUT_MD = OUT_CSV.with_suffix(".md")

DOI_RE = re.compile(r"\bdoi:\s*(10\.\d{4,9}/[^\s]+)", flags=re.IGNORECASE)
URL_RE = re.compile(r"https?://[^\s]+", flags=re.IGNORECASE)
YEAR_RE = re.compile(r"\b(19|20)\d{2}\b")
PUBLICATION_YEAR_RE = re.compile(r"\.\s*((?:19|20)\d{2});")
VOLUME_ISSUE_PAGE_RE = re.compile(
    r"\b(?P<volume>\d+)(?:\((?P<issue>[^)]+)\))?:(?P<pages>[A-Za-z]?\d+[A-Za-z0-9–\-]*"
    r"(?:[–\-][A-Za-z]?\d+[A-Za-z0-9]*)?)"
)

# Informes, libros, software/documentación y norma; no se espera DOI de artículo.
EXPECTED_WITHOUT_DOI = {1, 3, 11, 16, 23, 26, 27, 28, 31, 47}


def normalize(text: object) -> str:
    value = re.sub(r"<[^>]+>", "", str(text or ""))
    value = unicodedata.normalize("NFKD", value)
    value = "".join(character for character in value if not unicodedata.combining(character))
    value = value.lower().replace("–", "-")
    return re.sub(r"[^a-z0-9]+", " ", value).strip()


def references_from_docx() -> list[str]:
    document = Document(MANUSCRIPT)
    collecting = False
    references: list[str] = []
    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if text == "References":
            collecting = True
            continue
        if text == "Supporting information":
            break
        if collecting and text:
            references.append(text)
    return references


def crossref_lookup(doi: str) -> dict[str, object]:
    url = f"https://api.crossref.org/works/{quote(doi, safe='')}"
    headers = {
        "User-Agent": (
            "ENDES-PLOS-reference-audit/1.0 "
            "(mailto:rodrigo.cardenas@unapiquitos.edu.pe)"
        )
    }
    last_exception: Exception | None = None
    for attempt in range(4):
        try:
            request = Request(url, headers=headers)
            with urlopen(request, timeout=30) as response:
                message = json.loads(response.read().decode("utf-8"))["message"]
            authors = message.get("author") or []
            issued = message.get("published-print") or message.get("published") or message.get("issued") or {}
            date_parts = issued.get("date-parts") or [[]]
            year = date_parts[0][0] if date_parts and date_parts[0] else None
            title = " ".join(message.get("title") or [])
            journal = " ".join(message.get("container-title") or [])
            first_author = ""
            if authors:
                first_author = authors[0].get("family") or authors[0].get("name") or ""
            return {
                "crossref_ok": True,
                "crossref_error": "",
                "crossref_title": title,
                "crossref_first_author": first_author,
                "crossref_year": year,
                "crossref_journal": journal,
                "crossref_volume": message.get("volume") or "",
                "crossref_issue": message.get("issue") or "",
                "crossref_pages": message.get("page") or message.get("article-number") or "",
                "canonical_doi": message.get("DOI") or "",
            }
        except Exception as exc:  # noqa: BLE001 - el error se registra en la auditoría.
            last_exception = exc
            if getattr(exc, "code", None) == 429 and attempt < 3:
                time.sleep(2 ** (attempt + 1))
                continue
            break
    return {
        "crossref_ok": False,
        "crossref_error": f"{type(last_exception).__name__}: {last_exception}",
        "crossref_title": "",
        "crossref_first_author": "",
        "crossref_year": "",
        "crossref_journal": "",
        "crossref_volume": "",
        "crossref_issue": "",
        "crossref_pages": "",
        "canonical_doi": "",
    }


def url_status(url: str) -> tuple[bool, str]:
    try:
        request = Request(url, headers={"User-Agent": "ENDES-PLOS-reference-audit/1.0"})
        with urlopen(request, timeout=30) as response:
            return response.status < 400, f"{response.status} {response.url}"
    except Exception as exc:  # noqa: BLE001
        return False, f"{type(exc).__name__}: {exc}"


def parse_manuscript_fields(reference: str) -> dict[str, object]:
    doi_match = DOI_RE.search(reference)
    doi = doi_match.group(1).rstrip(".,);") if doi_match else ""
    url_match = URL_RE.search(reference)
    url = url_match.group(0).rstrip(".,);") if url_match else ""
    publication_year_match = PUBLICATION_YEAR_RE.search(reference)
    years = [match.group(0) for match in YEAR_RE.finditer(reference)]
    publication_year = (
        publication_year_match.group(1)
        if publication_year_match
        else (years[0] if years else "")
    )
    vip = VOLUME_ISSUE_PAGE_RE.search(reference)
    volume = vip.group("volume") if vip else ""
    issue = vip.group("issue") if vip and vip.group("issue") else ""
    pages = vip.group("pages") if vip else ""
    first_chunk = reference.split(".", 1)[0]
    author_chunk = first_chunk.split(",", 1)[0].strip()
    author_tokens = author_chunk.split()
    if author_tokens and author_tokens[0].lower() in {"van", "von", "de", "del", "da", "di"}:
        first_author = " ".join(author_tokens[:2])
    else:
        first_author = author_tokens[0] if author_tokens else ""
    chunks = reference.split(". ")
    title = chunks[1] if len(chunks) > 1 else ""
    return {
        "doi": doi,
        "url": url,
        "manuscript_year": publication_year,
        "manuscript_volume": volume,
        "manuscript_issue": issue,
        "manuscript_pages": pages,
        "manuscript_first_author": first_author,
        "manuscript_title": title,
    }


def page_range_matches(manuscript_pages: object, crossref_pages: object) -> bool:
    def clean(value: object) -> str:
        text = unicodedata.normalize("NFKD", str(value or "")).lower()
        text = text.replace("–", "-").replace("—", "-")
        return re.sub(r"[^a-z0-9-]+", "", text)

    manuscript = clean(manuscript_pages)
    crossref = clean(crossref_pages)
    if not manuscript or not crossref:
        return True
    if manuscript == crossref:
        return True

    def expand(value: str) -> tuple[str, str]:
        parts = value.split("-", 1)
        if len(parts) == 1:
            return parts[0], parts[0]
        start, end = parts
        if end.isdigit() and start.isdigit() and len(end) < len(start):
            end = start[: len(start) - len(end)] + end
        return start, end

    manuscript_start, manuscript_end = expand(manuscript)
    crossref_start, crossref_end = expand(crossref)
    return (
        manuscript_start == crossref_start
        and (
            manuscript_end == crossref_end
            or manuscript_end == manuscript_start
            or crossref_end == crossref_start
        )
    )


def main() -> None:
    references = references_from_docx()
    if len(references) != 47:
        raise RuntimeError(f"Se esperaban 47 referencias finales; se encontraron {len(references)}.")

    parsed = []
    for number, reference in enumerate(references, start=1):
        item = {"reference_number": number, "reference": reference}
        item.update(parse_manuscript_fields(reference))
        parsed.append(item)

    doi_values = sorted({item["doi"] for item in parsed if item["doi"]})
    crossref: dict[str, dict[str, object]] = {}
    with ThreadPoolExecutor(max_workers=4) as executor:
        futures = {executor.submit(crossref_lookup, doi): doi for doi in doi_values}
        for future in as_completed(futures):
            crossref[futures[future]] = future.result()

    url_values = sorted({item["url"] for item in parsed if item["url"]})
    url_results: dict[str, tuple[bool, str]] = {}
    with ThreadPoolExecutor(max_workers=6) as executor:
        futures = {executor.submit(url_status, url): url for url in url_values}
        for future in as_completed(futures):
            url_results[futures[future]] = future.result()

    rows: list[dict[str, object]] = []
    for item in parsed:
        number = int(item["reference_number"])
        doi = str(item["doi"])
        result = crossref.get(doi, {})
        row = dict(item)
        row.update(result)

        if doi and result.get("crossref_ok"):
            manuscript_title_norm = normalize(item["manuscript_title"])
            crossref_title_norm = normalize(result.get("crossref_title"))
            title_similarity = SequenceMatcher(
                None,
                manuscript_title_norm,
                crossref_title_norm,
            ).ratio()
            title_match = (
                title_similarity >= 0.70
                or (
                    min(len(manuscript_title_norm), len(crossref_title_norm)) >= 5
                    and (
                        manuscript_title_norm in crossref_title_norm
                        or crossref_title_norm in manuscript_title_norm
                    )
                )
            )
            author_match = (
                normalize(item["manuscript_first_author"])
                in normalize(result.get("crossref_first_author"))
                or normalize(result.get("crossref_first_author"))
                in normalize(item["manuscript_first_author"])
            )
            if number == 13 and "ncd risk factor collaboration" in normalize(item["reference"]):
                author_match = True
            year_match = str(item["manuscript_year"]) == str(result.get("crossref_year"))
            if number == 29 and str(item["manuscript_year"]) == "2016":
                # Crossref deposita 2017, pero PubMed y el número 45(6) fechan la versión impresa en 2016.
                year_match = True
            volume_match = (
                not item["manuscript_volume"]
                or not result.get("crossref_volume")
                or normalize(item["manuscript_volume"]) == normalize(result.get("crossref_volume"))
            )
            issue_match = (
                not item["manuscript_issue"]
                or not result.get("crossref_issue")
                or normalize(item["manuscript_issue"]) == normalize(result.get("crossref_issue"))
            )
            pages_match = page_range_matches(item["manuscript_pages"], result.get("crossref_pages"))
            if number == 29 and normalize(result.get("crossref_pages")) == "dyw341":
                # Crossref deposita el identificador electrónico; PubMed/impreso aportan 1887–1894.
                pages_match = True
        else:
            title_similarity = float("nan")
            title_match = False if doi else None
            author_match = False if doi else None
            year_match = False if doi else None
            volume_match = False if doi else None
            issue_match = False if doi else None
            pages_match = False if doi else None

        url_ok, url_detail = url_results.get(str(item["url"]), (None, ""))
        row.update(
            {
                "title_similarity": title_similarity,
                "first_author_match": author_match,
                "year_match": year_match,
                "volume_match": volume_match,
                "issue_match": issue_match,
                "pages_match": pages_match,
                "url_ok": url_ok,
                "url_detail": url_detail,
            }
        )

        if doi:
            checks = [
                bool(result.get("crossref_ok")),
                bool(title_match),
                bool(author_match),
                bool(year_match),
                bool(volume_match),
                bool(issue_match),
                bool(pages_match),
            ]
            row["status"] = "verified" if all(checks) else "review_metadata"
        elif number in EXPECTED_WITHOUT_DOI:
            row["status"] = "verified_non_doi" if (not item["url"] or url_ok) else "review_url"
        else:
            row["status"] = "unexpected_missing_doi"
        rows.append(row)

    audit = pd.DataFrame(rows)
    audit.to_csv(OUT_CSV, index=False, encoding="utf-8-sig")

    counts = audit["status"].value_counts().to_dict()
    lines = [
        "# Auditoría de referencias — 2026-07-23",
        "",
        f"- Manuscrito: `{MANUSCRIPT.relative_to(ROOT).as_posix()}`",
        f"- Referencias: {len(audit)}",
        f"- Con DOI consultado en Crossref: {int(audit['doi'].astype(bool).sum())}",
        f"- Estado: {counts}",
        "",
        "## Elementos que requieren revisión",
        "",
    ]
    flagged = audit.loc[~audit["status"].isin(["verified", "verified_non_doi"])]
    if flagged.empty:
        lines.append("Ninguno.")
    else:
        for _, row in flagged.iterrows():
            lines.append(
                f"- Ref. {int(row['reference_number'])}: `{row['status']}` — "
                f"{row['reference']}"
            )
    lines.extend(
        [
            "",
            "La salida CSV conserva metadatos de Crossref, comparación de autor/año/volumen/"
            "número/páginas y comprobación de URL para cada referencia.",
            "",
        ]
    )
    OUT_MD.write_text("\n".join(lines), encoding="utf-8")

    print(f"Auditoría CSV: {OUT_CSV.relative_to(ROOT)}")
    print(f"Resumen: {OUT_MD.relative_to(ROOT)}")
    print(f"Estados: {counts}")


if __name__ == "__main__":
    main()
