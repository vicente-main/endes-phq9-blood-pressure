"""Convierte las figuras vectoriales (PDF) del paquete de envio a TIFF 600 dpi
con compresion LZW, formato exigido por PLOS ONE (TIFF/EPS, 300-600 dpi).

Renderiza cada PDF con PyMuPDF a 600 dpi y lo guarda como TIFF (LZW) via Pillow,
grabando la resolucion en los metadatos. Solo aplica a figuras ya en ingles;
el DAG (S2) se regenera aparte en ingles antes de convertirlo.

Salida: <mismo nombre>.tif en ENVIO_PLOS_ONE_2019-2025/
"""
from __future__ import annotations
from pathlib import Path
import fitz  # PyMuPDF
from PIL import Image

ROOT = Path(__file__).resolve().parents[2]
ENVIO = ROOT / "ENVIO_PLOS_ONE_2019-2025"
DPI = 600

# PDFs vectoriales ya en ingles (el DAG se maneja por separado)
FIGURES = ["Fig1_STROBE.pdf", "S1_Fig_spline.pdf", "S3_Fig_forest_altitude.pdf"]


def pdf_to_tiff(pdf_path: Path, dpi: int = DPI) -> Path:
    doc = fitz.open(pdf_path)
    page = doc[0]
    pix = page.get_pixmap(dpi=dpi, alpha=False)
    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    out = pdf_path.with_suffix(".tif")
    img.save(out, format="TIFF", compression="tiff_lzw", dpi=(dpi, dpi))
    return out, pix.width, pix.height


def main():
    for name in FIGURES:
        p = ENVIO / name
        if not p.exists():
            print(f"  ⚠ no existe: {p}")
            continue
        out, w, h = pdf_to_tiff(p)
        mb = out.stat().st_size / 1e6
        print(f"  {name} -> {out.name}  {w}x{h}px @ {DPI}dpi  ({mb:.1f} MB)")


if __name__ == "__main__":
    main()
