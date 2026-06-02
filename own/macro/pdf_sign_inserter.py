#!/usr/bin/env python3
"""
Insert signature images from PDF files in the `sign` folder into PDF documents
from the `in` folder, and save results to the `out` folder.

Folder layout (relative to this script):
    in/   — multi-page PDF documents to receive images
    sign/ — single-page PDF documents containing images to insert
    out/  — output PDFs with images inserted

Insertion pattern:
    Given N sign documents, page k of an input PDF receives the image from
    sign document number (k-1) mod N (0-indexed).
    Example with 3 sign docs: pages 1,4,7→sign[0]; pages 2,5,8→sign[1];
    pages 3,6,9→sign[2].

Page size adjustment:
    If a page is not an integer multiple of A4 in both dimensions, it is
    padded/scaled to the nearest A4 multiple.

Usage:
    python3 pdf_sign_inserter.py [--in-dir IN] [--sign-dir SIGN] [--out-dir OUT]

Dependencies:
    pip install pymupdf
"""

import argparse
import sys
from pathlib import Path

try:
    import pymupdf
except ImportError:
    sys.exit("Error: PyMuPDF not installed. Run: pip install pymupdf")

# A4 dimensions in PDF points (72 pt = 1 inch; A4 = 210×297 mm)
A4_WIDTH_PT = 595.28   # 210 mm
A4_HEIGHT_PT = 841.89  # 297 mm


def ensure_dirs(*dirs: Path) -> None:
    """Create directories if they don't exist."""
    for d in dirs:
        d.mkdir(parents=True, exist_ok=True)


def check_non_empty(folder: Path, label: str) -> None:
    """Exit with a warning if the folder has no PDF files."""
    pdfs = sorted(folder.glob("*.pdf"))
    if not pdfs:
        sys.exit(
            f"Warning: The '{label}' folder ({folder}) contains no PDF files.\n"
            f"Please add PDF documents to '{folder}' and run the script again."
        )


def nearest_a4_multiple(size_pt: float, a4_dim: float) -> float:
    """Return the nearest integer multiple of a4_dim that is >= size_pt."""
    import math
    n = math.ceil(size_pt / a4_dim)
    return max(n, 1) * a4_dim


def is_a4_multiple(width: float, height: float, tol: float = 1.0) -> bool:
    """Return True if width and height are already integer multiples of A4."""
    import math
    w_mult = width / A4_WIDTH_PT
    h_mult = height / A4_HEIGHT_PT
    return (
        abs(w_mult - round(w_mult)) < tol / A4_WIDTH_PT
        and abs(h_mult - round(h_mult)) < tol / A4_HEIGHT_PT
    )


def render_sign_page_as_image(sign_doc: pymupdf.Document) -> pymupdf.Pixmap:
    """Render the first page of a sign document to a high-resolution pixmap."""
    page = sign_doc[0]
    mat = pymupdf.Matrix(2, 2)  # 2x zoom → 144 dpi
    return page.get_pixmap(matrix=mat, alpha=True)


def get_sign_rect_on_page(
    sign_doc: pymupdf.Document,
) -> tuple[pymupdf.Rect, pymupdf.Pixmap]:
    """
    Return the bounding rect of content on the first page of the sign doc,
    and a pixmap rendering of that content.
    """
    page = sign_doc[0]
    # Try to get the bbox of actual content (tighter than full page)
    content_rect = page.rect  # fall back to full page rect
    mat = pymupdf.Matrix(2, 2)
    clip_mat = pymupdf.Matrix(2, 2)
    pix = page.get_pixmap(matrix=clip_mat, alpha=True, clip=content_rect)
    return content_rect, pix


def adjust_page_size(page: pymupdf.Page) -> None:
    """
    Resize the page to the nearest integer multiple of A4 dimensions if needed.
    Content is kept at its current position; extra space is transparent.
    """
    rect = page.rect
    w, h = rect.width, rect.height

    if is_a4_multiple(w, h):
        return

    new_w = nearest_a4_multiple(w, A4_WIDTH_PT)
    new_h = nearest_a4_multiple(h, A4_HEIGHT_PT)

    new_rect = pymupdf.Rect(0, 0, new_w, new_h)
    page.set_mediabox(new_rect)


def insert_sign_into_page(
    out_page: pymupdf.Page,
    sign_doc: pymupdf.Document,
    sign_pix: pymupdf.Pixmap,
    sign_rect: pymupdf.Rect,
) -> None:
    """
    Overlay the sign image onto out_page.

    The sign image is placed at the same relative position it occupies on the
    sign PDF page (preserving its size in points).  If it doesn't fit on the
    target page, it is scaled down proportionally.
    """
    page_rect = out_page.rect

    sign_page = sign_doc[0]
    sign_page_rect = sign_page.rect

    # Position the sign's content rect on the target page (same offset from top-left)
    # Scale if the sign page is larger than the target page.
    scale_x = min(1.0, page_rect.width / sign_page_rect.width)
    scale_y = min(1.0, page_rect.height / sign_page_rect.height)
    scale = min(scale_x, scale_y)

    dest_x0 = sign_rect.x0 * scale
    dest_y0 = sign_rect.y0 * scale
    dest_x1 = sign_rect.x1 * scale
    dest_y1 = sign_rect.y1 * scale
    dest_rect = pymupdf.Rect(dest_x0, dest_y0, dest_x1, dest_y1)

    out_page.insert_image(dest_rect, pixmap=sign_pix, overlay=True)


def process_pdf(
    in_path: Path,
    sign_docs: list[pymupdf.Document],
    sign_pixmaps: list[tuple[pymupdf.Rect, pymupdf.Pixmap]],
    out_path: Path,
) -> None:
    """Process a single input PDF: adjust page sizes and insert sign images."""
    print(f"  Processing: {in_path.name}")
    doc = pymupdf.open(str(in_path))
    n_signs = len(sign_docs)

    for page_idx in range(doc.page_count):
        page = doc[page_idx]

        # Adjust page size to A4 multiple if needed
        adjust_page_size(page)

        # Select sign document for this page (0-indexed cycling)
        sign_idx = page_idx % n_signs
        sign_rect, sign_pix = sign_pixmaps[sign_idx]

        insert_sign_into_page(page, sign_docs[sign_idx], sign_pix, sign_rect)

    doc.save(str(out_path), deflate=True)
    doc.close()
    print(f"    Saved: {out_path.name}")


def main() -> None:
    script_dir = Path(__file__).resolve().parent

    parser = argparse.ArgumentParser(
        description=(
            "Insert signature images from PDFs in 'sign/' into PDFs in 'in/',\n"
            "saving results to 'out/'.\n\n"
            "Requires: pip install pymupdf"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--in-dir",
        default=str(script_dir / "in"),
        help="Folder with input PDFs (default: in/ next to script)",
    )
    parser.add_argument(
        "--sign-dir",
        default=str(script_dir / "sign"),
        help="Folder with sign PDFs (default: sign/ next to script)",
    )
    parser.add_argument(
        "--out-dir",
        default=str(script_dir / "out"),
        help="Output folder (default: out/ next to script)",
    )
    args = parser.parse_args()

    in_dir = Path(args.in_dir)
    sign_dir = Path(args.sign_dir)
    out_dir = Path(args.out_dir)

    # Step 1: Ensure required directories exist
    ensure_dirs(in_dir, sign_dir, out_dir)

    # Step 2 & 3: Validate that folders are not empty
    check_non_empty(in_dir, "in")
    check_non_empty(sign_dir, "sign")

    # Load and pre-render all sign documents
    sign_paths = sorted(sign_dir.glob("*.pdf"))
    print(f"Found {len(sign_paths)} sign document(s): {[p.name for p in sign_paths]}")

    sign_docs: list[pymupdf.Document] = []
    sign_pixmaps: list[tuple[pymupdf.Rect, pymupdf.Pixmap]] = []

    for sp in sign_paths:
        sdoc = pymupdf.open(str(sp))
        if sdoc.page_count == 0:
            print(f"  Warning: sign document '{sp.name}' has no pages — skipping.")
            continue
        rect, pix = get_sign_rect_on_page(sdoc)
        sign_docs.append(sdoc)
        sign_pixmaps.append((rect, pix))

    if not sign_docs:
        sys.exit(
            "Warning: No valid sign documents found in the 'sign' folder.\n"
            "Please add single-page PDF documents with images to the 'sign' folder."
        )

    # Step 4–6: Process each input PDF
    in_paths = sorted(in_dir.glob("*.pdf"))
    print(f"\nFound {len(in_paths)} input document(s): {[p.name for p in in_paths]}")
    print()

    for in_path in in_paths:
        out_path = out_dir / in_path.name
        try:
            process_pdf(in_path, sign_docs, sign_pixmaps, out_path)
        except Exception as exc:
            print(f"  Error processing '{in_path.name}': {exc}")

    # Cleanup
    for sdoc in sign_docs:
        sdoc.close()

    print("\nDone.")


if __name__ == "__main__":
    main()
