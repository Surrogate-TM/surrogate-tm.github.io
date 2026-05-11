#!/usr/bin/env python3
"""
Extract objects from a named layer in a single-page PDF and export to DWG/DXF.

Usage:
    python3 pdf_layer_to_dwg.py input.pdf [--layer LAYER_NAME] [--output output.dwg]

Dependencies:
    pip install pymupdf ezdxf

DWG export also requires the free ODA File Converter:
    https://www.opendesign.com/guestfiles/oda_file_converter
Without it, the script outputs a DXF file (AutoCAD-compatible).

Algorithm:
  1. Open the PDF and identify all Optional Content Groups (OCGs / layers).
  2. Turn ON only the target layer, turn OFF all others.
  3. Extract vector paths via page.get_drawings() (respects layer visibility).
  4. Extract text via page.get_text() (respects layer visibility).
  5. Convert PDF coordinates to DXF coordinates (flip Y axis: PDF top-left -> DXF bottom-left).
  6. Reconstruct geometry in ezdxf: lines, polylines, Bezier splines, text.
  7. Save as DXF; optionally convert to DWG via ODA File Converter.
"""

import argparse
import sys
import math
from pathlib import Path

try:
    import pymupdf
except ImportError:
    sys.exit("Error: PyMuPDF not installed. Run: pip install pymupdf")

try:
    import ezdxf
    from ezdxf.enums import TextEntityAlignment
    from ezdxf.math import Vec2
except ImportError:
    sys.exit("Error: ezdxf not installed. Run: pip install ezdxf")


# PDF uses points (1 pt = 1/72 inch). DXF model space typically uses mm.
# 1 pt = 25.4/72 mm ≈ 0.3528 mm. Set SCALE = 1.0 to keep native pt units.
SCALE = 1.0


def pdf_to_dxf_y(y_pdf: float, page_height: float) -> float:
    """Flip Y: PDF origin is top-left in MuPDF viewport; DXF origin is bottom-left."""
    return (page_height - y_pdf) * SCALE


def pdf_to_dxf_xy(x: float, y: float, page_height: float) -> tuple:
    return x * SCALE, pdf_to_dxf_y(y, page_height)


def get_layer_names(doc: pymupdf.Document) -> list:
    """Return list of layer names in the PDF."""
    return [layer["text"] for layer in doc.layer_ui_configs()]


def activate_single_layer(doc: pymupdf.Document, target_name: str) -> bool:
    """Turn on only the target layer, turn off all others. Returns True if found."""
    found = False
    for layer in doc.layer_ui_configs():
        if layer["text"] == target_name:
            doc.set_layer_ui_config(layer["number"], action=0)  # ON
            found = True
        else:
            doc.set_layer_ui_config(layer["number"], action=2)  # OFF
    return found


def bezier_to_polyline(p0, p1, p2, p3, steps: int = 16) -> list:
    """Approximate a cubic Bezier curve as a sequence of (x, y) points."""
    pts = []
    for i in range(steps + 1):
        t = i / steps
        u = 1 - t
        x = u**3 * p0[0] + 3*u**2*t * p1[0] + 3*u*t**2 * p2[0] + t**3 * p3[0]
        y = u**3 * p0[1] + 3*u**2*t * p1[1] + 3*u*t**2 * p2[1] + t**3 * p3[1]
        pts.append((x, y))
    return pts


def add_path_to_dxf(msp, drawing: dict, page_height: float) -> None:
    """Convert a single path dict from get_drawings() into DXF entities."""
    items = drawing.get("items", [])
    close_path = drawing.get("closePath", False)
    has_fill = drawing.get("fill") is not None

    # Collect segments; flush each contiguous run of beziers as a polyline.
    line_pts = []  # accumulated polyline points for lines/beziers

    def flush_polyline(pts: list, close: bool = False) -> None:
        if len(pts) >= 2:
            msp.add_lwpolyline(pts, close=close)

    for item in items:
        cmd = item[0]

        if cmd == "l":  # straight line segment
            p1 = item[1]
            p2 = item[2]
            x0, y0 = pdf_to_dxf_xy(p1.x, p1.y, page_height)
            x1, y1 = pdf_to_dxf_xy(p2.x, p2.y, page_height)
            if not line_pts:
                line_pts.append((x0, y0))
            line_pts.append((x1, y1))

        elif cmd == "c":  # cubic Bezier
            p0, cp1, cp2, p3 = item[1], item[2], item[3], item[4]
            bp0 = pdf_to_dxf_xy(p0.x, p0.y, page_height)
            bp1 = pdf_to_dxf_xy(cp1.x, cp1.y, page_height)
            bp2 = pdf_to_dxf_xy(cp2.x, cp2.y, page_height)
            bp3 = pdf_to_dxf_xy(p3.x, p3.y, page_height)
            seg = bezier_to_polyline(bp0, bp1, bp2, bp3)
            if line_pts:
                # Merge: skip duplicate start point
                line_pts.extend(seg[1:])
            else:
                line_pts.extend(seg)

        elif cmd == "re":  # rectangle
            flush_polyline(line_pts, close=close_path)
            line_pts = []
            rect = item[1]
            x0, y0 = pdf_to_dxf_xy(rect.x0, rect.y1, page_height)
            x1, y1 = pdf_to_dxf_xy(rect.x1, rect.y0, page_height)
            msp.add_lwpolyline(
                [(x0, y0), (x1, y0), (x1, y1), (x0, y1)], close=True
            )

        elif cmd == "qu":  # quad (4-corner polygon)
            flush_polyline(line_pts, close=close_path)
            line_pts = []
            quad = item[1]
            pts = [
                pdf_to_dxf_xy(quad.ul.x, quad.ul.y, page_height),
                pdf_to_dxf_xy(quad.ur.x, quad.ur.y, page_height),
                pdf_to_dxf_xy(quad.lr.x, quad.lr.y, page_height),
                pdf_to_dxf_xy(quad.ll.x, quad.ll.y, page_height),
            ]
            msp.add_lwpolyline(pts, close=True)

    flush_polyline(line_pts, close=close_path or has_fill)


def add_text_to_dxf(msp, page: pymupdf.Page, page_height: float) -> None:
    """Extract text from the visible layer and add TEXT entities to DXF."""
    text_dict = page.get_text("dict")
    for block in text_dict.get("blocks", []):
        if block.get("type") != 0:  # skip non-text blocks
            continue
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                text = span.get("text", "").strip()
                if not text:
                    continue
                fontsize = span.get("size", 12)
                origin = span.get("origin")  # baseline origin
                if origin:
                    x, y = pdf_to_dxf_xy(origin[0], origin[1], page_height)
                else:
                    bbox = span.get("bbox")
                    x, y = pdf_to_dxf_xy(bbox[0], bbox[3], page_height)

                # DXF text height in drawing units (1 pt ≈ 0.3528 mm)
                dxf_height = fontsize * SCALE
                entity = msp.add_text(
                    text,
                    dxfattribs={"height": dxf_height},
                )
                entity.set_placement((x, y), align=TextEntityAlignment.LEFT)


def add_images_to_dxf(msp, page: pymupdf.Page, doc: pymupdf.Document,
                       page_height: float, output_dir: Path) -> None:
    """Save raster images from the visible layer as PNG files and reference in DXF."""
    images = page.get_images(full=True)
    if not images:
        return

    output_dir.mkdir(parents=True, exist_ok=True)
    for idx, img_info in enumerate(images):
        xref = img_info[0]
        try:
            pix = pymupdf.Pixmap(doc, xref)
            if pix.n > 4:  # CMYK -> RGB
                pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
            img_path = output_dir / f"image_{idx:03d}.png"
            pix.save(str(img_path))

            # Find image placement on page via get_image_bbox
            try:
                bbox = page.get_image_bbox(img_info)
                x0, y0 = pdf_to_dxf_xy(bbox.x0, bbox.y1, page_height)
                x1, y1 = pdf_to_dxf_xy(bbox.x1, bbox.y0, page_height)
                w = (x1 - x0) * SCALE
                h = (y1 - y0) * SCALE
                msp.add_image_def_reactor  # just ensure available
                # Add image reference (requires image definition)
                img_def = msp.doc.add_image_def(str(img_path), size_in_pixel=(pix.width, pix.height))
                msp.add_image(
                    insert=(x0, y0),
                    size_in_units=(w, h),
                    image_def=img_def,
                )
            except Exception:
                pass  # image placement not critical
        except Exception:
            pass


def convert_to_dwg(dxf_doc: ezdxf.document.Drawing, output_path: Path) -> bool:
    """Try to export DXF as DWG using ODA File Converter. Returns True on success."""
    try:
        from ezdxf.addons import odafc
        dwg_path = output_path.with_suffix(".dwg")
        odafc.export_dwg(dxf_doc, str(dwg_path), replace=True)
        return True
    except Exception as exc:
        print(f"Warning: DWG conversion failed ({exc}).")
        print("Install ODA File Converter from https://www.opendesign.com/guestfiles/oda_file_converter")
        return False


def extract_layer_to_dxf(
    pdf_path: str,
    layer_name: str,
    output_path: str,
) -> Path:
    """
    Main entry point. Returns the path to the saved output file.
    Outputs DXF; also attempts DWG conversion via ODA File Converter.
    """
    pdf_path = Path(pdf_path)
    output_path = Path(output_path)

    if not pdf_path.exists():
        sys.exit(f"Error: PDF file not found: {pdf_path}")

    doc = pymupdf.open(str(pdf_path))

    if doc.page_count == 0:
        sys.exit("Error: PDF has no pages.")

    # List available layers
    available_layers = get_layer_names(doc)
    if not available_layers:
        sys.exit("Error: The PDF has no layers (Optional Content Groups).")

    print(f"PDF layers found: {available_layers}")

    if not activate_single_layer(doc, layer_name):
        sys.exit(
            f"Error: Layer '{layer_name}' not found.\n"
            f"Available layers: {available_layers}"
        )

    print(f"Extracting objects from layer: '{layer_name}'")

    page = doc[0]
    page_height = page.rect.height

    # Create DXF document (R2010 = AC1024, widely compatible)
    dxf_doc = ezdxf.new("R2010")
    dxf_doc.header["$INSUNITS"] = 0  # unitless (native PDF points)
    msp = dxf_doc.modelspace()

    # Vector paths
    drawings = page.get_drawings()
    print(f"  Vector paths: {len(drawings)}")
    for drawing in drawings:
        add_path_to_dxf(msp, drawing, page_height)

    # Text
    add_text_to_dxf(msp, page, page_height)

    # Images (embedded rasters)
    image_dir = output_path.parent / (output_path.stem + "_images")
    add_images_to_dxf(msp, page, doc, page_height, image_dir)

    # Save DXF
    dxf_path = output_path.with_suffix(".dxf")
    dxf_doc.saveas(str(dxf_path))
    print(f"Saved DXF: {dxf_path}")

    # Attempt DWG conversion
    dwg_path = output_path.with_suffix(".dwg")
    if convert_to_dwg(dxf_doc, dwg_path):
        print(f"Saved DWG: {dwg_path}")
        return dwg_path

    return dxf_path


def main():
    parser = argparse.ArgumentParser(
        description=(
            "Extract objects from a named PDF layer and export to DWG/DXF.\n"
            "Requires: pip install pymupdf ezdxf\n"
            "DWG output requires ODA File Converter (free): "
            "https://www.opendesign.com/guestfiles/oda_file_converter"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("pdf", help="Input PDF file path")
    parser.add_argument(
        "--layer",
        default="light",
        help="Name of the PDF layer to extract (default: 'light')",
    )
    parser.add_argument(
        "--output",
        default=None,
        help="Output file path (default: <pdf_name>_<layer>.dwg)",
    )
    parser.add_argument(
        "--list-layers",
        action="store_true",
        help="List all layers in the PDF and exit",
    )

    args = parser.parse_args()

    if args.list_layers:
        doc = pymupdf.open(args.pdf)
        layers = get_layer_names(doc)
        if layers:
            print("Layers in PDF:")
            for name in layers:
                print(f"  - {name}")
        else:
            print("No layers found in PDF.")
        return

    pdf_path = Path(args.pdf)
    output_path = Path(args.output) if args.output else (
        pdf_path.parent / f"{pdf_path.stem}_{args.layer}.dwg"
    )

    result = extract_layer_to_dxf(str(pdf_path), args.layer, str(output_path))
    print(f"\nDone. Output: {result}")


if __name__ == "__main__":
    main()
