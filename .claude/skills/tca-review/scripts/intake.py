#!/usr/bin/env python3
"""Convert supplied TCA screenshots to readable PNGs.

Handles HEIC/HEIF from a phone camera as well as ordinary screenshots. Applies
EXIF rotation, upscales small images, and optionally writes quadrant crops so a
dense table can be read a piece at a time.

    python intake.py --in reviews/x/images --out reviews/x/images/png --tile
"""
from __future__ import annotations

import argparse
import csv
import sys
from pathlib import Path

from PIL import Image, ImageOps

try:
    from pillow_heif import register_heif_opener

    register_heif_opener()
    HEIF = True
except Exception:  # pragma: no cover - depends on environment
    HEIF = False

SUFFIXES = {".heic", ".heif", ".jpg", ".jpeg", ".png", ".webp", ".tif", ".tiff"}
# A phone photo is far larger than anything that helps read a table, and a 20MB
# PNG costs minutes to write for no extra legibility. Cap both views instead:
# the full view is for the shape of the table, the tiles carry the digits.
FULL_EDGE = 2200
TILE_EDGE = 2400
TILE_MIN_EDGE = 1600  # only tile images with real detail to recover


def fit(im: Image.Image, edge: int) -> Image.Image:
    """Scale so the long edge is `edge`, up or down. LANCZOS both ways."""
    scale = edge / max(im.size)
    if 0.98 < scale < 1.02:
        return im
    return im.resize((max(1, round(im.width * scale)), max(1, round(im.height * scale))),
                     Image.LANCZOS)


def convert(src: Path, out_dir: Path, tile: bool) -> list[dict]:
    with Image.open(src) as im:
        im = ImageOps.exif_transpose(im).convert("RGB")

        rows = []
        full = fit(im, FULL_EDGE)
        full_path = out_dir / f"{src.stem}.png"
        full.save(full_path, "PNG", compress_level=1)
        rows.append(
            {
                "png": full_path.name,
                "source": src.name,
                "kind": "full",
                "width": full.width,
                "height": full.height,
            }
        )

        if tile and max(im.size) >= TILE_MIN_EDGE:
            W, H = im.size
            # overlapping quadrants: a row split across a seam stays whole in one tile
            ox, oy = int(W * 0.06), int(H * 0.06)
            boxes = {
                "tl": (0, 0, W // 2 + ox, H // 2 + oy),
                "tr": (W // 2 - ox, 0, W, H // 2 + oy),
                "bl": (0, H // 2 - oy, W // 2 + ox, H),
                "br": (W // 2 - ox, H // 2 - oy, W, H),
            }
            for name, box in boxes.items():
                tile_img = fit(im.crop(box), TILE_EDGE)
                path = out_dir / f"{src.stem}__{name}.png"
                tile_img.save(path, "PNG", compress_level=1)
                rows.append(
                    {
                        "png": path.name,
                        "source": src.name,
                        "kind": f"tile-{name}",
                        "width": tile_img.width,
                        "height": tile_img.height,
                    }
                )
        return rows


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--in", dest="src", required=True, help="folder of supplied images")
    ap.add_argument("--out", dest="dst", required=True, help="folder for PNG output")
    ap.add_argument("--tile", action="store_true", help="also write quadrant crops")
    args = ap.parse_args()

    src_dir, out_dir = Path(args.src), Path(args.dst)
    if not src_dir.is_dir():
        print(f"ERROR: no such folder: {src_dir}", file=sys.stderr)
        return 2
    out_dir.mkdir(parents=True, exist_ok=True)

    files = sorted(p for p in src_dir.iterdir() if p.suffix.lower() in SUFFIXES)
    if not files:
        print(f"ERROR: no images in {src_dir}", file=sys.stderr)
        print(f"       looked for: {', '.join(sorted(SUFFIXES))}", file=sys.stderr)
        return 2

    heic = [p for p in files if p.suffix.lower() in {".heic", ".heif"}]
    if heic and not HEIF:
        print("ERROR: HEIC files present but pillow-heif is not installed.", file=sys.stderr)
        print("       pip install pillow-heif", file=sys.stderr)
        return 2

    rows: list[dict] = []
    for path in files:
        if path.parent.resolve() == out_dir.resolve():
            continue
        try:
            rows.extend(convert(path, out_dir, args.tile))
        except Exception as exc:
            print(f"  FAILED {path.name}: {exc}", file=sys.stderr)

    inv = out_dir / "inventory.csv"
    with inv.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=["png", "source", "kind", "width", "height"])
        writer.writeheader()
        writer.writerows(rows)

    fulls = sum(1 for r in rows if r["kind"] == "full")
    print(f"converted {fulls} image(s) -> {len(rows)} PNG(s) in {out_dir}")
    print(f"inventory: {inv}")
    print("\nRead every 'full' PNG first for the shape of each table,")
    print("then the tiles for the digits. Unreadable cell -> NA, never a guess.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
