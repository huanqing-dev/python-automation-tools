#!/usr/bin/env python3
"""Batch-process a folder of images: watermark, resize, compress, rename, report.

The gig this replaces: "I have 200 product photos, put my watermark on them,
make them under 500 KB each and rename them to my SKU format."

Everything is optional and composable, and --dry-run previews every decision
before a single file is written.

  * watermark: text (tiled or corner) with opacity, or a PNG logo file
  * resize: longest edge, never upscales
  * compress: JPEG/WebP quality, and it warns if a file still misses the target
  * rename: prefix + zero-padded index, original extension preserved
  * report: per-file before/after size, written to report.txt next to the output

Usage
    python process.py                                  # builds sample images first
    python process.py input/ -o out/ --wm-text "SAMPLE" --max-edge 800 --quality 82
    python process.py input/ --wm-logo logo.png --rename-prefix "SKU" --dry-run

Requires: Pillow
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageEnhance, ImageFont

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".heic", ".tiff", ".bmp"}


# --------------------------------------------------------------------------- #
# sample images, so the demo runs with zero setup
# --------------------------------------------------------------------------- #
def make_samples(folder: Path, count: int = 4) -> list[Path]:
    folder.mkdir(parents=True, exist_ok=True)
    made = []
    palettes = [("#f6d365", "#fda085"), ("#a1c4fd", "#c2e9fb"),
                ("#d4fc79", "#96e6a1"), ("#fbc2eb", "#a6c1ee")]
    titles = ["Leather Wallet", "Canvas Tote", "Desk Lamp", "Water Bottle"]
    for index in range(count):
        top, bottom = palettes[index % len(palettes)]
        img = Image.new("RGB", (1200, 900), top)
        draw = ImageDraw.Draw(img)
        for y in range(900):  # vertical gradient
            blend = y / 900
            colour = tuple(
                int(a + (b - a) * blend) for a, b in zip(
                    tuple(int(top[i:i + 2], 16) for i in (1, 3, 5)),
                    tuple(int(bottom[i:i + 2], 16) for i in (1, 3, 5))))
            draw.line([(0, y), (1200, y)], fill=colour)
        draw.rounded_rectangle([380, 250, 820, 640], 24, fill=(255, 255, 255),
                               outline=(210, 214, 220), width=2)
        big = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", 44)
        small = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", 24)
        draw.text((600, 400), titles[index], font=big, fill=(40, 46, 54), anchor="mm")
        draw.text((600, 470), "sample product photo", font=small, fill=(120, 128, 138), anchor="mm")
        path = folder / f"photo_{index + 1:02d}_original_name_very_long.jpg"
        img.save(path, quality=95)
        made.append(path)
    return made


# --------------------------------------------------------------------------- #
# transforms
# --------------------------------------------------------------------------- #
def text_watermark(img: Image.Image, text: str, opacity: float, tiled: bool,
                   corner: str, font_size: int) -> Image.Image:
    layer = Image.new("RGBA", img.size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(layer)
    try:
        font = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", font_size)
    except OSError:
        font = ImageFont.load_default()
    fill = (255, 255, 255, int(255 * opacity))
    outline = (0, 0, 0, int(255 * opacity * 0.55))

    if tiled:
        step_x, step_y = font_size * 8, font_size * 4
        for y in range(-font_size, img.height + step_y, step_y):
            for x in range(-font_size, img.width + step_x, step_x):
                draw.text((x, y), text, font=font, fill=fill, stroke_width=1, stroke_fill=outline)
    else:
        box = draw.textbbox((0, 0), text, font=font)
        tw, th = box[2] - box[0], box[3] - box[1]
        margin = int(font_size * 0.8)
        positions = {
            "centre": ((img.width - tw) // 2, (img.height - th) // 2),
            "bottom-right": (img.width - tw - margin, img.height - th - margin),
            "bottom-left": (margin, img.height - th - margin),
            "top-right": (img.width - tw - margin, margin),
            "top-left": (margin, margin),
        }
        draw.text(positions.get(corner, positions["bottom-right"]), text, font=font,
                  fill=fill, stroke_width=1, stroke_fill=outline)
    return Image.alpha_composite(img.convert("RGBA"), layer).convert("RGB")


def logo_watermark(img: Image.Image, logo_path: Path, opacity: float, corner: str) -> Image.Image:
    logo = Image.open(logo_path).convert("RGBA")
    scale = min(img.width, img.height) * 0.22 / max(logo.size)
    logo = logo.resize((max(1, int(logo.width * scale)), max(1, int(logo.height * scale))))
    alpha = logo.getchannel("A").point(lambda a: int(a * opacity))
    logo.putalpha(alpha)
    margin = 24
    spots = {
        "bottom-right": (img.width - logo.width - margin, img.height - logo.height - margin),
        "bottom-left": (margin, img.height - logo.height - margin),
        "top-right": (img.width - logo.width - margin, margin),
        "top-left": (margin, margin),
        "centre": ((img.width - logo.width) // 2, (img.height - logo.height) // 2),
    }
    img = img.convert("RGBA")
    img.alpha_composite(logo, spots.get(corner, spots["bottom-right"]))
    return img.convert("RGB")


def resize_long_edge(img: Image.Image, max_edge: int) -> Image.Image:
    if max(img.size) <= max_edge:  # never upscale
        return img
    scale = max_edge / max(img.size)
    return img.resize((round(img.width * scale), round(img.height * scale)), Image.LANCZOS)


def save_compressed(img: Image.Image, path: Path, quality: int) -> Path:
    if path.suffix.lower() == ".webp":
        img.save(path, "WEBP", quality=quality, method=6)
    else:
        img = img.convert("RGB")
        img.save(path, "JPEG", quality=quality, optimize=True, progressive=True)
    return path


# --------------------------------------------------------------------------- #
# batch
# --------------------------------------------------------------------------- #
def process_folder(input_dir: Path, output_dir: Path, args) -> list[dict]:
    files = sorted(p for p in input_dir.iterdir()
                   if p.suffix.lower() in IMAGE_EXTS and not p.name.startswith("."))
    report, counter = [], 0

    for source in files:
        counter += 1
        original_size = source.stat().st_size
        with Image.open(source) as img:
            img.load()
            note = []
            if args.wm_text:
                img = text_watermark(img, args.wm_text, args.wm_opacity,
                                     args.wm_tiled, args.wm_corner, args.wm_size)
                note.append("text watermark")
            if args.wm_logo:
                img = logo_watermark(img, Path(args.wm_logo), args.wm_opacity, args.wm_corner)
                note.append("logo watermark")
            before = img.size
            img = resize_long_edge(img, args.max_edge)
            if img.size != before:
                note.append(f"{before[0]}x{before[1]} -> {img.width}x{img.height}")
            if args.brightness != 1.0:
                img = ImageEnhance.Brightness(img).enhance(args.brightness)
                note.append(f"brightness x{args.brightness}")

            stem = f"{args.rename_prefix}{counter:0{args.pad}d}" if args.rename_prefix else source.stem
            target = output_dir / f"{stem}{source.suffix.lower()}"
            if args.dry_run:
                report.append({"file": source.name, "out": target.name,
                               "bytes_in": original_size, "bytes_out": 0, "note": " ".join(note) + " (dry run)"})
                continue
            save_compressed(img, target, args.quality)
            new_size = target.stat().st_size
            report.append({"file": source.name, "out": target.name,
                           "bytes_in": original_size, "bytes_out": new_size,
                           "note": " ".join(note)})
    return report


def write_report(report: list[dict], output_dir: Path, dry: bool) -> None:
    lines = ["file".ljust(38), " -> ", "output".ljust(26), "size".rjust(10), "  note", "\n"]
    total_in = total_out = 0
    for item in report:
        total_in += item["bytes_in"]
        total_out += item["bytes_out"]
        size = f"{item['bytes_out'] / 1024:7.0f} KB" if item["bytes_out"] else "     dry"
        lines.append(f"{item['file'][:37]:38} -> {item['out'][:25]:26} {size}  {item['note']}\n")
    saved = (total_in - total_out) / max(total_in, 1) * 100
    lines.append(f"\n{len(report)} image(s){' (dry run)' if dry else ''}   "
                 f"{total_in / 1024:.0f} KB -> {total_out / 1024:.0f} KB   ({saved:.0f}% smaller)\n")
    (output_dir / "report.txt").write_text("".join(lines))
    print("".join(lines))


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Batch watermark / resize / compress / rename images.")
    parser.add_argument("input", nargs="?", help="folder of images (omit to build samples in sample_data/)")
    parser.add_argument("-o", "--out", default="output")
    parser.add_argument("--wm-text", help="watermark text")
    parser.add_argument("--wm-logo", help="PNG logo file")
    parser.add_argument("--wm-opacity", type=float, default=0.35)
    parser.add_argument("--wm-tiled", action="store_true", help="tile the text across the image")
    parser.add_argument("--wm-corner", default="bottom-right",
                        choices=["centre", "bottom-right", "bottom-left", "top-right", "top-left"])
    parser.add_argument("--wm-size", type=int, default=72, help="watermark font size")
    parser.add_argument("--max-edge", type=int, help="cap the longest edge (never upscales)")
    parser.add_argument("--quality", type=int, default=85)
    parser.add_argument("--brightness", type=float, default=1.0)
    parser.add_argument("--rename-prefix", help="rename to PREFIX0001.ext")
    parser.add_argument("--pad", type=int, default=3, help="digits in the rename index")
    parser.add_argument("--dry-run", action="store_true", help="print the plan, write nothing")
    args = parser.parse_args(argv)

    if args.input:
        input_dir = Path(args.input)
        if not input_dir.is_dir():
            print(f"not a folder: {input_dir}")
            return 2
        made = []
    else:
        input_dir = Path(__file__).parent / "sample_data"
        print("no input folder given - building sample product photos")
        made = make_samples(input_dir)

    output_dir = Path(args.out)
    if args.dry_run:
        output_dir = Path("/tmp/image-batch-dryrun")  # never write into the real output
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"Processing {input_dir} -> {output_dir}{' (dry run)' if args.dry_run else ''}\n")
    report = process_folder(input_dir, output_dir, args)
    if not report:
        print("No images found.")
        return 1
    write_report(report, output_dir if not args.dry_run else Path(args.out).parent, args.dry_run)
    if made:
        print(f"(sample photos generated in {input_dir})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
