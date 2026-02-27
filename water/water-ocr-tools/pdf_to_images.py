#!/usr/bin/env python3
"""
Convert PDF to high-resolution PNG images.

Usage:
    uv run --with pdf2image,Pillow pdf_to_images.py <input_pdf> [output_dir]

Requirements:
    - pdf2image
    - Pillow
    - poppler (system dependency: brew install poppler)
"""

import argparse
import json
import sys
from pathlib import Path
from pdf2image import convert_from_path


def pdf_to_images(pdf_path: str, output_dir: str = "./output_images", dpi: int = 300) -> list[dict]:
    """
    Convert PDF to PNG images at specified DPI.

    Args:
        pdf_path: Path to input PDF file
        output_dir: Directory to save output images
        dpi: Resolution in dots per inch (default: 300)

    Returns:
        List of dicts containing image metadata
    """
    pdf_file = Path(pdf_path)
    if not pdf_file.exists():
        raise FileNotFoundError(f"PDF file not found: {pdf_path}")

    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    print(f"Converting PDF: {pdf_file}")
    print(f"Output directory: {output_path}")
    print(f"DPI: {dpi}")
    print()

    # Convert PDF to images
    images = convert_from_path(pdf_file, dpi=dpi)

    metadata = []

    for i, image in enumerate(images, start=1):
        # Generate output filename
        output_filename = f"page_{i}.png"
        output_file = output_path / output_filename

        # Save image
        image.save(output_file, "PNG")

        # Collect metadata
        file_size = output_file.stat().st_size
        width, height = image.size

        meta = {
            "page": i,
            "filename": output_filename,
            "path": str(output_file.absolute()),
            "width": width,
            "height": height,
            "size_bytes": file_size,
            "size_mb": round(file_size / (1024 * 1024), 2),
            "dpi": dpi
        }

        metadata.append(meta)

        # Print to stdout
        print(f"Created: {output_file}")
        print(f"  Size: {width}x{height}px, {meta['size_mb']}MB")

    # Write metadata JSON
    metadata_file = output_path / "images_metadata.json"
    with open(metadata_file, "w") as f:
        json.dump({
            "source_pdf": str(pdf_file.absolute()),
            "total_pages": len(images),
            "dpi": dpi,
            "images": metadata
        }, f, indent=2)

    print()
    print(f"Metadata saved to: {metadata_file}")
    print(f"Total pages converted: {len(images)}")

    return metadata


def main():
    parser = argparse.ArgumentParser(
        description="Convert PDF to high-resolution PNG images",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  uv run --with pdf2image,Pillow pdf_to_images.py document.pdf
  uv run --with pdf2image,Pillow pdf_to_images.py document.pdf ./my_images
  uv run --with pdf2image,Pillow pdf_to_images.py document.pdf --dpi 600
        """
    )

    parser.add_argument(
        "input_pdf",
        help="Path to input PDF file"
    )

    parser.add_argument(
        "output_dir",
        nargs="?",
        default="./output_images",
        help="Output directory for images (default: ./output_images)"
    )

    parser.add_argument(
        "--dpi",
        type=int,
        default=300,
        help="Resolution in DPI (default: 300)"
    )

    args = parser.parse_args()

    try:
        pdf_to_images(args.input_pdf, args.output_dir, args.dpi)
    except FileNotFoundError as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Error converting PDF: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
