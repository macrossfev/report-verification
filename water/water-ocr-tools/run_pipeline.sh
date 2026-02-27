#!/usr/bin/env bash
set -euo pipefail

# Water Quality Report OCR Pipeline
# Usage: ./run_pipeline.sh <pdf_file> <excel_ground_truth> [output_dir]

PDF_FILE="${1:?Usage: ./run_pipeline.sh <pdf> <excel> [output_dir]}"
EXCEL_FILE="${2:?Usage: ./run_pipeline.sh <pdf> <excel> [output_dir]}"
OUTPUT_DIR="${3:-./pipeline_output}"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

echo "=== Water Quality OCR Pipeline ==="
echo "PDF:    $PDF_FILE"
echo "Excel:  $EXCEL_FILE"
echo "Output: $OUTPUT_DIR"
echo

# Step 1: Convert PDF to images
echo "--- Step 1: PDF → Images (300 DPI) ---"
uv run --with pdf2image,Pillow python "$SCRIPT_DIR/pdf_to_images.py" "$PDF_FILE" "$OUTPUT_DIR/images"
echo

# Step 2: Extract ground truth from Excel
echo "--- Step 2: Excel → Ground Truth JSON ---"
uv run --with openpyxl python "$SCRIPT_DIR/parse_excel.py" "$EXCEL_FILE" > "$OUTPUT_DIR/ground_truth.json"
echo "Ground truth saved to: $OUTPUT_DIR/ground_truth.json"
echo

# Step 3: Prompt user to run extraction
echo "--- Step 3: Run Extraction ---"
echo "Images are ready at: $OUTPUT_DIR/images/"
echo "Extraction instructions: $SCRIPT_DIR/EXTRACTION_INSTRUCTIONS.md"
echo
echo "To run blind extraction with a Claude agent:"
echo "  Give the agent EXTRACTION_INSTRUCTIONS.md and point it to $OUTPUT_DIR/images/"
echo "  Agent should output: $OUTPUT_DIR/extracted_results.json"
echo
echo "Once extracted_results.json exists, run comparison:"
echo "  uv run python $SCRIPT_DIR/compare_results.py $OUTPUT_DIR/extracted_results.json $OUTPUT_DIR/ground_truth.json"
