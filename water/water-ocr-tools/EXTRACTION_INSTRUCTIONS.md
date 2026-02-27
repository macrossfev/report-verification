# Water Quality Report - Data Extraction Task

## Objective
Extract ALL numerical and textual data from the provided water quality report images into structured JSON format.

## Source Material
You will receive 6 page images from a Chinese water quality monitoring report (国家城市供水水质监测网重庆监测站).

The images are located at:
- `/Users/Shared/projects/lockin/tg-mesh-proxy/water-ocr-tools/output_images/page_1.png` - Sample registration table (样品登记表)
- `/Users/Shared/projects/lockin/tg-mesh-proxy/water-ocr-tools/output_images/page_2.png` - Sampling records (采样原始记录)
- `/Users/Shared/projects/lockin/tg-mesh-proxy/water-ocr-tools/output_images/page_3.png` - Detection results summary (检测结果汇总表)
- `/Users/Shared/projects/lockin/tg-mesh-proxy/water-ocr-tools/output_images/page_4.png` - Chemical analysis results (分析结果汇总表)
- `/Users/Shared/projects/lockin/tg-mesh-proxy/water-ocr-tools/output_images/page_5.png` - Trihalomethanes results (三卤甲烷检测)
- `/Users/Shared/projects/lockin/tg-mesh-proxy/water-ocr-tools/output_images/page_6.png` - Radioactivity results (放射性检测)

## Required Output Format

Output a single JSON file at `/Users/Shared/projects/lockin/tg-mesh-proxy/water-ocr-tools/extracted_results.json` with this structure:

```json
{
  "detection_results": {
    "parameter_name(unit)": {
      "SAMPLE_ID": value_or_string
    }
  },
  "chemical_analysis": {
    "parameter_name(unit)": {
      "SAMPLE_ID": value_or_string
    }
  },
  "trihalomethanes": {
    "parameter_name(unit)": {
      "SAMPLE_ID": value_or_string
    }
  },
  "radioactivity": {
    "parameter_name(unit)": {
      "SAMPLE_ID": value_or_string
    }
  }
}
```

## Critical Rules

1. **Sample IDs** are in format: `W260127C05`, `W260127C06`, ..., `W260127C11`, `K260127C12`
2. **Numeric values**: Use numbers (e.g., `7.15`, `0.112`, `150`)
3. **Below detection limit**: Use strings with `<` prefix (e.g., `"<0.010"`, `"<0.0034"`)
4. **Uncertainty values**: Use strings with `±` (e.g., `"0.046±0.009"`)
5. **Text values**: Use strings (e.g., `"<5"` for 色度)
6. **Empty cells / diagonal lines**: Do NOT include these - only include cells with actual values
7. **Preserve exact precision**: If the document says `0.070`, output `0.070` not `0.07`. If it says `2.40` output `2.40` not `2.4`

## CRITICAL: Table Column Layout

### Understanding Diagonal Lines (斜线)
These tables use diagonal lines drawn through cells to indicate "not tested / not applicable". A diagonal line means NO VALUE — skip that cell entirely. Do NOT confuse a diagonal line cell with an adjacent cell's value.

### Page 3 Column Order (Detection Results - 供水检测部检测结果汇总表)
The header row lists sample IDs left to right in this EXACT order:

| Col 1 (项目/Parameter) | Col 2 | Col 3 | Col 4 | Col 5 | Col 6 | Col 7 | Col 8 | Col 9 |
|---|---|---|---|---|---|---|---|---|
| Parameter name | W260127C05 | W260127C06 | W260127C07 | W260127C08 | W260127C09 | W260127C10 | W260127C11 | K260127C12 |

**Key notes for Page 3:**
- C05-C08 are 出厂水/管网水 (treated/pipe water) — they have values for: 水温, pH, 色度, 高锰酸盐指数, 菌落总数, 总大肠菌群, 大肠埃希氏菌, 电导率
- C09 is a PARTIAL test sample — only has: 菌落总数, 总大肠菌群, 大肠埃希氏菌
- C10, C11 are 原水 (raw/source water) — they have: 水温, pH, 高锰酸盐指数, 溶解氧, 化学需氧量, 五日生化需氧量, 粪大肠菌群. NOTE: C10 and C11 are in SEPARATE columns. Do not merge or shift them
- K260127C12 is blank sample — only has: 菌落总数, 总大肠菌群, 大肠埃希氏菌
- Many cells between C09 and C12 contain diagonal lines — these are NOT values

### Page 4 Column Order (Chemical Analysis - 供水检测部分析结果汇总表)
The header row lists sample IDs. NOTE: This table has a slightly different layout:

| Col 1-2 (项目/Parameter) | Col 3 | Col 4 | Col 5 | Col 6 | Col 7 | Col 8 | Col 9 |
|---|---|---|---|---|---|---|---|
| Parameter name | W260127C05 | W260127C06 | W260127C07 | W260127C08 | W260127C09 | W260127C10 | W260127C11 |

**Key notes for Page 4:**
- C05-C08 are treated/pipe water — have most parameters (氟化物 through 溶解性总固体)
- C09 has LIMITED parameters: only 铜, 铁, 锰, 砷, 锌, 硒, 汞, 镉, 铅, 铝, 钙, 镁
- C10 has its OWN column (separate from C11): 氟化物, 氯化物, 硝酸盐, 硫酸盐, 铜, 铁, 锰, 砷, 锌, 硒, 汞, 镉, 铅, 阴离子合成洗涤剂, 挥发酚, 氰化物, 硫化物, 六价铬, 石油类, 总磷, 总氮, 氨
- C11 has its OWN column (rightmost data column): same parameters as C10
- **WATCH OUT**: C10 and C11 values are DIFFERENT numbers. If you see the same value in both, you likely misread the column boundary. Carefully trace each column from header to value
- Diagonal lines separate tested from untested parameters per sample

### Page 5 Column Order (Trihalomethanes)
Rows are samples, columns are parameters:

| 样品编号 | 三氯甲烷 | 四氯化碳 | 二氯一溴甲烷 | 一氯二溴甲烷 | 三溴甲烷 | 三卤甲烷(总量) |
|---|---|---|---|---|---|---|
| W260127C05 | ... | ... | ... | ... | ... | ... |
| W260127C06 | ... | ... | ... | ... | ... | ... |
| W260127C07 | ... | ... | ... | ... | ... | ... |
| W260127C08 | ... | ... | ... | ... | ... | ... |
| K260127C12 | ... | ... | ... | ... | ... | ... |

### Page 6 Column Order (Radioactivity)
Simple 2-column data table:

| 样品编号 | 总α(Bq/L) | 总β(Bq/L) |
|---|---|---|
| W260127C05 | value±uncertainty | value±uncertainty |
| W260127C06 | ... | ... |
| W260127C07 | ... | ... |
| W260127C08 | ... | ... |

## Extraction Approach

1. **Read page_1.png first** to identify all sample IDs and their types (出厂水/管网水/原水/空白样)
2. For each data page (3, 4, 5, 6):
   a. **First, read the column headers** to establish the exact left-to-right column mapping
   b. **For each row**, trace vertically from the header to ensure correct column alignment
   c. Diagonal lines = skip. Actual numbers/text = extract
   d. **Double-check C10 vs C11**: These are adjacent columns for raw water samples. They should have DIFFERENT values for most parameters
3. Write the JSON output file
4. **Self-check**: Verify that C10 and C11 don't have identical values for any parameter (they are different water sources: 清水塘水库 vs 天宝寺水库)
