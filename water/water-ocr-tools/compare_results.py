#!/usr/bin/env python3
"""
Compare extracted OCR results against ground truth.

Usage:
    uv run python compare_results.py <extracted.json> <ground_truth.json>

Both files should follow the same structure as ground_truth_structured.json.
Outputs a detailed comparison report with accuracy scoring.
"""

import sys
import json
import re
from typing import Any


def normalize_value(val: Any) -> str:
    """Normalize a value for comparison."""
    if val is None:
        return ""
    s = str(val).strip()
    # Normalize whitespace
    s = re.sub(r'\s+', ' ', s)
    # Normalize less-than symbols
    s = s.replace('＜', '<').replace('≤', '<=')
    # Remove trailing zeros after decimal point for numeric comparison
    try:
        f = float(s)
        # Keep original string format for "<" values
        return s
    except (ValueError, TypeError):
        pass
    return s


def values_match(extracted: Any, truth: Any) -> tuple[bool, str]:
    """
    Compare extracted value against ground truth.
    Returns (match: bool, detail: str).
    """
    e = normalize_value(extracted)
    t = normalize_value(truth)

    if not e and not t:
        return True, "both_empty"

    if not e:
        return False, "missing_in_extracted"

    if not t:
        return False, "missing_in_truth"

    # Exact string match
    if e == t:
        return True, "exact_match"

    # Both are "<X" format
    if e.startswith('<') and t.startswith('<'):
        try:
            ev = float(e[1:])
            tv = float(t[1:])
            if abs(ev - tv) < 1e-10:
                return True, "threshold_match"
        except (ValueError, TypeError):
            pass

    # Numeric comparison with tolerance
    try:
        ev = float(e)
        tv = float(t)
        if tv == 0:
            if ev == 0:
                return True, "zero_match"
            return False, f"expected=0, got={ev}"
        rel_error = abs(ev - tv) / abs(tv)
        if rel_error < 0.01:  # 1% tolerance
            return True, "numeric_close"
        elif rel_error < 0.05:  # 5% tolerance
            return True, "numeric_approx"
        else:
            return False, f"numeric_mismatch(err={rel_error:.1%})"
    except (ValueError, TypeError):
        pass

    # String similarity for ± values (radioactivity)
    if '±' in t:
        # Extract central value and uncertainty
        t_parts = t.split('±')
        e_parts = e.split('±') if '±' in e else [e]
        try:
            t_central = float(t_parts[0])
            e_central = float(e_parts[0])
            if abs(e_central - t_central) < 0.001:
                if len(e_parts) > 1 and len(t_parts) > 1:
                    t_unc = float(t_parts[1])
                    e_unc = float(e_parts[1])
                    if abs(e_unc - t_unc) < 0.001:
                        return True, "uncertainty_match"
                    return False, f"uncertainty_mismatch({e_unc} vs {t_unc})"
                return True, "central_value_match"
        except (ValueError, TypeError):
            pass

    return False, f"mismatch('{e}' vs '{t}')"


def compare_section(extracted: dict, truth: dict) -> dict:
    """Compare a section (detection_results, chemical_analysis, etc.)."""
    results = {
        "total_values": 0,
        "matched": 0,
        "mismatched": 0,
        "missing": 0,
        "extra": 0,
        "details": {}
    }

    # Get all parameter names (skip _ prefixed metadata)
    truth_params = {k for k in truth if not k.startswith('_')}
    extracted_params = {k for k in extracted if not k.startswith('_')}

    all_params = truth_params | extracted_params

    for param in sorted(all_params):
        param_result = {"values": {}}

        if param not in truth:
            param_result["status"] = "extra_parameter"
            results["extra"] += 1
            results["details"][param] = param_result
            continue

        if param not in extracted:
            truth_count = len([v for v in truth[param].values() if v is not None])
            param_result["status"] = "missing_parameter"
            param_result["missing_count"] = truth_count
            results["missing"] += truth_count
            results["total_values"] += truth_count
            results["details"][param] = param_result
            continue

        truth_samples = truth[param] if isinstance(truth[param], dict) else {}
        extracted_samples = extracted[param] if isinstance(extracted[param], dict) else {}

        all_samples = set(truth_samples.keys()) | set(extracted_samples.keys())

        for sample in sorted(all_samples):
            t_val = truth_samples.get(sample)
            e_val = extracted_samples.get(sample)

            if t_val is None:
                if e_val is not None:
                    results["extra"] += 1
                    param_result["values"][sample] = {
                        "status": "extra",
                        "extracted": e_val
                    }
                continue

            results["total_values"] += 1

            if e_val is None:
                results["missing"] += 1
                param_result["values"][sample] = {
                    "status": "missing",
                    "truth": t_val
                }
                continue

            match, detail = values_match(e_val, t_val)
            if match:
                results["matched"] += 1
                param_result["values"][sample] = {
                    "status": "correct",
                    "detail": detail,
                    "value": str(t_val)
                }
            else:
                results["mismatched"] += 1
                param_result["values"][sample] = {
                    "status": "wrong",
                    "detail": detail,
                    "extracted": str(e_val),
                    "truth": str(t_val)
                }

        results["details"][param] = param_result

    return results


def run_comparison(extracted_path: str, truth_path: str) -> dict:
    """Run full comparison between extracted and ground truth."""
    with open(extracted_path) as f:
        extracted = json.load(f)
    with open(truth_path) as f:
        truth = json.load(f)

    sections = ["detection_results", "chemical_analysis", "trihalomethanes", "radioactivity"]

    report = {
        "overall": {"total": 0, "matched": 0, "mismatched": 0, "missing": 0},
        "sections": {}
    }

    for section in sections:
        t_section = truth.get(section, {})
        e_section = extracted.get(section, {})

        if not t_section:
            report["sections"][section] = {"status": "no_ground_truth"}
            continue

        result = compare_section(e_section, t_section)
        report["sections"][section] = result
        report["overall"]["total"] += result["total_values"]
        report["overall"]["matched"] += result["matched"]
        report["overall"]["mismatched"] += result["mismatched"]
        report["overall"]["missing"] += result["missing"]

    # Calculate accuracy
    total = report["overall"]["total"]
    if total > 0:
        report["overall"]["accuracy"] = round(report["overall"]["matched"] / total * 100, 1)
        report["overall"]["error_rate"] = round(report["overall"]["mismatched"] / total * 100, 1)
        report["overall"]["missing_rate"] = round(report["overall"]["missing"] / total * 100, 1)
    else:
        report["overall"]["accuracy"] = 0
        report["overall"]["error_rate"] = 0
        report["overall"]["missing_rate"] = 0

    return report


def print_report(report: dict):
    """Print human-readable comparison report."""
    overall = report["overall"]
    print("=" * 70)
    print("WATER QUALITY OCR EXTRACTION ACCURACY REPORT")
    print("=" * 70)
    print()
    print(f"Overall Accuracy: {overall['accuracy']}%")
    print(f"  Total values:  {overall['total']}")
    print(f"  Matched:       {overall['matched']}")
    print(f"  Mismatched:    {overall['mismatched']} ({overall['error_rate']}%)")
    print(f"  Missing:       {overall['missing']} ({overall['missing_rate']}%)")
    print()

    for section_name, section in report["sections"].items():
        if section.get("status") == "no_ground_truth":
            continue

        total = section["total_values"]
        matched = section["matched"]
        acc = round(matched / total * 100, 1) if total > 0 else 0

        print(f"--- {section_name} ({acc}% accuracy, {matched}/{total}) ---")

        for param, detail in section.get("details", {}).items():
            if detail.get("status") == "missing_parameter":
                print(f"  ✗ {param}: ENTIRE PARAMETER MISSING ({detail['missing_count']} values)")
                continue

            for sample, val in detail.get("values", {}).items():
                if val["status"] == "wrong":
                    print(f"  ✗ {param} [{sample}]: extracted='{val['extracted']}' truth='{val['truth']}' ({val['detail']})")
                elif val["status"] == "missing":
                    print(f"  ○ {param} [{sample}]: MISSING (truth='{val['truth']}')")

        print()


def main():
    if len(sys.argv) < 3:
        print("Usage: uv run python compare_results.py <extracted.json> <ground_truth.json>")
        print()
        print("Both files should use the ground_truth_structured.json format.")
        sys.exit(1)

    extracted_path = sys.argv[1]
    truth_path = sys.argv[2]

    report = run_comparison(extracted_path, truth_path)
    print_report(report)

    # Also save JSON report
    report_path = extracted_path.replace('.json', '_report.json')
    with open(report_path, 'w') as f:
        json.dump(report, f, indent=2, ensure_ascii=False)
    print(f"\nDetailed JSON report saved to: {report_path}")

    # Return exit code based on accuracy
    accuracy = report["overall"]["accuracy"]
    if accuracy >= 95:
        print(f"\n✓ PASS: {accuracy}% accuracy")
        sys.exit(0)
    elif accuracy >= 80:
        print(f"\n⚠ ACCEPTABLE: {accuracy}% accuracy")
        sys.exit(0)
    else:
        print(f"\n✗ FAIL: {accuracy}% accuracy")
        sys.exit(1)


if __name__ == "__main__":
    main()
