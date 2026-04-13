#!/usr/bin/env python3
"""
Eureka Forbes P&L Analyzer — Main Entry Point

Usage:
    # Generate sample data and PPT:
    python main.py --sample

    # Analyze an existing Excel file:
    python main.py --input path/to/pnl.xlsx --month Mar --fy 2026

    # Analyze with auto-detected month (from Metadata sheet):
    python main.py --input path/to/pnl.xlsx
"""
import argparse
import os
import sys

from sample_data import generate_sample_excel
from analyzer import analyze_pnl
from ppt_generator import generate_ppt


def main():
    parser = argparse.ArgumentParser(description="Eureka Forbes P&L Analyzer & PPT Generator")
    parser.add_argument("--input", "-i", help="Path to P&L Excel file")
    parser.add_argument("--output", "-o", help="Output PPT path (default: auto-named)")
    parser.add_argument("--month", "-m", help="Review month (e.g., Mar, Jun, Sep, Dec)")
    parser.add_argument("--fy", type=int, help="Financial year (e.g., 2026)")
    parser.add_argument("--sample", action="store_true", help="Generate sample data first, then analyze")
    args = parser.parse_args()

    # Generate sample if requested
    if args.sample:
        review_month = args.month or "Mar"
        review_fy = args.fy or 2026
        sample_path = os.path.join(os.path.dirname(__file__), "sample_pnl_eureka_forbes.xlsx")
        print(f"Generating sample P&L data for {review_month} FY{review_fy % 100}...")
        generate_sample_excel(sample_path, review_month=review_month, review_fy=review_fy)
        print(f"  -> Sample Excel: {sample_path}")
        args.input = sample_path

    if not args.input:
        print("Error: Please provide --input path to P&L Excel file, or use --sample")
        sys.exit(1)

    if not os.path.exists(args.input):
        print(f"Error: File not found: {args.input}")
        sys.exit(1)

    # Determine output path
    if not args.output:
        base = os.path.splitext(os.path.basename(args.input))[0]
        month = args.month or "review"
        args.output = os.path.join(os.path.dirname(args.input) or ".", f"{base}_analysis_{month}.pptx")

    print(f"\nAnalyzing P&L: {args.input}")
    analysis = analyze_pnl(args.input, review_month=args.month, review_fy=args.fy)

    print(f"\n{'='*60}")
    print(f"  Company: {analysis.company}")
    print(f"  Period:  {analysis.review_month} {analysis.review_fy}")
    print(f"  {'Full Year Review' if analysis.review_month == 'Mar' else 'Monthly Review'}")
    print(f"{'='*60}")

    # Print summary
    ns = analysis.current_month.get("Total Net Sales", 0)
    ebitda = analysis.current_month.get("EBITDA (post allocation)", 0)
    pat = analysis.current_month.get("Profit After Tax", 0)
    print(f"\n  Net Sales:  Rs {ns:,.1f} Cr")
    print(f"  EBITDA:     Rs {ebitda:,.1f} Cr ({analysis.current_month.get('EBITDA %', 0):.1f}%)")
    print(f"  PAT:        Rs {pat:,.1f} Cr ({analysis.current_month.get('PAT %', 0):.1f}%)")

    print(f"\n  Key Highlights:")
    for h in analysis.highlights:
        print(f"    > {h}")

    print(f"\n  Outliers Detected: {len(analysis.outliers)}")
    by_type = {}
    for o in analysis.outliers:
        by_type[o.comparison_type] = by_type.get(o.comparison_type, 0) + 1
    for t, c in by_type.items():
        label = {"vs_AOP": "vs Budget", "MoM": "Month-over-Month",
                 "YoY": "Year-over-Year", "QoQ": "Quarter-over-Quarter"}.get(t, t)
        print(f"    - {label}: {c}")

    high_count = sum(1 for o in analysis.outliers if o.severity == "high")
    print(f"    - High severity: {high_count}")

    print(f"\nGenerating PowerPoint...")
    ppt_path = generate_ppt(analysis, args.output)
    print(f"  -> PPT saved: {ppt_path}")

    print(f"\nDone! Open the PPT to review the analysis.")
    return ppt_path


if __name__ == "__main__":
    main()
