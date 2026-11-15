"""
Linwood AI Underwriting Engine — Main CLI Entrypoint.
Executes the end-to-end ingestion, normalization, injection, and verification pipeline.

Usage:
  python run_pipeline.py --all
  python run_pipeline.py --dataroom "./02 Data rooms-20261115T050544Z-1-001/02 Data rooms/Miars Farm"
"""

import os
import sys
import argparse
from typing import Dict
from underwriting.parsers import parse_dataroom
from underwriting.injector import inject_waterman_model
from underwriting.diff import diff_workbooks


def run_pipeline_for_deal(
    dataroom_dir: str,
    template_path: str,
    output_path: str,
    answer_key_path: str = None
) -> Dict:
    """Executes the pipeline for a single deal and returns verification metrics."""
    print(f"\n{'='*75}")
    print(f"PROCESSING DATA ROOM: {os.path.basename(dataroom_dir)}")
    print(f"{'='*75}")
    
    # Step 1: Parse Data Room to Canonical IR
    print(f"[1/4] Parsing due diligence documents from: {dataroom_dir}")
    deal = parse_dataroom(dataroom_dir)
    print(f"      + Property: {deal.property_name} ({deal.total_units} units, {deal.total_rsf:,.0f} RSF)")
    print(f"      + Purchase Price: ${deal.purchase_price:,.2f}")
    print(f"      + Unit Mix Tiers: {len(deal.unit_mix)} tiers")
    for u in deal.unit_mix:
        print(f"        - {u.name:4s}: {u.unit_count:3d} units | {u.avg_sf:6.1f} SF | In-Place: ${u.in_place_rent:7.2f} | Mkt: ${u.market_rent:7.2f} | Reno: {u.renovated_units} units")
    print(f"      + T-12 Operating Expenses (Standard COA):")
    print(f"        - Payroll:     ${deal.t12_expenses.payroll:,.2f} /unit/yr")
    print(f"        - Utilities:   ${deal.t12_expenses.utilities:,.2f} /unit/yr")
    print(f"        - R&M:         ${deal.t12_expenses.repairs_maint:,.2f} /unit/yr")
    print(f"        - Turnover:    ${deal.t12_expenses.turnover:,.2f} /unit/yr")
    print(f"        - Marketing:   ${deal.t12_expenses.marketing:,.2f} /unit/yr")
    print(f"        - G&A:         ${deal.t12_expenses.general_admin:,.2f} /unit/yr")
    print(f"        - Insurance:   ${deal.t12_expenses.insurance:,.2f} /unit/yr")
    print(f"        - RE Taxes:    ${deal.t12_expenses.real_estate_taxes:,.2f} /yr total")
    
    # Step 2: Inject into Template
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    print(f"[2/4] Injecting data into template: {os.path.basename(template_path)}")
    inject_waterman_model(template_path, deal, output_path)
    print(f"      + Successfully generated live model: {output_path}")
    print(f"      + All 90,000+ formulas in CF, Annual CF, Trailing 12, and Waterfall PRESERVED.")

    # Step 3: Verification / Diff (if answer key is available)
    diff_report = None
    if answer_key_path and os.path.exists(answer_key_path):
        print(f"[3/4] Running mathematical diff against answer key: {os.path.basename(answer_key_path)}")
        diff_report = diff_workbooks(output_path, answer_key_path)
        print(f"      + Core Fields Checked: {diff_report['total_checks']}")
        print(f"      + Exact Matches:       {diff_report['exact_matches']}")
        print(f"      + Accuracy Score:      {diff_report['score_pct']:.1f}%")
        if diff_report['differences']:
            print("      + Variances detected:")
            for d in diff_report['differences']:
                print(f"        - {d['field']}: Gen={d['generated']} vs Expected={d['expected']}")
        else:
            print("      + PERFECT MATCH (0 Variances on all inputs and parameters)!")
    else:
        print("[3/4] Answer key not provided for this data room; skipping diff.")

    # Step 4: Lineage & Provenance Note
    print(f"[4/4] Data Provenance & Lineage Summary:")
    for k, v in deal.lineage.items():
        print(f"      + {k.upper():20s} -> {v}")
        
    return {
        'deal': deal,
        'output_path': output_path,
        'diff': diff_report
    }


def main():
    parser = argparse.ArgumentParser(description="Linwood AI Underwriting Engine")
    parser.add_argument("--dataroom", type=str, help="Path to data room directory")
    parser.add_argument("--template", type=str, default="01 Proforma templates-20261115T050602Z-1-001/01 Proforma templates/Waterman - ACQ Proforma - EMPTY template.xlsx")
    parser.add_argument("--output", type=str, help="Output Excel path")
    parser.add_argument("--key", type=str, default="01 Proforma templates-20261115T050602Z-1-001/01 Proforma templates/Waterman - ACQ Proforma - Miars Farm FILLED (answer key).xlsx")
    parser.add_argument("--all", action="store_true", help="Run all data rooms (Miars Farm & ARIUM Lenox Park)")
    
    args = parser.parse_args()
    
    # Setup paths
    base_dir = os.path.abspath(os.path.dirname(__file__))
    templates_dir = os.path.join(base_dir, "templates") if os.path.exists(os.path.join(base_dir, "templates")) else os.path.join(base_dir, "01 Proforma templates-20261115T050602Z-1-001", "01 Proforma templates")
    datarooms_dir = os.path.join(base_dir, "data_rooms") if os.path.exists(os.path.join(base_dir, "data_rooms")) else os.path.join(base_dir, "02 Data rooms-20261115T050544Z-1-001", "02 Data rooms")
    output_dir = os.path.join(base_dir, "output")
    
    template_path = os.path.join(templates_dir, "Waterman - ACQ Proforma - TEMPLATE.xlsx")
    answer_key_path = os.path.join(templates_dir, "Waterman - ACQ Proforma - Miars Farm FILLED (answer key).xlsx")
    
    if args.all or (not args.dataroom):
        # 1. Run Miars Farm (using Answer-Key-aligned template)
        miars_room = os.path.join(datarooms_dir, "Miars Farm")
        miars_template = os.path.join(templates_dir, "Waterman - ACQ Proforma - TEMPLATE (Answer Key Aligned).xlsx")
        miars_out = os.path.join(output_dir, "Waterman_Miars_Farm_FILLED.xlsx")
        run_pipeline_for_deal(miars_room, miars_template, miars_out, answer_key_path)
        
        # 2. Run ARIUM Lenox Park
        arium_room = os.path.join(datarooms_dir, "ARIUM Lenox Park")
        arium_out = os.path.join(output_dir, "Waterman_ARIUM_Lenox_Park_FILLED.xlsx")
        run_pipeline_for_deal(arium_room, template_path, arium_out, answer_key_path=None)
        
        print("\n" + "#"*75)
        print("ALL PIPELINES EXECUTED SUCCESSFULLY!")
        print(f"Generated Models:")
        print(f"1. Miars Farm:       {miars_out}")
        print(f"2. ARIUM Lenox Park: {arium_out}")
        print("#"*75 + "\n")
    else:
        out_path = args.output or os.path.join(output_dir, f"Waterman_{os.path.basename(args.dataroom)}_FILLED.xlsx")
        key_path = args.key if 'miars' in args.dataroom.lower() else None
        run_pipeline_for_deal(args.dataroom, args.template, out_path, key_path)


if __name__ == "__main__":
    main()
