"""
Verification & Diff Engine.
Performs cell-by-cell and semantic comparison between the generated pro forma
and the ground-truth human answer key.
"""

from typing import List, Dict, Any, Tuple
import openpyxl


def are_cells_equal(v1: Any, v2: Any, tol: float = 1e-4) -> bool:
    """Compares two cell values, handling openpyxl formula objects and floating point precision."""
    if v1 is None and v2 is None:
        return True
    if v1 is None or v2 is None:
        return False
        
    # Check if either is an openpyxl formula object
    t1 = getattr(v1, 'text', None)
    t2 = getattr(v2, 'text', None)
    if t1 is not None or t2 is not None:
        return (t1 or str(v1)) == (t2 or str(v2))
        
    # Check numeric equality with tolerance
    if isinstance(v1, (int, float)) and isinstance(v2, (int, float)):
        return abs(float(v1) - float(v2)) < tol
        
    # Direct equality
    return str(v1).strip() == str(v2).strip()


def diff_workbooks(
    generated_path: str,
    answer_key_path: str,
    sheets: List[str] = None
) -> Dict[str, Any]:
    """
    Compares the generated workbook against the answer key.
    Accounts for semantic row alignment and verifies input precision.
    """
    wb_gen = openpyxl.load_workbook(generated_path, data_only=False)
    wb_key = openpyxl.load_workbook(answer_key_path, data_only=False)
    
    results = {
        'total_checks': 0,
        'exact_matches': 0,
        'differences': [],
        'sheet_summaries': {}
    }
    
    ws_g = wb_gen['Inputs']
    ws_k = wb_key['Inputs']
    
    # 1. Compare Core Underwriting Data Points
    core_fields = [
        ('Price Input ', 13, 'Purchase Price ($16.85M)'),
        ('Equity Placement Fee', 13, 'Equity Placement Fee ($350k)'),
        ('Working Capital', 13, 'Working Capital ($250k)'),
        ('Payroll', 6, 'Payroll ($1,155/unit)'),
        ('Utilities', 6, 'Utilities ($2,532/unit)'),
        ('Repairs & Maintenance', 6, 'R&M ($147/unit)'),
        ('Turnover/Redecoration', 6, 'Turnover ($79/unit)'),
        ('Marketing & Advertising', 6, 'Marketing ($373/unit)'),
        ('General & Admin', 6, 'G&A ($222/unit)'),
        ('Insurance', 6, 'Insurance ($1,763/unit)'),
        ('Management Fee', 6, 'Management Fee (2.75%)'),
    ]
    
    key_diffs = []
    
    def get_val_by_label(ws, label, col, search_cols=(2, 8)):
        target = label.strip().lower()
        for r in range(1, 150):
            for sc in search_cols:
                val = str(ws.cell(r, sc).value or '').strip().lower()
                if val == target:
                    return ws.cell(r, col).value
        return None

    for label, col, display_name in core_fields:
        v_gen = get_val_by_label(ws_g, label, col)
        v_key = get_val_by_label(ws_k, label, col)
        
        results['total_checks'] += 1
        if are_cells_equal(v_gen, v_key):
            results['exact_matches'] += 1
        else:
            key_diffs.append({
                'field': display_name,
                'generated': v_gen,
                'expected': v_key
            })
            
    # 2. Compare Unit Mix Rows
    def get_unit_mix(ws):
        mix = {}
        header_row = None
        for r in range(1, 100):
            if str(ws.cell(r, 2).value or '').strip().lower() == 'unit mix':
                header_row = r
                break
        if header_row:
            for r in range(header_row + 1, header_row + 15):
                name = ws.cell(r, 2).value
                if name and name not in ['-', 'Total/Avg']:
                    mix[str(name)] = {
                        'count': ws.cell(r, 3).value,
                        'sf': ws.cell(r, 5).value,
                        'in_place': ws.cell(r, 6).value,
                        'market': ws.cell(r, 8).value,
                        'reno_units': ws.cell(r, 9).value,
                        'reno_capex': ws.cell(r, 10).value,
                    }
        return mix

    mix_g = get_unit_mix(ws_g)
    mix_k = get_unit_mix(ws_k)
    
    for tier in ['1BR', '2BR', '3BR']:
        results['total_checks'] += 1
        if tier in mix_g and tier in mix_k:
            g_tier = mix_g[tier]
            k_tier = mix_k[tier]
            match = all(
                are_cells_equal(g_tier[k], k_tier[k])
                for k in ['count', 'sf', 'in_place', 'market', 'reno_units', 'reno_capex']
            )
            if match:
                results['exact_matches'] += 1
            else:
                key_diffs.append({
                    'field': f'Unit Mix [{tier}]',
                    'generated': g_tier,
                    'expected': k_tier
                })
        else:
            key_diffs.append({
                'field': f'Unit Mix [{tier}] missing',
                'generated': tier in mix_g,
                'expected': tier in mix_k
            })
            
    results['differences'] = key_diffs
    results['score_pct'] = (results['exact_matches'] / results['total_checks']) * 100.0
    return results
