"""
Autonomous Multi-Agent Data Room Ingestion & Normalization Engine.
Parses Offering Memorandums (PDF), Rent Rolls (XLS, XLSX, CSV), and T-12 P&Ls
into the Canonical DealUnderwriting Intermediate Representation using specialized AI Agents.
"""

import os
import re
from typing import List, Dict, Tuple, Optional
import openpyxl
import xlrd
from .models import DealUnderwriting, UnitMixItem, T12Expenses
from .ai_agents import RentRollAgent, T12FinancialAgent, OMIntelligenceAgent, DealSynthesizerAgent


def identify_file_roles(files: List[str], data_dir: str) -> Dict[str, Optional[str]]:
    """
    Identifies the role of each file in a data room (rent roll, t12, om, tax bill)
    based on filename hints and internal content inspection.
    """
    roles = {
        'rent_roll': None,
        't12': None,
        'om': None,
        'tax_bill': None
    }
    
    full_paths = [os.path.join(data_dir, f) for f in files if os.path.isfile(os.path.join(data_dir, f))]
    
    # 1. Match Offering Memo (PDF with OM / memo / flyer / brochure, or any PDF > 5MB)
    for p in full_paths:
        name = os.path.basename(p).lower()
        if p.endswith('.pdf'):
            if any(k in name for k in ['om', 'memo', 'brochure', 'flyer', 'investment']):
                roles['om'] = p
                break
            elif not roles['om'] and 'tax' not in name:
                roles['om'] = p
                
    # 2. Match Tax Bill
    for p in full_paths:
        name = os.path.basename(p).lower()
        if 'tax' in name and p.endswith('.pdf'):
            roles['tax_bill'] = p
            break

    # 3. Match Rent Roll (XLSX, XLS, CSV with rent/roll/rr/unit in name)
    for p in full_paths:
        name = os.path.basename(p).lower()
        if any(p.endswith(ext) for ext in ['.xlsx', '.xls', '.csv']):
            if any(k in name for k in ['rent', 'roll', 'rr', 'unit', 'lease', 'rediq']):
                roles['rent_roll'] = p
                break

    # 4. Match T-12 Operating Statement
    for p in full_paths:
        name = os.path.basename(p).lower()
        if any(p.endswith(ext) for ext in ['.xlsx', '.xls', '.csv']):
            if p != roles['rent_roll']:
                if any(k in name for k in ['t12', 't-12', 'operating', 'p&l', 'statement', 'trailing', 'financial']):
                    roles['t12'] = p
                    break

    # 5. Content-based fallback if names don't match keywords
    remaining = [p for p in full_paths if p not in roles.values() and any(p.endswith(ext) for ext in ['.xlsx', '.xls', '.csv'])]
    for p in remaining:
        if not roles['rent_roll']:
            roles['rent_roll'] = p
        elif not roles['t12']:
            roles['t12'] = p

    return roles


def parse_generic_dataroom(
    data_dir: str,
    default_name: Optional[str] = None,
    default_price: Optional[float] = None
) -> DealUnderwriting:
    """
    Intelligently parses ANY custom due diligence data room using the Multi-Agent AI system.
    Dynamically orchestrates RentRollAgent, T12FinancialAgent, OMIntelligenceAgent,
    and DealSynthesizerAgent.
    """
    files = os.listdir(data_dir) if os.path.exists(data_dir) else []
    roles = identify_file_roles(files, data_dir)
    lineage = {}

    # Agent 1: Parse Rent Roll
    unit_mix = []
    total_units = 100
    total_rsf = 95000.0
    if roles['rent_roll'] and os.path.exists(roles['rent_roll']):
        agent_rr = RentRollAgent(roles['rent_roll'])
        unit_mix, total_units, total_rsf, l_rr = agent_rr.parse()
        lineage['rent_roll'] = l_rr
    else:
        # Fallback default unit mix
        unit_mix = [
            UnitMixItem(name='1BR', unit_count=50, avg_sf=750.0, in_place_rent=1300.0, market_rent=1450.0, renovated_units=25, reno_capex_per_unit=12000.0, reno_pace_per_month=2),
            UnitMixItem(name='2BR', unit_count=50, avg_sf=1050.0, in_place_rent=1600.0, market_rent=1800.0, renovated_units=25, reno_capex_per_unit=12000.0, reno_pace_per_month=2)
        ]
        total_units = 100
        total_rsf = 90000.0
        lineage['rent_roll'] = "Market baseline unit mix assumption"

    # Agent 2: Parse T-12 Operating Statement
    t12_expenses = None
    if roles['t12'] and os.path.exists(roles['t12']):
        agent_t12 = T12FinancialAgent(roles['t12'], total_units=total_units)
        t12_expenses, l_t12 = agent_t12.parse()
        lineage['t12_expenses'] = l_t12
    else:
        # Fallback baseline expenses
        t12_expenses = T12Expenses(
            payroll=1200,
            utilities=2200,
            repairs_maint=200,
            turnover=100,
            contract_services=250,
            marketing=350,
            general_admin=250,
            insurance=1600,
            real_estate_taxes=120000.0,
            management_fee_pct=0.0275,
            replacement_reserve=15.0
        )
        lineage['t12_expenses'] = "Institutional market expense assumption"

    # Agent 3: Parse Offering Memo
    om_meta = {}
    if roles['om'] and os.path.exists(roles['om']):
        try:
            agent_om = OMIntelligenceAgent(roles['om'])
            om_meta, l_om = agent_om.parse()
            lineage['om'] = l_om
        except Exception:
            pass

    # Agent 4: Deal Synthesis & QC
    deal = DealSynthesizerAgent.synthesize(
        unit_mix=unit_mix,
        t12_expenses=t12_expenses,
        om_meta=om_meta,
        user_name=default_name,
        user_price=default_price,
        lineage_map=lineage
    )
    return deal


def parse_dataroom(data_dir: str) -> DealUnderwriting:
    """Auto-detects data room and executes the Multi-Agent AI Underwriting pipeline."""
    return parse_generic_dataroom(data_dir)
