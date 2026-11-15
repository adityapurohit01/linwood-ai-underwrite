"""
Canonical Real Estate Underwriting Intermediate Representation (IR).
Decouples raw due diligence data room extraction from Excel model injection.
"""

from dataclasses import dataclass, field, asdict
from typing import List, Dict, Optional


@dataclass
class UnitMixItem:
    """Represents an aggregated bedroom/floorplan tier."""
    name: str                        # e.g., '1BR', '2BR', '3BR'
    unit_count: int                  # Number of units in this tier
    avg_sf: float                    # Net rentable square feet per unit
    in_place_rent: float             # Actual contractual in-place rent ($/unit/month)
    market_rent: float               # Market or post-renovation achievable rent ($/unit/month)
    renovated_units: int = 0         # Number of classic units slated for renovation
    reno_capex_per_unit: float = 0.0 # Renovation cost per unit ($)
    reno_pace_per_month: int = 1     # Monthly renovation velocity

    def to_dict(self) -> Dict:
        return asdict(self)


@dataclass
class T12Expenses:
    """Standardized commercial real estate Chart of Accounts (COA) per unit/year."""
    payroll: float                   # Personnel, leasing, maintenance staff ($/unit/yr)
    utilities: float                 # Water, sewer, gas, electric, waste ($/unit/yr)
    repairs_maint: float             # Routine & recurring maintenance ($/unit/yr)
    turnover: float                  # Unit turns, painting, carpet clean ($/unit/yr)
    marketing: float                 # Advertising, locator fees, signage ($/unit/yr)
    general_admin: float             # Software, telecom, legal, bank fees ($/unit/yr)
    insurance: float                 # Property, casualty, general liability ($/unit/yr)
    real_estate_taxes: float         # Annual real estate property taxes ($/yr total)
    management_fee_pct: float = 0.0275 # % of Effective Gross Income (EGI)
    replacement_reserve: float = 15.0  # Capex reserve buffer ($/unit/month)

    def to_dict(self) -> Dict:
        return asdict(self)


@dataclass
class DealUnderwriting:
    """Canonical model capturing all deal terms, facts, and underwriting assumptions."""
    property_name: str
    address: str
    city_state_zip: str
    total_units: int
    total_rsf: float
    purchase_price: float
    unit_mix: List[UnitMixItem]
    t12_expenses: T12Expenses
    year_built: str = "1998 / Renovated 2021"
    
    # Capitalization & Closing Costs
    equity_placement_fee: float = 350000.0
    working_capital: float = 250000.0
    closing_cost_pct: float = 0.01
    acquisition_fee_pct: float = 0.01
    
    # Financing Assumptions (Senior Debt)
    ltv: float = 0.65
    interest_rate: float = 0.06
    amort_months: int = 360
    io_months: int = 60
    rate_buydown_active: bool = True
    rate_reduction_bps: float = 45.0
    cost_of_buydown: float = 500000.0
    
    # Valuation & Hold
    hold_period_months: int = 60
    exit_cap_rate: float = 0.0525
    
    # Audit Provenance: Maps field name -> source location / assumption disclosure
    lineage: Dict[str, str] = field(default_factory=dict)

    def to_dict(self) -> Dict:
        return asdict(self)

