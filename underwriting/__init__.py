"""
Linwood AI Underwriting Engine
A deterministic, auditable pipeline for parsing commercial real estate data rooms
and injecting standardized inputs into institutional Excel models.
"""

from .models import DealUnderwriting, UnitMixItem, T12Expenses
from .parsers import parse_dataroom
from .injector import inject_waterman_model
from .diff import diff_workbooks

__all__ = [
    "DealUnderwriting",
    "UnitMixItem",
    "T12Expenses",
    "parse_dataroom",
    "inject_waterman_model",
    "diff_workbooks",
]
