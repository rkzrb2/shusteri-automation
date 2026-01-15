"""
Генераторы документов (ИСПРАВЛЕННАЯ ВЕРСИЯ БЕЗ ГРУППИРОВКИ)
"""
from .invoice import InvoiceGenerator
from .specification import SpecificationGenerator
from .packing_list import PackingListGenerator

__all__ = ['InvoiceGenerator', 'SpecificationGenerator', 'PackingListGenerator']
