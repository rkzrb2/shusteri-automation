"""
Модуль форматирования документов
"""
from .invoice_formatter import InvoiceFormatter
from .specification_formatter import SpecificationFormatter
from .packing_list_formatter import PackingListFormatter

__all__ = ['InvoiceFormatter', 'SpecificationFormatter', 'PackingListFormatter']
