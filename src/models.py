"""
Модели данных
"""
from pydantic import BaseModel, Field
from typing import Optional, List
from decimal import Decimal


class ProductLine(BaseModel):
    """Одна строка товара из входного файла"""
    row_number: int
    brand: str
    article: str
    code_raw: str
    hs_code_le24: str = ""  # ≤24см
    hs_code_gt24: str = ""  # >24см
    name: str
    material: str
    color: str
    lining: str
    sole: str
    heel_height: str
    composition: Optional[str] = None
    price: Decimal
    boxes: int
    net_weight_per_pair: Decimal
    gross_weight_per_box: Decimal

    # Количество по размерам
    qty_by_size: dict  # {35: 100, 36: 200, ...}

    @property
    def total_pairs(self) -> int:
        """Общее количество пар"""
        return sum(self.qty_by_size.values())

    @property
    def pairs_le24(self) -> int:
        """Пары ≤24см (размеры 35-37)"""
        return sum(qty for size, qty in self.qty_by_size.items()
                   if size in [35, 36, 37])

    @property
    def pairs_gt24(self) -> int:
        """Пары >24см (размеры 38+)"""
        return sum(qty for size, qty in self.qty_by_size.items()
                   if size in [38, 39, 40, 41, 42])


class OutputLine(BaseModel):
    """Строка для выходных документов"""
    brand: str
    hs_code: str
    article: str
    description: str
    color: str
    material: str
    lining: str
    sole: str
    heel_height: str
    insole_category: str  # "длина стельки до 24см" или "длина стельки более 24см"
    quantity: int
    net_weight: Decimal
    gross_weight: Decimal
    boxes: int
    original_boxes: int = 0
    price: Decimal
    amount: Decimal
    kiz_codes: List[str] = []


class DocumentMetadata(BaseModel):
    """Метаданные документа"""
    invoice_number: str
    date: str
    container_number: str = ""
    seller_name: str
    seller_name_en: str = ""
    seller_address: str
    seller_address_en: str = ""
    buyer_name: str
    buyer_address: str
    buyer_address_en: str = ""
    contract_number: str
    contract_date: str
    terms_of_delivery: str
    currency: str = "CNY"
