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


class ShipmentLine(BaseModel):
    """
    Одна строка товара из файла "Перечень отправки грузы"
    Представляет один товар одного размера с указанием типа полупары
    """

    # Идентификация
    row_number: int                         # Номер строки в Excel
    article: str                            # Артикул (пример: "SU29133-1R")
    brand: str                              # Марка (пример: "V.I.KONTY")

    # ТН ВЭД (единственный!)
    hs_code: str                            # Код ТН ВЭД (пример: "6403911100")

    # Характеристики товара
    shoe_type: str                          # Вид обуви (пример: "Ботинки женские")
    material: str                           # Материал верха (пример: "нат.замша")
    color: str                              # Цвет (пример: "чёрный")
    lining: str                             # Материал подкладки (пример: "текстиль")
    sole: str                               # Материал подошвы (пример: "нитрилкаучук")
    heel_height: str                        # Высота каблука (пример: "невыс. (ров./до 3 см.)")
    shaft_height: str                       # Высота голенища (пример: "12cm")
    composition: Optional[str] = None       # Процентный состав (пример: "0.95/0.05")
    perforation: str                        # С/без перфорации (пример: "без")

    # Размер и тип полупары
    size: int                               # Размер (36 или 37)
    halfpair_type: str                      # "левый полупарок" или "правый полупарок"

    # Количество и упаковка
    halfpairs_loaded: int                   # Количество полупар загружено (обычно 1)
    box_group: Optional[int] = None         # Номер коробки (группа объединенных ячеек)
    boxes_in_group: Optional[int] = None    # Количество коробок в группе (из столбца P)

    # Веса и объем
    net_weight_per_unit: Decimal            # Вес нетто на штук (пример: 0.37)
    gross_weight_per_box: Optional[Decimal] = None  # Вес брутто на коробку (пример: 15.4)
    box_volume: Optional[Decimal] = None    # Объём коробки

    # Цена
    price: Decimal                          # Цена за полупару (пример: 64.5)

    # Вычисляемые свойства
    @property
    def total_net_weight(self) -> Decimal:
        """Вес нетто = вес на единицу × количество полупар"""
        return self.net_weight_per_unit * Decimal(self.halfpairs_loaded)

    @property
    def total_amount(self) -> Decimal:
        """Сумма = цена × количество полупар"""
        return self.price * Decimal(self.halfpairs_loaded)

    @property
    def insole_category(self) -> str:
        """
        Категория длины стельки
        Для нового формата: только размеры 36-37, все ≤24см
        """
        return "длина стельки до 24см"


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

    # НОВЫЕ ПОЛЯ для расширенного описания (для режима shipment):
    shoe_type: Optional[str] = None         # Вид обуви
    shaft_height: Optional[str] = None      # Высота голенища
    halfpair_type: Optional[str] = None     # Левый/правый полупарок
    perforation: Optional[str] = None       # С/без перфорации

    # НОВОЕ ПОЛЕ для группировки коробок:
    box_group: Optional[int] = None         # Номер группы коробки


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
