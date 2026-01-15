"""
Парсер входного файла
"""
import pandas as pd
from typing import List, Tuple
from decimal import Decimal
from .models import ProductLine
import logging

logger = logging.getLogger(__name__)


class InputFileParser:
    """Парсер китайской таблицы"""

    def __init__(self, config: dict):
        self.config = config
        self.columns = config['input_columns']

    def parse(self, file_path: str) -> List[ProductLine]:
        """Читает и парсит входной файл"""

        # Читаем Excel
        df = pd.read_excel(file_path, sheet_name=0)

        # Удаляем пустые строки и итоговые строки
        df = df[df[self.columns['article']].notna()]
        df = df[df[self.columns['article']] != '']

        products = []

        for idx, row in df.iterrows():
            try:
                # Парсим код (формат: 6403911100/6403911800)
                code_raw = str(row[self.columns['code']])
                hs_le24, hs_gt24 = self._parse_hs_code(code_raw)

                # Собираем количество по размерам
                qty_by_size = {}
                for size in [35, 36, 37, 38, 39, 40, 41, 42]:
                    col_name = self.columns.get(f'qty_{size}')
                    if col_name and col_name in row and pd.notna(row[col_name]):
                        qty = int(row[col_name])
                        if qty > 0:
                            qty_by_size[size] = qty

                if not qty_by_size:
                    logger.warning(f"Строка {idx + 2}: нет количества по размерам, пропускаем")
                    continue

                product = ProductLine(
                    row_number=idx + 2,
                    brand=str(row[self.columns['brand']]),
                    article=str(row[self.columns['article']]),
                    code_raw=code_raw,
                    hs_code_le24=hs_le24,
                    hs_code_gt24=hs_gt24,
                    name=str(row[self.columns['name']]),
                    material=str(row[self.columns['material']]),
                    color=str(row[self.columns['color']]),
                    lining=str(row[self.columns['lining']]),
                    sole=str(row[self.columns['sole']]),
                    heel_height=str(row[self.columns['heel_height']]),
                    composition=str(row[self.columns['composition']]) if pd.notna(
                        row.get(self.columns['composition'])) else None,
                    price=Decimal(str(row[self.columns['price']])),
                    boxes=int(row[self.columns['boxes']]),
                    net_weight_per_pair=Decimal(str(row[self.columns['net_weight_per_pair']])),
                    gross_weight_per_box=Decimal(str(row[self.columns['gross_weight_per_box']])),
                    qty_by_size=qty_by_size
                )

                products.append(product)
                logger.info(f"✓ Обработана строка {idx + 2}: {product.article} ({product.total_pairs} пар)")

            except Exception as e:
                logger.error(f"Ошибка в строке {idx + 2}: {e}")
                continue

        logger.info(f"Всего обработано: {len(products)} позиций")
        return products

    def _parse_hs_code(self, code_raw: str) -> Tuple[str, str]:
        """
        Разделяет код вида 6403911100/6403911800
        Returns: (код для ≤24см, код для >24см)
        """
        if '/' in code_raw:
            parts = code_raw.split('/')
            return parts[0].strip(), parts[1].strip()
        else:
            # Если слэша нет - используем один код для обеих категорий
            return code_raw.strip(), code_raw.strip()