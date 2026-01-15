"""
Процессор данных - преобразование в строки для документов
"""
from typing import List
from decimal import Decimal
from .models import ProductLine, OutputLine


class DataProcessor:
    """Обработчик данных"""

    def __init__(self, config: dict):
        self.config = config

    def process(self, products: List[ProductLine]) -> List[OutputLine]:
        """
        Преобразует входные строки в выходные
        
        Логика:
        - Если коды ТН ВЭД РАЗНЫЕ (есть слэш во входном файле) → создаем ДВЕ строки: для ≤24см и >24см
        - Если коды ТН ВЭД ОДИНАКОВЫЕ (нет слэша во входном файле) → создаем ОДНУ строку со всеми размерами
        """
        output_lines = []

        for product in products:
            # Проверяем, одинаковые ли коды ТН ВЭД
            codes_are_same = (product.hs_code_le24 == product.hs_code_gt24)
            
            if codes_are_same:
                # Если код один - создаем ОДНУ строку со всеми парами
                if product.total_pairs > 0:
                    output_lines.append(self._create_output_line(
                        product,
                        category="all",  # Все размеры вместе
                        quantity=product.total_pairs
                    ))
            else:
                # Если коды разные - создаем ДВЕ строки (как раньше)
                
                # Строка для категории ≤24см
                if product.pairs_le24 > 0:
                    output_lines.append(self._create_output_line(
                        product,
                        category="le24",
                        quantity=product.pairs_le24
                    ))

                # Строка для категории >24см
                if product.pairs_gt24 > 0:
                    output_lines.append(self._create_output_line(
                        product,
                        category="gt24",
                        quantity=product.pairs_gt24
                    ))

        return output_lines

    def _create_output_line(
            self,
            product: ProductLine,
            category: str,
            quantity: int
    ) -> OutputLine:
        """Создает строку для документа"""

        # Определяем HS код и описание категории
        if category == "le24":
            hs_code = product.hs_code_le24
            insole_text = "длина стельки до 24см"
        elif category == "gt24":
            hs_code = product.hs_code_gt24
            insole_text = "длина стельки более 24см"
        else:  # category == "all"
            # Если код один для всех размеров - используем его
            hs_code = product.hs_code_le24  # Они одинаковые, можно взять любой
            
            # Определяем диапазон размеров автоматически
            sizes = sorted(product.qty_by_size.keys())
            if sizes:
                min_size = min(sizes)
                max_size = max(sizes)
                insole_text = f"все размеры {min_size}-{max_size}"
            else:
                insole_text = "все размеры"

        # Рассчитываем веса пропорционально количеству
        net_weight = product.net_weight_per_pair * quantity

        # Брутто: рассчитываем пропорционально от брутто за коробку
        # (предполагаем, что boxes указано для всех пар)
        if product.total_pairs > 0:
            gross_weight = (product.gross_weight_per_box * product.boxes * quantity) / product.total_pairs
        else:
            gross_weight = Decimal(0)

        # Количество коробок берем из исходного файла БЕЗ пересчета
        # Для разделенных строк (le24/gt24) boxes отображается только в первой строке
        boxes = product.boxes

        # Сумма
        amount = product.price * quantity

        return OutputLine(
            brand=product.brand,
            hs_code=hs_code,
            article=product.article,
            description=product.name,
            color=product.color,
            material=product.material,
            lining=product.lining,
            sole=product.sole,
            heel_height=product.heel_height,
            insole_category=insole_text,
            quantity=quantity,
            net_weight=round(net_weight, 3),
            gross_weight=round(gross_weight, 3),
            boxes=boxes if category in ["le24", "all"] else 0,  # Коробки указываем в первой строке или в единственной
            original_boxes=product.boxes,  # ← ДОБАВЛЕНО: Сохраняем исходное количество коробок
            price=product.price,
            amount=round(amount, 2),
            kiz_codes=[]
        )
