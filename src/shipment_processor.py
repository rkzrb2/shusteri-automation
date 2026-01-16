"""
Процессор для обработки данных формата "Перечень отправки грузы"
"""
from typing import List
from decimal import Decimal
from .models import ShipmentLine, OutputLine
import logging

logger = logging.getLogger(__name__)


class ShipmentProcessor:
    """
    Обработчик для нового формата

    Особенности:
    - БЕЗ разделения по длине стельки (все размеры ≤24см)
    - БЕЗ разделения по ТН ВЭД (у каждой строки свой код)
    - Каждая входная строка → одна выходная строка
    - Расширенное описание товара
    - Группировка по коробкам
    """

    def __init__(self, config: dict):
        """
        Args:
            config: словарь конфигурации
        """
        self.config = config

    def process(self, lines: List[ShipmentLine]) -> List[OutputLine]:
        """
        Обрабатывает список ShipmentLine и создает OutputLine для документов

        Args:
            lines: список входных строк ShipmentLine

        Returns:
            List[OutputLine]: список обработанных строк для документов
        """
        logger.info(f"Начинаю обработку {len(lines)} строк")

        output_lines = []

        for line in lines:
            output = self._convert_to_output(line)
            output_lines.append(output)

        logger.info(f"Обработано строк: {len(output_lines)}")
        return output_lines

    def _convert_to_output(self, line: ShipmentLine) -> OutputLine:
        """
        Конвертирует ShipmentLine в OutputLine

        Логика:
        - Каждая входная строка → одна выходная строка
        - БЕЗ группировки, БЕЗ разделения по стельке
        - Количество = halfpairs_loaded
        - Описание = расширенное

        Args:
            line: входная строка ShipmentLine

        Returns:
            OutputLine: выходная строка
        """
        # Создаем расширенное описание
        description = self._build_description(line)

        # Создаем выходную строку
        output = OutputLine(
            brand=line.brand,
            hs_code=line.hs_code,
            article=line.article,
            description=description,
            color=line.color,
            material=line.material,
            lining=line.lining,
            sole=line.sole,
            heel_height=line.heel_height,
            insole_category=line.insole_category,  # Всегда "до 24см"
            quantity=line.halfpairs_loaded,
            net_weight=line.total_net_weight,
            gross_weight=line.gross_weight_per_box if line.gross_weight_per_box else Decimal('0'),
            boxes=line.boxes_in_group if line.boxes_in_group else 0,
            original_boxes=line.boxes_in_group if line.boxes_in_group else 0,
            price=line.price,
            amount=line.total_amount,

            # НОВЫЕ ПОЛЯ для расширенного описания
            shoe_type=line.shoe_type,
            shaft_height=line.shaft_height,
            halfpair_type=line.halfpair_type,
            perforation=line.perforation,
            box_group=line.box_group
        )

        return output

    def _build_description(self, line: ShipmentLine) -> str:
        """
        Создает расширенное описание товара

        Формат:
        [Вид обуви], верх:[материал], цвет:[цвет], подкладка:[подкладка],
        подошва:[подошва], каблук:[высота каблука], голенище:[высота голенища],
        [левый/правый полупарок], длина стельки до 24см

        Args:
            line: ShipmentLine

        Returns:
            str: описание товара
        """
        parts = []

        # Вид обуви
        parts.append(line.shoe_type)

        # Материал верха
        parts.append(f"верх:{line.material}")

        # Цвет
        parts.append(f"цвет:{line.color}")

        # Подкладка
        parts.append(f"подкладка:{line.lining}")

        # Подошва
        parts.append(f"подошва:{line.sole}")

        # Высота каблука
        parts.append(f"каблук:{line.heel_height}")

        # Высота голенища (если есть)
        if line.shaft_height:
            parts.append(f"голенище:{line.shaft_height}")

        # Тип полупары
        parts.append(line.halfpair_type)

        # Категория стельки (всегда ≤24см)
        parts.append(line.insole_category)

        # Процентный состав (если есть)
        if line.composition:
            parts.append(f"состав:{line.composition}")

        return ", ".join(parts)
