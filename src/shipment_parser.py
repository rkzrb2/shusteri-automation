"""
Парсер для файлов формата "Перечень отправки грузы"
"""
import pandas as pd
from typing import List, Optional, Tuple
from decimal import Decimal
from .models import ShipmentLine
import logging

logger = logging.getLogger(__name__)


class ShipmentParser:
    """
    Парсер для нового формата входного файла

    Особенности:
    - Каждая строка = один товар одного размера
    - Обработка объединенных ячеек в столбце КОРОБОК (P)
    - Поддержка полупар (левый/правый)
    """

    def __init__(self, file_path: str):
        """
        Args:
            file_path: путь к Excel файлу
        """
        self.file_path = file_path
        self.df = None
        self.box_groups = []  # Список групп коробок

    def parse(self) -> List[ShipmentLine]:
        """
        Парсит Excel файл и возвращает список ShipmentLine

        Returns:
            List[ShipmentLine]: список распарсенных товаров
        """
        logger.info(f"Начинаю парсинг файла: {self.file_path}")

        # Читаем файл
        self.df = pd.read_excel(self.file_path)

        logger.info(f"Прочитано строк: {len(self.df)}")

        # Определяем группы коробок (объединенные ячейки в столбце P)
        self._detect_box_groups()

        logger.info(f"Обнаружено групп коробок: {len(self.box_groups)}")

        # Парсим каждую строку
        lines = []
        for idx, row in self.df.iterrows():
            try:
                line = self._parse_row(idx, row)
                if line:
                    lines.append(line)
            except Exception as e:
                logger.error(f"Ошибка парсинга строки {idx + 2}: {e}")
                continue

        logger.info(f"Успешно распарсено строк: {len(lines)}")
        return lines

    def _detect_box_groups(self):
        """
        Определяет группы коробок по столбцу КОРОБОК (P)

        Логика:
        - Если значение != NaN → начало новой группы
        - Все последующие NaN до следующего значения → та же группа

        Результат сохраняется в self.box_groups:
        [
            {'start_row': 0, 'end_row': 2, 'box_count': 1},
            {'start_row': 3, 'end_row': 5, 'box_count': 2},
            ...
        ]
        """
        box_col = 'КОРОБОК'  # Столбец P

        if box_col not in self.df.columns:
            logger.warning(f"Столбец '{box_col}' не найден в файле")
            return

        current_group = None
        group_id = 0

        for idx, value in enumerate(self.df[box_col]):
            if pd.notna(value):  # Начало новой группы
                # Закрываем предыдущую группу
                if current_group is not None:
                    current_group['end_row'] = idx - 1
                    self.box_groups.append(current_group)

                # Начинаем новую группу
                current_group = {
                    'id': group_id,
                    'start_row': idx,
                    'end_row': idx,  # Будет обновлено
                    'box_count': int(value)
                }
                group_id += 1

        # Закрываем последнюю группу
        if current_group is not None:
            current_group['end_row'] = len(self.df) - 1
            self.box_groups.append(current_group)

    def _get_box_group(self, row_idx: int) -> Tuple[Optional[int], Optional[int]]:
        """
        Определяет номер группы коробки и количество коробок для строки

        Args:
            row_idx: индекс строки

        Returns:
            (group_id, box_count) или (None, None)
        """
        for group in self.box_groups:
            if group['start_row'] <= row_idx <= group['end_row']:
                return group['id'], group['box_count']
        return None, None

    def _parse_row(self, idx: int, row: pd.Series) -> Optional[ShipmentLine]:
        """
        Парсит одну строку таблицы

        Args:
            idx: индекс строки
            row: Series с данными строки

        Returns:
            ShipmentLine или None если строка пустая
        """
        # Проверяем обязательные поля
        if 'артикул' not in row or pd.isna(row['артикул']):
            return None

        # Определяем группу коробки
        box_group, boxes_in_group = self._get_box_group(idx)

        # Парсим тип полупары
        halfpair_raw = str(row.get(' левый полупарок/ правый полупарок', '')).strip()

        # Создаем объект
        line = ShipmentLine(
            row_number=idx + 2,  # Excel нумерация с 1, заголовок = 1
            article=str(row['артикул']).strip(),
            brand=str(row['марка']).strip(),
            hs_code=str(row['таможенный код']).strip(),
            shoe_type=str(row['вид обуви']).strip(),
            material=str(row['материал верх']).strip(),
            color=str(row['цвет']).strip(),
            lining=str(row['материал подкладка']).strip(),
            sole=str(row['материал подошва']).strip(),
            heel_height=str(row['Высота каблука']).strip(),
            shaft_height=str(row['высота гленище']).strip() if pd.notna(row.get('высота гленище', '')) else '',
            composition=str(row['процентный состав']).strip() if pd.notna(row.get('процентный состав')) else None,
            perforation=str(row['с/без перфорации']).strip(),
            size=int(row['размер']),
            halfpair_type=halfpair_raw,
            halfpairs_loaded=int(row['полупар  ЗАГРУЖЕНЫ']) if pd.notna(row.get('полупар  ЗАГРУЖЕНЫ')) else 1,
            box_group=box_group,
            boxes_in_group=boxes_in_group,
            net_weight_per_unit=Decimal(str(row['вес нетто на штук'])),
            gross_weight_per_box=Decimal(str(row['вес брутто на коробку'])) if pd.notna(row.get('вес брутто на коробку')) else None,
            box_volume=Decimal(str(row['ОБЬЁМ КОРОБКИ'])) if pd.notna(row.get('ОБЬЁМ КОРОБКИ')) else None,
            price=Decimal(str(row['цена']))
        )

        return line
