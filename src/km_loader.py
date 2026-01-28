"""
Загрузчик справочника КМ (кодов маркировки) из системы Честный знак
"""
import pandas as pd
import re
import logging
from pathlib import Path
from typing import List, Dict, Tuple, Optional

logger = logging.getLogger(__name__)


class KMLoader:
    """Загрузчик и хранилище кодов маркировки (КМ)"""

    def __init__(self, file_path: str):
        """
        Args:
            file_path: путь к файлу выгрузки из Честный знак
        """
        self.file_path = Path(file_path)
        self._index: Dict[Tuple[str, int], List[str]] = {}
        self._load()

    def _extract_article(self, nomenclature: str) -> Optional[str]:
        """
        Извлекает артикул из строки номенклатуры.

        Примеры:
            'Туфли женские арт.GL24136-1(натуральная, натуральная кожа)' -> 'GL24136-1'
            'Полуботинки женские арт.СBF25086-1(серо-бежевый, натуральная кожа)' -> 'BF25086-1'
        """
        # Паттерн: точка, возможно мусорные символы, затем артикул (буквы + цифры + дефисы)
        match = re.search(r'\.[\s\W]*([A-Z]{1,4}[0-9\-]+[A-Za-z0-9\-]*)', str(nomenclature))
        if match:
            return match.group(1)

        # Запасной вариант — ищем просто артикул без привязки к точке
        match = re.search(r'([A-Z]{2,4}[0-9]{4,}[\-0-9A-Za-z]*)', str(nomenclature))
        return match.group(1) if match else None

    def _load(self):
        """Загружает данные из файла и строит индекс"""
        logger.info(f"Загрузка файла КМ: {self.file_path}")

        try:
            df = pd.read_excel(self.file_path, header=0)

            # Ожидаемые колонки: Номенклатура, Размер, GTIN, КМ
            if len(df.columns) < 4:
                raise ValueError(f"Ожидается минимум 4 колонки, получено {len(df.columns)}")

            df.columns = ['Номенклатура', 'Размер', 'GTIN', 'КМ']

            # Извлекаем артикулы
            df['Артикул'] = df['Номенклатура'].apply(self._extract_article)

            # Проверяем успешность извлечения
            failed = df['Артикул'].isna().sum()
            if failed > 0:
                logger.warning(f"Не удалось извлечь артикул для {failed} записей")

            # Строим индекс: (артикул, размер) -> [км1, км2, ...]
            for _, row in df.iterrows():
                article = row['Артикул']
                size = row['Размер']
                km = row['КМ']

                if article and pd.notna(km):
                    key = (article, int(size))
                    if key not in self._index:
                        self._index[key] = []
                    self._index[key].append(str(km))

            logger.info(f"Загружено {len(df)} записей, {len(self._index)} уникальных комбинаций артикул+размер")

            # Статистика по артикулам
            unique_articles = df['Артикул'].nunique()
            logger.info(f"Уникальных артикулов: {unique_articles}")

        except Exception as e:
            logger.error(f"Ошибка загрузки файла КМ: {e}")
            raise

    def get_km_codes(self, article: str, sizes: List[int]) -> List[str]:
        """
        Получает все КМ коды для артикула по указанным размерам.

        Args:
            article: артикул товара (например 'GL24136-1')
            sizes: список размеров (например [35, 36, 37] для ≤24см)

        Returns:
            Список КМ кодов для всех указанных размеров
        """
        result = []
        for size in sizes:
            key = (article, size)
            if key in self._index:
                result.extend(self._index[key])
        return result

    def get_km_codes_for_category(self, article: str, insole_category: str) -> List[str]:
        """
        Получает КМ коды для артикула по категории длины стельки.

        Args:
            article: артикул товара
            insole_category: категория стельки ('длина стельки до 24см' или 'длина стельки более 24см')

        Returns:
            Список КМ кодов
        """
        if "до 24см" in insole_category:
            sizes = [35, 36, 37]
        elif "более 24см" in insole_category:
            sizes = [38, 39, 40, 41]
        else:
            # Неизвестная категория — берём все размеры
            sizes = [35, 36, 37, 38, 39, 40, 41]

        return self.get_km_codes(article, sizes)

    def get_all_km_codes(self, article: str) -> List[str]:
        """
        Получает все КМ коды для артикула (все размеры).

        Args:
            article: артикул товара

        Returns:
            Список всех КМ кодов для артикула
        """
        return self.get_km_codes(article, [35, 36, 37, 38, 39, 40, 41])

    @property
    def articles(self) -> List[str]:
        """Возвращает список уникальных артикулов"""
        return list(set(article for article, _ in self._index.keys()))

    @property
    def total_codes(self) -> int:
        """Возвращает общее количество КМ кодов"""
        return sum(len(codes) for codes in self._index.values())
