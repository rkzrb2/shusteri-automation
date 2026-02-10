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

    def _normalize_article(self, article: str) -> str:
        """
        Нормализует артикул — убирает известные суффиксы материала упаковки.

        Примеры:
            'GL24136-1 PRG' -> 'GL24136-1'
            'GL24136-1PRG' -> 'GL24136-1'
            'Y5286N004-F002AL' -> 'Y5286N004-F002AL' (AL не убирается, это часть артикула)
        """
        if not article:
            return article

        # Список известных суффиксов упаковки (можно расширять)
        packaging_suffixes = ['PRG', 'BOX', 'PKG', 'CTN', 'PCS']

        # Убираем только известные суффиксы (с пробелом или без)
        article_stripped = str(article).strip()
        for suffix in packaging_suffixes:
            # Паттерн: опциональный пробел + суффикс в конце строки (case-insensitive)
            pattern = r'\s*' + suffix + r'$'
            article_stripped = re.sub(pattern, '', article_stripped, flags=re.IGNORECASE)

        return article_stripped

    def _extract_article(self, nomenclature: str) -> Optional[str]:
        """
        Извлекает артикул из строки номенклатуры.

        Поддерживаемые форматы:
            'Туфли женские арт.GL24136-1(...)' -> 'GL24136-1'
            'Полуботинки женские арт.BF25086-1(...)' -> 'BF25086-1'
            'Ботинки арт.A226C3305-7-1(...)' -> 'A226C3305-7-1'
            'Туфли арт.CA23C1030-153(...)' -> 'CA23C1030-153'
        """
        # Основной паттерн: ищем "арт." и берём всё до открывающей скобки
        # Универсальный подход: любая комбинация букв, цифр и дефисов после "арт."
        match = re.search(r'арт\.[\s\W]*([A-Z0-9][A-Z0-9\-]+?)(?:\s*[\(,]|\s+[А-Я])', str(nomenclature), re.IGNORECASE)
        if match:
            article = match.group(1).strip()
            # Убираем trailing дефисы если есть
            article = article.rstrip('-')
            return article

        # Запасной вариант: точка + артикул (старая логика)
        match = re.search(r'\.[\s\W]*([A-Z]{1,4}[0-9\-]+[A-Za-z0-9\-]*)', str(nomenclature))
        if match:
            return match.group(1)

        # Последний шанс: ищем любой паттерн похожий на артикул
        # 1-2 буквы, затем цифры, буквы, дефисы
        match = re.search(r'([A-Z]{1,2}[0-9]{2,}[A-Z]?[0-9\-]+[A-Z0-9]*)', str(nomenclature))
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
            article: артикул товара (например 'GL24136-1' или 'GL24136-1 PRG')
            sizes: список размеров (например [35, 36, 37] для ≤24см)

        Returns:
            Список КМ кодов для всех указанных размеров
        """
        # Нормализуем артикул — убираем суффиксы типа ' PRG'
        normalized_article = self._normalize_article(article)

        result = []
        for size in sizes:
            key = (normalized_article, size)
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
