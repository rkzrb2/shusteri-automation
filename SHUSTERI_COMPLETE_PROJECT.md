# SHUSTERI AUTOMATION - ПОЛНАЯ ДОКУМЕНТАЦИЯ И ИСХОДНЫЙ КОД

---

## 📚 ЧАСТЬ 1: ТЕХНИЧЕСКАЯ ДОКУМЕНТАЦИЯ

# Shusteri Automation - Техническая документация проекта

## 📋 Оглавление
1. [Обзор проекта](#обзор-проекта)
2. [Архитектура и поток данных](#архитектура-и-поток-данных)
3. [Структура проекта](#структура-проекта)
4. [Модели данных](#модели-данных)
5. [Бизнес-логика](#бизнес-логика)
6. [Модули и компоненты](#модули-и-компоненты)
   - 6.1. [Интеграция с системой "Честный знак"](#6-km_loaderpy---загрузчик-кодов-маркировки-честный-знак)
7. [Конфигурация](#конфигурация)
8. [Форматы документов](#форматы-документов)
9. [Особенности реализации](#особенности-реализации)
10. [Примеры данных](#примеры-данных)

---

## Обзор проекта

### Назначение
Система автоматизации генерации экспортных документов для обувной компании Shusteri. Преобразует китайскую таблицу поставок в три профессионально оформленных документа: Commercial Invoice, Specification и Packing List.

### Ключевые возможности
- Генерация трех типов документов (Invoice, Specification, Packing List)
- Поддержка пресетов клиентов (мультиклиентность)
- Автоматическое разделение товаров по категориям длины стельки (≤24см / >24см)
- Умная обработка кодов ТН ВЭД (одинаковые/разные коды)
- **Интеграция с системой "Честный знак"** - автоматическое добавление КМ кодов маркировки
- Профессиональное форматирование Excel документов
- Интерактивный CLI интерфейс
- Поддержка двух режимов: загрузка контейнера и отправка грузов (полупары)

### Технологический стек
```
Python 3.8+
├── pandas 2.1.4          # Работа с данными
├── openpyxl 3.1.2        # Чтение/запись Excel
├── xlsxwriter 3.1.9      # Дополнительное форматирование
├── rich 13.7.0           # Красивый CLI интерфейс
├── pydantic 2.5.0        # Валидация данных
├── PyYAML 6.0.1          # Конфигурация
└── click 8.1.7           # CLI инструменты
```

---

## Архитектура и поток данных

### Общая схема
```
┌─────────────────┐
│  Входной Excel  │  (китайская таблица с товарами)
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  InputFileParser│  (парсинг входных данных)
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│  DataProcessor  │  (бизнес-логика обработки)
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│   Generators    │  (создание структуры документов)
│  ├─Invoice      │
│  ├─Specification│
│  └─PackingList  │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│   Formatters    │  (профессиональное оформление)
│  ├─Invoice      │
│  ├─Specification│
│  └─PackingList  │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│ Выходные Excel  │  (3 файла или 1 файл с 3 листами)
└─────────────────┘
```

### Детальный поток обработки одного товара

```
Входная строка:
АРТИКУЛ: "ABC123"
КОД: "6403911100/6403911800"  ← РАЗНЫЕ коды
Размеры: 35(10шт), 36(20шт), 37(15шт), 38(25шт), 39(30шт)

     ↓ PARSER

ProductLine:
- article: "ABC123"
- hs_code_le24: "6403911100"
- hs_code_gt24: "6403911800"
- qty_by_size: {35:10, 36:20, 37:15, 38:25, 39:30}

     ↓ PROCESSOR (определяет: коды РАЗНЫЕ → создать 2 строки)

OutputLine #1:
- article: "ABC123"
- hs_code: "6403911100"
- insole_category: "длина стельки до 24см"
- quantity: 45 (10+20+15)
- boxes: 5

OutputLine #2:
- article: "ABC123"
- hs_code: "6403911800"
- insole_category: "длина стельки более 24см"
- quantity: 55 (25+30)
- boxes: 0  ← специально!

     ↓ GENERATOR (создает документ)

Таблица в Excel:
1  │ BRAND │ 6403911100 │ ABC123 │ описание + до 24см     │ 45  │ ... │ 5   │
   │       │ 6403911800 │        │ описание + более 24см  │ 55  │ ... │  ↑  │
                                                                        │ объединены │
```

---

## Структура проекта

```
shusteri-automation/
├── src/                          # Исходный код
│   ├── __init__.py
│   ├── models.py                 # Pydantic модели данных
│   ├── parser.py                 # Парсер входных Excel файлов
│   ├── processor.py              # Бизнес-логика обработки
│   ├── generators/               # Генераторы документов
│   │   ├── __init__.py
│   │   ├── invoice.py           # Генератор Invoice
│   │   ├── specification.py     # Генератор Specification
│   │   └── packing_list.py      # Генератор Packing List
│   └── formatters/               # Форматтеры Excel
│       ├── __init__.py
│       ├── invoice_formatter.py
│       ├── specification_formatter.py
│       └── packing_list_formatter.py
├── presets/                      # YAML пресеты клиентов
│   └── client_example.yaml
├── input/                        # Входные Excel файлы
├── output/                       # Выходные документы
├── config.yaml                   # Общая конфигурация
├── requirements.txt              # Python зависимости
├── run.py                        # Главный файл запуска
└── automation.log                # Лог файл
```

---

## Модели данных

### ProductLine (входная строка)
```python
class ProductLine(BaseModel):
    """Одна строка товара из входного файла"""
    row_number: int                    # Номер строки в Excel
    brand: str                         # Марка (例: "BETSY")
    article: str                       # Артикул (例: "998089/05-01")
    code_raw: str                      # Код в сыром виде (例: "6403911100/6403911800")
    hs_code_le24: str                  # HS код для ≤24см
    hs_code_gt24: str                  # HS код для >24см
    name: str                          # Наименование товара
    material: str                      # Материал верха
    color: str                         # Цвет
    lining: str                        # Подкладка
    sole: str                          # Подошва
    heel_height: str                   # Высота каблука
    composition: Optional[str]         # Процентный состав
    price: Decimal                     # Цена за пару
    boxes: int                         # Количество коробок
    net_weight_per_pair: Decimal       # Вес нетто на пару
    gross_weight_per_box: Decimal      # Вес брутто за коробку
    qty_by_size: dict                  # {35: 100, 36: 200, ...}
    
    # Вычисляемые свойства:
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
```

### OutputLine (строка для документа)
```python
class OutputLine(BaseModel):
    """Строка для выходных документов"""
    brand: str                         # Марка
    hs_code: str                       # Код ТН ВЭД (один!)
    article: str                       # Артикул
    description: str                   # Описание
    color: str                         # Цвет
    material: str                      # Материал верха
    lining: str                        # Подкладка
    sole: str                          # Подошва
    heel_height: str                   # Высота каблука
    insole_category: str               # "длина стельки до 24см" ИЛИ 
                                       # "длина стельки более 24см" ИЛИ
                                       # "все размеры 35-39"
    quantity: int                      # Количество пар в этой категории
    net_weight: Decimal                # Вес нетто (пропорционально)
    gross_weight: Decimal              # Вес брутто (пропорционально)
    boxes: int                         # Коробки (0 для второй строки пары)
    original_boxes: int                # ВАЖНО: исходное значение для объединения
    price: Decimal                     # Цена за пару
    amount: Decimal                    # Сумма (price * quantity)
    kiz_codes: List[str]               # КИЗ коды (пустой список обычно)
```

### DocumentMetadata (метаданные документа)
```python
class DocumentMetadata(BaseModel):
    """Метаданные для всех документов"""
    invoice_number: str                # Номер инвойса
    date: str                          # Дата
    container_number: str              # Номер контейнера (опционально)
    seller_name: str                   # Продавец (RU)
    seller_name_en: str                # Продавец (EN)
    seller_address: str                # Адрес продавца (RU)
    seller_address_en: str             # Адрес продавца (EN)
    buyer_name: str                    # Покупатель
    buyer_address: str                 # Адрес покупателя (RU)
    buyer_address_en: str              # Адрес покупателя (EN)
    contract_number: str               # Номер контракта
    contract_date: str                 # Дата контракта
    terms_of_delivery: str             # Условия поставки (例: "FCA")
    currency: str                      # Валюта (例: "CNY")
```

---

## Бизнес-логика

### 1. Таблица размеров и длины стельки
```yaml
size_to_insole:
  "35": 22.5  # ≤24см
  "36": 23.0  # ≤24см
  "37": 24.0  # ≤24см (граница)
  "38": 24.5  # >24см
  "39": 25.0  # >24см
  "40": 25.5  # >24см
  "41": 26.0  # >24см
  "42": 26.5  # >24см

insole_boundary: 24.0  # Граница категорий (СТРОГО)
```

### 2. Логика разделения по кодам ТН ВЭД (КРИТИЧНО!)

```python
def process(products: List[ProductLine]) -> List[OutputLine]:
    output_lines = []
    
    for product in products:
        # ПРОВЕРКА: одинаковые ли коды?
        codes_are_same = (product.hs_code_le24 == product.hs_code_gt24)
        
        if codes_are_same:
            # ВАРИАНТ 1: Код ОДИН → создаем ОДНУ строку со ВСЕМИ размерами
            output_lines.append(create_line(
                product,
                category="all",
                quantity=product.total_pairs
            ))
        else:
            # ВАРИАНТ 2: Коды РАЗНЫЕ → создаем ДВЕ строки
            
            # Строка 1: ≤24см
            if product.pairs_le24 > 0:
                output_lines.append(create_line(
                    product,
                    category="le24",
                    quantity=product.pairs_le24
                ))
            
            # Строка 2: >24см
            if product.pairs_gt24 > 0:
                output_lines.append(create_line(
                    product,
                    category="gt24",
                    quantity=product.pairs_gt24
                ))
    
    return output_lines
```

**КЛЮЧЕВОЙ МОМЕНТ:**
- Если во входном файле код написан БЕЗ слэша (например: `6403911100`) → создается ОДНА строка
- Если код написан СО слэшем (например: `6403911100/6403911800`) → создаются ДВЕ строки

### 3. Объединение ячеек boxes

**Правило:** Когда создаются ДВЕ строки для одного артикула (разные коды ТН ВЭД):
1. В первой строке (≤24см): `boxes = N`, где N - исходное количество коробок из файла
2. Во второй строке (>24см): `boxes = 0`
3. **При форматировании:** эти две ячейки объединяются вертикально, и в объединенной ячейке пишется `original_boxes`

**Пример:**
```
Входной файл:
АРТИКУЛ: ABC123, КОД: 6403911100/6403911800, КОРОБКИ: 10

После обработки (2 строки):
Line 1: boxes=10, original_boxes=10
Line 2: boxes=0,  original_boxes=10

В Excel (после объединения):
│ ABC123 │ ... │ ┌──┐ │
│        │ ... │ │10│ │  ← Объединенная ячейка
                └──┘
```

### 4. Подсчет итогов (Total)

**КРИТИЧНО:** При подсчете `total_boxes` НЕ дублируем коробки для второй строки пары!

```python
total_boxes = 0

for line in lines:
    # Проверяем: является ли это второй строкой пары?
    is_continuation = (
        prev_line is not None and
        line.article == prev_line.article and
        "более 24см" in line.insole_category and
        line.boxes == 0 and
        "до 24см" in prev_line.insole_category
    )
    
    # Добавляем boxes только для первых строк!
    if not is_continuation:
        total_boxes += line.boxes
```

---

## Модули и компоненты

### 1. models.py - Модели данных

**Назначение:** Определение структур данных с валидацией через Pydantic

**Классы:**
- `ProductLine` - входная строка из Excel файла
- `OutputLine` - строка для выходных документов
- `DocumentMetadata` - метаданные документа (инвойс, контракт и т.д.)

**Ключевые особенности:**
- Использование `Decimal` для финансовых расчетов (точность!)
- Вычисляемые свойства (@property) для общего количества пар
- Автоматическое разделение пар по категориям стельки

### 2. parser.py - Парсер входных файлов

**Назначение:** Чтение китайской таблицы Excel и преобразование в `ProductLine`

**Основной метод:** `parse(file_path: str) -> List[ProductLine]`

**Логика работы:**
1. Читает Excel файл через pandas
2. Удаляет пустые и итоговые строки
3. Для каждой строки:
   - Парсит код ТН ВЭД (разделяет по `/` если есть)
   - Собирает количество по размерам (35-42)
   - Создает объект `ProductLine`
4. Логирует результаты

**Обработка кодов ТН ВЭД:**
```python
def _parse_hs_code(self, code_raw: str) -> Tuple[str, str]:
    if '/' in code_raw:
        # Разные коды для ≤24см и >24см
        parts = code_raw.split('/')
        return parts[0].strip(), parts[1].strip()
    else:
        # Один код для всех размеров
        return code_raw.strip(), code_raw.strip()
```

### 3. processor.py - Обработчик данных

**Назначение:** Преобразование `ProductLine` → `OutputLine` с учетом бизнес-логики

**Основной метод:** `process(products: List[ProductLine]) -> List[OutputLine]`

**Логика разделения строк:**

```python
for product in products:
    codes_are_same = (product.hs_code_le24 == product.hs_code_gt24)
    
    if codes_are_same:
        # ОДИН код → ОДНА строка со всеми размерами
        output_line = create_line(
            product,
            category="all",
            quantity=product.total_pairs
        )
    else:
        # РАЗНЫЕ коды → ДВЕ строки (≤24см и >24см)
        line_le24 = create_line(product, category="le24", ...)
        line_gt24 = create_line(product, category="gt24", ...)
```

**Расчет весов:**
- Net weight: `net_weight_per_pair * quantity`
- Gross weight: пропорционально от общего брутто за все коробки

**Важно:** `boxes` устанавливается в 0 для второй строки пары, но `original_boxes` всегда содержит исходное значение!

### 4. generators/ - Генераторы документов

Три класса для генерации документов:
- `InvoiceGenerator` - Commercial Invoice (11 колонок)
- `SpecificationGenerator` - Specification (18 колонок)
- `PackingListGenerator` - Packing List (11 колонок)

**Общая структура метода generate():**

1. **Создание шапки документа:**
   - Название и адрес продавца
   - Телефон
   - Заголовок документа
   - Информация о покупателе
   - Контракт
   - Условия поставки
   - Номер контейнера

2. **Создание заголовка таблицы** (зависит от типа документа)

3. **Заполнение данных:**
   - Проход по всем `OutputLine`
   - Определение "продолжения" (вторая строка пары)
   - Нумерация только для первых строк
   - Пустые ячейки (Brand, Article) для вторых строк

4. **Подсчет итогов:**
   - `total_qty`, `total_net`, `total_gross`
   - **ВАЖНО:** `total_boxes` без дублирования для вторых строк!
   - `total_amount`

5. **Создание футера:**
   - Строка "Total / Итого"
   - Дополнительная информация (производитель, страна и т.д.)
   - **Только для Specification:** подписи продавца и покупателя

6. **Сохранение в Excel через pandas**

7. **Объединение ячеек boxes** для пар строк:
```python
if is_continuation:
    ws.merge_cells(f'I{current_row-1}:I{current_row}')
    ws[f'I{current_row-1}'] = line.original_boxes
```

8. **Применение форматирования** через соответствующий форматтер

### 5. formatters/ - Форматтеры документов

Три класса для форматирования:
- `InvoiceFormatter`
- `SpecificationFormatter`
- `PackingListFormatter`

**Основные задачи:**
- Установка ширины колонок
- Форматирование шапки (шрифты, размеры, выравнивание)
- Форматирование таблицы (границы, выравнивание, шрифты)
- Форматирование футера

**Общая структура метода format():**

```python
def format(self, file_path: str, data_start_row: int, data_end_row: int):
    wb = load_workbook(file_path)
    ws = wb.active
    
    self._set_column_widths(ws)
    self._format_header(ws)
    self._format_table(ws, data_start_row, data_end_row)
    self._format_footer(ws, data_end_row)
    
    wb.save(file_path)
    wb.close()  # ВАЖНО: Закрываем файл!
```

**Особенности форматирования:**

1. **Ширина колонок:** Точные значения в пунктах для каждой колонки
2. **Шрифты:** Arial с разными размерами (9-16)
3. **Выравнивание:**
   - Центр для номеров, кодов, артикулов
   - Влево для описаний
   - Вправо для чисел
4. **Границы:** Тонкие линии для всех ячеек таблицы
5. **Высота строк:** Настраивается индивидуально

### 6. km_loader.py - Загрузчик кодов маркировки "Честный знак"

**Назначение:** Модуль для загрузки и обработки выгрузок КМ (кодов маркировки) из системы "Честный знак" для обогащения документов уникальными кодами товаров.

**Основные возможности:**
1. **Парсинг файлов выгрузки** - извлечение артикулов из номенклатуры
2. **Нормализация артикулов** - удаление суффиксов (PRG, BOX и т.п.)
3. **Индексирование данных** - построение индекса `(артикул, размер) → [КМ коды]`
4. **Поиск кодов** - получение КМ кодов по артикулу и категории стельки

**Структура класса KMLoader:**

```python
class KMLoader:
    def __init__(self, file_path: str):
        """Загружает файл выгрузки и строит индекс"""
        self.file_path = Path(file_path)
        self._index: Dict[Tuple[str, int], List[str]] = {}
        self._load()

    def _normalize_article(self, article: str) -> str:
        """Убирает суффиксы типа ' PRG', ' BOX' из артикула"""
        # GL24136-1 PRG -> GL24136-1

    def _extract_article(self, nomenclature: str) -> Optional[str]:
        """
        Извлекает артикул из строки номенклатуры.

        Поддерживаемые форматы:
        - 'Туфли женские арт.GL24136-1(...)' -> 'GL24136-1'
        - 'Ботинки арт.A226C3305-7-1(...)' -> 'A226C3305-7-1'
        - 'Туфли арт.CA23C1030-153(...)' -> 'CA23C1030-153'

        Универсальный паттерн с поддержкой:
        - Старых артикулов: GL24136-1, BF25086-1
        - Новых артикулов: A226C3305-7-1, CA23C1030-153
        """

    def get_km_codes(self, article: str, sizes: List[int]) -> List[str]:
        """Получает КМ коды для артикула по размерам"""

    def get_km_codes_for_category(self, article: str, insole_category: str) -> List[str]:
        """Получает КМ коды по категории стельки (≤24см / >24см)"""
```

**Формат файла выгрузки:**

| Колонка | Название | Пример | Описание |
|---------|----------|--------|----------|
| A | Номенклатура | `Туфли женские арт.GL24136-1(натуральная кожа)` | Полное название с артикулом |
| B | Размер | `36` | Размер обуви |
| C | GTIN | `4630083360812` | Global Trade Item Number |
| D | КМ | `0104630083360812215_ijUKGGz/gsJ` | Код маркировки |

**Логика извлечения артикулов:**

```python
# Основной паттерн - универсальный
r'арт\.[\s\W]*([A-Z0-9][A-Z0-9\-]+?)(?:\s*[\(,]|\s+[А-Я])'

# Примеры извлечения:
"Туфли женские арт.GL24136-1(нат. кожа)" -> "GL24136-1"
"Ботинки арт.A226C3305-7-1(черная)" -> "A226C3305-7-1"
"Туфли арт.CA23C1030-153(кожа)" -> "CA23C1030-153"
"Туфли арт.СBF25086-1(серый)" -> "BF25086-1"  # кириллическая С
```

**Интеграция в run.py:**

```python
# Новые методы в ShusteriAutomation:

def get_km_files(self):
    """Получает список файлов из папки 'выгрузка честный знак'"""

def select_km_file(self):
    """Интерактивный выбор файла маркировки"""
    # 🏷️  Использовать файл маркировки (Честный знак)? [y/n]

def enrich_with_km_codes(self, output_lines, km_loader: KMLoader):
    """Обогащает output_lines кодами маркировки"""
    for line in output_lines:
        km_codes = km_loader.get_km_codes_for_category(
            line.article,
            line.insole_category
        )
        if km_codes:
            line.kiz_codes = km_codes
```

**Поток обработки с КМ кодами:**

```
1. Выбор файла КМ (опционально)
   └─> Интерактивный выбор из папки "выгрузка честный знак/"

2. Загрузка справочника
   └─> KMLoader парсит файл и строит индекс

3. Обработка входных данных
   └─> Обычный процесс (parser -> processor -> output_lines)

4. Обогащение КМ кодами
   └─> Для каждой OutputLine:
       - Берём артикул (нормализованный)
       - Определяем размеры по insole_category (≤24см = [35,36,37])
       - Получаем КМ коды из индекса
       - Записываем в line.kiz_codes

5. Генерация документов
   └─> Specification: колонка КИЗ (R)
       - Коды записываются через перенос строки (\n)
       - Каждый код на новой строке внутри ячейки
```

**Пример обогащения:**

```python
# Входная строка
output_line = OutputLine(
    article="GL24136-1",
    insole_category="длина стельки до 24см",  # размеры 35-37
    quantity=45,
    ...
)

# После обогащения
km_codes = km_loader.get_km_codes_for_category("GL24136-1", "длина стельки до 24см")
# → ['0104630083369471215...', '0104630083360812215...', ...]  (35 кодов)

output_line.kiz_codes = km_codes

# В Specification (колонка R):
cell_value = '\n'.join(km_codes)
# Результат в Excel:
# 0104630083369471215dbn9ieddgTn-
# 0104630083360812215_ijUKGGz/gsJ
# 0104630083360812215*6zCRuqR5SQQ
# ... (каждый код на новой строке)
```

**Особенности:**

1. **Нормализация артикулов:**
   - `GL24136-1 PRG` → `GL24136-1` (убирается суффикс упаковки)
   - Позволяет находить КМ коды для артикулов с разными суффиксами

2. **Универсальность паттернов:**
   - Поддержка старых форматов: `GL24136-1`, `BF25086-1`
   - Поддержка новых форматов: `A226C3305-7-1`, `CA23C1030-153`
   - Обработка кириллических символов в артикулах

3. **Группировка по категории стельки:**
   - `≤24см` → размеры [35, 36, 37]
   - `>24см` → размеры [38, 39, 40, 41]

4. **Формат вывода:**
   - КМ коды разделяются переносом строки `\n`
   - Excel автоматически применяет Wrap Text
   - Каждый код визуально на отдельной строке

---

## Конфигурация

### config.yaml - Общая конфигурация

```yaml
# Таблица соответствия EU размеров и длины стельки
size_to_insole:
  "35": 22.5
  "36": 23.0
  "37": 24.0
  "38": 24.5
  "39": 25.0
  "40": 25.5
  "41": 26.0
  "42": 26.5

# Граница категорий
insole_boundary: 24.0

# Маппинг колонок входного файла
input_columns:
  article: "АРТИКУЛ"
  size_range: "РАЗМЕР"
  brand: "МАРКА"
  code: "КОД"
  name: "НАИМЕНОВАНИЕ"
  material: "МАТЕРИАЛ ВЕРХА"
  color: "ЦВЕТ"
  lining: "ПОДКЛАДКА"
  sole: "ПОДОШВА"
  heel_height: "Высота каблука"
  price: "ЦЕНА"
  boxes: "КОРОБКИ ЗАГРУЖЕНЫ"
  net_weight_per_pair: "ВЕС НЕТТО НА ПАР"
  gross_weight_per_box: "ВЕС БРУТТО ЗА КОРОБКУ"
  composition: "процентный состав"
  qty_35: "кол.35 размер"
  qty_36: "кол.36 размер"
  # ... и т.д. до 42 размера
```

### Пресеты клиентов (presets/*.yaml)

Пресеты содержат данные конкретного клиента:

```yaml
preset_name: "Клиент 1"
description: "Описание клиента"

seller:
  name: "ООО РОМАШКА"
  name_en: "ROMASHKA LLC"
  address: "Адрес продавца на русском"
  address_en: "Seller address in English"
  phone: "+7 123 456 78 90"

buyer:
  name: "ООО ПОКУПАТЕЛЬ"
  address: "Адрес покупателя на русском"
  address_en: "Buyer address in English"

contract:
  number: "12345/2024"
  date: "01.01.2024"

delivery:
  terms: "FCA Guangzhou"
  currency: "CNY"
  country_of_origin: "Китай"
  country_of_origin_en: "China"
```

---

## Форматы документов

### 1. Commercial Invoice (11 колонок)

| № | Brand | Code | Factory code | Description | Qty | Net | Gross | Boxes | Price | Amount |
|---|-------|------|--------------|-------------|-----|-----|-------|-------|-------|--------|

**Особенности:**
- Нет Terms of delivery в футере (только в шапке)
- Нет подписей
- Boxes в колонке I

### 2. Specification (18 колонок)

| № | Brand | Code | Factory code | Description | Color | Material | Lining | Sole | Heel | Insole | Qty | Net | Gross | Boxes | Price | Amount | KIZ |
|---|-------|------|--------------|-------------|-------|----------|--------|------|------|--------|-----|-----|-------|-------|-------|--------|-----|

**Особенности:**
- Самый широкий документ (18 колонок)
- Есть Terms of delivery в футере
- Есть подписи продавца и покупателя
- Boxes в колонке O

### 3. Packing List (11 колонок)

| № | Brand | Code | Factory code | Description | Color | Qty | Net | Gross | Boxes | Type |
|---|-------|------|--------------|-------------|-------|-----|-----|-------|-------|------|

**Особенности:**
- Есть колонка "Type of packaging" (картонная коробка)
- Нет цен и сумм
- Boxes в колонке J

---

## Особенности реализации

### 1. Обработка файлов Excel

**Проблема:** Windows блокирует файлы Excel при открытии

**Решение:**
```python
# ПРАВИЛЬНО - используем context manager:
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    df.to_excel(writer, index=False, header=False)

# После работы с openpyxl:
wb.save(file_path)
wb.close()  # ОБЯЗАТЕЛЬНО!
```

### 2. Объединение ячеек boxes для парных строк

**Алгоритм:**
```python
prev_line = None
current_row = data_start_row

for line in lines:
    is_continuation = (определение...)
    
    if is_continuation:
        # Объединяем с предыдущей строкой
        ws.merge_cells(f'I{current_row-1}:I{current_row}')  # или J, или O
        ws[f'I{current_row-1}'] = line.original_boxes  # ← КЛЮЧ!
    elif line.boxes > 0:
        # Обычная строка
        ws[f'I{current_row}'] = line.boxes
    
    current_row += 1
    prev_line = line
```

### 3. Подсчет total_boxes

**ВАЖНО:** Не дублировать коробки для второй строки пары!

```python
for line in lines:
    is_continuation = (определение...)
    
    total_qty += line.quantity
    total_net += line.net_weight
    total_gross += line.gross_weight
    
    # КРИТИЧНО:
    if not is_continuation:
        total_boxes += line.boxes  # Только для первой строки!
    
    total_amount += line.amount
```

### 4. Форматирование сумм

```python
# CNY с пробелами вместо запятых:
f"¥{total_amount:,.2f}".replace(',', ' ')

# Веса с 3 знаками:
round(float(total_net), 3)
```

### 5. Интерактивный CLI

```python
from rich.console import Console
from rich.prompt import Prompt, Confirm
from rich.table import Table

console = Console()

# Таблица с вариантами
table = Table(show_header=True, header_style="bold magenta")
table.add_column("№", style="cyan")
table.add_column("Клиент", style="green")
console.print(table)

# Выбор
choice = Prompt.ask("Выберите пресет", choices=["1", "2", "3"])

# Подтверждение
if Confirm.ask("Создать еще один документ?", default=False):
    # ...
```

---

## Примеры данных

### Пример входной строки Excel

| АРТИКУЛ | РАЗМЕР | МАРКА | КОД | НАИМЕНОВАНИЕ | МАТЕРИАЛ ВЕРХА | ЦВЕТ | ... | кол.35 размер | кол.36 размер | ... |
|---------|--------|-------|-----|--------------|----------------|------|-----|---------------|---------------|-----|
| 998089/05-01 | 35-39 | BETSY | 6403911100/6403911800 | Ботинки женские | натуральная кожа | черный | ... | 10 | 20 | ... |

### Пример преобразования

**После Parser:**
```python
ProductLine(
    article="998089/05-01",
    brand="BETSY",
    code_raw="6403911100/6403911800",
    hs_code_le24="6403911100",
    hs_code_gt24="6403911800",
    qty_by_size={35: 10, 36: 20, 37: 15, 38: 25, 39: 30},
    total_pairs=100,
    pairs_le24=45,  # 10+20+15
    pairs_gt24=55,  # 25+30
    boxes=10,
    # ...
)
```

**После Processor (2 строки, т.к. коды разные):**
```python
# Строка 1:
OutputLine(
    article="998089/05-01",
    hs_code="6403911100",
    insole_category="длина стельки до 24см",
    quantity=45,
    boxes=10,
    original_boxes=10,
    # ...
)

# Строка 2:
OutputLine(
    article="998089/05-01",
    hs_code="6403911800",
    insole_category="длина стельки более 24см",
    quantity=55,
    boxes=0,  # ← ВАЖНО!
    original_boxes=10,  # ← ВАЖНО!
    # ...
)
```

**В документе Invoice:**
```
№ │ Brand │ Code       │ Article      │ Description              │ Qty │ ... │ Boxes │
1 │ BETSY │ 6403911100 │ 998089/05-01 │ ..., до 24см             │ 45  │ ... │       │
  │       │ 6403911800 │              │ ..., более 24см          │ 55  │ ... │  10   │
                                                                          │   ↑   │
                                                                          └───────┘
                                                                        объединены
```

### Пример с одинаковыми кодами

**Входная строка:**
```
АРТИКУЛ: "ABC123"
КОД: "6403911100"  ← БЕЗ слэша! (один код)
Размеры: 35(10), 36(20), 37(15), 38(25), 39(30)
```

**После Processor (1 строка):**
```python
OutputLine(
    article="ABC123",
    hs_code="6403911100",  # один код для всех
    insole_category="все размеры 35-39",  # НЕ "до/более 24см"!
    quantity=100,  # все пары вместе
    boxes=10,
    # ...
)
```

**В документе:**
```
№ │ Brand │ Code       │ Article │ Description           │ Qty │ ... │ Boxes │
1 │ BETSY │ 6403911100 │ ABC123  │ ..., все размеры 35-39│ 100 │ ... │  10   │
```

---

## Запуск и использование

### Команда запуска
```bash
python run.py
```

### Интерактивный процесс
```
1. Выбор пресета клиента (из presets/)
2. Выбор входного файла (из input/)
3. Ввод номера инвойса (по умолчанию: YYYYMMDD01)
4. Ввод номера контейнера (опционально)
5. Выбор формата вывода:
   - 1: Три отдельных файла
   - 2: Один файл с тремя листами
6. Обработка и генерация документов
7. Вывод статистики
8. Предложение создать еще один документ
```

### Результат
```
output/
├── Shusteri_2024010101_20240101_Invoice.xlsx
├── Shusteri_2024010101_20240101_Specification.xlsx
└── Shusteri_2024010101_20240101_PackingList.xlsx

ИЛИ

output/
└── Shusteri_2024010101_20240101_All.xlsx  (3 листа)
```

---

## Критические моменты для ИИ

### ⚠️ ВАЖНО понимать:

1. **Логика разделения строк зависит от кодов ТН ВЭД:**
   - Одинаковые → 1 строка
   - Разные → 2 строки

2. **Boxes только в первой строке пары:**
   - Вторая строка: boxes=0, но original_boxes=N
   - Ячейки объединяются вертикально

3. **Total boxes считается БЕЗ дублирования:**
   - Проверяем is_continuation перед добавлением

4. **Файлы нужно закрывать:**
   - wb.close() после сохранения

5. **Specification отличается от других:**
   - 18 колонок вместо 11
   - С Terms of delivery в футере
   - С подписями

6. **Нумерация строк:**
   - Вторая строка пары не имеет номера
   - Brand, Article тоже пустые для второй строки

---

## История изменений

### v1.2 — Форматирование документов по эталону (2026-02-11)

**Ветка:** `feature/minor-fixes`

#### Генераторы (контент):
1. **Все строки заполняются полностью** — убран паттерн `is_continuation`, теперь каждая строка получает свой номер и все поля (Brand, Article, Description, Color и т.д.) заполняются независимо от того, является ли строка "продолжением" предыдущего артикула с другим кодом ТНВЭД. Изменения в `invoice.py`, `specification.py`, `packing_list.py`.

2. **Шапка Specification переработана** — вместо 4 строк (3 строки текста + Container) теперь 2 строки:
   - Строка 1: `Спецификация к Контракту/ Specification to the Contract №{number} from/от {date}`
   - Строка 2: `Container No / Контейнер № {container}`
   - Текст размещён в колонке A (ранее был в колонке M)

3. **Container No в одну ячейку** — во всех документах (Invoice, Packing List) номер контейнера объединён в одну ячейку J:K вместо раздельных J и K.

4. **Английский адрес покупателя** — в Invoice и Packing List добавлен `buyer_address_en` из пресета. Формат: русский адрес + `\n(английский адрес)`.

#### Форматтеры (визуальное оформление):
5. **Specification formatter:**
   - Выравнивание шапки: `right` → `left`
   - Высоты строк: строка 1 = 17.4, строка 2 = 15.0, строка 8 = 22.8
   - Шрифт колонок A-D в данных: Arial 10 (было 9)
   - Высота строки "Payment of the cost..." в футере: 45

6. **Packing List formatter:**
   - Ширина колонки A: 3.33 → 7.55
   - Высота строк данных: 26.4
   - Container No: объединение J11:K11

7. **Invoice formatter:**
   - Жёлтая заливка (`PatternFill #FFFF00`) для строки Terms of delivery + Container
   - Container No: объединение J10:K10 с жёлтой заливкой
   - Высота строки "Payment of the cost..." в футере: 45

#### Пресет Shoesberry:
8. **Обновлён адрес покупателя** — новый адрес: Московская обл., г.о. Люберцы, пгт Томилино.

---

### v1.1 — Интеграция с системой "Честный знак" (2026-02-10)

**Ветка:** `feature/honest-mark`

- Создан модуль `src/km_loader.py` для загрузки КМ кодов из выгрузки "Честного знака"
- Интерактивный выбор файла с КМ кодами
- Обогащение данных Specification кодами маркировки (колонка R)
- Каждый КМ код на новой строке в ячейке (`'\n'.join()`)
- Нормализация артикулов (удаление суффиксов PRG, BOX, PKG, CTN, PCS)
- Универсальный regex для извлечения артикулов из номенклатуры

---

## Возможные улучшения

1. **Валидация данных:**
   - Проверка обязательных полей
   - Валидация HS кодов (10 цифр)
   - Проверка весов (net < gross)

2. **Тестирование:**
   - Unit тесты для processor.py
   - Интеграционные тесты
   - Тесты форматирования

3. **Обработка ошибок:**
   - Try-catch для каждого продукта
   - Детальные сообщения об ошибках
   - Rollback при критических ошибках

4. **Производительность:**
   - Batch обработка больших файлов
   - Асинхронная генерация документов
   - Кеширование пресетов

5. **Расширение функциональности:**
   - Поддержка других валют
   - Больше типов документов
   - Экспорт в PDF
   - Web интерфейс

---

**Конец технической документации**

_Версия: 1.2_
_Дата: 2026-02-11_
_Автор: Ростислав (Shusteri Automation Project)_

---
---
---

## 💻 ЧАСТЬ 2: ПОЛНЫЙ ИСХОДНЫЙ КОД ПРОЕКТА

### 📁 Структура проекта
```
shusteri-automation/
├── src/
│   ├── models.py
│   ├── parser.py
│   ├── processor.py
│   ├── generators/
│   │   ├── __init__.py
│   │   ├── invoice.py
│   │   ├── specification.py
│   │   └── packing_list.py
│   └── formatters/
│       ├── __init__.py
│       ├── invoice_formatter.py
│       ├── specification_formatter.py
│       └── packing_list_formatter.py
├── config.yaml
├── requirements.txt
├── README.md
└── run.py
```

---

### 📄 Файл: `config.yaml`

```yaml
# ===========================================
# КОНФИГУРАЦИЯ ДЛЯ ШУСТЕРИ
# ===========================================
# Технические настройки (размеры, колонки)
# Данные клиентов вынесены в папку presets/

# Таблица соответствия EU размеров и длины стельки
size_to_insole:
  "35": 22.5
  "36": 23.0
  "37": 24.0
  "38": 24.5
  "39": 25.0
  "40": 25.5
  "41": 26.0
  "42": 26.5

# Граница категорий (строго)
insole_boundary: 24.0

# Маппинг колонок входного файла
input_columns:
  article: "АРТИКУЛ"
  size_range: "РАЗМЕР"
  brand: "МАРКА"
  code: "КОД"
  name: "НАИМЕНОВАНИЕ"
  material: "МАТЕРИАЛ ВЕРХА"
  color: "ЦВЕТ"
  lining: "ПОДКЛАДКА"
  sole: "ПОДОШВА"
  heel_height: "Высота каблука"
  price: "ЦЕНА"
  boxes: "КОРОБКИ ЗАГРУЖЕНЫ"
  net_weight_per_pair: "ВЕС НЕТТО НА ПАР"
  gross_weight_per_box: "ВЕС БРУТТО ЗА КОРОБКУ"
  composition: "процентный состав"

  qty_35: "кол.35 размер"
  qty_36: "кол.36 размер"
  qty_37: "кол.37 размер"
  qty_38: "кол.38 размер"
  qty_39: "кол.39 размер"
  qty_40: "кол.40 размер"
  qty_41: "кол.41 размер"
  qty_42: "кол.42 размер"
```

---

### 📄 Файл: `requirements.txt`

```
pandas==2.1.4
openpyxl==3.1.2
xlsxwriter==3.1.9
click==8.1.7
rich==13.7.0
pydantic==2.5.0
PyYAML==6.0.1
```

---

### 📄 Файл: `README.md`

```markdown
# Shusteri Automation

Система автоматизации генерации документов (Invoice, Specification, Packing List) для компании Shusteri.

## Возможности

- Генерация Invoice (Коммерческий инвойс)
- Генерация Specification (Спецификация)
- Генерация Packing List (Упаковочный лист)
- Поддержка пресетов клиентов
- Автоматическое форматирование документов
- Разделение товаров по категориям длины стельки (≤24см / >24см)

## Установка

1. Установите зависимости:
```bash
pip install -r requirements.txt
```

2. Настройте пресеты клиентов в папке `presets/`

3. Поместите входные Excel файлы в папку `input/`

## Использование

Запустите программу:
```bash
python run.py
```

Следуйте интерактивным инструкциям:
1. Выберите пресет клиента
2. Выберите входной файл
3. Введите номер инвойса
4. Укажите номер контейнера (опционально)
5. Выберите формат вывода (3 файла или 1 файл с 3 листами)

Результаты будут сохранены в папке `output/`

## Структура проекта

```
shusteri-automation/
├── src/
│   ├── generators/       # Генераторы документов
│   ├── formatters/       # Форматирование Excel
│   ├── models.py         # Модели данных
│   ├── parser.py         # Парсер входных файлов
│   └── processor.py      # Обработка данных
├── presets/              # Пресеты клиентов
├── input/                # Входные файлы
├── output/               # Выходные файлы
├── config.yaml           # Конфигурация
└── run.py                # Главный файл запуска
```

## Требования

- Python 3.8+
- pandas
- openpyxl
- rich
- pydantic
- PyYAML

## Лицензия

Proprietary
```

---

### 📄 Файл: `run.py`

```python
#!/usr/bin/env python3
"""
Автоматизация документооборота для Шустери
Генерация Invoice, Specification и Packing List с форматированием

Интерактивный режим работы
"""

import yaml
from pathlib import Path
from datetime import datetime
import logging
from rich.console import Console
from rich.table import Table
from rich.prompt import Prompt, Confirm

from src.parser import InputFileParser
from src.processor import DataProcessor
from src.generators.invoice import InvoiceGenerator
from src.generators.specification import SpecificationGenerator
from src.generators.packing_list import PackingListGenerator
from src.models import DocumentMetadata

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('automation.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

console = Console()


class ShusteriAutomation:
    """Главный класс автоматизации"""

    def __init__(self, config_path: str = "config.yaml"):
        with open(config_path, 'r', encoding='utf-8') as f:
            self.config = yaml.safe_load(f)
        
        self.preset = None  # Текущий выбранный пресет
    
    def get_available_presets(self):
        """Получает список доступных пресетов"""
        presets_dir = Path("presets")
        if not presets_dir.exists():
            console.print("[bold red]❌ Папка presets/ не найдена![/bold red]")
            console.print("[yellow]Создайте папку presets/ и положите туда файлы пресетов[/yellow]")
            return []
        
        # Ищем YAML файлы
        preset_files = list(presets_dir.glob("*.yaml")) + list(presets_dir.glob("*.yml"))
        
        if not preset_files:
            console.print("[bold red]❌ Нет пресетов в папке presets/![/bold red]")
            return []
        
        # Загружаем информацию о пресетах
        presets = []
        for preset_file in sorted(preset_files):
            try:
                with open(preset_file, 'r', encoding='utf-8') as f:
                    preset_data = yaml.safe_load(f)
                    presets.append({
                        'file': preset_file,
                        'name': preset_data.get('preset_name', preset_file.stem),
                        'description': preset_data.get('description', ''),
                        'data': preset_data
                    })
            except Exception as e:
                logger.warning(f"Не удалось загрузить пресет {preset_file}: {e}")
        
        return presets
    
    def select_preset(self):
        """Интерактивный выбор пресета"""
        presets = self.get_available_presets()
        
        if not presets:
            return None
        
        # Показываем список пресетов
        console.print("\n[bold cyan]👥 Доступные пресеты клиентов:[/bold cyan]\n")
        
        table = Table(show_header=True, header_style="bold magenta")
        table.add_column("№", style="cyan", width=4)
        table.add_column("Клиент", style="green")
        table.add_column("Описание", style="yellow")
        
        for idx, preset in enumerate(presets, 1):
            table.add_row(
                str(idx),
                preset['name'],
                preset['description']
            )
        
        console.print(table)
        console.print()
        
        # Запрашиваем выбор
        if len(presets) == 1:
            console.print(f"[green]✓ Автоматически выбран пресет: {presets[0]['name']}[/green]\n")
            return presets[0]['data']
        else:
            choice = Prompt.ask(
                "Выберите пресет (введите номер)",
                choices=[str(i) for i in range(1, len(presets) + 1)],
                default="1"
            )
            selected_preset = presets[int(choice) - 1]
            console.print(f"[green]✓ Выбран пресет: {selected_preset['name']}[/green]\n")
            return selected_preset['data']

    def get_input_files(self):
        """Получает список файлов в папке input/"""
        input_dir = Path("input")
        if not input_dir.exists():
            console.print("[bold red]❌ Папка input/ не найдена![/bold red]")
            console.print("[yellow]Создайте папку input/ и положите туда Excel файлы[/yellow]")
            return []
        
        # Ищем Excel файлы
        excel_files = list(input_dir.glob("*.xlsx")) + list(input_dir.glob("*.xls"))
        excel_files = [f for f in excel_files if not f.name.startswith('~')]  # Исключаем временные файлы
        
        return sorted(excel_files)

    def select_input_file(self):
        """Интерактивный выбор входного файла"""
        files = self.get_input_files()
        
        if not files:
            console.print("[bold red]❌ Нет Excel файлов в папке input/![/bold red]")
            console.print("[yellow]Положите Excel файл в папку input/ и запустите программу снова[/yellow]")
            return None
        
        # Показываем список файлов
        console.print("\n[bold cyan]📁 Доступные файлы:[/bold cyan]\n")
        
        table = Table(show_header=True, header_style="bold magenta")
        table.add_column("№", style="cyan", width=4)
        table.add_column("Имя файла", style="green")
        table.add_column("Размер", style="yellow", justify="right")
        table.add_column("Дата изменения", style="blue")
        
        for idx, file in enumerate(files, 1):
            size_kb = file.stat().st_size / 1024
            mtime = datetime.fromtimestamp(file.stat().st_mtime).strftime('%d.%m.%Y %H:%M')
            table.add_row(str(idx), file.name, f"{size_kb:.1f} KB", mtime)
        
        console.print(table)
        console.print()
        
        # Запрашиваем выбор
        if len(files) == 1:
            # Если файл один - используем его автоматически
            console.print(f"[green]✓ Автоматически выбран файл: {files[0].name}[/green]\n")
            return files[0]
        else:
            # Если файлов несколько - даем выбрать
            choice = Prompt.ask(
                "Выберите файл (введите номер)",
                choices=[str(i) for i in range(1, len(files) + 1)],
                default="1"
            )
            selected_file = files[int(choice) - 1]
            console.print(f"[green]✓ Выбран файл: {selected_file.name}[/green]\n")
            return selected_file

    def get_invoice_number(self):
        """Запрос номера инвойса"""
        # Генерируем номер по умолчанию на основе текущей даты
        default_number = datetime.now().strftime('%Y%m%d01')
        
        invoice_number = Prompt.ask(
            "📋 Введите номер инвойса",
            default=default_number
        )
        
        return invoice_number

    def get_container_number(self):
        """Запрос номера контейнера (опционально)"""
        has_container = Confirm.ask(
            "📦 Есть номер контейнера?",
            default=False
        )
        
        if has_container:
            container = Prompt.ask("Введите номер контейнера")
            return container
        else:
            return ""
    
    def get_output_format(self):
        """Запрос формата вывода"""
        console.print("\n[cyan]📄 Выберите формат вывода:[/cyan]")
        console.print("  [yellow]1[/yellow] - Три отдельных файла (Invoice.xlsx, Specification.xlsx, PackingList.xlsx)")
        console.print("  [yellow]2[/yellow] - Один файл с тремя листами (All_Documents.xlsx)")
        
        choice = Prompt.ask(
            "\nВыберите вариант",
            choices=["1", "2"],
            default="1"
        )
        
        return choice
    
    def generate_combined_file(
            self,
            lines,
            metadata,
            output_path,
            base_name
    ):
        """Генерирует один файл с тремя листами"""
        from openpyxl import Workbook, load_workbook
        from copy import copy
        
        # Генерируем три временных файла
        temp_invoice = output_path / f"temp_invoice.xlsx"
        temp_spec = output_path / f"temp_spec.xlsx"
        temp_packing = output_path / f"temp_packing.xlsx"
        
        invoice_gen = InvoiceGenerator(self.config, self.preset)
        invoice_gen.generate(lines, metadata, str(temp_invoice))
        
        spec_gen = SpecificationGenerator(self.config, self.preset)
        spec_gen.generate(lines, metadata, str(temp_spec))
        
        packing_gen = PackingListGenerator(self.config, self.preset)
        packing_gen.generate(lines, metadata, str(temp_packing))
        
        # Создаем комбинированный файл
        combined_file = output_path / f"{base_name}_All.xlsx"
        wb_combined = Workbook()
        wb_combined.remove(wb_combined.active)  # Удаляем пустой лист по умолчанию
        
        # Копируем листы из временных файлов
        documents = [
            (temp_invoice, "Invoice"),
            (temp_spec, "Specification"),
            (temp_packing, "Packing List")
        ]
        
        for temp_file, sheet_name in documents:
            wb_temp = load_workbook(str(temp_file))
            source_ws = wb_temp.active
            
            # Создаем новый лист в комбинированном файле
            target_ws = wb_combined.create_sheet(title=sheet_name)
            
            # Копируем данные и форматирование
            for row in source_ws.iter_rows():
                for cell in row:
                    target_cell = target_ws[cell.coordinate]
                    target_cell.value = cell.value
                    if cell.has_style:
                        target_cell.font = copy(cell.font)
                        target_cell.border = copy(cell.border)
                        target_cell.fill = copy(cell.fill)
                        target_cell.number_format = copy(cell.number_format)
                        target_cell.protection = copy(cell.protection)
                        target_cell.alignment = copy(cell.alignment)
            
            # Копируем объединенные ячейки
            for merged_cell_range in source_ws.merged_cells.ranges:
                target_ws.merge_cells(str(merged_cell_range))
            
            # Копируем размеры колонок
            for col_letter, col_dim in source_ws.column_dimensions.items():
                target_ws.column_dimensions[col_letter].width = col_dim.width
            
            # Копируем высоты строк
            for row_num, row_dim in source_ws.row_dimensions.items():
                if row_dim.height:
                    target_ws.row_dimensions[row_num].height = row_dim.height
            
            wb_temp.close()
            temp_file.unlink()  # Удаляем временный файл
        
        # Сохраняем комбинированный файл
        wb_combined.save(str(combined_file))
        wb_combined.close()
        
        return combined_file

    def process(
            self,
            input_file: Path,
            invoice_number: str,
            container_number: str = "",
            output_format: str = "1",
            output_dir: str = "output"
    ):
        """Главный метод обработки"""

        console.print(f"\n[bold blue]🚀 Начало обработки Invoice #{invoice_number}[/bold blue]\n")

        try:
            # 1. Парсинг входного файла
            console.print("[cyan]📥 Чтение входного файла...[/cyan]")
            parser = InputFileParser(self.config)
            products = parser.parse(str(input_file))

            if not products:
                console.print("[bold red]❌ Не найдено ни одной позиции в файле[/bold red]")
                return

            console.print(f"[green]✓ Загружено {len(products)} позиций[/green]\n")

            # 2. Обработка данных
            console.print("[cyan]⚙️  Обработка данных...[/cyan]")
            processor = DataProcessor(self.config)
            output_lines = processor.process(products)

            console.print(f"[green]✓ Создано {len(output_lines)} строк для документов[/green]\n")

            # 3. Подготовка метаданных
            date = datetime.now().strftime('%d.%m.%Y')
            metadata = DocumentMetadata(
                invoice_number=invoice_number,
                date=date,
                container_number=container_number,
                seller_name=self.preset['seller']['name'],
                seller_name_en=self.preset['seller'].get('name_en', ''),
                seller_address=self.preset['seller']['address'],
                seller_address_en=self.preset['seller'].get('address_en', ''),
                buyer_name=self.preset['buyer']['name'],
                buyer_address=self.preset['buyer']['address'],
                buyer_address_en=self.preset['buyer'].get('address_en', ''),
                contract_number=self.preset['contract']['number'],
                contract_date=self.preset['contract']['date'],
                terms_of_delivery=self.preset['delivery']['terms'],
                currency=self.preset['delivery']['currency']
            )

            # 4. Создание директории для выходных файлов
            output_path = Path(output_dir)
            output_path.mkdir(exist_ok=True)

            # Имена файлов
            base_name = f"Shusteri_{invoice_number}_{datetime.now().strftime('%Y%m%d')}"

            # 5. Генерация документов с форматированием
            console.print("[cyan]📄 Генерация документов с форматированием...[/cyan]")

            if output_format == "1":
                # Три отдельных файла
                invoice_file = output_path / f"{base_name}_Invoice.xlsx"
                spec_file = output_path / f"{base_name}_Specification.xlsx"
                packing_file = output_path / f"{base_name}_PackingList.xlsx"

                invoice_gen = InvoiceGenerator(self.config, self.preset)
                invoice_gen.generate(output_lines, metadata, str(invoice_file))

                spec_gen = SpecificationGenerator(self.config, self.preset)
                spec_gen.generate(output_lines, metadata, str(spec_file))

                packing_gen = PackingListGenerator(self.config, self.preset)
                packing_gen.generate(output_lines, metadata, str(packing_file))
                
                generated_files = [
                    ("Invoice", invoice_file.name),
                    ("Specification", spec_file.name),
                    ("Packing List", packing_file.name)
                ]
            else:
                # Один файл с тремя листами
                combined_file = self.generate_combined_file(
                    output_lines,
                    metadata,
                    output_path,
                    base_name
                )
                
                generated_files = [
                    ("All Documents (3 sheets)", combined_file.name)
                ]

            console.print("[green]✓ Документы сгенерированы с форматированием[/green]\n")

            # Вывод результатов
            console.print(f"\n[bold green]✅ Обработка завершена успешно![/bold green]\n")

            result_table = Table(title="Сгенерированные файлы")
            result_table.add_column("Документ", style="cyan")
            result_table.add_column("Файл", style="green")

            for doc_name, file_name in generated_files:
                result_table.add_row(doc_name, file_name)

            console.print(result_table)

            # Статистика
            total_qty = sum(line.quantity for line in output_lines)
            total_amount = sum(line.amount for line in output_lines)
            total_net = sum(line.net_weight for line in output_lines)
            total_gross = sum(line.gross_weight for line in output_lines)

            stats_table = Table(title="Статистика")
            stats_table.add_column("Параметр", style="cyan")
            stats_table.add_column("Значение", style="magenta")

            stats_table.add_row("Всего пар", str(total_qty))
            stats_table.add_row("Сумма (CNY)", f"¥{total_amount:,.2f}")
            stats_table.add_row("Вес нетто (кг)", f"{total_net:,.3f}")
            stats_table.add_row("Вес брутто (кг)", f"{total_gross:,.3f}")
            stats_table.add_row("Позиций", str(len(products)))
            stats_table.add_row("Строк в документах", str(len(output_lines)))

            console.print(stats_table)

            console.print(f"\n[bold cyan]📁 Файлы сохранены в: {output_path.absolute()}[/bold cyan]\n")
            
            # Предложение создать еще один документ
            another = Confirm.ask("\n💡 Создать еще один документ?", default=False)
            return another

        except Exception as e:
            logger.error(f"Критическая ошибка: {e}", exc_info=True)
            console.print(f"\n[bold red]💥 Ошибка: {e}[/bold red]\n")
            raise


def main():
    """
    Интерактивный режим работы
    """
    console.print("\n[bold blue]═══════════════════════════════════════════════[/bold blue]")
    console.print("[bold blue]   Shusteri Automation - Генерация документов   [/bold blue]")
    console.print("[bold blue]═══════════════════════════════════════════════[/bold blue]\n")
    
    try:
        automation = ShusteriAutomation()
        
        # Выбор пресета клиента
        preset = automation.select_preset()
        if not preset:
            console.print("[bold red]Не удалось выбрать пресет. Завершение работы.[/bold red]")
            return
        
        automation.preset = preset
        
        # Цикл для создания нескольких документов
        while True:
            # 1. Выбор входного файла
            input_file = automation.select_input_file()
            if not input_file:
                break
            
            # 2. Запрос номера инвойса
            invoice_number = automation.get_invoice_number()
            
            # 3. Запрос номера контейнера
            container_number = automation.get_container_number()
            
            # 4. Запрос формата вывода
            output_format = automation.get_output_format()
            
            # 5. Обработка
            create_another = automation.process(
                input_file,
                invoice_number,
                container_number,
                output_format
            )
            
            # Если пользователь не хочет создавать еще документы - выходим
            if not create_another:
                break
        
        console.print("\n[bold green]👋 До свидания![/bold green]\n")

    except KeyboardInterrupt:
        console.print("\n\n[yellow]⚠️  Прервано пользователем[/yellow]\n")
    except Exception as e:
        console.print(f"\n[bold red]💥 Критическая ошибка: {e}[/bold red]\n")
        exit(1)


if __name__ == "__main__":
    main()
```

---

### 📄 Файл: `src/models.py`

```python
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
```

---

### 📄 Файл: `src/parser.py`

```python
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
```

---

### 📄 Файл: `src/processor.py`

```python
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
```

---

### 📄 Файл: `src/generators/__init__.py`

```python
"""
Генераторы документов (ИСПРАВЛЕННАЯ ВЕРСИЯ БЕЗ ГРУППИРОВКИ)
"""
from .invoice import InvoiceGenerator
from .specification import SpecificationGenerator
from .packing_list import PackingListGenerator

__all__ = ['InvoiceGenerator', 'SpecificationGenerator', 'PackingListGenerator']
```

---

### 📄 Файл: `src/generators/invoice.py`

```python
"""
Генератор Invoice БЕЗ группировки
Выводит все строки из output_lines подряд
"""
import pandas as pd
from typing import List
from ..models import OutputLine, DocumentMetadata


class InvoiceGenerator:
    """Генератор Инвойса"""

    def __init__(self, config: dict, preset: dict):
        self.config = config
        self.preset = preset

    def generate(
            self,
            lines: List[OutputLine],
            metadata: DocumentMetadata,
            output_path: str
    ) -> str:
        """
        Генерирует Invoice по шаблону
        
        Returns:
            путь к сохраненному файлу
        """

        rows = []

        # Шапка документа
        seller_name_full = f"{metadata.seller_name} ({metadata.seller_name_en})" if metadata.seller_name_en else metadata.seller_name
        rows.append([seller_name_full])
        rows.append([metadata.seller_address])
        rows.append([metadata.seller_address_en if metadata.seller_address_en else ''])
        rows.append([f"ТЕЛ/TEL:  {self.preset['seller']['phone']}"])
        rows.append([''])  # Пустая строка
        rows.append([f"COMMERCIAL INVOICE / КОММЕРЧЕСКИЙ ИНВОЙС № {metadata.invoice_number} from/от {metadata.date}"])
        rows.append([''])  # Пустая строка
        rows.append([f"Buyer / Покупатель: {metadata.buyer_name}\n{metadata.buyer_address}"])
        rows.append([f"Contract / Контракт №{metadata.contract_number} from/от {metadata.contract_date}"])
        
        container_text = f"Terms of delivery / Условия поставки: {metadata.terms_of_delivery}"
        container_row = [container_text, '', '', '', '', '', '', '', '', f"Container No / Контейнер №", metadata.container_number]
        rows.append(container_row)
        
        rows.append([''])  # Пустая строка

        # Заголовок таблицы
        header = [
            "№",
            "Brand / Марка",
            "Code / Код ТНВЭД",
            "Factory code / Артикул",
            " Description / Описание",
            "Qty of pairs\nКол-во пар",
            "Net weight\nВес нетто, кг",
            "Gross weight\nВес брутто, кг",
            "Qty of places\nКол-во мест",
            "Price / Цена, cny",
            "Amount / \nСумма, cny"
        ]
        rows.append(header)

        # КЛЮЧЕВОЕ ИЗМЕНЕНИЕ: БЕЗ ГРУППИРОВКИ - просто все строки подряд!
        total_qty = 0
        total_net = 0
        total_gross = 0
        total_boxes = 0
        total_amount = 0
        
        data_start_row = 13  # Строка, где начинаются данные (после заголовка)
        
        # Отслеживаем номер позиции и пары строк для объединения boxes
        item_number = 0
        prev_line = None

        for idx, line in enumerate(lines):
            # Определяем является ли строка "второй частью" артикула
            is_continuation = (
                prev_line is not None and
                line.article == prev_line.article and
                "более 24см" in line.insole_category and
                line.boxes == 0 and
                "до 24см" in prev_line.insole_category
            )
            
            # Увеличиваем номер только для первых строк
            if not is_continuation:
                item_number += 1
            
            rows.append([
                '' if is_continuation else item_number,  # Номер только для первой строки
                '' if is_continuation else line.brand,
                line.hs_code,
                '' if is_continuation else line.article,
                f"{line.description}, материал верха: {line.material}, {line.insole_category}",
                line.quantity,
                float(line.net_weight),
                float(line.gross_weight),
                '',  # Boxes - будет заполнено после объединения ячеек
                float(line.price),
                float(line.amount)
            ])

            total_qty += line.quantity
            total_net += line.net_weight
            total_gross += line.gross_weight
            # ИСПРАВЛЕНО: Правильный подсчет boxes
            # Для второй строки пары (is_continuation) НЕ добавляем boxes,
            # так как они уже были добавлены в первой строке
            if not is_continuation:
                total_boxes += line.boxes
            total_amount += line.amount
            
            prev_line = line

        data_end_row = len(rows)  # Последняя строка с данными

        # Итоговая строка
        rows.append([
            '',
            '',
            '',
            '',
            'Total / Итого:',
            total_qty,
            round(float(total_net), 3),
            round(float(total_gross), 3),
            total_boxes,
            '',
            f"¥{total_amount:,.2f}".replace(',', ' ')
        ])

        # Футер (только для Invoice - БЕЗ Terms of delivery и БЕЗ подписей!)
        rows.append([f"-Manufacturer / Производитель: {metadata.seller_name}"])
        rows.append([f"-Country of origin / Страна происхождения: {self.preset['delivery']['country_of_origin_en']} / {self.preset['delivery']['country_of_origin']}"])
        rows.append(["-Country of destination / Страна назначения: Russia / Россия"])
        rows.append(["-Product not for military use / Товар не для применения в военных целях"])
        rows.append(["-Terms of payment / Условия оплаты:"])
        rows.append([f"Payment of the cost of this transaction for the delivery of the goods specified above in the framework of the execution of Contract No. {metadata.contract_number} dated {metadata.contract_date} in the amount of ¥ {total_amount:,.2f} is payable no later than 120 days from the date of filing the Declaration for the goods in the country of Import./ Оплата стоимости данной сделки по поставке товара, указанного выше в рамках исполнения Контракта № {metadata.contract_number} от {metadata.contract_date} г. в размере ¥ {total_amount:,.2f} подлежит оплате не позднее 120 дней с даты подачи Декларации на товар в стране Импорта"])

        # Создаем DataFrame и сохраняем без форматирования
        df = pd.DataFrame(rows)
        # ВАЖНО: Используем контекстный менеджер чтобы файл корректно закрылся
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, header=False)
        
        # Объединяем ячейки boxes для пар строк
        from openpyxl import load_workbook
        wb = load_workbook(output_path)
        ws = wb.active
        
        # Проходим по строкам данных и объединяем boxes для пар
        prev_line = None
        current_row = data_start_row
        
        for line in lines:
            is_continuation = (
                prev_line is not None and
                line.article == prev_line.article and
                "более 24см" in line.insole_category and
                line.boxes == 0 and
                "до 24см" in prev_line.insole_category
            )
            
            if is_continuation:
                # Объединяем ячейки boxes (колонка I) для предыдущей и текущей строки
                ws.merge_cells(f'I{current_row-1}:I{current_row}')
                # Записываем ИСХОДНОЕ значение boxes (сумма обеих строк)
                ws[f'I{current_row-1}'] = line.original_boxes  # ← ИСПРАВЛЕНО: используем original_boxes
            elif line.boxes > 0:
                # Для обычных строк просто записываем значение
                ws[f'I{current_row}'] = line.boxes
            
            current_row += 1
            prev_line = line
        
        wb.save(output_path)
        wb.close()
        
        # Применяем форматирование
        from ..formatters import InvoiceFormatter
        formatter = InvoiceFormatter(self.config)
        formatter.format(output_path, data_start_row, data_end_row)
        
        return output_path
```

---

### 📄 Файл: `src/generators/specification.py`

```python
"""
Генератор Specification БЕЗ группировки
Выводит все строки из output_lines подряд
"""
import pandas as pd
from typing import List
from ..models import OutputLine, DocumentMetadata


class SpecificationGenerator:
    """Генератор Спецификации"""

    def __init__(self, config: dict, preset: dict):
        self.config = config
        self.preset = preset

    def generate(
            self,
            lines: List[OutputLine],
            metadata: DocumentMetadata,
            output_path: str
    ) -> str:
        """Генерирует Specification"""

        rows = []

        # Шапка
        rows.append(['', '', '', '', '', '', '', '', '', '', '', '', 'Спецификация к Контракту/', '', '', '', '', ''])
        rows.append(['', '', '', '', '', '', '', '', '', '', '', '', 'Specification to the Contract', '', '', '', '', ''])
        rows.append(['', '', '', '', '', '', '', '', '', '', '', '', f'№{metadata.contract_number} from/от {metadata.contract_date}', '', '', '', '', ''])
        rows.append(['', '', '', '', '', '', '', '', '', '', '', '', f'Container / Контейнер № {metadata.container_number}', '', '', '', '', ''])
        rows.append([''])  # Пустая строка 5
        rows.append([''])  # Пустая строка 6
        rows.append([''])  # Пустая строка 7
        rows.append([f'Specification / Спецификация № {metadata.invoice_number} from/от {metadata.date}'])

        rows.append([''])  # Пустая строка 9

        # Заголовок таблицы
        header = [
            "№",
            "Brand / Марка",
            "Code / Код ТНВЭД",
            "Factory code / Артикул",
            " Description / Описание",
            "Color / Цвет",
            "Top material / Материал верха",
            "Lining material / Материал подкладки",
            "Outsole material / Материал подошвы",
            "Heel height / Высота каблука",
            "Insole length / Длина стельки",
            "Quantity of pairs / Количество пар",
            "Net weight (kg) / Вес нетто (кг)",
            "Gross weight (kg) / Вес брутто (кг)",
            "Quantity of places / Количество мест",
            "Price / Цена, cny",
            "Amount / Сумма, cny",
            "КИЗ"
        ]
        rows.append(header)

        # КЛЮЧЕВОЕ ИЗМЕНЕНИЕ: БЕЗ ГРУППИРОВКИ - просто все строки подряд!
        total_qty = 0
        total_net = 0
        total_gross = 0
        total_boxes = 0
        total_amount = 0
        
        data_start_row = 11  # Строка, где начинаются данные
        
        # Отслеживаем номер позиции и пары строк
        item_number = 0
        prev_line = None

        for idx, line in enumerate(lines):
            # Определяем является ли строка "второй частью" артикула
            is_continuation = (
                prev_line is not None and
                line.article == prev_line.article and
                "более 24см" in line.insole_category and
                line.boxes == 0 and
                "до 24см" in prev_line.insole_category
            )
            
            # Увеличиваем номер только для первых строк
            if not is_continuation:
                item_number += 1

            rows.append([
                '' if is_continuation else item_number,
                '' if is_continuation else line.brand,
                line.hs_code,
                '' if is_continuation else line.article,
                '' if is_continuation else line.description,
                '' if is_continuation else line.color,
                '' if is_continuation else line.material,
                '' if is_continuation else line.lining,
                '' if is_continuation else line.sole,
                '' if is_continuation else line.heel_height,
                line.insole_category,
                line.quantity,
                float(line.net_weight),
                float(line.gross_weight),
                '',  # Boxes - будет заполнено после объединения
                float(line.price),
                float(line.amount),
                ' '.join(line.kiz_codes) if line.kiz_codes else ''
            ])

            total_qty += line.quantity
            total_net += line.net_weight
            total_gross += line.gross_weight
            # ИСПРАВЛЕНО: Правильный подсчет boxes
            # Для второй строки пары (is_continuation) НЕ добавляем boxes,
            # так как они уже были добавлены в первой строке
            if not is_continuation:
                total_boxes += line.boxes
            total_amount += line.amount
            
            prev_line = line

        data_end_row = len(rows)  # Последняя строка с данными

        # Итоги
        rows.append([
            '',
            '',
            '',
            '',
            'Total / Итого:',
            '',
            '',
            '',
            '',
            '',
            '',
            total_qty,
            round(float(total_net), 3),
            round(float(total_gross), 3),
            total_boxes,
            '',
            float(total_amount),
            ''
        ])

        # Футер с подписями (для Specification - С Terms of delivery и С подписями!)
        rows.append([f"-Manufacturer / Производитель: {metadata.seller_name}"])
        rows.append([f"-Country of origin / Страна происхождения: {self.preset['delivery']['country_of_origin_en']} / {self.preset['delivery']['country_of_origin']}"])
        rows.append(["-Country of destination / Страна назначения: Russia / Россия"])
        rows.append(["-Product not for military use / Товар не для применения в военных целях"])
        rows.append([f"-Terms of delivery / Условия поставки: {metadata.terms_of_delivery}"])
        rows.append(["-Terms of payment / Условия оплаты:"])
        rows.append([f"Payment of the cost of this transaction for the delivery of the goods specified above in the framework of the execution of Contract No. {metadata.contract_number} dated {metadata.contract_date} in the amount of ¥ {total_amount:,.2f} is payable no later than 120 days from the date of filing the Declaration for the goods in the country of Import./ Оплата стоимости данной сделки по поставке товара, указанного выше в рамках исполнения Контракта № {metadata.contract_number} от {metadata.contract_date} г. в размере ¥ {total_amount:,.2f} подлежит оплате не позднее 120 дней с даты подачи Декларации на товар в стране Импорта"])

        # Подписи (только в Specification!)
        rows.append(['', 'Buyer / Покупатель:', '', '', '', '', '', '', '', '', 'Seller / Продавец:', '', '', '', '', '', '', ''])
        rows.append(['', metadata.buyer_name.split('/')[0].strip(), '', '', '', '', '', '', '', '', metadata.seller_name, '', '', '', '', '', '', ''])

        # Создаем DataFrame и сохраняем
        df = pd.DataFrame(rows)
        # ВАЖНО: Используем контекстный менеджер чтобы файл корректно закрылся
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, header=False)
        
        # Объединяем ячейки boxes для пар строк
        from openpyxl import load_workbook
        wb = load_workbook(output_path)
        ws = wb.active
        
        # Проходим по строкам данных и объединяем boxes для пар
        prev_line = None
        current_row = data_start_row
        
        for line in lines:
            is_continuation = (
                prev_line is not None and
                line.article == prev_line.article and
                "более 24см" in line.insole_category and
                line.boxes == 0 and
                "до 24см" in prev_line.insole_category
            )
            
            if is_continuation:
                # Объединяем ячейки boxes (колонка O) для предыдущей и текущей строки
                ws.merge_cells(f'O{current_row-1}:O{current_row}')
                # Записываем ИСХОДНОЕ значение boxes (сумма обеих строк)
                ws[f'O{current_row-1}'] = line.original_boxes  # ← ИСПРАВЛЕНО: используем original_boxes
            elif line.boxes > 0:
                # Для обычных строк просто записываем значение
                ws[f'O{current_row}'] = line.boxes
            
            current_row += 1
            prev_line = line
        
        wb.save(output_path)
        wb.close()
        
        # Применяем форматирование
        from ..formatters import SpecificationFormatter
        formatter = SpecificationFormatter(self.config)
        formatter.format(output_path, data_start_row, data_end_row)
        
        return output_path
```

---

### 📄 Файл: `src/generators/packing_list.py`

```python
"""
Генератор Packing List БЕЗ группировки
Выводит все строки из output_lines подряд
"""
import pandas as pd
from typing import List
from ..models import OutputLine, DocumentMetadata


class PackingListGenerator:
    """Генератор Упаковочного листа"""

    def __init__(self, config: dict, preset: dict):
        self.config = config
        self.preset = preset

    def generate(
            self,
            lines: List[OutputLine],
            metadata: DocumentMetadata,
            output_path: str
    ) -> str:
        """Генерирует Packing List"""

        rows = []

        # Шапка
        rows.append([f"{metadata.seller_name} ({self.preset['seller']['name_en']})"])
        rows.append([metadata.seller_address])
        rows.append([metadata.seller_address_en])
        rows.append([f"ТЕЛ/TEL:  {self.preset['seller']['phone']}"])
        rows.append([''])  # Пустая строка
        rows.append([f"Packing list / Упаковочный лист № {metadata.invoice_number} from/от {metadata.date}"])
        rows.append([''])  # Пустая строка
        rows.append([f"Buyer / Покупатель: {metadata.buyer_name}"])
        rows.append([metadata.buyer_address])
        rows.append([f"Contract / Контракт №{metadata.contract_number} from/от {metadata.contract_date}"])
        
        container_text = f"Terms of delivery / Условия поставки: {metadata.terms_of_delivery}"
        container_row = [container_text, '', '', '', '', '', '', '', '', f"Container No / Контейнер №", metadata.container_number]
        rows.append(container_row)
        
        rows.append([''])  # Пустая строка

        # Заголовок
        header = [
            "№",
            "Brand / Марка",
            "Code / Код ТНВЭД",
            "Factory code / Артикул",
            " Description / Описание",
            "Color / Цвет",
            "Quantity of pairs / Количество пар",
            "Net weight, kg / Вес нетто, кг",
            "Gross weight, kg / Вес брутто, кг",
            "Quantity of places / Количество мест",
            "Type of packaging / Вид упаковки"
        ]
        rows.append(header)

        # КЛЮЧЕВОЕ ИЗМЕНЕНИЕ: БЕЗ ГРУППИРОВКИ - просто все строки подряд!
        total_qty = 0
        total_net = 0
        total_gross = 0
        total_boxes = 0
        
        data_start_row = 14  # Строка, где начинаются данные
        
        # Отслеживаем номер позиции и пары строк
        item_number = 0
        prev_line = None

        for idx, line in enumerate(lines):
            # Определяем является ли строка "второй частью" артикула
            is_continuation = (
                prev_line is not None and
                line.article == prev_line.article and
                "более 24см" in line.insole_category and
                line.boxes == 0 and
                "до 24см" in prev_line.insole_category
            )
            
            # Увеличиваем номер только для первых строк
            if not is_continuation:
                item_number += 1

            rows.append([
                '' if is_continuation else item_number,
                '' if is_continuation else line.brand,
                line.hs_code,
                '' if is_continuation else line.article,
                '' if is_continuation else line.description,
                '' if is_continuation else line.color,
                line.quantity,
                float(line.net_weight),
                float(line.gross_weight),
                '',  # Boxes - будет заполнено после объединения
                "cardboard box / \nкартонная коробка" if line.boxes > 0 or is_continuation else ''
            ])

            total_qty += line.quantity
            total_net += line.net_weight
            total_gross += line.gross_weight
            # ИСПРАВЛЕНО: Правильный подсчет boxes
            # Для второй строки пары (is_continuation) НЕ добавляем boxes,
            # так как они уже были добавлены в первой строке
            if not is_continuation:
                total_boxes += line.boxes
            
            prev_line = line

        data_end_row = len(rows)  # Последняя строка с данными

        # Итоги
        rows.append(['', '', '', '', '', 'Итого:', total_qty, round(float(total_net), 3), round(float(total_gross), 3), total_boxes, ''])
        rows.append([f"Total net weight, kgs / Общий вес нетто: {total_net:,.2f} кг"])
        rows.append([f"Total gross weight, kg / Общий вес брутто: {total_gross:,.2f} кг"])
        rows.append([f"Total quantity of pairs / Общее кол-во пар: {total_qty}"])
        rows.append([f"Total quantity of places / Общее кол-во мест: {total_boxes}"])

        # Создаем DataFrame и сохраняем
        df = pd.DataFrame(rows)
        # ВАЖНО: Используем контекстный менеджер чтобы файл корректно закрылся
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, header=False)
        
        # Объединяем ячейки boxes для пар строк
        from openpyxl import load_workbook
        wb = load_workbook(output_path)
        ws = wb.active
        
        # Проходим по строкам данных и объединяем boxes для пар
        prev_line = None
        current_row = data_start_row
        
        for line in lines:
            is_continuation = (
                prev_line is not None and
                line.article == prev_line.article and
                "более 24см" in line.insole_category and
                line.boxes == 0 and
                "до 24см" in prev_line.insole_category
            )
            
            if is_continuation:
                # Объединяем ячейки boxes (колонка J) для предыдущей и текущей строки
                ws.merge_cells(f'J{current_row-1}:J{current_row}')
                # Записываем ИСХОДНОЕ значение boxes (сумма обеих строк)
                ws[f'J{current_row-1}'] = line.original_boxes  # ← ИСПРАВЛЕНО: используем original_boxes
            elif line.boxes > 0:
                # Для обычных строк просто записываем значение
                ws[f'J{current_row}'] = line.boxes
            
            current_row += 1
            prev_line = line
        
        wb.save(output_path)
        wb.close()
        
        # Применяем форматирование
        from ..formatters import PackingListFormatter
        formatter = PackingListFormatter(self.config)
        formatter.format(output_path, data_start_row, data_end_row)
        
        return output_path
```

---

### 📄 Файл: `src/formatters/__init__.py`

```python
"""
Модуль форматирования документов
"""
from .invoice_formatter import InvoiceFormatter
from .specification_formatter import SpecificationFormatter
from .packing_list_formatter import PackingListFormatter

__all__ = ['InvoiceFormatter', 'SpecificationFormatter', 'PackingListFormatter']
```

---

### 📄 Файл: `src/formatters/invoice_formatter.py`

```python
"""
Форматтер для Invoice документа
"""
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter


class InvoiceFormatter:
    """Применяет форматирование к Invoice"""
    
    def __init__(self, config: dict):
        self.config = config
    
    def format(self, file_path: str, data_start_row: int, data_end_row: int):
        """
        Применяет форматирование к файлу
        
        Args:
            file_path: путь к файлу Excel
            data_start_row: номер строки, где начинаются данные таблицы (после заголовка)
            data_end_row: номер последней строки с данными (перед Total)
        """
        wb = load_workbook(file_path)
        ws = wb.active
        
        # Настройка ширины колонок
        self._set_column_widths(ws)
        
        # Форматирование шапки документа
        self._format_header(ws)
        
        # Форматирование таблицы
        self._format_table(ws, data_start_row, data_end_row)
        
        # Форматирование футера
        self._format_footer(ws, data_end_row)
        
        # Сохранение
        wb.save(file_path)
        wb.close()  # ВАЖНО: Закрываем файл чтобы Windows мог его удалить
    
    def _set_column_widths(self, ws):
        """Устанавливает ширину колонок"""
        widths = {
            'A': 4.66,
            'B': 11.0,
            'C': 11.83,
            'D': 14.66,
            'E': 91.66,
            'F': 11.66,
            'G': 12.33,
            'H': 13.66,
            'I': 13.33,
            'J': 9.66,
            'K': 14.83
        }
        
        for col, width in widths.items():
            ws.column_dimensions[col].width = width
    
    def _format_header(self, ws):
        """Форматирует шапку документа"""
        # Строка 1: Название продавца
        ws.merge_cells('A1:K1')
        cell = ws['A1']
        cell.font = Font(name='Arial', size=16, bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='top', wrap_text=True)
        ws.row_dimensions[1].height = 22.5
        
        # Строка 2: Адрес (строка 1)
        ws.merge_cells('A2:K2')
        cell = ws['A2']
        cell.font = Font(name='Arial', size=11)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.row_dimensions[2].height = 15.6
        
        # Строка 3: Адрес (строка 2)
        ws.merge_cells('A3:K3')
        cell = ws['A3']
        cell.font = Font(name='Arial', size=11)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.row_dimensions[3].height = 15.6
        
        # Строка 4: Телефон
        ws.merge_cells('A4:K4')
        cell = ws['A4']
        cell.font = Font(name='Arial', size=11)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[4].height = 14.25
        
        # Строка 5: пустая
        ws.row_dimensions[5].height = 6.0
        
        # Строка 6: Заголовок документа COMMERCIAL INVOICE
        ws.merge_cells('A6:K6')
        cell = ws['A6']
        cell.font = Font(name='Arial', size=16, bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='top', wrap_text=True)
        ws.row_dimensions[6].height = 21.0
        
        # Строка 7: пустая
        ws.row_dimensions[7].height = 12.75
        
        # Строка 8: Buyer
        ws.merge_cells('A8:K8')
        cell = ws['A8']
        cell.font = Font(name='Arial', size=11)
        cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
        ws.row_dimensions[8].height = 49.5
        
        # Строка 9: Contract
        ws.merge_cells('A9:K9')
        cell = ws['A9']
        cell.font = Font(name='Arial', size=10)
        cell.alignment = Alignment(horizontal='left', vertical='center')
        ws.row_dimensions[9].height = 14.25
        
        # Строка 10: Terms of delivery + Container
        ws.merge_cells('A10:I10')
        cell = ws['A10']
        cell.font = Font(name='Arial', size=10)
        cell.alignment = Alignment(horizontal='left', vertical='center')
        ws.row_dimensions[10].height = 14.25
        
        # Строка 11: пустая
        ws.row_dimensions[11].height = 6.0
    
    def _format_table(self, ws, data_start_row, data_end_row):
        """Форматирует таблицу с данными"""
        # Заголовок таблицы (строка 12)
        header_row = 12
        ws.row_dimensions[header_row].height = 33.75
        
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Форматирование заголовков
        for col in range(1, 12):  # A-K
            cell = ws.cell(row=header_row, column=col)
            cell.font = Font(name='Arial', size=10, bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = thin_border
        
        # Форматирование строк данных
        for row in range(data_start_row, data_end_row + 1):
            for col in range(1, 12):
                cell = ws.cell(row=row, column=col)
                cell.border = thin_border
                
                # Выравнивание для разных колонок
                if col == 1:  # №
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.font = Font(name='Arial', size=10)
                elif col in [2, 3, 4]:  # Brand, Code, Article
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    cell.font = Font(name='Arial', size=10)
                elif col == 5:  # Description
                    cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                    cell.font = Font(name='Arial', size=10)
                else:  # Числовые колонки
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                    cell.font = Font(name='Arial', size=10)
    
    def _format_footer(self, ws, data_end_row):
        """Форматирует футер документа"""
        total_row = data_end_row + 1
        
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Строка Total
        for col in range(1, 12):
            cell = ws.cell(row=total_row, column=col)
            cell.border = thin_border
            cell.font = Font(name='Arial', size=10, bold=True)
            
            if col == 5:  # "Total / Итого:"
                cell.alignment = Alignment(horizontal='right', vertical='center')
            elif col in [6, 7, 8, 9]:  # Числовые поля
                cell.alignment = Alignment(horizontal='right', vertical='center')
            elif col == 11:  # Amount
                cell.alignment = Alignment(horizontal='right', vertical='center')
            else:
                cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # Информационные строки после таблицы
        info_start = total_row + 1
        
        for row_offset in range(0, 7):  # 7 строк информации
            row = info_start + row_offset
            ws.merge_cells(f'A{row}:K{row}')
            cell = ws[f'A{row}']
            cell.font = Font(name='Arial', size=10)
            cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
```

---

### 📄 Файл: `src/formatters/specification_formatter.py`

```python
"""
Форматтер для Specification документа
"""
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


class SpecificationFormatter:
    """Применяет форматирование к Specification"""
    
    def __init__(self, config: dict):
        self.config = config
    
    def format(self, file_path: str, data_start_row: int, data_end_row: int):
        """
        Применяет форматирование к файлу
        
        Args:
            file_path: путь к файлу Excel
            data_start_row: номер строки, где начинаются данные таблицы
            data_end_row: номер последней строки с данными
        """
        wb = load_workbook(file_path)
        ws = wb.active
        
        # Настройка ширины колонок
        self._set_column_widths(ws)
        
        # Форматирование шапки документа
        self._format_header(ws)
        
        # Форматирование таблицы
        self._format_table(ws, data_start_row, data_end_row)
        
        # Форматирование футера
        self._format_footer(ws, data_end_row)
        
        # Сохранение
        wb.save(file_path)
        wb.close()  # ВАЖНО: Закрываем файл чтобы Windows мог его удалить
    
    def _set_column_widths(self, ws):
        """Устанавливает ширину колонок"""
        widths = {
            'A': 3.33,   # №
            'B': 13.33,  # Brand
            'C': 12.0,   # Code
            'D': 18.5,   # Factory code
            'E': 21.0,   # Description
            'F': 14.0,   # Color
            'G': 22.0,   # Top material
            'H': 14.33,  # Lining material
            'I': 15.5,   # Outsole material
            'J': 12.16,  # Heel height
            'K': 15.33,  # Insole length
            'L': 12.16,  # Quantity
            'M': 15.0,   # Net weight
            'N': 18.83,  # Gross weight
            'O': 13.16,  # Quantity of places
            'P': 11.0,   # Price
            'Q': 15.0,   # Amount
            'R': 25.0    # KIZ codes
        }
        
        for col, width in widths.items():
            ws.column_dimensions[col].width = width
    
    def _format_header(self, ws):
        """Форматирует шапку документа"""
        # Строки 1-3: Спецификация к Контракту (объединенные)
        for row in [1, 2, 3]:
            ws.merge_cells(f'A{row}:R{row}')
            cell = ws[f'A{row}']
            cell.font = Font(name='Arial', size=14, bold=True)
            cell.alignment = Alignment(horizontal='right', vertical='center')
        
        # Строка 4: Container
        ws.merge_cells('A4:R4')
        cell = ws['A4']
        cell.font = Font(name='Arial', size=12)
        cell.alignment = Alignment(horizontal='right', vertical='center')
        
        # Строка 5: пустая
        
        # Строка 6: пустая
        
        # Строка 7: пустая
        
        # Строка 8: Specification № ... from ...
        ws.merge_cells('A8:R8')
        cell = ws['A8']
        cell.font = Font(name='Arial', size=14, bold=True)
        cell.alignment = Alignment(horizontal='left', vertical='center')
        
        # Строка 9: пустая
    
    def _format_table(self, ws, data_start_row, data_end_row):
        """Форматирует таблицу с данными"""
        # Заголовок таблицы (строка 10)
        header_row = 10
        ws.row_dimensions[header_row].height = 60
        
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Форматирование заголовков (18 колонок A-R)
        for col in range(1, 19):
            cell = ws.cell(row=header_row, column=col)
            cell.font = Font(name='Arial', size=9, bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = thin_border
        
        # Форматирование строк данных
        for row in range(data_start_row, data_end_row + 1):
            for col in range(1, 19):
                cell = ws.cell(row=row, column=col)
                cell.border = thin_border
                cell.font = Font(name='Arial', size=9)
                
                # Выравнивание для разных колонок
                if col == 1:  # №
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                elif col in [2, 3, 4]:  # Brand, Code, Article
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                elif col in [5, 6, 7, 8, 9, 10, 11]:  # Текстовые поля
                    cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                elif col == 18:  # KIZ codes
                    cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
                else:  # Числовые колонки
                    cell.alignment = Alignment(horizontal='right', vertical='center')
    
    def _format_footer(self, ws, data_end_row):
        """Форматирует футер документа"""
        total_row = data_end_row + 1
        
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Строка Total
        for col in range(1, 19):
            cell = ws.cell(row=total_row, column=col)
            cell.border = thin_border
            cell.font = Font(name='Arial', size=10, bold=True)
            cell.alignment = Alignment(horizontal='right', vertical='center')
        
        # Информационные строки после таблицы
        info_start = total_row + 1
        
        for row_offset in range(0, 9):  # 9 строк информации (включая подписи)
            row = info_start + row_offset
            ws.merge_cells(f'A{row}:R{row}')
            cell = ws[f'A{row}']
            cell.font = Font(name='Arial', size=10)
            cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
        
        # Последние 2 строки для подписей (Buyer/Seller)
        signature_row = info_start + 9
        ws.merge_cells(f'A{signature_row}:I{signature_row}')
        ws.merge_cells(f'J{signature_row}:R{signature_row}')
        
        cell = ws[f'A{signature_row}']
        cell.font = Font(name='Arial', size=10)
        cell.alignment = Alignment(horizontal='left', vertical='center')
        
        cell = ws[f'J{signature_row}']
        cell.font = Font(name='Arial', size=10)
        cell.alignment = Alignment(horizontal='left', vertical='center')
        
        # Строка с названиями компаний
        company_row = signature_row + 1
        ws.merge_cells(f'A{company_row}:I{company_row}')
        ws.merge_cells(f'J{company_row}:R{company_row}')
        
        cell = ws[f'A{company_row}']
        cell.font = Font(name='Arial', size=10)
        cell.alignment = Alignment(horizontal='left', vertical='center')
        
        cell = ws[f'J{company_row}']
        cell.font = Font(name='Arial', size=10)
        cell.alignment = Alignment(horizontal='left', vertical='center')
```

---

### 📄 Файл: `src/formatters/packing_list_formatter.py`

```python
"""
Форматтер для Packing List документа
"""
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side


class PackingListFormatter:
    """Применяет форматирование к Packing List"""
    
    def __init__(self, config: dict):
        self.config = config
    
    def format(self, file_path: str, data_start_row: int, data_end_row: int):
        """
        Применяет форматирование к файлу
        
        Args:
            file_path: путь к файлу Excel
            data_start_row: номер строки, где начинаются данные таблицы
            data_end_row: номер последней строки с данными
        """
        wb = load_workbook(file_path)
        ws = wb.active
        
        # Настройка ширины колонок
        self._set_column_widths(ws)
        
        # Форматирование шапки документа
        self._format_header(ws)
        
        # Форматирование таблицы
        self._format_table(ws, data_start_row, data_end_row)
        
        # Форматирование футера
        self._format_footer(ws, data_end_row)
        
        # Сохранение
        wb.save(file_path)
        wb.close()  # ВАЖНО: Закрываем файл чтобы Windows мог его удалить
    
    def _set_column_widths(self, ws):
        """Устанавливает ширину колонок"""
        widths = {
            'A': 3.33,   # №
            'B': 13.5,   # Brand
            'C': 13.33,  # Code
            'D': 16.0,   # Factory code
            'E': 27.0,   # Description
            'F': 21.0,   # Color
            'G': 17.16,  # Quantity
            'H': 17.5,   # Net weight
            'I': 19.66,  # Gross weight
            'J': 18.66,  # Quantity of places
            'K': 21.16   # Type of packaging
        }
        
        for col, width in widths.items():
            ws.column_dimensions[col].width = width
    
    def _format_header(self, ws):
        """Форматирует шапку документа"""
        # Строка 1: Название продавца
        ws.merge_cells('A1:K1')
        cell = ws['A1']
        cell.font = Font(name='Arial', size=16, bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='top', wrap_text=True)
        ws.row_dimensions[1].height = 22.5
        
        # Строка 2: Адрес (строка 1)
        ws.merge_cells('A2:K2')
        cell = ws['A2']
        cell.font = Font(name='Arial', size=11)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.row_dimensions[2].height = 15.6
        
        # Строка 3: Адрес (строка 2)
        ws.merge_cells('A3:K3')
        cell = ws['A3']
        cell.font = Font(name='Arial', size=11)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.row_dimensions[3].height = 15.6
        
        # Строка 4: Телефон
        ws.merge_cells('A4:K4')
        cell = ws['A4']
        cell.font = Font(name='Arial', size=11)
        cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[4].height = 14.25
        
        # Строка 5: пустая
        ws.row_dimensions[5].height = 6.0
        
        # Строка 6: Заголовок документа Packing list
        ws.merge_cells('A6:K6')
        cell = ws['A6']
        cell.font = Font(name='Arial', size=16, bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='top', wrap_text=True)
        ws.row_dimensions[6].height = 21.0
        
        # Строка 7: пустая
        ws.row_dimensions[7].height = 12.75
        
        # Строка 8: Buyer
        ws.merge_cells('A8:K8')
        cell = ws['A8']
        cell.font = Font(name='Arial', size=11)
        cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
        ws.row_dimensions[8].height = 14.25
        
        # Строка 9: Адрес покупателя
        ws.merge_cells('A9:K9')
        cell = ws['A9']
        cell.font = Font(name='Arial', size=11)
        cell.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
        ws.row_dimensions[9].height = 49.5
        
        # Строка 10: Contract
        ws.merge_cells('A10:K10')
        cell = ws['A10']
        cell.font = Font(name='Arial', size=10)
        cell.alignment = Alignment(horizontal='left', vertical='center')
        ws.row_dimensions[10].height = 14.25
        
        # Строка 11: Terms of delivery + Container
        ws.merge_cells('A11:I11')
        cell = ws['A11']
        cell.font = Font(name='Arial', size=10)
        cell.alignment = Alignment(horizontal='left', vertical='center')
        ws.row_dimensions[11].height = 14.25
        
        # Строка 12: пустая
        ws.row_dimensions[12].height = 6.0
    
    def _format_table(self, ws, data_start_row, data_end_row):
        """Форматирует таблицу с данными"""
        # Заголовок таблицы (строка 13)
        header_row = 13
        ws.row_dimensions[header_row].height = 50
        
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Форматирование заголовков
        for col in range(1, 12):  # A-K
            cell = ws.cell(row=header_row, column=col)
            cell.font = Font(name='Arial', size=10, bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = thin_border
        
        # Форматирование строк данных
        for row in range(data_start_row, data_end_row + 1):
            for col in range(1, 12):
                cell = ws.cell(row=row, column=col)
                cell.border = thin_border
                cell.font = Font(name='Arial', size=10)
                
                # Выравнивание для разных колонок
                if col == 1:  # №
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                elif col in [2, 3, 4]:  # Brand, Code, Article
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                elif col in [5, 6]:  # Description, Color
                    cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                elif col == 11:  # Type of packaging
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                else:  # Числовые колонки
                    cell.alignment = Alignment(horizontal='right', vertical='center')
    
    def _format_footer(self, ws, data_end_row):
        """Форматирует футер документа"""
        total_row = data_end_row + 1
        
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Строка Total (Итого:)
        for col in range(1, 12):
            cell = ws.cell(row=total_row, column=col)
            cell.border = thin_border
            cell.font = Font(name='Arial', size=10, bold=True)
            
            if col == 6:  # "Итого:"
                cell.alignment = Alignment(horizontal='right', vertical='center')
            elif col in [7, 8, 9, 10]:  # Числовые поля
                cell.alignment = Alignment(horizontal='right', vertical='center')
            else:
                cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # Информационные строки после таблицы (4 строки)
        info_start = total_row + 1
        
        for row_offset in range(0, 4):
            row = info_start + row_offset
            ws.merge_cells(f'A{row}:K{row}')
            cell = ws[f'A{row}']
            cell.font = Font(name='Arial', size=11, bold=True)
            cell.alignment = Alignment(horizontal='left', vertical='center')
```

---

## 📌 КОНЕЦ ДОКУМЕНТАЦИИ

_Полная документация проекта Shusteri Automation_  
_Версия: 1.0 Complete_  
_Дата: 2025-01-16_  
_Автор: Rostislav_
