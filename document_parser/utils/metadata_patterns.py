"""
Паттерны и функции для извлечения метаданных из строительной документации.
Используется всеми парсерами (Excel, Image, PDF).
"""

import re
import pandas as pd
from typing import Dict, List, Any, Optional

from ..core.models import TableMetadata

# ============================================================
# ПАТТЕРНЫ ДЛЯ ПОИСКА МЕТАДАННЫХ
# ============================================================

METADATA_PATTERNS: Dict[str, List[str]] = {
    'document_name': [
        r'ведомост[ьи]?\s*№\s*[а-я\d/\-\.]+',
        r'ведомост[ьи]?[\sа-я]*',
        r'отч[её]т[уа]?\s*№\s*[а-я\d/\-\.]+',
        r'отч[её]т[уа]?[\sа-я]*',
        r'[\sа-я]*таблица[\sа-я]*',
        r'протокол\s*№\s*[а-я\d/\-\.]+',
        r'(?:ПРОТОКОЛ|Протокол)[\sа-я]*',
        r'график[\sа-я]*',
        r'нормы[\s\dа-я]*',
        r'журнал[\s\dа-я]*',
    ],

    'document_number': [
        r'приложение\s*[:—–-№]?\s*[\d/\-\.а-я]+',
        r'акт[у]?\s*[:—–-№]?\s*[\d/\-\.а-я]+',
    ],

    'object_name': [
        r'объект[уа]?\s*[:—–-]?\s*(.*)',
    ],

    'contract_number': [
        r'контракт[у]?\s*№\s*([а-я\d/\-\.]+)',
        r'договор[у]?\s*№\s*([а-я\d/\-\.]+)',
    ],

    'compilation_date': [
        # Дата со словом "дата", "от", и т.д.
        r'дат[аы]\s*(?:составления)?\s*[:—–-]?\s*(\d{1,2}[\./-]\d{1,2}[\./-]\d{2,4})',
        r'дат[аы]\s*(?:составления)?\s*[:—–-]?\s*(\d{2,4}[\./-]\d{1,2}[\./-]\d{1,2})',
        r'от\s*(\d{1,2}[\./-]\d{1,2}[\./-]\d{2,4})',
        r'от\s*(\d{2,4}[\./-]\d{1,2}[\./-]\d{1,2})',
        r'за\s*[:—–-]?\s*(\d{1,2}[\./-]\d{1,2}[\./-]\d{2,4})',
        r'за\s*[:—–-]?\s*(\d{2,4}[\./-]\d{1,2}[\./-]\d{1,2})',
        # Дата в формате "месяц год"
        r'(январ[ья]|феврал[ья]|март[а]?|апрел[ья]|ма[йя]|июн[ья]|июл[ья]|август[а]?|сентябр[ья]|октябр[ья]|ноябр[ья]|декабр[ья])\s*(\d{4})\s*г\.?',
        # Просто дата дд.мм.гггг
        r'(\d{1,2}[\./-]\d{1,2}[\./-]\d{2,4})',
        r'(\d{2,4}[\./-]\d{1,2}[\./-]\d{1,2})',
    ],

    'customer': [  # Заказчик
        r'заказчик\s*[:—–-]?\s+(.*[а-я]\.\s*[а-я]\.\s*[а-я]+)',
        r'заказчик\s*[:—–-]?\s+(.*[а-я]+\s*[а-я]\.\s*[а-я]\.?)',
    ],

    'contractor': [  # Подрядчик
        r'подрядчик\s*[:—–-]?\s+(.*[а-я]\.\s*[а-я]\.\s*[а-я]+)',
        r'подрядчик\s*[:—–-]?\s+(.*[а-я]+\s*[а-я]\.\s*[а-я]\.?)',
    ],

    'compiled': [  # Составил
        r'составил\s*[:—–-]?\s+(.*[а-я]\.\s*[а-я]\.\s*[а-я]+)',
        r'составил\s*[:—–-]?\s+(.*[а-я]+\s*[а-я]\.\s*[а-я]\.?)',
    ],

    'checked': [  # Проверил
        r'проверил\s*[:—–-]?\s+(.*[а-я]\.\s*[а-я]\.\s*[а-я]+)',
        r'проверил\s*[:—–-]?\s+(.*[а-я]+\s*[а-я]\.\s*[а-я]\.?)',
    ],

    'note': [  # Примечание
        r'примечание\s*[:—–-]?\s*(.+)',
    ],

    'extra': [
        r'^(заказчик|подрядчик|составил|проверил).*[а-я]\.\s*[а-я]\.\s*[а-я]+',
    ],
}

# Поля, которые нужно очищать от подчёркиваний
FIELDS_TO_CLEAN_UNDERSCORES = {'customer', 'contractor', 'compiled', 'checked'}

# Поля, для которых нужно брать первую группу захвата
FIELDS_WITH_CAPTURE_GROUP = {
    'document_number', 'contract_number', 'object_name',
    'compilation_date', 'note', 'customer', 'contractor',
    'compiled', 'checked'
}


# ============================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ============================================================

def _search_in_row(
        df: pd.DataFrame,
        row_idx: int,
        patterns: Optional[Dict[str, List[str]]],
        metadata: Dict
) -> None:
    """
    Ищет паттерны в конкретной строке DataFrame.
    """
    row_values = [str(x) for x in df.iloc[row_idx].values if pd.notna(x)]
    row_str = ' '.join(row_values)

    for key, pattern_list in patterns.items():
        if key in metadata:
            continue

        for pattern in pattern_list:
            match = re.search(pattern, row_str, re.IGNORECASE | re.DOTALL)
            if match:
                if key in ('document_number', 'contract_number', 'object_name', 'compilation_date', 'note'):
                    # Пытаемся взять группу, если есть
                    if match.groups():
                        value = match.group(1).strip()
                    else:
                        value = match.group().strip()
                    metadata[key] = value
                elif key == 'document_name':
                    text = match.group().strip()
                    if text == text.upper() and row_idx + 1 < len(df):
                        col_idx = _find_match_column(df, row_idx, pattern)
                        if col_idx is not None:
                            next_cell_value = df.iloc[row_idx + 1, col_idx]
                            if pd.notna(next_cell_value) and str(next_cell_value).strip():
                                next_text = str(next_cell_value).strip()
                                metadata[key] = text + " " + next_text
                                continue
                    metadata[key] = text
                elif key in FIELDS_TO_CLEAN_UNDERSCORES:
                    value = match.group(1).strip()
                    value = re.sub(r'\s*_{3,}\s*', ' ', value)
                    metadata[key] = value
                break


def _find_match_column(df: pd.DataFrame, row_idx: int, pattern: str) -> Optional[int]:
    """
    Находит индекс столбца, в котором паттерн совпал.
    """
    for col_idx in range(len(df.columns)):
        cell_value = df.iloc[row_idx, col_idx]
        if pd.notna(cell_value):
            cell_str = str(cell_value)
            if re.search(pattern, cell_str, re.IGNORECASE):
                return col_idx
    return None

# ============================================================
# ПУБЛИЧНЫЕ ФУНКЦИИ ДЛЯ ИЗВЛЕЧЕНИЯ МЕТАДАННЫХ
# ============================================================

def extract_metadata_from_text(
        text: str,
        patterns: Optional[Dict[str, List[str]]] = None
) -> TableMetadata | None:
    """
    Извлекает метаданные из произвольного текста.

    Args:
        text: Текст для поиска метаданных.
        patterns: Словарь с паттернами (если None, используется METADATA_PATTERNS).

    Returns:
        Словарь с найденными метаданными.
    """
    if patterns is None:
        patterns = METADATA_PATTERNS

    metadata = TableMetadata()

    lines = text.split('\n')

    for line in lines:
        line = line.strip()
        if not line:
            continue

        for key, pattern_list in patterns.items():
            if getattr(metadata, key):
                continue

            for pattern in pattern_list:
                match = re.search(pattern, line, re.IGNORECASE)
                if match:
                    if match.groups():
                        value = match.group(1).strip()
                    else:
                        value = match.group().strip()
                    setattr(metadata, key, value)
                    break

    return metadata


def extract_metadata_from_df(
        df: pd.DataFrame,
        table_start_idx: int,
        table_end_idx: int,
        patterns: Optional[Dict[str, List[str]]] = None
) -> TableMetadata | None:
    """
    Извлекает метаданные из DataFrame (до и после таблицы).

    Args:
        df: DataFrame с данными.
        table_start_idx: Индекс строки, с которой начинается таблица.
        table_end_idx: Индекс строки, которой заканчивается таблица.
        patterns: Словарь с паттернами (если None, используется METADATA_PATTERNS).

    Returns:
        Словарь с найденными метаданными.
    """
    if patterns is None:
        patterns = METADATA_PATTERNS

    metadata = {}

    for i in range(table_start_idx):
        _search_in_row(df, i, patterns, metadata)

    for i in range(table_end_idx + 1, len(df)):
        _search_in_row(df, i, patterns, metadata)

    return TableMetadata(**metadata) if metadata else None