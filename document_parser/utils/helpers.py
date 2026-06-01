"""
Вспомогательные функции для обработки данных.
"""

import re
import math
from typing import List, Dict, Any, Optional, Tuple
import pandas as pd
import numpy as np
from pathlib import Path

import pythoncom
import win32com.client


def convert_doc_to_docx_win32(doc_path: str, output_path: str = None) -> Path | None:
    """
    Конвертирует .doc в .docx через Microsoft Word (COM).
    ТОЛЬКО для Windows с установленным Microsoft Word.
    """
    doc_path = Path(doc_path)

    if output_path is None:
        output_path = doc_path.with_suffix('.docx')
    else:
        output_path = Path(output_path)

    try:
        # Инициализация COM для текущего потока
        pythoncom.CoInitialize()

        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False

        try:
            # Открываем с явным указанием формата
            doc = word.Documents.Open(
                str(doc_path.absolute()),
                ConfirmConversions=False,
                ReadOnly=True
            )
            # 16 = формат docx
            doc.SaveAs(str(output_path.absolute()), FileFormat=16)
            doc.Close(SaveChanges=False)

            return output_path
        except Exception as e:
            print(f"Ошибка при открытии/сохранении: {e}")
            return None
        finally:
            word.Quit()
            pythoncom.CoUninitialize()

    except ImportError:
        print("pywin32 не установлен. Установите: pip install pywin32")
        return None
    except Exception as e:
        print(f"Ошибка конвертации через Word: {e}")
        try:
            pythoncom.CoUninitialize()
        except:
            pass
        return None


def _clean_value(val: Any) -> Any:
    """
    Приводит значения к нативным типам Python для JSON сериализации.

    Args:
        val: Любое значение из pandas DataFrame.

    Returns:
        Очищенное значение (str, int, float, bool, None).
    """
    if pd.isna(val) or val is None:
        return None

    if isinstance(val, float) and math.isnan(val):
        return None

    if isinstance(val, (np.integer,)):
        return int(val)

    if isinstance(val, (np.floating,)):
        if math.isnan(val):
            return None
        return float(val)

    if isinstance(val, (pd.Timestamp, np.datetime64)):
        return val.isoformat()

    if isinstance(val, (np.bool_,)):
        return bool(val)

    if isinstance(val, str):
        return normalize_string(val.strip())

    return val


def normalize_string(string: str) -> str:
    """
    Очищает строку от переносов строк и лишних пробелов.

    Args:
        string: Строка.

    Returns:
        Нормализованная строка.
    """
    if not string:
        return ""

    normalized = ' '.join(str(string).replace('\n', ' ').replace('\r', ' ').split())
    return normalized


def align_rows_to_max_columns(rows: List[List[Any]], fill_value: Any = '') -> List[List[Any]]:
    """
    Выравнивает строки по максимальному количеству столбцов.
    Короткие строки дополняются fill_value справа.

    Args:
        rows: Список строк, где каждая строка - список значений.
        fill_value: Значение для заполнения недостающих ячеек.

    Returns:
        Список строк одинаковой длины.

    Examples:
        >>> rows = [['a', 'b'], ['c'], ['d', 'e', 'f']]
        >>> align_rows_to_max_columns(rows)
        [['a', 'b', ''], ['c', '', ''], ['d', 'e', 'f']]
    """
    if not rows:
        return []

    max_cols = max(len(row) for row in rows)

    aligned = []
    for row in rows:
        if len(row) < max_cols:
            row = list(row) + [fill_value] * (max_cols - len(row))
        aligned.append(row)

    return aligned


def convert_strings_to_dicts(
        df: pd.DataFrame,
        headers: List[str],
) -> List[Dict[str, Any]]:
    """
    Конвертирует строки DataFrame в список словарей.

    Args:
        df: DataFrame с данными.
        headers: Список заголовков столбцов.

    Returns:
        Список словарей, где ключи - заголовки, значения - очищенные данные.
    """
    num_cols = len(df.columns)
    num_headers = len(headers)
    if num_cols != num_headers:
        raise Exception(f"Несоответствие количества столбцов: в DataFrame {num_cols}, а в headers {num_headers}")

    df.columns = headers

    rows = []

    for _, row in df.iterrows():
        row_dict = {}
        for col in headers:
            row_dict[col] = _clean_value(row[col])
        rows.append(row_dict)

    return rows


def _is_empty(row: List[str]) -> bool:
    return all(element == '' for element in row)


def clean_empty_rows(matrix: List[List[str]]) -> List[List[str]]:
    clean_matrix = []

    for row in matrix:
        normalize_row = [normalize_string(element) for element in row]
        if not _is_empty(normalize_row):
            clean_matrix.append(row)

    return clean_matrix


def find_header_rows(
        df: pd.DataFrame,
        keywords: Optional[List[str]] = None,
        max_rows: int = 30,
        min_keywords: int = 2
) -> tuple[None, None] | tuple[int, int]:
    """
     Находит диапазон строк, которые являются заголовком таблицы.

    Args:
        df: DataFrame с данными.
        keywords: Список regex-паттернов для поиска.
        max_rows: Максимальное количество строк для поиска.
        min_keywords: Минимальное количество ключевых слов в строке-заголовке.

    Returns:
        Кортеж (start_row, end_row) — индексы первой и последней строки заголовка.
    """
    # TODO: вынести список регулярок как константу

    # ключевые слова для поиска
    if keywords is None:
        keywords = [
            r'№\s*п[\./]?п',  # № п/п, №п/п, № п.п
            r'\bpid\b',  # PID как отдельное слово
            r'\bнаименование\b',  # наименование
            r'ед[\.\s]*изм',  # ед. изм, ед.изм, ед изм
            r'\bединица\b',  # единица
            r'\bкол[\.\-]?во\b',  # кол-во, кол.во, колво
            r'\bколичество\b',  # количество
            r'\bхар\-ка\b',  # хар-ка
            r'\bхарактеристик[аи]\b',  # характеристика, характеристики
            r'\bдлина\b',  # длина
            r'\bширина\b',  # ширина
            r'\bглубина\b',  # глубина
            r'\bтолщина\b',  # толщина
            r'\bрасстояние\b',  # расстояние
            r'\bмаркировка\b',  # маркировка
            r'\bшифр\b',  # шифр
            r'\bгост\b',  # ГОСТ
            r'\bдата\b',  # дата
            r'\bитог[о]?\b',  # итог, итого
            r'\bстоимость\b',  # стоимость
            r'\bпримечание\b',  # примечание
            r'\bначал[оа]\b',  # начало, начала
            r'\bконец\b',  # конец
            r'\период\b',  # период
            r'параметр',  # параметр
        ]

    patterns = [re.compile(kw, re.IGNORECASE) for kw in keywords]

    header_start = None
    header_end = None

    for i in range(min(max_rows, len(df))):
        row_values = [str(x) for x in df.iloc[i].values if pd.notna(x)]
        text = ' '.join(row_values)

        matches = sum(len(p.findall(text)) for p in patterns)

        if matches >= min_keywords:
            if header_start is None:
                header_start = i
            header_end = i
        else:
            if header_start is not None:
                break

    # Проверяет, не пропустил ли он строки из-за того, что в них не было долтаточного числа ключевых слов
    # И решает проблему дублирующих заголовков
    building_headers = build_multi_level_headers(df, (header_start, header_end))
    while has_duplicate_strings(building_headers):
        header_end += 1
        if header_end >= len(df) or header_end - header_start + 1 > len(df.columns):
            return (None, None)
        building_headers = build_multi_level_headers(df, (header_start, header_end))

    return header_start, header_end


def has_duplicate_strings(strings: List[str]) -> bool:
    """
    Проверяет, есть ли в заголовках дублирующиеся названия.

    Args:
        strings: Список заголовков столбцов.

    Returns:
        True, если есть дубликаты, иначе False.
    """
    seen = set()

    for string in strings:
        if string in seen:
            return True

        seen.add(string)

    return False


def _is_numbering_row(row_values: List[str]) -> bool:
    """
    Определяет, является ли строка нумерацией столбцов (0, 1, 2, 3...).

    Args:
        row_values: Список значений в строке.

    Returns:
        True, если строка содержит последовательные числа и больше ничего значимого.
    """
    if not row_values:
        return False

    # Фильтруем непустые значения
    non_empty = [v for v in row_values if v and str(v).strip()]

    if len(non_empty) < 3:
        return False

    # Пытаемся преобразовать в числа
    numbers = []
    for v in non_empty:
        try:
            num = int(str(v).strip())
            numbers.append(num)
        except (ValueError, TypeError):
            return False

    # Проверяем последовательность
    is_sequential = all(
        numbers[i] == numbers[0] + i
        for i in range(len(numbers))
    )

    # Проверяем, что начинается с 0 или 1
    starts_correctly = numbers[0] in (0, 1)

    return is_sequential and starts_correctly


def _is_total_row(row_values: List[str]) -> bool:
    """
    Определяет, является ли строка строкой с итогами.

    Args:
        row_values: Список значений в строке.

    Returns:
        True, если строка содержит "Итого", "Итог", "Всего" и т.д.
    """
    if not row_values:
        return False

    row_str = ' '.join(row_values).lower()

    total_keywords = [
        r'итог[о]?\s*[:—–-]?',
        r'всего',
        r'total',
        r'сумма',
    ]

    for keyword in total_keywords:
        if re.search(keyword, row_str, re.IGNORECASE):
            return True

    return False


def remove_rows_between_headers_and_data(
        df: pd.DataFrame,
        data_start_idx: int,
        check_rows: int = 5,
) -> pd.DataFrame:
    """
    Удаляет информационные строки, которые находятся между заголовками и данными.

    Информационными строками являются:
    - строки с нумерацией столбцов
    - строки с итоговыми значениями

    Args:
        df: DataFrame с данными.
        data_start_idx: Индекс строки, где начинаются данные.

    Returns:
        DataFrame без информационных строк.
    """
    if data_start_idx >= len(df):
        return df

    rows_to_drop = []
    current_idx = data_start_idx

    # Проверяем несколько строк после заголовков
    for i in range(current_idx, min(current_idx + check_rows, len(df))):
        row_values = [
            str(v) for v in df.iloc[i].values
            if pd.notna(v) and str(v).strip()
        ]

        if _is_total_row(row_values) or _is_numbering_row(row_values):
            rows_to_drop.append(i)
        else:
            break

    if rows_to_drop:
        df = df.drop(df.index[rows_to_drop]).reset_index(drop=True)

    return df


def build_multi_level_headers(
        df: pd.DataFrame,
        header_rows: Tuple[int, int],
        columns_count: int | None = None,
) -> List[str]:
    """
    Собирает многоуровневые заголовки в плоский список.

    Args:
        df: DataFrame с данными.
        header_rows: Кортеж (start, end) — строки заголовка.
        columns_count: Количество столбцов, которое должно быть в таблице.

    Returns:
        Список финальных названий столбцов.
    """
    start, end = header_rows
    if start is None or end is None or start > end or start > len(df) or end > len(df):
        return []
    header_df = df.iloc[start:end + 1].copy()

    last_non_empty_col = _find_last_non_empty_column(header_df)
    if last_non_empty_col is not None and columns_count is None:
        header_df = header_df.iloc[:, :last_non_empty_col + 1]

    # Заполняем NaN из объединённых ячеек (ffill по столбцам)
    header_df = header_df.ffill(axis=0).ffill(axis=1)

    # Собираем названия столбцов: объединяем значения сверху вниз
    headers = []
    for col in header_df.columns:
        col_values = [
            normalize_string(str(v))
            for v in header_df[col].values
            if pd.notna(v) and str(v).strip()
        ]

        # Удаляем идущие подряд повторки
        deduped_values = []
        for val in col_values:
            if not deduped_values or deduped_values[-1] != val:
                deduped_values.append(val)

        if deduped_values:
            headers.append('/'.join(deduped_values))
        else:
            headers.append(f'Column_{col}')

    return headers


def _find_last_non_empty_column(df: pd.DataFrame) -> Optional[int]:
    """
    Находит индекс последнего столбца, в котором есть хотя бы одно непустое значение.

    Args:
        df: DataFrame для анализа.

    Returns:
        Индекс последнего непустого столбца или None, если все столбцы пустые.
    """
    last_non_empty = None

    for col_idx, col in enumerate(df.columns):
        # Проверяем, есть ли в столбце хотя бы одно непустое значение
        has_value = any(
            pd.notna(v) and str(v).strip() != ''
            for v in df[col].values
        )

        if has_value:
            last_non_empty = col_idx

    return last_non_empty


def find_table_end(df: pd.DataFrame, data_start: int) -> int:
    """
    Находит конец таблицы — первую полностью пустую строку после данных.

    Args:
        df: DataFrame с данными.
        data_start: Индекс строки, с которой начинаются данные.

    Returns:
        Индекс последней строки с данными (включительно).
    """
    table_end = data_start-1

    for i in range(data_start, len(df)):
        row_values = [x for x in df.iloc[i].values if pd.notna(x) and str(x).strip()]

        if not row_values:
            break

        table_end = i

    return table_end
