"""
Парсер Excel-файлов (.xls, .xlsx, .xlsm).
"""

import traceback
import pandas as pd
from typing import BinaryIO, Union, List, Optional
from pathlib import Path

from ..core.base_parser import BaseParser
from ..core.models import ParseResult, ParsedTable
from ..utils.helpers import (
    find_header_rows,
    build_multi_level_headers,
    remove_rows_between_headers_and_data,
    find_table_end,
    convert_strings_to_dicts,
)
from ..utils.metadata_patterns import (
    extract_metadata_from_df,
)


class ExcelParser(BaseParser):
    """
    Парсер для Excel-файлов строительной документации.

    Умеет:
    - Находить заголовки таблиц в произвольном месте листа.
    - Обрабатывать объединённые ячейки.
    - Извлекать метаданные (дату составления, название документа и т.д.).
    """

    def __init__(
            self,
            header_keywords: Optional[List[str]] = None,
            **kwargs
    ):
        """
        Инициализация парсера Excel.

        Args:
            header_keywords: Список ключевых слов для поиска строки заголовков.
            **kwargs: Дополнительные параметры.
        """
        super().__init__(**kwargs)
        self.header_keywords = header_keywords

    def parse(
            self,
            file: Union[BinaryIO, Path, str],
            filename: Optional[str] = None
    ) -> ParseResult:
        """
        Разбирает Excel-файл и возвращает структурированный результат.

        Args:
            file: Excel-файл для разбора.
            filename: Имя файла (опционально).

        Returns:
            ParseResult: Объект с результатами разбора.
        """
        filename = self._get_filename(file, filename)

        try:
            if isinstance(file, (str, Path)):
                xl = pd.ExcelFile(file)
            else:
                xl = pd.ExcelFile(file)

            all_tables = []

            for sheet_name in xl.sheet_names:
                # Читаем лист без заголовков
                df_raw = pd.read_excel(xl, sheet_name=sheet_name, header=None)
                if df_raw.empty:
                    raise Exception(f"Страница '{sheet_name}' пуста.")

                # Находим строки с заголовками
                header_start, header_end = find_header_rows(df_raw, self.header_keywords)

                if header_start is None or header_end is None:
                    metadata = extract_metadata_from_df(df_raw, -1, -1)
                    headers = []
                    rows = []
                else:
                    # Объединяем строки с заголовками в одну
                    headers = build_multi_level_headers(df_raw, (header_start, header_end))
                    # Удаляем промежуточные строки
                    df_raw = remove_rows_between_headers_and_data(df_raw, header_end + 1)

                    # Находим последнюю строку таблицы
                    table_end = find_table_end(df_raw, header_end + 1)

                    # Извлекаем метаданные из сырого DataFrame
                    metadata = extract_metadata_from_df(df_raw, header_start, table_end)

                    # Извлекаем данные таблицы
                    data_df = df_raw.iloc[header_end+1:table_end+1].copy()
                    if len(data_df.columns) > len(headers):
                        data_df = data_df.iloc[:, :len(headers)]

                    # Конвертируем строки в словари
                    rows = convert_strings_to_dicts(data_df, headers)

                all_tables.append(ParsedTable(
                    sheet_name=sheet_name,
                    headers=headers,
                    rows=rows,
                    metadata=metadata,
                ))

            return ParseResult(
                filename=filename,
                tables=all_tables,
                success=True,
            )

        except Exception as e:
            traceback.print_exc()
            return ParseResult(
                filename=filename,
                tables=[],
                success=False,
                error=str(e),
            )