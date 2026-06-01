"""
Вспомогательные функции для использования библиотеки document_parser.

Предоставляет готовые инструменты для:
- Выбора парсера по расширению файла
- Парсинга одного файла
- Экспорта результата в JSON
- Вывода информации о результате в консоль
"""

import json
import os
from pathlib import Path
from datetime import datetime
from typing import Union, Optional

from .parsers.excel_parser import ExcelParser
from .parsers.image_parser import ImageParser
from .parsers.word_parser import WordParser
from .core.models import ParseResult


# Карта расширений файлов
SUPPORTED_EXTENSIONS = {
    '.xls': 'excel',
    '.xlsx': 'excel',
    '.xlsm': 'excel',
    '.jpg': 'image',
    '.jpeg': 'image',
    '.png': 'image',
    '.tiff': 'image',
    '.tif': 'image',
    '.bmp': 'image',
    '.docx': 'word',
    '.doc': 'word',
}


def get_parser_for_file(
        filepath: Union[str, Path],
        tesseract_path: str | None = None,
        lang: str = 'rus+eng',
):
    """
    Возвращает подходящий парсер в зависимости от расширения файла.

    Args:
        filepath: Путь к файлу (строка или Path).
        tesseract_path: Путь к Tesseract (для Windows, для парсинга изображений).
        lang: Язык распознавания (для парсинга изображений).

    Returns:
        Экземпляр парсера (ExcelParser, ImageParser или WordParser).

    Raises:
        ValueError: Если расширение файла не поддерживается.

    Examples:
        >>> parser = get_parser_for_file("document.xlsx")
        >>> result = parser.parse("document.xlsx")
    """
    filepath = Path(filepath) if isinstance(filepath, str) else filepath
    ext = filepath.suffix.lower()

    parser_map = {
        'excel': ExcelParser,
        'image': ImageParser,
        'word': WordParser,
    }

    parser_type = SUPPORTED_EXTENSIONS.get(ext)

    if parser_type is None:
        raise ValueError(
            f"Неподдерживаемое расширение файла: {ext}. "
            f"Поддерживаются: {', '.join(sorted(SUPPORTED_EXTENSIONS.keys()))}"
        )

    if parser_type == 'image':
        usage_tesseract_path = tesseract_path
        if not usage_tesseract_path:
            usage_tesseract_path = os.path.join("C:", "Program Files", "Tesseract-OCR", "tesseract.exe")
        return ImageParser(tesseract_cmd=usage_tesseract_path, lang=lang)

    return parser_map[parser_type]()


def parse_file(filepath: Union[str, Path]) -> ParseResult:
    """
    Парсит один файл и возвращает результат.

    Args:
        filepath: Путь к файлу.

    Returns:
        ParseResult с результатами парсинга.

    Raises:
        FileNotFoundError: Если файл не найден.
        ValueError: Если расширение не поддерживается.

    Examples:
        >>> result = parse_file("ведомость.xls")
        >>> print(result.filename)
        >>> print(len(result.tables))
    """
    filepath = Path(filepath)

    if not filepath.exists():
        raise FileNotFoundError(f"Файл не найден: {filepath}")

    parser = get_parser_for_file(filepath)

    with open(filepath, "rb") as f:
        return parser.parse(f, filepath.name)


def print_result(result: ParseResult):
    """
    Выводит информацию о результате парсинга в консоль.

    Args:
        result: Результат парсинга (ParseResult).

    Examples:
        >>> result = parse_file("document.xlsx")
        >>> print_result(result)
    """
    if result.success:
        print(f"Файл: {result.filename}")
        print(f"Найдено таблиц: {len(result.tables)}")

        for i, table in enumerate(result.tables):
            print(f"\n--- Таблица {i + 1}: {table.sheet_name} ---")
            print(f"Заголовки: {table.headers}")
            print(f"Строк: {len(table.rows)}")

            if table.metadata:
                meta = table.metadata
                print(f"Метаданные: {table.metadata}")

            if table.rows:
                max_rows = min(2, len(table.rows))
                print(f"\n  Первые {max_rows} строк данных:")
                for j, row in enumerate(table.rows[:max_rows]):
                    print(f"  Строка {j + 1}:")
                    for key, value in row.items():
                        if value is not None:
                            str_value = str(value)
                            if len(str_value) > 80:
                                str_value = str_value[:80] + "..."
                            print(f"    {key}: {str_value}")
    else:
        print(f"Ошибка при парсинге: {result.error}")


def result_to_json(result: ParseResult) -> str:
    """
    Конвертирует результат парсинга в JSON-строку.

    Args:
        result: Результат парсинга.

    Returns:
        JSON-строка.

    Examples:
        >>> result = parse_file("document.xlsx")
        >>> json_str = result_to_json(result)
        >>> print(json_str)
    """
    return json.dumps(
        result.to_dict(),
        ensure_ascii=False,
        indent=2,
        default=_json_serializer
    )


def export_to_json(result: ParseResult, output_path: Union[str, Path]):
    """
    Сохраняет результат парсинга в JSON-файл.

    Args:
        result: Результат парсинга.
        output_path: Путь для сохранения JSON.

    Examples:
        >>> result = parse_file("document.xlsx")
        >>> export_to_json(result, "output.json")
    """
    output_path = Path(output_path)
    json_str = result_to_json(result)

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(json_str)

    print(f"Результат сохранён в: {output_path}")


def _json_serializer(obj):
    """Сериализатор для типов, не поддерживаемых json по умолчанию."""
    if isinstance(obj, datetime):
        return obj.isoformat()
    if hasattr(obj, 'isoformat'):
        return obj.isoformat()
    raise TypeError(f"Тип {type(obj)} не сериализуется")


def is_supported_file(filepath: Union[str, Path]) -> bool:
    """
    Проверяет, поддерживается ли файл для парсинга.

    Args:
        filepath: Путь к файлу.

    Returns:
        True, если файл поддерживается.

    Examples:
        >>> is_supported_file("document.xlsx")
        True
        >>> is_supported_file("document.pdf")
        False
    """
    ext = Path(filepath).suffix.lower()
    return ext in SUPPORTED_EXTENSIONS