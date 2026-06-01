"""
Document Parser

Инструмент для преобразования инженерной документации строительной отрасли
в структурированный формат JSON для автоматизированного анализа.
"""
from document_parser.parsers.word_parser import WordParser
from document_parser.parsers.excel_parser import ExcelParser
from document_parser.parsers.image_parser import ImageParser
from document_parser.core.models import (
    ParseResult,
    ParsedTable,
    TableMetadata,
)
from document_parser.usage_helpers import (
    get_parser_for_file,
    parse_file,
    print_result,
    result_to_json,
    export_to_json,
    is_supported_file,
)

__version__ = "0.1.0"
__author__ = "Elena Shangina"
__email__ = "arlantamom1999@gmail.com"

__all__ = [
    "ExcelParser",
    "ImageParser",
    "WordParser",

    "ParseResult",
    "ParsedTable",
    "TableMetadata",

    "get_parser_for_file",
    "parse_file",
    "print_result",
    "result_to_json",
    "export_to_json",
    "is_supported_file",
]