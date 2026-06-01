"""
Абстрактный базовый класс для всех парсеров документов.
"""

from abc import ABC, abstractmethod
from typing import BinaryIO, Union, Optional
from pathlib import Path
from .models import ParseResult


class BaseParser(ABC):
    """
    Базовый абстрактный класс для всех парсеров документов.

    Все конкретные реализации парсеров должны наследоваться от этого класса
    и реализовывать метод parse().
    """

    def __init__(self, **kwargs):
        """Инициализация парсера с произвольными параметрами."""
        self.config = kwargs

    @abstractmethod
    def parse(self, file: Union[BinaryIO, Path, str], filename: Optional[str] = None) -> ParseResult:
        """
        Разбирает документ и возвращает структурированный результат.

        Args:
            file: Файл для разбора. Может быть файловым объектом, путём Path или строкой пути.
            filename: Имя файла (опционально, если file — это путь, то извлекается из него).

        Returns:
            ParseResult: Объект с результатами разбора.
        """
        pass

    def _get_filename(self, file: Union[BinaryIO, Path, str], filename: Optional[str] = None) -> str:
        """Вспомогательный метод для получения имени файла."""
        if filename:
            return filename
        if isinstance(file, (str, Path)):
            return Path(file).name
        if hasattr(file, 'name'):
            return Path(file.name).name
        return "unknown"