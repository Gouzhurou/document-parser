"""
Вспомогательный класс для представления ячейки таблицы на изображении.
"""

from dataclasses import dataclass, field
from typing import Optional, Dict, Any, List


@dataclass
class Cell:
    """
    Представляет одну ячейку таблицы, найденную на изображении.
    """
    x: int
    y: int
    w: int
    h: int

    # Вычисляемые поля
    cx: int = field(init=False)  # центр X
    cy: int = field(init=False)  # центр Y
    area: int = field(init=False)  # площадь

    # Данные после OCR
    text: Optional[str] = None

    def __post_init__(self):
        """Вычисляет производные поля после инициализации."""
        self.cx = self.x + self.w // 2
        self.cy = self.y + self.h // 2
        self.area = self.w * self.h

    @property
    def x_end(self) -> int:
        """Правая граница ячейки."""
        return self.x + self.w

    @property
    def y_end(self) -> int:
        """Нижняя граница ячейки."""
        return self.y + self.h

    @property
    def bbox(self) -> tuple:
        """Кортеж (x1, y1, x2, y2) для совместимости с OpenCV."""
        return self.x, self.y, self.x_end, self.y_end

    @property
    def center(self) -> tuple:
        """Центр ячейки (cx, cy)."""
        return self.cx, self.cy

    @property
    def left_up(self) -> tuple:
        """Левый верхний угол (x, y)."""
        return self.x, self.y

    @property
    def right_down(self) -> tuple:
        """Правый нижний угол."""
        return self.x_end, self.y_end

    def is_inside(
            self,
            outer: Cell,
            tolerance: int,
    ) -> bool:
        """
        Проверяет, находится ли ячейка внутри outer с учетом погрешности.

        Args:
            outer: Внешняя ячейка.
            tolerance: Допуск, по которому ячейки будут считаться соприкасающимися.

        Returns:
            True, если ячейка находится внутри outer.
        """
        ox1, oy1, ox2, oy2 = outer.bbox

        ix1, iy1, ix2, iy2 = self.bbox

        is_inside = (
                ix1 >= ox1 - tolerance and
                iy1 >= oy1 - tolerance and
                ix2 <= ox2 + tolerance and
                iy2 <= oy2 + tolerance
        )

        return is_inside

    def is_pressed_to_borders(
            self,
            cell: Cell,
            tolerance: int,
    ) -> bool:
        """
        Проверяет, прижат ли cell к границам ячейки с учетом погрешности.

        Args:
            cell: Ячейка.
            tolerance: Допуск, по которому ячейки будут считаться соприкасающимися.

        Returns:
            True, если cell прижат к границам ячейки.
        """
        x1_1, y1_1, x1_2, y1_2 = cell.bbox
        x2_1, y2_1, x2_2, y2_2 = self.bbox

        # Пересечение границ ячеек
        x_intersection = x2_1 < x1_2 + tolerance and x2_2 > x1_1 - tolerance
        y_intersection = y2_1 < y1_2 + tolerance and y2_2 > y1_1 - tolerance

        # Чтобы ячейка стояла рядом не достаточно проверки на 1 координату
        near_left = abs(x2_1 - x1_1) <= tolerance and y_intersection
        near_right = abs(x2_2 - x1_2) <= tolerance and y_intersection
        near_top = abs(y2_1 - y1_1) <= tolerance and x_intersection
        near_bottom = abs(y2_2 - y1_2) <= tolerance and x_intersection

        is_near_border = near_left or near_right or near_top or near_bottom

        return is_near_border

    @classmethod
    def is_empty(cls, row: List[Cell]) -> bool:
        """
        Проверяет, является ли строка полностью пустой

        Args:
            row: Список с ячейками.
        """
        return all(cell.text == '' for cell in row)

    @classmethod
    def delete_empty_rows(cls, rows: List[List[Cell]]) -> List[List[Cell]]:
        result = []
        for row in rows:
            if not Cell.is_empty(row):
                result.append(row)
        return result

    @classmethod
    def get_max_row_height(cls, row: List[Cell]) -> int:
        """
        Находит максимальную высоту ячейки в строке.
        """
        max_height = 0
        for cell in row:
            if cell.h > max_height:
                max_height = cell.h
        return max_height