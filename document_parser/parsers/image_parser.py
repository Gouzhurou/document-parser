"""
Парсер для изображений (JPG, PNG, TIFF) с использованием OpenCV и Tesseract.
Находит таблицы по рамкам ячеек, группирует в строки и столбцы.
"""
import os

import cv2
import numpy as np
import pytesseract
from PIL import Image
from typing import BinaryIO, Union, Optional, List, Dict, Any, Tuple, Sequence
from pathlib import Path

from ..core.base_parser import BaseParser
from ..core.models import ParseResult, ParsedTable
from ..utils.metadata_patterns import (
    extract_metadata_from_text,
)
from ..utils.helpers import normalize_string
from ..utils.cell import Cell


class ImageParser(BaseParser):
    """
    Парсер для изображений с таблицами.

    Алгоритм:
    1. Извлекает метаданные из документа
    2. Находит все прямоугольные контуры (ячейки таблицы)
    3. Отфильтровывает шумовые контуры и контуры оболочки
    4. Группирует ячейки по таблицам
    5. Группирует их по строкам
    6. Распознаёт текст в каждой ячейке
    7. Объединяет многоуровневые строки в одну
    """

    def __init__(
        self,
        tesseract_cmd: Optional[str] = None,
        lang: str = 'rus+eng',
        config: str = '--psm 3 --oem 3',
        row_tolerance_ratio: float = 0.009,
        col_tolerance_ratio: float = 0.012,
        cell_tolerance_ratio: float = 0.003,
        min_cell_area_ratio: float = 0.0009,
        min_cell_x_ratio: float = 0.014,
        min_cell_y_ratio: float = 0.01,
        **kwargs
    ):
        """
        Инициализация парсера.

        Args:
            tesseract_cmd: Путь к Tesseract (для Windows).
            lang: Язык распознавания.
            config: Конфигурация для Tesseract.
            row_tolerance_ratio: Допуск по Y, по которому ячейки будут считаться находящимися в одной строке.
            col_tolerance_ratio: Допуск по X, по которому ячейки будут считаться находящимися в одном столбце.
            cell_tolerance_ratio: Допуск, по которому ячейки будут считаться соседними.
            min_cell_area_ratio: Минимальная площадь ячейки как доля от площади изображения.
            min_cell_x_ratio: Минимальная ширина ячейки как доля от ширины.
            min_cell_y_ratio: Минимальная высота ячейки как доля от высоты.
        """
        super().__init__(**kwargs)

        if tesseract_cmd:
            pytesseract.pytesseract.tesseract_cmd = tesseract_cmd

        self.lang = lang
        self.config = config

        self.row_tolerance_ratio = row_tolerance_ratio
        self.col_tolerance_ratio = col_tolerance_ratio
        self.cell_tolerance_ratio = cell_tolerance_ratio
        self.min_cell_area_ratio = min_cell_area_ratio
        self.min_cell_x_ratio = min_cell_x_ratio
        self.min_cell_y_ratio = min_cell_y_ratio

        # Эти значения будут вычислены при загрузке изображения
        self.row_tolerance = None
        self.col_tolerance = None
        self.cell_tolerance = None
        self.min_cell_area = None
        self.min_cell_x = None
        self.min_cell_y = None

    def _calculate_dynamic_params(self, img_width: int, img_height: int):
        """
        Вычисляет абсолютные значения параметров на основе размера изображения.

        Args:
            img_width: Ширина изображения в пикселях.
            img_height: Высота изображения в пикселях.
        """
        img_area = img_width * img_height

        self.row_tolerance = int(img_height * self.row_tolerance_ratio)
        self.col_tolerance = int(img_width * self.col_tolerance_ratio)
        self.cell_tolerance = int(max(img_width, img_height) * self.cell_tolerance_ratio)
        self.min_cell_area = int(img_area * self.min_cell_area_ratio)
        self.min_cell_x = int(img_width * self.min_cell_x_ratio)
        self.min_cell_y = int(img_height * self.min_cell_y_ratio)

    def parse(
        self,
        file: Union[BinaryIO, Path, str],
        filename: Optional[str] = None
    ) -> ParseResult:
        """
        Разбирает изображение и возвращает структурированный результат.
        """
        filename = self._get_filename(file, filename)

        try:
            if isinstance(file, (str, Path)):
                pil_image = Image.open(file)
            else:
                pil_image = Image.open(file)

            # Конвертируем в OpenCV формат
            img = cv2.cvtColor(np.array(pil_image), cv2.COLOR_RGB2BGR)

            # Вычисляем динамические параметры
            img_height, img_width = img.shape[:2]
            self._calculate_dynamic_params(img_width, img_height)

            # 1. Извлекаем текст со всего изображения для метаданных
            full_text = self._extract_text(img)

            # 2. Извлекаем метаданные
            metadata = extract_metadata_from_text(full_text)

            # 3. Находим и парсим таблицы
            tables = self._extract_tables_from_image(img, debug=True)

            parsed_tables = []
            for table in tables:
                parsed_table = ParsedTable(
                    sheet_name=table['sheet_name'],
                    headers=table['headers'],
                    rows=table['rows'],
                    metadata=metadata
                )
                parsed_tables.append(parsed_table)

            return ParseResult(
                filename=filename,
                tables=parsed_tables,
                success=True
            )

        except Exception as e:
            import traceback
            traceback.print_exc()
            return ParseResult(
                filename=filename,
                tables=[],
                success=False,
                error=str(e)
            )

    def _extract_text(self, img: np.ndarray) -> str:
        """
        Извлекает текст с изображения.
        """
        # Предобработка для лучшего распознавания
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        _, binary = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)

        text = pytesseract.image_to_string(
            binary,
            lang=self.lang,
            config=self.config
        )

        return text

    def _extract_tables_from_image(
            self,
            img: np.ndarray,
            debug: bool = False,
            debug_output_path: str = "debug_cells.jpg",
    ) -> List[Dict[str, Any]]:
        """
        Находит таблицы на изображении и парсит их.
        """
        # Предобработка для поиска линий
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        _, binary = cv2.threshold(gray, 200, 255, cv2.THRESH_BINARY_INV)

        contours, _ = cv2.findContours(binary, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)

        cells = self._filter_contours(img, contours, debug=debug, debug_output_path=debug_output_path)

        groups = self._find_connected_groups(cells)

        # Собираем словари с полями sheet_name, headers, rows
        result = []
        for i, group in enumerate(groups):
            rows = self._group_cells_by_rows(group)
            rows = self._ocr_text(img, rows)
            rows = Cell.delete_empty_rows(rows)

            str_rows = self._group_rows_by_multirows(rows)
            str_rows = self._align_rows(str_rows)

            result_dict = {
                'sheet_name': f"Таблица_{i}",
                'headers': str_rows[0],
                'rows': [],
            }
            for row in str_rows[1:]:
                result_row = {}
                for j, header in enumerate(str_rows[0]):
                    result_row[header] = row[j]
                result_dict['rows'].append(result_row)

            result.append(result_dict)

        return result

    def _align_rows(self, rows: List[List[str]]) -> List[List[str]]:
        # TODO: обработать ситуацию: в строке данных элементов меньше, чем в строке заголовков
        if not rows:
            return []

        count = len(rows[0])
        new_rows = [row[:count] for row in rows]

        return new_rows

    def _filter_contours(
            self,
            img: np.ndarray,
            contours: Sequence[Any],
            debug: bool = False,
            debug_output_path: str = "debug_cells.jpg",
    ) -> List[Cell]:
        # Создаём копию изображения для отладки
        if debug:
            debug_img = img.copy()

        # Фильтруем прямоугольные контуры (шум)
        all_cells = []
        for contour in contours:
            x, y, w, h = cv2.boundingRect(contour)
            area = w * h
            contour_area = cv2.contourArea(contour)

            if (
                    area > self.min_cell_area and
                    w > self.min_cell_x and
                    h > self.min_cell_y and
                    contour_area > 0 and
                    area / contour_area < 3
            ):
                cell = Cell(x=x, y=y, w=w, h=h)
                all_cells.append(cell)
            else:
                if debug:
                    cv2.rectangle(debug_img, (x, y), (x + w, y + h), (255, 0, 0), 1)

        all_cells.sort(key=lambda cell: cell.x)

        # Фильтруем ячейки (внешние рамки)
        without_cover_cells = []
        trash_cells = []
        for cell in all_cells:
            img_copy = img.copy()
            is_cover, inner_cells = self._is_cover(cell, all_cells, img_copy, debug)
            if not is_cover:
                without_cover_cells.append(cell)
                trash_cells.extend(inner_cells)
                if debug:
                    cv2.rectangle(debug_img, cell.left_up, cell.right_down, (0, 255, 0), 2)
                    cv2.putText(
                        debug_img,
                        f"({cell.x},{cell.y})",
                        (cell.x, cell.y - 5),
                        cv2.FONT_HERSHEY_SIMPLEX,
                        0.4,
                        (0, 0, 0),
                        1
                    )
            else:
                if debug:
                    cv2.rectangle(debug_img, cell.left_up, cell.right_down, (0, 0, 255), 1)

        # Удаляем шумовые ячейки
        cells = []
        for cell in without_cover_cells:
            if cell not in trash_cells:
                cells.append(cell)
            else:
                if debug:
                    cv2.rectangle(debug_img, cell.left_up, cell.right_down, (186, 2, 252), 2)

        if debug:
            output_path = str(os.path.join("data", debug_output_path))
            cv2.imwrite(output_path, debug_img)

        return cells

    def _is_cover(
            self,
            cell: Cell,
            all_cells: List[Cell],
            img: np.ndarray,
            debug: bool = False,
            debug_output_path: str = "debug_groups.jpg",
    ) -> Tuple[bool, List[Cell]]:
        """
        Определяет, является ли ячейка внешней оболочкой таблицы.

        Args:
            cell: Проверяемая ячейка.
            all_cells: Список всех ячеек.

        Returns:
            Является ли ячейка внешней оболочкой и список внутренних ячеек.
        """
        # Ищем внутренние ячейки
        inner_cells = []
        for other in all_cells:
            if other is cell:
                continue

            if other.is_inside(cell, self.cell_tolerance):
                inner_cells.append(other)

        # Если внутренних ячеек нет - не оболочка
        if not inner_cells:
            return False, []

        # Если внутри одна ячейка - не оболочка
        if len(inner_cells) == 1:
            return False, inner_cells

        groups = self._find_connected_groups(inner_cells)

        # Отрисовываем найденные группы
        if debug:
            colors = [
                (0, 0, 255),
                (0, 255, 0),
                (255, 0, 0),
            ]
            for i, group in enumerate(groups):
                color = colors[i % len(colors)]
                for cell in group:
                    cv2.rectangle(img, cell.left_up, cell.right_down, color, 2)
                    cv2.putText(
                        img,
                        f"({cell.x},{cell.y})",
                        (cell.x, cell.y - 5),
                        cv2.FONT_HERSHEY_SIMPLEX,
                        0.4,
                        (0, 0, 0),
                        1
                    )
            output_path = str(os.path.join("data", debug_output_path))
            cv2.imwrite(output_path, img)

        structured_groups = [g for g in groups if len(g) >= 2]
        if structured_groups:
            # Есть хотя бы одна структурированная группа - это оболочка
            return True, inner_cells

        # Все внутренние ячейки - одиночный шум
        return False, inner_cells

    def _find_connected_groups(self, cells: List[Cell]) -> List[List[Cell]]:
        """
        Находит связные группы ячеек (прижатых друг к другу).

        Args:
            cells: Список ячеек.

        Returns:
            Список групп, где каждая группа — список связанных ячеек (таблица).
        """
        if not cells:
            return []

        n = len(cells)

        # Строим матрицу смежности
        adjacency = {i: [] for i in range(n)}

        for i in range(n):
            for j in range(i + 1, n):
                if cells[i].is_pressed_to_borders(cells[j], self.cell_tolerance):
                    adjacency[i].append(j)
                    adjacency[j].append(i)

        # Поиск компонент связности (DFS)
        visited = set()
        groups = []

        for i in range(n):
            if i not in visited:
                component = []
                stack = [i]

                while stack:
                    node = stack.pop()
                    if node not in visited:
                        visited.add(node)
                        component.append(cells[node])
                        stack.extend(adjacency[node])

                groups.append(component)

        return groups

    def _is_structured_group(self, cells: List[Cell]) -> bool:
        """
        Проверяет, образуют ли ячейки структурированную группу (таблицу).

        Структурированная группа:
        - ≥2 ячеек
        - Все ячейки связаны (касаются друг друга)
        - Образуют подобие сетки (выровнены по строкам/столбцам)

        Args:
            cells: Список ячеек одной связной группы.

        Returns:
            True, если группа структурирована.
        """
        if len(cells) < 2:
            return False

        # Проверка выровненности по строкам и столбцам

        # Группируем по Y
        y_groups = {}
        for cell in cells:
            # Ищем ближайшую Y-группу
            matched = False
            for base_y in list(y_groups.keys()):
                if abs(cell.y - base_y) <= self.row_tolerance:
                    y_groups[base_y].append(cell)
                    matched = True
                    break
            if not matched:
                y_groups[cell.y] = [cell]

        rows = list(y_groups.values())

        # Группируем по X
        x_groups = {}
        for cell in cells:
            matched = False
            for base_x in list(x_groups.keys()):
                if abs(cell.x - base_x) <= self.col_tolerance:
                    x_groups[base_x].append(cell)
                    matched = True
                    break
            if not matched:
                x_groups[cell.x] = [cell]

        columns = list(x_groups.values())

        # Если есть и строки, и столбцы (хотя бы 2x1 или 1x2) — это структура
        # Также проверяем, что большинство строк содержат одинаковое количество ячеек
        if len(rows) >= 2 or len(columns) >= 2:
            # Проверяем регулярность строк
            row_sizes = [len(r) for r in rows]
            if len(row_sizes) >= 2:
                most_common_size = max(set(row_sizes), key=row_sizes.count)
                regular_rows = sum(1 for sz in row_sizes if sz == most_common_size)
                if regular_rows >= len(rows) * 0.6:
                    return True

            # Проверяем регулярность столбцов
            col_sizes = [len(c) for c in columns]
            if len(col_sizes) >= 2:
                most_common_size = max(set(col_sizes), key=col_sizes.count)
                regular_cols = sum(1 for sz in col_sizes if sz == most_common_size)
                if regular_cols >= len(columns) * 0.6:
                    return True

            # Если есть и строки, и столбцы — структура
            if len(rows) >= 1 and len(columns) >= 1:
                return True

        return False

    def _group_cells_by_rows(self, cells: List[Cell]) -> List[List[Cell]]:
        """
        Группирует ячейки по строкам (по Y-координате).
        """
        # TODO: добавить автодополнение нехватающих ячеек

        sorted_cells = sorted(cells, key=lambda c: c.y)

        rows = []
        current_row = []
        current_y = None

        for cell in sorted_cells:
            y = cell.y

            if current_y is None:
                current_y = y
                current_row.append(cell)
            elif abs(y - current_y) <= self.row_tolerance:
                current_row.append(cell)
            else:
                if current_row:
                    current_row.sort(key=lambda c: c.x)
                    rows.append(current_row)
                current_y = y
                current_row = [cell]

        # Последняя строка
        if current_row:
            current_row.sort(key=lambda c: c.x)
            rows.append(current_row)

        return rows

    def _group_rows_by_multirows(self, rows: List[List[Cell]]) -> List[List[str]]:
        """
        Объединяет многоуровневые строки в одну.

        Args:
            rows: Отсортированный список строк с ячейками.
        """
        if not rows:
            return []

        current_rows = []
        current_row = []
        current_start = 0
        current_height = 0
        row_num = 0
        for i, row in enumerate(rows):
            # Начинаем новую строку
            if not current_row:
                row_num = i
                current_start = row[0].y
                current_height = Cell.get_max_row_height(row)
                current_row.append(row)
                continue

            row_height = Cell.get_max_row_height(row)
            row_start = row[0].y
            # Новая строка является частью многоуровневой строки
            if row_start + row_height <= current_start + current_height + self.cell_tolerance:
                current_row.append(row)
            else:
                current_rows.append(self._group_rows_by_row(current_row, row_num))
                current_row = []
                row_num = i
                current_start = row[0].y
                current_height = Cell.get_max_row_height(row)
                current_row.append(row)

        # Последняя строка
        if current_row:
            current_rows.append(self._group_rows_by_row(current_row, row_num))

        return current_rows

    def _group_rows_by_row(self, rows: List[List[Cell]], row_num: int) -> List[str]:
        """
        Объединяет строки rows в одну.

        Args:
            rows: Список строк с ячейками.
            row_num: Номер первой строки в списке.

        Returns:
            Список строк с объединенными значениями столбцов.
        """
        if not rows:
            return []

        # Инициализируем базовые столбцы из первой строки
        # Каждый столбец описывается: {'x_start': int, 'x_end': int, 'text': str}
        current_columns = []
        for i, cell in enumerate(rows[0]):
            if cell.text == '':
                cell.text = f"None{row_num}_{i}"
            current_columns.append({
                'x_start': cell.x,
                'x_end': cell.x_end,
                'text': cell.text.strip()
            })

        # Обрабатываем последующие строки
        for row_idx in range(1, len(rows)):
            row_cells = rows[row_idx]

            if not row_cells:
                continue

            for i, cell in enumerate(row_cells):
                cell_x_start = cell.x
                cell_x_end = cell.x_end
                cell_text = cell.text.strip()

                if not cell_text:
                    cell_text = f"None{row_num + row_idx}_{i}"

                new_columns = []
                for col in current_columns:
                    # Проверяем пересечение по X (с погрешностью)
                    if (
                        col['x_start'] > cell_x_start - self.cell_tolerance and
                        col['x_end'] < cell_x_end + self.cell_tolerance or
                        cell_x_start > col['x_start'] - self.cell_tolerance and
                        cell_x_end < col['x_end'] + self.cell_tolerance
                    ):
                        new_columns.append({
                            'x_start': max(cell_x_start, col['x_start']),
                            'x_end': min(cell_x_end, col['x_end']),
                            'text': f"{col['text']}/{cell_text}"
                        })
                        # Если ячейка меньше, чем та, что выше
                        # И не является последней, то нужно добавить дополнительный столбец
                        if cell_x_end < col['x_end'] - self.cell_tolerance:
                            new_columns.append(col)
                    else:
                        new_columns.append(col)

                if new_columns:
                    # Сортируем по X
                    new_columns.sort(key=lambda c: c['x_start'])
                    current_columns = new_columns

        # Извлекаем итоговые тексты
        result = [col['text'] for col in current_columns]

        return result

    def _ocr_text(self, img: np.ndarray, rows: List[List[Cell]]) -> List[List[Cell]]:
        """
        Распознаёт текст в матрице ячеек. Добавляет ячейкам параметр text.
        """
        if not rows:
            return []

        rows_with_text = []
        # Распознаём текст во всех ячейках
        for i, row in enumerate(rows):
            row_with_text = []
            for j, cell in enumerate(row):
                # Вырезаем ячейку
                x1, y1, x2, y2 = cell.bbox
                cell_img = img[y1:y2, x1:x2]

                text = normalize_string(self._extract_text(cell_img))
                cell.text = text
                row_with_text.append(cell)

            rows_with_text.append(row_with_text)

        return rows_with_text
