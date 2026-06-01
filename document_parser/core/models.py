"""
Модели данных для представления результатов парсинга.
"""

from dataclasses import dataclass, field
from typing import List, Dict, Any, Optional


@dataclass
class TableMetadata:
    """Метаданные таблицы, извлечённые из документа."""
    document_name: Optional[str] = None
    # TODO: убрать взятие 2ой строки, если первая написана капсом. Нужно брать 2ую, если в первой одно слово
    document_number: Optional[str] = None
    object_name: Optional[str] = None
    contract_number: Optional[str] = None
    compilation_date: Optional[str] = None
    customer: Optional[str] = None
    contractor: Optional[str] = None
    compiled: Optional[str] = None
    checked: Optional[str] = None
    note: Optional[str] = None
    extra: Dict[str, Any] = field(default_factory=dict)

    # TODO: реализовать добавление всего текста, который никуда не подошел в экстра

    def __repr__(self) -> str:
        """Строковое представление для отладки."""
        fields = []

        if self.document_name:
            fields.append(f"document_name='{self.document_name}'")
        if self.document_number:
            fields.append(f"document_number='{self.document_number}'")
        if self.object_name:
            object_name_preview = self.object_name[:50] + "..." if len(self.object_name) > 50 else self.object_name
            fields.append(f"object_name='{object_name_preview}'")
        if self.contract_number:
            fields.append(f"contract_number='{self.contract_number}'")
        if self.compilation_date:
            fields.append(f"compilation_date='{self.compilation_date}'")
        if self.customer:
            fields.append(f"customer='{self.customer}'")
        if self.contractor:
            fields.append(f"contractor='{self.contractor}'")
        if self.compiled:
            fields.append(f"compiled='{self.compiled}'")
        if self.checked:
            fields.append(f"checked='{self.checked}'")
        if self.note:
            # Обрезаем длинное примечание
            note_preview = self.note[:50] + "..." if len(self.note) > 50 else self.note
            fields.append(f"note='{note_preview}'")
        if self.extra:
            fields.append(f"extra={self.extra}")

        return f"TableMetadata({', '.join(fields)})"

    def to_dict(self) -> Dict[str, Any]:
        """
        Сериализация метаданных в словарь.
        Включает только непустые поля.
        """
        result = {}

        if self.document_name is not None:
            result['document_name'] = self.document_name
        if self.document_number is not None:
            result['document_number'] = self.document_number
        if self.object_name is not None:
            result['object_name'] = self.object_name
        if self.contract_number is not None:
            result['contract_number'] = self.contract_number
        if self.compilation_date is not None:
            result['compilation_date'] = self.compilation_date
        if self.customer is not None:
            result['customer'] = self.customer
        if self.contractor is not None:
            result['contractor'] = self.contractor
        if self.compiled is not None:
            result['compiled'] = self.compiled
        if self.checked is not None:
            result['checked'] = self.checked
        if self.note is not None:
            result['note'] = self.note

        result.update(self.extra)

        return result


@dataclass
class ParsedTable:
    """
    Представляет одну таблицу, извлечённую из документа.
    """
    sheet_name: str | None
    headers: List[str]
    rows: List[Dict[str, Any]]
    metadata: Optional[TableMetadata] = None

    def to_dict(self) -> Dict[str, Any]:
        """Сериализация в словарь."""
        result = {
            "sheet_name": self.sheet_name,
            "headers": self.headers,
            "rows": self.rows,
        }
        if self.metadata:
            result["metadata"] = self.metadata.to_dict()
        return result


@dataclass
class ParseResult:
    """
    Результат парсинга всего документа.
    """
    filename: str | None
    tables: List[ParsedTable] = field(default_factory=list)
    success: bool = True
    error: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        """Сериализация в словарь."""
        return {
            "filename": self.filename,
            "success": self.success,
            "error": self.error,
            "tables": [t.to_dict() for t in self.tables],
        }