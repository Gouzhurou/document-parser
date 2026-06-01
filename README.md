# document-parser
A tool for converting engineering documentation into a structures JSON format for automated analysis

## Возможности

- Парсинг Excel-файлов (.xls, .xlsx, .xlsm)
- Парсинг документов Word (.docx, .doc)
- Парсинг изображений (.jpg, .jpeg, .png, .tiff, .tif, .bpm) с помощью OCR
- Автоматическое нахождение таблиц и многоуровневых заголовков
- Извлечение метаданных (название документа, объект, заказчик, подрядчик, дата и прочее)
- Экспорт результата в JSON

## Установка Tesseract OCR

1. Установка загрузчика с [github](https://github.com/UB-Mannheim/tesseract/wiki)
2. Запуск загрузчика и установка русского языка
2. Использование загрузчика Tesseract-OCR/tesseract.exe в коде

## Установка зависимостей 

```commandline
pip install -e .
```
