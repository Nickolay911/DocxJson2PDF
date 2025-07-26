# DOCX Template Processor → PDF

A Python script for automatic field replacement in Word documents and conversion to PDF.

## Installation

```bash
pip3 install -r requirements.txt
```

## Basic Usage

```bash
python3 template_processor.py <template.docx> <parameters.json> <output.pdf>
```

### Parameters:
- `template.docx` - source Word document with fields to replace
- `parameters.json` - JSON file with replacement values
- `output.pdf` - output PDF file

## Examples

### Basic example:
```bash
python3 template_processor.py tpl.docx params.json document.pdf
```

### JSON file format:
```json
{
    "{{fam}}": "Smith",
    "{{im}}": "John", 
    "{{ot}}": "Michael"
}
```

### Template fields:
Use fields in the format `{{field_name}}` in your Word document:
- `{{fam}}` - surname/family name
- `{{im}}` - first name  
- `{{ot}}` - middle name/patronymic

## Features
- ✅ Replacement in text, tables, headers, footers
- ✅ UTF-8 support for international characters
- ✅ Automatic conversion via LibreOffice
- ✅ Error checking and input validation
- ✅ Command-line interface with clear feedback
- ✅ Temporary file cleanup

## Requirements
- Python 3.6+
- LibreOffice (for PDF conversion)
- python-docx library

## Installing LibreOffice:

### Ubuntu/Debian:
```bash
sudo apt install libreoffice
```

### Windows:
Download from [LibreOffice official website](https://www.libreoffice.org/download/download/)

### macOS:
```bash
brew install --cask libreoffice
```

## Error Handling
The script includes comprehensive error handling for:
- Missing input files
- Invalid JSON format
- LibreOffice conversion failures
- File permission issues
- Timeout during conversion

## Author
Nickolay911

## License
This project is open source and available under standard terms. 

---

# Обработчик шаблонов DOCX → PDF

Скрипт для автоматической замены полей в документах Word и конвертации в PDF.

## Установка зависимостей

```bash
pip3 install -r requirements.txt
```

## Основное использование

```bash
python3 template_processor.py <шаблон.docx> <параметры.json> <результат.pdf>
```

### Параметры:
- `шаблон.docx` - исходный документ Word с полями для замены
- `параметры.json` - JSON файл с значениями для подстановки  
- `результат.pdf` - выходной PDF файл

## Примеры

### Базовый пример:
```bash
python3 template_processor.py tpl.docx params.json document.pdf
```

### Формат JSON файла:
```json
{
    "{{fam}}": "Иванов",
    "{{im}}": "Петр", 
    "{{ot}}": "Александрович"
}
```

### Поля в шаблоне:
В документе Word используйте поля вида `{{имя_поля}}`:
- `{{fam}}` - фамилия
- `{{im}}` - имя  
- `{{ot}}` - отчество

## Возможности
- ✅ Замена в тексте, таблицах, заголовках, футерах
- ✅ Поддержка UTF-8 для русских символов
- ✅ Автоматическая конвертация через LibreOffice
- ✅ Проверка ошибок и валидация входных данных

## Требования
- Python 3.6+
- LibreOffice (для конвертации в PDF)
- python-docx

## Установка LibreOffice (Ubuntu/Debian):
```bash
sudo apt install libreoffice
```
