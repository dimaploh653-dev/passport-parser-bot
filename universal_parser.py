import os
import pandas as pd
from docx import Document
import re

def extract_data_from_docx(path):
    """Извлекает данные из Word-файла"""
    doc = Document(path)
    text = "\n".join([p.text for p in doc.paragraphs])

    # Примеры извлечения данных (ты можешь адаптировать под свои шаблоны)
    name_match = re.search(r"Фамилия\s*[:\-]?\s*(\w+)", text)
    surname = name_match.group(1).capitalize() if name_match else ""

    passport_match = re.search(r"(\d{2}\s?\d{2}\s?\d{6})", text)
    passport = passport_match.group(1) if passport_match else ""

    english_name_match = re.search(r"Name\s*[:\-]?\s*([A-Za-z\s]+)", text)
    english_name = english_name_match.group(1).title() if english_name_match else ""

    return {
        "Фамилия": surname,
        "Фамилия (EN)": english_name,
        "Паспорт": passport,
        "Файл": os.path.basename(path)
    }

def process_word_files(file_paths, output_path):
    """Основная функция: получает список файлов и путь для сохранения"""
    all_data = []
    for path in file_paths:
        try:
            data = extract_data_from_docx(path)
            all_data.append(data)
        except Exception as e:
            all_data.append({
                "Фамилия": "",
                "Фамилия (EN)": "",
                "Паспорт": "",
                "Файл": f"{os.path.basename(path)} (ошибка: {e})"
            })

    df = pd.DataFrame(all_data)
    df.to_excel(output_path, index=False)
    print(f"✅ Результат сохранён в {output_path}")
