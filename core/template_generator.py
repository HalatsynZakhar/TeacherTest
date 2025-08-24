import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
import os

def create_test_template(output_path: str):
    """
    Створює Excel шаблон для тестів з детальними інструкціями.
    
    Args:
        output_path: Шлях для збереження шаблону
    """
    # Створюємо нову книгу
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Шаблон тесту"
    
    # Стилі для заголовків та інструкцій
    header_font = Font(bold=True, size=14, color="FFFFFF")
    instruction_font = Font(size=11, color="000000")
    example_font = Font(size=10, color="666666", italic=True)
    
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    instruction_fill = PatternFill(start_color="E7F3FF", end_color="E7F3FF", fill_type="solid")
    
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Інструкції в першому рядку
    instructions = [
        "ІНСТРУКЦІЯ ПО ЗАПОВНЕННЮ ТЕСТУ:",
        "1. Номер питання (колонка A): Вкажіть номер питання. Якщо декілька питань мають однаковий номер (наприклад, 1, 1, 1, 2, 2), система обере випадкове питання з кожної групи.",
        "2. Питання (колонка B): Текст питання.",
        "3. Правильна відповідь/Тип (колонка C): Для питань з варіантами відповідей - номер правильної відповіді (1, 2, 3, 4). Для відкритих питань - правильна відповідь текстом.",
        "4. Вага (колонка D): Числове значення складності питання (1-легке, 5-складне). Використовується для сортування від легкого до складного.",
        "5. Варіанти відповідей (колонки E-H): Варіанти 1, 2, 3, 4. Залиште порожніми для відкритих питань.",
        "ВАЖЛИВО: Заповнення починається з рядка 2! Видаліть цю інструкцію перед завантаженням файлу в систему."
    ]
    
    # Об'єднуємо комірки для інструкцій
    ws.merge_cells('A1:H1')
    ws['A1'] = '\n'.join(instructions)
    ws['A1'].font = instruction_font
    ws['A1'].fill = instruction_fill
    ws['A1'].alignment = Alignment(wrap_text=True, vertical='top')
    ws['A1'].border = border
    
    # Встановлюємо висоту першого рядка
    ws.row_dimensions[1].height = 150
    
    # Заголовки колонок у другому рядку
    headers = ['Номер питання', 'Питання', 'Правильна відповідь/Тип', 'Вага', 'Варіант 1', 'Варіант 2', 'Варіант 3', 'Варіант 4']
    
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border
    
    # Приклади даних
    example_data = [
        [1, "Яка мова програмування використовується для веб-розробки?", "1", 2, "JavaScript", "Python", "C++", "Java"],
        [1, "Що таке HTML?", "2", 1, "Мова програмування", "Мова розмітки", "База даних", "Операційна система"],
        [2, "Яке максимальне число в двійковій системі з 2 розрядів?", "3", 3, "", "", "", ""],
        [3, "Скільки буде 2+2?", "3", 1, "3", "5", "4", "6"]
    ]
    
    for row_idx, example in enumerate(example_data, 3):
        for col_idx, value in enumerate(example, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = example_font
            cell.border = border
            if col_idx <= 4:  # Основні колонки
                cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Налаштування ширини колонок
    column_widths = [15, 50, 20, 10, 25, 25, 25, 25]
    for col, width in enumerate(column_widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = width
    
    # Створюємо директорію якщо не існує
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    
    # Зберігаємо файл
    wb.save(output_path)