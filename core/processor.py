import pandas as pd
import random
import os
from typing import List, Dict, Any, Tuple, Optional
from fpdf import FPDF
import logging
from datetime import datetime
import textwrap
from docx import Document
from docx.shared import Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
import tempfile

log = logging.getLogger(__name__)

def ensure_temp_dir(prefix="temp_"):
    """Создает временную папку внутри проекта"""
    try:
        # Получаем корневую папку проекта
        project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        temp_dir = os.path.join(project_root, "temp")
        
        # Создаем папку temp если её нет
        os.makedirs(temp_dir, exist_ok=True)
        
        # Создаем подпапку с префиксом
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        specific_temp_dir = os.path.join(temp_dir, f"{prefix}{timestamp}")
        os.makedirs(specific_temp_dir, exist_ok=True)
        
        return specific_temp_dir
    except Exception as e:
        log.warning(f"Не удалось создать временную папку: {e}. Используем системную временную папку.")
        return tempfile.gettempdir()

def get_text_width(pdf: FPDF, text: str) -> float:
    """Получить ширину текста в текущем шрифте"""
    return pdf.get_string_width(text)

def fit_text_to_width(pdf: FPDF, text: str, max_width: float, min_font_size: int = 8) -> Tuple[str, int]:
    """Подогнать текст под заданную ширину, уменьшая шрифт или разбивая на строки"""
    current_font = pdf.font_family
    current_style = pdf.font_style
    current_size = pdf.font_size_pt
    
    # Попробуем уменьшить шрифт
    for font_size in range(int(current_size), min_font_size - 1, -1):
        pdf.set_font(current_font, current_style, font_size)
        if get_text_width(pdf, text) <= max_width:
            return text, font_size
    
    # Если не помогло уменьшение шрифта, разбиваем текст
    pdf.set_font(current_font, current_style, min_font_size)
    words = text.split()
    lines = []
    current_line = ""
    
    for word in words:
        test_line = current_line + (" " if current_line else "") + word
        if get_text_width(pdf, test_line) <= max_width:
            current_line = test_line
        else:
            if current_line:
                lines.append(current_line)
                current_line = word
            else:
                # Слово слишком длинное, принудительно разбиваем
                lines.append(word)
    
    if current_line:
        lines.append(current_line)
    
    return "\n".join(lines), min_font_size

def add_multiline_text(pdf: FPDF, text: str, max_width: float, line_height: float = 6, min_font_size: int = 8):
    """Добавить многострочный текст с автоматическим переносом"""
    fitted_text, font_size = fit_text_to_width(pdf, text, max_width, min_font_size)
    
    current_font = pdf.font_family
    current_style = pdf.font_style
    pdf.set_font(current_font, current_style, font_size)
    
    lines = fitted_text.split('\n')
    for line in lines:
        pdf.cell(0, line_height, line, ln=True)
    
    return len(lines) * line_height

def read_test_excel(file_path: str) -> pd.DataFrame:
    """
    Читает Excel файл с вопросами теста.
    
    Ожидаемая структура:
    - Столбец 0: Текст вопроса
    - Столбец 1: Номер правильного ответа (или пустой для нетестовых заданий)
    - Столбец 2: Вес задания (по умолчанию 1)
    - Столбцы 3+: Варианты ответов (опционально для нетестовых заданий)
    
    Args:
        file_path: Путь к Excel файлу
        
    Returns:
        DataFrame с вопросами теста
    """
    try:
        # Читаем Excel файл
        df = pd.read_excel(file_path, header=None)
        
        # Конвертируем только столбец с вопросами в строки
        # Остальные столбцы оставляем как есть для сохранения исходного форматирования
        df[df.columns[0]] = df[df.columns[0]].astype(str)
        
        # Проверяем минимальную структуру (вопрос + правильный ответ + вес + минимум 2 варианта)
        if df.shape[1] < 3:
            raise ValueError("Файл должен содержать минимум 3 столбца: вопрос, правильный ответ/тип, вес задания")
        
        # Удаляем пустые строки (проверяем только наличие вопроса)
        df = df[df.iloc[:, 0].notna() & (df.iloc[:, 0] != 'nan')]  # Удаляем строки где нет вопроса
        
        if df.empty:
            raise ValueError("Файл не содержит валидных данных")
        
        # Переименовываем столбцы для удобства
        columns = ['question', 'correct_answer', 'weight'] + [f'option_{i}' for i in range(1, df.shape[1] - 2)]
        df.columns = columns
        
        # Обрабатываем вес задания (по умолчанию 1)
        df['weight'] = pd.to_numeric(df['weight'], errors='coerce')
        df['weight'] = df['weight'].fillna(1.0)  # Заполняем пустые значения единицей
        
        # Определяем тип задания: тестовое (с вариантами) или открытое (без вариантов)
        # Подсчитываем количество непустых вариантов ответов для каждого вопроса
        option_cols = [col for col in df.columns if col.startswith('option_')]
        df['option_count'] = 0
        for col in option_cols:
            df['option_count'] += df[col].notna() & (df[col] != 'nan') & (df[col].astype(str).str.strip() != '')
        
        # Проверяем обязательное заполнение правильного ответа
        missing_answers = df['correct_answer'].isna() | (df['correct_answer'] == 'nan') | (df['correct_answer'].astype(str).str.strip() == '')
        if missing_answers.any():
            missing_questions = df[missing_answers]['question'].tolist()
            log.warning(f"Найдены вопросы без правильного ответа: {missing_questions[:3]}{'...' if len(missing_questions) > 3 else ''}")
            df = df[~missing_answers]  # Удаляем вопросы без правильного ответа
        
        # Задание считается тестовым если есть 2 или более вариантов ответов
        # Если 0 или 1 вариант ответа, то это открытое задание
        df['is_test_question'] = (df['option_count'] >= 2) & df['correct_answer'].notna() & (df['correct_answer'] != 'nan')
        
        # Приводим столбец correct_answer к object типу для избежания предупреждений
        df['correct_answer'] = df['correct_answer'].astype('object')
        
        # Для тестовых заданий проверяем правильные ответы
        test_mask = df['is_test_question']
        if test_mask.any():
            # Сначала преобразуем в числа для валидации
            numeric_answers = pd.to_numeric(df.loc[test_mask, 'correct_answer'], errors='coerce')
            # Удаляем тестовые вопросы с некорректными ответами
            valid_test_mask = test_mask & numeric_answers.notna()
            df = df[~(test_mask & numeric_answers.isna())]
            
            # Форматируем правильные ответы для тестовых заданий (убираем .0 для целых чисел)
            for idx in df[valid_test_mask].index:
                if idx in df.index:  # Проверяем, что индекс еще существует после фильтрации
                    answer_value = numeric_answers.loc[idx]
                    if answer_value == int(answer_value):
                        df.at[idx, 'correct_answer'] = str(int(answer_value))
                    else:
                        df.at[idx, 'correct_answer'] = str(answer_value)
        
        # Для открытых заданий сохраняем правильный ответ как текст
        open_mask = ~df['is_test_question'] & df['correct_answer'].notna() & (df['correct_answer'] != 'nan')
        if open_mask.any():
            # Для открытых заданий форматируем ответ правильно
            for idx in df[open_mask].index:
                answer_value = df.loc[idx, 'correct_answer']
                # Если это число, форматируем без лишних .0
                if isinstance(answer_value, (int, float)) and answer_value == int(answer_value):
                    df.at[idx, 'correct_answer'] = str(int(answer_value))
                else:
                    df.at[idx, 'correct_answer'] = str(answer_value).strip()
        
        # Удаляем временный столбец
        df = df.drop('option_count', axis=1)
        
        # Приводим столбец correct_answer к строковому типу для совместимости с PyArrow
        df['correct_answer'] = df['correct_answer'].astype(str)
        
        log.info(f"Загружено {len(df)} вопросов из файла {file_path}")
        return df
        
    except Exception as e:
        log.error(f"Ошибка при чтении файла {file_path}: {e}")
        raise

def generate_test_variants(df: pd.DataFrame, num_variants: int) -> List[Dict[str, Any]]:
    """
    Генерирует варианты тестов с перемешанными вопросами и ответами.
    
    Args:
        df: DataFrame с вопросами
        num_variants: Количество вариантов для генерации
        
    Returns:
        Список словарей с вариантами тестов
    """
    variants = []
    
    for variant_num in range(1, num_variants + 1):
        variant = {
            'variant_number': variant_num,
            'questions': [],
            'answer_key': []
        }
        
        # Перемешиваем порядок вопросов
        shuffled_df = df.sample(frac=1, random_state=variant_num).reset_index(drop=True)
        
        for idx, row in shuffled_df.iterrows():
            question_data = {
                'question_text': str(row['question']),
                'weight': float(row['weight']),
                'is_test_question': bool(row['is_test_question'])
            }
            
            if row['is_test_question']:
                # Тестовое задание с вариантами ответов
                options = []
                for col in df.columns:
                    if col.startswith('option_') and pd.notna(row[col]) and str(row[col]).strip() != '' and str(row[col]) != 'nan':
                        # Сохраняем исходное форматирование чисел
                        value = row[col]
                        if isinstance(value, (int, float)):
                            # Для чисел: целые без .0, дробные как есть
                            if isinstance(value, float) and value.is_integer():
                                options.append(str(int(value)))
                            else:
                                options.append(str(value))
                        else:
                            options.append(str(value).strip())
                
                if len(options) < 2:
                    log.warning(f"Тестовый вопрос '{row['question']}' имеет менее 2 вариантов ответов, пропускаем")
                    continue
                
                # Проверяем корректность индекса правильного ответа
                correct_answer_idx = int(row['correct_answer']) - 1  # -1 так как нумерация с 1
                if correct_answer_idx < 0 or correct_answer_idx >= len(options):
                    log.warning(f"Некорректный индекс правильного ответа {row['correct_answer']} для вопроса '{row['question']}', пропускаем")
                    continue
                
                # Находим правильный ответ по индексу
                correct_option_text = options[correct_answer_idx]
                
                # Перемешиваем варианты ответов
                random.seed(variant_num + idx)  # Для воспроизводимости
                shuffled_options = options.copy()
                random.shuffle(shuffled_options)
                
                # Находим новую позицию правильного ответа
                new_correct_position = shuffled_options.index(correct_option_text) + 1  # +1 для нумерации с 1
                
                question_data.update({
                    'options': shuffled_options,
                    'correct_answer': new_correct_position
                })
            else:
                # Открытое задание (без вариантов ответов)
                # Обрабатываем правильный ответ как текстовые данные
                formatted_answer = str(row['correct_answer']).strip()
                
                question_data.update({
                    'correct_answer': formatted_answer,
                    'options': []  # Открытые вопросы не имеют вариантов ответов
                })
            
            variant['questions'].append(question_data)
            
            # Добавляем в ключ ответов
            if row['is_test_question']:
                variant['answer_key'].append(new_correct_position)
            else:
                variant['answer_key'].append(formatted_answer)
        
        variants.append(variant)
        log.info(f"Сгенерирован вариант {variant_num} с {len(variant['questions'])} вопросами")
    
    return variants

def create_test_pdf(variants: List[Dict[str, Any]], output_dir: str, columns: int = 1) -> Tuple[str, str]:
    """
    Создает PDF файлы с тестами для учеников и ответами для учителя.
    
    Args:
        variants: Список вариантов тестов
        output_dir: Папка для сохранения файлов
        columns: Количество колонок для размещения вопросов (1-3)
        
    Returns:
        Кортеж (путь к файлу с тестами, путь к файлу с ответами)
    """
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
    test_pdf_path = os.path.join(output_dir, f"tests_{timestamp}.pdf")
    answers_pdf_path = os.path.join(output_dir, f"answers_{timestamp}.pdf")
    
    # Создаем PDF с тестами для учеников
    test_pdf = FPDF()
    test_pdf.add_font('Arial', '', 'c:/windows/fonts/arial.ttf', uni=True)
    test_pdf.add_font('Arial', 'B', 'c:/windows/fonts/arialbd.ttf', uni=True)
    
    page_width = test_pdf.w - 2 * test_pdf.l_margin  # Ширина страницы без полей
    
    for variant in variants:
        test_pdf.add_page()
        test_pdf.set_font('Arial', 'B', 16)
        test_pdf.cell(0, 10, f"Тест - Варіант {variant['variant_number']}", ln=True, align='C')
        test_pdf.ln(10)
        
        test_pdf.set_font('Arial', '', 12)
        test_pdf.cell(0, 8, "Інструкція: Оберіть правильну відповідь і впишіть її номер у таблицю внизу.", ln=True)
        test_pdf.ln(5)
        
        # Добавляем вопросы с поддержкой колонок
        columns = max(1, min(3, columns))  # Ограничиваем от 1 до 3 колонок
        questions = variant['questions']
        
        if columns == 1:
            # Одна колонка - простое размещение
            for q_idx, question in enumerate(questions):
                q_num = q_idx + 1
                
                # Проверяем место на странице
                if test_pdf.get_y() > 220:  # Оставляем место для таблицы ответов
                    test_pdf.add_page()
                    test_pdf.set_font('Arial', 'B', 16)
                    test_pdf.cell(0, 10, f"Тест - Варіант {variant['variant_number']} (продовження)", ln=True, align='C')
                    test_pdf.ln(10)
                
                test_pdf.set_font('Arial', 'B', 11)
                question_text = f"{q_num}. {question['question_text']}"
                add_multiline_text(test_pdf, question_text, page_width, 6, 9)
                
                test_pdf.set_font('Arial', '', 10)
                for opt_num, option in enumerate(question['options'], 1):
                    option_text = f"   {opt_num}) {option}"
                    add_multiline_text(test_pdf, option_text, page_width, 5, 8)
                
                test_pdf.ln(3)
        else:
            # Многоколоночное размещение
            column_width = (page_width - (columns - 1) * 10) / columns
            
            # Массив для отслеживания Y позиций каждой колонки
            column_y_positions = [test_pdf.get_y()] * columns
            start_y = test_pdf.get_y()
            
            for q_idx, question in enumerate(questions):
                q_num = q_idx + 1
                col = q_idx % columns
                x_pos = test_pdf.l_margin + col * (column_width + 10)
                
                # Проверяем место на странице - используем максимальную Y позицию среди всех колонок
                max_y = max(column_y_positions)
                if max_y > 200:  # Уменьшаем порог для лучшего контроля
                    test_pdf.add_page()
                    test_pdf.set_font('Arial', 'B', 16)
                    test_pdf.cell(0, 10, f"Тест - Варіант {variant['variant_number']} (продовження)", ln=True, align='C')
                    test_pdf.ln(10)
                    column_y_positions = [test_pdf.get_y()] * columns
                    start_y = test_pdf.get_y()
                
                # Устанавливаем позицию для текущей колонки
                test_pdf.set_xy(x_pos, column_y_positions[col])
                
                test_pdf.set_font('Arial', 'B', 10)
                question_text = f"{q_num}. {question['question_text']}"
                
                # Разбиваем текст на строки
                fitted_text, font_size = fit_text_to_width(test_pdf, question_text, column_width, 8)
                test_pdf.set_font('Arial', 'B', font_size)
                
                lines = fitted_text.split('\n')
                for line in lines:
                    test_pdf.cell(column_width, 5, line, ln=True)
                    test_pdf.set_x(x_pos)
                
                test_pdf.set_font('Arial', '', 9)
                for opt_num, option in enumerate(question['options'], 1):
                    option_text = f"   {opt_num}) {option}"
                    fitted_option, opt_font_size = fit_text_to_width(test_pdf, option_text, column_width, 7)
                    test_pdf.set_font('Arial', '', opt_font_size)
                    
                    opt_lines = fitted_option.split('\n')
                    for opt_line in opt_lines:
                        test_pdf.cell(column_width, 4, opt_line, ln=True)
                        test_pdf.set_x(x_pos)
                
                # Обновляем Y позицию для текущей колонки
                column_y_positions[col] = test_pdf.get_y() + 3
            
            # Устанавливаем позицию после всех колонок
            test_pdf.set_xy(test_pdf.l_margin, max(column_y_positions))
        
        # Добавляем таблицу для ответов
        test_pdf.ln(8)
        test_pdf.set_font('Arial', 'B', 11)
        test_pdf.cell(0, 6, "Таблиця відповідей:", ln=True)
        test_pdf.ln(3)
        
        # Создаем компактную таблицу ответов с одинаковыми ячейками
        test_pdf.set_font('Arial', '', 9)
        questions_per_row = 15  # Количество вопросов в строке
        num_questions = len(variant['questions'])
        
        # Фиксированная ширина ячейки для равномерности
        fixed_cell_width = page_width / questions_per_row
        
        for row_start in range(0, num_questions, questions_per_row):
            questions_in_row = min(questions_per_row, num_questions - row_start)
            
            # Номера вопросов - все ячейки одинакового размера
            test_pdf.set_font('Arial', 'B', 8)
            for i in range(questions_in_row):
                test_pdf.cell(fixed_cell_width, 6, f"№{row_start + i + 1}", 1, 0, 'C')
            # Заполняем оставшиеся ячейки пустыми для выравнивания
            for i in range(questions_in_row, questions_per_row):
                test_pdf.cell(fixed_cell_width, 6, "", 1, 0, 'C')
            test_pdf.ln()
            
            # Пустые ячейки для ответов - все одинакового размера
            test_pdf.set_font('Arial', '', 8)
            for i in range(questions_in_row):
                test_pdf.cell(fixed_cell_width, 8, "", 1, 0, 'C')
            # Заполняем оставшиеся ячейки пустыми для выравнивания
            for i in range(questions_in_row, questions_per_row):
                test_pdf.cell(fixed_cell_width, 8, "", 1, 0, 'C')
            test_pdf.ln()
            test_pdf.ln(2)
    
    test_pdf.output(test_pdf_path)
    
    # Создаем PDF с ответами для учителя
    answers_pdf = FPDF()
    answers_pdf.add_font('Arial', '', 'c:/windows/fonts/arial.ttf', uni=True)
    answers_pdf.add_font('Arial', 'B', 'c:/windows/fonts/arialbd.ttf', uni=True)
    
    answer_page_width = answers_pdf.w - 2 * answers_pdf.l_margin
    
    answers_pdf.add_page()
    answers_pdf.set_font('Arial', 'B', 16)
    answers_pdf.cell(0, 10, "Відповіді для вчителя", ln=True, align='C')
    answers_pdf.ln(10)
    
    # Группируем варианты по больше на страницу
    variants_per_page = 8  # Увеличиваем количество вариантов на странице
    for page_start in range(0, len(variants), variants_per_page):
        if page_start > 0:
            answers_pdf.add_page()
        
        for variant in variants[page_start:page_start + variants_per_page]:
            answers_pdf.set_font('Arial', 'B', 12)
            variant_text = f"Варіант {variant['variant_number']}"
            add_multiline_text(answers_pdf, variant_text, answer_page_width, 6, 9)
            
            answers_pdf.set_font('Arial', '', 10)
            # Выводим ответы более компактно
            answer_text = "Відповіді: " + ", ".join([f"{i+1}-{ans}" for i, ans in enumerate(variant['answer_key'])])
            add_multiline_text(answers_pdf, answer_text, answer_page_width, 5, 8)
            answers_pdf.ln(3)  # Уменьшаем отступ между вариантами
    
    answers_pdf.output(answers_pdf_path)
    
    log.info(f"Созданы PDF файлы: {test_pdf_path}, {answers_pdf_path}")
    return test_pdf_path, answers_pdf_path

def create_excel_answer_key(variants: List[Dict[str, Any]], output_dir: str, input_file_name: str = "") -> str:
    """
    Создает Excel файл-ключ с ответами для всех вариантов.
    
    Args:
        variants: Список вариантов тестов
        output_dir: Папка для сохранения файла
        
    Returns:
        Путь к созданному Excel файлу
    """
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
    if input_file_name:
        excel_path = os.path.join(output_dir, f"{input_file_name}_ключ_{timestamp}.xlsx")
    else:
        excel_path = os.path.join(output_dir, f"answer_key_{timestamp}.xlsx")
    
    # Подготавливаем данные для Excel
    data = []
    for variant in variants:
        # Создаем строку с ответами через запятую
        answers_str = ",".join(map(str, variant['answer_key']))
        # Создаем строку с весами через запятую
        weights_str = ",".join(str(q['weight']) for q in variant['questions'])
        data.append({
            'Вариант': variant['variant_number'],
            'Ответы': answers_str,
            'Веса': weights_str
        })
    
    # Создаем DataFrame и сохраняем в Excel
    df = pd.DataFrame(data)
    df.to_excel(excel_path, index=False)
    
    log.info(f"Создан Excel файл-ключ: {excel_path}")
    return excel_path

def check_student_answers(answer_key_file: str, variant_number: int, student_answers: List) -> Dict[str, Any]:
    """
    Проверяет ответы ученика по файлу-ключу.
    
    Args:
        answer_key_file: Путь к Excel файлу-ключу
        variant_number: Номер варианта ученика
        student_answers: Список ответов ученика (может содержать числа и строки)
        
    Returns:
        Словарь с результатами проверки
    """
    try:
        # Читаем файл-ключ с ответами
        key_df = pd.read_excel(answer_key_file)
        
        # Находим строку с нужным вариантом
        variant_row = key_df[key_df['Вариант'] == variant_number]
        if variant_row.empty:
            raise ValueError(f"Вариант {variant_number} не найден в файле-ключе")
        
        # Извлекаем ответы и веса
        answers_str = variant_row['Ответы'].iloc[0]
        weights_str = variant_row['Веса'].iloc[0]
        
        # Парсим ответы и веса
        answer_key = []
        weights = []
        
        for ans in str(answers_str).split(','):
            ans = ans.strip()
            try:
                answer_key.append(int(ans))
            except ValueError:
                answer_key.append(ans)
        
        for weight in str(weights_str).split(','):
            weights.append(float(weight.strip()))
        
        # Проверяем количество ответов
        if len(student_answers) != len(answer_key):
            raise ValueError(f"Количество ответов ученика ({len(student_answers)}) не совпадает с количеством вопросов ({len(answer_key)})")
        
        # Подсчитываем правильные ответы с учетом весов
        total_weight = sum(weights)
        total_points = 12  # Общее количество баллов за тест
        correct_weighted_score = 0
        detailed_results = []
        
        for i, (student_ans, correct_ans, weight) in enumerate(zip(student_answers, answer_key, weights)):
            question_points = (weight / total_weight) * total_points
            
            # Проверяем правильность ответа
            # Определяем тип вопроса по типу правильного ответа
            is_test_question = isinstance(correct_ans, int)
            
            if is_test_question:
                # Тестовое задание - сравниваем числа
                try:
                    student_ans_int = int(student_ans)
                    is_correct = student_ans_int == correct_ans
                except (ValueError, TypeError):
                    is_correct = False
                    student_ans_int = student_ans
            else:
                # Открытое задание - сравниваем строки с нормализацией
                student_str = str(student_ans).strip().lower()
                correct_str = str(correct_ans).strip().lower()
                is_correct = student_str == correct_str
                student_ans_int = student_ans
            
            if is_correct:
                correct_weighted_score += question_points
            
            detailed_results.append({
                'question_number': i + 1,
                'student_answer': student_ans_int,
                'correct_answer': correct_ans,
                'is_correct': is_correct,
                'weight': weight,
                'points': question_points if is_correct else 0,
                'max_points': question_points,
                'is_test_question': is_test_question
            })
        
        # Вычисляем процент
        score_percentage = (correct_weighted_score / total_points) * 100
        correct_count = sum(1 for r in detailed_results if r['is_correct'])
        
        result = {
            'variant_number': variant_number,
            'total_questions': len(answer_key),
            'correct_answers': correct_count,
            'score_percentage': score_percentage,
            'weighted_score': correct_weighted_score,
            'max_score': total_points,
            'detailed_results': detailed_results
        }
        
        log.info(f"Проверка завершена для варианта {variant_number}: {correct_count}/{len(answer_key)} ({score_percentage:.1f}%, {correct_weighted_score:.2f}/{total_points} баллов)")
        return result
        
    except Exception as e:
        log.error(f"Ошибка при проверке ответов: {e}")
        raise

def create_check_result_pdf(check_result: Dict[str, Any], output_dir: str) -> str:
    """
    Создает PDF файл с результатами проверки.
    
    Args:
        check_result: Результаты проверки
        output_dir: Папка для сохранения файла
        
    Returns:
        Путь к созданному PDF файлу
    """
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
    
    # Нормализуем путь и создаем папку
    try:
        output_dir = os.path.normpath(output_dir)
        os.makedirs(output_dir, exist_ok=True)
    except Exception as e:
        log.warning(f"Ошибка при работе с папкой {output_dir}: {e}. Используем временную папку.")
        output_dir = ensure_temp_dir("reports_")
    
    # Формируем имя файла в формате Класс_ПІБ_Вариант_Дата
    student_info = check_result.get('student_info', {})
    class_name = student_info.get('class', '').replace(' ', '_').replace('-', '_') or 'БезКласу'
    full_name = student_info.get('full_name', '').replace(' ', '_') or 'БезІмені'
    variant = check_result['variant_number']
    
    filename = f"{class_name}_{full_name}_Варіант{variant}_{timestamp}.pdf"
    pdf_path = os.path.join(output_dir, filename)
    
    pdf = FPDF()
    pdf.add_font('Arial', '', 'c:/windows/fonts/arial.ttf', uni=True)
    pdf.add_font('Arial', 'B', 'c:/windows/fonts/arialbd.ttf', uni=True)
    
    check_page_width = pdf.w - 2 * pdf.l_margin
    
    pdf.add_page()
    pdf.set_font('Arial', 'B', 16)
    pdf.cell(0, 10, "Результат перевірки тесту", ln=True, align='C')
    pdf.ln(10)
    
    # Данные ученика (если есть)
    student_info = check_result.get('student_info', {})
    if any(student_info.values()):
        pdf.set_font('Arial', 'B', 12)
        add_multiline_text(pdf, "Дані учня / учениці:", check_page_width, 8, 10)
        pdf.set_font('Arial', '', 10)
        
        if student_info.get('class'):
            add_multiline_text(pdf, f"Клас: {student_info['class']}", check_page_width, 6, 10)
        if student_info.get('full_name'):
            add_multiline_text(pdf, f"ПІБ: {student_info['full_name']}", check_page_width, 6, 10)
        pdf.ln(5)
    
    # Основная информация
    pdf.set_font('Arial', 'B', 12)
    # Используем взвешенные баллы
    weighted_score = check_result.get('weighted_score', 0)
    max_score = check_result.get('max_score', 12)
    info_texts = [
        f"Варіант: {check_result['variant_number']}",
        f"Всього питань: {check_result['total_questions']}",
        f"Правильних відповідей: {check_result['correct_answers']}",
        f"Відсоток: {check_result['score_percentage']:.1f}%",
        f"Бали: {weighted_score:.2f} з {max_score}"
    ]
    
    for info_text in info_texts:
        add_multiline_text(pdf, info_text, check_page_width, 8, 10)
    pdf.ln(10)
    
    # Детальные результаты
    pdf.set_font('Arial', 'B', 12)
    add_multiline_text(pdf, "Детальні результати:", check_page_width, 8, 10)
    pdf.ln(5)
    
    # Заголовки таблицы
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(20, 8, "Питання", 1, 0, 'C')
    pdf.cell(30, 8, "Відповідь учня", 1, 0, 'C')
    pdf.cell(45, 8, "Правильна відповідь", 1, 0, 'C')
    pdf.cell(30, 8, "Бали", 1, 0, 'C')
    pdf.cell(35, 8, "Результат", 1, 0, 'C')
    pdf.ln()
    
    # Строки таблицы
    pdf.set_font('Arial', '', 10)
    for result in check_result['detailed_results']:
        pdf.cell(20, 8, str(result['question_number']), 1, 0, 'C')
        pdf.cell(30, 8, str(result['student_answer']), 1, 0, 'C')
        pdf.cell(45, 8, str(result['correct_answer']), 1, 0, 'C')
        # Баллы за задание с учетом весов
        earned_points = result.get('points', 0)
        max_points = result.get('max_points', 0)
        points_text = f"{earned_points:.2f}/{max_points:.2f}"
        pdf.cell(30, 8, points_text, 1, 0, 'C')
        # Используем текст вместо символов, которые не поддерживаются шрифтом Arial
        result_text = "Правильно" if result['is_correct'] else "Неправильно"
        pdf.cell(35, 8, result_text, 1, 0, 'C')
        pdf.ln()
    
    try:
        pdf.output(pdf_path)
        log.info(f"Создан PDF с результатами проверки: {pdf_path}")
        return pdf_path
    except Exception as e:
        # Если ошибка связана с путем, пробуем использовать временную папку
        if 'Invalid argument' in str(e) or 'path' in str(e).lower():
            try:
                output_dir = ensure_temp_dir("reports_")
                log.warning(f"Ошибка с путем PDF, используем временную папку: {output_dir}")
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                pdf_path = os.path.join(output_dir, f"check_result_variant_{check_result['variant_number']}_{timestamp}.pdf")
                pdf.output(pdf_path)
                log.info(f"Создан PDF с результатами проверки (fallback): {pdf_path}")
                return pdf_path
            except Exception as fallback_error:
                log.error(f"Ошибка при создании PDF отчета (fallback): {fallback_error}")
                raise
        else:
            log.error(f"Ошибка при создании PDF отчета: {e}")
            raise

def create_check_result_word(check_result: Dict[str, Any], output_dir: str) -> str:
    """
    Создает Word файл с результатами проверки.
    
    Args:
        check_result: Результаты проверки
        output_dir: Папка для сохранения файла
        
    Returns:
        Путь к созданному Word файлу
    """
    try:
        # Нормализуем путь и создаем папку
        output_dir = os.path.normpath(output_dir)
        os.makedirs(output_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
        
        # Формируем имя файла в формате Класс_ПІБ_Вариант_Дата
        student_info = check_result.get('student_info', {})
        class_name = student_info.get('class', '').replace(' ', '_').replace('-', '_') or 'БезКласу'
        full_name = student_info.get('full_name', '').replace(' ', '_') or 'БезІмені'
        variant = check_result['variant_number']
        
        filename = f"{class_name}_{full_name}_Варіант{variant}_{timestamp}.docx"
        word_path = os.path.join(output_dir, filename)
        
        doc = Document()
        
        # Устанавливаем минимальные поля для экономии места
        section = doc.sections[0]
        section.top_margin = Inches(0.5)     # 1.27 см
        section.bottom_margin = Inches(0.5)  # 1.27 см
        section.left_margin = Inches(0.5)    # 1.27 см
        section.right_margin = Inches(0.5)   # 1.27 см
        
        # Заголовок
        heading = doc.add_heading('Результат перевірки тесту', level=1)
        heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        doc.add_paragraph()  # Пустая строка
        
        # Данные ученика (если есть)
        student_info = check_result.get('student_info', {})
        if any(student_info.values()):
            student_para = doc.add_paragraph()
            student_para.add_run('Дані учня / учениці:').bold = True
            
            if student_info.get('class'):
                doc.add_paragraph(f"Клас: {student_info['class']}", style='List Bullet')
            if student_info.get('full_name'):
                doc.add_paragraph(f"ПІБ: {student_info['full_name']}", style='List Bullet')
            
            doc.add_paragraph()  # Пустая строка
        
        # Основная информация
        info_para = doc.add_paragraph()
        info_para.add_run('Основна інформація:').bold = True
        
        # Используем взвешенные баллы
        weighted_score = check_result.get('weighted_score', 0)
        max_score = check_result.get('max_score', 12)
        info_texts = [
            f"Варіант: {check_result['variant_number']}",
            f"Всього питань: {check_result['total_questions']}",
            f"Правильних відповідей: {check_result['correct_answers']}",
            f"Відсоток: {check_result['score_percentage']:.1f}%",
            f"Бали: {weighted_score:.2f} з {max_score}"
        ]
        
        for info_text in info_texts:
            doc.add_paragraph(info_text, style='List Bullet')
        
        doc.add_paragraph()  # Пустая строка
        
        # Детальные результаты
        details_para = doc.add_paragraph()
        details_para.add_run('Детальні результати:').bold = True
        
        # Создаем таблицу
        table = doc.add_table(rows=1, cols=5)
        table.style = 'Table Grid'
        
        # Настраиваем ширину колонок
        table.columns[0].width = Inches(1.0)  # Питання
        table.columns[1].width = Inches(1.5)  # Відповідь учня
        table.columns[2].width = Inches(1.5)  # Правильна відповідь
        table.columns[3].width = Inches(0.8)  # Бали
        table.columns[4].width = Inches(1.8)  # Результат (увеличена для предотвращения переносов)
        
        # Заголовки таблицы
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = 'Питання'
        hdr_cells[1].text = 'Відповідь учня'
        hdr_cells[2].text = 'Правильна відповідь'
        hdr_cells[3].text = 'Бали'
        hdr_cells[4].text = 'Результат'
        
        # Делаем заголовки жирными
        for cell in hdr_cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.bold = True
        
        # Добавляем строки с результатами
        for result in check_result['detailed_results']:
            row_cells = table.add_row().cells
            row_cells[0].text = str(result['question_number'])
            row_cells[1].text = str(result['student_answer'])
            row_cells[2].text = str(result['correct_answer'])
            # Баллы за задание с учетом веса
            earned_points = result.get('earned_points', 0)
            max_points = result.get('max_points', 0)
            points = f"{earned_points:.2f} / {max_points:.2f}"
            row_cells[3].text = points
            
            # Результат с цветными символами
            result_paragraph = row_cells[4].paragraphs[0]
            result_paragraph.clear()
            if result['is_correct']:
                run = result_paragraph.add_run("✓ Правильно")
                run.font.color.rgb = RGBColor(0, 128, 0)  # Зеленый цвет
            else:
                run = result_paragraph.add_run("✗ Неправильно")
                run.font.color.rgb = RGBColor(255, 0, 0)  # Красный цвет
        
        # Добавляем разрыв страницы сразу после таблицы результатов
        doc.add_page_break()
        
        doc.save(word_path)
        log.info(f"Создан Word документ с результатами проверки: {word_path}")
        return word_path
        
    except Exception as e:
        # Если ошибка связана с путем, пробуем использовать временную папку
        if 'Invalid argument' in str(e) or 'path' in str(e).lower():
            try:
                output_dir = ensure_temp_dir("reports_")
                log.warning(f"Ошибка с путем, используем временную папку: {output_dir}")
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                word_path = os.path.join(output_dir, f"check_result_variant_{check_result['variant_number']}_{timestamp}.docx")
                
                doc = Document()
                
                # Заголовок
                heading = doc.add_heading('Результат перевірки тесту', level=1)
                heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
                
                doc.add_paragraph()  # Пустая строка
                
                # Основная информация
                info_para = doc.add_paragraph()
                info_para.add_run('Основна інформація:').bold = True
                
                # Расчет балла в 12-балльной системе
                score_12 = (check_result['correct_answers'] / check_result['total_questions']) * 12
                
                info_texts = [
                    f"Варіант: {check_result['variant_number']}",
                    f"Всього питань: {check_result['total_questions']}",
                    f"Правильних відповідей: {check_result['correct_answers']}",
                    f"Відсоток: {check_result['score_percentage']:.1f}%",
                    f"Оцінка (12-бальна система): {score_12:.2f}"
                ]
                
                for info_text in info_texts:
                    doc.add_paragraph(info_text, style='List Bullet')
                
                doc.add_paragraph()  # Пустая строка
                
                # Детальные результаты
                details_para = doc.add_paragraph()
                details_para.add_run('Детальні результати:').bold = True
                
                # Создаем таблицу
                table = doc.add_table(rows=1, cols=5)
                table.style = 'Table Grid'
                
                # Настраиваем ширину колонок
                table.columns[0].width = Inches(1.0)  # Питання
                table.columns[1].width = Inches(1.5)  # Відповідь учня
                table.columns[2].width = Inches(1.5)  # Правильна відповідь
                table.columns[3].width = Inches(0.8)  # Бали
                table.columns[4].width = Inches(1.8)  # Результат (увеличена для предотвращения переносов)
                
                # Заголовки таблицы
                hdr_cells = table.rows[0].cells
                hdr_cells[0].text = 'Питання'
                hdr_cells[1].text = 'Відповідь учня'
                hdr_cells[2].text = 'Правильна відповідь'
                hdr_cells[3].text = 'Бали'
                hdr_cells[4].text = 'Результат'
                
                # Делаем заголовки жирными
                for cell in hdr_cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.bold = True
                
                # Добавляем строки с результатами
                points_per_question = 12 / check_result['total_questions']
                for result in check_result['detailed_results']:
                    row_cells = table.add_row().cells
                    row_cells[0].text = str(result['question_number'])
                    row_cells[1].text = str(result['student_answer'])
                    row_cells[2].text = str(result['correct_answer'])
                    # Баллы за задание
                    points = f"{points_per_question:.2f}" if result['is_correct'] else "0.00"
                    row_cells[3].text = points
                    
                    # Результат с цветными символами
                    result_paragraph = row_cells[4].paragraphs[0]
                    result_paragraph.clear()
                    if result['is_correct']:
                        run = result_paragraph.add_run("✓ Правильно")
                        run.font.color.rgb = RGBColor(0, 128, 0)  # Зеленый цвет
                    else:
                        run = result_paragraph.add_run("✗ Неправильно")
                        run.font.color.rgb = RGBColor(255, 0, 0)  # Красный цвет
                
                # Добавляем разрыв страницы сразу после таблицы результатов
                doc.add_page_break()
                
                doc.save(word_path)
                log.info(f"Создан Word документ с результатами проверки (fallback): {word_path}")
                return word_path
                
            except Exception as fallback_error:
                log.error(f"Ошибка при создании Word отчета (fallback): {fallback_error}")
                raise
        else:
            log.error(f"Ошибка при создании Word отчета: {e}")
            raise


def create_test_word(variants: List[Dict[str, Any]], output_dir: str, columns: int = 1, input_file_name: str = "", answer_format: str = "list", space_optimization: bool = False, test_class: str = "", test_date: str = "") -> str:
    """Создать Word документ с тестами для всех вариантов
    
    Args:
        variants: Список вариантов тестов
        output_dir: Папка для сохранения файлов
        columns: Количество колонок для размещения вопросов (всегда 1)
        input_file_name: Имя входного файла
        answer_format: Формат вариантов ответов ('list' или 'table')
        space_optimization: Минимизировать переводы строк для экономии места
        test_class: Класс для отображения в заголовке (опционально)
        test_date: Дата теста для отображения в заголовке (опционально)
    """
    try:
        os.makedirs(output_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
        if input_file_name:
            word_path = os.path.join(output_dir, f"{input_file_name}_тести_{timestamp}.docx")
        else:
            word_path = os.path.join(output_dir, f"tests_{timestamp}.docx")
        
        doc = Document()
        
        # Устанавливаем минимальные поля для экономии места
        section = doc.sections[0]
        section.top_margin = Inches(0.5)     # 1.27 см
        section.bottom_margin = Inches(0.5)  # 1.27 см
        section.left_margin = Inches(0.5)    # 1.27 см
        section.right_margin = Inches(0.5)   # 1.27 см
        
        for variant in variants:
            # Заголовок варианта - общий по центру перед колонками
            title_parts = ["Тест"]
            if test_class:
                title_parts.append(f"Клас: {test_class}")
            if test_date:
                title_parts.append(f"Дата: {test_date}")
            title_parts.append(f"Варіант {variant['variant_number']}")
            
            heading = doc.add_heading(" - ".join(title_parts), level=1)
            heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Инструкция
            instruction = doc.add_paragraph("Інструкція: Оберіть правильну відповідь для кожного питання.")
            instruction.alignment = WD_ALIGN_PARAGRAPH.LEFT
            
            if not space_optimization:
                doc.add_paragraph()  # Пустая строка
            
            # Вычисляем общий вес всех заданий для нормализации баллов
            total_weight = sum(q['weight'] for q in variant['questions'])
            total_points = 12  # Общее количество баллов за тест
            
            # Вопросы в одноколоночной компоновке
            for i, question in enumerate(variant['questions'], 1):
                # Вычисляем баллы за это задание
                question_points = (question['weight'] / total_weight) * total_points
                # Форматируем баллы красиво
                if question_points == int(question_points):
                    points_str = f"({int(question_points)} балів)"
                else:
                    points_str = f"({question_points:.1f} балів)"
                
                # Номер задания и баллы отдельно от текста вопроса
                question_header = doc.add_paragraph(f"{i}. {points_str}")
                question_header.runs[0].bold = True
                
                # Текст вопроса отдельной строкой
                question_para = doc.add_paragraph(question['question_text'])
                question_para.style = 'Normal'
                
                # Варианты ответов в зависимости от типа задания
                if question['is_test_question']:
                    # Тестовое задание с вариантами ответов
                    if answer_format == 'table':
                        # Табличный формат - варианты ответов в таблице по ширине страницы
                        options = question['options']
                        num_options = len(options)
                        
                        # Создаем таблицу с одной строкой и количеством колонок равным количеству вариантов
                        table = doc.add_table(rows=1, cols=num_options)
                        table.style = 'Table Grid'
                        
                        # Растягиваем таблицу по всей ширине страницы
                        table.autofit = False
                        for col_idx, col in enumerate(table.columns):
                            col.width = Inches(6.5 / num_options)  # Равномерно распределяем по ширине
                        
                        # Заполняем ячейки вариантами ответов
                        cells = table.rows[0].cells
                        for j, option in enumerate(options):
                            cells[j].text = f"{j + 1}) {option}"
                            cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    else:
                        # Списочный формат - обычные варианты ответов
                        for j, option in enumerate(question['options'], 1):
                            option_para = doc.add_paragraph(f"   {j}) {option}")
                            option_para.style = 'Normal'
                else:
                    # Нетестовое задание - место для ответа
                    answer_para = doc.add_paragraph("Відповідь: ___________________________")
                    answer_para.style = 'Normal'
                
                if not space_optimization:
                    doc.add_paragraph()  # Пустая строка между вопросами
            
            # Таблица для ответов - сразу после теста, без разрыва страницы
            # Всегда по 15 элементов в строке, последняя строка дополняется пустыми ячейками
            doc.add_paragraph("Таблиця відповідей:")
            
            total_questions = len(variant['questions'])
            questions_per_row = 15  # Всегда 15 элементов в строке
            num_rows = (total_questions + questions_per_row - 1) // questions_per_row  # Округление вверх
            
            current_q = 0
            for row_idx in range(num_rows):
                # Всегда создаем строку с 15 колонками
                cols_in_row = questions_per_row
                
                # Создаем таблицу для текущей строки
                table = doc.add_table(rows=2, cols=cols_in_row)
                table.style = 'Table Grid'
                
                # Растягиваем таблицу по всей странице
                table.autofit = False
                for col in table.columns:
                    col.width = Inches(6.5 / cols_in_row)
                
                # Заголовки (номера вопросов)
                header_cells = table.rows[0].cells
                for i in range(cols_in_row):
                    if current_q < total_questions:
                        header_cells[i].text = str(current_q + 1)
                        current_q += 1
                        header_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                    else:
                        # Пустая ячейка без рамок
                        header_cells[i].text = ""
                        # Убираем рамки для пустых ячеек
                        from docx.oxml.shared import qn
                        tc = header_cells[i]._tc
                        tcPr = tc.get_or_add_tcPr()
                        tcBorders = tcPr.find(qn('w:tcBorders'))
                        if tcBorders is None:
                            tcBorders = parse_xml('<w:tcBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>')
                            tcPr.append(tcBorders)
                
                # Пустая строка для ответов
                answer_cells = table.rows[1].cells
                for i, cell in enumerate(answer_cells):
                    if i < total_questions - (row_idx * questions_per_row):
                        cell.text = ""
                    else:
                        # Пустая ячейка без рамок
                        cell.text = ""
                        from docx.oxml.shared import qn
                        tc = cell._tc
                        tcPr = tc.get_or_add_tcPr()
                        tcBorders = tcPr.find(qn('w:tcBorders'))
                        if tcBorders is None:
                            tcBorders = parse_xml('<w:tcBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>')
                            tcPr.append(tcBorders)
                
                # Добавляем небольшой отступ между строками таблиц
                if row_idx < num_rows - 1 and not space_optimization:
                    doc.add_paragraph()
            
            # Разрыв страницы между вариантами (кроме последнего)
            if variant != variants[-1]:
                doc.add_page_break()
        
        doc.save(word_path)
        log.info(f"Word документ с тестами создан: {word_path}")
        return word_path
        
    except Exception as e:
        log.error(f"Ошибка при создании Word документа: {e}")
        raise


def read_test_word(file_path: str) -> pd.DataFrame:
    """Прочитать тест из Word документа и преобразовать в DataFrame"""
    try:
        doc = Document(file_path)
        questions_data = []
        
        current_question = None
        current_options = []
        question_number = 0
        
        for para in doc.paragraphs:
            text = para.text.strip()
            
            if not text:
                continue
            
            # Проверяем, является ли это вопросом (начинается с числа и точки)
            if text and text[0].isdigit() and '. ' in text:
                # Сохраняем предыдущий вопрос
                if current_question and current_options:
                    questions_data.append({
                        'question': current_question,
                        'option_1': current_options[0] if len(current_options) > 0 else '',
                        'option_2': current_options[1] if len(current_options) > 1 else '',
                        'option_3': current_options[2] if len(current_options) > 2 else '',
                        'option_4': current_options[3] if len(current_options) > 3 else '',
                        'correct_answer': 1  # По умолчанию первый вариант
                    })
                
                # Начинаем новый вопрос
                question_number += 1
                current_question = text.split('. ', 1)[1] if '. ' in text else text
                current_options = []
            
            # Проверяем, является ли это вариантом ответа
            elif text and (text.startswith('   1)') or text.startswith('   2)') or 
                          text.startswith('   3)') or text.startswith('   4)')):
                option_text = text.split(') ', 1)[1] if ') ' in text else text
                current_options.append(option_text)
        
        # Сохраняем последний вопрос
        if current_question and current_options:
            questions_data.append({
                'question': current_question,
                'option_1': current_options[0] if len(current_options) > 0 else '',
                'option_2': current_options[1] if len(current_options) > 1 else '',
                'option_3': current_options[2] if len(current_options) > 2 else '',
                'option_4': current_options[3] if len(current_options) > 3 else '',
                'correct_answer': 1  # По умолчанию первый вариант
            })
        
        if not questions_data:
            raise ValueError("Не удалось найти вопросы в Word документе")
        
        df = pd.DataFrame(questions_data)
        log.info(f"Из Word документа загружено {len(df)} вопросов")
        return df
        
    except Exception as e:
        log.error(f"Ошибка при чтении Word документа: {e}")
        raise


def export_answers_to_word(variants: List[Dict[str, Any]], output_dir: str, input_file_name: str = "") -> str:
    """Экспортировать ответы всех вариантов в Word документ"""
    try:
        os.makedirs(output_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
        if input_file_name:
            word_path = os.path.join(output_dir, f"{input_file_name}_ключ_{timestamp}.docx")
        else:
            word_path = os.path.join(output_dir, f"answers_{timestamp}.docx")
        
        doc = Document()
        
        # Устанавливаем минимальные поля для экономии места
        section = doc.sections[0]
        section.top_margin = Inches(0.5)     # 1.27 см
        section.bottom_margin = Inches(0.5)  # 1.27 см
        section.left_margin = Inches(0.5)    # 1.27 см
        section.right_margin = Inches(0.5)   # 1.27 см
        
        # Заголовок
        heading = doc.add_heading('Відповіді до тестів', level=1)
        heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Инструкция
        instruction = doc.add_paragraph(
            "Цей документ містить правильні відповіді для всіх варіантів тестів. "
            "Використовуйте його для перевірки робіт учнів."
        )
        instruction.alignment = WD_ALIGN_PARAGRAPH.LEFT
        
        doc.add_paragraph()  # Пустая строка
        
        # Ответы для каждого варианта
        for variant in variants:
            # Заголовок варианта
            variant_heading = doc.add_heading(f'Варіант {variant["variant_number"]}', level=2)
            variant_heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
            
            # Ответы в виде таблицы
            table = doc.add_table(rows=1, cols=2)
            table.style = 'Table Grid'
            
            # Заголовки таблицы
            hdr_cells = table.rows[0].cells
            hdr_cells[0].text = 'Номер питання'
            hdr_cells[1].text = 'Правильна відповідь'
            
            # Делаем заголовки жирными
            for cell in hdr_cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.bold = True
            
            # Добавляем строки с ответами
            for i, answer in enumerate(variant['answer_key'], 1):
                row_cells = table.add_row().cells
                row_cells[0].text = str(i)
                row_cells[1].text = str(answer)
            
            # Также добавляем ответы в строку для удобства
            answers_line = doc.add_paragraph()
            answers_line.add_run('Відповіді в рядок: ').bold = True
            answers_line.add_run(', '.join([f"{i+1}-{ans}" for i, ans in enumerate(variant['answer_key'])]))
            
            doc.add_paragraph()  # Пустая строка между вариантами
        
        doc.save(word_path)
        log.info(f"Word документ с ответами создан: {word_path}")
        return word_path
        
    except Exception as e:
        log.error(f"Ошибка при экспорте ответов в Word: {e}")
        raise
