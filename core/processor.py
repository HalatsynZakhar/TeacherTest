import pandas as pd
import random
import os
from typing import List, Dict, Any, Tuple, Optional
from fpdf import FPDF
import logging
from datetime import datetime
import textwrap
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml

log = logging.getLogger(__name__)

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
    - Столбец 1: Номер правильного ответа
    - Столбцы 2+: Варианты ответов
    
    Args:
        file_path: Путь к Excel файлу
        
    Returns:
        DataFrame с вопросами теста
    """
    try:
        # Читаем Excel файл
        df = pd.read_excel(file_path, header=None)
        
        # Конвертируем все значения в строки для предотвращения ошибок с числовыми ячейками
        df = df.astype(str)
        
        # Проверяем минимальную структуру (вопрос + правильный ответ + минимум 2 варианта)
        if df.shape[1] < 4:
            raise ValueError("Файл должен содержать минимум 4 столбца: вопрос, правильный ответ, и минимум 2 варианта ответов")
        
        # Удаляем пустые строки (после конвертации в строки проверяем на 'nan')
        df = df[(df.iloc[:, 0] != 'nan') & (df.iloc[:, 1] != 'nan')]  # Удаляем строки где нет вопроса или правильного ответа
        
        if df.empty:
            raise ValueError("Файл не содержит валидных данных")
        
        # Переименовываем столбцы для удобства
        columns = ['question', 'correct_answer'] + [f'option_{i}' for i in range(1, df.shape[1] - 1)]
        df.columns = columns
        
        # Проверяем, что правильные ответы - числа
        try:
            df['correct_answer'] = pd.to_numeric(df['correct_answer'], errors='coerce')
            df = df.dropna(subset=['correct_answer'])
        except:
            raise ValueError("Столбец с правильными ответами должен содержать числа")
        
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
            # Собираем все варианты ответов
            options = []
            for col in df.columns:
                if col.startswith('option_') and pd.notna(row[col]):
                    options.append(str(row[col]))
            
            if len(options) < 2:
                log.warning(f"Вопрос '{row['question']}' имеет менее 2 вариантов ответов, пропускаем")
                continue
            
            # Перемешиваем варианты ответов
            random.seed(variant_num + idx)  # Для воспроизводимости
            shuffled_options = options.copy()
            random.shuffle(shuffled_options)
            
            # Находим новую позицию правильного ответа
            correct_option_text = options[int(row['correct_answer']) - 1]  # -1 так как нумерация с 1
            new_correct_position = shuffled_options.index(correct_option_text) + 1  # +1 для нумерации с 1
            
            question_data = {
                'question_text': str(row['question']),
                'options': shuffled_options,
                'correct_answer': new_correct_position
            }
            
            variant['questions'].append(question_data)
            variant['answer_key'].append(new_correct_position)
        
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
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
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

def create_excel_answer_key(variants: List[Dict[str, Any]], output_dir: str) -> str:
    """
    Создает Excel файл-ключ с ответами для всех вариантов.
    
    Args:
        variants: Список вариантов тестов
        output_dir: Папка для сохранения файла
        
    Returns:
        Путь к созданному Excel файлу
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_path = os.path.join(output_dir, f"answer_key_{timestamp}.xlsx")
    
    # Подготавливаем данные для Excel
    data = []
    for variant in variants:
        # Создаем строку с ответами через запятую
        answers_str = ",".join(map(str, variant['answer_key']))
        data.append({
            'Вариант': variant['variant_number'],
            'Ответы': answers_str
        })
    
    # Создаем DataFrame и сохраняем в Excel
    df = pd.DataFrame(data)
    df.to_excel(excel_path, index=False)
    
    log.info(f"Создан Excel файл-ключ: {excel_path}")
    return excel_path

def check_student_answers(answer_key_file: str, variant_number: int, student_answers: List[int]) -> Dict[str, Any]:
    """
    Проверяет ответы ученика по файлу-ключу.
    
    Args:
        answer_key_file: Путь к Excel файлу-ключу
        variant_number: Номер варианта ученика
        student_answers: Список ответов ученика
        
    Returns:
        Словарь с результатами проверки
    """
    try:
        # Читаем файл-ключ
        df = pd.read_excel(answer_key_file)
        
        # Находим строку с нужным вариантом
        variant_row = df[df['Вариант'] == variant_number]
        if variant_row.empty:
            raise ValueError(f"Вариант {variant_number} не найден в файле-ключе")
        
        # Получаем правильные ответы
        correct_answers_str = variant_row['Ответы'].iloc[0]
        correct_answers = [int(x.strip()) for x in correct_answers_str.split(',')]
        
        # Проверяем количество ответов
        if len(student_answers) != len(correct_answers):
            raise ValueError(f"Количество ответов ученика ({len(student_answers)}) не совпадает с количеством вопросов ({len(correct_answers)})")
        
        # Подсчитываем правильные ответы
        correct_count = 0
        detailed_results = []
        
        for i, (student_ans, correct_ans) in enumerate(zip(student_answers, correct_answers)):
            is_correct = student_ans == correct_ans
            if is_correct:
                correct_count += 1
            
            detailed_results.append({
                'question_number': i + 1,
                'student_answer': student_ans,
                'correct_answer': correct_ans,
                'is_correct': is_correct
            })
        
        # Вычисляем процент
        score_percentage = (correct_count / len(correct_answers)) * 100
        
        result = {
            'variant_number': variant_number,
            'total_questions': len(correct_answers),
            'correct_answers': correct_count,
            'score_percentage': score_percentage,
            'detailed_results': detailed_results
        }
        
        log.info(f"Проверка завершена для варианта {variant_number}: {correct_count}/{len(correct_answers)} ({score_percentage:.1f}%)")
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
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # Нормализуем путь и проверяем доступность
    try:
        output_dir = os.path.normpath(output_dir)
        # Если это сетевой путь и он недоступен, используем локальную папку
        if output_dir.startswith('\\\\') and not os.path.exists(output_dir):
            output_dir = os.path.expanduser('~/Desktop')
            log.warning(f"Сетевой путь недоступен, используем локальную папку: {output_dir}")
        
        # Создаем папку если она не существует
        os.makedirs(output_dir, exist_ok=True)
        
    except Exception as e:
        log.warning(f"Ошибка при работе с папкой {output_dir}: {e}. Используем текущую папку.")
        output_dir = os.getcwd()
    
    pdf_path = os.path.join(output_dir, f"check_result_variant_{check_result['variant_number']}_{timestamp}.pdf")
    
    pdf = FPDF()
    pdf.add_font('Arial', '', 'c:/windows/fonts/arial.ttf', uni=True)
    pdf.add_font('Arial', 'B', 'c:/windows/fonts/arialbd.ttf', uni=True)
    
    check_page_width = pdf.w - 2 * pdf.l_margin
    
    pdf.add_page()
    pdf.set_font('Arial', 'B', 16)
    pdf.cell(0, 10, "Результат перевірки тесту", ln=True, align='C')
    pdf.ln(10)
    
    # Основная информация
    pdf.set_font('Arial', 'B', 12)
    info_texts = [
        f"Варіант: {check_result['variant_number']}",
        f"Всього питань: {check_result['total_questions']}",
        f"Правильних відповідей: {check_result['correct_answers']}",
        f"Відсоток: {check_result['score_percentage']:.1f}%"
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
    pdf.cell(25, 8, "Питання", 1, 0, 'C')
    pdf.cell(35, 8, "Відповідь учня", 1, 0, 'C')
    pdf.cell(35, 8, "Правильна відповідь", 1, 0, 'C')
    pdf.cell(45, 8, "Результат", 1, 0, 'C')
    pdf.ln()
    
    # Строки таблицы
    pdf.set_font('Arial', '', 10)
    for result in check_result['detailed_results']:
        pdf.cell(25, 8, str(result['question_number']), 1, 0, 'C')
        pdf.cell(35, 8, str(result['student_answer']), 1, 0, 'C')
        pdf.cell(35, 8, str(result['correct_answer']), 1, 0, 'C')
        # Используем текст вместо символов, которые не поддерживаются шрифтом Arial
        result_text = "Правильно" if result['is_correct'] else "Неправильно"
        pdf.cell(45, 8, result_text, 1, 0, 'C')
        pdf.ln()
    
    try:
        pdf.output(pdf_path)
        log.info(f"Создан PDF с результатами проверки: {pdf_path}")
        return pdf_path
    except Exception as e:
        # Если ошибка связана с путем, пробуем использовать текущую папку
        if 'Invalid argument' in str(e) or 'path' in str(e).lower():
            try:
                output_dir = os.getcwd()
                log.warning(f"Ошибка с путем PDF, используем текущую папку: {output_dir}")
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
        # Нормализуем путь и проверяем доступность
        output_dir = os.path.normpath(output_dir)
        # Если это сетевой путь и он недоступен, используем локальную папку
        if output_dir.startswith('\\\\') and not os.path.exists(output_dir):
            output_dir = os.path.expanduser('~/Desktop')
            log.warning(f"Сетевой путь недоступен, используем локальную папку: {output_dir}")
        
        os.makedirs(output_dir, exist_ok=True)
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
        
        info_texts = [
            f"Варіант: {check_result['variant_number']}",
            f"Всього питань: {check_result['total_questions']}",
            f"Правильних відповідей: {check_result['correct_answers']}",
            f"Відсоток: {check_result['score_percentage']:.1f}%"
        ]
        
        for info_text in info_texts:
            doc.add_paragraph(info_text, style='List Bullet')
        
        doc.add_paragraph()  # Пустая строка
        
        # Детальные результаты
        details_para = doc.add_paragraph()
        details_para.add_run('Детальні результати:').bold = True
        
        # Создаем таблицу
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        
        # Заголовки таблицы
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = 'Питання'
        hdr_cells[1].text = 'Відповідь учня'
        hdr_cells[2].text = 'Правильна відповідь'
        hdr_cells[3].text = 'Результат'
        
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
            row_cells[3].text = "✓" if result['is_correct'] else "✗"
        
        doc.save(word_path)
        log.info(f"Создан Word документ с результатами проверки: {word_path}")
        return word_path
        
    except Exception as e:
        # Если ошибка связана с путем, пробуем использовать текущую папку
        if 'Invalid argument' in str(e) or 'path' in str(e).lower():
            try:
                output_dir = os.getcwd()
                log.warning(f"Ошибка с путем, используем текущую папку: {output_dir}")
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
                
                info_texts = [
                    f"Варіант: {check_result['variant_number']}",
                    f"Всього питань: {check_result['total_questions']}",
                    f"Правильних відповідей: {check_result['correct_answers']}",
                    f"Відсоток: {check_result['score_percentage']:.1f}%"
                ]
                
                for info_text in info_texts:
                    doc.add_paragraph(info_text, style='List Bullet')
                
                doc.add_paragraph()  # Пустая строка
                
                # Детальные результаты
                details_para = doc.add_paragraph()
                details_para.add_run('Детальні результати:').bold = True
                
                # Создаем таблицу
                table = doc.add_table(rows=1, cols=4)
                table.style = 'Table Grid'
                
                # Заголовки таблицы
                hdr_cells = table.rows[0].cells
                hdr_cells[0].text = 'Питання'
                hdr_cells[1].text = 'Відповідь учня'
                hdr_cells[2].text = 'Правильна відповідь'
                hdr_cells[3].text = 'Результат'
                
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
                    row_cells[3].text = "✓" if result['is_correct'] else "✗"
                
                doc.save(word_path)
                log.info(f"Создан Word документ с результатами проверки (fallback): {word_path}")
                return word_path
                
            except Exception as fallback_error:
                log.error(f"Ошибка при создании Word отчета (fallback): {fallback_error}")
                raise
        else:
            log.error(f"Ошибка при создании Word отчета: {e}")
            raise


def create_test_word(variants: List[Dict[str, Any]], output_dir: str, columns: int = 1) -> str:
    """Создать Word документ с тестами для всех вариантов
    
    Args:
        variants: Список вариантов тестов
        output_dir: Папка для сохранения файлов
        columns: Количество колонок для размещения вопросов (1-3)
    """
    try:
        os.makedirs(output_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        word_path = os.path.join(output_dir, f"tests_{timestamp}.docx")
        
        doc = Document()
        
        for variant in variants:
            # Заголовок варианта
            heading = doc.add_heading(f"Тест - Варіант {variant['variant_number']}", level=1)
            heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Инструкция
            instruction = doc.add_paragraph("Інструкція: Оберіть правильну відповідь для кожного питання.")
            instruction.alignment = WD_ALIGN_PARAGRAPH.LEFT
            
            doc.add_paragraph()  # Пустая строка
            
            # Вопросы с поддержкой колонок
            num_columns = max(1, min(3, columns))  # Ограничиваем от 1 до 3 колонок
            
            if num_columns == 1:
                # Обычная компоновка в одну колонку
                for i, question in enumerate(variant['questions'], 1):
                    # Текст вопроса
                    question_para = doc.add_paragraph(f"{i}. {question['question_text']}")
                    question_para.runs[0].bold = True
                    
                    # Варианты ответов
                    for j, option in enumerate(question['options'], 1):
                        option_para = doc.add_paragraph(f"   {j}) {option}")
                        # Убираем автоматическую нумерацию стиля, чтобы избежать глобального счетчика
                        option_para.style = 'Normal'
                    
                    doc.add_paragraph()  # Пустая строка между вопросами
            else:
                # Многоколоночная компоновка с использованием таблицы
                questions = variant['questions']
                questions_per_column = len(questions) // num_columns + (1 if len(questions) % num_columns > 0 else 0)
                
                # Создаем таблицу для колонок
                table = doc.add_table(rows=1, cols=num_columns)
                table.style = 'Table Grid'
                
                # Убираем границы таблицы
                for row in table.rows:
                    for cell in row.cells:
                        cell._element.get_or_add_tcPr().append(parse_xml('<w:tcBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>'))
                
                # Заполняем колонки
                for col in range(num_columns):
                    cell = table.cell(0, col)
                    cell_para = cell.paragraphs[0]
                    cell_para.clear()
                    
                    start_q = col * questions_per_column
                    end_q = min(start_q + questions_per_column, len(questions))
                    
                    for q_idx in range(start_q, end_q):
                        q_num = (q_idx - start_q) + 1  # Нумерация относительно начала колонки
                        question = questions[q_idx]
                        
                        # Добавляем вопрос
                        if q_idx > start_q:
                            cell.add_paragraph()
                        
                        question_para = cell.add_paragraph(f"{q_num}. {question['question_text']}")
                        question_para.runs[0].bold = True
                        
                        # Добавляем варианты ответов
                        for j, option in enumerate(question['options'], 1):
                            option_para = cell.add_paragraph(f"   {j}) {option}")
                        
                        if q_idx < end_q - 1:  # Не последний вопрос в колонке
                            cell.add_paragraph()
            
            # Таблица для ответов
            doc.add_paragraph("Таблиця відповідей:")
            table = doc.add_table(rows=1, cols=len(variant['questions']))
            table.style = 'Table Grid'
            
            # Заголовки таблицы
            header_cells = table.rows[0].cells
            for i in range(len(variant['questions'])):
                header_cells[i].text = str(i + 1)
                header_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Пустая строка для ответов
            row_cells = table.add_row().cells
            for cell in row_cells:
                cell.text = ""
            
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


def export_answers_to_word(variants: List[Dict[str, Any]], output_dir: str) -> str:
    """Экспортировать ответы всех вариантов в Word документ"""
    try:
        os.makedirs(output_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        word_path = os.path.join(output_dir, f"answers_{timestamp}.docx")
        
        doc = Document()
        
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
