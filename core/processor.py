import pandas as pd
import random
import os
from typing import List, Dict, Any, Tuple, Optional
from fpdf import FPDF
import logging
from datetime import datetime

log = logging.getLogger(__name__)

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
        
        # Проверяем минимальную структуру (вопрос + правильный ответ + минимум 2 варианта)
        if df.shape[1] < 4:
            raise ValueError("Файл должен содержать минимум 4 столбца: вопрос, правильный ответ, и минимум 2 варианта ответов")
        
        # Удаляем пустые строки
        df = df.dropna(subset=[0, 1])  # Удаляем строки где нет вопроса или правильного ответа
        
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

def create_test_pdf(variants: List[Dict[str, Any]], output_dir: str) -> Tuple[str, str]:
    """
    Создает PDF файлы с тестами для учеников и ответами для учителя.
    
    Args:
        variants: Список вариантов тестов
        output_dir: Папка для сохранения файлов
        
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
    
    for variant in variants:
        test_pdf.add_page()
        test_pdf.set_font('Arial', 'B', 16)
        test_pdf.cell(0, 10, f"Тест - Вариант {variant['variant_number']}", ln=True, align='C')
        test_pdf.ln(10)
        
        test_pdf.set_font('Arial', '', 12)
        test_pdf.cell(0, 8, "Инструкция: Выберите правильный ответ и впишите его номер в таблицу внизу.", ln=True)
        test_pdf.ln(5)
        
        # Добавляем вопросы
        for q_num, question in enumerate(variant['questions'], 1):
            test_pdf.set_font('Arial', 'B', 11)
            test_pdf.cell(0, 8, f"{q_num}. {question['question_text']}", ln=True)
            test_pdf.set_font('Arial', '', 10)
            
            for opt_num, option in enumerate(question['options'], 1):
                test_pdf.cell(0, 6, f"   {opt_num}) {option}", ln=True)
            test_pdf.ln(3)
        
        # Добавляем таблицу для ответов
        test_pdf.ln(10)
        test_pdf.set_font('Arial', 'B', 12)
        test_pdf.cell(0, 8, "Таблица ответов:", ln=True)
        test_pdf.ln(5)
        
        # Создаем таблицу ответов
        test_pdf.set_font('Arial', '', 10)
        questions_per_row = 10
        num_questions = len(variant['questions'])
        
        for row_start in range(0, num_questions, questions_per_row):
            # Номера вопросов
            test_pdf.set_font('Arial', 'B', 9)
            for i in range(row_start, min(row_start + questions_per_row, num_questions)):
                test_pdf.cell(18, 8, f"№{i+1}", 1, 0, 'C')
            test_pdf.ln()
            
            # Пустые ячейки для ответов
            test_pdf.set_font('Arial', '', 9)
            for i in range(row_start, min(row_start + questions_per_row, num_questions)):
                test_pdf.cell(18, 12, "", 1, 0, 'C')
            test_pdf.ln()
            test_pdf.ln(3)
    
    test_pdf.output(test_pdf_path)
    
    # Создаем PDF с ответами для учителя
    answers_pdf = FPDF()
    answers_pdf.add_font('Arial', '', 'c:/windows/fonts/arial.ttf', uni=True)
    answers_pdf.add_font('Arial', 'B', 'c:/windows/fonts/arialbd.ttf', uni=True)
    
    answers_pdf.add_page()
    answers_pdf.set_font('Arial', 'B', 16)
    answers_pdf.cell(0, 10, "Ответы для учителя", ln=True, align='C')
    answers_pdf.ln(10)
    
    # Группируем варианты по несколько на страницу
    variants_per_page = 4
    for page_start in range(0, len(variants), variants_per_page):
        if page_start > 0:
            answers_pdf.add_page()
        
        for variant in variants[page_start:page_start + variants_per_page]:
            answers_pdf.set_font('Arial', 'B', 14)
            answers_pdf.cell(0, 8, f"Вариант {variant['variant_number']}", ln=True)
            answers_pdf.set_font('Arial', '', 11)
            
            # Выводим ответы в строку
            answer_text = "Ответы: " + ", ".join([f"{i+1}-{ans}" for i, ans in enumerate(variant['answer_key'])])
            answers_pdf.cell(0, 6, answer_text, ln=True)
            answers_pdf.ln(5)
    
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
    pdf_path = os.path.join(output_dir, f"check_result_variant_{check_result['variant_number']}_{timestamp}.pdf")
    
    pdf = FPDF()
    pdf.add_font('Arial', '', 'c:/windows/fonts/arial.ttf', uni=True)
    pdf.add_font('Arial', 'B', 'c:/windows/fonts/arialbd.ttf', uni=True)
    
    pdf.add_page()
    pdf.set_font('Arial', 'B', 16)
    pdf.cell(0, 10, "Результат проверки теста", ln=True, align='C')
    pdf.ln(10)
    
    # Основная информация
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 8, f"Вариант: {check_result['variant_number']}", ln=True)
    pdf.cell(0, 8, f"Всего вопросов: {check_result['total_questions']}", ln=True)
    pdf.cell(0, 8, f"Правильных ответов: {check_result['correct_answers']}", ln=True)
    pdf.cell(0, 8, f"Процент: {check_result['score_percentage']:.1f}%", ln=True)
    pdf.ln(10)
    
    # Детальные результаты
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 8, "Детальные результаты:", ln=True)
    pdf.ln(5)
    
    # Заголовки таблицы
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(30, 8, "Вопрос", 1, 0, 'C')
    pdf.cell(40, 8, "Ответ ученика", 1, 0, 'C')
    pdf.cell(40, 8, "Правильный ответ", 1, 0, 'C')
    pdf.cell(30, 8, "Результат", 1, 0, 'C')
    pdf.ln()
    
    # Строки таблицы
    pdf.set_font('Arial', '', 10)
    for result in check_result['detailed_results']:
        pdf.cell(30, 8, str(result['question_number']), 1, 0, 'C')
        pdf.cell(40, 8, str(result['student_answer']), 1, 0, 'C')
        pdf.cell(40, 8, str(result['correct_answer']), 1, 0, 'C')
        pdf.cell(30, 8, "✓" if result['is_correct'] else "✗", 1, 0, 'C')
        pdf.ln()
    
    pdf.output(pdf_path)
    
    log.info(f"Создан PDF с результатами проверки: {pdf_path}")
    return pdf_path
