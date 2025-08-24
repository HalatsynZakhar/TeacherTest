#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from core.processor import read_test_excel, generate_test_variants, check_student_answers, create_check_result_pdf, create_check_result_word

def test_comma_formatting():
    """Тестирует форматирование чисел с запятой"""
    
    # Читаем тестовый файл
    df = read_test_excel('test_numbers.xlsx')
    print("Данные из Excel:")
    print(df)
    print()
    
    # Генерируем варианты
    variants = generate_test_variants(df, 2)
    print(f"Сгенерировано {len(variants)} вариантов")
    print()
    
    # Создаем ключ ответов для первого варианта
    variant_1 = variants[0]
    answer_key = []
    for q in variant_1['questions']:
        answer_key.append({
            'correct_answer': q['correct_answer'],
            'weight': q['weight'],
            'is_test_question': q['is_test_question']
        })
    
    # Симулируем ответы студента (некоторые правильные, некоторые неправильные)
    student_answers = [6.0, 1.0, 2.0]  # Первый правильный, второй неправильный, третий неправильный
    
    # Создаем результат проверки вручную
    total_weight = sum(q['weight'] for q in variant_1['questions'])
    total_points = 12
    detailed_results = []
    
    for i, (student_ans, answer_info) in enumerate(zip(student_answers, answer_key)):
        question_points = (answer_info['weight'] / total_weight) * total_points
        is_correct = student_ans == answer_info['correct_answer']
        
        detailed_results.append({
            'question_number': i + 1,
            'student_answer': student_ans,
            'correct_answer': answer_info['correct_answer'],
            'is_correct': is_correct,
            'weight': answer_info['weight'],
            'points': question_points if is_correct else 0,
            'max_points': question_points,
            'is_test_question': answer_info['is_test_question']
        })
    
    # Пересчитываем результаты
    correct_count = sum(1 for r in detailed_results if r['is_correct'])
    weighted_score = sum(r['points'] for r in detailed_results)
    score_percentage = (weighted_score / total_points) * 100
    
    check_result = {
        'variant_number': 1,
        'total_questions': len(student_answers),
        'correct_answers': correct_count,
        'score_percentage': score_percentage,
        'weighted_score': weighted_score,
        'max_score': total_points,
        'detailed_results': detailed_results
    }
    
    print("Результат проверки:")
    print(f"Правильных ответов: {check_result['correct_answers']}/{check_result['total_questions']}")
    print(f"Процент: {check_result['score_percentage']:.1f}%")
    print(f"Баллы: {check_result['weighted_score']:.2f}/{check_result['max_score']}")
    print()
    
    # Создаем PDF отчет
    try:
        pdf_path = create_check_result_pdf(check_result, '.')
        print(f"PDF отчет создан: {pdf_path}")
    except Exception as e:
        print(f"Ошибка создания PDF: {e}")
    
    # Создаем Word отчет
    try:
        word_path = create_check_result_word(check_result, '.')
        print(f"Word отчет создан: {word_path}")
    except Exception as e:
        print(f"Ошибка создания Word: {e}")

if __name__ == '__main__':
    test_comma_formatting()