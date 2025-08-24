#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from core.processor import create_check_result_pdf, create_check_result_word, create_test_word

def test_all_fixes():
    """Тест всех исправлений: баллы, ширина колонок, структура Word документа"""
    
    # Тестовые данные для отчетов
    check_result = {
        'variant_number': 1,
        'total_questions': 10,
        'correct_answers': 6,
        'score_percentage': 60.0,
        'detailed_results': [
            {'question_number': 1, 'student_answer': 'A', 'correct_answer': 'A', 'is_correct': True},
            {'question_number': 2, 'student_answer': 'B', 'correct_answer': 'C', 'is_correct': False},
            {'question_number': 3, 'student_answer': 'C', 'correct_answer': 'C', 'is_correct': True},
            {'question_number': 4, 'student_answer': 'A', 'correct_answer': 'B', 'is_correct': False},
            {'question_number': 5, 'student_answer': 'D', 'correct_answer': 'D', 'is_correct': True},
            {'question_number': 6, 'student_answer': 'A', 'correct_answer': 'A', 'is_correct': True},
            {'question_number': 7, 'student_answer': 'B', 'correct_answer': 'C', 'is_correct': False},
            {'question_number': 8, 'student_answer': 'C', 'correct_answer': 'C', 'is_correct': True},
            {'question_number': 9, 'student_answer': 'A', 'correct_answer': 'B', 'is_correct': False},
            {'question_number': 10, 'student_answer': 'D', 'correct_answer': 'D', 'is_correct': True}
        ]
    }
    
    # Тестовые данные для создания тестов
    test_variants = [{
        'variant_number': 1,
        'questions': [
            {
                'question_text': 'Какой язык программирования используется для веб-разработки?',
                'options': ['Python', 'JavaScript', 'C++', 'Java'],
                'correct_answer': 'B'
            },
            {
                'question_text': 'Что означает HTML?',
                'options': ['HyperText Markup Language', 'High Tech Modern Language', 'Home Tool Markup Language', 'Hyperlink and Text Markup Language'],
                'correct_answer': 'A'
            },
            {
                'question_text': 'Какой протокол используется для передачи веб-страниц?',
                'options': ['FTP', 'SMTP', 'HTTP', 'TCP'],
                'correct_answer': 'C'
            },
            {
                'question_text': 'Что такое CSS?',
                'options': ['Cascading Style Sheets', 'Computer Style Sheets', 'Creative Style Sheets', 'Colorful Style Sheets'],
                'correct_answer': 'A'
            },
            {
                'question_text': 'Какой тег используется для создания ссылки в HTML?',
                'options': ['<link>', '<a>', '<href>', '<url>'],
                'correct_answer': 'B'
            },
            {
                'question_text': 'Что означает SQL?',
                'options': ['Structured Query Language', 'Simple Query Language', 'Standard Query Language', 'System Query Language'],
                'correct_answer': 'A'
            }
        ]
    }]
    
    print("Тестирование всех исправлений...")
    
    # Создаем PDF отчет
    try:
        pdf_path = create_check_result_pdf(check_result, "test_fixes_reports")
        print(f"✅ PDF отчет создан: {pdf_path}")
    except Exception as e:
        print(f"❌ Ошибка создания PDF отчета: {e}")
    
    # Создаем Word отчет
    try:
        word_path = create_check_result_word(check_result, "test_fixes_reports")
        print(f"✅ Word отчет создан: {word_path}")
    except Exception as e:
        print(f"❌ Ошибка создания Word отчета: {e}")
    
    # Создаем Word тест с 2 колонками
    try:
        test_word_path = create_test_word(test_variants, "test_fixes_reports", columns=2)
        print(f"✅ Word тест создан: {test_word_path}")
    except Exception as e:
        print(f"❌ Ошибка создания Word теста: {e}")
    
    print("\nПроверьте созданные файлы:")
    print("1. PDF отчет: колонка 'Бали' с правильными баллами (1.20 за каждый вопрос), 12-балльная оценка (7.20)")
    print("2. Word отчет: цветные символы, колонка 'Бали' с правильными баллами, широкая колонка 'Результат', 12-балльная оценка")
    print("3. Word тест: правильная нумерация вопросов в 2 колонках (1, 2, 3... а не 1, 1, 2...)")
    
    # Проверяем расчет баллов
    points_per_question = 12 / check_result['total_questions']
    expected_score = check_result['correct_answers'] * points_per_question
    print(f"\nОжидаемые баллы за вопрос: {points_per_question:.2f}")
    print(f"Ожидаемая итоговая оценка: {expected_score:.2f}")

if __name__ == "__main__":
    test_all_fixes()