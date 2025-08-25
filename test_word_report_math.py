#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тест для перевірки обробки математичних формул у детальному Word звіті
"""

import os
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.processor import create_check_result_word

def test_word_report_with_math():
    """Тестуємо створення Word звіту з математичними формулами"""
    
    # Створюємо тестові дані з математичними формулами
    check_result = {
        'variant_number': 1,
        'total_questions': 4,
        'correct_answers': 2,
        'score_percentage': 50.0,
        'weighted_score': 6.0,
        'max_score': 12.0,
        'student_info': {
            'class': '10-А',
            'full_name': 'Іванов Іван Іванович'
        },
        'detailed_results': [
            {
                'question_number': 1,
                'question_text': 'Розв\'яжіть рівняння x^2 + 3x - 4 = 0',
                'question_type': 'Тестове',
                'question_options': ['x = 1, x = -4', 'x = 2, x = -2', 'x = 0, x = 3', 'x = -1, x = 4'],
                'student_answer': '1',
                'correct_answer': '1',
                'is_correct': True,
                'weight': 1.0,
                'points': 3.0,
                'max_points': 3.0,
                'is_test_question': True
            },
            {
                'question_number': 2,
                'question_text': 'Обчисліть значення виразу (a+b)/(c-d) при a=5, b=3, c=10, d=2',
                'question_type': 'Тестове',
                'question_options': ['1', '2', '3', '4'],
                'student_answer': '2',
                'correct_answer': '1',
                'is_correct': False,
                'weight': 1.0,
                'points': 0.0,
                'max_points': 3.0,
                'is_test_question': True
            },
            {
                'question_number': 3,
                'question_text': 'Знайдіть похідну функції f(x) = x^3 + 2x^2 - 5x + 1',
                'question_type': 'Тестове',
                'question_options': ['3x^2 + 4x - 5', '3x^2 + 2x - 5', 'x^2 + 4x - 5', '3x^3 + 4x^2 - 5x'],
                'student_answer': '1',
                'correct_answer': '1',
                'is_correct': True,
                'weight': 1.0,
                'points': 3.0,
                'max_points': 3.0,
                'is_test_question': True
            },
            {
                'question_number': 4,
                'question_text': 'Обчисліть sqrt(16) + 2^3',
                'question_type': 'Відкрите',
                'student_answer': '10',
                'correct_answer': '12',
                'is_correct': False,
                'weight': 1.0,
                'points': 0.0,
                'max_points': 3.0,
                'is_test_question': False
            }
        ]
    }
    
    # Створюємо Word звіт
    try:
        word_path = create_check_result_word(check_result, '.')
        print(f"Word звіт з математичними формулами створено: {word_path}")
        return True
    except Exception as e:
        print(f"Помилка при створенні Word звіту: {e}")
        return False

if __name__ == '__main__':
    test_word_report_with_math()