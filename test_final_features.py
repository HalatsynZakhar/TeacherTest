#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from core.processor import create_check_result_pdf, create_check_result_word

def test_final_features():
    """Тест финальных функций: цветные символы в Word, колонка с баллами, 12-балльная система"""
    
    # Тестовые данные
    check_result = {
        'variant_number': 1,
        'total_questions': 5,
        'correct_answers': 3,
        'score_percentage': 60.0,
        'detailed_results': [
            {'question_number': 1, 'student_answer': 'A', 'correct_answer': 'A', 'is_correct': True},
            {'question_number': 2, 'student_answer': 'B', 'correct_answer': 'C', 'is_correct': False},
            {'question_number': 3, 'student_answer': 'C', 'correct_answer': 'C', 'is_correct': True},
            {'question_number': 4, 'student_answer': 'A', 'correct_answer': 'B', 'is_correct': False},
            {'question_number': 5, 'student_answer': 'D', 'correct_answer': 'D', 'is_correct': True}
        ]
    }
    
    # Создаем PDF отчет
    pdf_path = create_check_result_pdf(check_result, "test_final_report.pdf")
    print(f"PDF отчет создан: {pdf_path}")
    
    # Создаем Word отчет
    word_path = create_check_result_word(check_result, "test_final_report.docx")
    print(f"Word отчет создан: {word_path}")
    
    print("\nПроверьте отчеты:")
    print("1. PDF: колонка 'Бали' с баллами за каждое задание, 12-балльная оценка")
    print("2. Word: цветные символы (зеленые галочки, красные крестики), колонка 'Бали', 12-балльная оценка")
    print("3. Ширина колонок должна быть правильно настроена")
    
    # Проверяем расчет 12-балльной системы
    expected_score_12 = (3 / 5) * 12  # 7.20
    print(f"\nОжидаемая оценка в 12-балльной системе: {expected_score_12:.2f}")
    
if __name__ == "__main__":
    test_final_features()