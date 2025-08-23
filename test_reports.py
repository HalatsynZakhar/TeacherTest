#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from core.processor import create_check_result_pdf, create_check_result_word

def test_reports():
    """Тестируем создание отчетов с результатами проверки"""
    
    # Создаем тестовые данные результата проверки
    test_check_result = {
        'variant_number': 1,
        'total_questions': 5,
        'correct_answers': 3,
        'score_percentage': 60.0,
        'detailed_results': [
            {
                'question_number': 1,
                'student_answer': 2,
                'correct_answer': 2,
                'is_correct': True
            },
            {
                'question_number': 2,
                'student_answer': 1,
                'correct_answer': 3,
                'is_correct': False
            },
            {
                'question_number': 3,
                'student_answer': 4,
                'correct_answer': 4,
                'is_correct': True
            },
            {
                'question_number': 4,
                'student_answer': 2,
                'correct_answer': 1,
                'is_correct': False
            },
            {
                'question_number': 5,
                'student_answer': 3,
                'correct_answer': 3,
                'is_correct': True
            }
        ]
    }
    
    output_dir = os.path.dirname(os.path.abspath(__file__))
    
    print("Создание тестовых отчетов...")
    
    # Создаем PDF отчет
    try:
        pdf_path = create_check_result_pdf(test_check_result, output_dir)
        print(f"✅ PDF отчет создан: {pdf_path}")
    except Exception as e:
        print(f"❌ Ошибка создания PDF отчета: {e}")
    
    # Создаем Word отчет
    try:
        word_path = create_check_result_word(test_check_result, output_dir)
        print(f"✅ Word отчет создан: {word_path}")
    except Exception as e:
        print(f"❌ Ошибка создания Word отчета: {e}")
    
    print("\nТест завершен. Проверьте созданные файлы.")

if __name__ == '__main__':
    test_reports()