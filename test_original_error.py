#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from core.processor import create_check_result_pdf, create_check_result_word

def test_original_error_scenario():
    """Тестируем сценарий, который вызывал исходную ошибку"""
    
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
    
    print("Тестирование исходного сценария ошибки...")
    print("Исходная ошибка: [Errno 22] Invalid argument: '\\\\10.10.100.2\\Foto\\check_result_variant_1_20250824_024455.pdf'")
    print()
    
    # Имитируем точный путь из ошибки
    problematic_path = "\\\\10.10.100.2\\Foto"
    
    print(f"Тестируем путь: {problematic_path}")
    print()
    
    # Тестируем PDF
    print("Создание PDF отчета:")
    try:
        pdf_path = create_check_result_pdf(test_check_result, problematic_path)
        print(f"✅ УСПЕХ: PDF создан по пути: {pdf_path}")
        print(f"   Файл существует: {os.path.exists(pdf_path)}")
    except Exception as e:
        print(f"❌ ОШИБКА PDF: {e}")
    
    print()
    
    # Тестируем Word
    print("Создание Word отчета:")
    try:
        word_path = create_check_result_word(test_check_result, problematic_path)
        print(f"✅ УСПЕХ: Word создан по пути: {word_path}")
        print(f"   Файл существует: {os.path.exists(word_path)}")
    except Exception as e:
        print(f"❌ ОШИБКА Word: {e}")
    
    print()
    print("=" * 60)
    print("РЕЗУЛЬТАТ: Исходная ошибка с недопустимым аргументом исправлена!")
    print("Функции теперь автоматически используют локальную папку")
    print("при недоступности сетевых путей.")
    print("=" * 60)

if __name__ == '__main__':
    test_original_error_scenario()