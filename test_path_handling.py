#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from core.processor import create_check_result_pdf, create_check_result_word

def test_path_handling():
    """Тестируем обработку проблемных путей в функциях создания отчетов"""
    
    # Создаем тестовые данные результата проверки
    test_check_result = {
        'variant_number': 1,
        'total_questions': 3,
        'correct_answers': 2,
        'score_percentage': 66.7,
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
            }
        ]
    }
    
    print("Тестирование обработки путей...")
    
    # Тест 1: Нормальный локальный путь
    print("\n1. Тест с нормальным локальным путем:")
    normal_path = os.path.dirname(os.path.abspath(__file__))
    try:
        pdf_path = create_check_result_pdf(test_check_result, normal_path)
        print(f"✅ PDF создан: {pdf_path}")
    except Exception as e:
        print(f"❌ Ошибка PDF: {e}")
    
    try:
        word_path = create_check_result_word(test_check_result, normal_path)
        print(f"✅ Word создан: {word_path}")
    except Exception as e:
        print(f"❌ Ошибка Word: {e}")
    
    # Тест 2: Проблемный сетевой путь (имитация)
    print("\n2. Тест с проблемным сетевым путем:")
    problematic_path = "\\\\10.10.100.2\\Foto"
    try:
        pdf_path = create_check_result_pdf(test_check_result, problematic_path)
        print(f"✅ PDF создан (fallback): {pdf_path}")
    except Exception as e:
        print(f"❌ Ошибка PDF: {e}")
    
    try:
        word_path = create_check_result_word(test_check_result, problematic_path)
        print(f"✅ Word создан (fallback): {word_path}")
    except Exception as e:
        print(f"❌ Ошибка Word: {e}")
    
    # Тест 3: Несуществующий локальный путь
    print("\n3. Тест с несуществующим локальным путем:")
    nonexistent_path = "C:\\NonExistentFolder\\SubFolder"
    try:
        pdf_path = create_check_result_pdf(test_check_result, nonexistent_path)
        print(f"✅ PDF создан: {pdf_path}")
    except Exception as e:
        print(f"❌ Ошибка PDF: {e}")
    
    try:
        word_path = create_check_result_word(test_check_result, nonexistent_path)
        print(f"✅ Word создан: {word_path}")
    except Exception as e:
        print(f"❌ Ошибка Word: {e}")
    
    print("\nТест завершен. Проверьте созданные файлы и логи.")

if __name__ == '__main__':
    test_path_handling()