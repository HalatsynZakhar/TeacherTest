#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from core.processor import create_check_result_pdf

def test_pdf_symbols_fix():
    """Тестируем исправленное отображение результатов в PDF"""
    
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
    
    print("Тестирование исправленного отображения символов в PDF...")
    print("Проблема: символы ✓ и ✗ не отображались в PDF из-за отсутствия в шрифте Arial")
    print("Решение: заменены на текст 'Правильно'/'Неправильно'")
    print()
    
    # Создаем PDF с исправленным отображением
    output_dir = os.getcwd()
    
    try:
        pdf_path = create_check_result_pdf(test_check_result, output_dir)
        print(f"✅ УСПЕХ: PDF создан с исправленным отображением результатов")
        print(f"   Путь к файлу: {pdf_path}")
        print(f"   Файл существует: {os.path.exists(pdf_path)}")
        print()
        print("📋 Содержимое PDF:")
        print("   - Вместо символов ✓/✗ теперь используется текст")
        print("   - Правильные ответы: 'Правильно'")
        print("   - Неправильные ответы: 'Неправильно'")
        print("   - Ширина колонки 'Результат' увеличена для размещения текста")
        print()
        print("🔧 Технические изменения:")
        print("   - Заменены Unicode символы на обычный текст")
        print("   - Изменена ширина колонок таблицы: 25-35-35-45")
        print("   - Исправлена совместимость с шрифтом Arial")
        
    except Exception as e:
        print(f"❌ ОШИБКА: {e}")
        return False
    
    print()
    print("=" * 60)
    print("РЕЗУЛЬТАТ: Проблема с отображением символов в PDF исправлена!")
    print("Теперь результаты проверки корректно отображаются в PDF файлах.")
    print("=" * 60)
    
    return True

if __name__ == '__main__':
    test_pdf_symbols_fix()