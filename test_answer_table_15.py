#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from core.processor import create_test_word

def test_answer_table_15():
    """Тестирует новую логику таблицы ответов с максимумом 15 элементов"""
    
    # Тест 1: 10 вопросов (меньше 15) - должна быть одна строка
    print("Тест 1: 10 вопросов (одна строка)")
    questions_10 = []
    for i in range(10):
        question = {
            'question_text': f"Вопрос {i+1}?",
            'options': ["A) Вариант 1", "B) Вариант 2", "C) Вариант 3", "D) Вариант 4"],
            'correct_answer': 'A'
        }
        questions_10.append(question)
    
    variant_10 = {'variant_number': 1, 'questions': questions_10}
    create_test_word([variant_10], "test_10_questions.docx", columns=2)
    print("✓ Создан файл test_10_questions.docx")
    
    # Тест 2: 20 вопросов (больше 15, меньше 30) - должно быть 2 строки
    print("\nТест 2: 20 вопросов (две строки)")
    questions_20 = []
    for i in range(20):
        question = {
            'question_text': f"Вопрос {i+1}?",
            'options': ["A) Вариант 1", "B) Вариант 2", "C) Вариант 3", "D) Вариант 4"],
            'correct_answer': 'A'
        }
        questions_20.append(question)
    
    variant_20 = {'variant_number': 1, 'questions': questions_20}
    create_test_word([variant_20], "test_20_questions.docx", columns=2)
    print("✓ Создан файл test_20_questions.docx")
    
    # Тест 3: 35 вопросов (больше 30) - должно быть 3 строки
    print("\nТест 3: 35 вопросов (три строки)")
    questions_35 = []
    for i in range(35):
        question = {
            'question_text': f"Вопрос {i+1}?",
            'options': ["A) Вариант 1", "B) Вариант 2", "C) Вариант 3", "D) Вариант 4"],
            'correct_answer': 'A'
        }
        questions_35.append(question)
    
    variant_35 = {'variant_number': 1, 'questions': questions_35}
    create_test_word([variant_35], "test_35_questions.docx", columns=2)
    print("✓ Создан файл test_35_questions.docx")
    
    # Тест 4: 17 вопросов (нечетное число больше 15) - должна добавиться пустая ячейка
    print("\nТест 4: 17 вопросов (нечетное число, должна добавиться пустая ячейка)")
    questions_17 = []
    for i in range(17):
        question = {
            'question_text': f"Вопрос {i+1}?",
            'options': ["A) Вариант 1", "B) Вариант 2", "C) Вариант 3", "D) Вариант 4"],
            'correct_answer': 'A'
        }
        questions_17.append(question)
    
    variant_17 = {'variant_number': 1, 'questions': questions_17}
    create_test_word([variant_17], "test_17_questions.docx", columns=2)
    print("✓ Создан файл test_17_questions.docx")
    
    print("\n=== Все тесты завершены ===")
    print("Проверьте созданные файлы:")
    print("- test_10_questions.docx: одна строка с 10 ячейками")
    print("- test_20_questions.docx: две строки по 10 ячеек")
    print("- test_35_questions.docx: три строки (12+12+11 или 12+12+12)")
    print("- test_17_questions.docx: две строки (9+8 с добавлением пустой ячейки)")
    print("\nВсе таблицы должны быть растянуты по всей ширине страницы.")
    print("Заголовок варианта должен быть по центру вне колонок.")
    print("Таблица ответов должна идти сразу после последнего вопроса.")

if __name__ == "__main__":
    test_answer_table_15()