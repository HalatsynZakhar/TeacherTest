#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from core.processor import create_test_word

def test_unlimited_rows():
    """Тестирует новую логику таблицы ответов с неограниченным количеством строк"""
    
    # Тест 1: 50 вопросов - должно быть 4 строки (15+15+15+5)
    print("Тест 1: 50 вопросов (4 строки: 15+15+15+5)")
    questions_50 = []
    for i in range(50):
        question = {
            'question_text': f"Вопрос {i+1}?",
            'options': ["A) Вариант 1", "B) Вариант 2", "C) Вариант 3", "D) Вариант 4"],
            'correct_answer': 'A'
        }
        questions_50.append(question)
    
    variant_50 = {'variant_number': 1, 'questions': questions_50}
    create_test_word([variant_50], "test_50_questions.docx", columns=2)
    print("✓ Создан файл test_50_questions.docx")
    
    # Тест 2: 75 вопросов - должно быть 5 строк (15+15+15+15+15)
    print("\nТест 2: 75 вопросов (5 строк по 15 элементов)")
    questions_75 = []
    for i in range(75):
        question = {
            'question_text': f"Вопрос {i+1}?",
            'options': ["A) Вариант 1", "B) Вариант 2", "C) Вариант 3", "D) Вариант 4"],
            'correct_answer': 'A'
        }
        questions_75.append(question)
    
    variant_75 = {'variant_number': 1, 'questions': questions_75}
    create_test_word([variant_75], "test_75_questions.docx", columns=2)
    print("✓ Создан файл test_75_questions.docx")
    
    # Тест 3: 100 вопросов - должно быть 7 строк (15+15+15+15+15+15+10)
    print("\nТест 3: 100 вопросов (7 строк: 15+15+15+15+15+15+10)")
    questions_100 = []
    for i in range(100):
        question = {
            'question_text': f"Вопрос {i+1}?",
            'options': ["A) Вариант 1", "B) Вариант 2", "C) Вариант 3", "D) Вариант 4"],
            'correct_answer': 'A'
        }
        questions_100.append(question)
    
    variant_100 = {'variant_number': 1, 'questions': questions_100}
    create_test_word([variant_100], "test_100_questions.docx", columns=2)
    print("✓ Создан файл test_100_questions.docx")
    
    # Тест 4: 47 вопросов (нечетное число) - должно быть 4 строки (15+15+15+2 с добавлением пустой ячейки)
    print("\nТест 4: 47 вопросов (нечетное число, должна добавиться пустая ячейка в последней строке)")
    questions_47 = []
    for i in range(47):
        question = {
            'question_text': f"Вопрос {i+1}?",
            'options': ["A) Вариант 1", "B) Вариант 2", "C) Вариант 3", "D) Вариант 4"],
            'correct_answer': 'A'
        }
        questions_47.append(question)
    
    variant_47 = {'variant_number': 1, 'questions': questions_47}
    create_test_word([variant_47], "test_47_questions.docx", columns=2)
    print("✓ Создан файл test_47_questions.docx")
    
    print("\n=== Все тесты завершены ===")
    print("Проверьте созданные файлы:")
    print("- test_50_questions.docx: 4 строки (15+15+15+5)")
    print("- test_75_questions.docx: 5 строк по 15 элементов")
    print("- test_100_questions.docx: 7 строк (15+15+15+15+15+15+10)")
    print("- test_47_questions.docx: 4 строки (15+15+15+2 с пустой ячейкой)")
    print("\nВсе таблицы должны быть растянуты по всей ширине страницы.")
    print("Максимум 15 элементов в строке, неограниченное количество строк.")

if __name__ == "__main__":
    test_unlimited_rows()