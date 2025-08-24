#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тест для проверки новой структуры Word документов:
- Таблица ответов ограничена 10 элементами в строке
- Таблица ответов размещается сразу после теста без разрыва страницы
- Заголовок варианта общий по центру перед колонками
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from core.processor import generate_test_variants, create_test_word
import pandas as pd

def create_test_data():
    """Создает тестовые данные с 25 вопросами для проверки структуры"""
    data = []
    for i in range(1, 26):  # 25 вопросов для тестирования таблицы ответов
        data.append({
            'question': f'Тестовый вопрос номер {i}?',
            'option_1': f'Вариант A для вопроса {i}',
            'option_2': f'Вариант B для вопроса {i}',
            'option_3': f'Вариант C для вопроса {i}',
            'option_4': f'Вариант D для вопроса {i}',
            'correct_answer': (i % 4) + 1  # Циклически 1,2,3,4
        })
    
    return pd.DataFrame(data)

def main():
    print("Создание тестовых данных...")
    df = create_test_data()
    
    print("Генерация вариантов тестов...")
    variants = generate_test_variants(df, 2)  # 2 варианта
    
    output_dir = "test_word_structure_reports"
    os.makedirs(output_dir, exist_ok=True)
    
    print("\nТестирование различных конфигураций колонок:")
    
    # Тест 1: Одна колонка
    print("1. Создание теста с одной колонкой...")
    word_path_1 = create_test_word(variants, output_dir, columns=1)
    print(f"   Создан: {word_path_1}")
    
    # Тест 2: Две колонки
    print("2. Создание теста с двумя колонками...")
    word_path_2 = create_test_word(variants, output_dir, columns=2)
    print(f"   Создан: {word_path_2}")
    
    # Тест 3: Три колонки
    print("3. Создание теста с тремя колонками...")
    word_path_3 = create_test_word(variants, output_dir, columns=3)
    print(f"   Создан: {word_path_3}")
    
    print("\nВсе тесты созданы успешно!")
    print("\nПроверьте следующие аспекты в созданных документах:")
    print("✓ Таблица ответов разбита на строки по 10 элементов")
    print("✓ Таблица ответов размещена сразу после вопросов")
    print("✓ Заголовок варианта расположен по центру перед колонками")
    print("✓ Вопросы корректно распределены по колонкам")
    print("✓ Нумерация вопросов непрерывная")

if __name__ == "__main__":
    main()