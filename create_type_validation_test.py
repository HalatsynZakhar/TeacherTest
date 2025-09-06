import pandas as pd

# Створюємо тестові дані тільки з невизначеним типом завдання
test_data = {
    'question_number': [1, 2, 3],
    'question': [
        'Тестове питання з правильним типом',
        'Питання з невизначеним типом',
        'Відкрите питання'
    ],
    'correct_answer': ['А', 'Б', '42'],
    'weight': [1, 1, 2],
    'option_a': ['Правильна відповідь', 'Неправильна 1', ''],
    'option_b': ['Неправильна 1', 'Правильна відповідь', ''],
    'option_c': ['Неправильна 2', 'Неправильна 2', ''],
    'option_d': ['Неправильна 3', 'Неправильна 3', ''],
    'task_type': ['ТО', 'Невідомий тип', 'ВП']  # Другий тип невідомий
}

df = pd.DataFrame(test_data)

# Зберігаємо в Excel файл
df.to_excel('test_type_validation.xlsx', index=False)
print("Створено тестовий файл test_type_validation.xlsx з невизначеним типом завдання")
print("\nВміст файлу:")
print(df.to_string(index=False))