import pandas as pd
from core.processor import read_test_excel, generate_test_variants, _process_optional_questions

# Створюємо тестові дані як у користувача
data = {
    'question_number': [1, 1, 2, 2, 3, 3, 4, 4, 5, 5],
    'question': [
        'Яке з чисел є натуральним?',
        'Який з наведених дробів є правильним?',
        'Оберіть числа, які діляться на 3 без залишку.',
        'Які з цих фігур є прямокутниками?',
        'Обчисліть: 100 - 5 * 10',
        'Який залишок від ділення 47 на 5?',
        'Оберіть усі парні числа.',
        'Оберіть усі числа, які діляться на 5.',
        'Напиши 5',
        'Слава Україні!'
    ],
    'correct_answer': ['А', 'Б', 'АВ', 'БГ', 'Б', 'В', 'АВДЖИ', 'АВДЖИ', '5', 'Героям Слава'],
    'weight': [1, 1, 2, 2, 2, 2, 3, 3, 1, 1],
    'option_1': ['12', '12 / 5', '27', 'Трикутник', '950', '0', '12', '10', '', ''],
    'option_2': ['0', '5 / 12', '13', 'Квадрат', '50', '1', '15', '12', '', ''],
    'option_3': ['', '', '81', 'Коло', '90', '2', '28', '25', '', ''],
    'option_4': ['', '', '44', 'Прямокутник', '95', '3', '31', '34', '', ''],
    'option_5': ['', '', '', '', '105', '4', '44', '50', '', ''],
    'option_6': ['', '', '', '', '15', '5', '57', '63', '', ''],
    'option_7': ['', '', '', '', '500', '7', '80', '', '', ''],
    'option_8': ['', '', '', '', '', '', '99', '', '', ''],
    'option_9': ['', '', '', '', '', '', '106', '', '', ''],
    'task_type': ['ТО', 'ТО', 'ТК', 'ТК', 'ТО', 'ТО', 'ТК', 'ТК', 'ВП', 'ВП']
}

df = pd.DataFrame(data)

# Додаємо поле is_test_question як це робить read_test_excel
df['is_test_question'] = df['task_type'] != 'ВП'

print("=== Початкові дані ===")
print(f"Всього питань: {len(df)}")
print("Питання за номерами:")
for num in sorted(df['question_number'].unique()):
    questions_for_num = df[df['question_number'] == num]
    print(f"  Номер {num}: {len(questions_for_num)} питань")
    for _, row in questions_for_num.iterrows():
        print(f"    - {row['question'][:50]}... (тип: {row['task_type']})")

print("\n=== Обробка опціональних питань ===")
processed_df = _process_optional_questions(df)
print(f"Після обробки: {len(processed_df)} питань")
print("Обрані питання:")
for _, row in processed_df.iterrows():
    print(f"  {row['question_number']}. {row['question'][:50]}... (тип: {row['task_type']})")

print("\n=== Перевірка чи є відкриті питання ===")
open_questions = processed_df[processed_df['task_type'] == 'ВП']
print(f"Відкритих питань після обробки: {len(open_questions)}")
if len(open_questions) > 0:
    for _, row in open_questions.iterrows():
        print(f"  - {row['question']} (відповідь: {row['correct_answer']})")
else:
    print("  Відкриті питання відсутні!")

print("\n=== Генерація тестів ===")
try:
    variants = generate_test_variants(processed_df, num_variants=1)
    print(f"Згенеровано {len(variants)} варіантів")
    
    variant = variants[0]
    print(f"Варіант містить {len(variant['questions'])} питань:")
    for i, q in enumerate(variant['questions'], 1):
        q_type = "Тестове" if q['is_test_question'] else "Відкрите"
        print(f"  {i}. {q['question_text'][:50]}... ({q_type})")
except Exception as e:
    print(f"Помилка при генерації: {e}")