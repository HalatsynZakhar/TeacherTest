import pandas as pd
import os
from core.processor import read_test_excel, generate_test_variants, create_test_word, _process_optional_questions

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

print("=== Генерація Word документу з індикаторами типу завдання ===")

try:
    # Генеруємо варіанти тестів
    variants = generate_test_variants(df, num_variants=2)
    print(f"Згенеровано {len(variants)} варіантів")
    
    # Створюємо папку для виводу
    output_dir = "test_output"
    os.makedirs(output_dir, exist_ok=True)
    
    # Створюємо Word документ з тестами
    word_path = create_test_word(
        variants=variants,
        output_dir=output_dir,
        input_file_name="тест_з_індикаторами",
        answer_format="list",
        space_optimization=False,
        test_class="5-А",
        test_date="2024-01-15"
    )
    
    print(f"Word документ створено: {word_path}")
    
    # Також створюємо компактну версію
    word_path_compact = create_test_word(
        variants=variants,
        output_dir=output_dir,
        input_file_name="тест_з_індикаторами_компактний",
        answer_format="list",
        space_optimization=True,
        test_class="5-А",
        test_date="2024-01-15"
    )
    
    print(f"Компактний Word документ створено: {word_path_compact}")
    
    print("\n=== Перевірка індикаторів у варіантах ===")
    for variant in variants:
        print(f"\nВаріант {variant['variant_number']}:")
        for i, question in enumerate(variant['questions'], 1):
            if question['is_test_question']:
                correct_answer = question.get('correct_answer', '')
                if isinstance(correct_answer, str) and len(correct_answer.split(',')) > 1:
                    indicator = "(Виберіть всі правильні відповіді)"
                else:
                    indicator = "(Виберіть одну правильну відповідь)"
            else:
                indicator = "(Запишіть відповідь)"
            
            print(f"  {i}. {question['question_text'][:40]}... {indicator}")
    
except Exception as e:
    print(f"Помилка при генерації: {e}")
    import traceback
    traceback.print_exc()