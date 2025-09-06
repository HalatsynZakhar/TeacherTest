import pandas as pd
import os
from core.processor import generate_test_variants, create_check_result_word

# Тестові дані з питанням, що має 9 варіантів відповідей
data = {
    'question_number': [1],
    'question': ['Оберіть усі парні числа.'],
    'correct_answer': ['АВДЖИ'],  # Правильні відповіді: А(28), В(80), Д(106), Ж(44), И(12)
    'weight': [4],
    'option_1': ['28'],   # А - правильно
    'option_2': ['15'],   # Б - неправильно
    'option_3': ['80'],   # В - правильно
    'option_4': ['31'],   # Г - неправильно
    'option_5': ['106'],  # Д - правильно
    'option_6': ['99'],   # Е - неправильно
    'option_7': ['44'],   # Ж - правильно
    'option_8': ['12'],   # З - неправильно (помилка в даних)
    'option_9': ['57'],   # И - неправильно (помилка в даних)
    'task_type': ['ТК']
}

# Виправляємо дані відповідно до правильних відповідей
data['option_8'] = ['12']  # И - правильно
data['option_9'] = ['57']  # К - неправильно

df = pd.DataFrame(data)
df['is_test_question'] = df['task_type'] != 'ВП'

print("=== Тест перевірки позначень у розширеному звіті ===\n")

try:
    # Генеруємо варіант
    variants = generate_test_variants(df, num_variants=1)
    variant = variants[0]
    
    print(f"Питання: {variant['questions'][0]['question_text']}")
    print(f"Правильна відповідь: {variant['questions'][0]['correct_answer']}")
    print(f"Варіанти відповідей: {variant['questions'][0]['options']}")
    print()
    
    # Створюємо директорію для результатів
    output_dir = "test_output"
    os.makedirs(output_dir, exist_ok=True)
    
    # Тестуємо різні варіанти відповідей учня
    test_cases = [
        ('БВ', 'Учень обрав Б(15) та В(80) - частково правильно'),
        ('АВДЖИ', 'Учень обрав всі правильні відповіді'),
        ('БГЕ', 'Учень обрав тільки неправильні відповіді'),
        ('АВД', 'Учень обрав частину правильних відповідей')
    ]
    
    for student_answer, description in test_cases:
        print(f"\n=== {description} ===")
        print(f"Відповідь учня: {student_answer}")
        
        # Створюємо check_result
        check_result = {
            'variant_number': 1,
            'total_questions': 1,
            'correct_answers': 1 if student_answer == 'АВДЖИ' else 0,
            'score_percentage': 100.0 if student_answer == 'АВДЖИ' else 0.0,
            'weighted_score': 4.0 if student_answer == 'АВДЖИ' else 0.0,
            'max_score': 4.0,
            'student_info': {
                'class': '5-А',
                'full_name': f'Тест_{student_answer}'
            },
            'detailed_results': [{
                'question_number': 1,
                'question_text': variant['questions'][0]['question_text'],
                'question_type': 'Тестове',
                'student_answer': student_answer,
                'correct_answer': variant['questions'][0]['correct_answer'],
                'is_correct': student_answer == variant['questions'][0]['correct_answer'],
                'weight': 4,
                'points': 4.0 if student_answer == 'АВДЖИ' else 0.0,
                'max_points': 4.0,
                'is_test_question': True,
                'question_options': variant['questions'][0]['options']
            }]
        }
        
        # Створюємо звіт
        result_path = create_check_result_word(check_result, output_dir)
        print(f"Звіт створено: {result_path}")
        
        # Показуємо, які позначення мають бути у звіті
        print("Очікувані позначення:")
        ukrainian_letters = ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ж', 'З', 'И']
        correct_letters = set('АВДЖИ')
        student_letters = set(student_answer)
        
        for i, option in enumerate(variant['questions'][0]['options'], 1):
            letter = ukrainian_letters[i-1]
            status = []
            
            if letter in student_letters and letter in correct_letters:
                status.append("✓ Учень відповів - ПРАВИЛЬНО")
            elif letter in student_letters:
                status.append("✗ Учень відповів - НЕПРАВИЛЬНО")
            elif letter in correct_letters:
                status.append("✓ Правильна відповідь")
            
            status_text = f" ({', '.join(status)})" if status else ""
            print(f"  {i}. {option}{status_text}")

except Exception as e:
    print(f"Помилка: {e}")
    import traceback
    traceback.print_exc()