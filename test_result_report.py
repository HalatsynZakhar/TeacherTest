import pandas as pd
import os
from core.processor import generate_test_variants, create_check_result_word, _process_optional_questions

# Створюємо тестові дані
data = {
    'question_number': [1, 2, 3, 4, 5],
    'question': [
        'Яке з чисел є натуральним?',
        'Оберіть числа, які діляться на 3 без залишку.',
        'Обчисліть: 100 - 5 * 10',
        'Оберіть усі парні числа.',
        'Напиши відповідь на питання'
    ],
    'correct_answer': ['А', 'АВ', 'Б', 'АВДЖИ', 'Правильна відповідь'],
    'weight': [1, 2, 2, 4, 1],
    'option_1': ['12', '27', '950', '28', ''],
    'option_2': ['0', '13', '50', '15', ''],
    'option_3': ['', '81', '90', '80', ''],
    'option_4': ['', '44', '95', '31', ''],
    'option_5': ['', '', '105', '106', ''],
    'option_6': ['', '', '15', '99', ''],
    'option_7': ['', '', '500', '44', ''],
    'option_8': ['', '', '', '12', ''],
    'option_9': ['', '', '', '57', ''],
    'task_type': ['ТО', 'ТК', 'ТО', 'ТК', 'ВП']
}

df = pd.DataFrame(data)
df['is_test_question'] = df['task_type'] != 'ВП'

print("=== Генерація розширеного звіту результатів з індикаторами ===")

try:
    # Генеруємо варіант тесту
    variants = generate_test_variants(df, num_variants=1)
    variant = variants[0]
    
    print(f"Згенеровано варіант з {len(variant['questions'])} питань")
    
    # Створюємо відповіді учня (для демонстрації)
    student_answers = []
    for i, question in enumerate(variant['questions']):
        if question['is_test_question']:
            # Для різних питань даємо різні відповіді
            if i == 0:  # Перше питання - правильна відповідь
                student_answers.append(question['correct_answer'])
            elif i == 3:  # Четверте питання (з 9 варіантами) - неправильна відповідь
                student_answers.append('БВ')  # Замість правильної 'АВДЖИ'
            else:
                student_answers.append('Б')  # Інші - неправильні
        else:
            # Для відкритих питань даємо частково правильну відповідь
            student_answers.append('Частково правильно')
    
    # Створюємо папку для виводу
    output_dir = "test_output"
    os.makedirs(output_dir, exist_ok=True)
    
    # Створюємо структуру check_result для звіту
    check_result = {
        'variant_number': variant['variant_number'],
        'total_questions': len(variant['questions']),
        'correct_answers': 0,
        'score_percentage': 0.0,
        'weighted_score': 0.0,
        'max_score': sum(q['weight'] for q in variant['questions']),
        'student_info': {
            'class': '5-А',
            'full_name': 'Іван Петренко'
        },
        'detailed_results': []
    }
    
    # Заповнюємо детальні результати
    for i, question in enumerate(variant['questions']):
        student_answer = student_answers[i]
        correct_answer = question.get('correct_answer', '')
        is_correct = student_answer == correct_answer
        
        if is_correct:
            check_result['correct_answers'] += 1
            points = question['weight']
            check_result['weighted_score'] += points
        else:
            points = 0
        
        result_item = {
            'question_number': i + 1,
            'question_text': question['question_text'],
            'question_type': 'Тестове' if question['is_test_question'] else 'Відкрите',
            'student_answer': student_answer,
            'correct_answer': correct_answer,
            'is_correct': is_correct,
            'weight': question['weight'],
            'points': points,
            'max_points': question['weight'],
            'is_test_question': question['is_test_question']
        }
        
        if question['is_test_question']:
            result_item['question_options'] = question.get('options', [])
        
        check_result['detailed_results'].append(result_item)
    
    check_result['score_percentage'] = (check_result['correct_answers'] / check_result['total_questions']) * 100
    
    # Створюємо розширений звіт результатів
    result_path = create_check_result_word(check_result, output_dir)
    
    print(f"Розширений звіт створено: {result_path}")
    
    print("\n=== Перевірка індикаторів у звіті ===")
    for i, question in enumerate(variant['questions'], 1):
        if question['is_test_question']:
            correct_answer = question.get('correct_answer', '')
            if isinstance(correct_answer, str) and len(correct_answer.split(',')) > 1:
                indicator = "(Виберіть всі правильні відповіді)"
            else:
                indicator = "(Виберіть одну правильну відповідь)"
            q_type = "Тестове"
        else:
            indicator = "(Запишіть відповідь)"
            q_type = "Відкрите"
        
        print(f"  {i}. {question['question_text'][:40]}... {q_type} {indicator}")
    
except Exception as e:
    print(f"Помилка при генерації: {e}")
    import traceback
    traceback.print_exc()