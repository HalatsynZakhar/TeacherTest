import pandas as pd
import os
from core.processor import check_student_answers, create_check_result_word
from datetime import datetime

def create_scoring_comparison_demo():
    """
    Створює демонстрацію різних підходів до оцінювання складних завдань.
    """
    print("=== Демонстрація системи часткового оцінювання ===")
    
    # Створюємо тестові дані
    test_data = {
        'Варіант': [1],
        'Відповіді': ['АВДЖИ,Б,В'],  # Складне завдання + 2 прості
        'Ваги': ['5,1,1']  # Складне завдання має більшу вагу
    }
    
    detailed_data = [
        {
            'Варіант': 1,
            'Номер_питання': 1,
            'Текст_питання': 'Оберіть всі правильні твердження про фотосинтез:',
            'Тип_питання': 'Тестове',
            'Правильна_відповідь': 'АВДЖИ',
            'Вага': 5,
            'Варіант_1': 'А) Відбувається в хлоропластах',
            'Варіант_2': 'Б) Потребує тільки води',
            'Варіант_3': 'В) Виділяє кисень',
            'Варіант_4': 'Г) Відбувається тільки вдень',
            'Варіант_5': 'Д) Перетворює CO₂ на глюкозу',
            'Варіант_6': 'Е) Потребує світла',
            'Варіант_7': 'Ж) Відбувається в мітохондріях',
            'Варіант_8': 'З) Споживає кисень',
            'Варіант_9': 'И) Створює органічні речовини'
        },
        {
            'Варіант': 1,
            'Номер_питання': 2,
            'Текст_питання': 'Яка формула води?',
            'Тип_питання': 'Тестове',
            'Правильна_відповідь': 'Б',
            'Вага': 1,
            'Варіант_1': 'А) H₂SO₄',
            'Варіант_2': 'Б) H₂O'
        },
        {
            'Варіант': 1,
            'Номер_питання': 3,
            'Текст_питання': 'Скільки хромосом у людини?',
            'Тип_питання': 'Тестове',
            'Правильна_відповідь': 'В',
            'Вага': 1,
            'Варіант_1': 'А) 44',
            'Варіант_2': 'Б) 45',
            'Варіант_3': 'В) 46'
        }
    ]
    
    # Створюємо тимчасовий Excel файл
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    test_file = f'demo_scoring_{timestamp}.xlsx'
    
    with pd.ExcelWriter(test_file, engine='openpyxl') as writer:
        pd.DataFrame(test_data).to_excel(writer, sheet_name='Основні_відповіді', index=False)
        pd.DataFrame(detailed_data).to_excel(writer, sheet_name='Детальна_інформація', index=False)
    
    print(f"Створено демонстраційний файл: {test_file}")
    
    # Різні сценарії відповідей учнів
    scenarios = [
        {
            'name': 'Відмінник Олексій',
            'description': 'Знає все ідеально',
            'answers': ['АВДЖИ', 'Б', 'В'],
            'color': 'зелений'
        },
        {
            'name': 'Хорошист Марія',
            'description': 'Знає майже все, але пропустила одну деталь',
            'answers': ['АВДЖ', 'Б', 'В'],  # Пропустила И
            'color': 'світло-зелений'
        },
        {
            'name': 'Середняк Петро',
            'description': 'Знає основи, але плутається в деталях',
            'answers': ['АВБ', 'Б', 'В'],  # А,В правильні, Б - помилка
            'color': 'жовтий'
        },
        {
            'name': 'Слабкий Іван',
            'description': 'Знає мало, багато помилок',
            'answers': ['АБЗ', 'А', 'А'],  # Тільки А правильна в складному завданні
            'color': 'помаранчевий'
        },
        {
            'name': 'Невстигаючий Сергій',
            'description': 'Відповідав навмання',
            'answers': ['БЗК', 'А', 'А'],  # Всі неправильні
            'color': 'червоний'
        }
    ]
    
    print("\n=== ПОРІВНЯННЯ СИСТЕМ ОЦІНЮВАННЯ ===")
    print("\nПравильні відповіді:")
    print("1. АВДЖИ (складне завдання, 5 балів)")
    print("2. Б (просте завдання, 1 бал)")
    print("3. В (просте завдання, 1 бал)")
    print("\nМаксимум: 12 балів")
    
    print("\n" + "="*80)
    print(f"{'Учень':<20} {'Відповіді':<15} {'Нова система':<15} {'Стара система':<15} {'Різниця':<10}")
    print("="*80)
    
    for scenario in scenarios:
        try:
            # Перевіряємо з новою системою
            result = check_student_answers(test_file, 1, scenario['answers'])
            new_score = result['weighted_score']
            new_percent = result['score_percentage']
            
            # Симулюємо стару систему (бінарне оцінювання)
            old_score = 0
            if scenario['answers'][0] == 'АВДЖИ':  # Складне завдання
                old_score += 5 * (12/7)  # 5 балів з 7 загальної ваги
            if scenario['answers'][1] == 'Б':  # Просте завдання
                old_score += 1 * (12/7)
            if scenario['answers'][2] == 'В':  # Просте завдання
                old_score += 1 * (12/7)
            
            difference = new_score - old_score
            
            answers_str = ','.join(scenario['answers'])
            
            print(f"{scenario['name']:<20} {answers_str:<15} {new_score:>6.1f}б ({new_percent:>4.1f}%) {old_score:>6.1f}б ({old_score/12*100:>4.1f}%) {difference:>+6.1f}б")
            
        except Exception as e:
            print(f"{scenario['name']:<20} ПОМИЛКА: {e}")
    
    print("="*80)
    
    # Створюємо детальні звіти для демонстрації
    print("\n=== Створення детальних звітів ===")
    
    output_dir = "demo_output"
    os.makedirs(output_dir, exist_ok=True)
    
    for scenario in scenarios:
        try:
            result = check_student_answers(test_file, 1, scenario['answers'])
            
            # Створюємо Word звіт
            doc_path = create_check_result_word(result, output_dir)
            
            print(f"Створено звіт для {scenario['name']}: {os.path.basename(doc_path)}")
            
        except Exception as e:
            print(f"Помилка створення звіту для {scenario['name']}: {e}")
    
    # Видаляємо тимчасовий файл
    try:
        os.remove(test_file)
        print(f"\nВидалено тимчасовий файл: {test_file}")
    except:
        pass
    
    print("\n=== ВИСНОВКИ ===")
    print("\n✅ ПЕРЕВАГИ НОВОЇ СИСТЕМИ:")
    print("• Справедливіше оцінювання складних завдань")
    print("• Часткові бали за правильні елементи")
    print("• Штрафи за неправильні відповіді")
    print("• Мотивація до більш ретельного вивчення")
    
    print("\n⚠️  ОСОБЛИВОСТІ:")
    print("• Прості завдання залишаються бінарними (0 або повний бал)")
    print("• Складні завдання (>1 літери) оцінюються частково")
    print("• Штраф: 50% від базового балу за кожну помилку")
    print("• Мінімальний бал: 0 (не може бути від'ємним)")
    
    print(f"\n📁 Детальні звіти збережено в папці: {output_dir}")
    print("\n=== Демонстрація завершена ===")

if __name__ == '__main__':
    create_scoring_comparison_demo()