from core.processor import read_test_excel

try:
    print("Тестуємо валідацію типів завдань...")
    df = read_test_excel('test_type_validation.xlsx')
    print("ПОМИЛКА: Валідація типів завдань не спрацювала! Файл був прочитаний без помилок.")
    print(f"Прочитано {len(df)} питань")
except ValueError as e:
    print("УСПІХ: Валідація типів завдань спрацювала правильно!")
    print(f"Помилка: {e}")
except Exception as e:
    print(f"Неочікувана помилка: {e}")