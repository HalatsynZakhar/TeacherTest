from core.processor import read_test_excel

try:
    print("Тестуємо валідацію з файлом, що містить помилки...")
    df = read_test_excel('test_validation.xlsx')
    print("ПОМИЛКА: Валідація не спрацювала! Файл був прочитаний без помилок.")
    print(f"Прочитано {len(df)} питань")
except ValueError as e:
    print("УСПІХ: Валідація спрацювала правильно!")
    print(f"Помилка: {e}")
except Exception as e:
    print(f"Неочікувана помилка: {e}")