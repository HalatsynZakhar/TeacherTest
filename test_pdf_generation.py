from core.processor import read_test_excel, generate_test_variants, create_test_pdf
import os

# Читаємо тестові дані
file_path = 'test_numbers.xlsx'
print(f"Читання файлу: {file_path}")

try:
    df = read_test_excel(file_path)
    print(f"Успішно прочитано {len(df)} питань")
    
    # Генеруємо варіанти тестів
    print("\n=== Генерація варіантів тестів ===")
    variants = generate_test_variants(
        df, 
        num_variants=3, 
        question_shuffle_mode='full',
        answer_shuffle_mode='random'
    )
    
    print(f"Згенеровано {len(variants)} варіантів")
    for i, variant in enumerate(variants, 1):
        print(f"Варіант {i}: {len(variant['questions'])} питань, ключ відповідей: {variant['answer_key']}")
    
    # Створюємо PDF
    print("\n=== Створення PDF ===")
    test_pdf_path, answers_pdf_path = create_test_pdf(
        variants, 
        output_dir='.',
        columns=1
    )
    
    print(f"PDF створено: {test_pdf_path}")
    
    # Перевіряємо чи створилися файли
    if os.path.exists(test_pdf_path):
        print(f"✅ PDF файл з тестами успішно створено: {test_pdf_path}")
        print(f"Розмір файлу: {os.path.getsize(test_pdf_path)} байт")
    else:
        print("❌ Помилка: PDF файл з тестами не створено")
        
    if os.path.exists(answers_pdf_path):
        print(f"✅ PDF файл з відповідями успішно створено: {answers_pdf_path}")
        print(f"Розмір файлу: {os.path.getsize(answers_pdf_path)} байт")
    else:
        print("❌ Помилка: PDF файл з відповідями не створено")
        
except Exception as e:
    print(f"Помилка: {e}")
    import traceback
    traceback.print_exc()