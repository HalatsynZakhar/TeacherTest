import os
import sys
import logging
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import traceback
from datetime import datetime
import pandas as pd
from typing import List, Dict, Any, Optional

# Add current directory to path
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.append(current_dir)

# Import modules
from core import processor
from utils import config_manager
from utils.config_manager import get_downloads_folder
import tempfile
import shutil

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('teacher_test.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

def ensure_temp_dir(prefix: str = "") -> str:
    """Створює тимчасову папку всередині проекту"""
    try:
        project_dir = os.path.dirname(__file__)
        temp_dir = os.path.join(project_dir, "temp")
        os.makedirs(temp_dir, exist_ok=True)
        
        if prefix:
            temp_subdir = os.path.join(temp_dir, f"{prefix}{datetime.now().strftime('%Y%m%d_%H%M%S')}")
            os.makedirs(temp_subdir, exist_ok=True)
            return temp_subdir
        
        return temp_dir
    except Exception as e:
        logger.warning(f"Не вдалося створити тимчасову папку: {e}. Використовуємо системну тимчасову папку.")
        return tempfile.gettempdir()

def cleanup_temp_files():
    """Очищення тимчасових файлів при запуску додатку"""
    try:
        project_dir = os.path.dirname(__file__)
        temp_dir = os.path.join(project_dir, "temp")
        
        if os.path.exists(temp_dir):
            import time
            for filename in os.listdir(temp_dir):
                file_path = os.path.join(temp_dir, filename)
                try:
                    if os.path.isfile(file_path):
                        # Перевіряємо вік файлу (видаляємо файли старші 1 дня)
                        file_age = time.time() - os.path.getmtime(file_path)
                        if file_age > 86400:  # 24 години в секундах
                            os.remove(file_path)
                            logger.info(f"Видалено старий тимчасовий файл: {filename}")
                except Exception as e:
                    logger.warning(f"Не вдалося видалити тимчасовий файл {filename}: {e}")
    except Exception as e:
        logger.error(f"Помилка при очищенні тимчасових файлів: {e}")

# Очистка при запуске
cleanup_temp_files()

class TeacherTestApp:
    def __init__(self, root):
        self.root = root
        self.root.title("TeacherTest - Генератор тестів для учнів")
        self.root.geometry("800x700")
        
        # Змінні для зберігання шляхів
        self.excel_file_path = tk.StringVar()
        self.output_folder_path = tk.StringVar(value=ensure_temp_dir("output_"))
        self.answer_key_file_path = tk.StringVar()
        
        # Змінні для налаштувань
        self.num_variants = tk.IntVar(value=10)
        self.variant_number = tk.IntVar(value=1)
        self.student_answers_text = tk.StringVar()
        
        # Режим роботи (1 - генерація тестів, 2 - перевірка робіт)
        self.mode = tk.IntVar(value=1)
        
        self.setup_ui()
        
    def setup_ui(self):
        """Налаштування користувацького інтерфейсу"""
        # Головний фрейм
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Налаштування розтягування
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Заголовок
        title_label = ttk.Label(main_frame, text="TeacherTest - Генератор тестів", font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Вибір режиму роботи
        mode_frame = ttk.LabelFrame(main_frame, text="Режим роботи", padding="10")
        mode_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Radiobutton(mode_frame, text="Режим 1: Генерація тестів", variable=self.mode, value=1, command=self.on_mode_change).grid(row=0, column=0, sticky=tk.W)
        ttk.Radiobutton(mode_frame, text="Режим 2: Перевірка робіт", variable=self.mode, value=2, command=self.on_mode_change).grid(row=0, column=1, sticky=tk.W)
        
        # Фрейм для режиму 1 (генерація тестів)
        self.mode1_frame = ttk.LabelFrame(main_frame, text="Генерація тестів", padding="10")
        self.mode1_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        self.mode1_frame.columnconfigure(1, weight=1)
        
        # Вибір Excel файлу з питаннями
        ttk.Label(self.mode1_frame, text="Excel файл з питаннями:").grid(row=0, column=0, sticky=tk.W, pady=2)
        ttk.Entry(self.mode1_frame, textvariable=self.excel_file_path, width=50).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(5, 5), pady=2)
        ttk.Button(self.mode1_frame, text="Огляд", command=self.browse_excel_file).grid(row=0, column=2, pady=2)
        
        # Кількість варіантів
        ttk.Label(self.mode1_frame, text="Кількість варіантів:").grid(row=1, column=0, sticky=tk.W, pady=2)
        ttk.Spinbox(self.mode1_frame, from_=1, to=100, textvariable=self.num_variants, width=10).grid(row=1, column=1, sticky=tk.W, padx=(5, 0), pady=2)
        
        # Папка для збереження
        ttk.Label(self.mode1_frame, text="Папка для збереження:").grid(row=2, column=0, sticky=tk.W, pady=2)
        ttk.Entry(self.mode1_frame, textvariable=self.output_folder_path, width=50).grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(5, 5), pady=2)
        ttk.Button(self.mode1_frame, text="Огляд", command=self.browse_output_folder).grid(row=2, column=2, pady=2)
        
        # Кнопка генерації
        ttk.Button(self.mode1_frame, text="Згенерувати тести", command=self.generate_tests).grid(row=3, column=0, columnspan=3, pady=10)
        
        # Фрейм для режиму 2 (перевірка робіт)
        self.mode2_frame = ttk.LabelFrame(main_frame, text="Перевірка робіт", padding="10")
        self.mode2_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        self.mode2_frame.columnconfigure(1, weight=1)
        
        # Файл-ключ з відповідями
        ttk.Label(self.mode2_frame, text="Excel файл-ключ:").grid(row=0, column=0, sticky=tk.W, pady=2)
        ttk.Entry(self.mode2_frame, textvariable=self.answer_key_file_path, width=50).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(5, 5), pady=2)
        ttk.Button(self.mode2_frame, text="Огляд", command=self.browse_answer_key_file).grid(row=0, column=2, pady=2)
        
        # Номер варіанту учня
        ttk.Label(self.mode2_frame, text="Номер варіанту учня:").grid(row=1, column=0, sticky=tk.W, pady=2)
        ttk.Spinbox(self.mode2_frame, from_=1, to=100, textvariable=self.variant_number, width=10).grid(row=1, column=1, sticky=tk.W, padx=(5, 0), pady=2)
        
        # Відповіді учня
        ttk.Label(self.mode2_frame, text="Відповіді учня (через кому):").grid(row=2, column=0, sticky=tk.W, pady=2)
        answers_entry = ttk.Entry(self.mode2_frame, textvariable=self.student_answers_text, width=50)
        answers_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(5, 5), pady=2)
        
        # Додаємо підказку
        help_label = ttk.Label(self.mode2_frame, text="💡 Для тестових: номери від 1 до кількості варіантів. Для відкритих: текстові відповіді", 
                              font=('TkDefaultFont', 8), foreground='gray')
        help_label.grid(row=2, column=2, sticky=tk.W, padx=(5, 0), pady=2)
        
        # Кнопка перевірки
        ttk.Button(self.mode2_frame, text="Перевірити роботу", command=self.check_answers).grid(row=3, column=0, columnspan=3, pady=10)
        
        # Прогресс-бар
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 5))
        
        # Статус
        self.status_var = tk.StringVar(value="Готовий до роботи")
        self.status_label = ttk.Label(main_frame, textvariable=self.status_var)
        self.status_label.grid(row=5, column=0, columnspan=3, pady=5)
        
        # Лог
        log_frame = ttk.LabelFrame(main_frame, text="Журнал операцій", padding="5")
        log_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(6, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=10, width=80)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Спочатку показуємо тільки режим 1
        self.on_mode_change()
        
    def on_mode_change(self):
        """Обробник зміни режиму роботи"""
        if self.mode.get() == 1:
            self.mode1_frame.grid()
            self.mode2_frame.grid_remove()
        else:
            self.mode1_frame.grid_remove()
            self.mode2_frame.grid()
    
    def browse_excel_file(self):
        """Вибір Excel файлу з питаннями"""
        filename = filedialog.askopenfilename(
            title="Виберіть Excel файл з питаннями",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.excel_file_path.set(filename)
    
    def browse_answer_key_file(self):
        """Вибір Excel файлу-ключа"""
        filename = filedialog.askopenfilename(
            title="Виберіть Excel файл-ключ",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.answer_key_file_path.set(filename)
    
    def browse_output_folder(self):
        """Вибір папки для збереження"""
        folder = filedialog.askdirectory(title="Виберіть папку для збереження")
        if folder:
            self.output_folder_path.set(folder)
    
    def log_message(self, message):
        """Додавання повідомлення в журнал"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def update_progress(self, current, total, extra_info=""):
        """Оновлення прогрес-бару"""
        if total > 0:
            progress = (current / total) * 100
            self.progress_var.set(progress)
            status_text = f"Прогрес: {current}/{total}"
            if extra_info:
                status_text += f" - {extra_info}"
            self.status_var.set(status_text)
        self.root.update_idletasks()
    
    def generate_tests(self):
        """Генерація тестів"""
        # Перевірка вхідних даних
        if not self.excel_file_path.get():
            messagebox.showerror("Помилка", "Виберіть Excel файл з питаннями")
            return
        
        if not os.path.exists(self.excel_file_path.get()):
            messagebox.showerror("Помилка", "Вибраний Excel файл не існує")
            return
        
        if not self.output_folder_path.get():
            messagebox.showerror("Помилка", "Виберіть папку для збереження")
            return
        
        if self.num_variants.get() < 1:
            messagebox.showerror("Помилка", "Кількість варіантів повинна бути більше 0")
            return
        
        # Запуск генерації в окремому потоці
        thread = threading.Thread(target=self._generate_tests_thread)
        thread.daemon = True
        thread.start()
    
    def _generate_tests_thread(self):
        """Потік для генерації тестів"""
        try:
            self.log_message("Початок генерації тестів...")
            self.status_var.set("Генерація тестів...")
            
            # Читання Excel файлу
            self.log_message(f"Читання файлу: {self.excel_file_path.get()}")
            df = processor.read_test_excel(self.excel_file_path.get())
            
            if df.empty:
                raise ValueError("Excel файл не містить даних")
            
            self.log_message(f"Знайдено {len(df)} питань")
            
            # Генерація варіантів
            self.log_message(f"Генерація {self.num_variants.get()} варіантів...")
            variants = processor.generate_test_variants(df, self.num_variants.get())
            
            # PDF файли відключені - використовуємо тільки Word та Excel
            
            # Створення Excel файлу-ключа
            self.log_message("Створення Excel файлу-ключа...")
            input_file_name = os.path.splitext(os.path.basename(self.excel_file_path.get()))[0]
            excel_key_path = processor.create_excel_answer_key(variants, self.output_folder_path.get(), input_file_name)
            
            self.log_message(f"Готово! Створені файли:")
            self.log_message(f"- Excel ключ: {excel_key_path}")
            
            self.status_var.set("Генерація завершена успішно")
            self.progress_var.set(100)
            
            # Показуємо повідомлення про успіх
            self.root.after(0, lambda: messagebox.showinfo(
                "Успіх", 
                f"Тести успішно згенеровано!\n\n"
                f"Файли збережено в: {self.output_folder_path.get()}"
            ))
            
        except Exception as e:
            error_msg = f"Помилка при генерації тестів: {str(e)}"
            self.log_message(error_msg)
            logger.error(error_msg, exc_info=True)
            self.status_var.set("Помилка генерації")
            self.root.after(0, lambda: messagebox.showerror("Помилка", error_msg))
    
    def check_answers(self):
        """Перевірка відповідей учня"""
        # Перевірка вхідних даних
        if not self.answer_key_file_path.get():
            messagebox.showerror("Помилка", "Виберіть Excel файл-ключ")
            return
        
        if not os.path.exists(self.answer_key_file_path.get()):
            messagebox.showerror("Помилка", "Вибраний файл-ключ не існує")
            return
        
        if self.variant_number.get() < 1:
            messagebox.showerror("Помилка", "Номер варіанту повинен бути більше 0")
            return
        
        if not self.student_answers_text.get().strip():
            messagebox.showerror("Помилка", "Введіть відповіді учня")
            return
        
        # Запуск перевірки в окремому потоці
        thread = threading.Thread(target=self._check_answers_thread)
        thread.daemon = True
        thread.start()
    
    def _check_answers_thread(self):
        """Потік для перевірки відповідей"""
        try:
            self.log_message("Початок перевірки відповідей...")
            self.status_var.set("Перевірка відповідей...")
            
            # Парсинг відповідей учня
            answers_text = self.student_answers_text.get().strip()
            try:
                student_answers = [x.strip() for x in answers_text.split(',') if x.strip()]
            except ValueError:
                raise ValueError("Відповіді повинні бути розділені комами")
            
            self.log_message(f"Відповіді учня: {student_answers}")
            
            # Перевірка відповідей
            check_result = processor.check_student_answers(
                self.answer_key_file_path.get(),
                self.variant_number.get(),
                student_answers
            )
            
            # Створюємо звіти з результатами
            self.log_message("Створення звітів...")
            
            # Створюємо PDF звіт
            pdf_report_path = processor.create_check_result_pdf(check_result, self.output_folder_path.get())
            self.log_message(f"PDF звіт створено: {pdf_report_path}")
            
            # Створюємо Word звіт
            word_report_path = processor.create_check_result_word(check_result, self.output_folder_path.get())
            self.log_message(f"Word звіт створено: {word_report_path}")
            
            # Виведення результатів
            weighted_score = check_result.get('weighted_score', 0)
            max_score = check_result.get('max_score', 12)
            self.log_message(f"Результат перевірки:")
            self.log_message(f"- Варіант: {check_result['variant_number']}")
            self.log_message(f"- Всього питань: {check_result['total_questions']}")
            self.log_message(f"- Правильних відповідей: {check_result['correct_answers']}")
            self.log_message(f"- Відсоток: {check_result['score_percentage']:.1f}%")
            self.log_message(f"- Бали: {weighted_score:.2f} з {max_score}")
            
            self.status_var.set("Перевірка завершена")
            
            # Показуємо результат
            result_text = (
                f"Перевірка завершена!\n\n"
                f"Варіант: {check_result['variant_number']}\n"
                f"Правильних відповідей: {check_result['correct_answers']} з {check_result['total_questions']}\n"
                f"Відсоток: {check_result['score_percentage']:.1f}%\n"
                f"Бали: {weighted_score:.2f} з {max_score}\n\n"
                f"Створено звіти:\n"
                f"- PDF: {os.path.basename(pdf_report_path)}\n"
                f"- Word: {os.path.basename(word_report_path)}"
            )
            
            self.root.after(0, lambda: messagebox.showinfo("Результат перевірки", result_text))
            
        except Exception as e:
            error_msg = f"Помилка при перевірці відповідей: {str(e)}"
            self.log_message(error_msg)
            logger.error(error_msg, exc_info=True)
            self.status_var.set("Помилка перевірки")
            self.root.after(0, lambda: messagebox.showerror("Помилка", error_msg))

def main():
    """Головна функція додатку"""
    try:
        root = tk.Tk()
        app = TeacherTestApp(root)
        root.mainloop()
    except Exception as e:
        logger.error(f"Критична помилка додатку: {e}", exc_info=True)
        messagebox.showerror("Критична помилка", f"Сталася критична помилка: {e}")

if __name__ == "__main__":
    main()