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

class TeacherTestApp:
    def __init__(self, root):
        self.root = root
        self.root.title("TeacherTest - Генератор тестов для учеников")
        self.root.geometry("800x700")
        
        # Переменные для хранения путей
        self.excel_file_path = tk.StringVar()
        self.output_folder_path = tk.StringVar(value=get_downloads_folder())
        self.answer_key_file_path = tk.StringVar()
        
        # Переменные для настроек
        self.num_variants = tk.IntVar(value=10)
        self.variant_number = tk.IntVar(value=1)
        self.student_answers_text = tk.StringVar()
        
        # Режим работы (1 - генерация тестов, 2 - проверка работ)
        self.mode = tk.IntVar(value=1)
        
        self.setup_ui()
        
    def setup_ui(self):
        """Настройка пользовательского интерфейса"""
        # Главный фрейм
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Настройка растягивания
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Заголовок
        title_label = ttk.Label(main_frame, text="TeacherTest - Генератор тестов", font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Выбор режима работы
        mode_frame = ttk.LabelFrame(main_frame, text="Режим работы", padding="10")
        mode_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Radiobutton(mode_frame, text="Режим 1: Генерация тестов", variable=self.mode, value=1, command=self.on_mode_change).grid(row=0, column=0, sticky=tk.W)
        ttk.Radiobutton(mode_frame, text="Режим 2: Проверка работ", variable=self.mode, value=2, command=self.on_mode_change).grid(row=0, column=1, sticky=tk.W)
        
        # Фрейм для режима 1 (генерация тестов)
        self.mode1_frame = ttk.LabelFrame(main_frame, text="Генерация тестов", padding="10")
        self.mode1_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        self.mode1_frame.columnconfigure(1, weight=1)
        
        # Выбор Excel файла с вопросами
        ttk.Label(self.mode1_frame, text="Excel файл с вопросами:").grid(row=0, column=0, sticky=tk.W, pady=2)
        ttk.Entry(self.mode1_frame, textvariable=self.excel_file_path, width=50).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(5, 5), pady=2)
        ttk.Button(self.mode1_frame, text="Обзор", command=self.browse_excel_file).grid(row=0, column=2, pady=2)
        
        # Количество вариантов
        ttk.Label(self.mode1_frame, text="Количество вариантов:").grid(row=1, column=0, sticky=tk.W, pady=2)
        ttk.Spinbox(self.mode1_frame, from_=1, to=100, textvariable=self.num_variants, width=10).grid(row=1, column=1, sticky=tk.W, padx=(5, 0), pady=2)
        
        # Папка для сохранения
        ttk.Label(self.mode1_frame, text="Папка для сохранения:").grid(row=2, column=0, sticky=tk.W, pady=2)
        ttk.Entry(self.mode1_frame, textvariable=self.output_folder_path, width=50).grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(5, 5), pady=2)
        ttk.Button(self.mode1_frame, text="Обзор", command=self.browse_output_folder).grid(row=2, column=2, pady=2)
        
        # Кнопка генерации
        ttk.Button(self.mode1_frame, text="Сгенерировать тесты", command=self.generate_tests).grid(row=3, column=0, columnspan=3, pady=10)
        
        # Фрейм для режима 2 (проверка работ)
        self.mode2_frame = ttk.LabelFrame(main_frame, text="Проверка работ", padding="10")
        self.mode2_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        self.mode2_frame.columnconfigure(1, weight=1)
        
        # Файл-ключ с ответами
        ttk.Label(self.mode2_frame, text="Excel файл-ключ:").grid(row=0, column=0, sticky=tk.W, pady=2)
        ttk.Entry(self.mode2_frame, textvariable=self.answer_key_file_path, width=50).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(5, 5), pady=2)
        ttk.Button(self.mode2_frame, text="Обзор", command=self.browse_answer_key_file).grid(row=0, column=2, pady=2)
        
        # Номер варианта ученика
        ttk.Label(self.mode2_frame, text="Номер варианта ученика:").grid(row=1, column=0, sticky=tk.W, pady=2)
        ttk.Spinbox(self.mode2_frame, from_=1, to=100, textvariable=self.variant_number, width=10).grid(row=1, column=1, sticky=tk.W, padx=(5, 0), pady=2)
        
        # Ответы ученика
        ttk.Label(self.mode2_frame, text="Ответы ученика (через запятую):").grid(row=2, column=0, sticky=tk.W, pady=2)
        ttk.Entry(self.mode2_frame, textvariable=self.student_answers_text, width=50).grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(5, 5), pady=2)
        
        # Кнопка проверки
        ttk.Button(self.mode2_frame, text="Проверить работу", command=self.check_answers).grid(row=3, column=0, columnspan=3, pady=10)
        
        # Прогресс-бар
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 5))
        
        # Статус
        self.status_var = tk.StringVar(value="Готов к работе")
        self.status_label = ttk.Label(main_frame, textvariable=self.status_var)
        self.status_label.grid(row=5, column=0, columnspan=3, pady=5)
        
        # Лог
        log_frame = ttk.LabelFrame(main_frame, text="Лог операций", padding="5")
        log_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(6, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=10, width=80)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Изначально показываем только режим 1
        self.on_mode_change()
        
    def on_mode_change(self):
        """Обработчик смены режима работы"""
        if self.mode.get() == 1:
            self.mode1_frame.grid()
            self.mode2_frame.grid_remove()
        else:
            self.mode1_frame.grid_remove()
            self.mode2_frame.grid()
    
    def browse_excel_file(self):
        """Выбор Excel файла с вопросами"""
        filename = filedialog.askopenfilename(
            title="Выберите Excel файл с вопросами",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.excel_file_path.set(filename)
    
    def browse_answer_key_file(self):
        """Выбор Excel файла-ключа"""
        filename = filedialog.askopenfilename(
            title="Выберите Excel файл-ключ",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.answer_key_file_path.set(filename)
    
    def browse_output_folder(self):
        """Выбор папки для сохранения"""
        folder = filedialog.askdirectory(title="Выберите папку для сохранения")
        if folder:
            self.output_folder_path.set(folder)
    
    def log_message(self, message):
        """Добавление сообщения в лог"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def update_progress(self, current, total, extra_info=""):
        """Обновление прогресс-бара"""
        if total > 0:
            progress = (current / total) * 100
            self.progress_var.set(progress)
            status_text = f"Прогресс: {current}/{total}"
            if extra_info:
                status_text += f" - {extra_info}"
            self.status_var.set(status_text)
        self.root.update_idletasks()
    
    def generate_tests(self):
        """Генерация тестов"""
        # Проверка входных данных
        if not self.excel_file_path.get():
            messagebox.showerror("Ошибка", "Выберите Excel файл с вопросами")
            return
        
        if not os.path.exists(self.excel_file_path.get()):
            messagebox.showerror("Ошибка", "Выбранный Excel файл не существует")
            return
        
        if not self.output_folder_path.get():
            messagebox.showerror("Ошибка", "Выберите папку для сохранения")
            return
        
        if self.num_variants.get() < 1:
            messagebox.showerror("Ошибка", "Количество вариантов должно быть больше 0")
            return
        
        # Запуск генерации в отдельном потоке
        thread = threading.Thread(target=self._generate_tests_thread)
        thread.daemon = True
        thread.start()
    
    def _generate_tests_thread(self):
        """Поток для генерации тестов"""
        try:
            self.log_message("Начало генерации тестов...")
            self.status_var.set("Генерация тестов...")
            
            # Чтение Excel файла
            self.log_message(f"Чтение файла: {self.excel_file_path.get()}")
            df = processor.read_test_excel(self.excel_file_path.get())
            
            if df.empty:
                raise ValueError("Excel файл не содержит данных")
            
            self.log_message(f"Найдено {len(df)} вопросов")
            
            # Генерация вариантов
            self.log_message(f"Генерация {self.num_variants.get()} вариантов...")
            variants = processor.generate_test_variants(df, self.num_variants.get())
            
            # PDF файлы отключены - используем только Word и Excel
            
            # Создание Excel файла-ключа
            self.log_message("Создание Excel файла-ключа...")
            input_file_name = os.path.splitext(os.path.basename(self.excel_file_path.get()))[0]
            excel_key_path = processor.create_excel_answer_key(variants, self.output_folder_path.get(), input_file_name)
            
            self.log_message(f"Готово! Созданы файлы:")
            self.log_message(f"- Excel ключ: {excel_key_path}")
            
            self.status_var.set("Генерация завершена успешно")
            self.progress_var.set(100)
            
            # Показываем сообщение об успехе
            self.root.after(0, lambda: messagebox.showinfo(
                "Успех", 
                f"Тесты успешно сгенерированы!\n\n"
                f"Файлы сохранены в: {self.output_folder_path.get()}"
            ))
            
        except Exception as e:
            error_msg = f"Ошибка при генерации тестов: {str(e)}"
            self.log_message(error_msg)
            logger.error(error_msg, exc_info=True)
            self.status_var.set("Ошибка генерации")
            self.root.after(0, lambda: messagebox.showerror("Ошибка", error_msg))
    
    def check_answers(self):
        """Проверка ответов ученика"""
        # Проверка входных данных
        if not self.answer_key_file_path.get():
            messagebox.showerror("Ошибка", "Выберите Excel файл-ключ")
            return
        
        if not os.path.exists(self.answer_key_file_path.get()):
            messagebox.showerror("Ошибка", "Выбранный файл-ключ не существует")
            return
        
        if self.variant_number.get() < 1:
            messagebox.showerror("Ошибка", "Номер варианта должен быть больше 0")
            return
        
        if not self.student_answers_text.get().strip():
            messagebox.showerror("Ошибка", "Введите ответы ученика")
            return
        
        # Запуск проверки в отдельном потоке
        thread = threading.Thread(target=self._check_answers_thread)
        thread.daemon = True
        thread.start()
    
    def _check_answers_thread(self):
        """Поток для проверки ответов"""
        try:
            self.log_message("Начало проверки ответов...")
            self.status_var.set("Проверка ответов...")
            
            # Парсинг ответов ученика
            answers_text = self.student_answers_text.get().strip()
            try:
                student_answers = [int(x.strip()) for x in answers_text.split(',')]
            except ValueError:
                raise ValueError("Ответы должны быть числами, разделенными запятыми")
            
            self.log_message(f"Ответы ученика: {student_answers}")
            
            # Проверка ответов
            check_result = processor.check_student_answers(
                self.answer_key_file_path.get(),
                self.variant_number.get(),
                student_answers
            )
            
            # Создаем отчеты с результатами
            self.log_message("Создание отчетов...")
            
            # Создаем PDF отчет
            pdf_report_path = processor.create_check_result_pdf(check_result, self.output_folder_path.get())
            self.log_message(f"PDF отчет создан: {pdf_report_path}")
            
            # Создаем Word отчет
            word_report_path = processor.create_check_result_word(check_result, self.output_folder_path.get())
            self.log_message(f"Word отчет создан: {word_report_path}")
            
            # Вывод результатов
            self.log_message(f"Результат проверки:")
            self.log_message(f"- Вариант: {check_result['variant_number']}")
            self.log_message(f"- Всего вопросов: {check_result['total_questions']}")
            self.log_message(f"- Правильных ответов: {check_result['correct_answers']}")
            self.log_message(f"- Процент: {check_result['score_percentage']:.1f}%")
            
            self.status_var.set("Проверка завершена")
            
            # Показываем результат
            result_text = (
                f"Проверка завершена!\n\n"
                f"Вариант: {check_result['variant_number']}\n"
                f"Правильных ответов: {check_result['correct_answers']} из {check_result['total_questions']}\n"
                f"Процент: {check_result['score_percentage']:.1f}%\n\n"
                f"Созданы отчеты:\n"
                f"- PDF: {os.path.basename(pdf_report_path)}\n"
                f"- Word: {os.path.basename(word_report_path)}"
            )
            
            self.root.after(0, lambda: messagebox.showinfo("Результат проверки", result_text))
            
        except Exception as e:
            error_msg = f"Ошибка при проверке ответов: {str(e)}"
            self.log_message(error_msg)
            logger.error(error_msg, exc_info=True)
            self.status_var.set("Ошибка проверки")
            self.root.after(0, lambda: messagebox.showerror("Ошибка", error_msg))

def main():
    """Главная функция приложения"""
    try:
        root = tk.Tk()
        app = TeacherTestApp(root)
        root.mainloop()
    except Exception as e:
        logger.error(f"Критическая ошибка приложения: {e}", exc_info=True)
        messagebox.showerror("Критическая ошибка", f"Произошла критическая ошибка: {e}")

if __name__ == "__main__":
    main()