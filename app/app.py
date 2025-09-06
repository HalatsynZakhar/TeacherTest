import streamlit as st
import os
import sys
import logging
import io
import time
import tempfile
import shutil
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, List, Optional, Tuple
import pandas as pd
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
import json
import traceback
# from fpdf import FPDF  # Отключено - используем только Word и Excel

# Добавляем корневую папку проекта в PYTHONPATH
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

# Используем относительные импорты
from utils import config_manager
from utils import excel_utils
from utils.config_manager import get_downloads_folder, ConfigManager
from core.processor import (
    read_test_excel, 
    generate_test_variants, 
    create_excel_answer_key,
    check_student_answers,
    create_check_result_pdf,
    create_check_result_word,
    create_test_word,
    read_test_word,
    export_answers_to_word,
    generate_test_template,
    generate_neural_query_document,
    save_student_result_to_excel
)

# Настройка логирования
log_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'logs')
os.makedirs(log_dir, exist_ok=True)

# Обмежуємо кількість файлів логів до 5 останніх
log_files = sorted([f for f in os.listdir(log_dir) if f.startswith('teacher_test_')])
if len(log_files) > 5:
    for old_log in log_files[:-5]:
        try:
            os.remove(os.path.join(log_dir, old_log))
        except:
            pass

# Створюємо новий лог-файл
log_file = os.path.join(log_dir, 'teacher_test_latest.log')
try:
    with open(log_file, 'w', encoding='utf-8') as f:
        f.write(f'{datetime.now().strftime("%Y-%m-%d %H:%M:%S")} - INFO - teacher_test - New log file created with UTF-8 encoding\n')
except Exception as e:
    print(f"Error creating log file: {e}")

log_stream = io.StringIO()
log_handler = logging.StreamHandler(log_stream)
log_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(name)s - %(message)s'))
log_handler.setLevel(logging.INFO)

file_handler = logging.FileHandler(log_file, encoding='utf-8')
file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(name)s - %(message)s'))
file_handler.setLevel(logging.DEBUG)

root_logger = logging.getLogger()
root_logger.setLevel(logging.DEBUG)
for handler in root_logger.handlers[:]:
    root_logger.removeHandler(handler)
root_logger.addHandler(log_handler)
root_logger.addHandler(file_handler)

# Устанавливаем кодировку для логирования
sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')

log = logging.getLogger(__name__)

# Визначаємо налаштування за замовчуванням для TeacherTest
default_settings = {
    "test_settings": {
        "default_variants_count": 10,
        "shuffle_questions": True,
        "shuffle_answers": True
    },
    "file_settings": {
        "max_size_mb": 50
    },
    "user_paths": {
        "results_excel_path": "",
        "results_excel_filename": "results.xlsx",
        "last_answer_key_file": "",
        "save_results_path": ""
    }
}

# Инициализация менеджера конфигурации
def init_config_manager():
    """Инициализировать менеджер конфигурации и установить значения по умолчанию"""
    if 'config_manager' not in st.session_state:
        presets_folder = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'settings_presets')
        config_manager_instance = config_manager.ConfigManager(presets_folder)
        config_manager_instance.load_settings()
        
        # Устанавливаем значения по умолчанию для TeacherTest
        if not config_manager_instance.get_setting('test_settings.default_variants_count'):
            config_manager_instance.set_setting('test_settings.default_variants_count', default_settings['test_settings']['default_variants_count'])
        
        if config_manager_instance.get_setting('test_settings.shuffle_questions') is None:
            config_manager_instance.set_setting('test_settings.shuffle_questions', default_settings['test_settings']['shuffle_questions'])
        
        if config_manager_instance.get_setting('test_settings.shuffle_answers') is None:
            config_manager_instance.set_setting('test_settings.shuffle_answers', default_settings['test_settings']['shuffle_answers'])
        
        if not config_manager_instance.get_setting('file_settings.max_size_mb'):
            config_manager_instance.set_setting('file_settings.max_size_mb', default_settings['file_settings']['max_size_mb'])
        
        # Ініціалізуємо налаштування шляхів користувача
        if not config_manager_instance.get_setting('user_paths.results_excel_filename'):
            config_manager_instance.set_setting('user_paths.results_excel_filename', default_settings['user_paths']['results_excel_filename'])
        
        if not config_manager_instance.get_setting('user_paths.results_excel_path'):
            config_manager_instance.set_setting('user_paths.results_excel_path', get_downloads_folder())
        
        if not config_manager_instance.get_setting('user_paths.last_answer_key_file'):
            config_manager_instance.set_setting('user_paths.last_answer_key_file', default_settings['user_paths']['last_answer_key_file'])
        
        if not config_manager_instance.get_setting('user_paths.save_results_path'):
            config_manager_instance.set_setting('user_paths.save_results_path', get_downloads_folder())
        
        config_manager_instance.save_settings()
        st.session_state.config_manager = config_manager_instance
        log.info("Менеджер конфігурації TeacherTest ініціалізовано")
    
    return st.session_state.config_manager

# Инициализация
config_folder = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'settings_presets')
config_manager.init_config_manager(config_folder)
init_config_manager()
cm = st.session_state.config_manager

# Настройка параметров приложения
st.set_page_config(
    page_title="TeacherTest - Генератор тестів",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Функція для створення тимчасових директорій
def ensure_temp_dir(prefix: str = "") -> str:
    """Создает и возвращает путь к временной директории"""
    project_dir = os.path.dirname(os.path.dirname(__file__))
    temp_dir = os.path.join(project_dir, "temp")
    
    try:
        os.makedirs(temp_dir, exist_ok=True)
        log.info(f"Створена/перевірена тимчасова директорія: {temp_dir}")
    except Exception as e:
        log.error(f"Помилка при створенні тимчасової директорії {temp_dir}: {e}")
        temp_dir = os.path.join(tempfile.gettempdir(), f"{prefix}teachertest")
        try:
            os.makedirs(temp_dir, exist_ok=True)
        except Exception as e2:
            log.error(f"Помилка при створенні системної тимчасової директорії: {e2}")
            raise e2
    
    return temp_dir

# Функція очищення тимчасових файлів
def cleanup_temp_files():
    """Очищення тимчасових файлів при запуску додатка"""
    try:
        project_dir = os.path.dirname(os.path.dirname(__file__))
        temp_dir = os.path.join(project_dir, "temp")
        
        if os.path.exists(temp_dir):
            for filename in os.listdir(temp_dir):
                file_path = os.path.join(temp_dir, filename)
                try:
                    if os.path.isfile(file_path):
                        # Перевіряємо вік файла (видаляємо файли старше 1 дня)
                        file_age = time.time() - os.path.getmtime(file_path)
                        if file_age > 86400:  # 24 часа в секундах
                            os.remove(file_path)
                            log.info(f"Видалено старий тимчасовий файл: {filename}")
                except Exception as e:
                    log.warning(f"Не вдалося видалити тимчасовий файл {filename}: {e}")
    except Exception as e:
        log.error(f"Помилка при очищенні тимчасових файлів: {e}")

# Очистка при запуске
cleanup_temp_files()

# Инициализация session state
if 'log_messages' not in st.session_state:
    st.session_state.log_messages = []
if 'df' not in st.session_state:
    st.session_state.df = None
if 'temp_file_path' not in st.session_state:
    st.session_state.temp_file_path = None
if 'processing_error' not in st.session_state:
    st.session_state.processing_error = None
if 'output_files' not in st.session_state:
    st.session_state.output_files = {}
if 'mode' not in st.session_state:
    st.session_state.mode = 1  # 1 - генерація тестів, 2 - перевірка робіт
if 'variants_count' not in st.session_state:
    st.session_state.variants_count = cm.get_setting('test_settings.default_variants_count', 10)
if 'answer_key_file' not in st.session_state:
    st.session_state.answer_key_file = None
if 'variant_number' not in st.session_state:
    st.session_state.variant_number = 1
if 'student_answers' not in st.session_state:
    st.session_state.student_answers = ""
if 'answer_format' not in st.session_state:
    st.session_state.answer_format = 'list'  # 'list' или 'table'
if 'space_optimization' not in st.session_state:
    st.session_state.space_optimization = False  # Оптимизация места
if 'student_class' not in st.session_state:
    st.session_state.student_class = ""
if 'student_full_name' not in st.session_state:
    st.session_state.student_full_name = ""
if 'input_file_name' not in st.session_state:
    st.session_state.input_file_name = ""
if 'test_class' not in st.session_state:
    st.session_state.test_class = ""
if 'test_date' not in st.session_state:
    st.session_state.test_date = ""
if 'question_shuffle_mode' not in st.session_state:
    st.session_state.question_shuffle_mode = 'full'  # 'full', 'easy_to_hard', 'none'
if 'answer_shuffle_mode' not in st.session_state:
    st.session_state.answer_shuffle_mode = 'random'  # 'random', 'none'
if 'last_error' not in st.session_state:
    st.session_state.last_error = None
if 'test_work_name' not in st.session_state:
    st.session_state.test_work_name = ""
if 'results_excel_path' not in st.session_state:
    # Завантажуємо збережений шлях або встановлюємо папку Завантаження як значення за замовчуванням
    saved_path = cm.get_setting('user_paths.results_excel_path')
    st.session_state.results_excel_path = saved_path if saved_path else get_downloads_folder()
if 'results_excel_filename' not in st.session_state:
    # Завантажуємо збережену назву файлу
    saved_filename = cm.get_setting('user_paths.results_excel_filename')
    st.session_state.results_excel_filename = saved_filename if saved_filename else "results.xlsx"
if 'save_results_path' not in st.session_state:
    # Завантажуємо збережений шлях для збереження результатів або використовуємо за замовчуванням
    saved_save_path = cm.get_setting('user_paths.save_results_path')
    st.session_state.save_results_path = saved_save_path if saved_save_path else get_downloads_folder()
if 'save_tests_path' not in st.session_state:
    # Завантажуємо збережений шлях для збереження згенерованих тестів або використовуємо за замовчуванням
    saved_tests_path = cm.get_setting('user_paths.save_tests_path')
    st.session_state.save_tests_path = saved_tests_path if saved_tests_path else get_downloads_folder()

def save_user_settings():
    """Збереження налаштувань користувача в конфігурацію"""
    try:
        cm.set_setting('user_paths.results_excel_path', st.session_state.results_excel_path)
        cm.set_setting('user_paths.results_excel_filename', st.session_state.results_excel_filename)
        cm.set_setting('user_paths.save_results_path', st.session_state.save_results_path)
        cm.set_setting('user_paths.save_tests_path', st.session_state.save_tests_path)
        if hasattr(st.session_state, 'answer_key_file') and st.session_state.answer_key_file:
            cm.set_setting('user_paths.last_answer_key_file', st.session_state.answer_key_file)
        cm.save_settings()
        log.info("Налаштування користувача збережено")
    except Exception as e:
        log.error(f"Помилка при збереженні налаштувань: {e}")

def create_custom_filename(work_name, student_class, student_full_name, variant, key_file_name, timestamp, extension):
    """Створює кастомну назву файлу з назвою роботи"""
    parts = []
    
    # Додаємо назву роботи
    if work_name:
        parts.append(work_name.replace(' ', '_').replace('/', '_').replace('\\', '_'))
    
    # Додаємо клас (якщо є)
    if student_class:
        parts.append(student_class.replace(' ', '_').replace('-', '_'))
    
    # Додаємо ім'я (якщо є)
    if student_full_name:
        parts.append(student_full_name.replace(' ', '_').replace('/', '_').replace('\\', '_'))
    
    # Додаємо варіант
    parts.append(f"Варіант{variant}")
    
    # Додаємо timestamp
    parts.append(timestamp)
    
    return f"{'_'.join(parts)}.{extension}"

def save_all_results():
    """Зберігає всі результати (PDF, Word, Excel) в указану папку"""
    try:
        if not hasattr(st.session_state, 'check_result') or not st.session_state.check_result:
            st.error("❌ Немає результатів для збереження. Спочатку перевірте роботу.")
            return False
        
        save_path = st.session_state.save_results_path
        if not save_path or not os.path.exists(save_path):
            st.error("❌ Вказаний шлях для збереження не існує.")
            return False
        
        result = st.session_state.check_result
        work_name = st.session_state.test_work_name
        key_file_name = os.path.basename(st.session_state.answer_key_file) if st.session_state.answer_key_file else ""
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
        
        # Створюємо кастомні назви файлів
        pdf_filename = create_custom_filename(
            work_name, st.session_state.student_class, st.session_state.student_full_name,
            result['variant_number'], key_file_name, timestamp, "pdf"
        )
        word_filename = create_custom_filename(
            work_name, st.session_state.student_class, st.session_state.student_full_name,
            result['variant_number'], key_file_name, timestamp, "docx"
        )
        
        # Зберігаємо PDF з кастомною назвою
        from core.processor import create_check_result_pdf
        import tempfile
        import shutil
        
        # Створюємо PDF у тимчасовій папці
        temp_pdf_path = create_check_result_pdf(result, tempfile.gettempdir())
        pdf_path = os.path.join(save_path, pdf_filename)
        shutil.move(temp_pdf_path, pdf_path)
        pdf_success = True
        
        # Зберігаємо Word з кастомною назвою
        from core.processor import create_check_result_word
        temp_word_path = create_check_result_word(result, tempfile.gettempdir())
        word_path = os.path.join(save_path, word_filename)
        shutil.move(temp_word_path, word_path)
        word_success = True
        
        # Зберігаємо в Excel
        excel_success = save_student_result_to_excel()
        excel_path = os.path.join(st.session_state.results_excel_path, st.session_state.results_excel_filename)
        
        # Перевіряємо результати
        success_count = sum([pdf_success, word_success, excel_success])
        
        if success_count == 3:
            work_name = st.session_state.test_work_name if st.session_state.test_work_name else "Тест"
            key_file_name = os.path.basename(st.session_state.answer_key_file) if st.session_state.answer_key_file else "Невідомий файл"
            st.success(f"✅ {work_name} - Всі результати збережено успішно!\n📁 PDF: {pdf_path}\n📄 Word: {word_path}\n📊 Excel: {excel_path}\n🔑 Файл-ключ: {key_file_name}")
            save_user_settings()
            return True
        elif success_count > 0:
            st.warning(f"⚠️ Частково збережено ({success_count}/3 файлів)")
            return True
        else:
            st.error("❌ Помилка при збереженні результатів")
            return False
            
    except Exception as e:
        st.error(f"❌ Помилка при збереженні всіх результатів: {str(e)}")
        add_log_message(f"Помилка при збереженні всіх результатів: {str(e)}", "ERROR")
        return False

def save_all_tests():
    """Збереження всіх згенерованих тестів в одну папку"""
    try:
        if not st.session_state.save_tests_path:
            st.error("Не вказано шлях для збереження")
            return
        
        # Створюємо папку якщо її немає
        os.makedirs(st.session_state.save_tests_path, exist_ok=True)
        
        saved_files = []
        
        # Зберігаємо Word файл з тестами
        if 'test_word' in st.session_state.output_files and os.path.exists(st.session_state.output_files['test_word']):
            filename = os.path.basename(st.session_state.output_files['test_word'])
            final_path = os.path.join(st.session_state.save_tests_path, filename)
            shutil.copy2(st.session_state.output_files['test_word'], final_path)
            saved_files.append(f"Тести: {filename}")
        
        # Зберігаємо Excel ключ
        if 'excel_key' in st.session_state.output_files and os.path.exists(st.session_state.output_files['excel_key']):
            filename = os.path.basename(st.session_state.output_files['excel_key'])
            final_path = os.path.join(st.session_state.save_tests_path, filename)
            shutil.copy2(st.session_state.output_files['excel_key'], final_path)
            saved_files.append(f"Ключ: {filename}")
        
        # Зберігаємо Word файл з відповідями
        if 'answers_word' in st.session_state.output_files and os.path.exists(st.session_state.output_files['answers_word']):
            filename = os.path.basename(st.session_state.output_files['answers_word'])
            final_path = os.path.join(st.session_state.save_tests_path, filename)
            shutil.copy2(st.session_state.output_files['answers_word'], final_path)
            saved_files.append(f"Відповіді: {filename}")
        
        if saved_files:
            st.success(f"✅ Всі файли тестів збережено успішно в папку: {st.session_state.save_tests_path}\n\n" +
                      "\n".join(saved_files))
            save_user_settings()
        else:
            st.warning("⚠️ Немає файлів для збереження")
            
    except Exception as e:
        st.error(f"❌ Помилка при збереженні файлів: {e}")
        add_log_message(f"Помилка збереження всіх тестів: {e}", "ERROR")

def add_log_message(message, level="INFO"):
    """Добавление сообщения в лог"""
    timestamp = datetime.now().strftime("%H:%M:%S")
    log_entry = f"[{timestamp}] {level}: {message}"
    
    if 'log_messages' not in st.session_state:
        st.session_state.log_messages = []
    
    st.session_state.log_messages.append(log_entry)
    
    # Обмежуємо кількість повідомлень у лозі
    if len(st.session_state.log_messages) > 100:
        st.session_state.log_messages = st.session_state.log_messages[-100:]
    
    # Логуємо у файл
    if level == "ERROR":
        log.error(message)
    elif level == "WARNING":
        log.warning(message)
    else:
        log.info(message)

def load_file(uploaded_file_arg=None):
    """Завантаження та обробка Excel або Word файла з питаннями"""
    try:
        if uploaded_file_arg is None:
            uploaded_file_arg = st.session_state.get('temp_file_path')
        
        if not uploaded_file_arg or not os.path.exists(uploaded_file_arg):
            st.session_state.df = None
            st.session_state.processing_error = "Файл не знайдено"
            return
        
        add_log_message(f"Завантаження файла: {os.path.basename(uploaded_file_arg)}")
        
        # Зберігаємо ім'я файла без розширення для використання в іменах вихідних файлів
        st.session_state.input_file_name = os.path.splitext(os.path.basename(uploaded_file_arg))[0]
        
        # Визначаємо тип файла за розширенням
        file_extension = os.path.splitext(uploaded_file_arg)[1].lower()
        
        if file_extension in ['.xlsx', '.xls']:
            # Читаємо Excel файл з питаннями
            df = read_test_excel(uploaded_file_arg)
        elif file_extension in ['.docx', '.doc']:
            # Читаємо Word файл з питаннями
            df = read_test_word(uploaded_file_arg)
        else:
            st.session_state.df = None
            st.session_state.processing_error = "Непідтримуваний формат файла. Використовуйте Excel (.xlsx, .xls) або Word (.docx, .doc)"
            add_log_message("Непідтримуваний формат файла", "ERROR")
            return
        
        if df.empty:
            st.session_state.df = None
            st.session_state.processing_error = "Файл не містить даних або має неправильну структуру"
            add_log_message("Файл порожній або має неправильну структуру", "ERROR")
            return
        
        st.session_state.df = df
        st.session_state.processing_error = None
        add_log_message(f"Файл успішно завантажено. Знайдено {len(df)} питань", "SUCCESS")
        
    except Exception as e:
        error_msg = f"Помилка при завантаженні файла: {str(e)}"
        st.session_state.processing_error = error_msg
        st.session_state.df = None
        add_log_message(error_msg, "ERROR")
        log.error(f"Помилка при завантаженні файла: {e}", exc_info=True)

def generate_tests():
    """Генерація тестів"""
    try:
        # Очищуємо попередні помилки
        st.session_state.last_error = None
        
        add_log_message("Початок генерації тестів")
        
        if st.session_state.df is None or st.session_state.df.empty:
            raise ValueError("Нет данных для генерации тестов")
        
        # Генерируем варианты тестов
        variants = generate_test_variants(
            st.session_state.df, 
            st.session_state.variants_count,
            question_shuffle_mode=st.session_state.question_shuffle_mode,
            answer_shuffle_mode=st.session_state.answer_shuffle_mode
        )
        add_log_message(f"Згенеровано {len(variants)} варіантів тестів")
        
        # Створюємо тимчасову папку для вихідних файлів
        output_dir = ensure_temp_dir("output_")
        
        # PDF файли відключені - використовуємо тільки Word та Excel
        
        # Створюємо Excel файл-ключ
        excel_key_path = create_excel_answer_key(variants, output_dir, st.session_state.input_file_name)
        add_log_message(f"Створено Excel файл-ключ")
        
        # Створюємо Word файл з тестами
        test_word_path = create_test_word(variants, output_dir, 1, st.session_state.input_file_name, st.session_state.answer_format, st.session_state.space_optimization, st.session_state.test_class, st.session_state.test_date)
        add_log_message(f"Створено Word файл з тестами")
        
        # Створюємо Word файл з відповідями
        answers_word_path = export_answers_to_word(variants, output_dir, st.session_state.input_file_name, st.session_state.test_class, st.session_state.test_date)
        add_log_message(f"Створено Word файл з відповідями")
        
        # Зберігаємо шляхи до файлів
        st.session_state.output_files = {
            'excel_key': excel_key_path,
            'test_word': test_word_path,
            'answers_word': answers_word_path
        }
        
        add_log_message("Генерація тестів завершена успішно", "SUCCESS")
        return True
        
    except Exception as e:
        error_msg = f"Помилка при генерації тестів: {str(e)}"
        st.session_state.last_error = error_msg
        add_log_message(error_msg, "ERROR")
        log.error(error_msg, exc_info=True)
        return False

def check_answers():
    """Перевірка відповідей учня"""
    try:
        # Очищуємо попередні помилки
        st.session_state.last_error = None
        
        add_log_message("Початок перевірки відповідей")
        
        if not st.session_state.answer_key_file:
            raise ValueError("Не обрано файл-ключ")
        
        # Парсимо відповіді учня
        try:
            # Розділяємо по комах, зберігаючи порожні відповіді
            raw_answers = st.session_state.student_answers.split(',')
            student_answers = []
            
            for i, answer in enumerate(raw_answers):
                answer = answer.strip()
                if not answer:  # Порожня відповідь
                    student_answers.append("")
                else:
                    student_answers.append(answer)
                    
        except ValueError:
            raise ValueError("Відповіді повинні бути розділені комами")
        
        # Перевіряємо відповіді
        check_result = check_student_answers(
            st.session_state.answer_key_file,
            st.session_state.variant_number,
            student_answers
        )
        
        # Добавляем данные ученика в результат
        student_info = {
            'class': st.session_state.student_class.strip(),
            'full_name': st.session_state.student_full_name.strip()
        }
        check_result['student_info'] = student_info
        
        # Створюємо звіти з результатами
        output_dir = ensure_temp_dir("reports_")
        
        # Створюємо PDF звіт
        pdf_report_path = create_check_result_pdf(check_result, output_dir)
        
        # Створюємо Word звіт
        word_report_path = create_check_result_word(check_result, output_dir)
        
        # Зберігаємо результат та шляхи до звітів
        st.session_state.check_result = check_result
        st.session_state.check_reports = {
            'pdf_report': pdf_report_path,
            'word_report': word_report_path
        }
        
        add_log_message(f"Перевірка завершена. Правильних відповідей: {check_result['correct_answers']} з {check_result['total_questions']}", "SUCCESS")
        return True
        
    except Exception as e:
        error_msg = f"Помилка при перевірці відповідей: {str(e)}"
        st.session_state.last_error = error_msg
        add_log_message(error_msg, "ERROR")
        log.error(error_msg, exc_info=True)
        return False

def save_student_result_to_excel():
    """Збереження результату учня у Excel файл"""
    try:
        if not hasattr(st.session_state, 'check_result') or not st.session_state.check_result:
            st.session_state.last_error = "Немає результатів для збереження"
            return False
        
        # Формуємо повний шлях до файлу
        full_path = os.path.join(st.session_state.results_excel_path, st.session_state.results_excel_filename)
        
        # Підготовуємо інформацію про учня
        student_info = {
            'class': st.session_state.student_class,
            'full_name': st.session_state.student_full_name
        }
        
        # Отримуємо назву файлу-ключа
        key_file_name = ""
        if hasattr(st.session_state, 'answer_key_file') and st.session_state.answer_key_file:
            key_file_name = os.path.basename(st.session_state.answer_key_file)
        
        # Викликаємо функцію збереження з processor.py
        from core.processor import save_student_result_to_excel as save_result_func
        save_result_func(
            check_result=st.session_state.check_result,
            student_info=student_info,
            work_name=st.session_state.test_work_name,
            excel_file_path=full_path,
            key_file_name=key_file_name
        )
        
        success = True
        
        if success:
            add_log_message(f"Результат збережено у файл {full_path}", "SUCCESS")
            # Зберігаємо налаштування користувача
            save_user_settings()
            return True
        else:
            st.session_state.last_error = "Помилка при збереженні результату"
            return False
            
    except Exception as e:
        error_msg = f"Помилка при збереженні: {str(e)}"
        st.session_state.last_error = error_msg
        add_log_message(error_msg, "ERROR")
        log.error(f"Error in save_student_result_to_excel: {e}", exc_info=True)
        return False

# Основной интерфейс
def main():
    st.title("📝 TeacherTest - Генератор тестів для учнів")
    
    # Секция скачивания шаблонов
    st.subheader("📋 Шаблони для роботи")
    col1, col2 = st.columns([1, 1])
    
    with col1:
        try:
            temp_dir = ensure_temp_dir("templates_")
            test_template_path = generate_test_template(temp_dir)
            
            with open(test_template_path, 'rb') as file:
                st.download_button(
                     label="📥 Завантажити шаблон тесту",
                     data=file.read(),
                     file_name="Шаблон_тесту.xlsx",
                     mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     help="Завантажити Excel шаблон для створення тестів"
                 )
            add_log_message("Шаблон тесту згенеровано успішно")
        except Exception as e:
            st.error(f"Помилка при створенні шаблону тесту: {e}")
            add_log_message(f"Помилка створення шаблону тесту: {e}", "ERROR")
    
    with col2:
        try:
            temp_dir = ensure_temp_dir("templates_")
            query_doc_path = generate_neural_query_document(temp_dir)
            
            with open(query_doc_path, 'rb') as file:
                st.download_button(
                     label="🤖 Запит для нейромережі",
                     data=file.read(),
                     file_name="Запит_для_нейромережі.docx",
                     mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                     help="Завантажити документ з запитом для генерації тестів через ШІ"
                 )
            add_log_message("Документ з запитом для нейромережі згенеровано успішно")
        except Exception as e:
            st.error(f"Помилка при створенні документа: {e}")
            add_log_message(f"Помилка створення документа: {e}", "ERROR")
    
    st.info("💡 **Підказка:** Завантажте шаблон для правильного форматування ваших тестів. Шаблон містить детальні інструкції по заповненню. Використовуйте запит для нейромережі, щоб швидко згенерувати тести через ШІ.")
     
    st.markdown("---")
    
    # Боковая панель с настройками
    with st.sidebar:
        st.header("⚙️ Налаштування")
        
        # Вибір режиму роботи
        st.session_state.mode = st.radio(
            "Режим роботи:",
            [1, 2],
            format_func=lambda x: "Генерація тестів" if x == 1 else "Перевірка робіт",
            index=st.session_state.mode - 1
        )
        
        st.markdown("---")
        
        if st.session_state.mode == 1:
            st.subheader("Налаштування генерації")
            st.session_state.variants_count = st.number_input(
                "Кількість варіантів:",
                min_value=1,
                max_value=100,
                value=st.session_state.variants_count
            )
            
            st.session_state.answer_format = st.selectbox(
                "Формат варіантів відповідей:",
                options=['list', 'table'],
                format_func=lambda x: 'Список' if x == 'list' else 'Таблиця',
                index=0 if st.session_state.answer_format == 'list' else 1,
                help="Виберіть формат відображення варіантів відповідей: список або таблиця"
            )
            
            st.session_state.space_optimization = st.checkbox(
                "Оптимізація місця",
                value=st.session_state.space_optimization,
                help="Мінімізує кількість переводів рядків для економії місця (може погіршити читабельність)"
            )
            
            # Налаштування перемішування
            st.subheader("Налаштування перемішування")
            
            st.session_state.question_shuffle_mode = st.selectbox(
                "Перемішування питань:",
                options=['full', 'easy_to_hard', 'none'],
                format_func=lambda x: {
                    'full': 'Повне перемішування',
                    'easy_to_hard': 'Від легкого до складного',
                    'none': 'Не перемішувати'
                }[x],
                index=['full', 'easy_to_hard', 'none'].index(st.session_state.question_shuffle_mode),
                help="Виберіть спосіб упорядкування питань у тесті"
            )
            
            st.session_state.answer_shuffle_mode = st.selectbox(
                "Перемішування варіантів:",
                options=['random', 'none'],
                format_func=lambda x: {
                    'random': 'Випадкове',
                    'none': 'Ні'
                }[x],
                index=['random', 'none'].index(st.session_state.answer_shuffle_mode),
                help="Виберіть спосіб перемішування варіантів відповідей"
            )
            
            # Дополнительные поля для заголовка теста
            st.subheader("Додаткова інформація")
            st.session_state.test_class = st.text_input(
                "Клас (опціонально):",
                value=st.session_state.test_class,
                help="Вкажіть клас для відображення в заголовку тесту"
            )
            
            st.session_state.test_date = st.text_input(
                "Дата тесту (опціонально):",
                value=st.session_state.test_date,
                help="Вкажіть дату тесту для відображення в заголовку"
            )
        else:
            st.subheader("Налаштування перевірки")
            st.session_state.variant_number = st.number_input(
                "Номер варіанта учня:",
                min_value=1,
                max_value=100,
                value=st.session_state.variant_number
            )
    
    # Основной контент
    if st.session_state.mode == 1:
        # Режим 1: Генерація тестів
        st.header("🎯 Режим 1: Генерація тестів")
        
        # Завантаження Excel або Word файла з питаннями
        uploaded_file = st.file_uploader(
            "Оберіть Excel або Word файл з питаннями",
            type=["xlsx", "xls", "docx", "doc"],
            help="Excel файл повинен містити: стовпець з питаннями, стовпець з номером правильної відповіді, стовпці з варіантами відповідей. Word файл повинен містити питання у форматі: '1. Питання' з варіантами відповідей '1) Варіант'"
        )
        
        if uploaded_file is not None:
            # Зберігаємо файл у тимчасову папку
            temp_dir = ensure_temp_dir()
            temp_file_path = os.path.join(temp_dir, uploaded_file.name)
            
            # Перевіряємо, чи потрібно оновити файл
            need_update = (
                not st.session_state.temp_file_path or 
                not os.path.exists(st.session_state.temp_file_path) or
                os.path.basename(st.session_state.temp_file_path) != uploaded_file.name or
                (os.path.exists(st.session_state.temp_file_path) and 
                 os.path.getsize(st.session_state.temp_file_path) != uploaded_file.size)
            )
            
            if need_update:
                with open(temp_file_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())
                st.session_state.temp_file_path = temp_file_path
                load_file()
            
            # Відображення інформації про файл
            if st.session_state.df is not None:
                st.success(f"✅ Файл завантажено: {uploaded_file.name}")
                st.info(f"📊 Знайдено питань: {len(st.session_state.df)}")
                
                # Предпросмотр данных
                with st.expander("👀 Попередній перегляд даних"):
                    # Создаем копию DataFrame для отображения и преобразуем object столбцы в строки
                    display_df = st.session_state.df.head(10).copy()
                    for col in display_df.select_dtypes(include=['object']).columns:
                        display_df[col] = display_df[col].astype(str)
                    # Дополнительно конвертируем option_ колонки в строки для совместимости с pyarrow
                    option_cols = [col for col in display_df.columns if col.startswith('option_')]
                    for col in option_cols:
                        display_df[col] = display_df[col].astype(str)
                    st.dataframe(display_df, use_container_width=True)
                
                # Кнопка генерации
                col1, col2, col3 = st.columns([1, 2, 1])
                with col2:
                    if st.button("🚀 Згенерувати тести", type="primary", use_container_width=True):
                        with st.spinner("Генерація тестів..."):
                            success = generate_tests()
                        
                        if success:
                            st.success("✅ Тести успішно згенеровано!")
                            st.rerun()
                    
                    # Відображення помилок під кнопкою
                    if st.session_state.get('last_error'):
                        st.error(st.session_state.last_error)
            
            elif st.session_state.processing_error:
                st.error(f"❌ {st.session_state.processing_error}")
        
        # Відображення результатів генерації
        if st.session_state.output_files:
            st.markdown("---")
            st.header("📥 Завантажити результати")
            
            # Поле для вибору шляху збереження
            st.subheader("📁 Налаштування збереження")
            new_save_tests_path = st.text_input(
                "Шлях для збереження файлів:",
                value=st.session_state.save_tests_path,
                help="Вкажіть папку для збереження згенерованих файлів. Натисніть Enter для збереження."
            )
            if new_save_tests_path != st.session_state.save_tests_path:
                st.session_state.save_tests_path = new_save_tests_path
                save_user_settings()
            
            # Кнопка "Завантажити все"
            st.markdown("### 🎯 Швидке збереження")
            if st.button("📦 Завантажити все", type="primary", use_container_width=True, help="Зберегти всі файли (тести, ключ, відповіді) в обрану папку"):
                save_all_tests()
            
            st.markdown("### 📋 Окремі файли")
            col1, col2, col3 = st.columns(3)
            
            # Тесты для учеников (Word)
            with col1:
                if os.path.exists(st.session_state.output_files['test_word']):
                    with open(st.session_state.output_files['test_word'], "rb") as file:
                        st.download_button(
                            label="📝 Тести Word",
                            data=file,
                            file_name=os.path.basename(st.session_state.output_files['test_word']),
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
            
            # Excel ключ
            with col2:
                if os.path.exists(st.session_state.output_files['excel_key']):
                    with open(st.session_state.output_files['excel_key'], "rb") as file:
                        st.download_button(
                            label="📊 Excel ключ",
                            data=file,
                            file_name=os.path.basename(st.session_state.output_files['excel_key']),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
            
            # Ответы (Word)
            with col3:
                if os.path.exists(st.session_state.output_files['answers_word']):
                    with open(st.session_state.output_files['answers_word'], "rb") as file:
                        st.download_button(
                            label="📋 Відповіді Word",
                            data=file,
                            file_name=os.path.basename(st.session_state.output_files['answers_word']),
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
    
    else:
        # Режим 2: Перевірка робіт
        st.header("✅ Режим 2: Перевірка робіт")
        
        # Завантаження файла-ключа
        answer_key_file = st.file_uploader(
            "Оберіть Excel файл-ключ",
            type=["xlsx", "xls"],
            help="Файл-ключ, створений при генерації тестів"
        )
        
        if answer_key_file is not None:
            # Зберігаємо файл-ключ
            temp_dir = ensure_temp_dir()
            key_file_path = os.path.join(temp_dir, answer_key_file.name)
            with open(key_file_path, "wb") as f:
                f.write(answer_key_file.getbuffer())
            st.session_state.answer_key_file = key_file_path
            # Зберігаємо назву файлу-ключа в налаштуваннях
            save_user_settings()
            st.success(f"✅ Файл-ключ завантажено: {answer_key_file.name}")
        
        # Ввод данных ученика
        if st.session_state.answer_key_file:
            st.subheader("👤 Дані учня та роботи")
            
            # Назва роботи
            st.session_state.test_work_name = st.text_input(
                "📝 Назва роботи:",
                value=st.session_state.test_work_name,
                placeholder="Наприклад: Контрольна робота №1 з математики"
            )
            
            col1, col2 = st.columns(2)
            with col1:
                st.session_state.student_class = st.text_input(
                    "Клас:",
                    value=st.session_state.student_class,
                    placeholder="Наприклад: 10-А"
                )
            with col2:
                st.session_state.student_full_name = st.text_input(
                    "ПІБ учня:",
                    value=st.session_state.student_full_name,
                    placeholder="Прізвище Ім'я По батькові"
                )
            
            # Налаштування файлу результатів
            st.subheader("📊 Налаштування файлу результатів")
            
            col1, col2 = st.columns([3, 1])
            with col1:
                st.session_state.results_excel_path = st.text_input(
                    "📁 Шлях до папки для збереження результатів:",
                    value=st.session_state.results_excel_path,
                    placeholder="Наприклад: C:\\Users\\Teacher\\Documents\\Results"
                )
            with col2:
                st.session_state.results_excel_filename = st.text_input(
                    "📄 Назва файлу:",
                    value=st.session_state.results_excel_filename,
                    placeholder="results.xlsx"
                )
            
            # Окремий шлях для збереження всіх результатів
            st.session_state.save_results_path = st.text_input(
                "💾 Шлях для збереження всіх результатів (PDF, Word, Excel):",
                value=st.session_state.save_results_path,
                placeholder="Наприклад: C:\\Users\\Teacher\\Downloads",
                help="Папка, куди будуть збережені всі файли результатів при натисканні кнопки 'Зберегти всі результати'"
            )
            
            st.markdown("---")
            
            # Ввод ответов ученика
            st.session_state.student_answers = st.text_input(
                "Відповіді учня (через кому):",
                value=st.session_state.student_answers,
                placeholder="Наприклад: А,В,Б,Г,А,Б або АБВ,Г,відкрита відповідь",
                help="Для тестових питань: введіть українські літери (А, Б, В, Г...). Для множинного вибору: АБВ. Для відкритих питань: введіть текстові відповіді. Розділяйте відповіді комами."
            )
            
            # Кнопка проверки
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                if st.button("🔍 Перевірити роботу", type="primary", use_container_width=True):
                    if st.session_state.student_answers.strip():
                        with st.spinner("Перевірка відповідей..."):
                            success = check_answers()
                        
                        if success:
                            st.success("✅ Перевірка завершена!")
                            st.rerun()
                    else:
                        st.error("❌ Введіть відповіді учня")
                
                # Відображення помилок під кнопкою
                if st.session_state.get('last_error'):
                    st.error(st.session_state.last_error)
        
        # Відображення результатів перевірки
        if hasattr(st.session_state, 'check_result') and st.session_state.check_result:
            st.markdown("---")
            st.header("📊 Результати перевірки")
            
            result = st.session_state.check_result
            
            # Метрики
            col1, col2, col3, col4, col5 = st.columns(5)
            with col1:
                st.metric("Варіант", result['variant_number'])
            with col2:
                st.metric("Всього питань", result['total_questions'])
            with col3:
                st.metric("Правильних відповідей", result['correct_answers'])
            with col4:
                st.metric("Відсоток", f"{result['score_percentage']:.1f}%")
            with col5:
                # Розрахунок балів по 12-бальній системі
                weighted_score = result.get('weighted_score', 0)
                max_score = result.get('max_score', result['total_questions'])
                if max_score > 0:
                    twelve_point_score = round((weighted_score / max_score) * 12, 1)
                else:
                    twelve_point_score = 0
                st.metric("Бали (12-бальна)", f"{twelve_point_score}")
            
            # Кнопки для скачивания отчетов
            if hasattr(st.session_state, 'check_reports') and st.session_state.check_reports:
                st.markdown("---")
                st.subheader("📥 Завантажити звіти")
                
                col1, col2 = st.columns(2)
                
                # PDF отчет
                with col1:
                    if os.path.exists(st.session_state.check_reports['pdf_report']):
                        # Створюємо кастомну назву для PDF
                        work_name = st.session_state.test_work_name if st.session_state.test_work_name else "Тест"
                        key_file_name = os.path.basename(st.session_state.answer_key_file) if st.session_state.answer_key_file else ""
                        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
                        
                        pdf_custom_name = create_custom_filename(
                            work_name, st.session_state.student_class, st.session_state.student_full_name,
                            st.session_state.check_result['variant_number'], key_file_name, timestamp, "pdf"
                        )
                        
                        with open(st.session_state.check_reports['pdf_report'], "rb") as file:
                            st.download_button(
                                label="📄 Короткий звіт (PDF)",
                                data=file,
                                file_name=pdf_custom_name,
                                mime="application/pdf",
                                use_container_width=True
                            )
                
                # Word отчет
                with col2:
                    if os.path.exists(st.session_state.check_reports['word_report']):
                        # Створюємо кастомну назву для Word
                        work_name = st.session_state.test_work_name if st.session_state.test_work_name else "Тест"
                        key_file_name = os.path.basename(st.session_state.answer_key_file) if st.session_state.answer_key_file else ""
                        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
                        
                        word_custom_name = create_custom_filename(
                            work_name, st.session_state.student_class, st.session_state.student_full_name,
                            st.session_state.check_result['variant_number'], key_file_name, timestamp, "docx"
                        )
                        
                        with open(st.session_state.check_reports['word_report'], "rb") as file:
                            st.download_button(
                                label="📝 Розширений звіт (Word)",
                                data=file,
                                file_name=word_custom_name,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True
                            )
                
                # Кнопка фіксації результату
                st.markdown("---")
                col1, col2, col3 = st.columns([1, 2, 1])
                with col2:
                    if st.button("💾 Фіксація результату учня у таблицю", type="secondary", use_container_width=True):
                        if st.session_state.results_excel_path and st.session_state.results_excel_filename:
                            with st.spinner("Збереження результату..."):
                                success = save_student_result_to_excel()
                            
                            if success:
                                st.success("✅ Результат успішно збережено у таблицю!")
                            else:
                                st.error("❌ Помилка при збереженні результату")
                        else:
                            st.error("❌ Вкажіть шлях до папки та назву файлу для збереження результатів")
                
                # Кнопка збереження всіх результатів
                st.markdown("---")
                col1, col2, col3 = st.columns([1, 2, 1])
                with col2:
                    if st.button("📦 Зберегти всі результати", type="primary", use_container_width=True, help="Зберегти PDF, Word звіти та зафіксувати результат у таблиці"):
                        if st.session_state.save_results_path and st.session_state.results_excel_path and st.session_state.results_excel_filename:
                            with st.spinner("Збереження всіх результатів..."):
                                success = save_all_results()
                            
                            if success:
                                st.balloons()
                        else:
                            st.error("❌ Вкажіть всі необхідні шляхи для збереження результатів")
    
    # Журнал событий
    with st.expander("📋 Журнал подій", expanded=False):
        if st.session_state.log_messages:
            for log_msg in st.session_state.log_messages[-20:]:  # Показываем последние 20 сообщений
                if "ERROR" in log_msg:
                    st.error(log_msg)
                elif "SUCCESS" in log_msg:
                    st.success(log_msg)
                elif "WARNING" in log_msg:
                    st.warning(log_msg)
                else:
                    st.info(log_msg)
        else:
            st.info("Журнал порожній")

if __name__ == "__main__":
    main()
