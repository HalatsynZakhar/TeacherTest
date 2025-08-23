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
from fpdf import FPDF

# Добавляем корневую папку проекта в PYTHONPATH
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

# Используем относительные импорты
from utils import config_manager
from utils import excel_utils
from utils.config_manager import get_downloads_folder, ConfigManager
from core.processor import (
    read_test_excel, 
    generate_test_variants, 
    create_test_pdf, 
    create_excel_answer_key,
    check_student_answers,
    create_check_result_pdf
)

# Настройка логирования
log_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'logs')
os.makedirs(log_dir, exist_ok=True)

# Ограничиваем количество файлов логов до 5 последних
log_files = sorted([f for f in os.listdir(log_dir) if f.startswith('teacher_test_')])
if len(log_files) > 5:
    for old_log in log_files[:-5]:
        try:
            os.remove(os.path.join(log_dir, old_log))
        except:
            pass

# Создаем новый лог-файл
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

# Определяем настройки по умолчанию для TeacherTest
default_settings = {
    "test_settings": {
        "default_variants_count": 10,
        "shuffle_questions": True,
        "shuffle_answers": True
    },
    "file_settings": {
        "max_size_mb": 50
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
        
        config_manager_instance.save_settings()
        st.session_state.config_manager = config_manager_instance
        log.info("Менеджер конфигурации TeacherTest инициализирован")
    
    return st.session_state.config_manager

# Инициализация
config_folder = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'settings_presets')
config_manager.init_config_manager(config_folder)
init_config_manager()
cm = st.session_state.config_manager

# Настройка параметров приложения
st.set_page_config(
    page_title="TeacherTest - Генератор тестов",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Функция для создания временных директорий
def ensure_temp_dir(prefix: str = "") -> str:
    """Создает и возвращает путь к временной директории"""
    project_dir = os.path.dirname(os.path.dirname(__file__))
    temp_dir = os.path.join(project_dir, "temp")
    
    try:
        os.makedirs(temp_dir, exist_ok=True)
        log.info(f"Создана/проверена временная директория: {temp_dir}")
    except Exception as e:
        log.error(f"Ошибка при создании временной директории {temp_dir}: {e}")
        temp_dir = os.path.join(tempfile.gettempdir(), f"{prefix}teachertest")
        try:
            os.makedirs(temp_dir, exist_ok=True)
        except Exception as e2:
            log.error(f"Ошибка при создании системной временной директории: {e2}")
            raise e2
    
    return temp_dir

# Функция очистки временных файлов
def cleanup_temp_files():
    """Очистка временных файлов при запуске приложения"""
    try:
        project_dir = os.path.dirname(os.path.dirname(__file__))
        temp_dir = os.path.join(project_dir, "temp")
        
        if os.path.exists(temp_dir):
            for filename in os.listdir(temp_dir):
                file_path = os.path.join(temp_dir, filename)
                try:
                    if os.path.isfile(file_path):
                        # Проверяем возраст файла (удаляем файлы старше 1 дня)
                        file_age = time.time() - os.path.getmtime(file_path)
                        if file_age > 86400:  # 24 часа в секундах
                            os.remove(file_path)
                            log.info(f"Удален старый временный файл: {filename}")
                except Exception as e:
                    log.warning(f"Не удалось удалить временный файл {filename}: {e}")
    except Exception as e:
        log.error(f"Ошибка при очистке временных файлов: {e}")

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
    st.session_state.mode = 1  # 1 - генерация тестов, 2 - проверка работ
if 'variants_count' not in st.session_state:
    st.session_state.variants_count = cm.get_setting('test_settings.default_variants_count', 10)
if 'answer_key_file' not in st.session_state:
    st.session_state.answer_key_file = None
if 'variant_number' not in st.session_state:
    st.session_state.variant_number = 1
if 'student_answers' not in st.session_state:
    st.session_state.student_answers = ""

def add_log_message(message, level="INFO"):
    """Добавление сообщения в лог"""
    timestamp = datetime.now().strftime("%H:%M:%S")
    log_entry = f"[{timestamp}] {level}: {message}"
    
    if 'log_messages' not in st.session_state:
        st.session_state.log_messages = []
    
    st.session_state.log_messages.append(log_entry)
    
    # Ограничиваем количество сообщений в логе
    if len(st.session_state.log_messages) > 100:
        st.session_state.log_messages = st.session_state.log_messages[-100:]
    
    # Логируем в файл
    if level == "ERROR":
        log.error(message)
    elif level == "WARNING":
        log.warning(message)
    else:
        log.info(message)

def load_excel_file(uploaded_file_arg=None):
    """Загрузка и обработка Excel файла с вопросами"""
    try:
        if uploaded_file_arg is None:
            uploaded_file_arg = st.session_state.get('temp_file_path')
        
        if not uploaded_file_arg or not os.path.exists(uploaded_file_arg):
            st.session_state.df = None
            st.session_state.processing_error = "Файл не найден"
            return
        
        add_log_message(f"Загрузка файла: {os.path.basename(uploaded_file_arg)}")
        
        # Читаем Excel файл с вопросами
        df = read_test_excel(uploaded_file_arg)
        
        if df.empty:
            st.session_state.df = None
            st.session_state.processing_error = "Файл не содержит данных или имеет неправильную структуру"
            add_log_message("Файл пуст или имеет неправильную структуру", "ERROR")
            return
        
        st.session_state.df = df
        st.session_state.processing_error = None
        add_log_message(f"Файл успешно загружен. Найдено {len(df)} вопросов", "SUCCESS")
        
    except Exception as e:
        error_msg = f"Ошибка при загрузке файла: {str(e)}"
        st.session_state.processing_error = error_msg
        st.session_state.df = None
        add_log_message(error_msg, "ERROR")
        log.error(f"Ошибка при загрузке файла: {e}", exc_info=True)

def generate_tests():
    """Генерация тестов"""
    try:
        add_log_message("Начало генерации тестов")
        
        if st.session_state.df is None or st.session_state.df.empty:
            raise ValueError("Нет данных для генерации тестов")
        
        # Генерируем варианты тестов
        variants = generate_test_variants(st.session_state.df, st.session_state.variants_count)
        add_log_message(f"Сгенерировано {len(variants)} вариантов тестов")
        
        # Создаем временную папку для выходных файлов
        output_dir = ensure_temp_dir("output_")
        
        # Создаем PDF файлы
        test_pdf_path, answers_pdf_path = create_test_pdf(variants, output_dir)
        add_log_message(f"Созданы PDF файлы: тесты и ответы")
        
        # Создаем Excel файл-ключ
        excel_key_path = create_excel_answer_key(variants, output_dir)
        add_log_message(f"Создан Excel файл-ключ")
        
        # Сохраняем пути к файлам
        st.session_state.output_files = {
            'test_pdf': test_pdf_path,
            'answers_pdf': answers_pdf_path,
            'excel_key': excel_key_path
        }
        
        add_log_message("Генерация тестов завершена успешно", "SUCCESS")
        return True
        
    except Exception as e:
        error_msg = f"Ошибка при генерации тестов: {str(e)}"
        add_log_message(error_msg, "ERROR")
        log.error(error_msg, exc_info=True)
        return False

def check_answers():
    """Проверка ответов ученика"""
    try:
        add_log_message("Начало проверки ответов")
        
        if not st.session_state.answer_key_file:
            raise ValueError("Не выбран файл-ключ")
        
        # Парсим ответы ученика
        try:
            student_answers = [int(x.strip()) for x in st.session_state.student_answers.split(',') if x.strip()]
        except ValueError:
            raise ValueError("Ответы должны быть числами, разделенными запятыми")
        
        # Проверяем ответы
        check_result = check_student_answers(
            st.session_state.answer_key_file,
            st.session_state.variant_number,
            student_answers
        )
        
        # Создаем PDF с результатом
        output_dir = ensure_temp_dir("check_")
        result_pdf_path = create_check_result_pdf(check_result, output_dir)
        
        # Сохраняем результат
        st.session_state.check_result = check_result
        st.session_state.check_result_pdf = result_pdf_path
        
        add_log_message(f"Проверка завершена. Правильных ответов: {check_result['correct_answers']} из {check_result['total_questions']}", "SUCCESS")
        return True
        
    except Exception as e:
        error_msg = f"Ошибка при проверке ответов: {str(e)}"
        add_log_message(error_msg, "ERROR")
        log.error(error_msg, exc_info=True)
        return False

# Основной интерфейс
def main():
    st.title("📝 TeacherTest - Генератор тестов для учеников")
    st.markdown("---")
    
    # Боковая панель с настройками
    with st.sidebar:
        st.header("⚙️ Настройки")
        
        # Выбор режима работы
        st.session_state.mode = st.radio(
            "Режим работы:",
            [1, 2],
            format_func=lambda x: "Генерация тестов" if x == 1 else "Проверка работ",
            index=st.session_state.mode - 1
        )
        
        st.markdown("---")
        
        if st.session_state.mode == 1:
            st.subheader("Настройки генерации")
            st.session_state.variants_count = st.number_input(
                "Количество вариантов:",
                min_value=1,
                max_value=100,
                value=st.session_state.variants_count
            )
        else:
            st.subheader("Настройки проверки")
            st.session_state.variant_number = st.number_input(
                "Номер варианта ученика:",
                min_value=1,
                max_value=100,
                value=st.session_state.variant_number
            )
    
    # Основной контент
    if st.session_state.mode == 1:
        # Режим 1: Генерация тестов
        st.header("🎯 Режим 1: Генерация тестов")
        
        # Загрузка Excel файла с вопросами
        uploaded_file = st.file_uploader(
            "Выберите Excel файл с вопросами",
            type=["xlsx", "xls"],
            help="Файл должен содержать: столбец с вопросами, столбец с номером правильного ответа, столбцы с вариантами ответов"
        )
        
        if uploaded_file is not None:
            # Сохраняем файл во временную папку
            temp_dir = ensure_temp_dir()
            temp_file_path = os.path.join(temp_dir, uploaded_file.name)
            
            # Проверяем, нужно ли обновить файл
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
                load_excel_file()
            
            # Отображение информации о файле
            if st.session_state.df is not None:
                st.success(f"✅ Файл загружен: {uploaded_file.name}")
                st.info(f"📊 Найдено вопросов: {len(st.session_state.df)}")
                
                # Предпросмотр данных
                with st.expander("👀 Предпросмотр данных"):
                    st.dataframe(st.session_state.df.head(10), use_container_width=True)
                
                # Кнопка генерации
                col1, col2, col3 = st.columns([1, 2, 1])
                with col2:
                    if st.button("🚀 Сгенерировать тесты", type="primary", use_container_width=True):
                        with st.spinner("Генерация тестов..."):
                            success = generate_tests()
                        
                        if success:
                            st.success("✅ Тесты успешно сгенерированы!")
                            st.rerun()
            
            elif st.session_state.processing_error:
                st.error(f"❌ {st.session_state.processing_error}")
        
        # Отображение результатов генерации
        if st.session_state.output_files:
            st.markdown("---")
            st.header("📥 Скачать результаты")
            
            col1, col2, col3 = st.columns(3)
            
            # Тесты для учеников
            with col1:
                if os.path.exists(st.session_state.output_files['test_pdf']):
                    with open(st.session_state.output_files['test_pdf'], "rb") as file:
                        st.download_button(
                            label="📄 Тесты для учеников",
                            data=file,
                            file_name=os.path.basename(st.session_state.output_files['test_pdf']),
                            mime="application/pdf",
                            use_container_width=True
                        )
            
            # Ответы для учителя
            with col2:
                if os.path.exists(st.session_state.output_files['answers_pdf']):
                    with open(st.session_state.output_files['answers_pdf'], "rb") as file:
                        st.download_button(
                            label="📋 Ответы для учителя",
                            data=file,
                            file_name=os.path.basename(st.session_state.output_files['answers_pdf']),
                            mime="application/pdf",
                            use_container_width=True
                        )
            
            # Excel ключ
            with col3:
                if os.path.exists(st.session_state.output_files['excel_key']):
                    with open(st.session_state.output_files['excel_key'], "rb") as file:
                        st.download_button(
                            label="📊 Excel ключ",
                            data=file,
                            file_name=os.path.basename(st.session_state.output_files['excel_key']),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
    
    else:
        # Режим 2: Проверка работ
        st.header("✅ Режим 2: Проверка работ")
        
        # Загрузка файла-ключа
        answer_key_file = st.file_uploader(
            "Выберите Excel файл-ключ",
            type=["xlsx", "xls"],
            help="Файл-ключ, созданный при генерации тестов"
        )
        
        if answer_key_file is not None:
            # Сохраняем файл-ключ
            temp_dir = ensure_temp_dir()
            key_file_path = os.path.join(temp_dir, answer_key_file.name)
            with open(key_file_path, "wb") as f:
                f.write(answer_key_file.getbuffer())
            st.session_state.answer_key_file = key_file_path
            st.success(f"✅ Файл-ключ загружен: {answer_key_file.name}")
        
        # Ввод ответов ученика
        if st.session_state.answer_key_file:
            st.session_state.student_answers = st.text_input(
                "Ответы ученика (через запятую):",
                value=st.session_state.student_answers,
                placeholder="Например: 1,3,2,4,1,2",
                help="Введите номера ответов ученика через запятую"
            )
            
            # Кнопка проверки
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                if st.button("🔍 Проверить работу", type="primary", use_container_width=True):
                    if st.session_state.student_answers.strip():
                        with st.spinner("Проверка ответов..."):
                            success = check_answers()
                        
                        if success:
                            st.success("✅ Проверка завершена!")
                            st.rerun()
                    else:
                        st.error("❌ Введите ответы ученика")
        
        # Отображение результатов проверки
        if hasattr(st.session_state, 'check_result') and st.session_state.check_result:
            st.markdown("---")
            st.header("📊 Результаты проверки")
            
            result = st.session_state.check_result
            
            # Метрики
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Вариант", result['variant_number'])
            with col2:
                st.metric("Всего вопросов", result['total_questions'])
            with col3:
                st.metric("Правильных ответов", result['correct_answers'])
            with col4:
                st.metric("Процент", f"{result['score_percentage']:.1f}%")
            
            # Скачать результат
            if hasattr(st.session_state, 'check_result_pdf') and os.path.exists(st.session_state.check_result_pdf):
                col1, col2, col3 = st.columns([1, 2, 1])
                with col2:
                    with open(st.session_state.check_result_pdf, "rb") as file:
                        st.download_button(
                            label="📄 Скачать результат проверки",
                            data=file,
                            file_name=os.path.basename(st.session_state.check_result_pdf),
                            mime="application/pdf",
                            use_container_width=True
                        )
    
    # Журнал событий
    with st.expander("📋 Журнал событий", expanded=False):
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
            st.info("Журнал пуст")

if __name__ == "__main__":
    main()
