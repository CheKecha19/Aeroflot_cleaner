# config.py
import os
import logging
from pathlib import Path
from datetime import datetime, timedelta

# Базовые пути
BASE_DIR = Path(__file__).parent
INPUT_DIR = BASE_DIR / "эксельки"
OUTPUT_DIR = BASE_DIR / "вывод"
LOG_DIR = OUTPUT_DIR / "log"
AD_EXPORT_DIR = INPUT_DIR / "AD"
SHTAT_DIR = INPUT_DIR / "штатка"
KONTUR_DIR = INPUT_DIR / "эдо_контур_диадок"
DIADOC_DIR = INPUT_DIR / "эдо_сфера_курьер"
ONEC_DIR = INPUT_DIR / "1С"

# Создаем директории, если они не существуют
INPUT_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)
LOG_DIR.mkdir(exist_ok=True)
AD_EXPORT_DIR.mkdir(exist_ok=True)
SHTAT_DIR.mkdir(exist_ok=True)
KONTUR_DIR.mkdir(exist_ok=True)
DIADOC_DIR.mkdir(exist_ok=True)
ONEC_DIR.mkdir(exist_ok=True)

# Настройка актуальности файлов (в днях)
MAX_FILE_AGE_DAYS = 180

# Генерация имени файла с датой и временем
current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
OUTPUT_FILE = OUTPUT_DIR / f"результат_обработки_{current_time}.xlsx"

# Файлы сотрудников и ГПХ (создаются автоматически)
EMPLOYEES_FILE = AD_EXPORT_DIR / "сотрудники.txt"
GPH_FILE = AD_EXPORT_DIR / "ГПХ.txt"

# Настройки обработки Excel
SHEET_NAME = "сравнение пользователей"
COMPARISON_SHEET = "сравнение AD и Штатки"
KONTUR_SHEET = "Контур Диадок данные"
DIADOC_SHEET = "Сфера Курьер данные"
ONEC_SHEET = "1С данные"
MAX_ROWS = 10000
RED_COLOR = (255, 199, 206)
YELLOW_COLOR = (255, 235, 156)

# ========== НАСТРОЙКИ ЛОГИРОВАНИЯ ==========
LOG_LEVEL = logging.INFO
LOG_FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'

def setup_logging():
    """Настройка логирования для всего приложения"""
    # Очищаем существующие обработчики
    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)
    
    # Основной логгер приложения
    logging.basicConfig(
        level=LOG_LEVEL,
        format=LOG_FORMAT,
        handlers=[
            logging.FileHandler(LOG_DIR / "log.txt", encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    
    # Дополнительный логгер для AD экспорта
    ad_logger = logging.getLogger('ad_export')
    ad_logger.setLevel(LOG_LEVEL)
    ad_logger.handlers.clear()
    
    ad_file_handler = logging.FileHandler(LOG_DIR / "ad_export.log", encoding='utf-8')
    ad_file_handler.setFormatter(logging.Formatter(LOG_FORMAT))
    ad_logger.addHandler(ad_file_handler)
    
    ad_console_handler = logging.StreamHandler()
    ad_console_handler.setFormatter(logging.Formatter(LOG_FORMAT))
    ad_logger.addHandler(ad_console_handler)
    
    ad_logger.propagate = False

# Инициализируем логирование при импорте конфига
setup_logging()