# main.py
import logging
import pandas as pd
from config import INPUT_DIR, OUTPUT_DIR, OUTPUT_FILE
from excel_processor import process_excel_data
from ad_export import export_ad_users

# Получаем логгер для этого модуля
logger = logging.getLogger(__name__)

def get_user_choice():
    """Получение выбора пользователя"""
    print("\n" + "="*50)
    print("Выберите опции для проверки (через пробел):")
    print("0 - Всё")
    print("1 - 1С")
    print("2 - Сфера Курьер")  # ← ИЗМЕНИЛ: Диадок → Сфера Курьер
    print("3 - Контур Диадок")  # ← ИЗМЕНИЛ: Контур → Контур Диадок
    print("="*50)
    
    while True:
        choice = input("Ваш выбор: ").strip()
        
        if not choice:
            print("Пожалуйста, введите хотя бы одну цифру")
            continue
            
        choices = choice.split()
        
        # Проверка на валидность ввода
        valid_choices = {'0', '1', '2', '3'}
        if all(c in valid_choices for c in choices):
            # Если выбран 0, добавляем все остальные опции
            if '0' in choices:
                return {0, 1, 2, 3}
            return set(int(c) for c in choices)
        else:
            print("Некорректный ввод. Пожалуйста, используйте цифры 0, 1, 2, 3 через пробел")

def get_employee_type_choice():
    """Получение выбора типа сотрудников"""
    print("\n" + "="*50)
    print("Выберите тип сотрудников для проверки (через пробел):")
    print("0 - Все")
    print("1 - Сотрудники")
    print("2 - ГПХ")
    print("="*50)
    
    while True:
        choice = input("Ваш выбор: ").strip()
        
        if not choice:
            print("Пожалуйста, введите хотя бы одну цифру")
            continue
            
        choices = choice.split()
        
        # Проверка на валидность ввода
        valid_choices = {'0', '1', '2'}
        if all(c in valid_choices for c in choices):
            # Если выбран 0, добавляем все остальные опции
            if '0' in choices:
                return {0, 1, 2}
            return set(int(c) for c in choices)
        else:
            print("Некорректный ввод. Пожалуйста, используйте цифры 0, 1, 2 через пробел")

def main():
    logger.info("Запуск обработки данных")
    
    # Получаем выбор пользователя
    selected_options = get_user_choice()
    selected_employee_types = get_employee_type_choice()
    logger.info(f"Выбранные опции: {selected_options}")
    logger.info(f"Выбранные типы сотрудников: {selected_employee_types}")
    
    # Экспорт данных из AD (всегда выполняется)
    try:
        logger.info("Экспорт пользователей из Active Directory")
        total_users, employees_count, gph_count = export_ad_users()
        logger.info(f"Экспорт AD завершен: {total_users} пользователей, {employees_count} сотрудников, {gph_count} ГПХ")
    except Exception as e:
        logger.error(f"Ошибка при экспорте из AD: {e}")
        logger.info("Продолжение обработки с пустыми данными AD")
        total_users, employees_count, gph_count = 0, 0, 0
    
    # Обработка Excel данных
    try:
        logger.info("Обработка Excel данных")
        results = process_excel_data(selected_options, selected_employee_types)
        
        logger.info("Обработка завершена. Результаты:")
        if 1 in selected_options or 0 in selected_options:
            logger.info(f"- Дубликаты между AD и 1С: {results.get('duplicates_ad_1c', 0)}")
            logger.info(f"- Внутренние дубликаты в 1С: {results.get('internal_duplicates_1c', 0)}")
            logger.info(f"- Пользователей для удаления из 1С: {len(results.get('users_to_remove_1c', pd.DataFrame()))}")
        if 2 in selected_options or 0 in selected_options:
            logger.info(f"- Дубликаты между AD и Сфера Курьер: {results.get('duplicates_ad_diadoc', 0)}")  # ← ИЗМЕНИЛ
            logger.info(f"- Внутренние дубликаты в Сфере Курьер: {results.get('internal_duplicates_diadoc', 0)}")  # ← ИЗМЕНИЛ
            logger.info(f"- Пользователей для удаления из Сферы Курьер: {len(results.get('users_to_remove_diadoc', pd.DataFrame()))}")  # ← ИЗМЕНИЛ
        if 3 in selected_options or 0 in selected_options:
            logger.info(f"- Дубликаты между AD и Контур Диадок: {results.get('duplicates_ad_kontur', 0)}")  # ← ИЗМЕНИЛ
            logger.info(f"- Внутренние дубликаты в Контур Диадок: {results.get('internal_duplicates_kontur', 0)}")  # ← ИЗМЕНИЛ
            logger.info(f"- Пользователей для удаления из Контур Диадок: {len(results.get('users_to_remove_kontur', pd.DataFrame()))}")  # ← ИЗМЕНИЛ
        logger.info(f"- Несоответствий между AD и Штатным расписанием: {results.get('comparison_count', 0)}")
        
    except Exception as e:
        logger.error(f"Ошибка при обработке Excel: {str(e)}")
    
    logger.info(f"Результаты сохранены в файл: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()