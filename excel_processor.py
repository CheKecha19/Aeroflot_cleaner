# excel_processor.py
import pandas as pd
import numpy as np
from config import OUTPUT_FILE, SHEET_NAME, COMPARISON_SHEET, MAX_ROWS, EMPLOYEES_FILE, GPH_FILE
from config import SHTAT_DIR
from utils import replace_yo, normalize_name, find_internal_duplicates
from utils import load_shtat_data, create_comparison_sheet
from processors.onec_processor import process_onec_data
from processors.kontur_processor import process_kontur_data
from processors.diadoc_processor import process_diadoc_data
from utils import replace_yo, normalize_name, find_internal_duplicates
from utils import load_shtat_data, create_comparison_sheet
import logging
logger = logging.getLogger(__name__)

def read_names_and_statuses_from_file(filename):
    """Чтение имен и статусов из файла в формате 'Name: ФИО' и 'Status: Статус'"""
    names = []
    statuses = []
    current_name = None
    current_status = "Неизвестно"
    
    try:
        with open(filename, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line.startswith('Name:'):
                    # Если уже есть текущее имя, сохраняем его с текущим статусом
                    if current_name is not None:
                        names.append(current_name)
                        statuses.append(current_status)
                    current_name = line.split(':', 1)[1].strip()
                    current_status = "Неизвестно"  # Сбрасываем статус для нового имени
                elif line.startswith('Status:'):
                    current_status = line.split(':', 1)[1].strip()
            
            # Добавляем последнее имя, если оно есть
            if current_name is not None:
                names.append(current_name)
                statuses.append(current_status)
                
        return names, statuses
    except Exception as e:
        logger.error(f"Ошибка при чтении файла {filename}: {e}")
        return [], []

def process_excel_data(selected_options=None, employee_types=None):
    """Основная функция обработки Excel данных"""
    if selected_options is None:
        selected_options = {0}  # По умолчанию проверяем всё
    
    if employee_types is None:
        employee_types = {0}  # По умолчанию все типы сотрудников
    
    # Создаем новый DataFrame с нужной структурой
    df = pd.DataFrame(index=range(MAX_ROWS), columns=[
        'Штатное_ФИО',
        'AD_сотрудники',
        'AD_Статус_сотрудники',
        'AD_ГПХ',
        'AD_Статус_ГПХ',
        'Контур_Диадок_ФИО',  # ← ИЗМЕНИЛ: Контур → Контур_Диадок
        'Контур_Диадок_Администратор',  # ← ИЗМЕНИЛ
        'Контур_Диадок_статус',  # ← ИЗМЕНИЛ
        'Сфера_Курьер_ФИО',  # ← ИЗМЕНИЛ: Диадок → Сфера_Курьер
        'Сфера_Курьер_Активен',  # ← ИЗМЕНИЛ
        'Сфера_Курьер_Администратор',  # ← ИЗМЕНИЛ
        '1C_ФИО',
        '1C_Активен'
    ])
    
    # Чтение сотрудников из AD с фильтрацией по типам
    employees_names, employees_statuses = read_names_and_statuses_from_file(EMPLOYEES_FILE)
    gph_names, gph_statuses = read_names_and_statuses_from_file(GPH_FILE)
    
    # Заполняем столбцы AD
    df['AD_сотрудники'] = pd.Series(employees_names)
    df['AD_Статус_сотрудники'] = pd.Series(employees_statuses)
    df['AD_ГПХ'] = pd.Series(gph_names)
    df['AD_Статус_ГПХ'] = pd.Series(gph_statuses)
    
    # Создаем объединенный DataFrame AD сотрудников для сравнения
    ad_employees_data = []
    
    # Добавляем сотрудников, если выбраны
    if 1 in employee_types or 0 in employee_types:
        for i, name in enumerate(employees_names):
            if i < len(employees_statuses):
                ad_employees_data.append({'AD_ФИО': name, 'AD_Статус': employees_statuses[i]})
    
    # Добавляем ГПХ, если выбраны
    if 2 in employee_types or 0 in employee_types:
        for i, name in enumerate(gph_names):
            if i < len(gph_statuses):
                ad_employees_data.append({'AD_ФИО': name, 'AD_Статус': gph_statuses[i]})
    
    # Создаем DataFrame для сравнения
    if ad_employees_data:
        ad_employees_df = pd.DataFrame(ad_employees_data)
    else:
        ad_employees_df = pd.DataFrame(columns=['AD_ФИО', 'AD_Статус'])
    
    # Загружаем данные из штатного расписания
    shtat_data = load_shtat_data()
    if not shtat_data.empty:
        df['Штатное_ФИО'] = pd.Series(shtat_data['Штатное_ФИО'])
    
    # Обработка данных из различных источников
    df, _ = process_onec_data(df, ad_employees_df, selected_options, employee_types)
    df, _ = process_kontur_data(df, ad_employees_df, selected_options, employee_types)
    df, _ = process_diadoc_data(df, ad_employees_df, selected_options, employee_types)
    
    # Замена ё на е во всех столбцах с ФИО
    for col in ['Штатное_ФИО', 'AD_сотрудники', 'AD_ГПХ', 'Контур_Диадок_ФИО', 'Сфера_Курьер_ФИО', '1C_ФИО']:  # ← ИЗМЕНИЛ
        if col in df.columns:
            df.loc[:, col] = df[col].apply(lambda x: replace_yo(x) if pd.notna(x) else x)
    
    # Удаляем полностью пустые строки
    df = df.replace('', np.nan).dropna(how='all')
    
    # Сохранение основного листа
    df.to_excel(OUTPUT_FILE, sheet_name=SHEET_NAME, index=False)
    
    # Создание листа сравнения AD и Штатного расписания
    shtat_names = shtat_data['Штатное_ФИО'].tolist() if not shtat_data.empty else []
    comparison_count = create_comparison_sheet(employees_names, shtat_names, OUTPUT_FILE)
    
    # Создаем набор всех AD имен для сравнения
    all_ad_names = set()
    for col in ['AD_сотрудники', 'AD_ГПХ']:
        if col in df.columns:
            names = df[col].dropna().apply(normalize_name)
            all_ad_names.update(names)
    
    # Сохранение результатов в отдельные листы
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl', mode='a') as writer:
        # Обрабатываем каждый сервис
        services = [
            {
                'name': 'Контур Диадок',  # ← ИЗМЕНИЛ
                'fio_col': 'Контур_Диадок_ФИО',  # ← ИЗМЕНИЛ
                'status_col': 'Контур_Диадок_статус',  # ← ИЗМЕНИЛ
                'active_value': 'активна',
                'remove_sheet': 'удалить из Контур Диадок',  # ← ИЗМЕНИЛ
                'duplicates_sheet': 'дубли в Контур Диадок'  # ← ИЗМЕНИЛ
            },
            {
                'name': 'Сфера Курьер',  # ← ИЗМЕНИЛ
                'fio_col': 'Сфера_Курьер_ФИО',  # ← ИЗМЕНИЛ
                'status_col': 'Сфера_Курьер_Активен',  # ← ИЗМЕНИЛ
                'active_value': 'Да',
                'remove_sheet': 'удалить из Сфера Курьер',  # ← ИЗМЕНИЛ
                'duplicates_sheet': 'дубли в Сфера Курьер'  # ← ИЗМЕНИЛ
            },
            {
                'name': '1С',
                'fio_col': '1C_ФИО',
                'status_col': '1C_Активен',
                'active_value': 'Да',
                'remove_sheet': 'удалить из 1С',
                'duplicates_sheet': 'дубли в 1С'
            }
        ]
        
        for service in services:
            fio_col = service['fio_col']
            status_col = service['status_col']
            active_value = service['active_value']
            remove_sheet = service['remove_sheet']
            duplicates_sheet = service['duplicates_sheet']
            
            # Пропускаем если столбцы не существуют
            if fio_col not in df.columns:
                logger.debug(f"Пропускаем {service['name']}: столбец {fio_col} не найден")
                continue
            
            # 1. Поиск и сохранение дубликатов
            service_fio_data = df[[fio_col]].dropna(subset=[fio_col])
            duplicates = find_internal_duplicates(service_fio_data, fio_col)
            if duplicates:
                duplicate_df = service_fio_data[service_fio_data[fio_col].apply(normalize_name).isin(duplicates)]
                if not duplicate_df.empty:
                    duplicate_df.to_excel(writer, sheet_name=duplicates_sheet, index=False)
                    logger.info(f"Создан лист {duplicates_sheet} с {len(duplicate_df)} записями")
            
            # 2. Поиск и сохранение пользователей для удаления
            if status_col not in df.columns:
                logger.debug(f"Пропускаем {remove_sheet}: столбец {status_col} не найден")
                continue
                
            # Берем только строки с заполненным ФИО
            service_data = df[[fio_col, status_col]].dropna(subset=[fio_col])
            # Приводим статус к строке и обрезаем пробелы
            service_data.loc[:, status_col] = service_data[status_col].astype(str).str.strip()
            
            # Фильтруем: активные пользователи, которых нет в AD
            mask = (service_data[status_col].str.lower() == active_value.lower()) & \
                (~service_data[fio_col].apply(normalize_name).isin(all_ad_names))
            users_to_remove = service_data[mask]
            
            if not users_to_remove.empty:
                users_to_remove.to_excel(writer, sheet_name=remove_sheet, index=False)
                logger.info(f"Создан лист {remove_sheet} с {len(users_to_remove)} записями")
            else:
                logger.debug(f"Нет данных для листа {remove_sheet}")
        
        # Дополнительная проверка для Контур Диадок
        if 'Контур_Диадок_ФИО' in df.columns and 'Контур_Диадок_статус' in df.columns:  # ← ИЗМЕНИЛ
            kontur_data = df[['Контур_Диадок_ФИО', 'Контур_Диадок_статус']].dropna(subset=['Контур_Диадок_ФИО'])  # ← ИЗМЕНИЛ
            kontur_data.loc[:, 'Контур_Диадок_статус'] = kontur_data['Контур_Диадок_статус'].astype(str).str.strip()  # ← ИЗМЕНИЛ
            
            # Проверяем, есть ли активные пользователи в Контур Диадок
            active_kontur_users = kontur_data[kontur_data['Контур_Диадок_статус'].str.lower() == 'активна']  # ← ИЗМЕНИЛ
            logger.debug(f"Активных пользователей в Контур Диадок: {len(active_kontur_users)}")  # ← ИЗМЕНИЛ
            
            # Проверяем, сколько из них нет в AD
            kontur_users_not_in_ad = active_kontur_users[
                ~active_kontur_users['Контур_Диадок_ФИО'].apply(normalize_name).isin(all_ad_names)  # ← ИЗМЕНИЛ
            ]
            logger.debug(f"Активных пользователей в Контур Диадок, которых нет в AD: {len(kontur_users_not_in_ad)}")  # ← ИЗМЕНИЛ
        
    return {'comparison_count': comparison_count}