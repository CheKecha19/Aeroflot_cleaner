# processors/onec_processor.py
import pandas as pd
import os
import logging
from utils import get_onec_file, is_file_recent, normalize_name, find_duplicates, find_internal_duplicates, find_users_to_remove

logger = logging.getLogger(__name__)

def load_onec_data_new_format():
    """
    Загрузка данных из 1С в новом формате с обработкой объединенных ячеек
    """
    try:
        onec_file = get_onec_file()
        if not onec_file or not is_file_recent(onec_file):
            logger.warning("Актуальный файл 1С не найден")
            return pd.DataFrame(columns=['1C_ФИО', '1C_Активен'])

        # Читаем Excel файл без пропуска строк, чтобы найти начало данных
        df_raw = pd.read_excel(onec_file, sheet_name='Лист_1', header=None)
        
        logger.debug(f"Размер сырых данных: {df_raw.shape}")
        
        # Находим строку с заголовком "Пользователь"
        start_row = None
        for i in range(len(df_raw)):
            # Проверяем первый столбец на наличие "Пользователь"
            if (pd.notna(df_raw.iloc[i, 0]) and 
                str(df_raw.iloc[i, 0]).strip() == 'Пользователь'):
                start_row = i
                logger.debug(f"Найдена строка с заголовком в позиции {i}")
                break
        
        if start_row is None:
            logger.error("Не найдена строка с заголовком 'Пользователь'")
            return pd.DataFrame(columns=['1C_ФИО', '1C_Активен'])
        
        # Читаем данные, начиная со строки после заголовка
        df_data = pd.read_excel(
            onec_file, 
            sheet_name='Лист_1',
            skiprows=start_row + 1,  # Пропускаем строку заголовка
            header=None
        )
        
        logger.debug(f"Размер данных после пропуска: {df_data.shape}")
        
        # Создаем список для данных
        data = []
        valid_rows = 0
        
        for index, row in df_data.iterrows():
            # Проверяем, что строка содержит данные
            if pd.isna(row.iloc[0]) or str(row.iloc[0]).strip() == '':
                continue
                
            # Извлекаем ФИО из первого столбца (индекс 0)
            user_name = str(row.iloc[0]).strip()
            
            # Извлекаем статус из пятого столбца (индекс 4) - столбец E "Недействителен"
            if len(row) > 4 and pd.notna(row.iloc[4]):
                status = str(row.iloc[4]).strip()
                is_active = status == 'Нет'  # "Нет" = активен, "Да" = неактивен
                active_status = 'Да' if is_active else 'Нет'
                logger.debug(f"Пользователь {user_name}: статус '{status}' -> активен: {is_active}")
            else:
                # Если статус не найден, считаем активным
                active_status = 'Да'
                logger.debug(f"Статус не найден для пользователя {user_name}, помечен как активный")
            
            # Пропускаем служебные записи
            if any(service in user_name.lower() for service in ['сервис', 'robot', 'робот']):
                logger.debug(f"Пропущена служебная запись: {user_name}")
                continue
                
            data.append({
                '1C_ФИО': user_name,
                '1C_Активен': active_status
            })
            valid_rows += 1
        
        logger.info(f"Загружено {valid_rows} записей из 1С")
        
        if valid_rows == 0:
            logger.warning("Не найдено валидных записей пользователей в файле 1С")
            # Выведем первые несколько строк для отладки
            logger.debug(f"Первые 5 строк данных: {df_data.head().values.tolist()}")
        
        return pd.DataFrame(data)
        
    except Exception as e:
        logger.error(f"Ошибка при загрузке данных 1С: {e}", exc_info=True)
        return pd.DataFrame(columns=['1C_ФИО', '1C_Активен'])

def process_onec_data(df, ad_employees_df, selected_options, employee_types):
    """Обработка данных из 1С"""
    if 1 not in selected_options and 0 not in selected_options:
        return df, {}
    
    logger.info("Обработка данных 1С...")
    
    # Проверяем наличие необходимых данных в AD
    if ad_employees_df.empty or 'AD_ФИО' not in ad_employees_df.columns:
        logger.warning("AD DataFrame пуст или не содержит столбец 'AD_ФИО'")
        return df, {
            'duplicates_ad_1c': 0,
            'internal_duplicates_1c': 0,
            'users_to_remove_1c': pd.DataFrame()
        }
    
    # Загружаем данные из 1С в новом формате
    onec_data = load_onec_data_new_format()
    
    if not onec_data.empty:
        logger.info(f"Загружено {len(onec_data)} записей из 1С")
        
        # Убедимся, что не превышаем MAX_ROWS
        onec_fio = onec_data['1C_ФИО'][:len(df)]
        onec_active = onec_data['1C_Активен'][:len(df)]  # ИСПРАВЛЕНО: '1C_Активen' -> '1C_Активен'
        
        df['1C_ФИО'] = pd.Series(onec_fio)
        df['1C_Активен'] = pd.Series(onec_active)
        
        # Логируем несколько примеров для проверки
        sample_users = onec_data.head(3)[['1C_ФИО', '1C_Активен']].values.tolist()
        logger.debug(f"Примеры пользователей из 1С: {sample_users}")
        
        # Статистика по активным/неактивным
        active_count = len(onec_data[onec_data['1C_Активен'] == 'Да'])
        inactive_count = len(onec_data[onec_data['1C_Активен'] == 'Нет'])
        logger.info(f"Статистика 1С: {active_count} активных, {inactive_count} неактивных")
    else:
        logger.warning("Данные из 1С не загружены или пусты")
    
    # Разделение на отдельные DataFrame
    onec_df = df[['1C_ФИО', '1C_Активен']].dropna(subset=['1C_ФИО'])
    
    # Инициализация результатов
    results = {
        'duplicates_ad_1c': 0,
        'internal_duplicates_1c': 0,
        'users_to_remove_1c': pd.DataFrame()
    }
    
    if not onec_df.empty:
        # Поиск дубликатов для 1С
        results['duplicates_ad_1c'] = len(find_duplicates(ad_employees_df, onec_df, 'AD_ФИО', '1C_ФИО'))
        results['internal_duplicates_1c'] = len(find_internal_duplicates(onec_df, '1C_ФИО'))
        results['users_to_remove_1c'] = find_users_to_remove(onec_df, ad_employees_df, ad_employees_df)
        
        logger.info(f"Найдено дубликатов между AD и 1С: {results['duplicates_ad_1c']}")
        logger.info(f"Найдено внутренних дубликатов в 1С: {results['internal_duplicates_1c']}")
        logger.info(f"Найдено пользователей для удаления из 1С: {len(results['users_to_remove_1c'])}")
    else:
        logger.warning("DataFrame 1С пуст, пропускаем поиск дубликатов")
    
    return df, results

# Сохраняем старые функции для тестирования
def parse_1c_users_report(file_path):
    """
    Парсит отчёт 1С по пользователям в новом формате (для тестирования)
    """
    try:
        # Читаем Excel файл
        df = pd.read_excel(file_path, sheet_name='Лист_1', header=None)
        
        print(f"Размер сырых данных: {df.shape}")
        
        # Находим строку с заголовком "Пользователь"
        start_row = None
        for i in range(len(df)):
            if (pd.notna(df.iloc[i, 0]) and 
                str(df.iloc[i, 0]).strip() == 'Пользователь'):
                start_row = i
                print(f"Найдена строка с заголовком в позиции {i}")
                break
        
        if start_row is None:
            raise ValueError("Не найдена строка с заголовком 'Пользователь'")
        
        # Читаем данные, начиная со строки после заголовка
        df_data = pd.read_excel(
            file_path, 
            sheet_name='Лист_1',
            skiprows=start_row + 1,
            header=None
        )
        
        print(f"Размер данных после пропуска: {df_data.shape}")
        print(f"Первые 5 строк данных:")
        for i in range(min(5, len(df_data))):
            print(f"Строка {i}: {df_data.iloc[i].values}")
        
        # Форматируем результат
        users_data = []
        for index, row in df_data.iterrows():
            if pd.isna(row.iloc[0]) or str(row.iloc[0]).strip() == '':
                continue
                
            user_name = str(row.iloc[0]).strip()
            
            # Из столбца E (индекс 4) - "Недействителен"
            if len(row) > 4 and pd.notna(row.iloc[4]):
                status = str(row.iloc[4]).strip()
                is_active = status == 'Нет'  # "Нет" = активен, "Да" = неактивен
                print(f"Пользователь {user_name}: статус '{status}' -> активен: {is_active}")
            else:
                is_active = True
                print(f"Пользователь {user_name}: статус не найден -> активен: {is_active}")
            
            user_data = {
                'user_name': user_name,
                'is_active': is_active,
            }
            users_data.append(user_data)
        
        return users_data
        
    except Exception as e:
        print(f"Ошибка при обработке файла: {e}")
        import traceback
        traceback.print_exc()
        return []

def process_users_data(users_data):
    """
    Обрабатывает данные пользователей (сохранена предыдущая функциональность)
    """
    active_users = [user for user in users_data if user['is_active']]
    inactive_users = [user for user in users_data if not user['is_active']]
    
    print(f"Всего пользователей: {len(users_data)}")
    print(f"Активных: {len(active_users)}")
    print(f"Неактивных: {len(inactive_users)}")
    
    # Пример дополнительной обработки
    for user in users_data[:5]:  # Покажем первые 5 пользователей
        print(f"\nПользователь: {user['user_name']}")
        print(f"  Активен: {'Да' if user['is_active'] else 'Нет'}")

# Использование
if __name__ == "__main__":
    file_path = "Users.1C.xlsx"  # Изменил имя файла на актуальное
    
    if os.path.exists(file_path):
        users_data = parse_1c_users_report(file_path)
        process_users_data(users_data)
    else:
        print("Файл не найден")