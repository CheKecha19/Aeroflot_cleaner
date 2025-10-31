# processors/kontur_processor.py
import pandas as pd
from utils import load_kontur_data, find_duplicates, find_internal_duplicates, find_users_to_remove
import logging
logger = logging.getLogger(__name__)

def process_kontur_data(df, ad_employees_df, selected_options, employee_types):
    """Обработка данных из Контур Диадок"""  # ← ИЗМЕНИЛ комментарий
    if 3 not in selected_options and 0 not in selected_options:
        return df, {}
    
    logger.info("Обработка данных Контур Диадок...")  # ← ИЗМЕНИЛ
    
    # Проверяем наличие необходимых данных в AD
    if ad_employees_df.empty or 'AD_ФИО' not in ad_employees_df.columns:
        logger.warning("AD DataFrame пуст или не содержит столбец 'AD_ФИО'")
        return df, {
            'duplicates_ad_kontur': 0,
            'internal_duplicates_kontur': 0,
            'users_to_remove_kontur': pd.DataFrame()
        }
    
    # Загружаем данные из Контур Диадок
    kontur_data = load_kontur_data()
    
    if not kontur_data.empty:
        # Убедимся, что не превышаем MAX_ROWS
        kontur_fio = kontur_data['Контур_Диадок_ФИО'][:len(df)]  # ← ИЗМЕНИЛ
        kontur_admin = kontur_data['Контур_Диадок_Администратор'][:len(df)]  # ← ИЗМЕНИЛ
        kontur_status = kontur_data['Контур_Диадок_статус'][:len(df)]  # ← ИЗМЕНИЛ
        
        df['Контур_Диадок_ФИО'] = pd.Series(kontur_fio)  # ← ИЗМЕНИЛ
        df['Контур_Диадок_Администратор'] = pd.Series(kontur_admin)  # ← ИЗМЕНИЛ
        df['Контур_Диадок_статус'] = pd.Series(kontur_status)  # ← ИЗМЕНИЛ
    
    # Разделение на отдельные DataFrame
    kontur_df = df[['Контур_Диадок_ФИО', 'Контур_Диадок_статус']].dropna(subset=['Контур_Диадок_ФИО'])  # ← ИЗМЕНИЛ
    
    # Инициализация результатов
    results = {
        'duplicates_ad_kontur': 0,
        'internal_duplicates_kontur': 0,
        'users_to_remove_kontur': pd.DataFrame()
    }
    
    # Поиск дубликатов для Контур Диадок
    results['duplicates_ad_kontur'] = len(find_duplicates(ad_employees_df, kontur_df, 'AD_ФИО', 'Контур_Диадок_ФИО'))  # ← ИЗМЕНИЛ
    results['internal_duplicates_kontur'] = len(find_internal_duplicates(kontur_df, 'Контур_Диадок_ФИО'))  # ← ИЗМЕНИЛ
    results['users_to_remove_kontur'] = find_users_to_remove(kontur_df, ad_employees_df, ad_employees_df)
    
    return df, results