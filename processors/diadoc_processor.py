# processors/diadoc_processor.py
import pandas as pd
from utils import load_diadoc_data, find_duplicates, find_internal_duplicates, find_users_to_remove
import logging
logger = logging.getLogger(__name__)

def process_diadoc_data(df, ad_employees_df, selected_options, employee_types):
    """Обработка данных из Сфера Курьер"""  # ← ИЗМЕНИЛ комментарий
    if 2 not in selected_options and 0 not in selected_options:
        return df, {}
    
    logger.info("Обработка данных Сфера Курьер...")  # ← ИЗМЕНИЛ
    
    # Проверяем наличие необходимых данных в AD
    if ad_employees_df.empty or 'AD_ФИО' not in ad_employees_df.columns:
        logger.warning("AD DataFrame пуст или не содержит столбец 'AD_ФИО'")
        return df, {
            'duplicates_ad_diadoc': 0,
            'internal_duplicates_diadoc': 0,
            'users_to_remove_diadoc': pd.DataFrame()
        }
    
    # Загружаем данные из Сфера Курьер
    diadoc_data = load_diadoc_data()
    
    if not diadoc_data.empty:
        # Убедимся, что не превышаем MAX_ROWS
        diadoc_fio = diadoc_data['Сфера_Курьер_ФИО'][:len(df)]  # ← ИЗМЕНИЛ
        diadoc_active = diadoc_data['Сфера_Курьер_Активен'][:len(df)]  # ← ИЗМЕНИЛ
        diadoc_admin = diadoc_data['Сфера_Курьер_Администратор'][:len(df)]  # ← ИЗМЕНИЛ
        
        df['Сфера_Курьер_ФИО'] = pd.Series(diadoc_fio)  # ← ИЗМЕНИЛ
        df['Сфера_Курьер_Активен'] = pd.Series(diadoc_active)  # ← ИЗМЕНИЛ
        df['Сфера_Курьер_Администратор'] = pd.Series(diadoc_admin)  # ← ИЗМЕНИЛ
    
    # Разделение на отдельные DataFrame
    diadoc_df = df[['Сфера_Курьер_ФИО', 'Сфера_Курьер_Активен']].dropna(subset=['Сфера_Курьер_ФИО'])  # ← ИЗМЕНИЛ
    
    # Инициализация результатов
    results = {
        'duplicates_ad_diadoc': 0,
        'internal_duplicates_diadoc': 0,
        'users_to_remove_diadoc': pd.DataFrame()
    }
    
    # Поиск дубликатов для Сфера Курьер
    results['duplicates_ad_diadoc'] = len(find_duplicates(ad_employees_df, diadoc_df, 'AD_ФИО', 'Сфера_Курьер_ФИО'))  # ← ИЗМЕНИЛ
    results['internal_duplicates_diadoc'] = len(find_internal_duplicates(diadoc_df, 'Сфера_Курьер_ФИО'))  # ← ИЗМЕНИЛ
    results['users_to_remove_diadoc'] = find_users_to_remove(diadoc_df, ad_employees_df, ad_employees_df)
    
    return df, results