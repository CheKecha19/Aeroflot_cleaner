# processors/diadoc_processor.py
import pandas as pd
from utils import get_diadoc_data_source, find_duplicates, find_internal_duplicates, find_users_to_remove
import logging
logger = logging.getLogger(__name__)

def process_diadoc_data(df, ad_employees_df, selected_options, employee_types):
    """
    Обработка данных из Сфера Курьер.
    Источник данных определяется настройкой USE_SFERA_API в config.py.
    """
    if 2 not in selected_options and 0 not in selected_options:
        return df, {}

    logger.info("Обработка данных Сфера Курьер...")

    # Проверяем наличие необходимых данных в AD
    if ad_employees_df.empty or 'AD_ФИО' not in ad_employees_df.columns:
        logger.warning("AD DataFrame пуст или не содержит столбец 'AD_ФИО'")
        return df, {
            'duplicates_ad_diadoc': 0,
            'internal_duplicates_diadoc': 0,
            'users_to_remove_diadoc': pd.DataFrame()
        }

    # Загружаем данные из единого источника (API или Excel)
    diadoc_data = get_diadoc_data_source()

    if not diadoc_data.empty:
        # Убедимся, что не превышаем MAX_ROWS
        diadoc_fio = diadoc_data['Сфера_Курьер_ФИО'][:len(df)]
        diadoc_active = diadoc_data['Сфера_Курьер_Активен'][:len(df)]
        diadoc_admin = diadoc_data['Сфера_Курьер_Администратор'][:len(df)]

        df['Сфера_Курьер_ФИО'] = pd.Series(diadoc_fio)
        df['Сфера_Курьер_Активен'] = pd.Series(diadoc_active)
        df['Сфера_Курьер_Администратор'] = pd.Series(diadoc_admin)
    else:
        logger.warning("Данные Сфера Курьер не загружены (пустой DataFrame)")

    # Разделение на отдельные DataFrame для анализа
    diadoc_df = df[['Сфера_Курьер_ФИО', 'Сфера_Курьер_Активен']].dropna(subset=['Сфера_Курьер_ФИО'])

    # Инициализация результатов
    results = {
        'duplicates_ad_diadoc': 0,
        'internal_duplicates_diadoc': 0,
        'users_to_remove_diadoc': pd.DataFrame()
    }

    # Поиск дубликатов и пользователей для удаления
    results['duplicates_ad_diadoc'] = len(find_duplicates(ad_employees_df, diadoc_df, 'AD_ФИО', 'Сфера_Курьер_ФИО'))
    results['internal_duplicates_diadoc'] = len(find_internal_duplicates(diadoc_df, 'Сфера_Курьер_ФИО'))
    results['users_to_remove_diadoc'] = find_users_to_remove(diadoc_df, ad_employees_df, ad_employees_df)

    logger.info(f"Дубликатов AD-Сфера: {results['duplicates_ad_diadoc']}")
    logger.info(f"Внутренних дубликатов в Сфере: {results['internal_duplicates_diadoc']}")
    logger.info(f"Пользователей для удаления: {len(results['users_to_remove_diadoc'])}")

    return df, results
