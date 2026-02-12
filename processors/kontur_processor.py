# processors/kontur_processor.py
import pandas as pd
from utils import get_kontur_data_source, find_duplicates, find_internal_duplicates, find_users_to_remove
import logging
logger = logging.getLogger(__name__)

def process_kontur_data(df, ad_employees_df, selected_options, employee_types):
    """
    Обработка данных из Контур Диадок.
    Источник данных определяется настройкой USE_DIADOC_API в config.py.
    """
    if 3 not in selected_options and 0 not in selected_options:
        return df, {}

    logger.info("Обработка данных Контур Диадок...")

    # Проверяем наличие необходимых данных в AD
    if ad_employees_df.empty or 'AD_ФИО' not in ad_employees_df.columns:
        logger.warning("AD DataFrame пуст или не содержит столбец 'AD_ФИО'")
        return df, {
            'duplicates_ad_kontur': 0,
            'internal_duplicates_kontur': 0,
            'users_to_remove_kontur': pd.DataFrame()
        }

    # Загружаем данные из единого источника (API или Excel)
    kontur_data = get_kontur_data_source()

    if not kontur_data.empty:
        # Убедимся, что не превышаем MAX_ROWS
        kontur_fio = kontur_data['Контур_Диадок_ФИО'][:len(df)]
        kontur_admin = kontur_data['Контур_Диадок_Администратор'][:len(df)]
        kontur_status = kontur_data['Контур_Диадок_статус'][:len(df)]

        df['Контур_Диадок_ФИО'] = pd.Series(kontur_fio)
        df['Контур_Диадок_Администратор'] = pd.Series(kontur_admin)
        df['Контур_Диадок_статус'] = pd.Series(kontur_status)
    else:
        logger.warning("Данные Контур Диадок не загружены (пустой DataFrame)")

    # Разделение на отдельные DataFrame для анализа
    kontur_df = df[['Контур_Диадок_ФИО', 'Контур_Диадок_статус']].dropna(subset=['Контур_Диадок_ФИО'])

    # Инициализация результатов
    results = {
        'duplicates_ad_kontur': 0,
        'internal_duplicates_kontur': 0,
        'users_to_remove_kontur': pd.DataFrame()
    }

    # Поиск дубликатов и пользователей для удаления
    results['duplicates_ad_kontur'] = len(find_duplicates(ad_employees_df, kontur_df, 'AD_ФИО', 'Контур_Диадок_ФИО'))
    results['internal_duplicates_kontur'] = len(find_internal_duplicates(kontur_df, 'Контур_Диадок_ФИО'))
    results['users_to_remove_kontur'] = find_users_to_remove(kontur_df, ad_employees_df, ad_employees_df)

    logger.info(f"Дубликатов AD-Контур: {results['duplicates_ad_kontur']}")
    logger.info(f"Внутренних дубликатов Контур: {results['internal_duplicates_kontur']}")
    logger.info(f"Пользователей для удаления: {len(results['users_to_remove_kontur'])}")

    return df, results
