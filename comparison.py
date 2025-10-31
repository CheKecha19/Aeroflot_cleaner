# comparison.py
import pandas as pd
from utils import normalize_name
import logging
logger = logging.getLogger(__name__)

def find_duplicates(df1, df2, col1, col2):
    """Поиск дубликатов между двумя DataFrame"""
    names1 = set(df1[col1].apply(normalize_name).dropna())
    names2 = set(df2[col2].apply(normalize_name).dropna())
    
    return names1.intersection(names2)

def find_internal_duplicates(df, column):
    """Поиск дубликатов внутри одного столбца"""
    normalized_names = df[column].apply(normalize_name)
    value_counts = normalized_names.value_counts()
    return set(value_counts[value_counts > 1].index)

def find_users_to_remove(edo_df, staff_df, gph_df):
    """Поиск пользователей для удаления из ЭДО"""
    staff_names = set(staff_df.iloc[:, 0].apply(normalize_name).dropna())
    gph_names = set(gph_df.iloc[:, 0].apply(normalize_name).dropna())
    all_valid_names = staff_names.union(gph_names)
    
    users_to_remove = []
    
    for _, row in edo_df.iterrows():
        # Первый столбец - ФИО
        fio_column = edo_df.columns[0]
        normalized_name = normalize_name(row[fio_column])
        
        # Проверяем условия для удаления (нет в AD и активен/не заблокирован)
        if normalized_name not in all_valid_names:
            # Для Контура проверяем дату блокировки
            if 'Контур_Дата_блокировки' in edo_df.columns and pd.isna(row['Контур_Дата_блокировки']):
                users_to_remove.append(row)
            # Для Диадока проверяем активность
            elif 'Диадок_Активен' in edo_df.columns and row['Диадок_Активен'] == 'Да':
                users_to_remove.append(row)
            # Для 1С проверяем активность
            elif '1C_Активен' in edo_df.columns and row['1C_Активен'] == 'Да':
                users_to_remove.append(row)
    
    return pd.DataFrame(users_to_remove)