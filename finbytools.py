import pandas as pd
import numpy as np


def get_table_data(file, sheet_name, table_name):
    sheet = file[sheet_name]
    table_obj = sheet.tables[table_name]
    table_refs = sheet[table_obj.ref]

    data = []
    for row in table_refs:
        cols = []
        for col in row:
            cols.append(col.value)
        data.append(cols)

    df = pd.DataFrame(
        data=data[1:], 
        index=None,
        columns=data[0]
    ).fillna(value=np.nan)

    return df


def get_executor_perc(outer_df, inner_df, filter_by_project, filter_by_task):
    # Проверка на NULL project и task
    temp_df = inner_df[
        ( inner_df['Проект'] == filter_by_project if not pd.isnull(filter_by_project) else pd.isnull(inner_df['Проект']) ) & 
        ( inner_df['Задача'] == filter_by_task if not pd.isnull(filter_by_task) else pd.isnull(inner_df['Задача']) )
    ]
    temp_df = temp_df.merge(outer_df, on='Проект', how='left')

    # replace null to '' for startswith func
    filter_by_task = '' if pd.isnull(filter_by_task) else filter_by_task

    # _x = outer, _y = inner
    temp_df = temp_df[
        temp_df.apply(
            lambda row: True if pd.isnull(row['Задача_y']) else filter_by_task.startswith(row['Задача_y']),  # is np.nan
            axis=1
        )
    ]

    # select columns and get distinct values
    temp_df = temp_df[['Проект', 'Задача_x', 'Задача_y',
                    'Постановщик', '% выполнения (факт)']].drop_duplicates()

    return [temp_df.iloc[0]['Постановщик'], temp_df.iloc[0]['% выполнения (факт)']] if len(temp_df) > 0 else [np.nan, np.nan]


def get_planner_perc(df, filter_by_planner, filter_by_executor, filter_by_project, filter_by_task):
    filter_by_task = '' if pd.isnull(filter_by_task) else filter_by_task

    temp_df = df[
        df.apply(
            lambda row: 
                ( row['Постановщик'] == filter_by_planner if not pd.isnull(filter_by_planner) else pd.isnull(row['Постановщик']) ) and
                ( row['Исполнитель'] == filter_by_executor if not pd.isnull(filter_by_executor) else pd.isnull(row['Исполнитель']) ) and
                ( row['Проект'] == filter_by_project if not pd.isnull(filter_by_project) else pd.isnull(row['Проект']) ) and
                ( True if pd.isnull(row['Задача']) else filter_by_task.startswith(row['Задача']) ), #  is np.nan
            axis=1
        )
    ]
    
    temp_df = temp_df.sort_values(by='Задача', ascending=False)

    return temp_df.iloc[0]['% выполнения (Постановщик)'] if len(temp_df) > 0 else np.nan