from logging import StrFormatStyle
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import json

def get_staffing_report_cell(lpn, week, df):
    df = df.loc[(df['LPN'] == lpn) & (df['Период'] == week)]
    df = df.groupby('Job', as_index=False).sum()  # группировка для суммирования дубликатов строк
    df = df[['Job', 'Hours']]

    if df.empty:
        return '="0"'

    report_cell_value = '='
    for job_name, hours in df.values.tolist():
        report_cell_value += f'"{job_name} ({hours:.0f})"&char(10)&'
    report_cell_value = report_cell_value[:-10]
    return report_cell_value



if __name__ == '__main__':
    # загрузка данных из xlsx
    staffing_df_raw = pd.read_excel('./data/1c_data.xlsx',
                                    converters={
                                        'Период': lambda x: datetime.strptime(x, "%d.%m.%Y").date(),
                                        'LPN': str
                                        }
                                    )
    
    staffing_df = staffing_df_raw.copy()

    # data preprocessing
    staffing_df['Job'] = staffing_df['Job'].str.strip()
    staffing_df['Position'] = staffing_df['Position'].str.strip()
    staffing_df['Position'] = staffing_df['Position'].fillna('')
    staffing_df['Staff'] = staffing_df['Staff'].str.replace(', ', ' ')
    
    # убираем данные старше чем недельной давности
    report_date_from = datetime.today().date() - timedelta(weeks=1)
    staffing_df = staffing_df[staffing_df['Период'] > report_date_from]
    
    # получаем уникальные даты для header
    week_cols = staffing_df['Период'].unique()

    # получаем уникальных сотрудников
    staff = staffing_df[['LPN', 'Staff', 'Position']].drop_duplicates()
    
    with open('grades.json', 'r') as f:
        grades = json.load(f)
        grades_order = {key: n for (n, key) in enumerate(grades.keys())}
    
    staff['Grade'] = staff['Position'].map(grades)
    staff['Grade_order'] = staff['Position'].map(grades_order)
    staff.sort_values(by=['Grade_order', 'Staff'], inplace=True, ignore_index=True)
    staff.drop(columns=['Position',	'Grade_order'], inplace=True)

    # # формируем отчет
    for week in week_cols:
        for lpn in staff['LPN']:
            res = get_staffing_report_cell(lpn, week, staffing_df)
            week_plus_2_days = week + timedelta(days=2) # в отчет нужно выводить дату понедельника, а не субботы
            week_plus_2_days_fmt = week_plus_2_days.strftime('%d %B %Y')
            staff.loc[staff['LPN'] == lpn, week_plus_2_days_fmt] = res
    
    staff.to_excel('unformatted_report.xlsx', index=False)
    