import pandas as pd
import numpy as np
import xlrd
import xlsxwriter
import time
import datetime
import ast

from re import sub
from os import system
from glob import glob
from sys import exit

clear = lambda: system('cls')


def get_report_cell(gpn, date, df):
    df = df.loc[[gpn], ['Engagement name', date]]
    df = df.loc[df[date] != 0]
    df = df.groupby('Engagement name', as_index=False).sum()  # группировка для суммирования дубликатов строк

    if not df.values.tolist():
        return '="0"'

    res = '='
    for name, hours in df.values.tolist():
        res += '"'
        res += sub(r' - \d{8}', '', str(name))
        res += ' ({0:g}'.format(hours)
        res += ')"&char(10)&'
    res = str(res[:-10])
    return res


def generate_report(source_file_name):
    # загрузка исходника в dataframe
    df = pd.read_excel(source_file_name, sheet_name='Daten Retain')

    # замена кода дожности стажерам
    df.loc[df['*Current grade'] == 'Client Serving Contractor 1', 'Current Grade Order'] = 500

    # сортировка по коду грейда
    # df.sort_values(by=['Current Grade Order', '*Resource name'], inplace=True, ascending=True, ignore_index=True)

    # замена наименования должностей
    grades = {
        "Partner/Principal 1": "P",
        "Senior Manager 3": "SM3+",
        "Senior Manager 2": "SM2",
        "Senior Manager 1": "SM1",
        "Manager 3": "M3+",
        "Manager 2": "M2",
        "Manager 1": "M1",
        "Senior 3": "S3+",
        "Senior 2": "S2",
        "Senior 1": "S1",
        "Staff/Assistant 3": "ASt",
        "Staff/Assistant 2": "ASt",
        "Staff/Assistant 1": "St",
        "Client Serving Contractor 1": "Int"
    }

    for g in grades.keys():
        df.loc[df['*Current grade'] == g, '*Current grade'] = grades[g]

    # split имени сотрудника и gpn
    df[['Specialist', 'GPN']] = df['*Resource name'].str.split(' - ', expand=True)
    df.Specialist = df.Specialist.str.replace(',', '')
    df.drop(columns=['*Resource name'], inplace=True)

    # замена порядка колонок
    cols = df.columns.tolist()
    cols_new_order = cols[0:5] + cols[-2:] + cols[5:-2]
    df = df[cols_new_order]

    # переименование колонок
    rename_dict = {}  # обрезка "Time (Hours) " с использованием словаря
    for i in cols_new_order[7:]:
        rename_dict[i] = str(i[13:])

    rename_dict['*Current grade'] = 'Grade'
    rename_dict['Current Grade Order'] = 'Grade code'
    rename_dict['*Sub-management unit'] = 'Sub-management unit'
    rename_dict['*Engagement name'] = 'Engagement name'

    df.rename(columns=rename_dict, inplace=True)

    # создание тотал репорта (общее количество часов у сотрудника в неделю)
    df_total = df.groupby(by=['GPN', 'Specialist', 'Grade', 'Grade code'], as_index=False).sum()
    df_total.sort_values(by=['Grade code', 'Specialist'], inplace=True, ignore_index=True)
    df_total.drop(columns=['Engagement number'], inplace=True)
    df_total.set_index('GPN', inplace=True)

    # создание репорта (разбивка часов по проектам в виде длинной строки)
    df_report = df_total.copy()
    date_cols = df_report.columns.to_list()[3:]

    for i in date_cols:  # присвоение пустых значений для дат
        df_report[i] = np.nan

    df.set_index('GPN', inplace=True)

    for gpn in df_report.index.to_list():  # цикл по заполнению клеток отчета
        for date in df_report.columns.to_list()[3:]:
            df_report.loc[gpn, date] = get_report_cell(gpn, date, df)

    df_report = df_report.replace(r'^\s*$', 0, regex=True)

    return df_report, df_total


if __name__ == '__main__':
    while True:
        clear()
        print('SQUARE tool v0.7\n')
        print('Выберите необходимую функцию:')
        print('1. Преобразование формата отчета Retain')
        print('2. Сверка Staffing vs Charging')

        answer = input("\nВведите номер варианта --> ")

        if answer.isdigit():
            answer = int(answer)
            if (answer in (1, 2)):
                break

        print("Введен недопустимый номер, повторите ввод...")
        time.sleep(1)
        continue

    if (answer == 1):
        # преобразование формата отчета

        while True:
            clear()
            print('Выберите вариант раскраски отчета:')
            print('1. Формальный: белый + желтый + зеленый + красный')
            print('2. Внутренний мониторинг: вариант 1 + бордовый + темно-серый')

            answer = input("Введите номер варианта --> ")

            if answer.isdigit():
                answer = int(answer)
                if (answer in (1, 2)):
                    break

            print("Введен недопустимый номер, повторите ввод...")
            time.sleep(1)
            continue

        try:
            print('Загружаю исходные данные отчета...')

            staffing_file_name = glob('*Retain*.xls')[0]
            staffing_df = pd.read_excel(staffing_file_name, sheet_name='Daten Retain')
            print('Стаффинг загружен --> ' + staffing_file_name)

        except IndexError:
            print('Ошибка. Проверьте наличие всех необходимых для сверки файлов.')
            exit()

        clear()
        print("Приступаю к обработке отчета...")

        # формирование отчета по стаффингу
        df_report, df_total_count = generate_report(staffing_file_name)
        df_report.set_index('Specialist', inplace=True)
        df_report.drop(columns=['Grade code'], inplace=True)
        df_total_count.set_index('Specialist', inplace=True)

        # вывод в эксель файл с условным форматирвоанием
        output_file = 'Staffing_ITRA_byPerson-w-.xlsx'
        print('Вывожу результат в новый файл -> ' + output_file)
        workbook = xlsxwriter.Workbook(output_file)
        worksheet = workbook.add_worksheet('Staffing_report')

        worksheet.set_zoom(50)

        r_cols = list(df_report.columns)
        r_index = list(df_report.index.unique())

        # заполнение заголовка
        header = workbook.add_format({
            'bold': True,
            'font_size': 16,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#a9a9a9',
            'border': 1})

        worksheet.write('A1', 'Specialist', header)
        row = 0
        col = 1
        for name in r_cols:
            worksheet.write(row, col, name, header)
            col += 1

        # заполнение специалиста и грейда
        spec = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'border': 1})
        worksheet.set_column(0, 0, 25)  # ширина конолки с именанми
        worksheet.set_column(1, 1, 10)  # ширина конолки с грейдами
        row = 1
        col = 0
        for ind in r_index:
            worksheet.write(row, col, ind, spec)
            row += 1

        row = 1
        col = 1
        for grade in list(df_report.loc[:, 'Grade']):
            worksheet.write(row, col, grade, spec)
            row += 1

        # заполнение инфы по стаффингу
        worksheet.set_column(2, len(r_cols), 50)  # ширина колонок с инфой
        # worksheet.set_row(1, len(r_index), 290) # высота колонок с инфорй

        formats = {
            'white': workbook.add_format({
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'text_wrap': True}),
            'yellow': workbook.add_format({
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'bg_color': '#FFFF00',
                'text_wrap': True}),
            'green': workbook.add_format({
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'bg_color': '#90ee90',
                'text_wrap': True}),
            'red': workbook.add_format({
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'bg_color': '#ff5050',
                'text_wrap': True}),
            'bordo': workbook.add_format({
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'bg_color': '#b00000',
                'text_wrap': True,
                'font_color': 'white'}),
            'dark_gray': workbook.add_format({
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'bg_color': '#565656',
                'text_wrap': True,
                'font_color': 'white'})
        }

        print("Применяю " + str(answer) + " вариант раскраски отчета")

        with open('fmt.txt', 'r') as f:
            fmt_dict = ast.literal_eval(f.read())
            fmt_settings = fmt_dict[answer]

        row = 1
        for name in r_index:
            col = 2
            for date in r_cols[1:]:
                total = int(df_total_count.loc[name, date])

                for clr, rng in fmt_settings.items():
                    if total in range(rng[0], rng[1]):
                        format = formats[clr]

                worksheet.write(row, col, str(df_report.loc[name, date]), format)
                col += 1
            row += 1

        for n in range(1, len(r_index) + 1):
            worksheet.set_row(n, 200)

        workbook.close()

        print('Отчет успешно сформирован!')
        input('Нажмите любую клавишу...')

        exit()

    if (answer == 2):
        clear()
        # загрузка данных для сверки
        try:
            print('Загружаю документы для сверки...')

            counsellors_file_name = glob('*Counsellors*.xls')[0]
            counsellors_df = pd.read_excel(counsellors_file_name, sheet_name='Data')
            print('Справочник канселоров загружен --> ' + counsellors_file_name)

            staffing_file_name = glob('*Retain*.xls')[0]
            staffing_df = pd.read_excel(staffing_file_name, sheet_name='Daten Retain')
            print('Стаффинг загружен --> ' + staffing_file_name)

            charging_cyber_file_name = glob('*Staff Charging Details_Cyber Security*.xls')[0]
            charging_cyber_df = pd.read_excel(charging_cyber_file_name, sheet_name='Staff_Charging', skiprows=10,
                                              index_col=None)
            print('Чарджинг Cyber загружен --> ' + charging_cyber_file_name)

            charging_tech_file_name = glob('*Staff Charging Details_Technology Risk*.xls')[0]
            charging_tech_df = pd.read_excel(charging_tech_file_name, sheet_name='Staff_Charging', skiprows=10,
                                             index_col=None)
            print('Чарджинг Technology загружен --> ' + charging_tech_file_name)

        except IndexError:
            print('Ошибка. Проверьте наличие всех необходимых для сверки файлов.')
            exit()

            time.sleep(2)

        # получение периода для формирования отчета
        clear()
        while True:
            try:
                print('Введите дату начала недели в формте ДД-ММ-ГГГГ')
                answer = str(input("Дата --> "))
                start_date = datetime.datetime.strptime(answer, '%d-%m-%Y').date()
                break
            except ValueError:
                print("Введено недопустимое значение, повторите ввод...")
                time.sleep(1)
                clear()
                continue

        clear()
        while True:
            try:
                print('Введите дату конца недели в формте ДД-ММ-ГГГГ')
                answer = str(input("Дата --> "))
                end_date = datetime.datetime.strptime(answer, '%d-%m-%Y').date()
                break
            except ValueError:
                print("Введено недопустимое значение, повторите ввод...")
                time.sleep(1)
                clear()
                continue

        if (start_date > end_date):
            print("Ошибка. Дата конца раньше даты начала!")
            exit()
        elif ((end_date - start_date).days > 5):
            print("Ошибка. Интервал превышает 5 дней!")
            exit()
        else:
            clear()
            print("Выполняю сверку на неделе с " + start_date.strftime("%d.%m") + " по " + end_date.strftime("%d.%m"))

        # временное присвоение дат
        # start_date = datetime.datetime.strptime('29-03-2021', '%d-%m-%Y').date()
        # end_date = datetime.datetime.strptime('02-04-2021', '%d-%m-%Y').date()

        # формирование отчета по стаффингу
        staffing_report_df, staffing_report_total_df = generate_report(staffing_file_name)

        # оставляем в сформированных отчетах только неделю с началом в start_date
        columns = list(staffing_report_df.columns)
        dates = columns[3:]  # выбираем все колонки с датами
        tmp_dict = {}  # создаем временный словарь

        for i in dates:
            tmp_dict[datetime.datetime.strptime(i.split(' - ')[0], '%d %b %Y').date()] = i

        staffing_report_df = staffing_report_df[['Specialist', 'Grade', 'Grade code', str(tmp_dict[start_date])]]

        # в total отчете тоже берем только интересующую нас неделю
        staffing_report_total_df = staffing_report_total_df[[str(tmp_dict[start_date])]]

        # в total отчете переименовываем неделю в total
        column_name = str(staffing_report_total_df.columns.to_list()[0])
        staffing_report_total_df.rename(columns={column_name: 'Total'}, inplace=True)

        # подгрузка канселора
        counsellors_df = counsellors_df[["GPN", "Counselor Name"]]
        counsellors_df.set_index("GPN", inplace=True)

        # подгрузка чарджинга
        charging_cyber_df = charging_cyber_df[charging_cyber_df['Eng. \nType'] == 'C']
        charging_cyber_df = charging_cyber_df[['GPN', 'Hrs', 'Timesheet \nDate']]
        charging_cyber_df['Timesheet \nDate'] = charging_cyber_df['Timesheet \nDate'].dt.date

        charging_cyber_df = charging_cyber_df[
            (charging_cyber_df['Timesheet \nDate'] >= start_date) & (charging_cyber_df['Timesheet \nDate'] <= end_date)]

        charging_cyber_df = charging_cyber_df.groupby('GPN').sum()

        charging_tech_df = charging_tech_df[charging_tech_df['Eng. \nType'] == 'C']
        charging_tech_df = charging_tech_df[['GPN', 'Hrs', 'Timesheet \nDate']]
        charging_tech_df['Timesheet \nDate'] = charging_tech_df['Timesheet \nDate'].dt.date

        charging_tech_df = charging_tech_df[
            (charging_tech_df['Timesheet \nDate'] >= start_date) & (charging_tech_df['Timesheet \nDate'] <= end_date)]

        charging_tech_df = charging_tech_df.groupby('GPN').sum()

        charging_union = pd.concat([charging_cyber_df, charging_tech_df])

        # формируем финальный отчет по сверке
        result = pd.merge(staffing_report_df, staffing_report_total_df, how='left', left_index=True, right_index=True,
                          sort=False)
        result = pd.merge(result, counsellors_df, how='left', left_index=True, right_index=True, sort=False)
        result = pd.merge(result, charging_union, how='left', left_index=True, right_index=True, sort=False)

        result.fillna(value={'Hrs': 0}, inplace=True)

        result.sort_values(by=['Grade code', 'Specialist'], inplace=True, ascending=True)

        result['Charging - Staffing'] = result['Hrs'] - result['Total']
        result['Project manager'] = ''

        # меняем имя и порядок колонок
        rename_columns = {}
        rename_columns[str(result.columns.to_list()[3])] = 'Staffing'
        rename_columns['Counselor Name'] = 'Counselor'
        rename_columns['Total'] = 'Staffing (total)'
        rename_columns['Hrs'] = 'Charged on client codes'

        result.rename(columns=rename_columns, inplace=True)

        new_column_order = ['Grade code', 'Specialist', 'Grade', 'Counselor', 'Staffing', 'Staffing (total)',
                            'Project manager', 'Charged on client codes', 'Charging - Staffing']
        result = result[new_column_order]

        # добавляем комменты для интернов и отпусков
        result['Comment'] = np.where(result['Staffing'].str.contains('Vacation'), 'Vacation', '')
        result.loc[result['Grade'] == 'Int', 'Comment'] = 'Intern'

        # удаляем лишних людей
        to_delete = ['Masiagina Viktoria', 'Tokareva Ekaterina']
        for i in to_delete:
            result.drop(result[result.Specialist == i].index, inplace=True)

        # result.to_csv('res_tmp.csv')

        # вывод сверки в файл
        output_file_name = 'Staffing vs Charging_week ' + start_date.strftime("%d.%m") + '-' + end_date.strftime(
            "%d.%m.%Y") + '.xlsx'
        print('Вывожу результат в новый файл -> ' + output_file_name)

        workbook = xlsxwriter.Workbook(output_file_name)
        worksheet = workbook.add_worksheet('Report')

        r_cols = list(result.columns)
        r_index = list(result.index)

        worksheet.set_zoom(65)
        worksheet.set_row(0, 66)
        worksheet.set_row(1, 46)

        for row in range(2, len(r_index) + 2):
            worksheet.set_row(row, 104)

        worksheet.set_column(0, 0, 32)
        worksheet.set_column(1, 1, 10)
        worksheet.set_column(2, 2, 32)
        worksheet.set_column(3, 3, 72)
        for col in range(4, 9):
            worksheet.set_column(col, col, 30)

        worksheet.freeze_panes(2, 0)

        # заполнение заголовка
        header_fmt = workbook.add_format({
            'bold': True,
            'font_size': 18,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#d9d9d9',
            'border': 1,
            'text_wrap': True})

        first_row = 'Staffing vs Charging report\n' + start_date.strftime("%d.%m") + ' - ' + end_date.strftime(
            "%d.%m.%Y")
        worksheet.merge_range('A1:I1', first_row, header_fmt)

        row = 1
        col = 0
        for name in r_cols[1:]:
            worksheet.write(row, col, name, header_fmt)
            col += 1

        # вывод основной части отчета
        formats = {
            'spec_fmt': workbook.add_format({
                'bold': True,
                'font_size': 18,
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'text_wrap': True}),
            'spec_fmt_red': workbook.add_format({
                'bold': True,
                'font_size': 18,
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'text_wrap': True,
                'font_color': 'red'}),
            'main_fmt': workbook.add_format({
                'bold': False,
                'font_size': 18,
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'text_wrap': True}),
            'white': workbook.add_format({
                'align': 'center',
                'valign': 'vcenter',
                # 'font_size': 9,
                'border': 1,
                'text_wrap': True}),
            'yellow': workbook.add_format({
                'align': 'center',
                'valign': 'vcenter',
                # 'font_size': 9,
                'border': 1,
                'bg_color': '#FFFF00',
                'text_wrap': True}),
            'green': workbook.add_format({
                'align': 'center',
                'valign': 'vcenter',
                # 'font_size': 9,
                'border': 1,
                'bg_color': '#90ee90',
                'text_wrap': True})
        }

        with open('fmt.txt', 'r') as f:
            fmt_dict = ast.literal_eval(f.read())
            fmt_settings = fmt_dict[3]

        row = 2
        for gpn in r_index:

            col = 0

            # вывод специалиста
            worksheet.write(row, col, result.loc[gpn, 'Specialist'], formats['spec_fmt'])
            col += 1

            # вывод грейда
            worksheet.write(row, col, result.loc[gpn, 'Grade'], formats['main_fmt'])
            col += 1

            # вывод канселора
            try:
                worksheet.write(row, col, result.loc[gpn, 'Counselor'], formats['main_fmt'])
            except:
                worksheet.write(row, col, '', formats['main_fmt'])

            col += 1

            # вывод стаффинга
            total = int(result.loc[gpn, 'Staffing (total)'])
            for clr, rng in fmt_settings.items():
                if total in range(rng[0], rng[1]):
                    format = formats[clr]

            worksheet.write(row, col, result.loc[gpn, 'Staffing'], format)
            col += 1

            # вывод стаффинга (тотал)
            worksheet.write(row, col, result.loc[gpn, 'Staffing (total)'], formats['main_fmt'])
            col += 1

            # вывод пустой колонки с манагерами
            worksheet.write(row, col, result.loc[gpn, 'Project manager'], formats['main_fmt'])
            col += 1

            # вывод чарджинга
            worksheet.write(row, col, result.loc[gpn, 'Charged on client codes'], formats['main_fmt'])
            col += 1

            # вывод разницы
            diff = float(result.loc[gpn, 'Charging - Staffing'])
            diff = round(diff, 2)

            if (diff < 0):
                fmt = formats['spec_fmt_red']
                res = str(diff)

            if (diff == 0):
                fmt = formats['spec_fmt']
                res = str(diff)

            if (diff > 0):
                fmt = formats['spec_fmt']
                res = "+" + str(diff)

            res = res.replace('.0', '')

            worksheet.write(row, col, res, fmt)
            col += 1

            # вывод комментария
            worksheet.write(row, col, result.loc[gpn, 'Comment'], formats['main_fmt'])
            col += 1

            # вывод одной строки закончен, переходим к следующей
            row += 1

        workbook.close()
        print('Сверка успешно сформирована!')
        input('Нажмите любую клавишу...')

        exit()
