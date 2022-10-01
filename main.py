import json
import tkinter as tk
from unittest import result
import pandas as pd
import xlsxwriter
import numpy as np
from datetime import datetime, timedelta
from tkinter import messagebox as mb
from tkinter import filedialog as fd
from tkinter import ttk
from threading import Thread


class CellFormatter:

    def __init__(self, fmt_type=1) -> None:
        try:
            with open('formats.json', 'r') as f:
                self.color_ranges = json.load(f)[str(fmt_type)]
        except:
            raise Exception('Ошибка файла formats.json')
        
        self.base_format = {
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'text_wrap': True
        }

        self.colors = {
            'white': '#FFFFFF',
            'yellow': '#FFFF00',
            'green': '#90ee90',
            'red': '#ff5050',
            'bordo': '#b00000',
            'dark_gray': '#565656',
            'header_gray': '#d9d9d9'
        }

    def get_staffing_cell_format(self, total_hours):

        format = self.base_format.copy()

        for color, rng in self.color_ranges.items():
            if rng[0] <= total_hours < rng[1]:
                color_match = color
        
        try:
            format['bg_color'] = self.colors[color_match]
            if color_match in ('bordo', 'dark_gray'):
                format['font_color'] = 'white'
        except UnboundLocalError:
            raise Exception('В файле formats.json есть разрыв периода')

        return format

    def get_header_format(self, font_size=None):

        format = self.base_format.copy()

        if font_size is not None:
            format['font_size'] = font_size

        format['bg_color'] = self.colors['header_gray']
        format['bold'] = True

        return format
    
    def get_base_format(self, font_size=None, bold=False, font_color=None):

        format = self.base_format.copy()

        if font_size is not None:
            format['font_size'] = font_size

        if bold:
            format['bold'] = True

        if font_color is not None:
            format['font_color'] = font_color

        return format

class StaffingReportCellGenerator:

    def __init__(self, gpn, week, df) -> None:
        self.df_filtered = df.loc[(df['GPN'] == gpn) & (df['Период'] == week)]
        self.generate_cell_text()
        self.calculate_cell_total()

    def generate_cell_text(self):
        job_hours_df = self.df_filtered[['Job', 'Hours']].groupby('Job', as_index=False).sum()
        hours_list = [hours for _, hours in job_hours_df.values.tolist()]
        
        if job_hours_df.empty or not(any(hours_list)):
            self.text = '="0"'
        else:
            self.text = '='
            for job_name, hours in job_hours_df.values.tolist():
                if hours != 0:
                    self.text += f'"{job_name} ({hours:.0f})"&char(10)&'
                
            self.text = self.text[:-10]

    def calculate_cell_total(self):
        staff_hours_df = self.df_filtered[['Staff', 'Hours']].groupby('Staff', as_index=False).sum()
        self.total = 0 if staff_hours_df.empty else staff_hours_df['Hours'].values[0]

class StaffingDataLoader:

    def __init__(self, path_to_file):
        self.data_path = path_to_file
        self.load_data()
        self.preprocess_data()
        self.get_staff_list()
        # self.remove_old_periods_data()
        self.get_week_cols()

    def load_data(self):
        self.raw_df = pd.read_excel(self.data_path,
                                    converters={
                                        'Период': lambda x: datetime.strptime(x, "%d.%m.%Y").date(),
                                        'GPN': str,
                                        'MU': str
                                        }
                                    )

    def preprocess_data(self):
        df = self.raw_df.copy()
        df['Job'] = df['Job'].str.strip()
        df['Position'] = df['Position'].str.strip()
        df['Position'] = df['Position'].fillna('')
        df['Staff'] = df['Staff'].str.replace(', ', ' ')
        df = df[df['Staff.Suspended'] == 'Нет']
        df = df[df['MU'] == '00217']
        df = df[['GPN', 'Период', 'Job', 'Hours', 'Staff', 'Position']]
        self.df = df

    def get_week_cols(self):
        self.week_cols = self.df['Период'].unique().tolist()
        self.week_cols.sort()

    def get_staff_list(self):
        try:
            with open('grades.json', 'r') as f:
                grades = json.load(f)
                grades_order = {key: n for (n, key) in enumerate(grades.keys())}
        except:
            raise Exception('Ошибка файла grades.json')
        
        staff_df = self.df[['GPN', 'Staff', 'Position']].drop_duplicates()
        staff_df['Grade'] = staff_df['Position'].map(grades)
        staff_df['Grade_order'] = staff_df['Position'].map(grades_order)
        staff_df.sort_values(by=['Grade_order', 'Staff'], inplace=True, ignore_index=True)
        staff_df.drop(columns=['Position',	'Grade_order'], inplace=True)
        staff_df.fillna(value='', inplace=True)
        self.staff_list = staff_df.values.tolist()

    # def remove_old_periods_data(self):
    #     report_date_from = datetime.today().date() - timedelta(weeks=1)
    #     self.df = self.df[self.df['Период'] > report_date_from]

    def get_total_df(self, date_from=None, date_to=None):
        date_from -= timedelta(days=2)
        date_to -= timedelta(days=2)
        df = self.df.copy()
        if date_from is not None:
            df = df[df['Период'] >= date_from]
        if date_to is not None:
            df = df[df['Период'] <= date_to]
        df = df.groupby('GPN').sum()
        df = df.rename({'Hours': 'Staffing (total)'}, axis=1)
        return df

class StaffingReportGenerator:

    def __init__(self, selected_file_path, report_type) -> None:
        self.loader = StaffingDataLoader(selected_file_path)
        self.cell_formatter = CellFormatter(report_type)
        week_name = (self.loader.week_cols[0] + timedelta(days=2)).strftime('%d-%m-%Y')
        self.save_path = f'Staffing_ITRA_byPerson-w-{week_name}.xlsx'
        self.set_up_excel_workbook()
        self.set_formats()
        self.print_staff_info()
        self.print_week_cols()
        self.print_report_data()
        self.workbook.close()


    def set_up_excel_workbook(self):
        workbook = xlsxwriter.Workbook(self.save_path)
        worksheet = workbook.add_worksheet('Staffing_report')
        worksheet.freeze_panes(1, 2)
        worksheet.set_zoom(50)
        worksheet.set_column(0, 0, 25)  # ширина колонки с именанми
        worksheet.set_column(1, 1, 10)  # ширина колонки с грейдами
        worksheet.set_column(2, len(self.loader.week_cols) + 1, 50)  # ширина колонок с инфой
        for n in range(1, len(self.loader.staff_list) + 1):
            worksheet.set_row(n, 150)
        
        self.workbook = workbook
        self.worksheet = worksheet

    def set_formats(self):
        self.header_format = self.workbook.add_format({
            'bold': True,
            'font_size': 16,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#a9a9a9',
            'border': 1
        })
        self.spec_format = self.workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })


    def print_staff_info(self):
        self.worksheet.write(0, 0, 'Specialist', self.header_format)
        self.worksheet.write(0, 1, 'Grade', self.header_format)

        for n, staff in enumerate(self.loader.staff_list):
            self.worksheet.write(n + 1, 0, staff[1], self.spec_format)
            self.worksheet.write(n + 1, 1, staff[2], self.spec_format)

    def print_week_cols(self):
        for n, week in enumerate(self.loader.week_cols):
            week_label = (week + timedelta(days=2)).strftime('%d %B %Y')
            self.worksheet.write(0, n + 2, week_label, self.header_format)

    def print_report_data(self):
        for staff_n, staff in enumerate(self.loader.staff_list):
            for week_n, week in enumerate(self.loader.week_cols):
                cell = StaffingReportCellGenerator(staff[0], week, self.loader.df)
                cell_format = self.workbook.add_format(self.cell_formatter.get_staffing_cell_format(cell.total))
                self.worksheet.write(staff_n + 1, week_n + 2, cell.text, cell_format)

class StaffingReportGenerationThread(Thread):

    def __init__(self, selected_file_path, report_type):
        super().__init__()
        self.selected_file_path = selected_file_path
        self.report_type = report_type
        self.save_path = str()

    def run(self):
        generator = StaffingReportGenerator(self.selected_file_path, self.report_type)
        self.save_path = generator.save_path

class ChargingDataLoader:

    def __init__(self, file_path):
        self.file_path = file_path
        self.load_data()
        self.preprocess_data()
        self.filter_data()


    def load_data(self):
        self.raw_df = pd.read_excel(self.file_path,
                                    sheet_name='Details',
                                    skiprows=5,
                                    index_col=None,
                                    converters={'GPN': str})

    def preprocess_data(self):
        df = self.raw_df.copy()
        df.columns = df.columns.str.replace('\n','')
        df['Timesheet Date'] = df['Timesheet Date'].dt.date
        self.df = df

    def filter_data(self):
        self.df = self.df[self.df['Eng. Type'] == 'C']
        self.df = self.df[['GPN', 'Hrs', 'Timesheet Date']]

    def get_total_by_gpn(self, gpn, date_from=None, date_to=None):
        df = self.df.copy()
        if date_from is not None:
            df = df[df['Timesheet Date'] >= date_from]
        if date_to is not None:
            df = df[df['Timesheet Date'] <= date_to]
        df = df.groupby('GPN').sum()
        
        return float(df.loc[gpn])
    
    def get_total_df(self, date_from=None, date_to=None):
        df = self.df.copy()
        if date_from is not None:
            df = df[df['Timesheet Date'] >= date_from]
        if date_to is not None:
            df = df[df['Timesheet Date'] <= date_to]
        df = df.groupby('GPN').sum()
        df = df.rename({'Hrs': 'Charged on client codes'}, axis=1)
        return df

class StaffingVsChargingReportGenerator:
    
    def __init__(self, date_f, date_t) -> None:
        date_from=datetime.strptime(date_f, '%Y-%m-%d').date()
        date_to=datetime.strptime(date_t, '%Y-%m-%d').date()

        employee = EmployeeDataLoader('.\data\ITRA Counsellors.xlsx').df

        staffing = StaffingDataLoader('.\data\staffing.xlsx')
        staffing_total = staffing.get_total_df(date_from, date_to)
        
        charging_cyber = ChargingDataLoader('.\data\Cyber_Staff Charging Details.xlsx')
        charging_tech = ChargingDataLoader('.\data\TR_Staff Charging Details.xlsx')
        charging_total = pd.concat([charging_cyber.get_total_df(date_from, date_to),
                                    charging_tech.get_total_df(date_from, date_to)])

        report = pd.merge(employee, staffing_total, how='left', left_index=True, right_index=True, sort=False)
        report = pd.merge(report, charging_total, how='left', left_index=True, right_index=True, sort=False)
        report.fillna(value={'Charged on client codes': 0, 'Staffing (total)': 0}, inplace=True)
        report['Charging - Staffing'] = report['Charged on client codes'] - report['Staffing (total)']
        report.sort_values(by=['grade_order', 'Name'], inplace=True)
        report['Project manager'] = ''

        # вывод сверки в файл
        formatter = CellFormatter()
        output_file_name = 'Staffing vs Charging_week {}-{}.xlsx'.format(date_from.strftime("%d.%m"),
                                                                         date_to.strftime("%d.%m.%Y"))

        workbook = xlsxwriter.Workbook(output_file_name)
        worksheet = workbook.add_worksheet('Report')
        worksheet.set_zoom(65)
        worksheet.set_row(0, 66)
        worksheet.set_row(1, 46)

        for row in range(2, len(report.index) + 2):
            worksheet.set_row(row, 104)

        worksheet.set_column(0, 0, 32)
        worksheet.set_column(1, 1, 10)
        worksheet.set_column(2, 2, 32)
        worksheet.set_column(3, 3, 72)
        for col in range(4, 9):
            worksheet.set_column(col, col, 30)

        worksheet.freeze_panes(2, 0)

        # заполнение заголовка
        header_fmt = workbook.add_format(formatter.get_header_format(font_size=18))
        base_fmt = workbook.add_format(formatter.get_base_format(font_size=18))
        base_fmt_bold = workbook.add_format(formatter.get_base_format(font_size=18, bold=True))
        base_fmt_bold_red = workbook.add_format(formatter.get_base_format(font_size=18, bold=True, font_color='red'))
        first_row_text = 'Staffing vs Charging report\n{} - {}'.format(date_from.strftime("%d.%m"),
                                                                       date_to.strftime("%d.%m.%Y"))
        worksheet.merge_range('A1:I1', first_row_text, header_fmt)

        row = 1
        col = 0
        for name in ['Specialist', 'Grade', 'Counselor', 'Staffing',
                    'Staffing (total)', 'Project manager', 'Charged on client codes',
                    'Charging - Staffing', 'Comment']:
            worksheet.write(row, col, name, header_fmt)
            col += 1

        row = 2
        for gpn in report.index:

            col = 0

            # вывод специалиста
            worksheet.write(row, col, report.loc[gpn, 'Name'], base_fmt_bold)
            col += 1

            # вывод грейда
            worksheet.write(row, col, report.loc[gpn, 'Short Grade'], base_fmt)
            col += 1

            # вывод канселора
            worksheet.write(row, col, report.loc[gpn, 'Counselor'], base_fmt)
            col += 1

            # вывод стаффинга
            staffing_cell = StaffingReportCellGenerator(gpn, date_from - timedelta(days=2), staffing.df)
            staffing_cell_format = workbook.add_format(formatter.get_staffing_cell_format(staffing_cell.total))
            worksheet.write(row, col, staffing_cell.text, staffing_cell_format)
            col += 1

            # вывод стаффинга (тотал)
            worksheet.write(row, col, report.loc[gpn, 'Staffing (total)'], base_fmt)
            col += 1

            # вывод пустой колонки с манагерами
            worksheet.write(row, col, report.loc[gpn, 'Project manager'], base_fmt)
            col += 1

            # вывод чарджинга
            worksheet.write(row, col, report.loc[gpn, 'Charged on client codes'], base_fmt)
            col += 1

            # вывод разницы
            diff = float(report.loc[gpn, 'Charging - Staffing'])
            diff = round(diff, 2)

            if (diff < 0):
                fmt = base_fmt_bold_red
                res = str(diff)

            if (diff == 0):
                fmt = base_fmt_bold
                res = str(diff)

            if (diff > 0):
                fmt = base_fmt_bold
                res = "+" + str(diff)

            res = res.replace('.0', '')

            worksheet.write(row, col, res, fmt)
            col += 1

            # вывод комментария
            comment_text = 'Vacation' if 'Vacation' in staffing_cell.text else ''
            worksheet.write(row, col, comment_text, base_fmt)
            col += 1

            # вывод одной строки закончен, переходим к следующей
            row += 1

        workbook.close()

class EmployeeDataLoader:

    def __init__(self, path_to_file) -> None:
        grades_df = pd.read_excel(path_to_file, sheet_name='Grades', index_col=None)
        grades = {v[0]: v[1] for v in grades_df.values}
        grades_order = {key: n for (n, key) in enumerate(grades.keys())}
        
        data_df = pd.read_excel(path_to_file, converters={'GPN': str}, sheet_name='Data')
        data_df.set_index('GPN', inplace=True)
        data_df['grade_order'] = data_df['Grade'].map(grades_order)
        data_df.sort_values(by=['grade_order', 'Name'], inplace=True)

        self.df = data_df

class View(tk.Tk):

    def __init__(self, master=None) -> None:
        super().__init__()

        # TODO: добавить проверку наличия файлов grades.json и formats.json

        self.title('ITRA reports')
        self.minsize(400, 200)

        self.main_frame = ttk.Frame(self)
        self.main_frame.pack(side='top', fill='both', expand=True, padx=10, pady=10)
        self.main_frame.rowconfigure(0, weight=2)
        self.main_frame.rowconfigure(1, weight=1)
        self.main_frame.rowconfigure(2, weight=1)
        self.main_frame.rowconfigure(3, weight=1)
        self.main_frame.rowconfigure(4, weight=4)
        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.columnconfigure(1, weight=4)

        select_file_label = ttk.Label(self.main_frame, text='Выгрузка из 1C')
        select_file_label.grid(row=0, column=0, sticky='w')
        self.selected_file_path = tk.StringVar()
        select_file_button = ttk.Button(self.main_frame, text='Выбрать файл', command=self.select_file)
        select_file_button.grid(row=0, column=1, sticky='we')

        report_type_label = ttk.Label(self.main_frame, text='Тип отчета:')
        report_type_label.grid(row=1, column=0, sticky='w')

        self.report_type = tk.IntVar()
        self.report_type.set(1)
        
        rb0 = ttk.Radiobutton(self.main_frame, text='Формальный: белый, желтый, зеленый, красный', variable=self.report_type, value=1)
        rb0.grid(row=2, column=0, columnspan=2, sticky='w')

        rb1 = ttk.Radiobutton(self.main_frame, text='Внутренний мониторинг: вариант 1 + бордовый и темно-серый', variable=self.report_type, value=2)
        rb1.grid(row=3, column=0, columnspan=2, sticky='w')

        self.generate_report_button = ttk.Button(self.main_frame, text='Сформировать отчет', width=40, command=self.generate_report)
        self.generate_report_button.grid(row=4, column=0, columnspan=2)
        self.generate_report_button['state'] = 'disabled'

    def select_file(self):
        filetypes = (
            ('Excel files', '*.xlsx'),
            ('All files', '*.*')
        )

        file_name = fd.askopenfilename(
            title='Выберите файл выгрузки из 1C',
            initialdir='.',
            filetypes=filetypes)

        if file_name == '':
            pass
        else:
            self.selected_file_path.set(file_name)
            self.generate_report_button['state'] = 'normal'

    def generate_report(self):
        self.generate_report_button['state'] = 'disabled'
        self.main_frame.config(cursor='wait')
        thread = StaffingReportGenerationThread(self.selected_file_path.get(), self.report_type.get())
        thread.start()
        self.monitor(thread)


    def monitor(self, thread):
        if thread.is_alive():
            self.after(100, lambda: self.monitor(thread))
        else:
            self.generate_report_button['state'] = 'normal'
            self.main_frame.config(cursor='')
            mb.showinfo(title='Готово', message=f'Отчет сохранен в файл {thread.save_path}')
            
        
    def main(self):
        self.mainloop()


if __name__ == '__main__':
    view = View()
    view.main()

    # dates = [
    #     # ['2022-08-29', '2022-09-02'],
    #     # ['2022-09-05', '2022-09-09'],
    #     ['2022-09-12', '2022-09-16']
    # ]

    # for d in dates:
    #     print(d)
    #     StaffingVsChargingReportGenerator(*d)