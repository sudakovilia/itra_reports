import json
import pandas as pd
import xlsxwriter
from datetime import datetime, timedelta


class ReportCellFormatter:

    def __init__(self, fmt_type) -> None:
        try:
            with open('formats.json', 'r') as f:
                self.color_ranges = json.load(f)[str(fmt_type)]
        except:
            raise Exception('Ошибка файла formats.json')

    def get_cell_format(self, total_hours):

        base_format = {
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'text_wrap': True
        }

        colors = {
            'white': '#FFFFFF',
            'yellow': '#FFFF00',
            'green': '#90ee90',
            'red': '#ff5050',
            'bordo': '#b00000', #'font_color': 'white'}),
            'dark_gray': '#565656' #'font_color': 'white'}),
        }
        
        for color, rng in self.color_ranges.items():
            if total_hours in range(*rng):
                color_match = color
        
        try:
            format = base_format.copy()
            format['bg_color'] = colors[color_match]
            if color_match in ('bordo', 'dark_gray'):
                format['font_color'] = 'white'
        except UnboundLocalError:
            raise Exception('В файле formats.json есть разрыв периода')

        return format

class ReportCellGenerator:

    def __init__(self, lpn, week, df) -> None:
        self.df_filtered = df.loc[(df['LPN'] == lpn) & (df['Период'] == week)]
        self.generate_cell_text()
        self.calculate_cell_total()

    def generate_cell_text(self):
        job_hours_df = self.df_filtered[['Job', 'Hours']].groupby('Job', as_index=False).sum()
        if job_hours_df.empty:
            self.text = '="0"'
        else:
            self.text = '='
            for job_name, hours in job_hours_df.values.tolist():
                self.text += f'"{job_name} ({hours:.0f})"&char(10)&'
            self.text = self.text[:-10]

    def calculate_cell_total(self):
        staff_hours_df = self.df_filtered[['Staff', 'Hours']].groupby('Staff', as_index=False).sum()
        self.total = 0 if staff_hours_df.empty else staff_hours_df['Hours'].values[0]

class DataLoader:

    def __init__(self, data_path):
        self.data_path = data_path
        self.load_data()
        self.preprocess_data()
        self.get_week_cols()
        self.get_staff_list()


    def load_data(self):
        self.raw_df = pd.read_excel(self.data_path,
                                    converters={
                                        'Период': lambda x: datetime.strptime(x, "%d.%m.%Y").date(),
                                        'LPN': str
                                        }
                                    )

    def preprocess_data(self):
        df = self.raw_df.copy()
        df['Job'] = df['Job'].str.strip()
        df['Position'] = df['Position'].str.strip()
        df['Position'] = df['Position'].fillna('')
        df['Staff'] = df['Staff'].str.replace(', ', ' ')
        report_date_from = datetime.today().date() - timedelta(weeks=1)
        df = df[df['Период'] > report_date_from]
        self.df = df

    def get_week_cols(self):
        self.week_cols = self.df['Период'].unique().tolist()

    def get_staff_list(self):
        try:
            with open('grades.json', 'r') as f:
                grades = json.load(f)
                grades_order = {key: n for (n, key) in enumerate(grades.keys())}
        except:
            raise Exception('Ошибка файла grades.json')
        
        staff_df = self.df[['LPN', 'Staff', 'Position']].drop_duplicates()
        staff_df['Grade'] = staff_df['Position'].map(grades)
        staff_df['Grade_order'] = staff_df['Position'].map(grades_order)
        staff_df.sort_values(by=['Grade_order', 'Staff'], inplace=True, ignore_index=True)
        staff_df.drop(columns=['Position',	'Grade_order'], inplace=True)
        staff_df.fillna(value='', inplace=True)
        self.staff_list = staff_df.values.tolist()

class ReportGenerator:

    def __init__(self) -> None:
        self.loader = DataLoader('./data/1c_data.xlsx')
        self.cell_formatter = ReportCellFormatter(2)
        week_name = (self.loader.week_cols[0] + timedelta(days=2)).strftime('%d-%m-%Y')
        self.save_path = f'./data/Staffing_ITRA_byPerson-w-{week_name}.xlsx'
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
        worksheet.set_column(0, 0, 25)  # ширина конолки с именанми
        worksheet.set_column(1, 1, 10)  # ширина конолки с грейдами
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
                cell = ReportCellGenerator(staff[0], week, self.loader.df)
                cell_format = self.workbook.add_format(self.cell_formatter.get_cell_format(cell.total))
                self.worksheet.write(staff_n + 1, week_n + 2, cell.text, cell_format)


if __name__ == '__main__':
    ReportGenerator()
    
    