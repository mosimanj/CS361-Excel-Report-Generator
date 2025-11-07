import pandas
import openpyxl # needed for generating Excel
from random import randint
from datetime import date
import os

#TODO: add doc strings
class ReportGenerator:
    def __init__(self, report_data):
        self.dataframe = pandas.DataFrame(report_data['Rows'], columns=report_data['Headers'])
        self.sort_by = False if 'sort_by' not in report_data else report_data['sort_by']
        self.style = False if 'style' not in report_data else report_data['style']

    def needs_style(self):
        return True if self.style else False

    def needs_sort(self):
        return True if self.sort_by else False

    #TODO: complete sort_report(), style_report()
    def sort_report(self):
        pass

    def style_report(self):
        pass

    def generate_excel(self):
        file_name = f'reports\\{date.today()}-report-{randint(300,9000)}.xlsx'
        self.dataframe.to_excel(file_name, index=False)
        cur_directory = os.path.abspath(__file__)
        return f'{cur_directory}\\{file_name}'