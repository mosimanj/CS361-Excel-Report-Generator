import pandas
import openpyxl # needed for generating Excel
from random import randint
from datetime import date
import os

class ReportGenerator:
    """
    Generates Excel report by creating a DataFrame from inputted dictionary. Allows sorting by a specified column and
    application of templated styling.

    :param report_data: Dictionary containing the headers and rows to be turned into a DataFrame, along with indicators
    for what header the report should be sorted by (if any) and what style template should be applied (if any).
    """
    def __init__(self, report_data):
        self.dataframe = pandas.DataFrame(report_data['Rows'], columns=report_data['Headers']) #Convert to DataFrame
        self.sort_by = False if 'sort_by' not in report_data else report_data['sort_by']
        self.style = False if 'style' not in report_data else report_data['style']

    def needs_style(self):
        """
        Indicates whether styling needs to be applied to the report, based on the original input dictionary.
        :return:    True if styling is needed, False if not.
        """
        return True if self.style else False

    def needs_sort(self):
        """
        Indicates whether sorting needs to be applied to the report, based on original input dictionary.
        :return:    True if sorting is needed, False if not.
        """
        return True if self.sort_by else False

    def sort_report(self):
        """
        Sorts the DataFrame by the column indicated by the sort_by value of original input dictionary. The indicated
        column will be sorted in ascending order.
        :return:    None.
        """
        self.dataframe = self.dataframe.sort_values(by=self.sort_by)

    #TODO: style_report() + docstring
    #TODO: After creating style_report() and associated templates, add documentation (incl. screenshots) to readme.
    def style_report(self):
        pass

    def generate_excel(self):
        """
        Generates an Excel file from the DataFrame, stores it in the reports directory, and returns the created files
        absolute filepath. File name is in [today's date]-report-randint format.
        :return:    String containing the file path of the created Excel file.
        """
        file_name = f'reports\\{date.today()}-report-{randint(300,9000)}.xlsx'
        self.dataframe.to_excel(file_name, index=False)
        cur_directory = os.path.abspath(__file__)
        return f'{cur_directory}\\{file_name}'