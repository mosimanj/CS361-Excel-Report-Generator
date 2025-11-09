import pandas
import openpyxl # needed for generating Excel
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from random import randint
from datetime import date
import os

STYLE_TEMPLATES = {
    "default": {
        "header_fill": PatternFill(start_color="1F4E78", fill_type="solid"),
        "header_font": Font(color="FFFFFF", bold=True),
        "header_alignment": Alignment(horizontal="center", vertical="center"),
        "row_font": Font(color="000000"),
        "alignment": Alignment(horizontal="center", vertical="center"),
        "border": Border(
            left=Side(style="thin"), right=Side(style="thin"),
            top=Side(style="thin"), bottom=Side(style="thin")
        )
    }
}

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

    #TODO: After creating style_report() and associated templates, add documentation (incl. screenshots) to readme.
    def style_report(self, file_path):
        """
        Applies predefined style template to the generated report.  Loads excel file from argument file path, applies 
        template styling to header and body rows, and saves Excel file with updates.
        :returns: original Excel file path
        """
        wb = openpyxl.load_workbook(file_path)
        ws = wb['Sheet1']

        template = STYLE_TEMPLATES[self.style]

        # Header row
        for cell in ws[1]:
            cell.fill = template["header_fill"]
            cell.font = template["header_font"]
            cell.alignment = template["header_alignment"]
            cell.border = template["border"]
        
        # Body rows
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.font = template["row_font"]
                cell.alignment = template["header_alignment"]
                cell.border = template["border"]
        
        wb.save(file_path)
        return file_path


    # def generate_excel(self):
    #     """
    #     Generates an Excel file from the DataFrame, stores it in the reports directory, and returns the created files
    #     absolute filepath. File name is in [today's date]-report-randint format.
    #     :return:    String containing the file path of the created Excel file.
    #     """
    #     file_name = f'reports\\{date.today()}-report-{randint(300,9000)}.xlsx'
    #     if self.needs_style():
    #         file_name = self.style_report(file_name)
    #     self.dataframe.to_excel(file_name, index=False)
    #     cur_directory = os.path.abspath(__file__)
    #     return f'{cur_directory}\\{file_name}'
    
def generate_excel(self):
        """
        Generates an Excel file from the DataFrame, stores it in the reports directory, and returns the created files
        absolute filepath. File name is in [today's date]-report-randint format.
        :return:    String containing the file path of the created Excel file.
        """
        # Ensure 'reports' directory exists
        if not os.path.exists("reports"):
            os.mkdir("reports")

        file_name = f"reports/{date.today()}-report-{randint(300,9000)}.xlsx"

        # Save DataFrame to excel file
        self.dataframe.to_excel(file_name, index=False)

        # Apply style
        if self.needs_style():
            file_name = self.style_report(file_name)

        return os.path.abspath(file_name)