from erg import ReportGenerator

report_data = {
    "Headers": ["Name", "Age", "Department"],
    "Rows": [
        ["Alice", 30, "Engineering"],
        ["Bob", 25, "Marketing"],
        ["Charlie", 35, "HR"]
    ],
    "sort_by": "Age",      
    "style": "green"     
}

report = ReportGenerator(report_data)

# if report.needs_sort():
#     report.sort_report()

file_path = report.generate_excel()
print(f"File created at: {file_path}")
