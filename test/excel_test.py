from ExcelManager import ExcelManager

excel:ExcelManager = ExcelManager(file_path="example.xlsx")

screenshots = excel.capture_screenshot("Expenses",  "screenshot.png","A1:AE32")