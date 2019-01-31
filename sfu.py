from sheetfu import SpreadsheetApp

spreadsheet_key = open('spreadsheet_key','r').read()
spreadsheet = SpreadsheetApp('auth.json').open_by_id(spreadsheet_key)
sheet = spreadsheet.get_sheet_by_name('Sheet1')
data_range = sheet.get_data_range()

values = data_range.get_values()
backgrounds = data_range.get_backgrounds()

print(values)
values[2][0] = 345
data_range.set_values(values)
#data_range.set_background('#000000')
#data_range.set_font_color('#ffffff')
