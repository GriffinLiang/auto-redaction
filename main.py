import openpyxl
from openpyxl.styles import PatternFill, GradientFill

def find_key_word(cell, key_words):
    for key_word in key_words:
        if isinstance(cell.value, str) and key_word in cell.value:
            return True, key_word
    return False, ''

def process_sheet(sheet, key_words_dict):
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            has_key_word, key_word = find_key_word(cell, key_words)
            if has_key_word:
                key_words_dict[key_word].append('{}'.format(cell))
                cell.value = ''
                cell.fill = PatternFill("solid", fgColor="000000")

def run_redaction(file_name):
    # open an excel document
    wb = openpyxl.load_workbook(file_name)
    sheet_dict = {}
    for input_sheet in wb.worksheets:
        key_words_dict = {k:[] for k in key_words}
        process_sheet(input_sheet, key_words_dict)
        sheet_dict[input_sheet.title] = key_words_dict
    wb.save('test.xlsx')

    for key_word in key_words:
        print('========= {} ========='.format(key_word))
        for sheet_name, key_words_dict in sheet_dict.items():
            key_words_result = key_words_dict[key_word]
            if len(key_words_result) > 0:
                print(sheet_name, key_words_result)

file_name = 'DOC-000015544_test.xlsx'
run_redaction(file_name)
