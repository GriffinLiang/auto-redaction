import os.path as osp
from datetime import datetime
from glob import glob
import openpyxl
from openpyxl.styles import PatternFill, GradientFill

def load_file_to_list(filename):
    out_list = []
    with open(filename) as f:
        for line in f.readlines():
            out_list.append(line.strip())
    return out_list

def find_key_word(cell, key_words):
    for key_word in key_words:
        if isinstance(cell.value, str) and key_word in cell.value:
            return key_word
    return None

def process_sheet(sheet, key_words_dict):
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            key_word = find_key_word(cell, key_words)
            if key_word is not None:
                key_words_dict[key_word].append('{}'.format(cell))
                cell.value = ''
                cell.fill = PatternFill("solid", fgColor="000000")

def run_redaction(file_name, key_words, logger):
    # open an excel document
    wb = openpyxl.load_workbook(file_name)
    sheet_dict = {}
    for input_sheet in wb.worksheets:
        key_words_dict = {k:[] for k in key_words}
        process_sheet(input_sheet, key_words_dict)
        sheet_dict[input_sheet.title] = key_words_dict
    wb.save(osp.join('output_files', osp.basename(file_name)))

    for key_word in key_words:
        logger.write('========= {} =========\n'.format(key_word))
        for sheet_name, key_words_dict in sheet_dict.items():
            key_words_result = key_words_dict[key_word]
            if len(key_words_result) > 0:
                logger.write('{}, {}'.format(sheet_name, key_words_result))

if __name__ == '__main__':
    current_time = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
    logger = open('output_files/log-{}.txt'.format(current_time), 'w')
    input_files = glob('input_files/*.xlsx')
    key_words = load_file_to_list('input_files/keywords.txt')
    for input_file in input_files:
        print(input_file)
        run_redaction(input_file, key_words, logger)
        break
    logger.close()
