import os
import os.path as osp
import sys
from datetime import datetime
from glob import glob
import openpyxl
from openpyxl.styles import PatternFill
import warnings

warnings.filterwarnings("ignore")

def load_file_to_list(filename):
    out_list = []
    with open(filename, 'rb') as f:
        for line in f.readlines():
            out_list.append(line.strip())
    return out_list

def find_key_word(cell, key_words):
    for key_word in key_words:
        # from IPython import embed; embed()
        if isinstance(cell.value, str) and key_word.decode("utf-8") in cell.value:
            return key_word
    return None

def process_sheet(sheet, key_words_dict):
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            key_word = find_key_word(cell, key_words)
            if key_word is not None:
                key_words_dict[key_word].append('{}'.format(cell.coordinate))
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
        logger.write('========= {} =========\n'.format(key_word.decode("utf-8")).encode("utf-8"))
        for sheet_name, key_words_dict in sheet_dict.items():
            key_words_result = key_words_dict[key_word]
            if len(key_words_result) > 0:
                logger.write('{}: {}\n'.format(sheet_name, ','.join(key_words_result)).encode("utf-8"))
                # logger.write('{}, {}'.format(sheet_name, key_words_result))
                # print('{}, {}'.format(sheet_name, key_words_result))

if __name__ == '__main__':
    # check input xlsx files and keyword txt
    input_files = glob('input_files/*.xlsx')
    if len(input_files) == 0:
        print('No xlsx files are found in input_files directory.')
        input("Press Enter to Quit...")
        sys.exit(0)

    if not osp.exists('input_files/keywords.txt'):
        print('Keywords.txt can not be found in input_files directory.')
        input("Press Enter to Quit...")
        sys.exit(0)
    key_words = load_file_to_list('input_files/keywords.txt')

    # check output directory
    if not osp.exists('output_files'):
        os.mkdir('output_files')
    current_time = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
    logger = open('output_files/log-{}.txt'.format(current_time), 'wb')
    print('Hi, I am processing your files one by one...')
    for input_file in input_files:
        print(input_file)
        run_redaction(input_file, key_words, logger)
    logger.close()
    print('All files have been processed.')
    input("Press Enter to Quit...")