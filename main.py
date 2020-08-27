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
        if isinstance(cell.value, str) and key_word.decode("utf-8") in cell.value:
            return key_word
    return None

def if_contain_chaos(keyword):
    try:
        keyword.encode("gb2312")
    except UnicodeEncodeError:
        return True
    return False


def process_sheet(sheet, key_words_dict, logger, mode):
    for row in sheet.iter_rows():
        for cell in row:
            try:
                if cell.value is None:
                    continue
                key_word = find_key_word(cell, key_words)
                if key_word is not None:
                    key_words_dict[key_word].append('{}'.format(cell.coordinate))
                    if mode == 'b':
                        cell.value = ''
                        cell.fill = PatternFill("solid", fgColor="000000")
                    else:
                        # replace all the keywords
                        for kw in key_words:
                            cell.value = cell.value.replace(kw.decode("utf-8"), '[DP Redaction]')
            except:
                logger.write('ERROR: cell {} has wrong value.\n'.format(cell).encode("utf-8"))

def run_redaction(file_name, key_words, logger, mode):
    # open an excel document
    wb = openpyxl.load_workbook(file_name)
    sheet_dict = {}
    for input_sheet in wb.worksheets:
        print(input_sheet.title)
        key_words_dict = {k:[] for k in key_words}
        process_sheet(input_sheet, key_words_dict, logger, mode)
        sheet_dict[input_sheet.title] = key_words_dict
    file_name, suffix = osp.splitext(osp.basename(file_name))
    wb.save(osp.join('output_files', '{}_redacted{}'.format(file_name, suffix)))

    for key_word in key_words:
        logger.write('* {} \n'.format(key_word.decode("utf-8")).encode("utf-8"))
        for sheet_name, key_words_dict in sheet_dict.items():
            key_words_result = key_words_dict[key_word]
            if len(key_words_result) > 0:
                logger.write('{}: {}\n'.format(sheet_name, ','.join(key_words_result)).encode("utf-8"))
                # logger.write('{}, {}'.format(sheet_name, key_words_result))
                # print('{}, {}'.format(sheet_name, key_words_result))

if __name__ == '__main__':

    print('''
                  _                           _            _   _             
       __ _ _   _| |_ ___        _ __ ___  __| | __ _  ___| |_(_) ___  _ __  
      / _` | | | | __/ _ \ _____| '__/ _ \/ _` |/ _` |/ __| __| |/ _ \| '_ \ 
     | (_| | |_| | || (_) |_____| | |  __/ (_| | (_| | (__| |_| | (_) | | | |
      \__,_|\__,_|\__\___/      |_|  \___|\__,_|\__,_|\___|\__|_|\___/|_| |_|
                                                                             
    ''')

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
    print('Some chaos keywords are listed below:')
    is_chaos = False
    for key_word_index, key_word in enumerate(key_words):
        if if_contain_chaos(key_word.decode("utf-8")):
            is_chaos = True
            print('Keyword {}:{}'.format(key_word_index, key_word.decode("utf-8")))
    if is_chaos:
        input("Press correct the above keywords.")
        key_word_correct = 'n'
    else:
        input("No chaos keywords are found. Please press enter to continue.")
        key_word_correct = 'y'

    if key_word_correct is 'y':
        mode = input("Use [r]eplace or [b]lack?\n")
        print('You choose {}'.format(mode))
        while mode not in ['r', 'b']:
            mode = input('Please input r or b. r -> repalce, b -> black.')
        # check output directory
        if not osp.exists('output_files'):
            os.mkdir('output_files')
        current_time = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
        logger = open('output_files/log-{}.txt'.format(current_time), 'wb')

        # start process files
        print('Hi, I am processing your files one by one...')
        for input_file in input_files:
            print(input_file)
            logger.write('=========== {} ===========\n'.format(input_file).encode("utf-8"))
            run_redaction(input_file, key_words, logger, mode)
            logger.write('\n'.encode("utf-8"))
            logger.flush()
        logger.close()
        print('All files have been processed.')
    else:
        print('Please modify keywords.txt.')

    input("Press Enter to Quit...")
    # from IPython import embed; embed()