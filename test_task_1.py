#!/usr/bin/env python
# encoding:     UTF-8
# filename:     test_task_1.py
# project:      TEST_TASK_1
# ver/date:	    1.0.2 / 2020.10.27
# company:
# author:       Starichenko Andrey
# language:     Python 3.8.6rc1 (AMD64)
import sys; print(sys.version)
# OS:           Windows 8.1 SL RU (AMD64)
import platform; print(platform.platform(), platform.machine())
# IDE:          PyCharm Community 2020.2
# licence: 	    GPLv3

# todo:1: check ID for unique! in all column at start!!!
# todo:10: close all queue threads on finish

# TODO:0: implement AUTOTEST!
# TODO:0: use PATHLIB for names

"""ЗАДАЧА
Имеется файл "Задача.xlsx" со следующими данными:
Код_КА	Документы для формирования	Дата документа	Отправить
PA_0001	акт,счет	30.06.2020	1
PA_0002	акт,счет	30.06.2020	1
PA_0003	акт	25.06.2020	0
PA_0004	акт,счет	23.06.2020	1
PA_0005	акт,счет	27.06.2020	1
PA_0006	акт	17.06.2020	0
PA_0007	счет	01.07.2020	0
PA_0008	акт	24.06.2020	0
PA_0009	счет	15.06.2020	1
PA_00010	счет	19.06.2020	1
PA_00011	акт,счет	28.06.2020	1
...

1) Создайте программный код для формирования пустых файлов в формате *.txt
с названием по шаблону "КА_<Код КА>_<тип_документа>_<дата>" по списку файла "Задание.xlsx"
    <Код КА> - значение из столбца А
    <тип_документа> - значение из столбца B. Значения могут быть указаны через запятую, по каждому из них нужно сформировать отдельный файл
    <дата> - значение из столбца C
2) Распределите сформированные файлы на подпапки:
    Папка1) Файлы, по которым в мастер-файле в столбце D установлено значение 1
    Папка2) Файлы, у которых в названии файла 2 последние цифры <Кода_КА> одинаковые, например, PA_00011, PA_00022
    Папка3) Файлы, с датой в промежутке от 20.06.2020 до 10.07.2020
Если файл подходит под несколько условий, то необходимо делать его копию
"""


import os
import time
import glob
import re
from openpyxl import load_workbook, Workbook    # pip install openpyxl
import timeit
import winsound
from threading import Thread
import threading

SOURCE_FILE_NAME = "Задание.xlsx"

# PATH FORMAT: without starting and finishing slashes!
PATH1 = "Папка1"
PATH2 = "Папка2"
PATH3 = "Папка3"

# DATE FORMAT: "dd.mm.yyyy"!
DATE_START = "20.06.2020"
DATE_FINISH = "10.07.2020"


date_start_t = time.strptime(DATE_START, "%d.%m.%Y")
date_finish_t = time.strptime(DATE_FINISH, "%d.%m.%Y")


ERROR_LIST = [
    ["ID", "MSG_PRINT"],
    [1, "[ERROR/filenames]CAN\'T SAVE FILE, check SYMBOLS IN TEXT!!!"],
    [2, "[ERROR/dateOrOther] may be incorrect date (or other unknown error)"],
]


def main():
    make_directories()

    wb_defects = Workbook()
    ws_defects = wb_defects.active

    # todo:1: check ID for unique! in all column at start!!!
    for source_line in read_source_lines(SOURCE_FILE_NAME):
        defected_line = False

        try:
            for filename in make_filename_group(source_line):
                for pathname in filter_and_make_path_group(source_line):
                    try:
                        make_file(pathname, filename)
                    except:
                        error_i = ERROR_LIST[1]
                        print(error_i[1], pathname, filename)
                        defected_line = True

                if defected_line:
                    break #stop listing "for filename"

        except:
            error_i = ERROR_LIST[2]
            print(error_i[1], source_line)
            defected_line = True

        if defected_line:
            delete_files_for_defected_row(source_line)
            ws_defects.append(list(source_line)+[error_i[1]])

            alert_error()

    if list(ws_defects.values) != []:
        defected_xls_filename_nameonly = "DEFECTED_DATA=check manually.xlsx"
        defected_xls_filename = PATH_RESULT + defected_xls_filename_nameonly

        max_defected_rows = ws_defects.max_row
        print("*"*80)
        print(f"количество дефектных строк=[{max_defected_rows}]")
        wb_defects.save(defected_xls_filename)
        wb_defects.close()

        make_copy_this_script_to_result_folder(defected_xls_filename_nameonly)


def make_directories():
    """make directories for saving results"""
    global PATH1, PATH2, PATH3, PATH_RESULT
    PATH_RESULT = "РЕЗУЛЬТАТЫ_выгрузки_" + time.strftime("%Y.%m.%d_%H.%M.%S", time.localtime()) + "/"
    PATH1 = PATH_RESULT + PATH1
    PATH2 = PATH_RESULT + PATH2
    PATH3 = PATH_RESULT + PATH3

    os.mkdir(PATH_RESULT)
    os.mkdir(PATH1)
    os.mkdir(PATH2)
    os.mkdir(PATH3)


def _filter1(row_tuple):  # D-column
    return str(row_tuple[3]) == "1"


def _filter2(row_tuple):  # last 2 digits in CA
    return row_tuple[0][-2] == row_tuple[0][-1]


def _filter3(row_tuple):  # date
    date_doc_t = time.strptime(row_tuple[2], "%d.%m.%Y")
    return date_start_t < date_doc_t < date_finish_t


def read_source_lines(filename):
    """
    Process: from Excel file read lines.
    Input: Excel filename(str)  ="Задание.xlsx".
    Output: rows (iter)         =<generator object Worksheet._cells_by_row at 0x000000D23AF32270>
    """
    wb = load_workbook(filename)
    ws = wb.active
    iter = ws.iter_rows(values_only=True)
    return iter


def filter_and_make_path_group(source_line):
    """
    1. Filter actual source line data from Excel
    2. Make corresponding list of path where will actually NEED saving files.
    Input: actual Excel row (tuple)         =('PA_0001', 'акт,счет', '30.06.2020', 1)
    Output: path list to save files (list)  =['РЕЗУЛЬТАТЫ_выгрузки_2020.10.06_17.58.48/Папка1', 'РЕЗУЛЬТАТЫ_выгрузки_2020.10.06_17.58.48/Папка3']
    """
    output = []

    if _filter1(source_line): output.append(PATH1)
    if _filter2(source_line): output.append(PATH2)
    if _filter3(source_line): output.append(PATH3)
    return output


def make_filename_group(row_tuple):
    """
    in relation to data in 2nd column make filenames need to save
    Input: actual Excel row (tuple)     =('PA_0001', 'акт,счет', '30.06.2020', 1)
    Output: filename list (list)        =['КА_PA_0001_акт_30.06.2020.txt', 'КА_PA_0001_счет_30.06.2020.txt']
    """
    output = []

    if row_tuple[0] in ['Код_КА', ""]:
        pass
    else:
        source_list_copy = list(row_tuple)
        for i in row_tuple[1].split(","):
            source_list_copy[1] = i
            output.append(_make_filename(source_list_copy))
    return output


def _make_filename(cells_data):
    """
    make filename only from Excel cells by template
    Input: isolated data for one file (list)    =['PA_0001', 'акт', '30.06.2020', 1]
    Output: filename (str)                      ="КА_PA_0001_акт_30.06.2020.txt"
    """
    if cells_data == []: return None

    part1_ca = cells_data[0]
    part2_type = cells_data[1]
    part3_date = cells_data[2]
    # make filename ("КА_<Код КА>_<тип_документа>_<дата>.txt")
    file_name_i = f"КА_{part1_ca}_{part2_type}_{part3_date}.txt"
    return file_name_i


def make_file(pathname, filename):
    """
    actually make file
    Input: path (str) and filename (str)    ="РЕЗУЛЬТАТЫ_выгрузки_2020.10.06_18.16.48/Папка1", "КА_PA_0001_акт_30.06.2020.txt"
    Output: save file (None)                =
    """
    if pathname in [None, ""] or filename in [None, ""]:
        pass
    else:
        full_filename = f"{pathname}/{filename}"
        newfile = open(full_filename, "w")
        newfile.close()
    return


def delete_files_for_defected_row(source_line):
    """
    if get error trying save any file, there are possibly already saved previous file in group.
    delete all group!
    Input: actual Excel row (tuple)     =('PA_0001', 'акт,счет', '30.06.2020', 1)
    Output: delete files (None)         =
    """
    delete_files_list = glob.glob(f"{PATH_RESULT}*/*{source_line[0]}_*.txt")
    for i in delete_files_list:
        os.remove(i)


def make_copy_this_script_to_result_folder(defected_xls_filename_nameonly):
    """
    make copy this script to result folder near the defected data excel file
    and chance link there to process it in next time.
    so you can manually edit defected excel file and start process again in the same folder.
    Input: filename (str)               ="DEFECTED_DATA=check manually.xlsx"
    Output: make edited copy (None)     =
    """
    this_script_copy_dstname = PATH_RESULT + os.path.basename(__file__)

    fp_this_script = open(__file__, "r", encoding="utf-8")
    fp_this_script_copy = open(this_script_copy_dstname, "w", encoding="utf-8")

    for line in fp_this_script:
        pattern = "SOURCE_FILE_NAME\s*=\s*.*[.]xls.*"
        repl = f'SOURCE_FILE_NAME = "{defected_xls_filename_nameonly}"\n'
        new_line = re.sub(pattern, repl, line)
        fp_this_script_copy.write(new_line)

    fp_this_script.close()
    fp_this_script_copy.close()


def alert_error():
    try:
        def play_my_sound(sound_winreg_alias_name):
            winsound.PlaySound(sound=sound_winreg_alias_name, flags=winsound.SND_ALIAS)

        sound_winreg_alias_name = 'SystemAsterisk'
        th = Thread(target=play_my_sound, args=(sound_winreg_alias_name,))
        th.start()
    except:
        pass


if __name__ == '__main__':
    print("*" * 80)
    time_of_process = timeit.timeit(stmt="main()", number=1, globals=globals())
    print("*"*80)
    print(f"Время выполнения: [{time_of_process}]секунд")
    print("!!!основной ПРОЦЕСС ЗАВЕРШЕН!!!")

    while threading.active_count() > 5:   # todo:10: close all queue threads on finish
        time.sleep(1)
        print(threading.active_count())
