# Randomopgaver
# -*- coding: utf-8 -*-
import sys
import os
import re
import time
import csv
import comtypes.client
from comtypes.gen import Word
from comtypes.gen import Excel


def create_randomopgaver(
        display_line, error_handler, doc_file, table_file, output_folder):

    global doc, word, row, headers, pattern, name, clean_up, csvfile

    # def clean_state():
    #     global word, excel, csvfile
    #     try:
    #         csvfile.close()
    #     except:
    #         pass
    #     try:
    #         word.Quit(SaveChanges=False)
    #     except:
    #         pass
    #     try:
    #         excel.Quit(SaveChanges=False)
    #     except:
    #         pass
    #     try:
    #         remove_file(os.path.join(
    #             output_folder,
    #             'temp_csv_file_32KJ4H23J4KH23JK4H.txt')
    #         )
    #     except:
    #         pass
    #     try:
    #         remove_file(os.path.join(
    #             output_folder,
    #             'temp_csv_file_32KJ4H23J4KH23JK4H.txt_replace')
    #         )
    #     except:
    #         pass


    def remove_file(file):
        if os.path.isfile(file):
            os.remove(file)

    def mkdir(path):
        if not os.path.exists(path):
            os.mkdir(path)

    def replace(show_answers):
        global doc, word
        # Open scaffold document
        try:
            doc = word.Documents.Open(doc_file, False, True)
        except:
            word.Quit(SaveChanges=False)
            csvfile.close()
            time.sleep(0.3)
            remove_file(table_file)
            error_handler('Fejl under åbning af Word-dockment: ' + sys.exc_info()[0])
            raise Exception

        # Replace tags with text
        for i, text in enumerate(row):
            # First check if header begins with $
            if headers[i] and headers[i][0] == '$':
                if not show_answers:
                    if pattern.match(headers[i]):
                        text = ''
                find_str = headers[i]
                replace_str = text
                word.Selection.Find.Execute(
                    find_str, False, False, False, False, False, True,
                    Word.wdFindContinue, False, replace_str, Word.wdReplaceAll
                )

    def save(subfolder):
        global doc, word
        folder = os.path.join(output_folder, subfolder)
        mkdir(folder)
        path = os.path.join(folder, name + '.pdf')
        try:
            doc.SaveAs(path, FileFormat=Word.wdFormatPDF)
        except:
            word.Quit(SaveChanges=False)
            csvfile.close()
            time.sleep(0.3)
            remove_file(table_file)
            error_handler('Fejl: Kunne ikke gemme til filen:'
                          '\n' + path +
                          '\nSørg venligst for at ingen output-filer er '
                          'åbne i andre programmer.')
            raise Exception
        else:
            doc.Close(SaveChanges=False)
            display_line('Gemte: ' + name + '.pdf til mappen ' + subfolder)

    # Whether cleanup of a temporary file is neccessary
    clean_up = False
    # Whether there are answers in the table
    answers = False

    # If table_file is an Excel file, then convert it to a .csv file
    if table_file[-5:] == ".xlsx" or table_file[-4:] == ".xls":
        display_line('Excel-fil givet som input. Genererer .csv-fil...')
        # Open excel file invisibly
        try:
            excel = comtypes.client.CreateObject('Excel.Application')
        except:
            csvfile.close()
            time.sleep(0.3)
            remove_file(table_file)
            error_handler('Uventet fejl: Kunne ikke åbne Excel.')
            return
        excel.Visible = False
        excel.DisplayAlerts = False
        try:
            wb = excel.Workbooks.Open(table_file)
        except:
            excel.Quit(SaveChanges=False)
            remove_file(table_file)
            error_handler('Uventet fejl under åbning af Excel-ark:' + sys.exc_info()[0])
            return

        # Save .csv file
        table_file = os.path.join(output_folder,
                                  'temp_csv_file_32KJ4H23J4KH23JK4H.txt')
        if os.path.isfile(table_file):
            os.remove(table_file)
            time.sleep(0.3)
        try:
            wb.SaveAs(table_file, FileFormat=Excel.xlUnicodeText)
        except:
            excel.Quit(SaveChanges=False)
            error_handler("Uventet fejl da Excel-ark skulle gemmes: " + sys.exc_info()[0])
            return
        else:
            excel.Quit(SaveChanges=False)

        # Replace dash with en dash in .csv file
        with open(table_file, 'r') as f1:
            with open(table_file + '_replace', 'w') as f2:
                for line in f1:
                    line = line.replace('-', '–').replace('\r', '')
                    f2.write(line)
        time.sleep(0.3)
        remove_file(table_file)
        time.sleep(0.3)
        os.rename(table_file + '_replace', table_file)
        clean_up = True

    # Open .csv file
    with open(table_file, newline='', encoding='utf-16-le') as csvfile:
        # Load .csv file into special object (reader)
        reader = csv.reader(csvfile, delimiter='\t', quotechar='"')

        # Get headers and test for answer tags
        headers = reader.__next__()
        headers[0] = headers[0].replace(u'\ufeff', '')     # remove BOM
        pattern = re.compile("\$\$.*[^!$]+\$\$$")
        for header in headers:
            if pattern.match(header):
                answers = True
                break

        # Open Word invisibly
        try:
            word = comtypes.client.CreateObject('Word.Application')
        except:
            csvfile.close()
            time.sleep(0.3)
            remove_file(table_file)
            error_handler('Uventet fejl: Kunne ikke åbne Word.')
            return
        word.Visible = False
        word.DisplayAlerts = False

        # Process the _rest_ of the CSV file
        display_line('Programmet gemmer .pdf-filer i mappen:'
                     '\n\t' + output_folder)
        for row in reader:
            # Get name from list
            name = row[0]

            if name:
                # Replace tags with text and save assignments
                try:
                    replace(show_answers=False)
                    save(subfolder='Opgaver')

                    # Save answers if applicable
                    if answers:
                        replace(show_answers=True)
                        save(subfolder='Svar')
                except:
                    return False

        # Quit Word
        word.Quit(SaveChanges=False)

    # Delete temporary .csv file if created
    if clean_up:
        remove_file(table_file)

    display_line('Randomopgaver blev genereret med succes.')
    return True
