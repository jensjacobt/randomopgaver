# Randomopgaver
# -*- coding: utf-8 -*-
import sys
import os
from filegenerator import create_randomopgaver

global doc, word, doc_file, row, headers, pattern, name, output_folder


def mkdir(path):
    if not os.path.exists(path):
        os.mkdir(path)


def error_handler(error_string):
    print(error_string)
    sys.exit(1)


def main():
    # Global variables
    global doc, word, doc_file, row, headers, pattern, name, output_folder

    # Handle arguments and check existence
    if len(sys.argv) == 3 or len(sys.argv) == 4:
        doc_file = os.path.abspath(sys.argv[1])
        table_file = os.path.abspath(sys.argv[2])
        if not os.path.isfile(doc_file):
            error_handler('Fejl: Den angivne sti for Word-dokumentet er ikke korrekt.')
        if not os.path.isfile(table_file):
            error_handler('Fejl: Den angivne sti for Excel-arket/CSV-filen er ikke korrekt.')
    else:
        print('Genererer randomopgaver i PDF-format fra et Word-dokument and en tabelfil.'
              '\n'
              '\nrandomopgaver dokument tabelfil [output-mappe]'
              '\n'
              '\ndokument        Stien til det Word-dokumentet hvori nogle tags skal udskiftes.'
              '\ntabelfil        Stien til en tabelfil med tags og værdier til udskiftning.'
              '\n                Filen kan enten være et Excel-ark eller en TXT-fil.'
              '\n                I tilfælde af en TXT-fil anvendes separator (;), '
              '\n                citationstegn (") og tegnkodning UTF-16LE.'
              '\noutput-mappe    Mappen hvor de genererede PDF-filer gemmes. Hvis ingen mappe'
              '\n                angives, så vil PDF-filerne gemmes i undermapperne Opgaver og'
              '\n                Svar til mappen, hvori Word-dokumentet ligger. Output-mappen'
              '\n                oprettes, hvis den ikke findes i forvejen.'
              )
    #                                                                                            |
        return
    if len(sys.argv) == 4:
        output_folder = sys.argv[3]
    else:
        output_folder = os.path.abspath(os.path.dirname(doc_file))
    mkdir(output_folder)

    create_randomopgaver(print, error_handler, doc_file, table_file, output_folder)


main()
