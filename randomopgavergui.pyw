# The GUI of Randomopgaver
# -*- coding: utf-8 -*-
import os
import tkinter
from tkinter import *
from tkinter import scrolledtext
from tkinter import messagebox
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory
from filegenerator import create_randomopgaver

global doc_file, tab_file, out_dir, st, status


# GUI FUNCTIONS


def set_doc():
    doc_file = askopenfilename(
        initialdir='',
        filetypes=(('Alle Word-dokumenter', '*.docx;*.doc'),
                   ('Alle filer', '*.*')),
        title='Vælg et Word-dokument',
        multiple=False
    )
    e1.delete(0, END)
    e1.insert(0, os.path.abspath(doc_file))


def set_tab():
    tab_file = askopenfilename(
        initialdir='',
        filetypes=(('Alle Excel-ark', '*.xlsx;*.xls'),
                   ('Alle CSV-filer', '*.csv'),
                   ('Alle TXT-filer', '*.txt'),
                   ("Alle filer", "*.*")),
        title="Vælg et Excel-ark (eller en CSV-fil)",
        multiple=False
    )
    e2.delete(0, END)
    e2.insert(0, os.path.abspath(tab_file))


def set_out():
    out_dir = askdirectory(
        initialdir="",
        title="Vælg en mappe til output"
    )
    e3.delete(0, END)
    e3.insert(0, os.path.abspath(out_dir))


def generate():
    global st, status
    if not os.path.isfile(e1.get()):
        error_handler(
            'Fejl: Den angivne sti for Word-dokumentet er ikke korrekt.')
        return
    if not os.path.isfile(e2.get()):
        error_handler(
            'Fejl: Den angivne sti for Excel-arket/CSV-filen er ikke korrekt.')
        return
    out_folder = (e3.get() if e3.get() else os.path.abspath(os.path.dirname(e1.get())))
    st = tkinter.scrolledtext.ScrolledText(
        master=root,
        wrap=tkinter.WORD,
        height=15,
        width=5
    )
    st.grid(row=4, column=0, columnspan=3, padx=xpad, pady=ypad, sticky=EW)
    root.grid_columnconfigure(0, weight=1)
    st.delete(1.0, END)
    st.insert(1.0, 'Genererering af randomopgaver påbegyndt.')
    status['text'] = 'Genererer randomopgaver...'
    st.update()
    res = create_randomopgaver(write, error_handler, e1.get(), e2.get(), out_folder)
    if res:
        status['text'] = 'Randomopgaver blev genereret med succes.'


def write(string):
    global st
    st.insert(END, '\n' + string)
    st.see(END)
    st.update()


def error_handler(error_string):
    global status
    status['text'] = 'Der opstod en fejl under genereringen, som nu er stoppet.'
    if tkinter.messagebox.askretrycancel(
            'Der opstod en fejl',
            error_string,
            icon=messagebox.ERROR,
            default=messagebox.CANCEL
    ):
        generate()


# GUI

root = tkinter.Tk()
root.wm_title('Randomopgaver')
root.resizable(width=False, height=False)
# https://commons.wikimedia.org/wiki/File:One_die.jpeg
root.iconbitmap(os.path.join(os.path.dirname(os.path.realpath(__file__)), 'icon.ico'))

ypad = 5
xpad = 6

l1 = Label(root, text='Word-dokument:')
l2 = Label(root, text='Excel-ark:')
l3 = Label(root, text='Evt. output-mappe:')
status = Label(root, text='')

l1.grid(row=0, sticky=W, padx=xpad)
l2.grid(row=1, sticky=W, padx=xpad)
l3.grid(row=2, sticky=W, padx=xpad)
status.grid(row=3, columnspan=2, sticky=W, padx=xpad)

e1 = Entry(root, width=40)
e2 = Entry(root, width=40)
e3 = Entry(root, width=40)

e1.grid(row=0, column=1, pady=ypad)
e2.grid(row=1, column=1, pady=ypad)
e3.grid(row=2, column=1, pady=ypad)

b1 = Button(root, text='Gennemse...', command=set_doc, width=14)
b2 = Button(root, text='Gennemse...', command=set_tab, width=14)
b3 = Button(root, text='Gennemse...', command=set_out, width=14)
b4 = Button(root, text='Generer opgaver', command=generate, width=14)

b1.grid(row=0, column=2, sticky=E, padx=xpad, pady=ypad)
b2.grid(row=1, column=2, sticky=E, padx=xpad, pady=ypad)
b3.grid(row=2, column=2, sticky=E, padx=xpad, pady=ypad)
b4.grid(row=3, column=2, sticky=E, padx=xpad, pady=ypad)

root.columnconfigure(1, weight=5)
root.rowconfigure(1, weight=5)

root.mainloop()
