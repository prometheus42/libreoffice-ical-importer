#! /usr/bin/env python3

import os
import sys
import logging
import logging.handlers
from tkinter import Tk, Frame, messagebox, StringVar, HORIZONTAL, VERTICAL
from tkinter.ttk import Button, Label, Separator
from tkinter.filedialog import askopenfilename, asksaveasfilename

from ical2csv import write_csv_file, read_ical_file


logger = logging.getLogger('ical2csv_gui')

LOG_FILENAME = 'ical2csv_gui.log'
WIDTH = 650
HEIGHT = 200
PADX = 20
PADY = 7
FILE_LABEL_FONT = ('Courier', 10)

ical_file = None
csv_file = None


def center(win):
    """
    Centers a tkinter window.

    :param win: the root or Toplevel window to center

    Source: https://stackoverflow.com/a/10018670
    """
    win.update_idletasks()
    width = win.winfo_width()
    frm_width = win.winfo_rootx() - win.winfo_x()
    win_width = width + 2 * frm_width
    height = win.winfo_height()
    titlebar_height = win.winfo_rooty() - win.winfo_y()
    win_height = height + titlebar_height + frm_width
    x = win.winfo_screenwidth() // 2 - win_width // 2
    y = win.winfo_screenheight() // 2 - win_height // 2
    win.geometry('{}x{}+{}+{}'.format(width, height, x, y))
    win.deiconify()

def create_logger():
    # create logger for this application
    global logger
    logger.setLevel(logging.DEBUG)
    log_to_file = logging.handlers.RotatingFileHandler(LOG_FILENAME, maxBytes=262144,
                                                       backupCount=5, encoding='utf-8')
    log_to_file.setLevel(logging.DEBUG)
    logger.addHandler(log_to_file)
    log_to_screen = logging.StreamHandler(sys.stdout)
    log_to_screen.setLevel(logging.INFO)
    logger.addHandler(log_to_screen)

def do_chose_ical_file():
    global ical_file, csv_file
    filename = askopenfilename(initialdir='.', title = "iCalendar-Datei auswählen...",
                            filetypes =(('iCalendar-Datei', '*.ics'),('Alle Dateien','*.*')))
    logger.info('iCalendar file was selected: {}'.format(filename))
    ical_file.set(filename)
    csv_file.set('{}.csv'.format(filename))

def do_chose_csv_file():
    global csv_file
    filename = asksaveasfilename(initialdir=os.path.dirname(csv_file.get()), title = "CSV-Datei auswählen...",
                                filetypes =(('CSV-Datei', '*.csv'),('Alle Dateien','*.*')),
                                initialfile=os.path.basename(csv_file.get()))
    logger.info('csv file was selected: {}'.format(filename))
    csv_file.set(filename)

def do_convert():
    logger.info('Starting conversion...')
    write_csv_file(read_ical_file(ical_file.get()), csv_file.get())
    logger.info('Conversion completed.')
    messagebox.showinfo('Konvertierung erfolgreich', 'Die Konvertierung der iCalendar-Datei war erfolgreich.')

def show_gui():
    global ical_file, csv_file
    
    window = Tk()
    ical_file = StringVar(window, ' '*60)
    csv_file = StringVar(window, ' '*60)
    
    window.title("iCalendar2csv")
    window.minsize(WIDTH, HEIGHT)
    window.maxsize(WIDTH, HEIGHT)
    window.geometry('{}x{}'.format(WIDTH, HEIGHT))
    window.resizable(0, 0)
    window.eval('tk::PlaceWindow %s center' % window.winfo_pathname(window.winfo_id()))
    #center(window)

    label1 = Label(window, text='iCalendar-Datei:')
    label1.grid(column=0, row=0, padx=PADX, pady=PADY, sticky="w")
    ical_file_label = Label(window, textvariable=ical_file, font=FILE_LABEL_FONT)
    ical_file_label.grid(column=1, row=0, padx=PADX, pady=PADY, sticky="w")

    file_button = Button(window,text='iCalendar-Datei auswählen...', command=do_chose_ical_file)
    file_button.grid(column=0, row=1, columnspan=2, padx=PADX, pady=PADY)

    label2 = Label(window, text='CSV-Datei:')
    label2.grid(column=0, row=3, padx=PADX, pady=PADY, sticky="w")
    csv_file_label = Label(window, textvariable=csv_file, font=FILE_LABEL_FONT)
    csv_file_label.grid(column=1, row=3, padx=PADX, pady=PADY, sticky="w")

    file_button2 = Button(window,text='CSV-Datei auswählen...', command=do_chose_csv_file)
    file_button2.grid(column=0, row=4, columnspan=2, padx=PADX, pady=PADY)

    convert_button = Button(window,text='iCalendar-Datei zu CSV-Datei konvertieren...', command=do_convert)
    convert_button.grid(column=0, row=5, columnspan=2, padx=PADX, pady=PADY)

    logger.info('Starting GUI...')
    window.mainloop()

if __name__ == '__main__':
    create_logger()
    show_gui()
