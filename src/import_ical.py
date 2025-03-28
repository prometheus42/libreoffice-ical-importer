#! /usr/bin/env python3

#
# All modules are located in the 'pythonpath' sub-directory of the directory
# 'src'. The script provider adds the directory to sys.path before the main
# script is executed.
#
# Source: https://forum.openoffice.org/en/forum/viewtopic.php?f=45&t=69540
#

import sys
import logging
from pathlib import Path
from datetime import timedelta

from arrow.arrow import Arrow
from ics import Calendar
from ics.grammar.parse import ParseError

import uno
import msgbox
import unohelper
from com.sun.star.util import DateTime
from com.sun.star.util import Time
from com.sun.star.lang import Locale
from com.sun.star.uno import RuntimeException
from com.sun.star.task import XJobExecutor
from com.sun.star.ui.dialogs.TemplateDescription import FILEOPEN_SIMPLE


def fill_table(ctx, filename):
    """Fills the first worksheet with data from an iCalendar file.

    All events from a given iCalendar file are imported. Every column
    represents a attribute from the file and has a corrsponding header. The
    cells will be configured with a number format depending on the type of data
    of that column (datetime, timedelta, string, etc.)."""
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    doc = desktop.getCurrentComponent()
    sheet = doc.getCurrentController().getActiveSheet()
    # write attributes into table header
    attr = ['name', 'begin', 'end', 'duration', 'uid', 'description', 'created',
            'last_modified', 'location', 'url', 'transparent', 'alarms',
            'attendees', 'categories', 'status', 'organizer', 'classification']
    for column, name in enumerate(attr):
        cell = sheet.getCellByPosition(column, 0)
        cell.String = name
    # iterate over all events and add data to table
    for row, e in enumerate(read_ical_file(filename)):
        for column, name in enumerate(attr):
            cell = sheet.getCellByPosition(column, row+1)
            fill_cell_with_data(doc, getattr(e, name), cell)
    # mark all columns and set them to optimal width
    selection = sheet.getCellRangeByPosition(0, 1, len(attr)-1, 1)
    for c in selection.Columns:
        c.OptimalWidth = True


def read_ical_file(filename):
    """Reads an iCalendar file and returns all events from that file."""
    with open(filename, 'r') as calendar_file:
        c = Calendar(calendar_file.read())
    return c.events


def fill_cell_with_data(doc, data, cell):
    """Fills a single cell with given data.

    Before the data is stored in the cell of an worksheet, the type of the data
    is checked. All necessary conversion and type casts are done before the
    value of the cell is set."""
    number_format = doc.getNumberFormats()
    local_settings = Locale()
    local_settings.Language = 'de'
    if not data:
        cell.String = ''
    elif isinstance(data, str):
        cell.String = data
    elif isinstance(data, Arrow):
        # write datetime into cell
        cell.String = data.format('DD.MM.YYYY HH:mm:ss')
        # set number format for cell
        format_string = 'TT.MM.JJJJ HH:MM:SS'
        number_format_id = number_format.queryKey(
            format_string, local_settings, True)
        # check whether a value was found for the requested format (-1 == no value)
        if number_format_id == -1:
            number_format_id = number_format.addNew(
                format_string, local_settings)
        cell.NumberFormat = number_format_id
    elif isinstance(data, timedelta):
        # write datetime into cell
        cell.String = str(data)
        # set number format for cell
        format_string = 'HH:MM:SS'
        number_format_id = number_format.queryKey(
            format_string, local_settings, True)
        # check whether a value was found for the requested format (-1 == no value)
        if number_format_id == -1:
            number_format_id = number_format.addNew(
                format_string, local_settings)
        cell.NumberFormat = number_format_id
    elif isinstance(data, set):
        cell.String = ', '.join([x for x in data])
    else:
        print('Unsupported type: {}'.format(type(data)))
        # return data without conversion to trigger an exception
        return data


class ImportButton(unohelper.Base, XJobExecutor):
    """Handles a click on the menu item for importing an iCalendar file.

    After the click on the menu item, a file picker dialog is shown for the
    user to choose the file that should be imported. Only the file extension
    .ics is supported and only one file is imported!"""
    def __init__(self, ctx):
        self.ctx = ctx
 
    def trigger(self, arg):
        print("ImportButton was pressed.")
        smgr = self.ctx.ServiceManager
        file_dialog = smgr.createInstanceWithContext('com.sun.star.ui.dialogs.FilePicker', self.ctx)
        file_dialog.initialize((FILEOPEN_SIMPLE,))
        file_dialog.appendFilter('iCalendar-Datei (.ics)', '*.ics')
        file_dialog.setTitle('Lade iCalendar-Datei')
        file_dialog.setMultiSelectionMode(False)
        ok = file_dialog.execute()
        if ok:
            list_of_files = file_dialog.getFiles()
            file_dialog.dispose()
            try:
                # TODO: Allow multiple files to be imported into a single worksheet.
                fill_table(self.ctx, uno.fileUrlToSystemPath(list_of_files[0]))
            except UnicodeDecodeError as e:
                show_message_box(self.ctx, 'Fehler', 'Fehler beim Einlesen der Datei.')
                print(e)
            except RuntimeException as e:
                show_message_box(self.ctx, 'Fehler', 'Fehler im UNO-System.')
                print(e)
            except NotImplementedError as e:
                show_message_box(self.ctx, 'Fehler', 'Datei enthält mehrere Kalender. Diese Funktion wird noch nicht unterstützt.')
                print(e)
            except ParseError as e:
                show_message_box(self.ctx, 'Fehler', 'Kalender-Datei ist fehlerhaft.')
                print(e)
            show_message_box(self.ctx, 'Kalender importiert', 'Die Kalender-Datei wurde erfolgreich importiert.')


def show_message_box(ctx, title, message):
    """Shows a message box with a single button to acknowledge the information.

    Source: https://docs.libreoffice.org/scripting/html/msgbox_8py_source.html
    """
    message_box = msgbox.MsgBox(ctx)
    message_box.addButton('OK')
    message_box.renderFromBoxSize(150)
    message_box.numberOflines = 2
    message_box.show(message, 0, title)


g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    ImportButton,
    'com.platformedia.libreoffice.extensions.import_ical.ImportButton',
    ('com.sun.star.task.Job',),
)


if __name__ == '__main__':
    import os

    # start OpenOffice.org, listen for connections and open testing document
    os.system("/usr/bin/libreoffice --calc '--accept=socket,host=localhost,port=2002;urp;' &")

    # get local context info
    localContext = uno.getComponentContext()
    resolver = localContext.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", localContext)

    ctx = None

    # wait until the OO.o starts and connection is established
    while ctx == None:
        try:
            ctx = resolver.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        except Exception as e:
            pass

    print("Testing ImportButton...")

    # trigger our job
    testjob = ImportButton(ctx)
    testjob.trigger(())

    print("ImportButton tested.")
