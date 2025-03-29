
# Import iCalendar

Simple extension for LibreOffice for importing iCalendar files (.ics) into LibreOffice Calc.

Based on a [tutorial by Andreas Mantke](https://amantke.de/wp-content/uploads/2018/08/extensionsbook_20180831.pdf) and his [template](https://github.com/andreasma/extensionbook/tree/master/extensiontemplates/extension_basic_design)

Another nice example can be found [here](https://blog.mdda.net/oss/2011/10/07/python-libreoffice).

As introduction for scripting LibreOffice with Python see these sources:

* [Designing & Developing Python Applications](https://wiki.documentfoundation.org/Macros/Python_Design_Guide)
* [OpenOffice Developer Guide: Extensions](https://wiki.openoffice.org/wiki/Documentation/DevGuide/Extensions/Extensions)
* [Transfer from Basic to Python](https://wiki.openoffice.org/wiki/Python/Transfer_from_Basic_to_Python)
* [Interface-oriented programming in OpenOffice / LibreOffice: automate your office tasks with Python Macros](https://christopher5106.github.io/office/2015/12/06/openoffice-libreoffice-automate-your-office-tasks-with-python-macros.html)
* [Debugging Python components in LibreOffice](https://wiki.documentfoundation.org/Development/How_to_debug#Debugging_Python_components_in_LibreOffice)

# Building LibreOffice extension

Before the extension can be packaged, all translations have to be compiled as
.mo files to be used by gettext:

    make compile-translations

To build a LibreOffice extension use the Makefile that is provided with the project:

    make zip

To install the extension in LibreOffice execute the following command:

    make install

# Building standalone applications

The file src/ical2csv_gui.py and src/ical2csv.py contain a CLI and a GUI application to convert iCalendar files into CSV files.

To create a single EXE file for Windows user, a pyinstaller .spec file has to be created:

    pyinstaller --onefile --windowed -i images/icon.ico --add-data C:\Users\christian\AppData\Local\Programs\Python\Python38-32\Lib\site-packages\ics\grammar\contentline.ebnf;ics\grammar\ src\ical2csv_gui.py

Alternativly, you can use the provided spec file in the repo:

    pyinstaller ical2csv_gui.spec 

# Usage in LibreOffice

1. Download LibreOffice extension: https://github.com/prometheus42/libreoffice-ical-importer/releases/tag/v0.2
2. Install extension in LibreOffice
3. Open LibreOffice Calc
4. Click on the tools ("Werkzeuge") menu and start the import function
5. Choose iCal file

The fields of the iCal file should be stored as columns in the table.

# Debugging

By default, there is no default stdout on Windows. All output from the
extension will just vanish. To get a console to be used as stdout, you can
start LibreOffice using soffice.com instead of soffice.exe as described in the
[documentation](https://wiki.documentfoundation.org/Macros/Python_Design_Guide#Output_to_Consoles):

    start "C:\Program Files\LibreOffice\program\soffice.com"

All debug messages will also be written to log file:

- /home/\<username\>/.config/libreoffice/4/user/ (Linux)
- C:\Users\\<username\>\AppData\Roaming\LibreOffice\4\user\ (Windows)

# File list

* META-INF/manifest.xml -> manifest declaring all parts of the extension
* description.xml -> XML file with all information about the extension
* gui.xcu -> XML file for all GUI elements of the extension
* src/import_ical.py -> python code to read iCalendar file and write data into worksheet
* registration/license_*.txt -> license files in various languages
* description/description_*.txt -> info text for extension in various languages
* images/icon.png -> icon for extension
* extensionname.txt -> contains the name of the extension for build script
* build.py -> python script to build .oxt file

# Contributions

* Icon for IcalImporter from GNOME Desktop Tango icon set. Licensed under GNU General Public License version 2.

# License

This extension is released under the MIT License.

# Requirements

IcalImporter runs under Python 3.6 and newer. Currently (March 2025) LibreOffice seems to ship with Python 3.10.

The following Python packages are necessary:

* [click: Composable command line interface toolkit](https://palletsprojects.com/p/click/) licensed under the BSD License.
* [Ics.py: iCalendar for Humans](https://github.com/C4ptainCrunch/ics.py/), licensed under the Apache 2.0 License.
* [Arrow: Better dates & times for Python](https://github.com/crsmithdev/arrow/), licensed under the Apache 2.0 License.
* [dateutil](https://github.com/dateutil/dateutil/) as dependency of Arrow, licensed under the Apache License and BSD License.
* [tatsu](https://github.com/neogeny/tatsu) as dependency of ics, licensed under a BSD-style License.
* [six](https://github.com/benjaminp/six) as dependency of ics, licensed under the MIT License.
