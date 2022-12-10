
# Import iCalendar

Simple extension for LibreOffice for importing iCalendar files (.ics) into LibreOffice Calc.

Based on a [tutorial by Andreas Mantke](https://amantke.de/wp-content/uploads/2018/08/extensionsbook_20180831.pdf) and his [template](https://github.com/andreasma/extensionbook/tree/master/extensiontemplates/extension_basic_design)

Another nice example can be found [here](https://blog.mdda.net/oss/2011/10/07/python-libreoffice).

As introduction for scripting LibreOffice with Python see these sources:

* [Transfer from Basic to Python](https://wiki.openoffice.org/wiki/Python/Transfer_from_Basic_to_Python)

# Building standalone applications

The file src/ical2csv_gui.py and src/ical2csv.py contain a CLI and a GUI application to convert iCalendar files into CSV files.

To create a single EXE file for Windows user, a pyinstaller .spec file has to be created:

    pyinstaller --onefile --windowed -i images/icon.ico --add-data C:\Users\christian\AppData\Local\Programs\Python\Python38-32\Lib\site-packages\ics\grammar\contentline.ebnf;ics\grammar\ src\ical2csv_gui.py

Alternativly, you can use the provided spec file in the repo:

    pyinstaller ical2csv_gui.spec 

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

IcalImporter runs under Python 3.6 and newer. Currently (December 2019) LibreOffice seems to ship with Python 3.5.

The following Python packages are necessary:

* [click: Composable command line interface toolkit](https://palletsprojects.com/p/click/) licensed under the BSD License.
* [Ics.py: iCalendar for Humans](https://github.com/C4ptainCrunch/ics.py/), licensed under the Apache 2.0 License.
* [Arrow: Better dates & times for Python](https://github.com/crsmithdev/arrow/), licensed under the Apache 2.0 License.
* [dateutil](https://github.com/dateutil/dateutil/) as dependency of Arrow, licensed under the Apache License and BSD License.
* [tatsu](https://github.com/neogeny/tatsu) as dependency of ics, licensed under a BSD-style License.
* [six](https://github.com/benjaminp/six) as dependency of ics, licensed under the MIT License.
