
# Import iCalendar

Simple extension for LibreOffice for importing iCalendar files (.ics) into LibreOffice Calc.

Based on a [tutorial by Andreas Mantke](https://amantke.de/wp-content/uploads/2018/08/extensionsbook_20180831.pdf) and his [template](https://github.com/andreasma/extensionbook/tree/master/extensiontemplates/extension_basic_design)

Another nice example can be found [here](https://blog.mdda.net/oss/2011/10/07/python-libreoffice).

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

IcalImporter should run under Python 3.5 and newer. The following Python packages are necessary:

* [Ics.py: iCalendar for Humans](https://github.com/C4ptainCrunch/ics.py/)
* [Arrow: Better dates & times for Python](https://github.com/crsmithdev/arrow/)
