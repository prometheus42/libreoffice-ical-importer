all:	clean compile-translations zip install

clean:
	unopkg remove de.ichmann.libreoffice.import_ical
	rm import_ical.oxt

zip:
	zip -r import_ical.oxt \
		description.xml \
		README.md \
		META-INF/manifest.xml \
		gui.xcu \
		src/import_ical.py \
		images/icon.png \
		images/menu_icon.png \
		description/description_de.txt \
		description/description_en.txt \
		registration/license_de.txt \
		registration/license_en.txt \
		localizations/* \
		src/pythonpath/*

build:
	/usr/bin/env python3 ./build.py

install:
	unopkg add import_ical.oxt

update-translations:
	xgettext src/import_ical.py -p localizations/en/LC_MESSAGES/
	xgettext src/import_ical.py -p localizations/de/LC_MESSAGES/

compile-translations:
	msgfmt -o localizations/de/LC_MESSAGES/import_ical.mo localizations/de/LC_MESSAGES/import_ical.po
	msgfmt -o localizations/en/LC_MESSAGES/import_ical.mo localizations/en/LC_MESSAGES/import_ical.po
