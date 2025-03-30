all:	clean compile-translations zip install

clean:
	unopkg remove de.ichmann.libreoffice.import_ical
	rm import_ical.oxt
	rm localizations/en/LC_MESSAGES/import_ical.mo
	rm localizations/de/LC_MESSAGES/import_ical.mo

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

install:
	unopkg add import_ical.oxt

update-translations:
	xgettext src/import_ical.py -o localizations/en/LC_MESSAGES/messages.pot
	xgettext src/import_ical.py -o localizations/de/LC_MESSAGES/messages.pot
	msgmerge -U --backup=none localizations/en/LC_MESSAGES/import_ical.po localizations/en/LC_MESSAGES/messages.pot
	msgmerge -U --backup=none localizations/de/LC_MESSAGES/import_ical.po localizations/de/LC_MESSAGES/messages.pot
	rm localizations/en/LC_MESSAGES/messages.pot
	rm localizations/de/LC_MESSAGES/messages.pot

compile-translations:
	msgfmt -o localizations/de/LC_MESSAGES/import_ical.mo localizations/de/LC_MESSAGES/import_ical.po
	msgfmt -o localizations/en/LC_MESSAGES/import_ical.mo localizations/en/LC_MESSAGES/import_ical.po
