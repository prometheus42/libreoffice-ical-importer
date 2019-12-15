all:	clean build install

clean:
	unopkg remove de.ichmann.libreoffice.import_ical
	rm import_ical.oxt
	
zip:
	zip -r import_ical.oxt \
		description.xml \
		META-INF/manifest.xml \
		gui.xcu \
		src/import_ical.py \
		images/icon.png \
		description/description_de.txt \
		description/description_en.txt \
		registration/license_de.txt \
		registration/license_en.txt 

build:
	/usr/bin/env python3 ./build.py

install:
	unopkg add import_ical.oxt
