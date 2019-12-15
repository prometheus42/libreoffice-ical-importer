#!/usr/bin/env python3
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

# Source: https://gerrit.libreoffice.org/plugins/gitiles/sdk-examples/+/HEAD/BundesGit/build

from zipfile import ZipFile
import os, sys

scriptdir = os.path.dirname(os.path.abspath(sys.argv[0]))
extensionname = open(os.path.join(scriptdir, 'extensionname.txt')).readlines()[0].rstrip('\n')
with ZipFile(extensionname, 'w') as tuesdayzip:
    os.chdir(scriptdir)
    for root, dirs, files in os.walk('.'):
        for name in files:
            if not name == extensionname:
                tuesdayzip.write(os.path.join(root, name))
