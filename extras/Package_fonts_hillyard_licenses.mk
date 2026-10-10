# -*- Mode: makefile-gmake; tab-width: 4; indent-tabs-mode: t -*-
#
# This file is part of the LibreOffice project.
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Licences and manifest for the Hillyard fonts. Kept out of
# extras_fonts_hillyard: every file in the ooo_fonts install module is flagged
# as a font, and the Windows installer fails trying to register a .txt file
# (Error 1907).

$(eval $(call gb_Package_Package,extras_fonts_hillyard_licenses,$(SRCDIR)/extras/source/truetype/hillyard))

$(eval $(call gb_Package_add_files,extras_fonts_hillyard_licenses,$(LIBO_SHARE_FOLDER)/fonts/truetype/hillyard-licenses,\
	LICENSE-gelasio.txt \
	LICENSE-selawik.txt \
	LICENSE-texgyre.txt \
	LICENSE-urw.txt \
	LICENSE-urw-COPYING.txt \
	MANIFEST-hillyard.txt \
))

# vim: set noet sw=4 ts=4:
