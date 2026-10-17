# -*- Mode: makefile-gmake; tab-width: 4; indent-tabs-mode: t -*-
#
# This file is part of the LibreOffice project.
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Hillyard Office's Office Defaults: Office's new-document templates, keyboard
# shortcuts and start-up behaviour, as a bundled extension. Its settings live in
# the bundled-extension configuration layer, so the product has them while
# unit tests (core layer only) keep LibreOffice's defaults.

$(eval $(call gb_Package_Package,extras_officedefaults_ext,$(SRCDIR)/extras/source/extensions/office-defaults))

$(eval $(call gb_Package_add_files_with_dir,extras_officedefaults_ext,$(LIBO_SHARE_FOLDER)/extensions/office-defaults,\
	Accelerators.xcu \
	Common.xcu \
	Impress.xcu \
	META-INF/manifest.xml \
	Setup.xcu \
	Writer.xcu \
	description.xml \
))

# vim: set noet sw=4 ts=4:
