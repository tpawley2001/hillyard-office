# -*- Mode: makefile-gmake; tab-width: 4; indent-tabs-mode: t -*-
#
# This file is part of the LibreOffice project.
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#
# Hillyard Office's Office Parity extension (Batch Replace, Office
# compatibility check), shipped as a bundled extension. The tracked copy is
# refreshed from custom/extension by custom/extension/sync_bundled.py.

$(eval $(call gb_Package_Package,extras_office_parity,$(SRCDIR)/extras/source/extensions/office-parity))

$(eval $(call gb_Package_add_files_with_dir,extras_office_parity,$(LIBO_SHARE_FOLDER)/extensions/office-parity,\
	Addons.xcu \
	Jobs.xcu \
	META-INF/manifest.xml \
	batch_job.py \
	compat_job.py \
	description.xml \
	pythonpath/backends.py \
	pythonpath/batch_dialog.py \
	pythonpath/batch_engine.py \
	pythonpath/batch_service.py \
	pythonpath/calc_batch_engine.py \
	pythonpath/compat_check.py \
	pythonpath/compat_inventory.py \
))

# vim: set noet sw=4 ts=4:
