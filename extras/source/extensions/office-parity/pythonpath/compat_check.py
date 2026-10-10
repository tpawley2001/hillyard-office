# SPDX-License-Identifier: MPL-2.0
"""Run the compatibility check on an open document.

check(ctx, document) saves a temporary copy in the Office format the document
belongs to (or the Office format it was last saved in), reloads that copy
hidden, and compares inventories. check_saved(ctx, document) does the same
against the file the document was just saved to, which is what the automatic
after-save check uses.
"""
import os
from pathlib import Path
import tempfile

import uno
from com.sun.star.beans import PropertyValue

import compat_inventory

OFFICE_FILTERS = {
    "writer": ("MS Word 2007 XML", ".docx", "Word"),
    "calc": ("Calc MS Excel 2007 XML", ".xlsx", "Excel"),
    "impress": ("Impress MS PowerPoint 2007 XML", ".pptx", "PowerPoint"),
}
# Filters for which a save is checked automatically, with the name shown.
MS_FILTERS = {
    "MS Word 2007 XML": "Word", "Office Open XML Text": "Word", "MS Word 97": "Word 97-2003",
    "MS Word 2007 XML Template": "Word template", "Rich Text Format": "RTF",
    "Calc MS Excel 2007 XML": "Excel", "Calc Office Open XML": "Excel",
    "MS Excel 97": "Excel 97-2003", "Calc MS Excel 2007 XML Template": "Excel template",
    "Impress MS PowerPoint 2007 XML": "PowerPoint", "Impress Office Open XML": "PowerPoint",
    "MS PowerPoint 97": "PowerPoint 97-2003",
}
# Larger files are not checked automatically after a save: reloading them
# would hold up the user. The menu command still checks them.
AUTO_CHECK_LIMIT = 15 * 1024 * 1024


def _prop(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


def application(document):
    if document.supportsService("com.sun.star.text.TextDocument"):
        return "writer"
    if document.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
        return "calc"
    if document.supportsService("com.sun.star.presentation.PresentationDocument"):
        return "impress"
    return None


def current_filter(document):
    for arg in document.getArgs():
        if arg.Name == "FilterName":
            return arg.Value
    return ""


def _load_hidden(ctx, url):
    desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    return desktop.loadComponentFromURL(url, "_blank", 0, (
        _prop("Hidden", True), _prop("ReadOnly", True),
        _prop("MacroExecutionMode", 0), _prop("UpdateDocMode", 0)))


def _inventory_of_file(ctx, url):
    copy = _load_hidden(ctx, url)
    if copy is None:
        raise RuntimeError("The saved copy could not be opened again.")
    try:
        return compat_inventory.inventory(copy)
    finally:
        copy.close(True)


def check(ctx, document):
    """(findings, format name) for saving the open document in an Office format."""
    app = application(document)
    if app is None:
        raise ValueError("Only text documents, spreadsheets and presentations can be checked.")
    filter_name = current_filter(document)
    if filter_name in MS_FILTERS:
        name = MS_FILTERS[filter_name]
        suffix = os.path.splitext(document.getURL())[1] or OFFICE_FILTERS[app][1]
    else:
        filter_name, suffix, name = OFFICE_FILTERS[app]
    before = compat_inventory.inventory(document)
    handle, path = tempfile.mkstemp(prefix="hillyard-compat-", suffix=suffix)
    os.close(handle)
    try:
        url = uno.systemPathToFileUrl(path)
        document.storeToURL(url, (_prop("FilterName", filter_name), _prop("Overwrite", True)))
        after = _inventory_of_file(ctx, url)
    finally:
        try:
            os.unlink(path)
        except OSError:
            pass
    return compat_inventory.compare(before, after), name


def check_saved(ctx, document):
    """(findings, format name) for the file the document was just saved to, or
    None when the save was not to an Office format or the file is too large."""
    name = MS_FILTERS.get(current_filter(document))
    url = document.getURL()
    if name is None or not url.startswith("file:"):
        return None
    try:
        if Path(uno.fileUrlToSystemPath(url)).stat().st_size > AUTO_CHECK_LIMIT:
            return None
    except OSError:
        return None
    before = compat_inventory.inventory(document)
    after = _inventory_of_file(ctx, url)
    return compat_inventory.compare(before, after), name
