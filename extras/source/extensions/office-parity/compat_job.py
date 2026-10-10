# SPDX-License-Identifier: MPL-2.0
r"""Office compatibility check: a menu command and an after-save check.

CompatCheckJob is the menu command (Office Parity > Check Office
Compatibility...). It checks the current document and shows the result.
Automation passes a percent-encoded JSON object instead of nothing:

  {"report": "/path/report.json"}

and gets the findings written there, so the packaged component can be tested
without a display. "mode" is "check" (the default, as the menu), "saved"
(check the file the document was saved to) or "last" (what the after-save
check found the last time it ran on this document).

SaveCheckJob runs on the OnSaveDone and OnSaveAsDone events (Jobs.xcu). When
the document was saved in a Microsoft format, it reloads the saved file and,
if anything was lost or changed, shows a warning bar on the document with a
Details button. A clean save shows nothing.
"""
import json
from pathlib import Path
from urllib.parse import unquote

import unohelper

from com.sun.star.task import XJob, XJobExecutor

import compat_check
import compat_inventory

MENU_IMPLEMENTATION = "org.officeparity.compat.Check"
SAVE_IMPLEMENTATION = "org.officeparity.compat.SaveCheck"
INFOBAR_ID = "org.officeparity.compat"
# The last after-save result per document, for the Details button.
LAST_RESULTS = {}


def current_document(ctx):
    desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    return desktop.getCurrentComponent()


def show_message(ctx, document, title, text, warning):
    from com.sun.star.awt.MessageBoxType import INFOBOX, WARNINGBOX
    toolkit = ctx.ServiceManager.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
    parent = document.getCurrentController().getFrame().getContainerWindow()
    box = toolkit.createMessageBox(parent, WARNINGBOX if warning else INFOBOX, 1, title, text)
    box.execute()
    box.dispose()


def show_findings(ctx, document, findings, name):
    show_message(ctx, document, "Office Compatibility",
                 compat_inventory.describe(findings, name), bool(findings))


class CompatCheckJob(unohelper.Base, XJobExecutor):
    def __init__(self, ctx):
        self.ctx = ctx

    def trigger(self, args):
        document = current_document(self.ctx)
        if args.strip() == "details":
            result = LAST_RESULTS.get(document.getURL())
            if result is not None:
                show_findings(self.ctx, document, *result)
                return
        if args.strip() and args.strip() != "details":
            self.run_automation(document, json.loads(unquote(args)))
            return
        try:
            findings, name = compat_check.check(self.ctx, document)
        except Exception as error:
            show_message(self.ctx, document, "Office Compatibility",
                         f"The check could not run: {error}", True)
            return
        show_findings(self.ctx, document, findings, name)

    def run_automation(self, document, options):
        report = {"status": "observed"}
        try:
            mode = options.get("mode", "check")
            if mode == "last":
                # What the after-save check last found for this document.
                result = LAST_RESULTS.get(document.getURL())
                if result is None:
                    report["status"] = "not_checked"
                    result = ([], None)
            elif mode == "saved":
                result = compat_check.check_saved(self.ctx, document)
                if result is None:
                    report["status"] = "not_checked"
                    result = ([], None)
            else:
                result = compat_check.check(self.ctx, document)
            report["findings"], report["format"] = result
            report["summary"] = compat_inventory.describe(*result) if result[1] else None
        except Exception as error:
            report = {"status": "error", "error": f"{type(error).__name__}: {error}"}
        Path(options["report"]).expanduser().write_text(
            json.dumps(report, indent=2) + "\n", encoding="utf-8")


class SaveCheckJob(unohelper.Base, XJob):
    def __init__(self, ctx):
        self.ctx = ctx

    def execute(self, args):
        environment = next((a.Value for a in args if a.Name == "Environment"), ())
        document = next((e.Value for e in environment if e.Name == "Model"), None)
        if document is None or compat_check.application(document) is None:
            return None
        try:
            result = compat_check.check_saved(self.ctx, document)
        except Exception:
            return None          # never let the check get in the way of a save
        controller = document.getCurrentController()
        if result is None or controller is None:
            return None
        findings, name = result
        LAST_RESULTS[document.getURL()] = result
        try:
            if controller.hasInfobar(INFOBAR_ID):
                controller.removeInfobar(INFOBAR_ID)
            if findings:
                from com.sun.star.beans import StringPair
                from com.sun.star.frame.InfobarType import WARNING
                losses = sum(f["severity"] == "loss" for f in findings)
                button = StringPair("Details", f"service:{MENU_IMPLEMENTATION}?details")
                primary = (f"Saved as {name}, but {losses} thing(s) were lost."
                           if losses else f"Saved as {name}; some things changed.")
                controller.appendInfobar(INFOBAR_ID, primary,
                                         "; ".join(f["message"] for f in findings[:3]),
                                         WARNING, (button,), True)
        except Exception:
            pass                 # controllers without infobars (headless) are fine
        return None


g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    CompatCheckJob, MENU_IMPLEMENTATION, ("com.sun.star.task.Job",))
g_ImplementationHelper.addImplementation(
    SaveCheckJob, SAVE_IMPLEMENTATION, ("com.sun.star.task.Job",))
