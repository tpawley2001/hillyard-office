# SPDX-License-Identifier: MPL-2.0
r"""Batch replacement job for Writer and Calc.

A menu invocation passes no arguments and opens the rules dialog for the
current document, choosing the backend from the document type. Automation
passes a percent-encoded JSON object instead:

  {"mode": "apply", "rules": "/path/rules.json", "report": "/path/report.json",
   "selection": false}

"selection": true limits the batch to the document's current text selection
and fails if nothing is selected.

The URL is built with urllib.parse.quote, because the dispatch layer may
percent-encode the argument string before it reaches trigger(). The automation
form exists so the packaged component can be exercised without a display; it
runs only for a caller that can already dispatch a command URL.
"""
import json
from pathlib import Path
from urllib.parse import unquote

import unohelper

from com.sun.star.task import XJobExecutor

import backends
import batch_dialog
import batch_service


IMPLEMENTATION_NAME = "org.officeparity.writer.batchreplace.Job"


def current_document(ctx):
    desktop = ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.frame.Desktop", ctx)
    return desktop.getCurrentComponent()


class BatchReplaceJob(unohelper.Base, XJobExecutor):
    def __init__(self, ctx):
        self.ctx = ctx

    def trigger(self, args):
        if args.strip():
            self.run_automation(json.loads(unquote(args)))
        else:
            document = current_document(self.ctx)
            backend = backends.for_document(document)
            batch_dialog.RuleDialog(self.ctx, document, backend).show()

    def run_automation(self, options):
        report = {"mode": options.get("mode", "apply"), "rules": [], "status": "observed"}
        try:
            document = current_document(self.ctx)
            backend = backends.for_document(document)
            report["application"] = backend.name
            report["scope"] = backend.scope
            report["document"] = document.getURL()
            rules = backend.load_file(options["rules"],
                                      options.get("case_sensitive", True),
                                      options.get("second_option", False))
            within = None
            if options.get("selection", False):
                within = backend.selection(document)
                if within is None:
                    raise ValueError("Selection scope was requested but nothing is selected.")
                report["scope"] = "selection"
            plan = backend.preview(document, rules, within)
            report["rules"] = batch_service.rule_reports(rules, plan.counts)
            mismatch = batch_service.mismatches(rules, plan.counts)
            if mismatch:
                report["status"] = "expectation_mismatch"
                report["error"] = mismatch
            elif report["mode"] == "apply":
                report["replaced"] = backend.apply(plan)
        except Exception as error:
            report["status"] = "error"
            report["error"] = f"{type(error).__name__}: {error}"
        Path(options["report"]).expanduser().write_text(
            json.dumps(report, indent=2) + "\n", encoding="utf-8")
        return report


g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    BatchReplaceJob, IMPLEMENTATION_NAME, ("com.sun.star.task.Job",))
