# SPDX-License-Identifier: MIT
# Copyright (c) 2026 Hillyard Tech
"""Interactive rules dialog for the batch replacement engine.

Preview is mandatory: applying always uses a plan built against the document's
current text, so a stale preview cannot edit a document that changed underneath
it. Replacement is a single undo action.
"""
import unohelper

from com.sun.star.awt import XActionListener

import batch_service


INTRO = "One 'find => replacement' rule per line. Blank lines and # comments are ignored."
PLACEHOLDER = "32nd => 33rd\nSunday February 8, 2026 => Sunday February 7, 2027"


class RuleDialog(unohelper.Base, XActionListener):
    def __init__(self, ctx, document, backend, initial_rules=PLACEHOLDER):
        self.ctx = ctx
        self.document = document
        self.backend = backend
        self.plan = None
        self.model = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.awt.UnoControlDialogModel", ctx)
        self.model.Title = "Batch Replace"
        self.model.Width = 280
        self.model.Height = 172
        self.add_label("intro", INTRO, 5, 5, 270, 8)
        self.add_edit("rules", 5, 15, 270, 55, initial_rules)
        self.add_label("results_label", "Preview", 5, 73, 270, 8)
        self.add_edit("results", 5, 83, 270, 55, "", read_only=True)
        # The same two options Word's and Excel's Find and Replace dialogs offer;
        # the second is named per application.
        self.add_checkbox("match_case", "Match case", 5, 141, 85, 8, True)
        self.add_checkbox("whole_word", backend.second_option_label, 95, 141, 110, 8, False)
        # Word's Replace All acts on the selection when there is one, so start
        # checked exactly when text is selected.
        selected = backend.supports_selection and backend.selection(document) is not None
        self.add_checkbox("selection_only", "Selection only", 205, 141, 70, 8, selected)
        self.model.getByName("selection_only").Enabled = backend.supports_selection
        self.add_button("preview", "Preview", 5, 152, 55, 14)
        self.add_button("apply", "Replace All", 65, 152, 55, 14)
        self.add_button("close", "Close", 125, 152, 55, 14)
        self.dialog = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.awt.UnoControlDialog", ctx)
        self.dialog.setModel(self.model)
        self.set_apply(False)

    def add_label(self, name, label, x, y, width, height):
        control = self.model.createInstance("com.sun.star.awt.UnoControlFixedTextModel")
        control.Name = name
        control.Label = label
        control.PositionX, control.PositionY = x, y
        control.Width, control.Height = width, height
        self.model.insertByName(name, control)

    def add_edit(self, name, x, y, width, height, text, read_only=False):
        control = self.model.createInstance("com.sun.star.awt.UnoControlEditModel")
        control.Name = name
        control.PositionX, control.PositionY = x, y
        control.Width, control.Height = width, height
        control.MultiLine = True
        control.VScroll = True
        control.ReadOnly = read_only
        control.Text = text
        self.model.insertByName(name, control)

    def add_button(self, name, label, x, y, width, height):
        control = self.model.createInstance("com.sun.star.awt.UnoControlButtonModel")
        control.Name = name
        control.Label = label
        control.PositionX, control.PositionY = x, y
        control.Width, control.Height = width, height
        self.model.insertByName(name, control)

    def add_checkbox(self, name, label, x, y, width, height, checked):
        control = self.model.createInstance("com.sun.star.awt.UnoControlCheckBoxModel")
        control.Name = name
        control.Label = label
        control.PositionX, control.PositionY = x, y
        control.Width, control.Height = width, height
        control.State = 1 if checked else 0
        self.model.insertByName(name, control)

    def options(self):
        """The two search options as Word's Find and Replace dialog presents them."""
        return (self.model.getByName("match_case").State == 1,
                self.model.getByName("whole_word").State == 1)

    def set_apply(self, enabled):
        self.model.getByName("apply").Enabled = enabled

    def do_preview(self):
        """Build a plan for the current rules; report problems in the results box."""
        try:
            case_sensitive, second_option = self.options()
            rules = self.backend.parse_text(self.model.getByName("rules").Text,
                                            case_sensitive, second_option)
            within = None
            if self.model.getByName("selection_only").State == 1:
                within = self.backend.selection(self.document)
                if within is None:
                    raise ValueError("Nothing is selected. Select text, or clear Selection only.")
            plan = self.backend.preview(self.document, rules, within)
        except ValueError as error:
            self.plan = None
            self.model.getByName("results").Text = str(error)
            self.set_apply(False)
            return False
        self.plan = plan
        text = batch_service.format_preview(rules, plan.counts)
        if getattr(plan, "selection", False):
            text = "In the selection only:\n" + text
        self.model.getByName("results").Text = text
        self.set_apply(all(plan.counts))
        return True

    def do_apply(self):
        if self.plan is None and not self.do_preview():
            return False
        try:
            replaced = self.backend.apply(self.plan)
        except ValueError as error:
            self.plan = None
            self.model.getByName("results").Text = str(error)
            self.set_apply(False)
            return False
        self.plan = None
        self.model.getByName("results").Text = (
            f"Replaced {replaced} match(es). Use Undo once to revert the whole batch.")
        self.set_apply(False)
        return True

    def actionPerformed(self, event):
        name = event.Source.getModel().Name
        if name == "preview":
            self.do_preview()
        elif name == "apply":
            self.do_apply()
        elif name == "close":
            self.dialog.endExecute()

    def disposing(self, event):
        """XEventListener: the control that we listen to is going away."""
        pass

    def create_peer(self):
        """Build the real widgets; separated from show() so tests can drive the
        dialog without entering its modal loop."""
        toolkit = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.awt.Toolkit", self.ctx)
        try:
            parent = self.document.CurrentController.Frame.ContainerWindow
        except Exception:
            parent = None
        self.dialog.createPeer(toolkit, parent)
        for name in ("preview", "apply", "close"):
            self.dialog.getControl(name).addActionListener(self)
        return self.dialog

    def show(self):
        self.create_peer()
        self.dialog.execute()
