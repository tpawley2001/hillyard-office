# SPDX-License-Identifier: MIT
# Copyright (c) 2026 Hillyard Tech
"""Checked batch replacement for Calc, through the spreadsheet document model.

Rules mirror the options Excel's Find and Replace offers: "Match case" and
"Match entire cell contents". Scope is the whole workbook, which is Excel's
"Within: Workbook".

Matches are found cell by cell rather than with the range API's findAll().
That is deliberate: Calc's findAll() returns *merged ranges*, so adjacent
matches collapse into one entry. On a sheet holding "2026" in A1 and
"2026 season" in A2, findAll() reports 2 while replaceAll() changes 3 cells.
A preview built on that count would disagree with the replacement it is
supposed to be previewing. See custom/local/probe_calc.py for the observation.

Every rule is matched against the original cell text and all edits to a cell
are computed together, so a replacement can never become the input of a later
rule.
"""
from dataclasses import dataclass, field

MAXCOL, MAXROW = 1023, 1048575


@dataclass(frozen=True)
class Rule:
    search: str
    replacement: str
    expected: object = None
    case_sensitive: bool = True
    entire_cell: bool = False
    look_in: str = "values"


@dataclass
class Cell:
    sheet: object
    sheet_name: str
    column: int
    row: int
    original: str
    replaced: str
    rules: tuple
    formula: bool

    @property
    def address(self):
        return f"{self.sheet_name}!{column_to_name(self.column)}{self.row + 1}"


@dataclass
class Plan:
    document: object
    rules: list
    cells: list = field(default_factory=list)
    counts: list = field(default_factory=list)


def as_rule(entry):
    if isinstance(entry, Rule):
        return entry
    if isinstance(entry, (tuple, list)) and len(entry) == 2:
        return Rule(entry[0], entry[1])
    raise ValueError("Each rule must be a Rule or a (find, replacement) pair.")


def find_spans(text, rule):
    """Non-overlapping (start, end) spans of rule.search within text."""
    if not text:
        return []
    if rule.entire_cell:
        if rule.case_sensitive:
            return [(0, len(text))] if text == rule.search else []
        return [(0, len(text))] if text.lower() == rule.search.lower() else []
    haystack = text if rule.case_sensitive else text.lower()
    needle = rule.search if rule.case_sensitive else rule.search.lower()
    spans, start = [], 0
    while True:
        index = haystack.find(needle, start)
        if index < 0:
            return spans
        spans.append((index, index + len(needle)))
        start = index + len(needle)


def combine(source, spans, rules):
    """Apply every span to source at once, so rules cannot feed each other."""
    out, position = [], 0
    for start, end, rule_index in spans:
        out.append(source[position:start])
        out.append(rules[rule_index].replacement)
        position = end
    out.append(source[position:])
    return "".join(out)


def used_area(sheet):
    """The sheet's used area as (last_column, last_row), or None when empty."""
    cursor = sheet.createCursor()
    cursor.gotoEndOfUsedArea(False)
    address = cursor.RangeAddress
    if address.EndColumn < 0 or address.EndRow < 0:
        return None
    return address.EndColumn, address.EndRow


def preview(document, rules):
    if document.isReadonly():
        raise ValueError("This spreadsheet is read-only. Open an editable copy first.")
    if not rules:
        raise ValueError("Add at least one replacement.")
    rules = [as_rule(entry) for entry in rules]
    seen = set()
    for rule in rules:
        if not rule.search:
            raise ValueError("Find text cannot be empty.")
        if rule.search == rule.replacement:
            raise ValueError("Find and replacement text are identical.")
        if rule.look_in not in ("values", "formulas"):
            raise ValueError("Search must look in values or in formulas.")
        key = (rule.search, rule.case_sensitive, rule.entire_cell, rule.look_in)
        if key in seen:
            raise ValueError("Each find phrase and option combination may appear only once.")
        seen.add(key)

    sheets = document.Sheets
    cells, counts = [], [0] * len(rules)
    for index in range(sheets.Count):
        sheet = sheets.getByIndex(index)
        area = used_area(sheet)
        if area is None:
            continue
        last_column, last_row = area
        for row in range(last_row + 1):
            for column in range(last_column + 1):
                cell = sheet.getCellByPosition(column, row)
                formula = cell.getFormula().startswith("=")
                source = cell.getFormula() if formula else cell.getString()
                if not source:
                    continue
                spans = []
                for rule_index, rule in enumerate(rules):
                    if (rule.look_in == "formulas") != formula:
                        continue
                    for start, end in find_spans(source, rule):
                        spans.append((start, end, rule_index))
                if not spans:
                    continue
                spans.sort()
                for left, right in zip(spans, spans[1:]):
                    if left[1] > right[0]:
                        raise ValueError(
                            f"Two rules overlap in {sheet.Name} cell "
                            f"{column_to_name(column)}{row + 1}. "
                            "Use separate, more specific phrases.")
                replaced = combine(source, spans, rules)
                if replaced == source:
                    continue
                for _, _, rule_index in spans:
                    counts[rule_index] += 1
                cells.append(Cell(sheet, sheet.Name, column, row, source, replaced,
                                  tuple(sorted({span[2] for span in spans})), formula))
    return Plan(document, list(rules), cells, counts)


def column_to_name(column):
    """0 -> A, 25 -> Z, 26 -> AA."""
    name = ""
    column += 1
    while column:
        column, remainder = divmod(column - 1, 26)
        name = chr(65 + remainder) + name
    return name


def apply(plan):
    document = plan.document
    if document.isReadonly():
        raise ValueError("The spreadsheet is now read-only.")
    for cell in plan.cells:
        current = cell.sheet.getCellByPosition(cell.column, cell.row)
        actual = current.getFormula() if cell.formula else current.getString()
        if actual != cell.original:
            raise ValueError("A cell changed after preview. Preview the replacements again.")
    if not all(plan.counts):
        raise ValueError("At least one phrase has no matches. Correct or remove it before applying.")
    undo = document.getUndoManager()
    if undo.isLocked():
        raise ValueError("Undo is unavailable. Finish the current operation before replacing.")
    document.lockControllers()
    undo.enterUndoContext("Batch replace")
    changed = False
    error = None
    try:
        for cell in plan.cells:
            target = cell.sheet.getCellByPosition(cell.column, cell.row)
            if cell.formula:
                target.setFormula(cell.replaced)
            else:
                target.setString(cell.replaced)
            changed = True
    except Exception as exc:
        error = exc
    finally:
        undo.leaveUndoContext()
        document.unlockControllers()
    if error is not None:
        if changed:
            undo.undo()
        raise error
    return sum(plan.counts)
