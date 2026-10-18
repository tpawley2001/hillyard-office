# SPDX-License-Identifier: MIT
# Copyright (c) 2026 Hillyard Tech
"""Application backends for the batch replacement command.

Writer and Calc expose the same user-facing capability over different document
models, so the job and dialog talk to a backend rather than to one engine.
Each backend provides the same small surface: parse rules, load a rule file,
preview, apply, and format a preview.

The two search options follow Word and Excel respectively. Writer's second
option is "Find whole words only"; Calc's is "Match entire cell contents",
which is Excel's wording for the same idea at cell granularity.
"""
import json
from pathlib import Path

import batch_engine
import batch_service
import calc_batch_engine


class WriterBackend:
    name = "Writer"
    scope = "document"
    second_option_label = "Find whole words only"
    supports_selection = True

    @staticmethod
    def selection(document):
        return batch_engine.text_selection(document)

    @staticmethod
    def parse_text(text, case_sensitive, second_option):
        return batch_service.parse_rules_text(text, case_sensitive, second_option)

    @staticmethod
    def load_file(path, case_sensitive, second_option):
        return batch_service.load_rules_file(path, case_sensitive, second_option)

    @staticmethod
    def preview(document, rules, within=None):
        return batch_engine.preview(document, rules, within)

    @staticmethod
    def apply(plan):
        return batch_engine.apply(plan)


class CalcBackend:
    name = "Calc"
    scope = "workbook"
    second_option_label = "Match entire cell contents"
    # Calc selects cells, not text; selection scope there needs its own design.
    supports_selection = False

    @staticmethod
    def selection(document):
        return None

    @staticmethod
    def parse_text(text, case_sensitive, second_option):
        rules = []
        for number, line in enumerate(text.splitlines(), 1):
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            search, separator, replacement = line.partition("=>")
            if not separator:
                raise ValueError(f"Line {number}: write 'find => replacement'.")
            search, replacement = search.strip(), replacement.strip()
            if not search:
                raise ValueError(f"Line {number}: the find phrase is empty.")
            rules.append(calc_batch_engine.Rule(search, replacement, None,
                                                case_sensitive, second_option))
        if not rules:
            raise ValueError("Enter at least one 'find => replacement' rule.")
        return rules

    @staticmethod
    def load_file(path, case_sensitive, second_option):
        data = json.loads(Path(path).read_text(encoding="utf-8"))
        if not isinstance(data, list):
            raise ValueError("A rule file must contain a JSON array.")
        rules = []
        for index, entry in enumerate(data, 1):
            if isinstance(entry, dict):
                try:
                    search, replacement = entry["search"], entry["replacement"]
                except KeyError as missing:
                    raise ValueError(f"Rule {index} has no {missing} field.") from None
                expected = entry.get("expected")
                rule_case = entry.get("case_sensitive", case_sensitive)
                rule_second = entry.get("entire_cell", second_option)
                look_in = entry.get("look_in", "values")
            elif isinstance(entry, (list, tuple)) and len(entry) in (2, 3):
                search, replacement = entry[0], entry[1]
                expected = entry[2] if len(entry) == 3 else None
                rule_case, rule_second, look_in = case_sensitive, second_option, "values"
            else:
                raise ValueError(f"Rule {index} must be a [search, replacement] pair "
                                 "or an object with search and replacement fields.")
            if not isinstance(search, str) or not search:
                raise ValueError(f"Rule {index} needs a nonempty text search phrase.")
            if not isinstance(replacement, str):
                raise ValueError(f"Rule {index} needs a text replacement.")
            for flag in (rule_case, rule_second):
                if not isinstance(flag, bool):
                    raise ValueError(f"Rule {index} needs its options to be true or false.")
            if look_in not in ("values", "formulas"):
                raise ValueError(f"Rule {index} must look in values or in formulas.")
            if expected is not None and (not isinstance(expected, int)
                                         or isinstance(expected, bool) or expected < 0):
                raise ValueError(f"Rule {index} has a non-negative integer count, or none.")
            rules.append(calc_batch_engine.Rule(search, replacement, expected,
                                                rule_case, rule_second, look_in))
        if not rules:
            raise ValueError("The rule file has no rules.")
        return rules

    @staticmethod
    def preview(document, rules, within=None):
        if within is not None:
            raise ValueError("Calc batches always cover the whole workbook.")
        return calc_batch_engine.preview(document, rules)

    @staticmethod
    def apply(plan):
        return calc_batch_engine.apply(plan)


def for_document(document):
    """Pick the backend for a document, or explain why there is none."""
    if document is None:
        raise ValueError("Open a Writer document or a Calc spreadsheet first.")
    if document.supportsService("com.sun.star.text.TextDocument"):
        return WriterBackend
    if document.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
        return CalcBackend
    raise ValueError("Batch Replace works in Writer documents and Calc spreadsheets.")
