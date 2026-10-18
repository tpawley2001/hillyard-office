# SPDX-License-Identifier: MIT
# Copyright (c) 2026 Hillyard Tech
"""Rule input and reporting for the batch replacement engine.

This module imports no UNO modules, so its parsing and formatting can be
unit-tested without an office runtime. Rules are batch_engine.Rule values.
"""
import json
from pathlib import Path

from batch_engine import Rule


def parse_rules_text(text, case_sensitive=True, whole_word=False):
    """Read one 'search => replacement' rule per line; blank lines and # are skipped.

    The two options apply to every rule typed here, matching how Word's Find and
    Replace dialog options apply to the whole operation.
    """
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
        rules.append(Rule(search, replacement, None, case_sensitive, whole_word))
    if not rules:
        raise ValueError("Enter at least one 'find => replacement' rule.")
    return rules


def load_rules_file(path, case_sensitive=True, whole_word=False):
    """Load the project rule formats: [[search, replacement, expected], ...] or
    [{"search": ..., "replacement": ..., "expected": ..., "case_sensitive": ...,
    "whole_word": ...}, ...].

    Array entries cannot carry options and use the two defaults. Object entries
    may override either option per rule.
    """
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
            rule_word = entry.get("whole_word", whole_word)
            for name, value in (("case_sensitive", rule_case), ("whole_word", rule_word)):
                if not isinstance(value, bool):
                    raise ValueError(f"Rule {index} needs {name} to be true or false.")
        elif isinstance(entry, (list, tuple)) and len(entry) in (2, 3):
            search, replacement = entry[0], entry[1]
            expected = entry[2] if len(entry) == 3 else None
            rule_case, rule_word = case_sensitive, whole_word
        else:
            raise ValueError(f"Rule {index} must be a [search, replacement] pair "
                             "or an object with search and replacement fields.")
        if not isinstance(search, str) or not search:
            raise ValueError(f"Rule {index} needs a nonempty text search phrase.")
        if not isinstance(replacement, str):
            raise ValueError(f"Rule {index} needs a text replacement.")
        if expected is not None and (not isinstance(expected, int) or isinstance(expected, bool)
                                     or expected < 0):
            raise ValueError(f"Rule {index} has a non-negative integer count, or none.")
        rules.append(Rule(search, replacement, expected, rule_case, rule_word))
    if not rules:
        raise ValueError("The rule file has no rules.")
    return rules


def second_option(rule):
    """The application's second search option: whole words, or entire cell."""
    value = getattr(rule, "whole_word", None)
    if value is None:
        value = getattr(rule, "entire_cell", False)
    return bool(value)


def describe_options(rule):
    """A short suffix naming any non-default search option, for reports.

    Writer rules carry whole_word; Calc rules carry entire_cell. The label
    follows the application's own wording.
    """
    label = "whole words" if hasattr(rule, "whole_word") else "entire cell"
    second = second_option(rule)
    if not rule.case_sensitive and second:
        return f" [ignore case, {label}]"
    if not rule.case_sensitive:
        return " [ignore case]"
    if second:
        return f" [{label}]"
    return ""


def rule_reports(rules, counts):
    """Per-rule observations, in the same shape as the runtime probe report."""
    reports = []
    for rule, found in zip(rules, counts):
        observation = {"search": rule.search, "replacement": rule.replacement, "found": found}
        if not rule.case_sensitive:
            observation["case_sensitive"] = False
        if second_option(rule):
            observation["second_option"] = True
        if rule.expected is not None:
            observation["expected"] = rule.expected
            observation["matches_expected"] = found == rule.expected
        reports.append(observation)
    return reports


def mismatches(rules, counts):
    """Return a message for the first rule whose count differs from its expectation."""
    for rule, found in zip(rules, counts):
        if rule.expected is not None and found != rule.expected:
            return (f"{rule.search!r} was found {found} times but {rule.expected} "
                    "replacement(s) were expected.")
    return None


def format_preview(rules, counts):
    """Human-readable preview lines for a dialog or log."""
    lines = []
    for rule, found in zip(rules, counts):
        line = f"{found} x  {rule.search}  ->  {rule.replacement}{describe_options(rule)}"
        if rule.expected is not None:
            line += f"   [expected {rule.expected}: {'ok' if found == rule.expected else 'MISMATCH'}]"
        lines.append(line)
    return "\n".join(lines)
