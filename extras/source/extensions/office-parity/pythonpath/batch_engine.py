# SPDX-License-Identifier: MIT
# Copyright (c) 2026 Hillyard Tech
"""Literal batch replacement through Writer's document model.

Preview freezes the original matches. Application runs from the end of each
text container, so replacements never become inputs to later rules.

Rules carry the two search options Word's Find and Replace dialog exposes,
"Match case" and "Find whole words only". Case sensitivity defaults to on,
which is this tool's established behavior and not Word's default; callers that
want Word's default pass case_sensitive=False.

Like Word's Replace All, a batch can be limited to the current selection:
preview() then keeps only matches that lie wholly inside a selected range.
"""
from dataclasses import dataclass
from functools import cmp_to_key


@dataclass(frozen=True)
class Rule:
    search: str
    replacement: str
    expected: object = None
    case_sensitive: bool = True
    whole_word: bool = False


def as_rule(entry):
    """Accept a Rule, or a plain (find, replacement) pair."""
    if isinstance(entry, Rule):
        return entry
    if isinstance(entry, (tuple, list)) and len(entry) == 2:
        return Rule(entry[0], entry[1])
    raise ValueError("Each rule must be a Rule or a (find, replacement) pair.")


def matches_search(actual, rule):
    """Whether document text is the rule's phrase under the rule's options."""
    if rule.case_sensitive:
        return actual == rule.search
    return actual.lower() == rule.search.lower()


def same_value(left, right):
    """Whether two UNO property values compare equal; unknown types do not."""
    try:
        return left == right
    except Exception:
        return False


@dataclass
class Match:
    range: object
    rule: int
    original: str


@dataclass
class Plan:
    document: object
    rules: list
    groups: list
    counts: list
    selection: bool = False


def text_selection(document):
    """The document's selected text ranges, or None when nothing is selected.

    A collapsed cursor is not a selection. Selections that are not text
    (a table-cell block, a shape, a picture) are also treated as none, so the
    caller falls back to the whole document, as Word does.
    """
    try:
        selection = document.getCurrentController().getSelection()
    except Exception:
        return None
    if selection is None or not selection.supportsService("com.sun.star.text.TextRanges"):
        return None
    ranges = [selection.getByIndex(i) for i in range(selection.getCount())]
    ranges = [r for r in ranges if r.getString()]
    return ranges or None


def inside(match, ranges):
    """Whether a found range lies wholly inside one of the selected ranges.

    Ranges in a different text container (a table cell or frame inside a body
    selection) cannot be compared by Writer and are treated as outside.
    """
    for selected in ranges:
        owner = selected.getText()
        try:
            if (owner.compareRegionStarts(selected, match) >= 0
                    and owner.compareRegionEnds(match, selected) >= 0):
                return True
        except Exception:
            continue
    return False


def preview(document, rules, within=None):
    """Freeze the matches for each rule; `within` limits them to selected ranges."""
    if document.isReadonly():
        raise ValueError("This document is read-only. Open an editable copy first.")
    if not rules:
        raise ValueError("Add at least one replacement.")
    rules = [as_rule(entry) for entry in rules]
    seen = set()
    for rule in rules:
        if not rule.search:
            raise ValueError("Find text cannot be empty.")
        if rule.search == rule.replacement:
            raise ValueError("Find and replacement text are identical.")
        key = (rule.search, rule.case_sensitive, rule.whole_word)
        if key in seen:
            raise ValueError("Each find phrase and option combination may appear only once.")
        seen.add(key)
    groups = []
    counts = []
    for index, rule in enumerate(rules):
        descriptor = document.createSearchDescriptor()
        descriptor.SearchString = rule.search
        descriptor.SearchCaseSensitive = rule.case_sensitive
        descriptor.SearchRegularExpression = False
        descriptor.SearchWords = rule.whole_word
        descriptor.SearchStyles = False
        found = document.findAll(descriptor)
        targets = [found.getByIndex(i) for i in range(found.getCount())]
        if within is not None:
            targets = [target for target in targets if inside(target, within)]
        counts.append(len(targets))
        for target in targets:
            # A case-insensitive match returns the document's own text, which
            # may differ in case from the search phrase. Replacements must act
            # on what is actually there.
            actual = target.getString()
            if not matches_search(actual, rule):
                raise ValueError("A match contains a field or unsupported text. Use a more specific phrase.")
            owner = target.getText()
            group = next((group for group in groups if group[0] == owner), None)
            if group is None:
                group = (owner, [])
                groups.append(group)
            group[1].append(Match(target, index, actual))
    for owner, matches in groups:
        matches.sort(key=cmp_to_key(lambda a, b: -owner.compareRegionStarts(a.range, b.range)))
        for left, right in zip(matches, matches[1:]):
            if owner.compareRegionStarts(left.range.getEnd(), right.range.getStart()) < 0:
                raise ValueError("Two replacements overlap. Use separate, more specific phrases.")
    return Plan(document, list(rules), groups, counts, within is not None)


def replacement_operations(owner, match, replacement):
    """Prepare and validate all cursors before making any changes."""
    if len(match.original) == len(replacement):
        for index in reversed(range(len(replacement))):
            if match.original[index] == replacement[index]:
                continue
            cursor = owner.createTextCursorByRange(match.range.getStart())
            # Writer's cursor moves by characters, including surrogate pairs.
            if index and not cursor.goRight(index, False):
                raise ValueError("A match moved outside its text container.")
            if not cursor.goRight(1, True):
                raise ValueError("A match could not be selected for replacement.")
            if cursor.getString() != match.original[index]:
                raise ValueError("This character sequence needs manual replacement.")
            names = tuple(p.Name for p in cursor.getPropertySetInfo().getProperties()
                          if p.Name.startswith("Char") and not p.Attributes & 16)
            values = cursor.getPropertyValues(names)
            states = cursor.getPropertyStates(names)
            defaults = cursor.getPropertyDefaults(names)
            # Writer reports character formatting that covers a whole paragraph
            # as DEFAULT_VALUE even though its value differs from the default,
            # because it lives on the paragraph rather than on the text node.
            # Replacing the text loses it, so carry every non-default value
            # rather than trusting the reported state alone.
            direct = [(name, value) for name, value, state, default
                      in zip(names, values, states, defaults)
                      if value is not None
                      and (state.value == "DIRECT_VALUE" or not same_value(value, default))]
            reset = tuple(name for name, value, state, default
                          in zip(names, values, states, defaults)
                          if state.value == "DEFAULT_VALUE" and same_value(value, default))
            yield owner, match, index, replacement[index], direct, reset
    else:
        yield owner, match, None, replacement, [], ()


def apply(plan):
    document = plan.document
    if document.isReadonly():
        raise ValueError("The document is now read-only.")
    for _, matches in plan.groups:
        for match in matches:
            if match.range.getString() != match.original:
                raise ValueError("The document changed after preview. Preview the replacements again.")
    if not all(plan.counts):
        raise ValueError("At least one phrase has no matches. Correct or remove it before applying.")
    undo = document.getUndoManager()
    if undo.isLocked():
        raise ValueError("Undo is unavailable. Finish the current operation before replacing text.")
    operations = [operation for owner, matches in plan.groups
                  for match in reversed(matches)
                  for operation in replacement_operations(owner, match, plan.rules[match.rule].replacement)]
    changed = False
    document.lockControllers()
    undo.enterUndoContext("Review replacements")
    error = None
    try:
        for owner, match, index, replacement, direct, reset in operations:
            # Live UNO ranges at an adjacent insertion boundary may expand.
            # Recreate a precise cursor from the original start for each edit.
            target = owner.createTextCursorByRange(match.range.getStart())
            if index:
                target.goRight(index, False)
            target.goRight(1 if index is not None else len(match.original), True)
            expected = match.original[index] if index is not None else match.original
            if target.getString() != expected:
                raise ValueError("A match changed during replacement. Preview again.")
            target.setString(replacement)
            changed = True
            for name in reset:
                target.setPropertyToDefault(name)
            if direct:
                target.setPropertyValues(tuple(name for name, _ in direct),
                                         tuple(value for _, value in direct))
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
