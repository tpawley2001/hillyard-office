# SPDX-License-Identifier: MPL-2.0
"""What a document contains, so a saved copy can be checked for losses.

inventory(document) walks a loaded Writer, Calc or Impress document through
the UNO API and returns plain data: counts, names and shape geometry. compare()
takes two inventories (the open document and the file reloaded after saving)
and lists what changed, worded for a person rather than a developer.

This is the check Word's Compatibility Checker stands for, done the PDF Forge
way: instead of predicting losses from a list of known unsupported features,
save, reload, and compare what is actually there.

compare() and describe() import no UNO modules and are unit-tested without an
office.
"""
from collections import Counter

# Shape geometry is compared in 1/100 mm; Word stores EMUs and twips, and a
# round trip through either can move an object by a rounding step.
GEOMETRY_TOLERANCE = 50          # 0.5 mm


def _names(container):
    try:
        return sorted(container.getElementNames())
    except Exception:
        return []


def _count(container):
    try:
        return container.getCount()
    except Exception:
        return 0


def _column_counts(sections):
    counts = []
    for name in _names(sections):
        try:
            columns = sections.getByName(name).getPropertyValue("TextColumns").ColumnCount
        except Exception:
            continue
        if columns > 1:
            counts.append(columns)
    return counts


def _shape_kind(shape):
    kind = shape.getShapeType().rsplit(".", 1)[-1]
    if kind == "CustomShape":
        try:
            if shape.getPropertyValue("TextBox"):
                return "TextBox"
        except Exception:
            pass
    return kind


def _shapes(draw_page):
    shapes = []
    for i in range(_count(draw_page)):
        shape = draw_page.getByIndex(i)
        entry = {"kind": _shape_kind(shape)}
        try:
            pos, size = shape.getPosition(), shape.getSize()
            entry.update(x=pos.X, y=pos.Y, width=size.Width, height=size.Height)
        except Exception:
            pass
        for name, key in (("AnchorType", "anchor"), ("TextWrap", "wrap")):
            try:
                entry[key] = shape.getPropertyValue(name).value
            except Exception:
                pass
        try:
            entry["text"] = shape.getString()[:200]
        except Exception:
            pass
        shapes.append(entry)
    return shapes


def _writer(doc):
    # The document's own counts are computed from its text. DocumentProperties'
    # DocumentStatistics are only the numbers stored in the file (for a DOCX,
    # Word's docProps/app.xml), so they cannot show a loss.
    counts = {name: doc.getPropertyValue(name)
              for name in ("ParagraphCount", "WordCount", "CharacterCount")}
    # The view's PageCount lays the document out without moving the user's
    # cursor, which matters when the open document is the one being checked.
    pages = doc.CurrentController.getPropertyValue("PageCount")
    fields = Counter()
    comments = []
    form_controls = 0
    enum = doc.TextFields.createEnumeration()
    while enum.hasMoreElements():
        field = enum.nextElement()
        services = field.getSupportedServiceNames()
        kind = next((s.rsplit(".", 1)[-1] for s in services
                     if s.startswith("com.sun.star.text.TextField.")
                     or s.startswith("com.sun.star.text.textfield.")), "Field")
        if kind == "Annotation":
            try:
                comments.append(field.getPropertyValue("Content")[:200])
            except Exception:
                comments.append("")
        elif kind == "DropDown" or "com.sun.star.text.Fieldmark" in services:
            # Word form controls: legacy form fields and date pickers come in
            # as fieldmarks or fields and may come back as content controls;
            # to Word they are the same control, so they are counted together.
            form_controls += 1
        else:
            fields[kind] += 1
    content_controls = 0
    hyperlinks = 0
    body_characters = 0
    para_enum = doc.Text.createEnumeration()
    while para_enum.hasMoreElements():
        para = para_enum.nextElement()
        if para.supportsService("com.sun.star.text.TextTable"):
            for name in para.getCellNames():
                body_characters += len("".join(para.getCellByName(name).getString().split()))
            continue
        if not para.supportsService("com.sun.star.text.Paragraph"):
            continue
        body_characters += len("".join(para.getString().split()))
        portions = para.createEnumeration()
        previous_url = ""
        while portions.hasMoreElements():
            portion = portions.nextElement()
            try:
                if portion.getPropertyValue("TextPortionType") == "ContentControl":
                    content_controls += 1
                url = portion.getPropertyValue("HyperLinkURL")
                # A link split over several formatting runs is one link.
                if url and url != previous_url:
                    hyperlinks += 1
                previous_url = url
            except Exception:
                previous_url = ""
    headers = footers = 0
    styles = doc.StyleFamilies.getByName("PageStyles")
    for name in styles.getElementNames():
        style = styles.getByName(name)
        if not style.isInUse():
            continue
        headers += bool(style.getPropertyValue("HeaderIsOn"))
        footers += bool(style.getPropertyValue("FooterIsOn"))
    redlines = 0
    try:
        redlines = _count(doc.Redlines)
    except Exception:
        pass
    return {
        "application": "writer", "pages": pages,
        "paragraphs": counts["ParagraphCount"], "words": counts["WordCount"],
        "characters": counts["CharacterCount"],
        "body_characters": body_characters,
        "tables": _names(doc.TextTables), "frames": _names(doc.TextFrames),
        "images": _count(doc.GraphicObjects), "objects": _count(doc.EmbeddedObjects),
        "sections": _names(doc.TextSections),
        # Writer represents continuous section breaks as text sections, and how
        # many it makes can change on a round trip with no change in the file;
        # what matters is which parts of the text are set in columns.
        "column_sections": sorted(_column_counts(doc.TextSections)),
        "bookmarks": _names(doc.Bookmarks),
        "footnotes": _count(doc.Footnotes), "endnotes": _count(doc.Endnotes),
        "indexes": _count(doc.DocumentIndexes), "fields": dict(fields),
        "comments": comments, "tracked_changes": redlines,
        "content_controls": content_controls + form_controls, "hyperlinks": hyperlinks,
        "headers": headers, "footers": footers,
        "shapes": _shapes(doc.DrawPage),
    }


def _calc(doc):
    sheets = []
    for i in range(doc.Sheets.getCount()):
        sheet = doc.Sheets.getByIndex(i)
        cursor = sheet.createCursor()
        cursor.gotoEndOfUsedArea(False)
        end = cursor.getRangeAddress()
        used = sheet.getCellRangeByPosition(0, 0, end.EndColumn, end.EndRow)
        formulas = 0
        try:
            from com.sun.star.sheet.CellFlags import FORMULA
            ranges = used.queryContentCells(FORMULA)
            for j in range(ranges.getCount()):
                a = ranges.getByIndex(j).getRangeAddress()
                formulas += (a.EndColumn - a.StartColumn + 1) * (a.EndRow - a.StartRow + 1)
        except Exception:
            pass
        conditional = 0
        try:
            conditional = _count(sheet.ConditionalFormats)
        except Exception:
            pass
        charts = _count(sheet.getCharts())
        sheets.append({
            "name": sheet.getName(), "formulas": formulas,
            "used_range": [end.EndColumn + 1, end.EndRow + 1],
            "comments": _count(sheet.getAnnotations()),
            "charts": charts, "pivot_tables": _count(sheet.getDataPilotTables()),
            "conditional_formats": conditional,
            "merged_ranges": _merged(used), "shapes": _shapes(sheet.DrawPage),
            "protected": bool(sheet.isProtected()),
            "visible": bool(sheet.getPropertyValue("IsVisible")),
        })
    return {"application": "calc", "sheets": sheets,
            "named_ranges": _names(doc.NamedRanges),
            "database_ranges": _names(doc.DatabaseRanges)}


def _merged(cell_range):
    """Merged areas in a range. Calc has no direct query, but a merge is a cell
    attribute, so the unique-format ranges split at every merge boundary."""
    try:
        merged = 0
        ranges = cell_range.getUniqueCellFormatRanges()
        for i in range(ranges.getCount()):
            block = ranges.getByIndex(i)
            for j in range(block.getCount()):
                if block.getByIndex(j).getIsMerged():
                    merged += 1
        return merged
    except Exception:
        return None


def _impress(doc):
    slides = []
    for i in range(doc.DrawPages.getCount()):
        page = doc.DrawPages.getByIndex(i)
        notes = ""
        try:
            notes_page = page.getNotesPage()
            notes = " ".join(notes_page.getByIndex(j).getString()
                             for j in range(notes_page.getCount())
                             if notes_page.getByIndex(j).getShapeType().endswith("NotesShape"))
        except Exception:
            pass
        transition = None
        for name in ("TransitionType", "Effect"):
            try:
                value = page.getPropertyValue(name)
                transition = getattr(value, "value", value)
                break
            except Exception:
                pass
        hidden = False
        try:
            hidden = not page.getPropertyValue("Visible")
        except Exception:
            pass
        slides.append({"shapes": _shapes(page), "notes": notes[:400],
                       "transition": transition, "hidden": hidden,
                       "master": page.MasterPage.getName()})
    return {"application": "impress", "slides": slides,
            "masters": _count(doc.MasterPages)}


def inventory(doc):
    if doc.supportsService("com.sun.star.text.TextDocument"):
        return _writer(doc)
    if doc.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
        return _calc(doc)
    if doc.supportsService("com.sun.star.presentation.PresentationDocument"):
        return _impress(doc)
    raise ValueError("Only Writer, Calc and Impress documents can be checked.")


# --- comparison (no UNO below this line) ---------------------------------

# frames, images and objects are not compared on their own: every one of them
# is also a shape on the draw page, and _compare_shapes covers it by category.
WRITER_LABELS = {
    "tables": "tables", "column_sections": "parts set in columns", "bookmarks": "bookmarks",
    "footnotes": "footnotes", "endnotes": "endnotes", "indexes": "tables of contents and indexes",
    "comments": "comments", "tracked_changes": "tracked changes",
    "content_controls": "form controls", "hyperlinks": "hyperlinks",
    "headers": "page styles with a header", "footers": "page styles with a footer",
}


def _size(value):
    return len(value) if isinstance(value, (list, tuple, dict)) else (value or 0)


def _finding(severity, area, message, before=None, after=None):
    return {"severity": severity, "area": area, "message": message,
            "before": before, "after": after}


# Shape types that store the same kind of object. An ellipse may come back as
# a custom shape with ellipse geometry and look the same, so only a change of
# category is a finding.
SHAPE_CATEGORIES = {
    "FrameShape": "frame", "GraphicObjectShape": "picture", "TextBox": "text box",
    "GroupShape": "group", "OLE2Shape": "embedded object", "LineShape": "line",
    "ConnectorShape": "line", "MediaShape": "media",
}


def _category(shape):
    return SHAPE_CATEGORIES.get(shape["kind"], "drawing shape")


def _compare_shapes(before, after, where):
    findings = []
    kinds_before = Counter(_category(s) for s in before)
    kinds_after = Counter(_category(s) for s in after)
    for kind in sorted(set(kinds_before) | set(kinds_after)):
        b, a = kinds_before[kind], kinds_after[kind]
        if a < b:
            findings.append(_finding("loss", "shapes",
                                     f"{where}{b - a} of {b} {kind}(s) would be lost", b, a))
        elif a > b:
            findings.append(_finding("change", "shapes",
                                     f"{where}{a - b} extra {kind}(s) would appear", b, a))
    if kinds_before != kinds_after:
        return findings
    moved = 0
    reanchored = 0
    lost_text = []
    for s, t in zip(sorted(before, key=_category), sorted(after, key=_category)):
        if any(abs(s.get(k, 0) - t.get(k, 0)) > GEOMETRY_TOLERANCE
               for k in ("x", "y", "width", "height")):
            moved += 1
        if s.get("anchor") is not None and s.get("anchor") != t.get("anchor"):
            reanchored += 1
        if s.get("text") and not t.get("text"):
            lost_text.append(s["text"])
    if moved:
        findings.append(_finding("change", "layout",
                                 f"{where}{moved} object(s) moved or changed size", None, moved))
    if reanchored:
        findings.append(_finding("change", "layout",
                                 f"{where}{reanchored} object(s) changed their anchor", None, reanchored))
    if lost_text:
        findings.append(_finding("loss", "text",
                                 f"{where}text inside {len(lost_text)} object(s) is gone",
                                 lost_text[:3], None))
    return findings


def _compare_writer(before, after):
    findings = []
    for key, label in WRITER_LABELS.items():
        b, a = _size(before.get(key)), _size(after.get(key))
        if a < b:
            findings.append(_finding("loss", key, f"{b - a} of {b} {label} would be lost", b, a))
        elif a > b:
            findings.append(_finding("change", key, f"{a - b} extra {label} would appear", b, a))
    fb, fa = before.get("fields", {}), after.get("fields", {})
    for kind in sorted(set(fb) | set(fa)):
        b, a = fb.get(kind, 0), fa.get(kind, 0)
        if a < b:
            findings.append(_finding("loss", "fields",
                                     f"{b - a} of {b} {kind} field(s) would become plain text or vanish", b, a))
    if before.get("characters") and after.get("characters") is not None:
        b, a = before["characters"], after["characters"]
        if a < b:
            findings.append(_finding("loss", "text", f"{b - a} characters of text would be lost", b, a))
    if before.get("pages") != after.get("pages"):
        findings.append(_finding("change", "layout",
                                 f"the document would have {after.get('pages')} pages instead of "
                                 f"{before.get('pages')}", before.get("pages"), after.get("pages")))
    findings += _compare_shapes(before.get("shapes", []), after.get("shapes", []), "")
    return findings


def _compare_calc(before, after):
    findings = []
    nb, na = before["named_ranges"], after["named_ranges"]
    lost = sorted(set(nb) - set(na))
    if lost:
        findings.append(_finding("loss", "named_ranges",
                                 f"{len(lost)} named range(s) would be lost: {', '.join(lost[:5])}",
                                 len(nb), len(na)))
    after_sheets = {s["name"]: s for s in after["sheets"]}
    for sheet in before["sheets"]:
        other = after_sheets.get(sheet["name"])
        where = f"Sheet {sheet['name']}: "
        if other is None:
            findings.append(_finding("loss", "sheets", f"{where}the sheet would be lost"))
            continue
        for key, label in (("formulas", "formulas"), ("comments", "comments"),
                           ("charts", "charts"), ("pivot_tables", "pivot tables"),
                           ("conditional_formats", "conditional formats"),
                           ("merged_ranges", "merged cell ranges")):
            b, a = sheet.get(key) or 0, other.get(key) or 0
            if a < b:
                findings.append(_finding("loss", key, f"{where}{b - a} of {b} {label} would be lost", b, a))
        if sheet["protected"] and not other["protected"]:
            findings.append(_finding("loss", "protection", f"{where}sheet protection would be removed"))
        findings += _compare_shapes(sheet["shapes"], other["shapes"], where)
    return findings


def _compare_impress(before, after):
    findings = []
    b, a = len(before["slides"]), len(after["slides"])
    if a != b:
        findings.append(_finding("loss" if a < b else "change", "slides",
                                 f"the presentation would have {a} slides instead of {b}", b, a))
    for n, (s, t) in enumerate(zip(before["slides"], after["slides"]), 1):
        where = f"Slide {n}: "
        findings += _compare_shapes(s["shapes"], t["shapes"], where)
        if s["notes"].strip() and not t["notes"].strip():
            findings.append(_finding("loss", "notes", f"{where}speaker notes would be lost"))
        if s["transition"] and not t["transition"]:
            findings.append(_finding("loss", "transitions", f"{where}the slide transition would be lost"))
        if s["hidden"] != t["hidden"]:
            findings.append(_finding("change", "slides", f"{where}hidden/shown state would change"))
    return findings


def compare(before, after):
    """Findings, most serious first: each {severity, area, message, before, after}."""
    if before["application"] != after["application"]:
        raise ValueError("The two inventories are of different kinds of document.")
    findings = {"writer": _compare_writer, "calc": _compare_calc,
                "impress": _compare_impress}[before["application"]](before, after)
    order = {"loss": 0, "change": 1}
    return sorted(findings, key=lambda f: order[f["severity"]])


def describe(findings, format_name):
    if not findings:
        return f"Nothing was lost saving as {format_name}."
    losses = sum(f["severity"] == "loss" for f in findings)
    head = (f"Saving as {format_name} loses or changes {len(findings)} thing(s)"
            + (f", {losses} of them losses" if losses else "") + ":")
    return "\n".join([head] + [f"  - {f['message']}" for f in findings])
