#!/usr/bin/env python3
"""Fill the UNAL Word formats (constancia / informe) from a JSON operation list.

The templates are edited in place through python-docx so the UNAL logo, headers
and table borders survive untouched. Addressing is by raw grid coordinates
(`tr` / `tc` on the XML rows and cells), because both templates use merged
cells that make python-docx's `row.cells` view collapse.

Usage:
    docx_fill.py --docx FILE.docx --ops ops.json [--out OUT.docx]

ops.json:
    {"ops": [
      {"op": "replace", "old": "[SEDE]", "new": "SEDE MANIZALES"},
      {"op": "set_cell", "scope": "header", "table": 1, "tr": 0, "tc": 1,
       "text": "NOMBRE APELLIDO CONTRATISTA"},
      {"op": "blanks", "scope": "body", "table": 0, "tr": 5, "tc": 0,
       "texts": ["9876543210", "02/09/2026", "SEPTIEMBRE DE 2026"]},
      {"op": "checkbox", "scope": "body", "table": 0, "tr": 6, "tc": 0, "index": 2},
      {"op": "ensure_rows", "scope": "body", "table": 1, "count": 10},
      {"op": "repeat_header", "scope": "body", "table": 1, "rows": 2}
    ]}

Every op is applied in order and reported on stdout as JSON.
"""

from __future__ import annotations

import argparse
import copy
import json
import os
import pathlib
import random
import re
import sys
import tempfile
import zipfile

import docx
from docx.oxml.ns import qn
from lxml import etree

BLANK = re.compile(r"_{3,}")
UNCHECKED = "☐"
CHECKED = "☒"


def tables_for(document, scope: str):
    if scope == "body":
        return document.tables
    section = document.sections[0]
    if scope == "header":
        return section.header.tables
    if scope == "footer":
        return section.footer.tables
    raise ValueError(f"unknown scope {scope!r}")


def paragraphs_for(document, scope: str):
    if scope == "body":
        return document.paragraphs
    section = document.sections[0]
    if scope == "header":
        return section.header.paragraphs
    if scope == "footer":
        return section.footer.paragraphs
    raise ValueError(f"unknown scope {scope!r}")


def flatten_content_controls(document) -> int:
    """Replace every `w:sdt` with its content, everywhere in the document.

    The constancia wraps its modalidad dropdown in a `w:sdt` that sits as a
    sibling of `w:tc` inside the row, which silently shifts every column index,
    and a control left in place keeps printing its own placeholder instead of
    the text written next to it. Flattening removes both problems at once.
    """
    removed = 0
    roots = [document.element.body]
    section = document.sections[0]
    for part in (section.header, section.footer):
        roots.append(part._element)

    for root in roots:
        while True:
            node = root.find(".//" + qn("w:sdt"))
            if node is None:
                break
            parent = node.getparent()
            index = list(parent).index(node)
            content = node.find(qn("w:sdtContent"))
            children = list(content) if content is not None else []
            for offset, child in enumerate(children):
                parent.insert(index + offset, child)
            parent.remove(node)
            removed += 1
    return removed


def grid_cell(table, tr: int, tc: int):
    """Address a cell by its raw XML position, ignoring merge expansion."""
    rows = table._tbl.findall(qn("w:tr"))
    cells = rows[tr].findall(qn("w:tc"))
    return cells[tc]


def cell_paragraphs(tc):
    return tc.findall(qn("w:p"))


def paragraph_text(node) -> str:
    return "".join(t.text or "" for t in node.iter(qn("w:t")))


def set_paragraph_text(node, text: str) -> None:
    """Rewrite the paragraph as a single run, keeping its character formatting.

    Everything except `w:pPr` is dropped, content controls (`w:sdt`) included:
    both templates wrap their dropdowns and ballot boxes in content controls,
    and a control left in place keeps rendering its own placeholder in Word no
    matter what text sits next to it.
    """
    first_run = node.find(".//" + qn("w:r"))
    run_properties = None
    if first_run is not None:
        found = first_run.find(qn("w:rPr"))
        if found is not None:
            run_properties = copy.deepcopy(found)

    for child in list(node):
        if child.tag != qn("w:pPr"):
            node.remove(child)

    run = node.makeelement(qn("w:r"), {})
    if run_properties is not None:
        run.append(run_properties)
    t = run.makeelement(qn("w:t"), {})
    t.set(qn("xml:space"), "preserve")
    t.text = text
    run.append(t)
    node.append(run)


def replace_span(node, start: int, end: int, text: str) -> bool:
    """Replace characters [start, end) of the paragraph's concatenated text.

    Edits the single `w:t` that holds the span whenever possible, so runs,
    fonts and line breaks around it survive. Returns True when that fast path
    applied, False when the paragraph had to be rebuilt.
    """
    offset = 0
    for t in node.iter(qn("w:t")):
        current = t.text or ""
        if offset <= start and end <= offset + len(current):
            t.text = current[: start - offset] + text + current[end - offset :]
            return True
        offset += len(current)
    joined = paragraph_text(node)
    set_paragraph_text(node, joined[:start] + text + joined[end:])
    return False


def replace_in_paragraph(node, old: str, new: str) -> bool:
    """Replace inside a single run when possible, else merge and replace."""
    for t in node.iter(qn("w:t")):
        if old in (t.text or ""):
            t.text = t.text.replace(old, new)
            return True
    joined = paragraph_text(node)
    if old not in joined:
        return False
    set_paragraph_text(node, joined.replace(old, new))
    return True


def replace_everywhere(document, old: str, new: str) -> int:
    hits = 0
    for scope in ("body", "header", "footer"):
        for node in (p._p for p in paragraphs_for(document, scope)):
            hits += replace_in_paragraph(node, old, new)
        for table in tables_for(document, scope):
            for tr in table._tbl.findall(qn("w:tr")):
                for tc in tr.findall(qn("w:tc")):
                    for node in cell_paragraphs(tc):
                        hits += replace_in_paragraph(node, old, new)
    return hits


def set_cell_text(document, scope: str, table: int, tr: int, tc: int, text: str) -> None:
    cell = grid_cell(tables_for(document, scope)[table], tr, tc)
    paragraphs = cell_paragraphs(cell)
    set_paragraph_text(paragraphs[0], text)
    for extra in paragraphs[1:]:
        cell.remove(extra)


def fill_blanks(document, scope: str, table: int, tr: int, tc: int, texts: list) -> int:
    """Fill the runs of underscores inside one cell, left to right.

    `texts` is positional: entry i goes into the i-th blank of the cell, and a
    null entry leaves that blank untouched. Filling happens in one pass because
    replacing a blank removes it, which would shift every later index.
    """
    cell = grid_cell(tables_for(document, scope)[table], tr, tc)
    nodes = cell_paragraphs(cell)
    spans = [
        (node_index, match.span())
        for node_index, node in enumerate(nodes)
        for match in BLANK.finditer(paragraph_text(node))
    ]
    if len(texts) > len(spans):
        raise IndexError(
            f"cell ({tr},{tc}) of table {table} has {len(spans)} blank(s), got {len(texts)} value(s)"
        )
    filled = 0
    # Right to left so earlier spans keep their offsets.
    for index in range(len(texts) - 1, -1, -1):
        if texts[index] is None:
            continue
        node_index, (start, end) = spans[index]
        replace_span(nodes[node_index], start, end, str(texts[index]))
        filled += 1
    return filled


def tick(document, scope: str, table: int, tr: int, tc: int, index: int) -> None:
    """Mark the `index`-th ballot box inside one cell as checked.

    Checked and unchecked boxes are both counted, so indices stay absolute no
    matter how many ticks were already applied to the same cell.
    """
    cell = grid_cell(tables_for(document, scope)[table], tr, tc)
    seen = 0
    for node in cell_paragraphs(cell):
        joined = paragraph_text(node)
        positions = [i for i, char in enumerate(joined) if char in (UNCHECKED, CHECKED)]
        if seen + len(positions) <= index:
            seen += len(positions)
            continue
        at = positions[index - seen]
        replace_span(node, at, at + 1, CHECKED)
        return
    raise IndexError(f"checkbox #{index} not found in table {table} cell ({tr},{tc})")


# w14:paraId (and, by convention, w14:textId) must be 8 uppercase hex digits,
# never 00000000, and strictly below 80000000 per the OOXML wordml schema.
W14_ID_MIN = 1
W14_ID_MAX = 0x7FFFFFFF


def _collect_w14_ids(document) -> set:
    """Every `w14:paraId`/`w14:textId` value already used in the document body."""
    ids = set()
    for element in document.element.body.iter():
        for attr in (qn("w14:paraId"), qn("w14:textId")):
            value = element.get(attr)
            if value:
                ids.add(value)
    return ids


def _new_w14_id(used: set) -> str:
    """A fresh id in the valid range, guaranteed distinct from every id in `used`."""
    while True:
        candidate = f"{random.randint(W14_ID_MIN, W14_ID_MAX):08X}"
        if candidate not in used:
            used.add(candidate)
            return candidate


def _regenerate_w14_ids(clone, used: set) -> None:
    """Give a cloned `w:tr`/`w:p` subtree fresh paraId/textId values.

    `copy.deepcopy` copies these identifiers verbatim, so every row added by
    `ensure_rows` would otherwise share the template row's paraId/textId --
    something native Word never produces. Covers the row itself plus any
    nested `w:tr`/`w:p` descendants that carry the attributes.
    """
    for element in list(clone.iter(qn("w:tr"))) + list(clone.iter(qn("w:p"))):
        if element.get(qn("w14:paraId")) is not None:
            element.set(qn("w14:paraId"), _new_w14_id(used))
        if element.get(qn("w14:textId")) is not None:
            element.set(qn("w14:textId"), _new_w14_id(used))


def ensure_rows(document, scope: str, table: int, count: int, template_row: int = -1) -> int:
    """Grow a table to `count` data rows by cloning `template_row`."""
    tbl = tables_for(document, scope)[table]._tbl
    rows = tbl.findall(qn("w:tr"))
    source = rows[template_row]
    used_ids = _collect_w14_ids(document)
    added = 0
    while len(tbl.findall(qn("w:tr"))) < count:
        clone = copy.deepcopy(source)
        _regenerate_w14_ids(clone, used_ids)
        for tc in clone.findall(qn("w:tc")):
            for node in cell_paragraphs(tc):
                set_paragraph_text(node, "")
        source.addnext(clone)
        source = clone
        added += 1
    return added


# Successors of `w:tblHeader` inside `w:trPr`, per the OOXML `CT_TrPrBase`
# sequence (cnfStyle, divId, gridBefore, gridAfter, wBefore, wAfter, cantSplit,
# trHeight, tblHeader, tblCellSpacing, jc, hidden, ins, del, trPrChange).
_TBLHEADER_SUCCESSORS = ("w:tblCellSpacing", "w:jc", "w:hidden", "w:ins", "w:del", "w:trPrChange")


def repeat_header(document, scope: str, table: int, rows: int) -> int:
    """Mark the first `rows` rows of a table as repeating header rows.

    Word only repeats a row across page breaks when its `w:trPr` carries
    `<w:tblHeader/>`; without it, a table that crosses a page shows its data
    rows with no column titles above them. Idempotent: a row that already
    carries the flag is left untouched.
    """
    tbl = tables_for(document, scope)[table]._tbl
    trs = tbl.findall(qn("w:tr"))[:rows]
    marked = 0
    for tr in trs:
        trPr = tr.get_or_add_trPr()
        if trPr.find(qn("w:tblHeader")) is not None:
            continue
        tblHeader = trPr.makeelement(qn("w:tblHeader"), {})
        trPr.insert_element_before(tblHeader, *_TBLHEADER_SUCCESSORS)
        marked += 1
    return marked


# python-docx re-serializes every zip part it parses through lxml, even one
# nobody touched: the XML declaration's quotes flip from `"` to `'` and the
# byte count shifts by one. UNAL's intake rejected a submission over exactly
# that -- a diff on parts that carry the format itself (styles, numbering,
# settings) with no actual content change. The functions below rebuild the
# saved .docx so only genuinely-changed parts differ from the source file.

_XML_SUFFIXES = (".xml", ".rels")


def _read_zip_parts(path) -> tuple[list, dict, dict]:
    """Return (name order, {name: ZipInfo}, {name: raw bytes}) for a .docx."""
    with zipfile.ZipFile(path) as zf:
        order = zf.namelist()
        infos = {name: zf.getinfo(name) for name in order}
        blobs = {name: zf.read(name) for name in order}
    return order, infos, blobs


def _clone_zipinfo(info: zipfile.ZipInfo) -> zipfile.ZipInfo:
    """Copy every metadata field of a ZipInfo, so the entry re-writes identically."""
    clone = zipfile.ZipInfo(info.filename, date_time=info.date_time)
    clone.compress_type = info.compress_type
    clone.external_attr = info.external_attr
    clone.internal_attr = info.internal_attr
    clone.create_system = info.create_system
    clone.create_version = info.create_version
    clone.extract_version = info.extract_version
    clone.flag_bits = info.flag_bits
    clone.volume = info.volume
    clone.comment = info.comment
    clone.extra = info.extra
    return clone


def _semantically_equal_xml(a: bytes, b: bytes) -> bool:
    """True when two XML blobs parse to the same tree, ignoring serialization."""
    try:
        return etree.canonicalize(a.decode("utf-8")) == etree.canonicalize(b.decode("utf-8"))
    except (etree.XMLSyntaxError, UnicodeDecodeError):
        return False


def _content_types_equal(a: bytes, b: bytes) -> bool:
    """Order-insensitive equality for `[Content_Types].xml`.

    Per the OPC spec, `Default`/`Override` declarations are an unordered set:
    nothing depends on which one comes first. python-docx rebuilds this part
    from an internal dict and serializes it in a different (alphabetical)
    order than the template's, with no semantic effect -- but plain
    tree-order canonicalization treats reordered siblings as a real change,
    which would wrongly keep this part off the byte-identical list. Compare
    it as a multiset of declarations instead.
    """
    try:
        root_a, root_b = etree.fromstring(a), etree.fromstring(b)
    except etree.XMLSyntaxError:
        return False
    if root_a.tag != root_b.tag or dict(root_a.attrib) != dict(root_b.attrib):
        return False

    def declarations(root):
        return sorted((child.tag, tuple(sorted(child.attrib.items()))) for child in root)

    return declarations(root_a) == declarations(root_b)


def _xml_parts_equal(name: str, a: bytes, b: bytes) -> bool:
    if name == "[Content_Types].xml":
        return _content_types_equal(a, b)
    return _semantically_equal_xml(a, b)


def rebuild_docx_preserving_parts(orig_order, orig_infos, orig_blobs, saved_path, final_path) -> dict:
    """Merge python-docx's fresh save with the caller's original zip parts.

    `orig_*` describe the source file, read into memory before python-docx
    touched anything. `saved_path` is what `document.save()` just produced.
    For every part: keep the ORIGINAL bytes (and ZipInfo, so compression and
    metadata match too) when the new bytes are identical or -- for XML parts
    -- only differ in serialization; otherwise keep the new bytes. Entry
    order follows the original zip, with any part the save introduced
    appended at the end, so no incidental zip-structure diff is introduced.
    """
    new_order, new_infos, new_blobs = _read_zip_parts(saved_path)

    final_order = list(orig_order)
    final_order += [name for name in new_order if name not in orig_infos]

    preserved, changed, added, dropped = [], [], [], []

    # Build into a scratch file next to the destination and swap it in only
    # once it is complete: `final_path` may equal the caller's original
    # `--docx` path, and a mid-loop failure must never leave that file
    # truncated or holding a half-written zip.
    final_path = pathlib.Path(final_path)
    fd, scratch_name = tempfile.mkstemp(
        prefix=".docx_fill-rebuild-", suffix=".docx", dir=str(final_path.parent)
    )
    os.close(fd)
    scratch_path = pathlib.Path(scratch_name)
    try:
        with zipfile.ZipFile(scratch_path, "w") as out:
            for name in final_order:
                orig_blob = orig_blobs.get(name)
                new_blob = new_blobs.get(name)

                if orig_blob is None:
                    out.writestr(_clone_zipinfo(new_infos[name]), new_blob)
                    added.append(name)
                    continue
                if new_blob is None:
                    dropped.append(name)
                    continue

                keep_original = orig_blob == new_blob
                if not keep_original and name.lower().endswith(_XML_SUFFIXES):
                    keep_original = _xml_parts_equal(name, orig_blob, new_blob)

                if keep_original:
                    out.writestr(_clone_zipinfo(orig_infos[name]), orig_blob)
                    preserved.append(name)
                else:
                    out.writestr(_clone_zipinfo(new_infos[name]), new_blob)
                    changed.append(name)
        os.replace(scratch_path, final_path)
    except BaseException:
        scratch_path.unlink(missing_ok=True)
        raise

    return {"preserved": preserved, "changed": changed, "added": added, "dropped": dropped}


def apply(document, op: dict) -> str:
    kind = op["op"]
    if kind == "replace":
        hits = replace_everywhere(document, op["old"], op["new"])
        return f"replace {op['old']!r} -> {hits} hit(s)"
    if kind == "set_cell":
        set_cell_text(document, op.get("scope", "body"), op["table"], op["tr"], op["tc"], op["text"])
        return f"set_cell {op.get('scope','body')}:{op['table']}({op['tr']},{op['tc']})"
    if kind == "blanks":
        filled = fill_blanks(
            document, op.get("scope", "body"), op["table"], op["tr"], op["tc"], op["texts"]
        )
        return f"blanks {op['table']}({op['tr']},{op['tc']}) -> {filled} filled"
    if kind == "checkbox":
        tick(document, op.get("scope", "body"), op["table"], op["tr"], op["tc"], op["index"])
        return f"checkbox #{op['index']} in {op['table']}({op['tr']},{op['tc']})"
    if kind == "ensure_rows":
        added = ensure_rows(
            document, op.get("scope", "body"), op["table"], op["count"], op.get("template_row", -1)
        )
        return f"ensure_rows table {op['table']} -> +{added}"
    if kind == "repeat_header":
        marked = repeat_header(document, op.get("scope", "body"), op["table"], op["rows"])
        return f"repeat_header table {op['table']} -> {marked} row(s) marked"
    raise ValueError(f"unknown op {kind!r}")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--docx", required=True, type=pathlib.Path)
    parser.add_argument("--ops", required=True, type=pathlib.Path)
    parser.add_argument("--out", type=pathlib.Path)
    args = parser.parse_args()

    # Read the source zip into memory before python-docx opens (and, for an
    # in-place run, later overwrites) it, so the untouched parts can be
    # restored verbatim afterwards -- this also keeps a chained run (ops
    # applied on an already-diligenced .docx) preserving whatever the
    # previous run already inherited from the original template.
    orig_order, orig_infos, orig_blobs = _read_zip_parts(args.docx)

    document = docx.Document(str(args.docx))
    flattened = flatten_content_controls(document)
    ops = json.loads(args.ops.read_text(encoding="utf-8"))["ops"]

    log = []
    for index, op in enumerate(ops):
        try:
            log.append({"i": index, "ok": True, "detail": apply(document, op)})
        except Exception as error:  # surface the failing op, keep the rest legible
            log.append({"i": index, "ok": False, "detail": f"{type(error).__name__}: {error}"})

    out = (args.out or args.docx).resolve()

    # Save to a scratch file first: `out` may be the same path as `args.docx`,
    # and rebuilding needs both the pristine original bytes (already in
    # memory) and python-docx's fresh output at once.
    fd, tmp_name = tempfile.mkstemp(prefix=".docx_fill-", suffix=".docx", dir=str(out.parent))
    os.close(fd)
    tmp_path = pathlib.Path(tmp_name)
    try:
        document.save(str(tmp_path))
        parts = rebuild_docx_preserving_parts(orig_order, orig_infos, orig_blobs, tmp_path, out)
    finally:
        tmp_path.unlink(missing_ok=True)

    failed = [entry for entry in log if not entry["ok"]]
    print(json.dumps({"docx": str(out), "content_controls_flattened": flattened,
                      "applied": log, "failed": len(failed),
                      "parts_preserved": {"count": len(parts["preserved"]), "parts": parts["preserved"]},
                      "parts_changed": {"count": len(parts["changed"]), "parts": parts["changed"]},
                      "parts_added": parts["added"], "parts_dropped": parts["dropped"]},
                     ensure_ascii=False, indent=2))
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
