#!/usr/bin/env python3
"""Validate XLSX package diffs against strict topology invariants."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import xml.etree.ElementTree as ET
from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path
from collections.abc import Mapping
from typing import cast
from zipfile import ZipFile


XML_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
XML_NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CELL_TAG = f"{{{XML_NS_MAIN}}}c"
CELL_VALUE_TAG = f"{{{XML_NS_MAIN}}}v"
CELL_INLINE_TAG = f"{{{XML_NS_MAIN}}}is"


@dataclass(frozen=True)
class WorkbookSnapshot:
    members: frozenset[str]
    member_sha256: dict[str, str]
    sheet_order: list[str]
    sheet_states: dict[str, str]
    defined_names: list[tuple[str, str | None, str]]
    sheet_xml_paths: dict[str, str]
    merges_by_sheet: dict[str, tuple[str, ...]]
    media_count: int
    drawing_count: int


def _read_json(path: Path) -> dict[str, object]:
    return cast(dict[str, object], json.loads(path.read_text(encoding="utf-8")))


def _normalize_xl_path(target: str) -> str:
    normalized = target.lstrip("/")
    if normalized.startswith("xl/"):
        return normalized
    return f"xl/{normalized}"


def _sheet_xml_paths(zf: ZipFile, workbook_root: ET.Element) -> dict[str, str]:
    rels_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_map = {
        rel.attrib["Id"]: _normalize_xl_path(rel.attrib["Target"])
        for rel in rels_root
        if rel.attrib.get("Id") and rel.attrib.get("Target")
    }
    paths: dict[str, str] = {}
    for sheet in workbook_root.findall(f".//{{{XML_NS_MAIN}}}sheet"):
        sheet_name = sheet.attrib.get("name")
        rel_id = sheet.attrib.get(f"{{{XML_NS_REL}}}id")
        if sheet_name and rel_id and rel_id in rel_map:
            paths[sheet_name] = rel_map[rel_id]
    return paths


def _defined_names(workbook_root: ET.Element) -> list[tuple[str, str | None, str]]:
    names: list[tuple[str, str | None, str]] = []
    for defined in workbook_root.findall(f".//{{{XML_NS_MAIN}}}definedName"):
        names.append(
            (
                defined.attrib.get("name", ""),
                defined.attrib.get("localSheetId"),
                (defined.text or "").strip(),
            )
        )
    return sorted(names)


def _sheet_merges(
    zf: ZipFile, sheet_paths: dict[str, str]
) -> dict[str, tuple[str, ...]]:
    merges: dict[str, tuple[str, ...]] = {}
    for sheet_name, sheet_path in sheet_paths.items():
        root = ET.fromstring(zf.read(sheet_path))
        refs = sorted(
            merge.attrib["ref"]
            for merge in root.findall(f".//{{{XML_NS_MAIN}}}mergeCell")
            if merge.attrib.get("ref")
        )
        merges[sheet_name] = tuple(refs)
    return merges


def _sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def snapshot_workbook(path: Path) -> WorkbookSnapshot:
    with ZipFile(path) as zf:
        members = frozenset(zf.namelist())
        member_sha256 = {name: _sha256_bytes(zf.read(name)) for name in members}

        workbook_root = ET.fromstring(zf.read("xl/workbook.xml"))
        sheets = workbook_root.findall(f".//{{{XML_NS_MAIN}}}sheet")
        sheet_order = [sheet.attrib.get("name", "") for sheet in sheets]
        sheet_states = {
            sheet.attrib.get("name", ""): sheet.attrib.get("state", "visible")
            for sheet in sheets
            if sheet.attrib.get("name")
        }

        sheet_paths = _sheet_xml_paths(zf, workbook_root)
        merges = _sheet_merges(zf, sheet_paths)

        media_count = sum(1 for name in members if name.startswith("xl/media/"))
        drawing_count = sum(1 for name in members if name.startswith("xl/drawings/"))

    return WorkbookSnapshot(
        members=members,
        member_sha256=member_sha256,
        sheet_order=sheet_order,
        sheet_states=sheet_states,
        defined_names=_defined_names(workbook_root),
        sheet_xml_paths=sheet_paths,
        merges_by_sheet=merges,
        media_count=media_count,
        drawing_count=drawing_count,
    )


def _compile_allowlist(
    template: WorkbookSnapshot, allowlist: Mapping[str, object]
) -> tuple[set[str], list[re.Pattern[str]], list[str], dict[str, frozenset[str]]]:
    allowed_parts = {
        str(item)
        for item in cast(list[object], allowlist.get("allowed_changed_parts", []))
    }
    allowed_prefixes = [
        str(item)
        for item in cast(
            list[object], allowlist.get("allowed_changed_part_prefixes", [])
        )
    ]
    writable_cells_by_part: dict[str, frozenset[str]] = {}

    if allowlist.get("include_contract_writable_surface_sheet_xml"):
        contract_path = Path(
            str(
                allowlist.get(
                    "contract_path", ".sisyphus/contracts/xlsx-template-contract.json"
                )
            )
        )
        contract = _read_json(contract_path)
        writable_surface = cast(dict[str, object], contract.get("writable_surface", {}))
        primary_target = cast(
            dict[str, object], writable_surface.get("primary_target", {})
        )
        linked_card_targets = cast(
            dict[str, object], writable_surface.get("linked_card_targets", {})
        )
        writable_cells_by_sheet = _expand_contract_writable_surface(
            primary_target=primary_target,
            linked_card_targets=linked_card_targets,
        )
        for sheet_name, allowed_cells in writable_cells_by_sheet.items():
            if sheet_name not in template.sheet_xml_paths:
                raise ValueError(
                    f"allowlist references writable sheet missing in template: {sheet_name}"
                )
            part_name = template.sheet_xml_paths[sheet_name]
            allowed_parts.add(part_name)
            writable_cells_by_part[part_name] = frozenset(allowed_cells)

    for sheet_name in cast(
        list[object], allowlist.get("allowed_changed_sheet_names", [])
    ):
        sheet_name_text = str(sheet_name)
        if sheet_name_text not in template.sheet_xml_paths:
            raise ValueError(f"allowlist references unknown sheet name: {sheet_name}")
        allowed_parts.add(template.sheet_xml_paths[sheet_name_text])

    patterns = [
        re.compile(str(pattern))
        for pattern in cast(
            list[object], allowlist.get("allowed_changed_part_patterns", [])
        )
    ]

    return allowed_parts, patterns, allowed_prefixes, writable_cells_by_part


def _expand_contract_writable_surface(
    *,
    primary_target: Mapping[str, object],
    linked_card_targets: Mapping[str, object],
) -> dict[str, set[str]]:
    writable_cells: dict[str, set[str]] = {}
    primary_sheet_name = str(primary_target["sheet_name"])
    writable_cells.setdefault(primary_sheet_name, set()).update(
        _expand_ranges(cast(list[object], primary_target.get("allowed_ranges", [])))
    )
    for sheet_name, cell_refs in linked_card_targets.items():
        writable_cells.setdefault(str(sheet_name), set()).update(
            str(cell_ref) for cell_ref in cast(list[object], cell_refs)
        )
    return writable_cells


def _expand_ranges(ranges: list[object]) -> set[str]:
    expanded: set[str] = set()
    for range_item in ranges:
        start_ref, end_ref = str(range_item).split(":", maxsplit=1)
        start_col, start_row = _split_cell_ref(start_ref)
        end_col, end_row = _split_cell_ref(end_ref)
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                expanded.add(_cell_ref(col, row))
    return expanded


def _split_cell_ref(cell_ref: str) -> tuple[int, int]:
    match = re.fullmatch(r"([A-Z]+)(\d+)", cell_ref)
    if match is None:
        raise ValueError(f"invalid cell reference in contract allowlist: {cell_ref}")
    return _column_index(match.group(1)), int(match.group(2))


def _column_index(column_letters: str) -> int:
    result = 0
    for char in column_letters:
        result = result * 26 + (ord(char) - ord("A") + 1)
    return result


def _cell_ref(column_index: int, row_index: int) -> str:
    return f"{_column_letters(column_index)}{row_index}"


def _column_letters(column_index: int) -> str:
    chars: list[str] = []
    current = column_index
    while current > 0:
        current, remainder = divmod(current - 1, 26)
        chars.append(chr(ord("A") + remainder))
    return "".join(reversed(chars))


def _is_allowlisted_part(
    part_name: str,
    *,
    allowed_parts: set[str],
    allowed_patterns: list[re.Pattern[str]],
    allowed_prefixes: list[str],
) -> bool:
    if part_name in allowed_parts:
        return True
    if any(part_name.startswith(prefix) for prefix in allowed_prefixes):
        return True
    return any(pattern.fullmatch(part_name) for pattern in allowed_patterns)


def validate_package_diff(
    template_path: Path, output_path: Path, allowlist_path: Path
) -> list[str]:
    template = snapshot_workbook(template_path)
    output = snapshot_workbook(output_path)
    allowlist = _read_json(allowlist_path)
    (
        allowed_parts,
        allowed_patterns,
        allowed_prefixes,
        writable_cells_by_part,
    ) = _compile_allowlist(template, allowlist)

    errors: list[str] = []

    unexpected_added = sorted(output.members - template.members)
    unexpected_removed = sorted(template.members - output.members)
    if unexpected_added:
        errors.append(f"unexpected added package parts: {unexpected_added}")
    if unexpected_removed:
        errors.append(f"unexpected removed package parts: {unexpected_removed}")

    changed_parts = sorted(
        name
        for name in template.members & output.members
        if template.member_sha256[name] != output.member_sha256[name]
    )
    unexpected_changed = [
        name
        for name in changed_parts
        if not _is_allowlisted_part(
            name,
            allowed_parts=allowed_parts,
            allowed_patterns=allowed_patterns,
            allowed_prefixes=allowed_prefixes,
        )
    ]
    if unexpected_changed:
        errors.append(f"unexpected changed package parts: {unexpected_changed}")

    writable_sheet_errors = _validate_allowlisted_writable_sheet_parts(
        template_path=template_path,
        output_path=output_path,
        changed_parts=changed_parts,
        writable_cells_by_part=writable_cells_by_part,
    )
    errors.extend(writable_sheet_errors)

    if template.sheet_order != output.sheet_order:
        errors.append(
            f"sheet order changed: template={template.sheet_order}, output={output.sheet_order}"
        )
    if template.sheet_states != output.sheet_states:
        errors.append(
            f"sheet hidden-state map changed: template={template.sheet_states}, output={output.sheet_states}"
        )
    if template.sheet_xml_paths != output.sheet_xml_paths:
        errors.append(
            f"sheet relationship targets changed: template={template.sheet_xml_paths}, output={output.sheet_xml_paths}"
        )
    if template.defined_names != output.defined_names:
        errors.append("defined names changed")

    all_sheet_names = sorted(
        set(template.merges_by_sheet.keys()) | set(output.merges_by_sheet.keys())
    )
    merge_changes = [
        sheet_name
        for sheet_name in all_sheet_names
        if template.merges_by_sheet.get(sheet_name)
        != output.merges_by_sheet.get(sheet_name)
    ]
    if merge_changes:
        errors.append(f"merge ranges changed for sheets: {merge_changes}")

    if template.media_count != output.media_count:
        errors.append(
            f"media count changed: template={template.media_count}, output={output.media_count}"
        )
    if template.drawing_count != output.drawing_count:
        errors.append(
            f"drawing count changed: template={template.drawing_count}, output={output.drawing_count}"
        )

    return errors


def _validate_allowlisted_writable_sheet_parts(
    *,
    template_path: Path,
    output_path: Path,
    changed_parts: list[str],
    writable_cells_by_part: Mapping[str, frozenset[str]],
) -> list[str]:
    target_parts = [part for part in changed_parts if part in writable_cells_by_part]
    if not target_parts:
        return []

    errors: list[str] = []
    with ZipFile(template_path) as template_zip, ZipFile(output_path) as output_zip:
        for part_name in target_parts:
            template_xml = template_zip.read(part_name)
            output_xml = output_zip.read(part_name)
            errors.extend(
                _validate_writable_sheet_value_surface(
                    part_name=part_name,
                    template_xml=template_xml,
                    output_xml=output_xml,
                    writable_cells=writable_cells_by_part[part_name],
                )
            )
    return errors


def _validate_writable_sheet_value_surface(
    *,
    part_name: str,
    template_xml: bytes,
    output_xml: bytes,
    writable_cells: frozenset[str],
) -> list[str]:
    template_root = ET.fromstring(template_xml)
    output_root = ET.fromstring(output_xml)

    template_norm = _normalize_worksheet_for_writable_surface(
        template_root, writable_cells
    )
    output_norm = _normalize_worksheet_for_writable_surface(output_root, writable_cells)

    mismatch = _element_mismatch_reason(template_norm, output_norm, path="worksheet")
    if mismatch is None:
        return []
    return [
        (
            "allowlisted writable sheet contains structural/layout drift "
            f"outside approved cell payloads: part={part_name}; detail={mismatch}"
        )
    ]


def _normalize_worksheet_for_writable_surface(
    root: ET.Element, writable_cells: frozenset[str]
) -> ET.Element:
    writable = {cell.upper() for cell in writable_cells}
    normalized = deepcopy(root)
    for cell in normalized.findall(f".//{CELL_TAG}"):
        cell_ref = cell.attrib.get("r", "").upper()
        if cell_ref not in writable:
            continue
        if "t" in cell.attrib:
            cell.attrib["t"] = "__WILDCARD_CELL_TYPE__"
        for child in list(cell):
            if child.tag in {CELL_VALUE_TAG, CELL_INLINE_TAG}:
                cell.remove(child)
    return normalized


def _element_mismatch_reason(
    left: ET.Element,
    right: ET.Element,
    *,
    path: str,
) -> str | None:
    if left.tag != right.tag:
        return f"tag mismatch at {path}: {left.tag} != {right.tag}"
    if left.attrib != right.attrib:
        return f"attribute mismatch at {path}: {left.attrib} != {right.attrib}"

    left_text = _normalized_xml_text(left.text)
    right_text = _normalized_xml_text(right.text)
    if left_text != right_text:
        return f"text mismatch at {path}: {left_text!r} != {right_text!r}"

    left_children = list(left)
    right_children = list(right)
    if len(left_children) != len(right_children):
        return (
            f"child count mismatch at {path}: "
            f"{len(left_children)} != {len(right_children)}"
        )

    for idx, (left_child, right_child) in enumerate(zip(left_children, right_children)):
        child_path = f"{path}/{_local_name(left_child.tag)}[{idx}]"
        mismatch = _element_mismatch_reason(left_child, right_child, path=child_path)
        if mismatch is not None:
            return mismatch
        left_tail = _normalized_xml_text(left_child.tail)
        right_tail = _normalized_xml_text(right_child.tail)
        if left_tail != right_tail:
            return f"tail mismatch at {child_path}: {left_tail!r} != {right_tail!r}"

    return None


def _normalized_xml_text(value: str | None) -> str:
    if value is None:
        return ""
    if value.strip() == "":
        return ""
    return value


def _local_name(tag: str) -> str:
    if "}" in tag:
        return tag.rsplit("}", 1)[1]
    return tag


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Check XLSX package diff against topology and allowlist invariants"
    )
    _ = parser.add_argument(
        "--template", required=True, help="Template/source xlsx path"
    )
    _ = parser.add_argument(
        "--output", required=True, help="Candidate output xlsx path"
    )
    _ = parser.add_argument(
        "--allowlist",
        required=True,
        help="JSON allowlist for package parts permitted to change",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)

    template = Path(cast(str, args.template))
    output = Path(cast(str, args.output))
    allowlist = Path(cast(str, args.allowlist))

    for path in (template, output, allowlist):
        if not path.exists():
            print(f"error: missing path: {path}")
            return 2

    try:
        errors = validate_package_diff(template, output, allowlist)
    except Exception as exc:
        print(f"error: validator failed: {exc}")
        return 2

    if errors:
        print("FAIL: package invariant validation failed")
        for error in errors:
            print(f" - {error}")
        return 1

    print("PASS: package diff is allowlisted and topology invariants are preserved")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
