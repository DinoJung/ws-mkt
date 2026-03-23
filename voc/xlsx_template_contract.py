from __future__ import annotations

import hashlib
import json
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping, Sequence
from zipfile import ZipFile


CONTRACT_PATH = Path(".sisyphus/contracts/xlsx-template-contract.json")
XML_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
XML_NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XML_NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"


@dataclass(frozen=True)
class WorkbookContractError(ValueError):
    code: str
    detail: str

    def __str__(self) -> str:
        return f"{self.code}: {self.detail}"


def load_template_contract(path: Path | None = None) -> dict[str, Any]:
    contract_path = path or CONTRACT_PATH
    return json.loads(contract_path.read_text(encoding="utf-8"))


def supported_workbook_fingerprint(path: Path) -> dict[str, Any]:
    with ZipFile(path) as zf:
        workbook_xml = zf.read("xl/workbook.xml")
        workbook = ET.fromstring(workbook_xml)
    return {
        "file_sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
        "workbook_xml_sha256": hashlib.sha256(workbook_xml).hexdigest(),
        "sheet_order": [sheet.attrib["name"] for sheet in _workbook_sheets(workbook)],
        "sheet_states": {
            sheet.attrib["name"]: sheet.attrib.get("state", "visible")
            for sheet in _workbook_sheets(workbook)
        },
    }


def derive_linked_card_targets(
    path: Path, contract: Mapping[str, Any]
) -> dict[str, list[str]]:
    rules = contract["writable_surface"]["linked_card_rules"]
    anchor_row = int(rules["card_anchor_row"])
    header_row = int(rules["header_row"])
    link_row = int(rules["link_row"])
    write_row = int(rules["write_row"])
    link_offset = int(rules["link_column_offset_from_anchor"])
    write_offset = int(rules["write_column_offset_from_anchor"])
    anchor_label = str(rules["card_anchor_label"])
    header_label = str(rules["header_label"])

    with ZipFile(path) as zf:
        shared_strings = _shared_strings(zf)
        sheet_paths = _sheet_xml_paths(zf)
        linked_targets: dict[str, list[str]] = {}
        for sheet_name in contract["preserved_topology"]["sheet_order"]:
            sheet_path = sheet_paths.get(sheet_name)
            if sheet_path is None:
                continue
            cell_text, hyperlinks = _worksheet_cell_maps(zf, sheet_path, shared_strings)
            anchor_cols = [
                col
                for ref, value in cell_text.items()
                if value == anchor_label and _split_cell_ref(ref)[1] == anchor_row
                for col, _ in [_split_cell_ref(ref)]
            ]
            cells: list[str] = []
            for anchor_col in sorted(anchor_cols):
                if cell_text.get(_cell_ref(anchor_col, header_row)) != header_label:
                    continue
                link_ref = _cell_ref(anchor_col + link_offset, link_row)
                link_value = hyperlinks.get(link_ref) or cell_text.get(link_ref)
                if not link_value:
                    continue
                cells.append(_cell_ref(anchor_col + write_offset, write_row))
            if cells:
                linked_targets[sheet_name] = cells
    return linked_targets


def expand_allowed_writable_surface(
    contract: Mapping[str, Any],
) -> dict[str, frozenset[str]]:
    allowed: dict[str, set[str]] = {}
    primary_target = contract["writable_surface"]["primary_target"]
    allowed.setdefault(str(primary_target["sheet_name"]), set()).update(
        _expand_ranges(primary_target["allowed_ranges"])
    )
    for sheet_name, cell_refs in contract["writable_surface"][
        "linked_card_targets"
    ].items():
        allowed.setdefault(str(sheet_name), set()).update(
            str(cell_ref) for cell_ref in cell_refs
        )
    return {
        sheet_name: frozenset(cell_refs) for sheet_name, cell_refs in allowed.items()
    }


def validate_supported_workbook(
    path: Path, contract: Mapping[str, Any]
) -> dict[str, Any]:
    fingerprint = supported_workbook_fingerprint(path)
    topology = contract["preserved_topology"]
    family = contract["supported_workbook_family"]
    primary_target = contract["writable_surface"]["primary_target"]

    if int(topology["sheet_count"]) != len(fingerprint["sheet_order"]):
        raise WorkbookContractError(
            "unsupported_sheet_count",
            "sheet count differs from supported workbook family",
        )
    if str(primary_target["sheet_name"]) not in fingerprint["sheet_states"]:
        raise WorkbookContractError(
            "unsupported_missing_target_sheet",
            f"missing required sheet {primary_target['sheet_name']}",
        )
    if fingerprint["sheet_order"] != list(topology["sheet_order"]):
        raise WorkbookContractError(
            "unsupported_sheet_order",
            "sheet order differs from supported workbook family",
        )
    if fingerprint["sheet_states"] != dict(topology["sheet_states"]):
        raise WorkbookContractError(
            "unsupported_sheet_state",
            "sheet hidden-state map differs from supported workbook family",
        )

    linked_card_targets = derive_linked_card_targets(path, contract)
    if linked_card_targets != dict(contract["writable_surface"]["linked_card_targets"]):
        raise WorkbookContractError(
            "unsupported_missing_linked_card_layout",
            "linked-card write targets differ from supported workbook family",
        )

    family_workbook_hashes = _supported_family_workbook_hashes(family)
    if family_workbook_hashes and (
        fingerprint["workbook_xml_sha256"] not in family_workbook_hashes
    ):
        raise WorkbookContractError(
            "unsupported_workbook_family_signature",
            "workbook.xml fingerprint differs from supported workbook family",
        )

    fingerprint["linked_card_targets"] = linked_card_targets
    return fingerprint


def enforce_writable_surface(
    updates: Sequence[Mapping[str, str]], contract: Mapping[str, Any]
) -> None:
    allowed = expand_allowed_writable_surface(contract)
    for update in updates:
        sheet_name = update["sheet_name"]
        cell_ref = update["cell_ref"]
        if sheet_name not in allowed:
            raise WorkbookContractError(
                "unsupported_write_sheet",
                f"write targets unsupported sheet {sheet_name}",
            )
        if cell_ref not in allowed[sheet_name]:
            raise WorkbookContractError(
                "unsupported_write_cell",
                f"write targets unsupported cell {sheet_name}!{cell_ref}",
            )


def _workbook_sheets(workbook: ET.Element) -> list[ET.Element]:
    return list(workbook.findall(f".//{{{XML_NS_MAIN}}}sheet"))


def _shared_strings(zf: ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    strings: list[str] = []
    for string_item in root.findall(f"{{{XML_NS_MAIN}}}si"):
        text_nodes = string_item.findall(f".//{{{XML_NS_MAIN}}}t")
        strings.append("".join(node.text or "" for node in text_nodes))
    return strings


def _sheet_xml_paths(zf: ZipFile) -> dict[str, str]:
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_map = {
        rel.attrib["Id"]: _normalize_xl_path(rel.attrib["Target"])
        for rel in rels
        if rel.attrib.get("Id") and rel.attrib.get("Target")
    }
    return {
        sheet.attrib["name"]: rel_map[sheet.attrib[f"{{{XML_NS_REL}}}id"]]
        for sheet in _workbook_sheets(workbook)
        if sheet.attrib.get("name")
        and sheet.attrib.get(f"{{{XML_NS_REL}}}id") in rel_map
    }


def _worksheet_cell_maps(
    zf: ZipFile, sheet_path: str, shared_strings: list[str]
) -> tuple[dict[str, str], dict[str, str]]:
    root = ET.fromstring(zf.read(sheet_path))
    cell_text = {
        cell_ref: value
        for cell in root.findall(f".//{{{XML_NS_MAIN}}}c")
        if (cell_ref := cell.attrib.get("r"))
        and (value := _cell_value(cell, shared_strings)) is not None
    }
    return cell_text, _worksheet_hyperlinks(zf, root, sheet_path)


def _worksheet_hyperlinks(
    zf: ZipFile, root: ET.Element, sheet_path: str
) -> dict[str, str]:
    rels_path = _sheet_rels_path(sheet_path)
    if rels_path not in zf.namelist():
        return {}
    rels = ET.fromstring(zf.read(rels_path))
    rel_map = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in rels
        if rel.attrib.get("Id") and rel.attrib.get("Target")
    }
    hyperlinks: dict[str, str] = {}
    for hyperlink in root.findall(f".//{{{XML_NS_MAIN}}}hyperlink"):
        cell_ref = hyperlink.attrib.get("ref")
        rel_id = hyperlink.attrib.get(f"{{{XML_NS_REL}}}id")
        if cell_ref and rel_id and rel_id in rel_map:
            hyperlinks[cell_ref] = rel_map[rel_id]
    return hyperlinks


def _cell_value(cell: ET.Element, shared_strings: list[str]) -> str | None:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        text_nodes = cell.findall(f".//{{{XML_NS_MAIN}}}t")
        return "".join(node.text or "" for node in text_nodes)
    value_node = cell.find(f"{{{XML_NS_MAIN}}}v")
    if value_node is None or value_node.text is None:
        return None
    if cell_type == "s":
        shared_idx = int(value_node.text)
        if 0 <= shared_idx < len(shared_strings):
            return shared_strings[shared_idx]
        return None
    return value_node.text


def _sheet_rels_path(sheet_path: str) -> str:
    prefix, filename = sheet_path.rsplit("/", 1)
    return f"{prefix}/_rels/{filename}.rels"


def _normalize_xl_path(target: str) -> str:
    target = target.lstrip("/")
    if target.startswith("xl/"):
        return target
    return f"xl/{target}"


def _split_cell_ref(cell_ref: str) -> tuple[int, int]:
    match = re.fullmatch(r"([A-Z]+)(\d+)", cell_ref)
    if match is None:
        raise WorkbookContractError(
            "unsupported_write_cell", f"invalid cell reference {cell_ref}"
        )
    return _column_index(match.group(1)), int(match.group(2))


def _cell_ref(column_index: int, row_index: int) -> str:
    return f"{_column_letters(column_index)}{row_index}"


def _expand_ranges(ranges: Sequence[Any]) -> set[str]:
    cells: set[str] = set()
    for range_item in ranges:
        start_ref, end_ref = str(range_item).split(":", maxsplit=1)
        start_col, start_row = _split_cell_ref(start_ref)
        end_col, end_row = _split_cell_ref(end_ref)
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                cells.add(_cell_ref(col, row))
    return cells


def _column_index(column_letters: str) -> int:
    result = 0
    for char in column_letters:
        result = result * 26 + (ord(char) - ord("A") + 1)
    return result


def _column_letters(column_index: int) -> str:
    letters: list[str] = []
    current = column_index
    while current > 0:
        current, remainder = divmod(current - 1, 26)
        letters.append(chr(ord("A") + remainder))
    return "".join(reversed(letters))


def _supported_family_workbook_hashes(family: Mapping[str, Any]) -> set[str]:
    hashes: set[str] = set()
    single = family.get("workbook_xml_sha256")
    if single:
        hashes.add(str(single))
    for item in family.get("workbook_xml_sha256_allowlist", []):
        hashes.add(str(item))
    return hashes
