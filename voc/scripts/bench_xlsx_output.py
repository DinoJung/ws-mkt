#!/usr/bin/env python3
"""Benchmark harness for VOC XLSX output paths."""

from __future__ import annotations

import argparse
import importlib
import json
import statistics
import sys
import tempfile
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

classify_voc = importlib.import_module("classify_voc")
openpyxl = importlib.import_module("openpyxl")

JsonDict = dict[str, Any]

OUTPUT_STAGE_SPEEDUP_TARGET = 3.0
OUTPUT_STAGE_REDUCTION_TARGET = 0.60


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _collect_benchmark_updates(
    targets: list[JsonDict], card_index: dict[str, list[tuple[str, int, int]]]
) -> tuple[list[JsonDict], int]:
    updates: list[JsonDict] = []
    classified_count = 0
    for item in targets:
        updates.extend(
            classify_voc._build_sheet_updates(
                item["row"],
                item.get("link"),
                "반응_기타",
                card_index,
            )
        )
        classified_count += 1
    return updates, classified_count


def _measure_run(writer: str, fixture: Path, output_path: Path) -> JsonDict:
    started = time.perf_counter()

    t0 = time.perf_counter()
    wb = openpyxl.load_workbook(fixture)
    load_s = time.perf_counter() - t0

    t0 = time.perf_counter()
    card_index = classify_voc.build_card_index(wb)
    targets = classify_voc.read_target_rows(wb)
    manifest_build_s = time.perf_counter() - t0

    t0 = time.perf_counter()
    updates, classified_count = _collect_benchmark_updates(targets, card_index)
    classify_s = time.perf_counter() - t0

    t0 = time.perf_counter()
    classify_voc._apply_updates_to_workbook(wb, updates)
    if writer == "legacy":
        classify_voc._apply_output_layout_to_workbook(wb)
    write_s = time.perf_counter() - t0

    t0 = time.perf_counter()
    if writer == "legacy":
        classify_voc.save_workbook_with_preserved_media_legacy(
            fixture, output_path, updates
        )
    elif writer == "primary":
        classify_voc.save_workbook_with_preserved_media(fixture, output_path, updates)
    else:
        raise ValueError(f"unsupported writer: {writer}")
    save_s = time.perf_counter() - t0

    total_s = time.perf_counter() - started
    return {
        "writer": writer,
        "load_seconds": load_s,
        "manifest_build_seconds": manifest_build_s,
        "classify_seconds": classify_s,
        "write_seconds": write_s,
        "save_seconds": save_s,
        "output_stage_seconds": write_s + save_s,
        "total_seconds": total_s,
        "update_count": len(updates),
        "classified_count": classified_count,
    }


def _run_candidate(
    candidate: str, fixture: Path, runs: int
) -> tuple[list[JsonDict], bytes | None]:
    if candidate not in {"legacy", "primary"}:
        raise ValueError(f"unsupported candidate writer: {candidate}")

    run_data: list[JsonDict] = []
    emitted_output: bytes | None = None
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        for idx in range(1, runs + 1):
            output_path = tmp / f"{candidate}-run-{idx}.xlsx"
            row = _measure_run(candidate, fixture, output_path)
            row["run"] = idx
            row["output_artifact"] = output_path.name
            run_data.append(row)
            if emitted_output is None:
                emitted_output = output_path.read_bytes()

    return run_data, emitted_output


def _median_metrics(rows: list[JsonDict]) -> JsonDict:
    keys = (
        "load_seconds",
        "manifest_build_seconds",
        "classify_seconds",
        "write_seconds",
        "save_seconds",
        "output_stage_seconds",
        "total_seconds",
    )
    return {key: statistics.median([row[key] for row in rows]) for key in keys}


def _build_comparison_metrics(
    candidate_median: JsonDict, baseline_median: JsonDict
) -> JsonDict:
    candidate_out = candidate_median["output_stage_seconds"]
    baseline_out = baseline_median["output_stage_seconds"]
    ratio_candidate_vs_baseline = candidate_out / baseline_out if baseline_out else None
    reduction_fraction = (
        (baseline_out - candidate_out) / baseline_out if baseline_out else None
    )
    speedup_factor = baseline_out / candidate_out if candidate_out else None
    return {
        "output_stage_ratio_candidate_vs_baseline": ratio_candidate_vs_baseline,
        "output_stage_reduction_fraction": reduction_fraction,
        "output_stage_speedup_factor_baseline_vs_candidate": speedup_factor,
    }


def evaluate_output_stage_guardrail(
    comparison: JsonDict,
    *,
    justification: str | None,
) -> JsonDict:
    speedup_factor = comparison.get("output_stage_speedup_factor_baseline_vs_candidate")
    reduction_fraction = comparison.get("output_stage_reduction_fraction")
    has_justification = bool(justification and justification.strip())
    passed_speedup = bool(
        speedup_factor is not None and speedup_factor >= OUTPUT_STAGE_SPEEDUP_TARGET
    )
    passed_reduction_with_justification = bool(
        reduction_fraction is not None
        and reduction_fraction >= OUTPUT_STAGE_REDUCTION_TARGET
        and has_justification
    )
    passed = passed_speedup or passed_reduction_with_justification
    if passed_speedup:
        route = "speedup"
    elif passed_reduction_with_justification:
        route = "reduction_with_justification"
    else:
        route = "failed"

    return {
        "passed": passed,
        "route": route,
        "target": {
            "speedup_factor_min": OUTPUT_STAGE_SPEEDUP_TARGET,
            "reduction_fraction_min": OUTPUT_STAGE_REDUCTION_TARGET,
            "reduction_requires_justification": True,
        },
        "justification": justification if has_justification else None,
    }


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Benchmark VOC XLSX output writers")
    parser.add_argument("--fixture", required=True, help="Input workbook fixture path")
    parser.add_argument(
        "--candidate",
        required=True,
        choices=["legacy", "primary"],
        help="Writer path to benchmark",
    )
    parser.add_argument(
        "--baseline",
        choices=["legacy", "primary"],
        help="Optional baseline writer for ratio comparison",
    )
    parser.add_argument(
        "--json-out", required=True, help="Where to write benchmark JSON"
    )
    parser.add_argument(
        "--guardrail-justification",
        help="Documented justification required when only the 60%% reduction threshold is used",
    )
    parser.add_argument(
        "--enforce-guardrail",
        action="store_true",
        help="Exit non-zero when benchmark does not meet Task 7 threshold policy",
    )
    parser.add_argument(
        "--emit-known-bad",
        help="Optional output workbook path for reproducible legacy fixture",
    )
    parser.add_argument("--runs", type=int, default=3, help="Benchmark iteration count")
    return parser


def _write_json_atomically(path: Path, payload: JsonDict):
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
    temp.replace(path)


def main(argv: list[str] | None = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)

    fixture = Path(args.fixture)
    json_out = Path(args.json_out)

    if args.runs < 1:
        print("error: --runs must be >= 1")
        return 2
    if not fixture.exists():
        print(f"error: missing fixture: {fixture}")
        return 2

    results: dict[str, JsonDict] = {}

    try:
        candidate_rows, candidate_output = _run_candidate(
            args.candidate, fixture, args.runs
        )
        results[args.candidate] = {
            "runs": candidate_rows,
            "median_seconds": _median_metrics(candidate_rows),
        }

        if args.baseline and args.baseline != args.candidate:
            baseline_rows, _ = _run_candidate(args.baseline, fixture, args.runs)
            results[args.baseline] = {
                "runs": baseline_rows,
                "median_seconds": _median_metrics(baseline_rows),
            }
        else:
            baseline_rows = None

        payload: JsonDict = {
            "generated_at": _now_iso(),
            "fixture": str(fixture),
            "runs": args.runs,
            "candidate": args.candidate,
            "baseline": args.baseline,
            "results": results,
        }

        if args.baseline and args.baseline in results:
            comparison = _build_comparison_metrics(
                results[args.candidate]["median_seconds"],
                results[args.baseline]["median_seconds"],
            )
            comparison["guardrail"] = evaluate_output_stage_guardrail(
                comparison,
                justification=args.guardrail_justification,
            )
            payload["comparison"] = comparison

        _write_json_atomically(json_out, payload)

        if args.enforce_guardrail:
            if "comparison" not in payload:
                raise ValueError("--enforce-guardrail requires distinct --baseline")
            guardrail = payload["comparison"].get("guardrail")
            if not isinstance(guardrail, dict) or not guardrail.get("passed"):
                print("error: benchmark guardrail check failed")
                return 1

        if args.emit_known_bad:
            if args.candidate != "legacy":
                raise ValueError(
                    "--emit-known-bad is supported only with --candidate legacy"
                )
            if candidate_output is None:
                raise RuntimeError("legacy run produced no output to emit")
            known_bad_path = Path(args.emit_known_bad)
            known_bad_path.parent.mkdir(parents=True, exist_ok=True)
            known_bad_path.write_bytes(candidate_output)

    except Exception as exc:
        print(f"error: benchmark failed: {exc}")
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
