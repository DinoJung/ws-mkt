import importlib
import json
from pathlib import Path


def _bench_module():
    return importlib.import_module("scripts.bench_xlsx_output")


def test_benchmark_guardrail_passes_on_speedup_threshold():
    bench = _bench_module()

    comparison = {
        "output_stage_speedup_factor_baseline_vs_candidate": 3.1,
        "output_stage_reduction_fraction": 0.4,
    }
    guardrail = bench.evaluate_output_stage_guardrail(comparison, justification=None)

    assert guardrail["passed"] is True
    assert guardrail["route"] == "speedup"


def test_benchmark_guardrail_requires_justification_for_reduction_route():
    bench = _bench_module()

    comparison = {
        "output_stage_speedup_factor_baseline_vs_candidate": 2.2,
        "output_stage_reduction_fraction": 0.62,
    }

    guardrail_without_justification = bench.evaluate_output_stage_guardrail(
        comparison, justification=None
    )
    assert guardrail_without_justification["passed"] is False
    assert guardrail_without_justification["route"] == "failed"

    guardrail_with_justification = bench.evaluate_output_stage_guardrail(
        comparison,
        justification="Workbook-specific layout parity work keeps us on reduction gate",
    )
    assert guardrail_with_justification["passed"] is True
    assert guardrail_with_justification["route"] == "reduction_with_justification"


def test_benchmark_guardrail_fails_on_regression():
    bench = _bench_module()

    comparison = {
        "output_stage_speedup_factor_baseline_vs_candidate": 1.2,
        "output_stage_reduction_fraction": 0.16,
    }
    guardrail = bench.evaluate_output_stage_guardrail(
        comparison,
        justification="regression should still fail regardless of note",
    )

    assert guardrail["passed"] is False
    assert guardrail["route"] == "failed"


def test_benchmark_guardrail_json_includes_comparison_metrics(monkeypatch, tmp_path):
    bench = _bench_module()

    fixture = tmp_path / "fixture.xlsx"
    fixture.write_bytes(b"fixture")
    json_out = tmp_path / "bench.json"

    def fake_run_candidate(candidate: str, fixture_path: Path, runs: int):
        assert fixture_path == fixture
        assert runs == 1
        if candidate == "primary":
            rows = [
                {
                    "load_seconds": 1.0,
                    "manifest_build_seconds": 0.2,
                    "classify_seconds": 1.0,
                    "write_seconds": 2.0,
                    "save_seconds": 1.0,
                    "output_stage_seconds": 3.0,
                    "total_seconds": 5.0,
                    "update_count": 10,
                    "classified_count": 5,
                    "writer": "primary",
                    "run": 1,
                    "output_artifact": "primary-run-1.xlsx",
                }
            ]
            return rows, b"primary"
        if candidate == "legacy":
            rows = [
                {
                    "load_seconds": 1.0,
                    "manifest_build_seconds": 0.2,
                    "classify_seconds": 1.0,
                    "write_seconds": 9.0,
                    "save_seconds": 3.0,
                    "output_stage_seconds": 12.0,
                    "total_seconds": 14.0,
                    "update_count": 10,
                    "classified_count": 5,
                    "writer": "legacy",
                    "run": 1,
                    "output_artifact": "legacy-run-1.xlsx",
                }
            ]
            return rows, b"legacy"
        raise AssertionError(f"unexpected candidate: {candidate}")

    monkeypatch.setattr(bench, "_run_candidate", fake_run_candidate)

    exit_code = bench.main(
        [
            "--fixture",
            str(fixture),
            "--candidate",
            "primary",
            "--baseline",
            "legacy",
            "--runs",
            "1",
            "--json-out",
            str(json_out),
            "--enforce-guardrail",
        ]
    )

    assert exit_code == 0
    payload = json.loads(json_out.read_text(encoding="utf-8"))
    assert payload["candidate"] == "primary"
    assert payload["baseline"] == "legacy"
    assert (
        payload["results"]["primary"]["median_seconds"]["manifest_build_seconds"] == 0.2
    )
    assert (
        payload["results"]["primary"]["median_seconds"]["output_stage_seconds"] == 3.0
    )
    assert (
        payload["results"]["legacy"]["median_seconds"]["output_stage_seconds"] == 12.0
    )
    assert payload["comparison"]["output_stage_ratio_candidate_vs_baseline"] == 0.25
    assert payload["comparison"]["output_stage_reduction_fraction"] == 0.75
    assert (
        payload["comparison"]["output_stage_speedup_factor_baseline_vs_candidate"]
        == 4.0
    )
    assert payload["comparison"]["guardrail"]["passed"] is True


def test_benchmark_guardrail_enforce_fails_when_threshold_not_met(
    monkeypatch, tmp_path
):
    bench = _bench_module()

    fixture = tmp_path / "fixture.xlsx"
    fixture.write_bytes(b"fixture")
    json_out = tmp_path / "bench.json"

    def fake_run_candidate(candidate: str, fixture_path: Path, runs: int):
        assert fixture_path == fixture
        assert runs == 1
        rows = [
            {
                "load_seconds": 1.0,
                "manifest_build_seconds": 0.2,
                "classify_seconds": 1.0,
                "write_seconds": 8.0,
                "save_seconds": 2.0,
                "output_stage_seconds": 10.0 if candidate == "primary" else 11.0,
                "total_seconds": 13.0,
                "update_count": 10,
                "classified_count": 5,
                "writer": candidate,
                "run": 1,
                "output_artifact": f"{candidate}-run-1.xlsx",
            }
        ]
        return rows, candidate.encode("utf-8")

    monkeypatch.setattr(bench, "_run_candidate", fake_run_candidate)

    exit_code = bench.main(
        [
            "--fixture",
            str(fixture),
            "--candidate",
            "primary",
            "--baseline",
            "legacy",
            "--runs",
            "1",
            "--json-out",
            str(json_out),
            "--enforce-guardrail",
        ]
    )

    assert exit_code == 1
