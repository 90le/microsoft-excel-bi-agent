#!/usr/bin/env python3
"""Build an audit report for saved cross-agent fresh-session responses."""

from __future__ import annotations

import argparse
import json
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


PASS = "pass"
FAIL = "fail"
PLACEHOLDER_RE = re.compile(
    "|".join([r"\b" + "TO" + "DO" + r"\b", r"\[" + "TO" + "DO" + r"\]", r"FIX" + "ME"]),
    re.IGNORECASE,
)
STUB_MARKERS = [
    "Fresh-agent response:",
    "Paste the fresh-session response below this line before scoring.",
]
FIXTURE_MARKERS = [
    "Evidence terms covered:",
    "This intentionally incomplete response is used to verify that the scorer fails missing evidence.",
]


def now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def read_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def write_json(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def response_path(responses_dir: Path, prompt_entry: dict[str, Any]) -> Path:
    agent = str(prompt_entry.get("agent", "")).strip()
    skill = str(prompt_entry.get("skill", "")).strip()
    return responses_dir / agent / f"{skill}.md"


def classify_response(path: Path) -> dict[str, Any]:
    result: dict[str, Any] = {
        "path": str(path),
        "exists": path.is_file(),
        "kind": "missing",
        "byteCount": 0,
        "lineCount": 0,
        "reviewable": False,
        "placeholder": False,
        "notes": [],
    }
    if not path.is_file():
        result["notes"].append("response file is missing")
        return result

    text = path.read_text(encoding="utf-8")
    stripped = text.strip()
    result["byteCount"] = len(text.encode("utf-8"))
    result["lineCount"] = len(text.splitlines())
    result["placeholder"] = bool(PLACEHOLDER_RE.search(text))
    is_stub = any(marker in text for marker in STUB_MARKERS)
    is_fixture = any(marker in text for marker in FIXTURE_MARKERS) or stripped.startswith("# Forward-test response for ")

    if is_stub:
        result["kind"] = "stub"
        result["notes"].append("response is a generated collection stub")
    elif is_fixture:
        result["kind"] = "generated-fixture"
        result["notes"].append("response appears to be generated scorer fixture evidence")
    elif len(stripped) < 120:
        result["kind"] = "too-short"
        result["notes"].append("response is too short to review")
    elif result["placeholder"]:
        result["kind"] = "placeholder"
        result["notes"].append("response contains placeholder marker")
    else:
        result["kind"] = "candidate-fresh-response"
        result["reviewable"] = True

    return result


def validate_manifest(manifest: dict[str, Any]) -> list[str]:
    failures: list[str] = []
    prompts = [item for item in manifest.get("prompts", []) if isinstance(item, dict)]
    prompt_count = manifest.get("promptCount")
    if manifest.get("status") != PASS:
        failures.append(f"manifest status={manifest.get('status')}")
    if not prompts:
        failures.append("manifest has no prompt entries")
    if isinstance(prompt_count, int) and prompt_count != len(prompts):
        failures.append(f"manifest promptCount={prompt_count}, prompts={len(prompts)}")
    for index, prompt in enumerate(prompts, start=1):
        if not str(prompt.get("agent", "")).strip():
            failures.append(f"prompt {index} missing agent")
        if not str(prompt.get("skill", "")).strip():
            failures.append(f"prompt {index} missing skill")
    return failures


def summarize_score(score: dict[str, Any] | None, expected_count: int) -> dict[str, Any]:
    if score is None:
        return {
            "provided": False,
            "status": "not-provided",
            "expectedResponseCount": None,
            "passedCount": None,
            "failedCount": None,
            "consistentWithManifest": False,
            "allExpectedPassed": False,
        }

    expected = score.get("expectedResponseCount")
    passed = score.get("passedCount")
    failed = score.get("failedCount")
    return {
        "provided": True,
        "status": score.get("status"),
        "expectedResponseCount": expected,
        "passedCount": passed,
        "failedCount": failed,
        "consistentWithManifest": expected == expected_count,
        "allExpectedPassed": score.get("status") == PASS and expected == expected_count and passed == expected_count and failed == 0,
    }


def determine_evidence_status(counts: dict[str, int], score_summary: dict[str, Any]) -> tuple[str, bool]:
    expected = counts["expectedResponseCount"]
    if counts["missingResponseCount"] or counts["stubResponseCount"]:
        return "collecting", False
    if counts["placeholderResponseCount"] or counts["tooShortResponseCount"]:
        return "needs-response-cleanup", False
    if counts["generatedFixtureResponseCount"] and counts["candidateFreshResponseCount"] == 0:
        return "fixture-only", False
    if counts["generatedFixtureResponseCount"]:
        return "mixed-fixture-review", False
    if counts["candidateFreshResponseCount"] != expected:
        return "incomplete-candidate-set", False
    if not score_summary["provided"]:
        return "needs-scoring", False
    if not score_summary["allExpectedPassed"]:
        return "scored-failed", False
    return "external-proof-ready", True


def build_report(manifest_json: Path, responses_dir: Path, score_json: Path | None = None) -> dict[str, Any]:
    manifest_json = manifest_json.expanduser().resolve()
    responses_dir = responses_dir.expanduser().resolve()
    manifest = read_json(manifest_json)
    prompts = [item for item in manifest.get("prompts", []) if isinstance(item, dict)]
    failures = validate_manifest(manifest)

    responses: list[dict[str, Any]] = []
    for prompt in prompts:
        path = response_path(responses_dir, prompt)
        classification = classify_response(path)
        classification["agent"] = str(prompt.get("agent", ""))
        classification["skill"] = str(prompt.get("skill", ""))
        classification["title"] = str(prompt.get("title", ""))
        responses.append(classification)

    counts = {
        "expectedResponseCount": len(prompts),
        "existingResponseCount": sum(1 for item in responses if item["exists"]),
        "missingResponseCount": sum(1 for item in responses if item["kind"] == "missing"),
        "stubResponseCount": sum(1 for item in responses if item["kind"] == "stub"),
        "generatedFixtureResponseCount": sum(1 for item in responses if item["kind"] == "generated-fixture"),
        "candidateFreshResponseCount": sum(1 for item in responses if item["kind"] == "candidate-fresh-response"),
        "tooShortResponseCount": sum(1 for item in responses if item["kind"] == "too-short"),
        "placeholderResponseCount": sum(1 for item in responses if item["kind"] == "placeholder" or item.get("placeholder")),
        "reviewableCandidateCount": sum(1 for item in responses if item.get("reviewable")),
    }

    score = read_json(score_json.expanduser().resolve()) if score_json is not None else None
    score_summary = summarize_score(score, counts["expectedResponseCount"])
    evidence_status, external_proof_ready = determine_evidence_status(counts, score_summary)

    return {
        "generatedAt": now_iso(),
        "status": PASS if not failures else FAIL,
        "evidenceStatus": evidence_status,
        "externalProofReady": external_proof_ready,
        "manifestJson": str(manifest_json),
        "responsesDir": str(responses_dir),
        "scoreJson": str(score_json.expanduser().resolve()) if score_json is not None else None,
        "packVersion": manifest.get("packVersion"),
        "manifestStatus": manifest.get("status"),
        "counts": counts,
        "score": score_summary,
        "responses": responses,
        "failures": failures,
        "boundaries": [
            "This report audits response collection state; it is not external proof unless externalProofReady=true.",
            "Generated response stubs and scorer fixtures validate workflow mechanics only.",
            "Real proof requires fresh-session agent outputs saved under responses/<agent>/<skill>.md and a passing scorer report.",
        ],
    }


def markdown_report(report: dict[str, Any]) -> str:
    counts = report["counts"]
    score = report["score"]
    lines = [
        "# Cross-Agent Response Collection Report",
        "",
        f"- Status: `{report['status']}`",
        f"- Evidence status: `{report['evidenceStatus']}`",
        f"- External proof ready: `{report['externalProofReady']}`",
        f"- Expected responses: `{counts['expectedResponseCount']}`",
        f"- Candidate fresh responses: `{counts['candidateFreshResponseCount']}`",
        f"- Generated fixture responses: `{counts['generatedFixtureResponseCount']}`",
        f"- Stub responses: `{counts['stubResponseCount']}`",
        f"- Missing responses: `{counts['missingResponseCount']}`",
        f"- Score status: `{score['status']}`",
        "",
        "## Response Contract",
        "",
        "Save real fresh-session outputs under:",
        "",
        "```text",
        "responses/<agent>/<skill>.md",
        "```",
        "",
        "| Agent | Skill | Kind | Reviewable | Notes |",
        "|---|---|---|---:|---|",
    ]
    for item in report["responses"]:
        notes = "; ".join(str(note) for note in item.get("notes", []))
        lines.append(
            f"| {item.get('agent', '')} | {item.get('skill', '')} | {item.get('kind', '')} | "
            f"{item.get('reviewable', False)} | {notes} |"
        )
    if report.get("failures"):
        lines.extend(["", "## Failures", ""])
        for failure in report["failures"]:
            lines.append(f"- {failure}")
    lines.extend(
        [
            "",
            "## Boundary",
            "",
            "This report is not external proof unless `externalProofReady=true`. Stubs and generated fixtures are useful workflow evidence, not fresh-agent behavior evidence.",
            "",
        ]
    )
    return "\n".join(lines)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest-json", required=True, type=Path, help="Manifest JSON from build_cross_agent_forward_test_pack.py")
    parser.add_argument("--responses-dir", required=True, type=Path, help="Directory containing responses/<agent>/<skill>.md files")
    parser.add_argument("--score-json", type=Path, help="Optional score JSON from score_cross_agent_forward_test_results.py")
    parser.add_argument("--out-json", type=Path, help="Optional collection report JSON path")
    parser.add_argument("--out-md", type=Path, help="Optional collection report Markdown path")
    parser.add_argument("--require-pass", action="store_true", help="Return non-zero when report generation or manifest validation fails")
    parser.add_argument("--require-external-proof", action="store_true", help="Return non-zero unless all responses are candidate fresh responses and scorer evidence passes")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    report = build_report(args.manifest_json, args.responses_dir, args.score_json)

    if args.out_json:
        write_json(args.out_json.expanduser().resolve(), report)
    if args.out_md:
        out_md = args.out_md.expanduser().resolve()
        out_md.parent.mkdir(parents=True, exist_ok=True)
        out_md.write_text(markdown_report(report), encoding="utf-8")

    print(
        json.dumps(
            {
                "status": report["status"],
                "evidenceStatus": report["evidenceStatus"],
                "externalProofReady": report["externalProofReady"],
                "counts": report["counts"],
            },
            ensure_ascii=False,
        )
    )
    if args.require_pass and report["status"] != PASS:
        return 1
    if args.require_external_proof and not report["externalProofReady"]:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
