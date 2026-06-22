#!/usr/bin/env bash
set -euo pipefail

usage() {
  cat <<'EOF'
Usage:
  excel_bi_structural.sh sanitized-bundle --out-dir DIR [--clean] [--validate] [--out-json PATH] [--out-md PATH]
  excel_bi_structural.sh pq-lineage --query-dir DIR [--out-json PATH] [--out-md PATH] [--fail-on-high-risk]
  excel_bi_structural.sh provider-baseline-fixture --out-dir DIR [--clean]
  excel_bi_structural.sh release-gate [run_release_gate.py options]

Portable structural helper for Git Bash, Linux, and macOS. It only runs
OpenXML/static-source paths and does not claim Excel COM, VBA execution,
Power Query refresh, Power Pivot calculation, Solver, ADO, or ADOMD runtime
behavior.
EOF
}

script_dir="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
project_root="$(cd "$script_dir/.." && pwd)"

is_windows_shell=0
case "$(uname -s 2>/dev/null || echo unknown)" in
  MINGW*|MSYS*|CYGWIN*) is_windows_shell=1 ;;
esac

to_native_path() {
  local path="$1"
  if [[ "$is_windows_shell" -eq 1 ]] && command -v cygpath >/dev/null 2>&1; then
    cygpath -w "$path"
  else
    printf '%s\n' "$path"
  fi
}

find_python() {
  local candidate
  for candidate in python python3 py; do
    if command -v "$candidate" >/dev/null 2>&1; then
      printf '%s\n' "$candidate"
      return 0
    fi
  done
  return 1
}

run_py() {
  "$py_exe" "$@"
}

require_arg() {
  local label="$1"
  local value="$2"
  if [[ -z "$value" ]]; then
    echo "Missing $label." >&2
    usage
    exit 2
  fi
}

safe_clean_dir() {
  local target="$1"
  run_py - "$target" <<'PY'
import shutil
import sys
from pathlib import Path

target = Path(sys.argv[1]).expanduser().resolve()
home = Path.home().resolve()
if str(target) == target.anchor:
    raise SystemExit(f"refusing to remove filesystem root: {target}")
if target == home:
    raise SystemExit(f"refusing to remove user home directory: {target}")
if len(target.parts) < 3:
    raise SystemExit(f"refusing to remove shallow directory: {target}")
if target.exists():
    shutil.rmtree(target)
PY
}

if [[ $# -lt 1 || "${1:-}" == "-h" || "${1:-}" == "--help" ]]; then
  usage
  exit 0
fi

py_exe="$(find_python)" || {
  echo "Cannot find python, python3, or py on PATH." >&2
  exit 2
}

cmd="$1"
shift

case "$cmd" in
  sanitized-bundle)
    out_dir=""
    out_json=""
    out_md=""
    clean=""
    validate=""
    while [[ $# -gt 0 ]]; do
      case "$1" in
        --out-dir) out_dir="$2"; shift 2 ;;
        --out-json) out_json="$2"; shift 2 ;;
        --out-md) out_md="$2"; shift 2 ;;
        --clean) clean="1"; shift ;;
        --validate) validate="1"; shift ;;
        -h|--help) usage; exit 0 ;;
        *) echo "Unknown sanitized-bundle option: $1" >&2; usage; exit 2 ;;
      esac
    done
    require_arg "--out-dir DIR" "$out_dir"

    bundle_args=("--out-dir" "$(to_native_path "$out_dir")")
    [[ -n "$clean" ]] && bundle_args+=("--clean")
    [[ -n "$out_json" ]] && bundle_args+=("--out-json" "$(to_native_path "$out_json")")
    [[ -n "$out_md" ]] && bundle_args+=("--out-md" "$(to_native_path "$out_md")")

    run_py "$(to_native_path "$script_dir/build_sanitized_fixture_bundle.py")" "${bundle_args[@]}"

    if [[ -n "$validate" ]]; then
      validation_dir="$out_dir/_validation"
      mkdir -p "$validation_dir"

      run_py "$(to_native_path "$script_dir/inspect_excel_bi_workbook.py")" "$(to_native_path "$out_dir/cube_formula_fixture.xlsx")" --out-json "$(to_native_path "$validation_dir/cube_openxml.json")"
      run_py "$(to_native_path "$script_dir/build_cube_dependency_report.py")" --openxml-json "$(to_native_path "$validation_dir/cube_openxml.json")" --model-json "$(to_native_path "$out_dir/cube_model_summary.json")" --out-json "$(to_native_path "$validation_dir/cube_report.json")"
      run_py "$(to_native_path "$script_dir/inspect_excel_bi_workbook.py")" "$(to_native_path "$out_dir/external_dependency_fixture.xlsx")" --out-json "$(to_native_path "$validation_dir/external_openxml.json")"
      run_py "$(to_native_path "$script_dir/build_external_dependency_report.py")" --openxml-json "$(to_native_path "$validation_dir/external_openxml.json")" --out-json "$(to_native_path "$validation_dir/external_readiness.json")"
      run_py "$(to_native_path "$script_dir/inspect_excel_bi_workbook.py")" "$(to_native_path "$out_dir/pure_deliverable_fixture.xlsx")" --out-json "$(to_native_path "$validation_dir/pure_openxml.json")"
      run_py "$(to_native_path "$script_dir/build_external_dependency_report.py")" --openxml-json "$(to_native_path "$validation_dir/pure_openxml.json")" --out-json "$(to_native_path "$validation_dir/pure_readiness.json")"
      run_py "$(to_native_path "$script_dir/build_power_query_lineage_report.py")" "$(to_native_path "$out_dir/power_query_lineage/safe")" --out-json "$(to_native_path "$validation_dir/pq_lineage_safe.json")" --fail-on-high-risk
      run_py "$(to_native_path "$script_dir/build_power_query_lineage_report.py")" "$(to_native_path "$out_dir/power_query_lineage/risky")" --out-json "$(to_native_path "$validation_dir/pq_lineage_risky.json")"

      run_py - "$out_dir" "$validation_dir/structural_wrapper_summary.json" <<'PY'
import json
import sys
from pathlib import Path

bundle = Path(sys.argv[1])
out = Path(sys.argv[2])
validation = out.parent

def read_json(path: Path):
    return json.loads(path.read_text(encoding="utf-8"))

manifest = read_json(bundle / "fixture-bundle.json")
cube = read_json(validation / "cube_report.json")
external = read_json(validation / "external_readiness.json")
pure = read_json(validation / "pure_readiness.json")
pq_safe = read_json(validation / "pq_lineage_safe.json")
pq_risky = read_json(validation / "pq_lineage_risky.json")

summary = {
    "fixtureCount": manifest.get("fixtureCount"),
    "fixtureIds": [item.get("id") for item in manifest.get("fixtures", []) if isinstance(item, dict)],
    "cubeFormulaCount": cube.get("cubeFormulaCount"),
    "externalReadiness": external.get("summary", {}).get("readiness"),
    "externalCredentialLikeConnectionCount": external.get("summary", {}).get("credentialLikeConnectionCount"),
    "pureReadiness": pure.get("summary", {}).get("readiness"),
    "pqSafeQueryCount": pq_safe.get("summary", {}).get("queryCount"),
    "pqSafeFindingCount": pq_safe.get("summary", {}).get("findingCount"),
    "pqRiskyQueryCount": pq_risky.get("summary", {}).get("queryCount"),
    "pqRiskyHighFindingCount": pq_risky.get("summary", {}).get("highFindingCount"),
    "pqRiskyCodes": sorted({finding.get("code") for finding in pq_risky.get("findings", [])}),
}
out.parent.mkdir(parents=True, exist_ok=True)
out.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps({"summary": str(out), "fixtureCount": summary["fixtureCount"]}, ensure_ascii=False))
PY
    fi
    ;;

  pq-lineage)
    query_dir=""
    out_json=""
    out_md=""
    fail_on_high_risk=""
    while [[ $# -gt 0 ]]; do
      case "$1" in
        --query-dir) query_dir="$2"; shift 2 ;;
        --out-json) out_json="$2"; shift 2 ;;
        --out-md) out_md="$2"; shift 2 ;;
        --fail-on-high-risk) fail_on_high_risk="1"; shift ;;
        -h|--help) usage; exit 0 ;;
        *) echo "Unknown pq-lineage option: $1" >&2; usage; exit 2 ;;
      esac
    done
    require_arg "--query-dir DIR" "$query_dir"
    args=("$(to_native_path "$query_dir")")
    [[ -n "$out_json" ]] && args+=("--out-json" "$(to_native_path "$out_json")")
    [[ -n "$out_md" ]] && args+=("--out-md" "$(to_native_path "$out_md")")
    [[ -n "$fail_on_high_risk" ]] && args+=("--fail-on-high-risk")
    run_py "$(to_native_path "$script_dir/build_power_query_lineage_report.py")" "${args[@]}"
    ;;

  provider-baseline-fixture)
    out_dir=""
    clean=""
    while [[ $# -gt 0 ]]; do
      case "$1" in
        --out-dir) out_dir="$2"; shift 2 ;;
        --clean) clean="1"; shift ;;
        -h|--help) usage; exit 0 ;;
        *) echo "Unknown provider-baseline-fixture option: $1" >&2; usage; exit 2 ;;
      esac
    done
    require_arg "--out-dir DIR" "$out_dir"
    if [[ -n "$clean" && -e "$out_dir" ]]; then
      safe_clean_dir "$out_dir"
    fi
    mkdir -p "$out_dir"

    fixture_dir="$out_dir/fixture"
    reports_dir="$out_dir/reports"
    mkdir -p "$reports_dir"
    manifest_json="$reports_dir/provider_environment_fixture.json"
    matching_report_json="$reports_dir/provider_matching_report.json"
    matching_report_md="$reports_dir/provider_matching_report.md"
    drift_report_json="$reports_dir/provider_drift_report.json"
    drift_report_md="$reports_dir/provider_drift_report.md"
    summary_json="$reports_dir/provider_baseline_wrapper_summary.json"

    run_py "$(to_native_path "$script_dir/create_provider_environment_fixture.py")" \
      --out-dir "$(to_native_path "$fixture_dir")" \
      --out-json "$(to_native_path "$manifest_json")"

    run_py "$(to_native_path "$script_dir/build_provider_environment_report.py")" \
      --project-root "$(to_native_path "$project_root")" \
      --probe-json "$(to_native_path "$fixture_dir/provider_fixture_probe.json")" \
      --baseline-json "$(to_native_path "$fixture_dir/provider_matching_baseline.json")" \
      --fail-on-drift \
      --out-json "$(to_native_path "$matching_report_json")" \
      --out-md "$(to_native_path "$matching_report_md")" \
      --require-pass

    set +e
    run_py "$(to_native_path "$script_dir/build_provider_environment_report.py")" \
      --project-root "$(to_native_path "$project_root")" \
      --probe-json "$(to_native_path "$fixture_dir/provider_fixture_probe.json")" \
      --baseline-json "$(to_native_path "$fixture_dir/provider_drifting_baseline.json")" \
      --fail-on-drift \
      --out-json "$(to_native_path "$drift_report_json")" \
      --out-md "$(to_native_path "$drift_report_md")" \
      --require-pass
    drift_status=$?
    set -e
    if [[ "$drift_status" -eq 0 ]]; then
      echo "Expected drifting provider baseline to fail, but it passed." >&2
      exit 1
    fi

    run_py - "$manifest_json" "$matching_report_json" "$drift_report_json" "$summary_json" <<'PY'
import json
import sys
from pathlib import Path

manifest_path = Path(sys.argv[1])
matching_path = Path(sys.argv[2])
drift_path = Path(sys.argv[3])
summary_path = Path(sys.argv[4])

def read_json(path: Path):
    return json.loads(path.read_text(encoding="utf-8"))

manifest = read_json(manifest_path)
matching = read_json(matching_path)
drift = read_json(drift_path)
expected = manifest.get("expected", {})
summary = {
    "matchingChangedCount": matching.get("comparison", {}).get("changedCount"),
    "driftChangedCount": drift.get("comparison", {}).get("changedCount"),
    "driftPaths": sorted(
        item.get("path")
        for item in drift.get("comparison", {}).get("changes", [])
        if isinstance(item, dict)
    ),
    "expectedMinimumDriftCount": expected.get("driftingMinimumChangedCount"),
    "requiredDriftPaths": expected.get("driftPaths", []),
    "matchingStatus": matching.get("status"),
    "driftStatus": drift.get("status"),
}
summary_path.parent.mkdir(parents=True, exist_ok=True)
summary_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps({"summary": str(summary_path), "driftChangedCount": summary["driftChangedCount"]}, ensure_ascii=False))
PY
    ;;

  release-gate)
    "$script_dir/run_release_gate.sh" "$@"
    ;;

  *)
    echo "Unknown command: $cmd" >&2
    usage
    exit 2
    ;;
esac
