#!/usr/bin/env bash
set -euo pipefail

usage() {
  cat <<'EOF'
Usage:
  run_release_gate.sh [run_release_gate.py options]

Runs the Microsoft Excel BI Agent Pack release gate from Git Bash, Linux, or
macOS. If --project-root is omitted, the plugin root is inferred from this
script location. If --out-json or --out-md is omitted, reports are written to a
temporary directory.

Examples:
  tools/run_release_gate.sh
  tools/run_release_gate.sh --profile structural
  tools/run_release_gate.sh --out-json /tmp/excel_bi_gate.json --out-md /tmp/excel_bi_gate.md
  tools/run_release_gate.sh --project-root /path/to/microsoft_excel_bi_agent_pack
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
  local fallback=""
  for candidate_name in python python3 py; do
    if command -v "$candidate_name" >/dev/null 2>&1; then
      candidate="$(command -v "$candidate_name")"
      [[ -n "$fallback" ]] || fallback="$candidate"
      if "$candidate" -c 'import yaml' >/dev/null 2>&1; then
        printf '%s\n' "$candidate"
        return 0
      fi
    fi
  done
  if [[ -n "$fallback" ]]; then
    printf '%s\n' "$fallback"
    return 0
  fi
  return 1
}

has_arg() {
  local key="$1"
  shift
  local arg
  for arg in "$@"; do
    case "$arg" in
      "$key"|"$key"=*) return 0 ;;
    esac
  done
  return 1
}

if [[ "${1:-}" == "-h" || "${1:-}" == "--help" ]]; then
  usage
  exit 0
fi

py_exe="$(find_python)" || {
  echo "Cannot find python3 or python on PATH." >&2
  exit 2
}

tmp_base="${TMPDIR:-/tmp}"
if [[ "$is_windows_shell" -eq 1 && -n "${TEMP:-}" ]]; then
  tmp_base="$TEMP"
fi

timestamp="$(date +%Y%m%d%H%M%S)"
tmp_dir="$(mktemp -d "$tmp_base/excel_bi_release_gate_${timestamp}_XXXXXX" 2>/dev/null || true)"
if [[ -z "$tmp_dir" ]]; then
  tmp_dir="$tmp_base/excel_bi_release_gate_${timestamp}_$$"
  mkdir -p "$tmp_dir"
fi

out_json="$tmp_dir/release_gate.json"
out_md="$tmp_dir/release_gate.md"

args=("$@")
if ! has_arg "--project-root" "${args[@]}"; then
  args+=("--project-root" "$(to_native_path "$project_root")")
fi
if ! has_arg "--out-json" "${args[@]}"; then
  args+=("--out-json" "$(to_native_path "$out_json")")
fi
if ! has_arg "--out-md" "${args[@]}"; then
  args+=("--out-md" "$(to_native_path "$out_md")")
fi

gate_script="$(to_native_path "$script_dir/run_release_gate.py")"

set +e
"$py_exe" "$gate_script" "${args[@]}"
status=$?
set -e

printf 'Release gate JSON: %s\n' "$out_json"
printf 'Release gate Markdown: %s\n' "$out_md"
exit "$status"
