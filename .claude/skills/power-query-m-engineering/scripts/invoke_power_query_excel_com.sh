#!/usr/bin/env bash
set -euo pipefail

usage() {
  cat <<'EOF'
Usage:
  invoke_power_query_excel_com.sh export -w WORKBOOK -d OUT_DIR
  invoke_power_query_excel_com.sh manage -w WORKBOOK -a ACTION [options]
  invoke_power_query_excel_com.sh refresh -w WORKBOOK [options]
  invoke_power_query_excel_com.sh fixture -o OUTPUT_WORKBOOK [options]

This wrapper is for Windows Git Bash/MSYS/Cygwin. It calls the PowerShell
Excel COM exporter after converting paths with cygpath when available.
On Linux/macOS without Windows Excel COM, use inspect_power_query_openxml.py.
EOF
}

script_dir="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

case "$(uname -s 2>/dev/null || echo unknown)" in
  MINGW*|MSYS*|CYGWIN*) ;;
  *)
    echo "Excel COM automation requires Windows desktop Excel. Use inspect_power_query_openxml.py for structural inspection." >&2
    exit 2
    ;;
esac

if command -v powershell.exe >/dev/null 2>&1; then
  ps_exe="powershell.exe"
elif command -v pwsh.exe >/dev/null 2>&1; then
  ps_exe="pwsh.exe"
else
  echo "Cannot find powershell.exe or pwsh.exe on PATH." >&2
  exit 2
fi

to_win_path() {
  local path="$1"
  if command -v cygpath >/dev/null 2>&1; then
    cygpath -w "$path"
  else
    printf '%s\n' "$path"
  fi
}

if [[ $# -lt 1 ]]; then
  usage
  exit 2
fi

cmd="$1"
shift

case "$cmd" in
  export)
    workbook=""
    out_dir=""
    while [[ $# -gt 0 ]]; do
      case "$1" in
        -w|--workbook) workbook="$2"; shift 2 ;;
        -d|--out-dir) out_dir="$2"; shift 2 ;;
        -h|--help) usage; exit 0 ;;
        *) echo "Unknown export option: $1" >&2; usage; exit 2 ;;
      esac
    done
    [[ -n "$workbook" && -n "$out_dir" ]] || { echo "Missing -w WORKBOOK or -d OUT_DIR." >&2; exit 2; }
    "$ps_exe" -NoProfile -ExecutionPolicy Bypass \
      -File "$(to_win_path "$script_dir/export_power_queries_excel_com.ps1")" \
      -WorkbookPath "$(to_win_path "$workbook")" \
      -OutDir "$(to_win_path "$out_dir")"
    ;;
  manage)
    workbook=""
    action=""
    query_name=""
    formula_path=""
    description=""
    output=""
    out_json=""
    inplace=""
    while [[ $# -gt 0 ]]; do
      case "$1" in
        -w|--workbook) workbook="$2"; shift 2 ;;
        -a|--action) action="$2"; shift 2 ;;
        -q|--query-name) query_name="$2"; shift 2 ;;
        -f|--formula-path) formula_path="$2"; shift 2 ;;
        --description) description="$2"; shift 2 ;;
        -o|--output) output="$2"; shift 2 ;;
        --out-json) out_json="$2"; shift 2 ;;
        --in-place) inplace="1"; shift ;;
        -h|--help) usage; exit 0 ;;
        *) echo "Unknown manage option: $1" >&2; usage; exit 2 ;;
      esac
    done
    [[ -n "$workbook" && -n "$action" ]] || { echo "Missing -w WORKBOOK or -a ACTION." >&2; exit 2; }
    args=(-NoProfile -ExecutionPolicy Bypass -File "$(to_win_path "$script_dir/manage_power_queries_excel_com.ps1")" -WorkbookPath "$(to_win_path "$workbook")" -Action "$action")
    [[ -n "$query_name" ]] && args+=(-QueryName "$query_name")
    [[ -n "$formula_path" ]] && args+=(-FormulaPath "$(to_win_path "$formula_path")")
    [[ -n "$description" ]] && args+=(-Description "$description")
    [[ -n "$output" ]] && args+=(-OutputWorkbookPath "$(to_win_path "$output")")
    [[ -n "$out_json" ]] && args+=(-OutJson "$(to_win_path "$out_json")")
    [[ -n "$inplace" ]] && args+=(-InPlace)
    "$ps_exe" "${args[@]}"
    ;;
  refresh)
    workbook=""
    query_name=""
    output=""
    out_json=""
    timeout=""
    disable_background=""
    calculate_full=""
    while [[ $# -gt 0 ]]; do
      case "$1" in
        -w|--workbook) workbook="$2"; shift 2 ;;
        -q|--query-name) query_name="$2"; shift 2 ;;
        -o|--output) output="$2"; shift 2 ;;
        --out-json) out_json="$2"; shift 2 ;;
        --timeout-seconds) timeout="$2"; shift 2 ;;
        --disable-background-refresh) disable_background="1"; shift ;;
        --calculate-full) calculate_full="1"; shift ;;
        -h|--help) usage; exit 0 ;;
        *) echo "Unknown refresh option: $1" >&2; usage; exit 2 ;;
      esac
    done
    [[ -n "$workbook" ]] || { echo "Missing -w WORKBOOK." >&2; exit 2; }
    args=(-NoProfile -ExecutionPolicy Bypass -File "$(to_win_path "$script_dir/refresh_power_queries_excel_com.ps1")" -WorkbookPath "$(to_win_path "$workbook")")
    [[ -n "$query_name" ]] && args+=(-QueryName "$query_name")
    [[ -n "$output" ]] && args+=(-OutputWorkbookPath "$(to_win_path "$output")")
    [[ -n "$out_json" ]] && args+=(-OutJson "$(to_win_path "$out_json")")
    [[ -n "$timeout" ]] && args+=(-TimeoutSeconds "$timeout")
    [[ -n "$disable_background" ]] && args+=(-DisableBackgroundRefresh)
    [[ -n "$calculate_full" ]] && args+=(-CalculateFull)
    "$ps_exe" "${args[@]}"
    ;;
  fixture)
    output=""
    query_name=""
    table_name=""
    formula_path=""
    out_json=""
    while [[ $# -gt 0 ]]; do
      case "$1" in
        -o|--output) output="$2"; shift 2 ;;
        -q|--query-name) query_name="$2"; shift 2 ;;
        --table-name) table_name="$2"; shift 2 ;;
        -f|--formula-path) formula_path="$2"; shift 2 ;;
        --out-json) out_json="$2"; shift 2 ;;
        -h|--help) usage; exit 0 ;;
        *) echo "Unknown fixture option: $1" >&2; usage; exit 2 ;;
      esac
    done
    [[ -n "$output" ]] || { echo "Missing -o OUTPUT_WORKBOOK." >&2; exit 2; }
    args=(-NoProfile -ExecutionPolicy Bypass -File "$(to_win_path "$script_dir/create_power_query_fixture_excel_com.ps1")" -OutputWorkbookPath "$(to_win_path "$output")")
    [[ -n "$query_name" ]] && args+=(-QueryName "$query_name")
    [[ -n "$table_name" ]] && args+=(-TableName "$table_name")
    [[ -n "$formula_path" ]] && args+=(-FormulaPath "$(to_win_path "$formula_path")")
    [[ -n "$out_json" ]] && args+=(-OutJson "$(to_win_path "$out_json")")
    "$ps_exe" "${args[@]}"
    ;;
  -h|--help)
    usage
    ;;
  *)
    echo "Unknown command: $cmd" >&2
    usage
    exit 2
    ;;
esac
