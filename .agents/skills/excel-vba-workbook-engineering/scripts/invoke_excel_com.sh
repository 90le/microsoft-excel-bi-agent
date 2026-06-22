#!/usr/bin/env bash
set -euo pipefail

usage() {
  cat <<'EOF'
Usage:
  invoke_excel_com.sh inspect -w WORKBOOK [-o OUT_JSON]
  invoke_excel_com.sh export  -w WORKBOOK -d OUT_DIR
  invoke_excel_com.sh import  -w WORKBOOK -s SOURCE_DIR -o OUTPUT_WORKBOOK

This wrapper is for Windows Git Bash/MSYS/Cygwin. It calls the PowerShell
Excel COM scripts after converting paths with cygpath when available.
On Linux/macOS without Windows Excel COM, use inspect_openxml.py instead.
EOF
}

script_dir="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

case "$(uname -s 2>/dev/null || echo unknown)" in
  MINGW*|MSYS*|CYGWIN*) ;;
  *)
    echo "Excel COM automation requires Windows desktop Excel. Use inspect_openxml.py for cross-platform structural inspection." >&2
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

run_ps() {
  "$ps_exe" -NoProfile -ExecutionPolicy Bypass -File "$@"
}

if [[ $# -lt 1 ]]; then
  usage
  exit 2
fi

cmd="$1"
shift

case "$cmd" in
  inspect)
    workbook=""
    out_json=""
    while [[ $# -gt 0 ]]; do
      case "$1" in
        -w|--workbook) workbook="$2"; shift 2 ;;
        -o|--out-json) out_json="$2"; shift 2 ;;
        -h|--help) usage; exit 0 ;;
        *) echo "Unknown inspect option: $1" >&2; usage; exit 2 ;;
      esac
    done
    [[ -n "$workbook" ]] || { echo "Missing -w WORKBOOK." >&2; exit 2; }
    ps_script="$(to_win_path "$script_dir/inspect_workbook.ps1")"
    workbook="$(to_win_path "$workbook")"
    if [[ -n "$out_json" ]]; then
      out_json="$(to_win_path "$out_json")"
      run_ps "$ps_script" -WorkbookPath "$workbook" -OutJson "$out_json"
    else
      run_ps "$ps_script" -WorkbookPath "$workbook"
    fi
    ;;
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
    run_ps "$(to_win_path "$script_dir/export_vba.ps1")" -WorkbookPath "$(to_win_path "$workbook")" -OutDir "$(to_win_path "$out_dir")"
    ;;
  import)
    workbook=""
    source_dir=""
    output=""
    while [[ $# -gt 0 ]]; do
      case "$1" in
        -w|--workbook) workbook="$2"; shift 2 ;;
        -s|--source-dir) source_dir="$2"; shift 2 ;;
        -o|--output) output="$2"; shift 2 ;;
        -h|--help) usage; exit 0 ;;
        *) echo "Unknown import option: $1" >&2; usage; exit 2 ;;
      esac
    done
    [[ -n "$workbook" && -n "$source_dir" && -n "$output" ]] || { echo "Missing -w WORKBOOK, -s SOURCE_DIR, or -o OUTPUT." >&2; exit 2; }
    run_ps "$(to_win_path "$script_dir/import_vba.ps1")" -WorkbookPath "$(to_win_path "$workbook")" -SourceDir "$(to_win_path "$source_dir")" -OutputWorkbookPath "$(to_win_path "$output")"
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
