#!/usr/bin/env bash
set -euo pipefail

usage() {
  cat <<'EOF'
Usage:
  invoke_excel_bi_com.sh inspect-model -w WORKBOOK [--include-columns] [--out-json PATH]
  invoke_excel_bi_com.sh model-report -w WORKBOOK [--include-columns] [--out-md PATH] [--out-json PATH] [--keep-intermediate]
  invoke_excel_bi_com.sh cube-report -w WORKBOOK [--include-model|--model-json PATH] [--out-md PATH] [--out-json PATH] [--out-mermaid PATH] [--detail-limit N] [--keep-intermediate]
  invoke_excel_bi_com.sh ado-query -w WORKBOOK --sql SQL [--create-fixture] [--include-schema] [--out-json PATH] [--provider PROVIDER]
  invoke_excel_bi_com.sh adomd-query [--probe-only] [--connection-string TEXT] [--mdx TEXT|--mdx-file PATH] [--max-cells N] [--out-json PATH]
  invoke_excel_bi_com.sh provider-probe [--out-json PATH] [--excel-com] [--ado-workbook-smoke] [--smoke-workbook PATH] [--provider PROVIDER]

This wrapper is for Windows Git Bash/MSYS/Cygwin. Most commands call
PowerShell scripts that require desktop Excel COM. The ado-query command
requires a compatible ADO/OLEDB provider and only needs Excel COM when
--create-fixture is used. The provider-probe command diagnoses local Office,
ADO/OLEDB, MSOLAP, and ADOMD availability. The adomd-query command validates
an explicit ADOMD/MSOLAP endpoint connection string and MDX query. On Linux/macOS, use
inspect_excel_bi_workbook.py for structural OpenXML inspection.
EOF
}

script_dir="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

case "$(uname -s 2>/dev/null || echo unknown)" in
  MINGW*|MSYS*|CYGWIN*) ;;
  *)
    echo "Excel COM automation requires Windows desktop Excel. Use inspect_excel_bi_workbook.py for structural inspection." >&2
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

if command -v python >/dev/null 2>&1; then
  py_exe="python"
elif command -v python3 >/dev/null 2>&1; then
  py_exe="python3"
else
  py_exe=""
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
  inspect-model)
    workbook=""
    out_json=""
    include_columns=""
    while [[ $# -gt 0 ]]; do
      case "$1" in
        -w|--workbook) workbook="$2"; shift 2 ;;
        --out-json) out_json="$2"; shift 2 ;;
        --include-columns) include_columns="1"; shift ;;
        -h|--help) usage; exit 0 ;;
        *) echo "Unknown inspect-model option: $1" >&2; usage; exit 2 ;;
      esac
    done
    [[ -n "$workbook" ]] || { echo "Missing -w WORKBOOK." >&2; exit 2; }
    args=(-NoProfile -ExecutionPolicy Bypass -File "$(to_win_path "$script_dir/inspect_excel_data_model_com.ps1")" -WorkbookPath "$(to_win_path "$workbook")")
    [[ -n "$out_json" ]] && args+=(-OutJson "$(to_win_path "$out_json")")
    [[ -n "$include_columns" ]] && args+=(-IncludeColumns)
    "$ps_exe" "${args[@]}"
    ;;
  model-report)
    workbook=""
    out_md=""
    out_json=""
    include_columns=""
    keep_intermediate=""
    while [[ $# -gt 0 ]]; do
      case "$1" in
        -w|--workbook) workbook="$2"; shift 2 ;;
        --out-md) out_md="$2"; shift 2 ;;
        --out-json) out_json="$2"; shift 2 ;;
        --include-columns) include_columns="1"; shift ;;
        --keep-intermediate) keep_intermediate="1"; shift ;;
        -h|--help) usage; exit 0 ;;
        *) echo "Unknown model-report option: $1" >&2; usage; exit 2 ;;
      esac
    done
    [[ -n "$workbook" ]] || { echo "Missing -w WORKBOOK." >&2; exit 2; }
    [[ -n "$py_exe" ]] || { echo "Cannot find python or python3 on PATH." >&2; exit 2; }

    tmp_dir="$(mktemp -d 2>/dev/null || printf '%s/excel_bi_model_report_%s' "${TMPDIR:-/tmp}" "$$")"
    mkdir -p "$tmp_dir"
    model_json="$tmp_dir/data_model.json"
    openxml_json="$tmp_dir/openxml.json"

    ps_args=(-NoProfile -ExecutionPolicy Bypass -File "$(to_win_path "$script_dir/inspect_excel_data_model_com.ps1")" -WorkbookPath "$(to_win_path "$workbook")" -OutJson "$(to_win_path "$model_json")")
    [[ -n "$include_columns" ]] && ps_args+=(-IncludeColumns)
    "$ps_exe" "${ps_args[@]}"

    "$py_exe" "$(to_win_path "$script_dir/inspect_excel_bi_workbook.py")" "$(to_win_path "$workbook")" --out-json "$(to_win_path "$openxml_json")"

    report_args=("--model-json" "$(to_win_path "$model_json")" "--openxml-json" "$(to_win_path "$openxml_json")")
    [[ -n "$out_md" ]] && report_args+=("--out-md" "$(to_win_path "$out_md")")
    [[ -n "$out_json" ]] && report_args+=("--out-json" "$(to_win_path "$out_json")")
    [[ -z "$out_md" && -z "$out_json" ]] && report_args+=("--print")
    "$py_exe" "$(to_win_path "$script_dir/build_excel_bi_model_report.py")" "${report_args[@]}"

    if [[ -z "$keep_intermediate" ]]; then
      rm -rf "$tmp_dir"
    else
      echo "Kept intermediate files in: $tmp_dir"
    fi
    ;;
  cube-report)
    workbook=""
    model_json=""
    out_md=""
    out_json=""
    out_mermaid=""
    detail_limit=""
    include_model=""
    keep_intermediate=""
    while [[ $# -gt 0 ]]; do
      case "$1" in
        -w|--workbook) workbook="$2"; shift 2 ;;
        --model-json) model_json="$2"; shift 2 ;;
        --include-model) include_model="1"; shift ;;
        --out-md) out_md="$2"; shift 2 ;;
        --out-json) out_json="$2"; shift 2 ;;
        --out-mermaid) out_mermaid="$2"; shift 2 ;;
        --detail-limit) detail_limit="$2"; shift 2 ;;
        --keep-intermediate) keep_intermediate="1"; shift ;;
        -h|--help) usage; exit 0 ;;
        *) echo "Unknown cube-report option: $1" >&2; usage; exit 2 ;;
      esac
    done
    [[ -n "$workbook" ]] || { echo "Missing -w WORKBOOK." >&2; exit 2; }
    [[ -n "$py_exe" ]] || { echo "Cannot find python or python3 on PATH." >&2; exit 2; }

    tmp_dir="$(mktemp -d 2>/dev/null || printf '%s/excel_bi_cube_report_%s' "${TMPDIR:-/tmp}" "$$")"
    mkdir -p "$tmp_dir"
    openxml_json="$tmp_dir/openxml.json"

    "$py_exe" "$(to_win_path "$script_dir/inspect_excel_bi_workbook.py")" "$(to_win_path "$workbook")" --out-json "$(to_win_path "$openxml_json")"

    if [[ -n "$include_model" && -z "$model_json" ]]; then
      model_json="$tmp_dir/data_model.json"
      ps_args=(-NoProfile -ExecutionPolicy Bypass -File "$(to_win_path "$script_dir/inspect_excel_data_model_com.ps1")" -WorkbookPath "$(to_win_path "$workbook")" -OutJson "$(to_win_path "$model_json")")
      "$ps_exe" "${ps_args[@]}"
    fi

    cube_args=("--openxml-json" "$(to_win_path "$openxml_json")")
    [[ -n "$model_json" ]] && cube_args+=("--model-json" "$(to_win_path "$model_json")")
    [[ -n "$out_md" ]] && cube_args+=("--out-md" "$(to_win_path "$out_md")")
    [[ -n "$out_json" ]] && cube_args+=("--out-json" "$(to_win_path "$out_json")")
    [[ -n "$out_mermaid" ]] && cube_args+=("--out-mermaid" "$(to_win_path "$out_mermaid")")
    [[ -n "$detail_limit" ]] && cube_args+=("--detail-limit" "$detail_limit")
    [[ -z "$out_md" && -z "$out_json" && -z "$out_mermaid" ]] && cube_args+=("--print")
    "$py_exe" "$(to_win_path "$script_dir/build_cube_dependency_report.py")" "${cube_args[@]}"

    if [[ -z "$keep_intermediate" ]]; then
      rm -rf "$tmp_dir"
    else
      echo "Kept intermediate files in: $tmp_dir"
    fi
    ;;
  ado-query)
    workbook=""
    sql=""
    out_json=""
    create_fixture=""
    include_schema=""
    provider=""
    while [[ $# -gt 0 ]]; do
      case "$1" in
        -w|--workbook) workbook="$2"; shift 2 ;;
        --sql) sql="$2"; shift 2 ;;
        --out-json) out_json="$2"; shift 2 ;;
        --create-fixture) create_fixture="1"; shift ;;
        --include-schema) include_schema="1"; shift ;;
        --provider) provider="$2"; shift 2 ;;
        -h|--help) usage; exit 0 ;;
        *) echo "Unknown ado-query option: $1" >&2; usage; exit 2 ;;
      esac
    done
    [[ -n "$workbook" ]] || { echo "Missing -w WORKBOOK." >&2; exit 2; }
    [[ -n "$sql" ]] || { echo "Missing --sql SQL." >&2; exit 2; }

    ado_args=(-NoProfile -ExecutionPolicy Bypass -File "$(to_win_path "$script_dir/test_excel_ado_sql_access.ps1")" -WorkbookPath "$(to_win_path "$workbook")" -SqlText "$sql")
    [[ -n "$out_json" ]] && ado_args+=(-OutJson "$(to_win_path "$out_json")")
    [[ -n "$create_fixture" ]] && ado_args+=(-CreateFixture)
    [[ -n "$include_schema" ]] && ado_args+=(-IncludeSchema)
    [[ -n "$provider" ]] && ado_args+=(-Provider "$provider")
    "$ps_exe" "${ado_args[@]}"
    ;;
  adomd-query)
    probe_only=""
    connection_string=""
    mdx=""
    mdx_file=""
    out_json=""
    max_cells=""
    while [[ $# -gt 0 ]]; do
      case "$1" in
        --probe-only) probe_only="1"; shift ;;
        --connection-string) connection_string="$2"; shift 2 ;;
        --mdx) mdx="$2"; shift 2 ;;
        --mdx-file) mdx_file="$2"; shift 2 ;;
        --out-json) out_json="$2"; shift 2 ;;
        --max-cells) max_cells="$2"; shift 2 ;;
        -h|--help) usage; exit 0 ;;
        *) echo "Unknown adomd-query option: $1" >&2; usage; exit 2 ;;
      esac
    done

    adomd_args=(-NoProfile -ExecutionPolicy Bypass -File "$(to_win_path "$script_dir/test_excel_adomd_query.ps1")")
    [[ -n "$probe_only" ]] && adomd_args+=(-ProbeOnly)
    [[ -n "$connection_string" ]] && adomd_args+=(-ConnectionString "$connection_string")
    [[ -n "$mdx" ]] && adomd_args+=(-Mdx "$mdx")
    [[ -n "$mdx_file" ]] && adomd_args+=(-MdxPath "$(to_win_path "$mdx_file")")
    [[ -n "$out_json" ]] && adomd_args+=(-OutJson "$(to_win_path "$out_json")")
    [[ -n "$max_cells" ]] && adomd_args+=(-MaxCells "$max_cells")
    "$ps_exe" "${adomd_args[@]}"
    ;;
  provider-probe)
    out_json=""
    excel_com=""
    ado_workbook_smoke=""
    smoke_workbook=""
    provider=""
    while [[ $# -gt 0 ]]; do
      case "$1" in
        --out-json) out_json="$2"; shift 2 ;;
        --excel-com) excel_com="1"; shift ;;
        --ado-workbook-smoke) ado_workbook_smoke="1"; shift ;;
        --smoke-workbook) smoke_workbook="$2"; shift 2 ;;
        --provider) provider="$2"; shift 2 ;;
        -h|--help) usage; exit 0 ;;
        *) echo "Unknown provider-probe option: $1" >&2; usage; exit 2 ;;
      esac
    done

    probe_args=(-NoProfile -ExecutionPolicy Bypass -File "$(to_win_path "$script_dir/probe_excel_bi_providers.ps1")")
    [[ -n "$out_json" ]] && probe_args+=(-OutJson "$(to_win_path "$out_json")")
    [[ -n "$excel_com" ]] && probe_args+=(-RunExcelComSmoke)
    [[ -n "$ado_workbook_smoke" ]] && probe_args+=(-RunAdoWorkbookSmoke)
    [[ -n "$smoke_workbook" ]] && probe_args+=(-SmokeWorkbookPath "$(to_win_path "$smoke_workbook")")
    [[ -n "$provider" ]] && probe_args+=(-AdoSmokeProvider "$provider")
    "$ps_exe" "${probe_args[@]}"
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
