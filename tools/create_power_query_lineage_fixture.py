#!/usr/bin/env python3
"""Create generic Power Query M source fixtures for lineage/source-risk reports."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any


def write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def write_json(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def write_manifest(query_dir: Path, queries: list[tuple[str, Path, str]]) -> None:
    write_json(
        query_dir / "power_queries.json",
        {
            "workbookPath": "power_query_lineage_fixture.xlsx",
            "outDir": str(query_dir),
            "queryCount": len(queries),
            "queries": [
                {
                    "index": index,
                    "name": name,
                    "description": description,
                    "formulaFile": str(path),
                }
                for index, (name, path, description) in enumerate(queries, start=1)
            ],
            "connections": [],
        },
    )


def create_safe_fixture(out_dir: Path) -> dict[str, Any]:
    safe_dir = out_dir / "safe"
    source_sales = safe_dir / "001_SourceSales.m"
    clean_sales = safe_dir / "002_CleanSales.m"
    path_config = safe_dir / "003_PathConfig.m"
    parameterized_folder = safe_dir / "004_ParameterizedFolder.m"
    write_text(
        source_sales,
        "\n".join(
            [
                "let",
                '    Source = #table({"Id", "Amount"}, {{1, 10}, {2, 20}})',
                "in",
                "    Source",
                "",
            ]
        ),
    )
    write_text(
        clean_sales,
        "\n".join(
            [
                "let",
                '    Source = #"SourceSales",',
                '    ChangedType = Table.TransformColumnTypes(Source, {{"Id", Int64.Type}, {"Amount", type number}})',
                "in",
                "    ChangedType",
                "",
            ]
        ),
    )
    write_text(
        path_config,
        "\n".join(
            [
                "let",
                '    Source = Excel.CurrentWorkbook(){[Name="SourceFolderTable"]}[Content],',
                '    FolderPath = Text.From(Source{0}[FolderPath])',
                "in",
                "    FolderPath",
                "",
            ]
        ),
    )
    write_text(
        parameterized_folder,
        "\n".join(
            [
                "let",
                '    FolderPath = #"PathConfig",',
                "    Source = Folder.Files(FolderPath),",
                '    Visible = Table.SelectRows(Source, each [Attributes]?[Hidden]? <> true and not Text.StartsWith([Name], "~$"))',
                "in",
                "    Visible",
                "",
            ]
        ),
    )
    write_manifest(
        safe_dir,
        [
            ("SourceSales", source_sales, "inline safe source"),
            ("CleanSales", clean_sales, "query dependency safe transform"),
            ("PathConfig", path_config, "workbook table parameter source"),
            ("ParameterizedFolder", parameterized_folder, "parameterized local folder source"),
        ],
    )
    return {
        "queryDir": str(safe_dir),
        "expected": {
            "readiness": "clean",
            "queryCount": 4,
            "dependencyCount": 2,
            "findingCount": 0,
            "sourceKindCounts": {
                "local-file": 1,
                "workbook-config": 2,
            },
        },
    }


def create_risky_fixture(out_dir: Path) -> dict[str, Any]:
    risky_dir = out_dir / "risky"
    local_files = risky_dir / "001_LocalFiles.m"
    web_data = risky_dir / "002_WebData.m"
    merged = risky_dir / "003_Merged.m"
    sql_data = risky_dir / "004_SqlData.m"
    odata_data = risky_dir / "005_ODataData.m"
    dataflow_data = risky_dir / "006_DataflowData.m"
    azure_blob_data = risky_dir / "007_AzureBlobData.m"
    native_sql_data = risky_dir / "008_NativeSqlData.m"
    api_with_header = risky_dir / "009_ApiWithHeader.m"
    cycle_a = risky_dir / "010_CycleA.m"
    cycle_b = risky_dir / "011_CycleB.m"
    write_text(
        local_files,
        "\n".join(
            [
                "let",
                '    Source = Folder.Files("C:\\Users\\analyst\\Desktop\\data"),',
                '    Visible = Table.SelectRows(Source, each [Attributes]?[Hidden]? <> true and not Text.StartsWith([Name], "~$"))',
                "in",
                "    Visible",
                "",
            ]
        ),
    )
    write_text(
        web_data,
        "\n".join(
            [
                "let",
                '    Source = Web.Contents("https://example.com/data.csv"),',
                "    Csv = Csv.Document(Source)",
                "in",
                "    Csv",
                "",
            ]
        ),
    )
    write_text(
        merged,
        "\n".join(
            [
                "let",
                '    Local = #"LocalFiles",',
                '    Web = #"WebData",',
                '    Merged = Table.NestedJoin(Local, {"Name"}, Web, {"Name"}, "WebRows", JoinKind.LeftOuter)',
                "in",
                "    Merged",
                "",
            ]
        ),
    )
    write_text(
        sql_data,
        "\n".join(
            [
                "let",
                '    Source = Sql.Database("warehouse.example.internal", "FinanceModel"),',
                '    Data = Source{[Schema="dbo", Item="FactSales"]}[Data]',
                "in",
                "    Data",
                "",
            ]
        ),
    )
    write_text(
        odata_data,
        "\n".join(
            [
                "let",
                '    Source = OData.Feed("https://services.example.com/finance/odata"),',
                '    Sales = Source{[Name="Sales", Signature="table"]}[Data]',
                "in",
                "    Sales",
                "",
            ]
        ),
    )
    write_text(
        dataflow_data,
        "\n".join(
            [
                "let",
                "    Source = PowerPlatform.Dataflows(null),",
                '    Workspace = Source{[workspaceName="Finance"]}[Data],',
                '    Dataflow = Workspace{[dataflowName="Budget"]}[Data]',
                "in",
                "    Dataflow",
                "",
            ]
        ),
    )
    write_text(
        azure_blob_data,
        "\n".join(
            [
                "let",
                '    Source = AzureStorage.Blobs("https://storage.example.com/reports"),',
                '    Visible = Table.SelectRows(Source, each [Extension] = ".csv")',
                "in",
                "    Visible",
                "",
            ]
        ),
    )
    write_text(
        native_sql_data,
        "\n".join(
            [
                "let",
                '    Source = Sql.Database("warehouse.example.internal", "FinanceModel"),',
                '    Data = Value.NativeQuery(Source, "select Region, Amount from dbo.FactSales where Period >= @StartPeriod", [StartPeriod = "2026Q1"])',
                "in",
                "    Data",
                "",
            ]
        ),
    )
    write_text(
        api_with_header,
        "\n".join(
            [
                "let",
                '    Source = Web.Contents("https://api.example.com/report", [Headers=[Authorization="Bearer redacted-example-token"]]),',
                "    Json = Json.Document(Source)",
                "in",
                "    Json",
                "",
            ]
        ),
    )
    write_text(
        cycle_a,
        "\n".join(["let", '    Source = #"CycleB"', "in", "    Source", ""]),
    )
    write_text(
        cycle_b,
        "\n".join(["let", '    Source = #"CycleA"', "in", "    Source", ""]),
    )
    write_manifest(
        risky_dir,
        [
            ("LocalFiles", local_files, "hard-coded local source"),
            ("WebData", web_data, "web source"),
            ("Merged", merged, "mixed-source dependency lineage"),
            ("SqlData", sql_data, "database source"),
            ("ODataData", odata_data, "OData feed source"),
            ("DataflowData", dataflow_data, "Power Platform dataflow source"),
            ("AzureBlobData", azure_blob_data, "Azure Storage source"),
            ("NativeSqlData", native_sql_data, "native SQL review source"),
            ("ApiWithHeader", api_with_header, "web source with authorization-like header"),
            ("CycleA", cycle_a, "cycle fixture A"),
            ("CycleB", cycle_b, "cycle fixture B"),
        ],
    )
    return {
        "queryDir": str(risky_dir),
        "expected": {
            "readiness": "blocked-for-delivery",
            "queryCount": 11,
            "dependencyCount": 4,
            "highFindingCount": 3,
            "mediumFindingCount": 9,
            "sourceKindCounts": {
                "cloud-service": 2,
                "database": 2,
                "local-file": 2,
                "web": 4,
            },
            "requiredCodes": [
                "hard-coded-local-path",
                "web-source",
                "database-source",
                "cloud-service-source",
                "native-query-review",
                "credential-like-literal",
                "mixed-source-lineage",
                "query-dependency-cycle",
            ],
        },
    }


def create_fixture(out_dir: Path) -> dict[str, Any]:
    out_dir.mkdir(parents=True, exist_ok=True)
    safe = create_safe_fixture(out_dir)
    risky = create_risky_fixture(out_dir)
    return {
        "safe": safe,
        "risky": risky,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out-dir", required=True, type=Path, help="Directory for generated fixture query folders")
    parser.add_argument("--out-json", type=Path, help="Optional manifest path")
    args = parser.parse_args()
    manifest = create_fixture(args.out_dir.expanduser().resolve())
    if args.out_json:
        write_json(args.out_json.expanduser().resolve(), manifest)
    else:
        print(json.dumps(manifest, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
