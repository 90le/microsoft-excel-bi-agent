# Power BI And Excel Host Boundary

This skill is a portability and review skill unless a supported Power BI model source is supplied.

## Supported By This Plugin

- Review DAX portability between Excel Power Pivot and Power BI semantic models.
- Flag host-specific functions or assumptions.
- Review exported metadata, TMDL-like source, or documented model summaries when provided.
- Route Excel workbook Data Model work back to `power-pivot-dax-modeling`.

## Not Claimed By Default

- Decode or rewrite PBIX binaries.
- Deploy semantic models to the Power BI service.
- Refresh Power BI datasets.
- Validate Power BI report visuals.
- Modify XMLA endpoints without explicit connection details and user intent.

## Practical Rule

When the deliverable is an Excel workbook, use the stricter Excel Power Pivot compatibility path. When the deliverable is Power BI, verify current product behavior against official Microsoft/Fabric documentation and supported tooling.
