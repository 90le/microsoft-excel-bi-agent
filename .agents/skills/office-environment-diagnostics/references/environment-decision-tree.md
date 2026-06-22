# Environment Decision Tree

Use this when an Excel BI automation task fails before workbook logic can be tested.

## First Split

1. **Excel cannot start through COM**
   - Check desktop Excel installation, licensing, process conflicts, and automation policy.
   - Do not debug workbook formulas yet.

2. **VBA project access fails**
   - Check Trust Center setting for trusted access to the VBA project object model.
   - Confirm the workbook format is macro-capable when importing modules.

3. **ADO/OLEDB workbook SQL fails**
   - Check ACE provider availability and Office/provider bitness.
   - Run the generated-workbook ADO smoke before blaming the customer file.

4. **Power Pivot or ADOMD fails**
   - Check MSOLAP/ADOMD COM activation separately.
   - Treat endpoint query execution as explicit-user-input only.

5. **Power Query refresh fails**
   - Separate connector/provider availability from credentials, privacy levels, gateway ownership, and source reachability.

## Evidence Order

Run provider probe first, then build the provider environment report, then compare against a known baseline if available.
