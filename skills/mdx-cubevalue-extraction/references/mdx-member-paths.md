# MDX Member Path Patterns

Read this file when CUBE formulas return `#N/A` or need dimension/member references.

## Typical Unique Name Shapes

```text
[Measures].[Measure Name]
[Table].[Column].[All].[Member]
[Table].[Column].&[Key]
[Dimension].[Hierarchy].[Level].&[Key]
```

The exact path depends on the model metadata. Do not guess when a workbook can be inspected.

## Debug Checklist

1. Confirm connection name.
2. Confirm measure display name.
3. Confirm table/column/hierarchy names.
4. Confirm whether member keys or captions are used.
5. Test the member with `CUBEMEMBER` before using it in `CUBEVALUE`.
