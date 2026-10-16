## 2024-05-24 - [PowerShell Delete Optimization]
**Learning:** Optimizing `Get-ChildItem | Remove-Item` pipelines by assigning to a variable and bulk-removing (`Remove-Item -Path $files.FullName`) causes an error if no files match the filter, because `$null` is passed to the path parameter.
**Action:** Always wrap bulk removal optimizations in a `$null -ne $files` conditional check, or revert to the pipeline which handles empty sets safely if edge cases are complex.
## 2024-05-24 - [PowerShell Delete Optimization Failed]
**Learning:** Replacing `Get-ChildItem | Remove-Item` pipeline binding with `$files = Get-ChildItem` and `Remove-Item -Path $files.FullName` breaks on valid filenames with wildcards like `[` and `]`. Pipeline binding safely uses `-LiteralPath` behind the scenes, whereas `-Path` evaluates wildcards. Furthermore, file deletion is highly I/O bound, so bypassing the pipeline for this operation is a micro-optimization with negligible impact.
**Action:** Do not bypass PowerShell pipelines for I/O bound operations. If you must pass paths explicitly, always use `-LiteralPath` instead of `-Path` to handle wildcard characters in filenames correctly.
