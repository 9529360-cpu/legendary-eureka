# Cleanup App.tsx
$appPath = "c:\Users\M1133\excel-copilot-addin\src\taskpane\components\App.tsx"

$c = Get-Content $appPath -Encoding UTF8
Write-Host "Original lines: $($c.Length)"

# Keep first 369 lines
$before = $c[0..368]

# Keep from line 975
$after = $c[974..($c.Length - 1)]

# Create replacement comment
$comment = @(
  "",
  "  // v2.9.12: Moved to separate modules",
  "  // - parseFormulaReferences, analyzeFormulaComplexity: utils/dataAnalysis.ts",
  "  // - scanWorkbook, verifyOperationResult: services/ExcelScanner.ts",
  "  // - generateDataSummary, generateProactiveSuggestions: utils/dataAnalysis.ts",
  "  // - Workbook scan useEffect: useWorkbookContext hook",
  ""
)

# Merge content
$newContent = $before + $comment + $after

# Write file
$newContent | Set-Content $appPath -Encoding UTF8

Write-Host "Cleanup done!"
Write-Host "New lines: $($newContent.Count)"
Write-Host "Deleted: $($c.Length - $newContent.Count) lines"
