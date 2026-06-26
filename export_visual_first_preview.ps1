$ErrorActionPreference = "Stop"

$Root = "F:\Users\user\Documents\测金属摩擦系数"
$Pptx = Join-Path $Root "outputs\metal_ice_99fid_visual_first.pptx"
$PreviewDir = Join-Path $Root "outputs\metal_ice_99fid_visual_first_preview"

if (Test-Path -LiteralPath $PreviewDir) {
    Remove-Item -LiteralPath $PreviewDir -Recurse -Force
}
New-Item -ItemType Directory -Path $PreviewDir | Out-Null

$app = New-Object -ComObject PowerPoint.Application
$app.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$pres = $app.Presentations.Open($Pptx, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
$pres.Export($PreviewDir, "PNG", 2048, 1152)
$pres.Close()
$app.Quit()

[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null
Write-Output $PreviewDir
