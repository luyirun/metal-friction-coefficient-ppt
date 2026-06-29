$ErrorActionPreference = "Stop"
Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint
Add-Type -AssemblyName Office

$msoTrue = -1
$msoFalse = 0
$msoShapeRectangle = 1
$msoAnimEffectFade = [Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectFade
$msoAnimateLevelNone = [Microsoft.Office.Interop.PowerPoint.MsoAnimateByLevel]::msoAnimateLevelNone
$msoTriggerOnClick = [Microsoft.Office.Interop.PowerPoint.MsoAnimTriggerType]::msoAnimTriggerOnShapeClick
$ppSaveAsOpenXMLPresentation = 24

$root = "F:\桌面\Keynote缩略图\因素矩阵"
$outDir = Join-Path $root "editable_static_overview"
$gordenDir = Join-Path $root "gorden_detail_pages_strict"
$basePath = Join-Path $outDir "factor_matrix_static_editable_overview.pptx"
$outPath = Join-Path $outDir "factor_matrix_static_editable_overview_all_factors_animation_v03_gorden_preview_layers.pptx"
$previewDir = Join-Path $outDir "all_factors_animation_v03_gorden_preview_layers_previews"
$notesPath = Join-Path $outDir "factor_matrix_all_factors_animation_v03_gorden_preview_layers_notes.md"

$factors = @(
    @{ Id = "f01"; Code = "F01"; Pattern = "gorden_detail_01_temperature_f01"; Trigger = @(42, 218, 142, 136); Return = @(1294, 816, 160, 44) },
    @{ Id = "f02"; Code = "F02"; Pattern = "gorden_detail_02_pressure_f02"; Trigger = @(194, 218, 142, 136); Return = @(1294, 816, 160, 44) },
    @{ Id = "f03"; Code = "F03"; Pattern = "gorden_detail_03_material_f03"; Trigger = @(346, 218, 142, 136); Return = @(1294, 816, 160, 44) },
    @{ Id = "f04"; Code = "F04"; Pattern = "gorden_detail_04_roughness_f04"; Trigger = @(42, 378, 142, 136); Return = @(1294, 816, 160, 44) },
    @{ Id = "f05"; Code = "F05"; Pattern = "gorden_detail_05_contact_area_f05"; Trigger = @(194, 378, 142, 136); Return = @(1294, 816, 160, 44) },
    @{ Id = "f06"; Code = "F06"; Pattern = "gorden_detail_06_ice_composition_f06"; Trigger = @(346, 378, 142, 136); Return = @(1294, 816, 160, 44) }
)

function Convert-HexColor([string]$hex) {
    $value = $hex.TrimStart("#")
    $r = [Convert]::ToInt32($value.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($value.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($value.Substring(4, 2), 16)
    return $r + ($g * 256) + ($b * 65536)
}

function Add-HitArea($slide, [string]$name, [double]$left, [double]$top, [double]$width, [double]$height) {
    $shape = $slide.Shapes.AddShape($msoShapeRectangle, $left, $top, $width, $height)
    $shape.Name = $name
    $shape.Fill.Visible = $msoTrue
    $shape.Fill.ForeColor.RGB = Convert-HexColor "#FFFFFF"
    $shape.Fill.Transparency = 0.99
    $shape.Line.Visible = $msoFalse
    return $shape
}

function Add-TriggeredFade($sequence, $shape, $triggerShape, [bool]$exit, [double]$duration) {
    $effect = $sequence.AddEffect($shape, $msoAnimEffectFade, $msoAnimateLevelNone, $msoTriggerOnClick)
    $effect.Timing.TriggerShape = $triggerShape
    $effect.Timing.Duration = $duration
    if ($exit) { $effect.Exit = $msoTrue }
    return $effect
}

function Assert-StrictPage($pageDir) {
    $manifest = Join-Path $pageDir "imagegen-assets-manifest.json"
    $preview = Join-Path $pageDir "out\powerpoint_preview\slide_01.png"
    $layout = Join-Path $pageDir "editable\01\layout.json"
    $qa = Join-Path $pageDir "editable\01\qa-visual\report.json"
    foreach ($path in @($manifest, $preview, $layout, $qa)) {
        if (-not (Test-Path -LiteralPath $path)) { throw "Missing strict evidence: $path" }
    }
    $raw = Get-Content -LiteralPath $manifest -Raw -Encoding UTF8
    if ($raw -match "programmatic|local layer generator|PIL|SVG|HTML|Canvas|matplotlib|screenshot renderer|null prompt_file") {
        throw "Non-Gorden image layer source found in $manifest"
    }
    if ($raw -notmatch "imagegen") { throw "Manifest lacks imagegen evidence: $manifest" }
    return $preview
}

New-Item -ItemType Directory -Force -Path $previewDir | Out-Null
if (-not (Test-Path -LiteralPath $basePath)) { throw "Base overview PPT not found: $basePath" }
if (Test-Path -LiteralPath $outPath) { Remove-Item -LiteralPath $outPath -Force }
Copy-Item -LiteralPath $basePath -Destination $outPath -Force

$app = New-Object -ComObject PowerPoint.Application
$app.Visible = $msoTrue
$presentation = $null
try {
    $presentation = $app.Presentations.Open($outPath, $msoFalse, $msoFalse, $msoTrue)
    $slide = $presentation.Slides.Item(1)
    $slide.Name = "all_factors_animation_from_gorden_preview_layers"
    $slide.SlideShowTransition.AdvanceOnClick = $msoFalse
    $slide.SlideShowTransition.AdvanceOnTime = $msoFalse
    $seq = $slide.TimeLine.InteractiveSequences.Add()

    $detailShapes = @{}
    $factorTriggers = @{}
    $returnTriggers = @{}

    foreach ($factor in $factors) {
        $pageDir = Join-Path $gordenDir $factor.Pattern
        $preview = Assert-StrictPage $pageDir
        $detail = $slide.Shapes.AddPicture($preview, $msoFalse, $msoTrue, 0, 0, 960, 640)
        $detail.Name = "anim_$($factor.Id)_detail_preview_layer"
        $detail.Visible = $msoFalse
        $detailShapes[$factor.Id] = $detail
        $factorTriggers[$factor.Id] = Add-HitArea $slide "trigger_factor_$($factor.Id)" $factor.Trigger[0] $factor.Trigger[1] $factor.Trigger[2] $factor.Trigger[3]
        $returnTriggers[$factor.Id] = Add-HitArea $slide "trigger_return_$($factor.Id)" $factor.Return[0] $factor.Return[1] $factor.Return[2] $factor.Return[3]
        $returnTriggers[$factor.Id].Visible = $msoFalse
    }

    foreach ($factor in $factors) {
        Add-TriggeredFade $seq $detailShapes[$factor.Id] $factorTriggers[$factor.Id] $false 0.18 | Out-Null
        foreach ($other in $factors) {
            if ($other.Id -ne $factor.Id) {
                Add-TriggeredFade $seq $detailShapes[$other.Id] $factorTriggers[$factor.Id] $true 0.08 | Out-Null
            }
        }
        Add-TriggeredFade $seq $detailShapes[$factor.Id] $returnTriggers[$factor.Id] $true 0.18 | Out-Null
    }

    $presentation.SaveAs($outPath, $ppSaveAsOpenXMLPresentation)

    foreach ($shape in $detailShapes.Values) { $shape.Visible = $msoFalse }
    foreach ($shape in $returnTriggers.Values) { $shape.Visible = $msoFalse }
    $slide.Export((Join-Path $previewDir "state_00_overview_initial.png"), "PNG", 1536, 1024)

    foreach ($factor in $factors) {
        foreach ($shape in $detailShapes.Values) { $shape.Visible = $msoFalse }
        foreach ($shape in $returnTriggers.Values) { $shape.Visible = $msoFalse }
        $detailShapes[$factor.Id].Visible = $msoTrue
        $returnTriggers[$factor.Id].Visible = $msoTrue
        $slide.Export((Join-Path $previewDir "state_$($factor.Id)_detail.png"), "PNG", 1536, 1024)
    }

    foreach ($shape in $detailShapes.Values) { $shape.Visible = $msoFalse }
    foreach ($shape in $returnTriggers.Values) { $shape.Visible = $msoFalse }
    $presentation.SaveAs($outPath, $ppSaveAsOpenXMLPresentation)

    $notes = @(
        "# Factor Matrix Interactive Gorden Preview-Layer Build",
        "",
        "Output PPT: $outPath",
        "Strict Gorden detail dir: $gordenDir",
        "",
        "This one-slide interactive deck uses each strict Gorden page's real PowerPoint preview as a full-slide detail state layer.",
        "The source F01-F06 strict PPTX files remain editable; this interactive shell is optimized for reliable click-through presentation and verified state previews.",
        "",
        "States exported:",
        "- state_00_overview_initial.png",
        ($factors | ForEach-Object { "- state_$($_.Id)_detail.png" })
    )
    $notes -join "`r`n" | Set-Content -LiteralPath $notesPath -Encoding UTF8
    Write-Output "Generated=$outPath"
    Write-Output "PreviewDir=$previewDir"
    Write-Output "Notes=$notesPath"
}
finally {
    if ($presentation -ne $null) { try { $presentation.Close() } catch {} }
    if ($app -ne $null) { try { $app.Quit() } catch {} }
}
