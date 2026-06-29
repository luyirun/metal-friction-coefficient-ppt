$ErrorActionPreference = "Stop"
Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint
Add-Type -AssemblyName Office

$msoTrue = -1
$msoFalse = 0
$msoShapeRectangle = 1
$msoAnimEffectFade = [Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectFade
$msoAnimEffectWipe = [Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectWipe
$msoAnimateLevelNone = [Microsoft.Office.Interop.PowerPoint.MsoAnimateByLevel]::msoAnimateLevelNone
$msoTriggerOnClick = [Microsoft.Office.Interop.PowerPoint.MsoAnimTriggerType]::msoAnimTriggerOnShapeClick
$msoTriggerWithPrevious = [Microsoft.Office.Interop.PowerPoint.MsoAnimTriggerType]::msoTriggerWithPrevious
$msoDirectionRight = [Microsoft.Office.Interop.PowerPoint.MsoAnimDirection]::msoAnimDirectionRight
$ppSaveAsOpenXMLPresentation = 24

$root = "F:\桌面\Keynote缩略图\因素矩阵"
$outDir = Join-Path $root "editable_static_overview"
$strictDir = Join-Path $root "gorden_detail_pages_strict"
$basePath = Join-Path $outDir "factor_matrix_static_editable_overview.pptx"
$outPath = Join-Path $outDir "factor_matrix_static_editable_overview_two_factors_f04_f05_gorden_static_pages_test.pptx"
$previewDir = Join-Path $outDir "two_factors_f04_f05_gorden_static_pages_test_previews"
$notesPath = Join-Path $outDir "factor_matrix_two_factors_f04_f05_gorden_static_pages_test_notes.md"

$factors = @(
    @{ Id = "f04"; Code = "F04"; Name = "roughness"; Pptx = (Join-Path $strictDir "gorden_detail_04_roughness_f04.pptx"); Trigger = @(42, 378, 142, 136); Return = @(1294, 816, 160, 44) },
    @{ Id = "f05"; Code = "F05"; Name = "contact_area"; Pptx = (Join-Path $strictDir "gorden_detail_05_contact_area_f05.pptx"); Trigger = @(194, 378, 142, 136); Return = @(1294, 816, 160, 44) }
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

function Add-TriggeredEffect($sequence, $shape, $triggerShape, $effectType, [bool]$exit, [double]$duration) {
    $effect = $sequence.AddEffect($shape, $effectType, $msoAnimateLevelNone, $msoTriggerOnClick)
    $effect.Timing.TriggerShape = $triggerShape
    $effect.Timing.Duration = $duration
    if ($effectType -eq $msoAnimEffectWipe) {
        try { $effect.EffectParameters.Direction = $msoDirectionRight } catch {}
    }
    if ($exit) { $effect.Exit = $msoTrue }
    return $effect
}

function Add-WithPreviousEffect($sequence, $shape, $effectType, [bool]$exit, [double]$duration) {
    $effect = $sequence.AddEffect($shape, $effectType, $msoAnimateLevelNone, $msoTriggerWithPrevious)
    $effect.Timing.Duration = $duration
    if ($effectType -eq $msoAnimEffectWipe) {
        try { $effect.EffectParameters.Direction = $msoDirectionRight } catch {}
    }
    if ($exit) { $effect.Exit = $msoTrue }
    return $effect
}

function Export-StrictPagePreview($app, [string]$pptx, [string]$outPng) {
    if (-not (Test-Path -LiteralPath $pptx)) { throw "Missing strict static page: $pptx" }
    $pres = $null
    try {
        $pres = $app.Presentations.Open($pptx, $msoFalse, $msoFalse, $msoFalse)
        $pres.Slides.Item(1).Export($outPng, "PNG", 1536, 1024)
    }
    finally {
        if ($pres -ne $null) { try { $pres.Close() } catch {} }
    }
}

function Assert-StrictEvidence([string]$pptx) {
    $pageDir = Join-Path (Split-Path -Parent $pptx) ([System.IO.Path]::GetFileNameWithoutExtension($pptx))
    $manifest = Join-Path $pageDir "imagegen-assets-manifest.json"
    $layout = Join-Path $pageDir "editable\01\layout.json"
    $qa = Join-Path $pageDir "editable\01\qa-visual\report.json"
    foreach ($path in @($manifest, $layout, $qa)) {
        if (-not (Test-Path -LiteralPath $path)) { throw "Missing strict evidence: $path" }
    }
    $raw = Get-Content -LiteralPath $manifest -Raw -Encoding UTF8
    if ($raw -notmatch "imagegen") { throw "Manifest lacks imagegen evidence: $manifest" }
    if ($raw -match "programmatic|local layer generator|PIL|SVG|HTML|Canvas|matplotlib|screenshot renderer|null prompt_file") {
        throw "Manifest contains blocked source token: $manifest"
    }
}

function Get-OverviewNames($slide) {
    $names = @()
    for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
        $n = $slide.Shapes.Item($i).Name
        if ($n -like "overview*" -or $n -like "flow*" -or $n -like "grp_*" -or $n -like "radar*" -or $n -like "conclusion*" -or $n -eq "group_title") { $names += $n }
    }
    return $names
}

New-Item -ItemType Directory -Force -Path $previewDir | Out-Null
if (-not (Test-Path -LiteralPath $basePath)) { throw "Base overview PPT not found: $basePath" }
if (Test-Path -LiteralPath $outPath) { Remove-Item -LiteralPath $outPath -Force }
Copy-Item -LiteralPath $basePath -Destination $outPath -Force

$app = New-Object -ComObject PowerPoint.Application
$app.Visible = $msoTrue
$presentation = $null
try {
    foreach ($factor in $factors) { Assert-StrictEvidence $factor.Pptx }
    foreach ($factor in $factors) {
        $factor.Preview = Join-Path $previewDir ("source_" + $factor.Id + "_powerpoint_preview.png")
        Export-StrictPagePreview $app $factor.Pptx $factor.Preview
    }

    $presentation = $app.Presentations.Open($outPath, $msoFalse, $msoFalse, $msoTrue)
    $slide = $presentation.Slides.Item(1)
    $slide.Name = "two_factors_f04_f05_gorden_static_pages_test"
    $slide.SlideShowTransition.AdvanceOnClick = $msoFalse
    $slide.SlideShowTransition.AdvanceOnTime = $msoFalse

    $overviewNames = @(Get-OverviewNames $slide)
    if ($overviewNames.Count -lt 2) { throw "No overview objects found to group." }
    $overviewGroup = $slide.Shapes.Range($overviewNames).Group()
    $overviewGroup.Name = "anim_overview_group"

    $detailGroups = @{}
    $factorTriggers = @{}
    $returnTriggers = @{}

    foreach ($factor in $factors) {
        $pic = $slide.Shapes.AddPicture($factor.Preview, $msoFalse, $msoTrue, 0, 0, 960, 640)
        $pic.Name = "anim_$($factor.Id)_detail_group"
        $pic.Visible = $msoFalse
        $detailGroups[$factor.Id] = $pic
        $factorTriggers[$factor.Id] = Add-HitArea $slide "trigger_$($factor.Id)_card" $factor.Trigger[0] $factor.Trigger[1] $factor.Trigger[2] $factor.Trigger[3]
        $returnTriggers[$factor.Id] = Add-HitArea $slide "trigger_$($factor.Id)_return" $factor.Return[0] $factor.Return[1] $factor.Return[2] $factor.Return[3]
        $returnTriggers[$factor.Id].Visible = $msoFalse
    }

    foreach ($factor in $factors) {
        $seq = $slide.TimeLine.InteractiveSequences.Add()
        [void](Add-TriggeredEffect $seq $overviewGroup $factorTriggers[$factor.Id] $msoAnimEffectFade $true 0.12)
        foreach ($other in $factors) {
            if ($other.Id -ne $factor.Id) {
                [void](Add-WithPreviousEffect $seq $detailGroups[$other.Id] $msoAnimEffectFade $true 0.08)
            }
        }
        [void](Add-WithPreviousEffect $seq $detailGroups[$factor.Id] $msoAnimEffectWipe $false 0.24)

        $seqReturn = $slide.TimeLine.InteractiveSequences.Add()
        [void](Add-TriggeredEffect $seqReturn $detailGroups[$factor.Id] $returnTriggers[$factor.Id] $msoAnimEffectFade $true 0.12)
        [void](Add-WithPreviousEffect $seqReturn $overviewGroup $msoAnimEffectWipe $false 0.24)
    }

    $presentation.SaveAs($outPath, $ppSaveAsOpenXMLPresentation)

    foreach ($shape in $detailGroups.Values) { $shape.Visible = $msoFalse }
    foreach ($shape in $returnTriggers.Values) { $shape.Visible = $msoFalse }
    $overviewGroup.Visible = $msoTrue
    $slide.Export((Join-Path $previewDir "state_00_overview_initial.png"), "PNG", 1536, 1024)

    foreach ($factor in $factors) {
        $overviewGroup.Visible = $msoFalse
        foreach ($shape in $detailGroups.Values) { $shape.Visible = $msoFalse }
        foreach ($shape in $returnTriggers.Values) { $shape.Visible = $msoFalse }
        $detailGroups[$factor.Id].Visible = $msoTrue
        $returnTriggers[$factor.Id].Visible = $msoTrue
        $slide.Export((Join-Path $previewDir ("state_" + $factor.Id + "_detail.png")), "PNG", 1536, 1024)
    }

    foreach ($shape in $detailGroups.Values) { $shape.Visible = $msoFalse }
    foreach ($shape in $returnTriggers.Values) { $shape.Visible = $msoFalse }
    $overviewGroup.Visible = $msoTrue
    $presentation.SaveAs($outPath, $ppSaveAsOpenXMLPresentation)

    @(
        "# Two-factor Gorden static-page animation test",
        "",
        "Output PPT: $outPath",
        "",
        "Static detail pages used:",
        "- F04: $($factors[0].Pptx)",
        "- F05: $($factors[1].Pptx)",
        "",
        "Logic:",
        "1. Click F04 card: overview fades out, F04 strict static page preview wipes in.",
        "2. Click F05 card: overview fades out, F05 strict static page preview wipes in; other detail state fades out if visible.",
        "3. Click return button inside either detail state: detail fades out, overview wipes back in.",
        "",
        "This test intentionally uses the same strict static page paths as the final interaction source of truth."
    ) | Set-Content -LiteralPath $notesPath -Encoding UTF8

    Write-Output "Generated=$outPath"
    Write-Output "PreviewDir=$previewDir"
    Write-Output "Notes=$notesPath"
}
finally {
    if ($presentation -ne $null) { try { $presentation.Close() } catch {} }
    try { $app.Quit() } catch {}
}
