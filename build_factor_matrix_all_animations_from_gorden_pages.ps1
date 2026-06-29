param(
    [string]$Root = "",
    [string]$GordenDir = "",
    [string]$OutDir = "",
    [switch]$AllowInspiredFallback
)

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
$msoTriggerWithPrevious = [Microsoft.Office.Interop.PowerPoint.MsoAnimTriggerType]::msoAnimTriggerWithPrevious
$msoDirectionRight = [Microsoft.Office.Interop.PowerPoint.MsoAnimDirection]::msoAnimDirectionRight

function U([string]$base64) {
    return [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($base64))
}

if ([string]::IsNullOrWhiteSpace($Root)) {
    $Root = U "Rjpc5qGM6Z2iXEtleW5vdGXnvKnnlaXlm75c5Zug57Sg55+p6Zi1"
}

if ([string]::IsNullOrWhiteSpace($GordenDir)) {
    $GordenDir = Join-Path $Root "gorden_detail_pages_strict"
}
if ([string]::IsNullOrWhiteSpace($OutDir)) {
    $OutDir = Join-Path $Root "editable_static_overview"
}

$basePath = Join-Path $OutDir "factor_matrix_static_editable_overview.pptx"
$outPath = Join-Path $OutDir "factor_matrix_static_editable_overview_all_factors_animation_v02_gorden_layers.pptx"
$previewDir = Join-Path $OutDir "all_factors_animation_v02_gorden_previews"
$notesPath = Join-Path $OutDir "factor_matrix_all_factors_animation_v02_gorden_notes.md"

$factors = @(
    @{ Id = "f01"; Code = "F01"; Pattern = "*01_temperature_f01*.pptx"; Trigger = @(42, 218, 142, 136) },
    @{ Id = "f02"; Code = "F02"; Pattern = "*02_pressure_f02*.pptx"; Trigger = @(194, 218, 142, 136) },
    @{ Id = "f03"; Code = "F03"; Pattern = "*03_material_f03*.pptx"; Trigger = @(346, 218, 142, 136) },
    @{ Id = "f04"; Code = "F04"; Pattern = "*04_roughness_f04*.pptx"; Trigger = @(42, 378, 142, 136) },
    @{ Id = "f05"; Code = "F05"; Pattern = "*05_contact_area_f05*.pptx"; Trigger = @(194, 378, 142, 136) },
    @{ Id = "f06"; Code = "F06"; Pattern = "*06_ice_composition_f06*.pptx"; Trigger = @(346, 378, 142, 136) }
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

function Get-OverviewNames($slide) {
    $names = @()
    for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
        $n = $slide.Shapes.Item($i).Name
        if (
            $n -like "overview*" -or
            $n -like "flow*" -or
            $n -like "grp_*" -or
            $n -like "radar*" -or
            $n -like "conclusion*" -or
            $n -eq "group_title"
        ) {
            $names += $n
        }
    }
    return $names
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

function Assert-GordenStrictInput([string]$dir) {
    if (-not (Test-Path -LiteralPath $dir)) {
        throw "Gorden detail directory not found: $dir"
    }

    $manifestFiles = @(Get-ChildItem -LiteralPath $dir -Recurse -Filter "imagegen-assets-manifest.json" -File -ErrorAction SilentlyContinue)
    if ($manifestFiles.Count -eq 0) {
        if ($AllowInspiredFallback) { return }
        throw "Strict Gorden manifest not found under $dir. Do not use gorden_static_single_pages unless explicitly allowing fallback."
    }

    foreach ($manifestFile in $manifestFiles) {
        $raw = Get-Content -LiteralPath $manifestFile.FullName -Raw
        if ($raw -match "programmatic|local layer generator|PIL|SVG|HTML|Canvas|matplotlib|screenshot renderer|null prompt_file") {
            throw "Non-Gorden image layer source found in $($manifestFile.FullName). Rebuild that page with GordenImage2PPTX."
        }
        if ($raw -notmatch "imagegen") {
            throw "Manifest does not show imagegen evidence: $($manifestFile.FullName)"
        }
    }
}

function Resolve-DetailPptx([string]$dir, [string]$pattern) {
    $matches = @(Get-ChildItem -LiteralPath $dir -Recurse -Filter $pattern -File | Where-Object { $_.Extension -ieq ".pptx" } | Sort-Object FullName)
    if ($matches.Count -eq 0) {
        throw "Missing detail PPTX matching $pattern under $dir"
    }
    return $matches[0].FullName
}

function Import-FullSlideAsGroup($app, $targetSlide, [string]$sourcePath, [string]$groupName) {
    $presentation = $targetSlide.Parent
    $insertAfter = $presentation.Slides.Count
    [void]$presentation.Slides.InsertFromFile($sourcePath, $insertAfter, 1, 1)
    $importedSlide = $presentation.Slides.Item($insertAfter + 1)

    try {
        $sourceIndexes = @()
        for ($i = 1; $i -le $importedSlide.Shapes.Count; $i++) {
            $sourceIndexes += $i
        }
        if ($sourceIndexes.Count -eq 0) {
            throw "No shapes were imported from $sourcePath"
        }

        $countBefore = $targetSlide.Shapes.Count
        $importedSlide.Shapes.Range($sourceIndexes).Cut()
        Start-Sleep -Milliseconds 250
        [void]$targetSlide.Shapes.Paste()

        $newIndexes = @()
        for ($i = $countBefore + 1; $i -le $targetSlide.Shapes.Count; $i++) {
            $newIndexes += $i
        }
        if ($newIndexes.Count -eq 0) {
            throw "No shapes were moved from imported slide for $sourcePath"
        }

        $group = $targetSlide.Shapes.Range($newIndexes).Group()
        $group.Name = $groupName
        return $group
    }
    finally {
        try { $importedSlide.Delete() } catch {}
    }
}

Assert-GordenStrictInput $GordenDir
New-Item -ItemType Directory -Force -Path $previewDir | Out-Null

if (-not (Test-Path -LiteralPath $basePath)) {
    throw "Base overview PPT not found: $basePath"
}
if (Test-Path -LiteralPath $outPath) {
    Remove-Item -LiteralPath $outPath -Force
}
Copy-Item -LiteralPath $basePath -Destination $outPath -Force

$app = New-Object -ComObject PowerPoint.Application
$app.Visible = $msoTrue
$presentation = $null

try {
    $presentation = $app.Presentations.Open($outPath, $msoFalse, $msoFalse, $msoTrue)
    $slide = $presentation.Slides.Item(1)
    $slide.Name = "all_factors_animation_from_gorden_pages"
    $slide.SlideShowTransition.AdvanceOnClick = $msoFalse
    $slide.SlideShowTransition.AdvanceOnTime = $msoFalse

    $overviewNames = @(Get-OverviewNames $slide)
    if ($overviewNames.Count -eq 0) {
        throw "No overview objects found in base PPT."
    }
    $overviewGroup = $slide.Shapes.Range($overviewNames).Group()
    $overviewGroup.Name = "anim_overview_group"

    $detailGroups = @{}
    $factorTriggers = @{}
    $returnTriggers = @{}

    foreach ($factor in $factors) {
        $detailPath = Resolve-DetailPptx $GordenDir $factor.Pattern
        $detailGroups[$factor.Id] = Import-FullSlideAsGroup $app $slide $detailPath "anim_$($factor.Id)_detail_group"
        $factorTriggers[$factor.Id] = Add-HitArea $slide "trigger_factor_$($factor.Id)" $factor.Trigger[0] $factor.Trigger[1] $factor.Trigger[2] $factor.Trigger[3]
        $returnTriggers[$factor.Id] = Add-HitArea $slide "trigger_back_$($factor.Id)" 804 532 106 26
    }

    foreach ($factor in $factors) {
        $seqShow = $slide.TimeLine.InteractiveSequences.Add()
        [void](Add-TriggeredEffect $seqShow $overviewGroup $factorTriggers[$factor.Id] $msoAnimEffectFade $true 0.18)
        [void](Add-WithPreviousEffect $seqShow $detailGroups[$factor.Id] $msoAnimEffectWipe $false 0.32)
        [void](Add-WithPreviousEffect $seqShow $returnTriggers[$factor.Id] $msoAnimEffectFade $false 0.01)

        $seqBack = $slide.TimeLine.InteractiveSequences.Add()
        [void](Add-TriggeredEffect $seqBack $detailGroups[$factor.Id] $returnTriggers[$factor.Id] $msoAnimEffectFade $true 0.18)
        [void](Add-WithPreviousEffect $seqBack $returnTriggers[$factor.Id] $msoAnimEffectFade $true 0.01)
        [void](Add-WithPreviousEffect $seqBack $overviewGroup $msoAnimEffectWipe $false 0.32)
    }

    foreach ($factor in $factors) {
        $detailGroups[$factor.Id].Visible = $msoFalse
        $returnTriggers[$factor.Id].Visible = $msoFalse
    }
    $overviewGroup.Visible = $msoTrue
    $slide.Export((Join-Path $previewDir "state_00_overview_initial.png"), "PNG", 1536, 1024)

    foreach ($factor in $factors) {
        $overviewGroup.Visible = $msoFalse
        foreach ($f in $factors) {
            $detailGroups[$f.Id].Visible = $msoFalse
            $returnTriggers[$f.Id].Visible = $msoFalse
        }
        $detailGroups[$factor.Id].Visible = $msoTrue
        $returnTriggers[$factor.Id].Visible = $msoTrue
        $slide.Export((Join-Path $previewDir "state_$($factor.Id)_detail.png"), "PNG", 1536, 1024)
    }

    $overviewGroup.Visible = $msoTrue
    foreach ($factor in $factors) {
        $detailGroups[$factor.Id].Visible = $msoTrue
        $returnTriggers[$factor.Id].Visible = $msoTrue
    }
    $presentation.Save()

    @(
        "# All factors animation from strict Gorden pages",
        "",
        "Base overview PPT: $basePath",
        "Strict Gorden detail dir: $GordenDir",
        "Output PPT: $outPath",
        "",
        "This script imports each full Gorden detail slide as one animation group.",
        "It intentionally refuses Gorden-inspired/programmatic pages unless -AllowInspiredFallback is passed."
    ) | Set-Content -LiteralPath $notesPath -Encoding UTF8

    Write-Output "Generated=$outPath"
    Write-Output "PreviewDir=$previewDir"
    Write-Output "Notes=$notesPath"
}
finally {
    if ($presentation -ne $null) { try { $presentation.Close() } catch {} }
    try { $app.Quit() } catch {}
}
