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
$msoTriggerAfterPrevious = [Microsoft.Office.Interop.PowerPoint.MsoAnimTriggerType]::msoAnimTriggerAfterPrevious
$msoDirectionRight = [Microsoft.Office.Interop.PowerPoint.MsoAnimDirection]::msoAnimDirectionRight

$outRoot = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String("Rjpc5qGM6Z2iXEtleW5vdGXnvKnnlaXlm75c5Zug57Sg55+p6Zi1XGdvcmRlbl9zdGF0aWNfc2luZ2xlX3BhZ2Vz"))
$overviewPath = Join-Path $outRoot "factor_matrix_static_00_overview.pptx"
$f01Path = Join-Path $outRoot "factor_matrix_static_01_temperature_f01.pptx"
$testPath = Join-Path $outRoot "factor_matrix_onepage_overview_f01_animation_test.pptx"
$previewPath = Join-Path $outRoot "factor_matrix_onepage_overview_f01_animation_test_preview.png"
$notesPath = Join-Path $outRoot "factor_matrix_onepage_animation_test_notes.md"

function Convert-HexColor([string]$hex) {
    $value = $hex.TrimStart("#")
    $r = [Convert]::ToInt32($value.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($value.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($value.Substring(4, 2), 16)
    return $r + ($g * 256) + ($b * 65536)
}

function Get-NamesForOverviewGroup($slide) {
    $names = @()
    for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
        $shape = $slide.Shapes.Item($i)
        if (($shape.Left -ge 500) -and ($shape.Top -lt 580)) {
            $names += $shape.Name
        }
    }
    return $names
}

function Get-NamesForF01DetailLayer($slide) {
    $names = @()
    for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
        $shape = $slide.Shapes.Item($i)
        $isRightDetail = (($shape.Left -ge 500) -and ($shape.Top -lt 580))
        $isActiveArrow = ($shape.Name -eq "active_arrow_temperature")
        $isActiveCardPart = (($shape.Left -ge 40) -and (($shape.Left + $shape.Width) -le 190) -and ($shape.Top -ge 215) -and (($shape.Top + $shape.Height) -le 355))
        if ($isRightDetail -or $isActiveArrow -or $isActiveCardPart) {
            $names += $shape.Name
        }
    }
    return $names
}

function Group-ByNames($slide, [string[]]$names, [string]$groupName) {
    if ($names.Count -lt 2) {
        throw "Need at least 2 shapes to group for $groupName"
    }
    $group = $slide.Shapes.Range($names).Group()
    $group.Name = $groupName
    return $group
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
    if ($exit) {
        $effect.Exit = $msoTrue
    }
    return $effect
}

function Add-WithPreviousEffect($sequence, $shape, $effectType, [bool]$exit, [double]$duration) {
    $effect = $sequence.AddEffect($shape, $effectType, $msoAnimateLevelNone, $msoTriggerWithPrevious)
    $effect.Timing.Duration = $duration
    if ($effectType -eq $msoAnimEffectWipe) {
        try { $effect.EffectParameters.Direction = $msoDirectionRight } catch {}
    }
    if ($exit) {
        $effect.Exit = $msoTrue
    }
    return $effect
}

if (Test-Path -LiteralPath $testPath) {
    Remove-Item -LiteralPath $testPath -Force
}
Copy-Item -LiteralPath $overviewPath -Destination $testPath -Force

$app = New-Object -ComObject PowerPoint.Application
$app.Visible = $msoTrue
$presentation = $null
$sourcePresentation = $null

try {
    $presentation = $app.Presentations.Open($testPath, $msoFalse, $msoFalse, $msoTrue)
    $sourcePresentation = $app.Presentations.Open($f01Path, $msoFalse, $msoFalse, $msoTrue)

    $slide = $presentation.Slides.Item(1)
    $sourceSlide = $sourcePresentation.Slides.Item(1)
    $slide.Name = "overview_f01_group_animation_test"
    $slide.SlideShowTransition.AdvanceOnClick = $msoFalse
    $slide.SlideShowTransition.AdvanceOnTime = $msoFalse

    $overviewNames = Get-NamesForOverviewGroup $slide
    $overviewGroup = Group-ByNames $slide $overviewNames "anim_overview_group"

    $detailNames = Get-NamesForF01DetailLayer $sourceSlide
    $sourceSlide.Shapes.Range($detailNames).Copy()
    $pastedRange = $slide.Shapes.Paste()
    $detailGroup = $pastedRange.Group()
    $detailGroup.Name = "anim_f01_detail_group"

    $triggerTemperature = Add-HitArea $slide "trigger_temperature_card_f01" 42 218 142 136
    $triggerBack = Add-HitArea $slide "trigger_back_to_overview" 804 532 106 26

    $triggerTemperature.ZOrder(0) | Out-Null
    $triggerBack.ZOrder(0) | Out-Null

    $mainSeq = $slide.TimeLine.MainSequence
    $initialHide = $mainSeq.AddEffect($detailGroup, $msoAnimEffectFade, $msoAnimateLevelNone, $msoTriggerAfterPrevious)
    $initialHide.Exit = $msoTrue
    $initialHide.Timing.Duration = 0.01

    $seqShowDetail = $slide.TimeLine.InteractiveSequences.Add()
    [void](Add-TriggeredEffect $seqShowDetail $overviewGroup $triggerTemperature $msoAnimEffectFade $true 0.15)
    [void](Add-WithPreviousEffect $seqShowDetail $detailGroup $msoAnimEffectWipe $false 0.30)

    $seqReturn = $slide.TimeLine.InteractiveSequences.Add()
    [void](Add-TriggeredEffect $seqReturn $detailGroup $triggerBack $msoAnimEffectFade $true 0.15)
    [void](Add-WithPreviousEffect $seqReturn $overviewGroup $msoAnimEffectWipe $false 0.30)

    $presentation.Save()
    $slide.Export($previewPath, "PNG", 1536, 1024)

    @(
        "# One-page animation test notes",
        "",
        "File: `factor_matrix_onepage_overview_f01_animation_test.pptx`",
        "",
        "## Groups",
        "",
        "- `anim_overview_group`: overview panel group.",
        "- `anim_f01_detail_group`: F01 detail panel group, including the active F01 card overlay and arrow.",
        "- `trigger_temperature_card_f01`: transparent click area over the F01 card.",
        "- `trigger_back_to_overview`: transparent click area over the return button.",
        "",
        "## Logic",
        "",
        "1. Initial slide show state shows overview.",
        "2. Click F01 card: overview fades out, F01 detail wipes in from left to right.",
        "3. Click return button: F01 detail fades out, overview wipes in from left to right.",
        "",
        "This test only contains overview + F01. Other factors are not merged yet."
    ) | Set-Content -LiteralPath $notesPath -Encoding UTF8

    Write-Output "Generated $testPath"
    Write-Output "Preview $previewPath"
    Write-Output "Notes $notesPath"
}
finally {
    if ($sourcePresentation -ne $null) { try { $sourcePresentation.Close() } catch {} }
    if ($presentation -ne $null) { try { $presentation.Close() } catch {} }
    try { $app.Quit() } catch {}
}
