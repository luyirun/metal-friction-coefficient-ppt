$ErrorActionPreference = "Stop"

$pptx = Join-Path (Get-Location) "projects\slide09_whitepaper\exports\slide09_whitepaper_page09_interactive.pptx"

$msoAnimEffectFade = 10
$msoAnimateLevelNone = 0
$msoAnimTriggerAfterPrevious = 3
$msoAnimTriggerOnShapeClick = 4
$msoTrue = -1

function Get-ShapesByPrefix($slide, [string]$prefix) {
    $items = @()
    for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
        $shape = $slide.Shapes.Item($i)
        if ($shape.Name -like "$prefix*") {
            $items += $shape
        }
    }
    return $items
}

function Get-ShapesByNamesAndPrefixes($slide, [string[]]$prefixes) {
    $items = @()
    for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
        $shape = $slide.Shapes.Item($i)
        foreach ($prefix in $prefixes) {
            if ($shape.Name -like "$prefix*") {
                $items += $shape
                break
            }
        }
    }
    return $items
}

function Group-Prefix($slide, [string]$prefix, [string]$groupName) {
    $names = @()
    for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
        $shape = $slide.Shapes.Item($i)
        if ($shape.Name -like "$prefix*") {
            $names += $shape.Name
        }
    }
    if ($names.Count -gt 1) {
        $range = $slide.Shapes.Range($names)
        $group = $range.Group()
        $group.Name = $groupName
        return $group
    }
    if ($names.Count -eq 1) {
        $shape = $slide.Shapes.Item($names[0])
        $shape.Name = $groupName
        return $shape
    }
    throw "No shapes found for $prefix"
}

function Add-TriggeredFadeIn($interactiveSeq, $shape, $triggerShape, [double]$duration) {
    $effect = $interactiveSeq.AddEffect($shape, $msoAnimEffectFade, $msoAnimateLevelNone, $msoAnimTriggerOnShapeClick)
    $effect.Timing.TriggerShape = $triggerShape
    $effect.Timing.Duration = $duration
    return $effect
}

function Add-TriggeredFadeOut($interactiveSeq, $shape, $triggerShape, [double]$duration) {
    $effect = $interactiveSeq.AddEffect($shape, $msoAnimEffectFade, $msoAnimateLevelNone, $msoAnimTriggerOnShapeClick)
    $effect.Exit = $msoTrue
    $effect.Timing.TriggerShape = $triggerShape
    $effect.Timing.Duration = $duration
    return $effect
}

function Add-AutoFadeIn($seq, $shape, [double]$delay, [double]$duration) {
    $effect = $seq.AddEffect($shape, $msoAnimEffectFade, $msoAnimateLevelNone, $msoAnimTriggerAfterPrevious)
    $effect.Timing.TriggerDelayTime = $delay
    $effect.Timing.Duration = $duration
    return $effect
}

$app = New-Object -ComObject PowerPoint.Application
$app.Visible = $msoTrue
$presentation = $null
try {
    $presentation = $app.Presentations.Open($pptx, $false, $false, $false)
    $slide = $presentation.Slides.Item(1)
    $seq = $slide.TimeLine.MainSequence

    $keys = @("temperature", "pressure", "material", "roughness", "area", "ice")
    $triggerMap = @{}
    foreach ($key in $keys) {
        $triggerMap[$key] = Group-Prefix $slide "factor_node_$key" "factor_node_$key"
    }

    $entryPrefixes = @(
        "slide09_title_index",
        "slide09_title_main",
        "slide09_title_sub",
        "slide09_factor_matrix",
        "slide09_factor_matrix_title",
        "slide09_factor_matrix_hint",
        "slide09_experiment_model_overview"
    )
    $delay = 0.05
    foreach ($prefix in $entryPrefixes) {
        foreach ($shape in (Get-ShapesByPrefix $slide $prefix)) {
            [void](Add-AutoFadeIn $seq $shape $delay 0.25)
        }
        $delay += 0.05
    }

    $overviewShapes = Get-ShapesByNamesAndPrefixes $slide @("slide09_experiment_model_overview", "overview_")
    $allAnalysis = @{}
    foreach ($key in $keys) {
        $allAnalysis[$key] = Get-ShapesByNamesAndPrefixes $slide @("analysis_group_$key", "table_group_$key", "chart_group_$key", "conclusion_$key")
    }

    foreach ($key in $keys) {
        $trigger = $triggerMap[$key]
        $interactiveSeq = $slide.TimeLine.InteractiveSequences.Add()
        foreach ($shape in $overviewShapes) {
            [void](Add-TriggeredFadeOut $interactiveSeq $shape $trigger 0.16)
        }
        foreach ($other in $keys) {
            if ($other -eq $key) { continue }
            foreach ($shape in $allAnalysis[$other]) {
                [void](Add-TriggeredFadeOut $interactiveSeq $shape $trigger 0.12)
            }
        }
        foreach ($shape in $allAnalysis[$key]) {
            [void](Add-TriggeredFadeIn $interactiveSeq $shape $trigger 0.22)
        }
    }

    $presentation.Save()
    Write-Output "Added slide09 trigger animations to $pptx"
}
finally {
    if ($presentation -ne $null) { $presentation.Close() }
    $app.Quit()
}
