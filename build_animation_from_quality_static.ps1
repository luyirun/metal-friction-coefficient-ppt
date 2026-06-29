$ErrorActionPreference = "Stop"

Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint
Add-Type -AssemblyName Office

$msoTrue = -1
$msoFalse = 0
$msoTextOrientationHorizontal = 1
$msoShapeRectangle = 1
$msoShapeRoundedRectangle = 5
$msoShapeOval = 9
$msoShapeRightArrow = 33
$msoAnimEffectFade = [Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectFade
$msoAnimEffectWipe = [Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectWipe
$msoAnimateLevelNone = [Microsoft.Office.Interop.PowerPoint.MsoAnimateByLevel]::msoAnimateLevelNone
$msoTriggerOnClick = [Microsoft.Office.Interop.PowerPoint.MsoAnimTriggerType]::msoAnimTriggerOnShapeClick
$msoTriggerWithPrevious = [Microsoft.Office.Interop.PowerPoint.MsoAnimTriggerType]::msoAnimTriggerWithPrevious
$msoTriggerAfterPrevious = [Microsoft.Office.Interop.PowerPoint.MsoAnimTriggerType]::msoAnimTriggerAfterPrevious
$msoDirectionRight = [Microsoft.Office.Interop.PowerPoint.MsoAnimDirection]::msoAnimDirectionRight

$outRoot = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String("Rjpc5qGM6Z2iXEtleW5vdGXnvKnnlaXlm75c5Zug57Sg55+p6Zi1XGVkaXRhYmxlX3N0YXRpY19vdmVydmlldw=="))
$basePath = Join-Path $outRoot "factor_matrix_static_editable_overview.pptx"
$outPath = Join-Path $outRoot "factor_matrix_static_editable_overview_f01_animation_test.pptx"
$previewPath = Join-Path $outRoot "factor_matrix_static_editable_overview_f01_animation_test_preview.png"
$initialStatePath = Join-Path $outRoot "factor_matrix_static_editable_overview_f01_state_initial.png"
$detailStatePath = Join-Path $outRoot "factor_matrix_static_editable_overview_f01_state_detail.png"
$notesPath = Join-Path $outRoot "factor_matrix_static_editable_overview_f01_animation_notes.md"

function U([string]$base64) {
    return [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($base64))
}

function Convert-HexColor([string]$hex) {
    $value = $hex.TrimStart("#")
    $r = [Convert]::ToInt32($value.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($value.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($value.Substring(4, 2), 16)
    return $r + ($g * 256) + ($b * 65536)
}

function Set-NoFill($shape) {
    $shape.Fill.Visible = $msoFalse
}

function Set-Line($shape, [string]$hex, [double]$weight, [double]$transparency = 0) {
    $shape.Line.Visible = $msoTrue
    $shape.Line.ForeColor.RGB = Convert-HexColor $hex
    $shape.Line.Weight = $weight
    $shape.Line.Transparency = $transparency
}

function Set-GlassFill($shape, [string]$hex = "#FFFFFF", [double]$transparency = 0.24) {
    $shape.Fill.Visible = $msoTrue
    $shape.Fill.ForeColor.RGB = Convert-HexColor $hex
    $shape.Fill.Transparency = $transparency
}

function Set-TextStyle($shape, [double]$fontSize, [string]$hex, [int]$bold = 0) {
    Set-NoFill $shape
    $shape.Line.Visible = $msoFalse
    $shape.TextFrame2.MarginLeft = 0
    $shape.TextFrame2.MarginRight = 0
    $shape.TextFrame2.MarginTop = 0
    $shape.TextFrame2.MarginBottom = 0
    $shape.TextFrame2.TextRange.Font.NameFarEast = "Microsoft YaHei"
    $shape.TextFrame2.TextRange.Font.Name = "Aptos"
    $shape.TextFrame2.TextRange.Font.Size = $fontSize
    $shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Convert-HexColor $hex
    if ($bold -ne 0) { $shape.TextFrame2.TextRange.Font.Bold = $msoTrue }
}

function Add-Text($slide, [string]$name, [string]$text, [double]$left, [double]$top, [double]$width, [double]$height, [double]$fontSize, [string]$hex, [int]$bold = 0) {
    $shape = $slide.Shapes.AddTextbox($msoTextOrientationHorizontal, $left, $top, $width, $height)
    $shape.Name = $name
    $shape.TextFrame2.TextRange.Text = $text
    Set-TextStyle $shape $fontSize $hex $bold
    return $shape
}

function Add-LineShape($slide, [string]$name, [double]$x1, [double]$y1, [double]$x2, [double]$y2, [string]$hex, [double]$weight, [double]$transparency = 0) {
    $line = $slide.Shapes.AddLine($x1, $y1, $x2, $y2)
    $line.Name = $name
    $line.Line.ForeColor.RGB = Convert-HexColor $hex
    $line.Line.Weight = $weight
    $line.Line.Transparency = $transparency
    return $line
}

function Add-Panel($slide, [string]$name, [double]$left, [double]$top, [double]$width, [double]$height, [string]$lineHex = "#5C95C9", [double]$lineWeight = 0.85, [double]$fillTransparency = 0.24) {
    $shape = $slide.Shapes.AddShape($msoShapeRoundedRectangle, $left, $top, $width, $height)
    $shape.Name = $name
    Set-GlassFill $shape "#FFFFFF" $fillTransparency
    Set-Line $shape $lineHex $lineWeight 0.02
    $shape.Shadow.Visible = $msoTrue
    $shape.Shadow.ForeColor.RGB = Convert-HexColor "#6F9FBD"
    $shape.Shadow.Transparency = 0.78
    $shape.Shadow.Blur = 8
    $shape.Shadow.OffsetX = 1
    $shape.Shadow.OffsetY = 1.5
    return $shape
}

function Add-HighlightPanel($slide, [string]$name, [double]$left, [double]$top, [double]$width, [double]$height) {
    $shape = $slide.Shapes.AddShape($msoShapeRoundedRectangle, $left, $top, $width, $height)
    $shape.Name = $name
    Set-NoFill $shape
    Set-Line $shape "#16A8FF" 1.55 0
    $shape.Shadow.Visible = $msoTrue
    $shape.Shadow.ForeColor.RGB = Convert-HexColor "#16A8FF"
    $shape.Shadow.Transparency = 0.55
    $shape.Shadow.Blur = 10
    $shape.Shadow.OffsetX = 0
    $shape.Shadow.OffsetY = 0
    return $shape
}

function Add-FactorIcon($slide, [string]$key, [double]$left, [double]$top) {
    if ($key -eq "temperature") {
        [void](Add-LineShape $slide "f01detail_icon_temp_stem" ($left + 16) ($top + 4) ($left + 16) ($top + 36) "#005A99" 1.8 0)
        $dot = $slide.Shapes.AddShape($msoShapeOval, $left + 8, $top + 31, 16, 16)
        $dot.Name = "f01detail_icon_temp_dot"; Set-NoFill $dot; Set-Line $dot "#005A99" 1.8 0
        [void](Add-Text $slide "f01detail_icon_temp_snow" "*" ($left + 28) ($top + 14) 18 18 16 "#005A99" 1)
    }
}

function Add-Table($slide, [string]$prefix, [double]$left, [double]$top, [double]$width, [double]$height) {
    $items = @()
    $items += Add-Panel $slide "$prefix`_box" $left $top $width $height "#B2CCE0" 0.55 0.68
    $rows = 5
    $rowH = $height / $rows
    $c1 = $width * 0.22
    $c2 = $width * 0.39
    $c3 = $width - $c1 - $c2
    for ($i = 1; $i -lt $rows; $i++) {
        $items += Add-LineShape $slide "$prefix`_h_$i" $left ($top + $i * $rowH) ($left + $width) ($top + $i * $rowH) "#B7D4E8" 0.55 0.1
    }
    $items += Add-LineShape $slide "$prefix`_v_1" ($left + $c1) $top ($left + $c1) ($top + $height) "#B7D4E8" 0.55 0.1
    $items += Add-LineShape $slide "$prefix`_v_2" ($left + $c1 + $c2) $top ($left + $c1 + $c2) ($top + $height) "#B7D4E8" 0.55 0.1

    $headers = @((U "5rC05bmz"), (U "5Y+Y6YeP5Y+W5YC8"), (U "5ruR77yI5pGp5pOm57O75pWw77yJ"))
    $data = @(
        @("T1", (U "LTE1wrBD"), "0.118"),
        @("T2", (U "LTEwwrBD"), "0.103"),
        @("T3", (U "LTXCsEM="), "0.089"),
        @("T4", (U "MMKwQw=="), "0.076")
    )
    $x1 = $left
    $x2 = $left + $c1
    $x3 = $left + $c1 + $c2
    $xs = @($x1, $x2, $x3)
    $ws = @($c1, $c2, $c3)
    for ($c = 0; $c -lt 3; $c++) {
        $items += Add-Text $slide "$prefix`_header_$c" $headers[$c] ($xs[$c] + 4) ($top + 10) ($ws[$c] - 8) 12 7.7 "#003B73" 1
        $items[-1].TextFrame2.TextRange.ParagraphFormat.Alignment = 2
    }
    for ($r = 0; $r -lt $data.Count; $r++) {
        for ($c = 0; $c -lt 3; $c++) {
            $items += Add-Text $slide "$prefix`_cell_$r`_$c" $data[$r][$c] ($xs[$c] + 4) ($top + (($r + 1) * $rowH) + 10) ($ws[$c] - 8) 12 8 "#003B73" 0
            $items[-1].TextFrame2.TextRange.ParagraphFormat.Alignment = 2
        }
    }
    return $items
}

function Add-TrendChart($slide, [string]$prefix, [double]$left, [double]$top, [double]$width, [double]$height) {
    $items = @()
    $x0 = $left + 42
    $y0 = $top + $height - 38
    $x1 = $left + $width - 16
    $y1 = $top + 28
    $items += Add-LineShape $slide "$prefix`_axis_x" $x0 $y0 $x1 $y0 "#7EAAD0" 0.8 0
    $items += Add-LineShape $slide "$prefix`_axis_y" $x0 $y0 $x0 $y1 "#7EAAD0" 0.8 0
    for ($i = 0; $i -le 4; $i++) {
        $v = 0.14 * $i / 4
        $ty = $y0 - ($y0 - $y1) * $i / 4
        $items += Add-LineShape $slide "$prefix`_tick_$i" ($x0 - 3) $ty $x0 $ty "#7EAAD0" 0.55 0
        $items += Add-Text $slide "$prefix`_tick_label_$i" ("{0:N2}" -f $v) ($left + 4) ($ty - 5) 32 10 6.9 "#003B73" 0
        $items[-1].TextFrame2.TextRange.ParagraphFormat.Alignment = 3
        if ($i -gt 0) {
            $items += Add-LineShape $slide "$prefix`_grid_$i" $x0 $ty $x1 $ty "#D0E2EF" 0.35 0.25
        }
    }
    $labels = @("T1", "T2", "T3", "T4")
    $values = @(0.118, 0.103, 0.089, 0.076)
    $points = @()
    for ($i = 0; $i -lt $values.Count; $i++) {
        $px = $x0 + ($x1 - $x0) * $i / 3
        $py = $y0 - ($y0 - $y1) * ($values[$i] / 0.14)
        $points += ,@($px, $py)
    }
    for ($i = 0; $i -lt 3; $i++) {
        $items += Add-LineShape $slide "$prefix`_line_$i" $points[$i][0] $points[$i][1] $points[$i + 1][0] $points[$i + 1][1] "#0077D9" 1.8 0
    }
    for ($i = 0; $i -lt 4; $i++) {
        $dot = $slide.Shapes.AddShape($msoShapeOval, $points[$i][0] - 3.5, $points[$i][1] - 3.5, 7, 7)
        $dot.Name = "$prefix`_dot_$i"; Set-GlassFill $dot "#EAF6FF" 0; Set-Line $dot "#0077D9" 1.5 0
        $items += $dot
        $items += Add-Text $slide "$prefix`_value_$i" ("{0:N3}" -f $values[$i]) ($points[$i][0] - 16) ($points[$i][1] - 22) 34 10 6.8 "#003B73" 0
        $items[-1].TextFrame2.TextRange.ParagraphFormat.Alignment = 2
        $items += Add-Text $slide "$prefix`_label_$i" $labels[$i] ($points[$i][0] - 12) ($y0 + 8) 24 10 7.2 "#003B73" 0
        $items[-1].TextFrame2.TextRange.ParagraphFormat.Alignment = 2
    }
    $items += Add-Text $slide "$prefix`_xlabel" (U "5rip5bqm5rC05bmz77yIwrBD77yJ") ($x0 + 35) ($y0 + 26) 160 12 7.4 "#003B73" 0
    $items[-1].TextFrame2.TextRange.ParagraphFormat.Alignment = 2
    $items += Add-Text $slide "$prefix`_ylabel" (U "5ruR77yI5pGp5pOm57O75pWw77yJ") ($left + 24) ($top + 10) 82 12 7.2 "#003B73" 0
    $items += Add-LineShape $slide "$prefix`_legend_line" ($left + $width - 72) ($top + 10) ($left + $width - 58) ($top + 10) "#0077D9" 1.6 0
    $legendDot = $slide.Shapes.AddShape($msoShapeOval, $left + $width - 67, $top + 7.8, 4, 4)
    $legendDot.Name = "$prefix`_legend_dot"; Set-GlassFill $legendDot "#EAF6FF" 0; Set-Line $legendDot "#0077D9" 1 0
    $items += $legendDot
    $items += Add-Text $slide "$prefix`_legend" (U "5ruR77yI5pGp5pOm57O75pWw77yJ") ($left + $width - 55) ($top + 5) 72 12 6.8 "#003B73" 0
    return $items
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
        $shape = $slide.Shapes.Item($i)
        $n = $shape.Name
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

function Group-Items($slide, [string]$groupName, [object[]]$items) {
    $names = @()
    foreach ($item in $items) { $names += $item.Name }
    $group = $slide.Shapes.Range($names).Group()
    $group.Name = $groupName
    return $group
}

function Add-F01DetailLayer($slide) {
    $items = @()
    $items += Add-HighlightPanel $slide "f01detail_active_card_outline" 42 218 142 136
    $arrow = $slide.Shapes.AddShape($msoShapeRightArrow, 178, 290, 44, 26)
    $arrow.Name = "f01detail_active_arrow"
    $arrow.Fill.Visible = $msoTrue
    $arrow.Fill.ForeColor.RGB = Convert-HexColor "#29A8FF"
    $arrow.Fill.Transparency = 0.08
    $arrow.Line.Visible = $msoFalse
    $items += $arrow

    $items += Add-Panel $slide "f01detail_panel" 502 118 430 448 "#70A8D0" 1.1 0.24
    $items += Add-Text $slide "f01detail_title" (U "5rip5bqm5Y+Y6YeP5YiG5p6Q") 524 142 178 26 17 "#003B73" 1
    $items += Add-Panel $slide "f01detail_badge" 654 146 28 18 "#70A8D0" 0.6 0.62
    $items += Add-Text $slide "f01detail_badge_text" "F01" 659 149 20 10 8 "#003B73" 1
    $items[-1].TextFrame2.TextRange.ParagraphFormat.Alignment = 2
    $items += Add-Text $slide "f01detail_en" "Temperature" 526 171 110 14 8.5 "#003B73" 0
    $items += Add-Text $slide "f01detail_big" "F01" 812 130 88 42 30 "#99BDD8" 1
    $items += Add-Text $slide "f01detail_close" "X" 905 134 18 18 14 "#003B73" 1
    $items[-1].TextFrame2.TextRange.ParagraphFormat.Alignment = 2

    $items += Add-Panel $slide "f01detail_content_box" 518 198 398 332 "#B2CCE0" 0.55 0.62
    $items += Add-Text $slide "f01detail_table_title" (U "5a6e6aqM5rC05bmz6KGo") 556 222 106 20 12.5 "#003B73" 1
    $items += Add-Text $slide "f01detail_chart_title" (U "6LaL5Yq/5Zu+") 716 222 80 20 12.5 "#003B73" 1
    $items += Add-Text $slide "f01detail_table_icon" "^" 534 222 16 18 15 "#005A99" 1
    $items += Add-Text $slide "f01detail_chart_icon" (U "4pal") 696 222 16 18 13 "#005A99" 1

    $items += Add-Table $slide "f01detail_table" 524 258 172 166
    $items += Add-TrendChart $slide "f01detail_chart" 704 254 198 174

    $items += Add-Panel $slide "f01detail_conclusion_box" 518 454 398 64 "#B2CCE0" 0.55 0.5
    $items += Add-Text $slide "f01detail_conclusion_icon" "!" 536 466 16 16 13 "#005A99" 1
    $items += Add-Text $slide "f01detail_conclusion_title" (U "57uT6K66") 556 465 60 18 13 "#003B73" 1
    $items += Add-Text $slide "f01detail_conclusion_text" (U "5rip5bqm5Y2H6auY5pe277yM5Yaw6KGo6Z2i5ray6Iac5aKe5by677yM5pGp5pOm57O75pWwIM68IOaVtOS9k+S4i+mZjeOAgg==") 536 491 356 18 8.5 "#003B73" 1
    $items += Add-Text $slide "f01detail_note" (U "5rOo77ya5b2T5YmN5pWw5o2u5Li656S65L6L77yM5ZCO57ut5Y+v5pu/5o2i5Li655yf5a6e5a6e6aqM5pWw5o2u44CC") 526 544 250 14 7.4 "#476378" 0
    $items += Add-Panel $slide "f01detail_back_button" 804 532 106 26 "#005A99" 0.9 0.35
    $items += Add-Text $slide "f01detail_back_text" (U "PCAg6L+U5Zue55+p6Zi1") 816 540 82 10 9 "#003B73" 1
    $items[-1].TextFrame2.TextRange.ParagraphFormat.Alignment = 2
    return Group-Items $slide "anim_f01_detail_group" $items
}

function Add-TriggeredEffect($sequence, $shape, $triggerShape, $effectType, [bool]$exit, [double]$duration) {
    $effect = $sequence.AddEffect($shape, $effectType, $msoAnimateLevelNone, $msoTriggerOnClick)
    $effect.Timing.TriggerShape = $triggerShape
    $effect.Timing.Duration = $duration
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

if (Test-Path -LiteralPath $outPath) { Remove-Item -LiteralPath $outPath -Force }
Copy-Item -LiteralPath $basePath -Destination $outPath -Force

$app = New-Object -ComObject PowerPoint.Application
$app.Visible = $msoTrue
$presentation = $null

try {
    $presentation = $app.Presentations.Open($outPath, $msoFalse, $msoFalse, $msoTrue)
    $slide = $presentation.Slides.Item(1)
    $slide.Name = "quality_overview_f01_animation_test"
    $slide.SlideShowTransition.AdvanceOnClick = $msoFalse
    $slide.SlideShowTransition.AdvanceOnTime = $msoFalse

    $overviewNames = @(Get-OverviewNames $slide)
    $overviewGroup = $slide.Shapes.Range($overviewNames).Group()
    $overviewGroup.Name = "anim_overview_group"
    $detailGroup = Add-F01DetailLayer $slide

    $triggerTemperature = Add-HitArea $slide "trigger_temperature_card_f01" 42 218 142 136
    $triggerBack = Add-HitArea $slide "trigger_back_to_overview" 804 532 106 26
    $triggerTemperature.ZOrder(0) | Out-Null
    $triggerBack.ZOrder(0) | Out-Null

    $seqShow = $slide.TimeLine.InteractiveSequences.Add()
    [void](Add-TriggeredEffect $seqShow $overviewGroup $triggerTemperature $msoAnimEffectFade $true 0.15)
    [void](Add-WithPreviousEffect $seqShow $detailGroup $msoAnimEffectWipe $false 0.30)

    $seqBack = $slide.TimeLine.InteractiveSequences.Add()
    [void](Add-TriggeredEffect $seqBack $detailGroup $triggerBack $msoAnimEffectFade $true 0.15)
    [void](Add-WithPreviousEffect $seqBack $overviewGroup $msoAnimEffectWipe $false 0.30)

    $detailGroup.Visible = $msoFalse
    $overviewGroup.Visible = $msoTrue
    $slide.Export($previewPath, "PNG", 1536, 1024)
    $slide.Export($initialStatePath, "PNG", 1536, 1024)
    $detailGroup.Visible = $msoTrue
    $overviewGroup.Visible = $msoFalse
    $slide.Export($detailStatePath, "PNG", 1536, 1024)
    $overviewGroup.Visible = $msoTrue
    $detailGroup.Visible = $msoTrue
    $presentation.Save()

    @(
        "# Animation test based on quality static overview",
        "",
        "Base file: editable_static_overview/factor_matrix_static_editable_overview.pptx",
        "Output file: factor_matrix_static_editable_overview_f01_animation_test.pptx",
        "",
        "Groups:",
        "- anim_overview_group: original right-side overview objects from the quality static PPT.",
        "- anim_f01_detail_group: newly added F01 detail group using the same visual style.",
        "- trigger_temperature_card_f01: transparent click area over the F01 card.",
        "- trigger_back_to_overview: transparent click area over the return button.",
        "",
        "Logic:",
        "1. Click F01 card: overview fades out; F01 detail wipes in from left to right.",
        "2. Click return: F01 detail fades out; overview wipes in from left to right.",
        "",
        "This file only tests overview + F01."
    ) | Set-Content -LiteralPath $notesPath -Encoding UTF8

    Write-Output "Generated $outPath"
    Write-Output "Preview $previewPath"
    Write-Output "Initial state $initialStatePath"
    Write-Output "Detail state $detailStatePath"
}
finally {
    if ($presentation -ne $null) { try { $presentation.Close() } catch {} }
    try { $app.Quit() } catch {}
}
