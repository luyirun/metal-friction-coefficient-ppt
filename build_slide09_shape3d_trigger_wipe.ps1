$ErrorActionPreference = "Stop"

Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint
Add-Type -AssemblyName Office

$msoTrue = -1
$msoFalse = 0
$msoTextOrientationHorizontal = 1
$msoShapeRectangle = 1
$msoShapeRoundedRectangle = 5
$msoShapeOval = 9
$ppLayoutBlank = 12
$ppSaveAsOpenXMLPresentation = 24

$root = Get-Location
$outDir = Join-Path $root "projects\slide09_whitepaper\exports"
$assetDir = Join-Path $root "projects\slide09_whitepaper\assets"
$outPath = Join-Path $outDir "slide09_shape3d_trigger_wipe_lr.pptx"
$modelPath = Join-Path $assetDir "mvp_native_trigger_plate.obj"
New-Item -ItemType Directory -Force -Path $outDir, $assetDir | Out-Null

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

function Set-Line($shape, [string]$hex, [double]$weight) {
    $shape.Line.Visible = $msoTrue
    $shape.Line.ForeColor.RGB = Convert-HexColor $hex
    $shape.Line.Weight = $weight
}

function Set-TextStyle($shape, [double]$fontSize, [string]$hex, [int]$bold = 0) {
    $shape.Fill.Visible = $msoFalse
    $shape.Line.Visible = $msoFalse
    $shape.TextFrame2.MarginLeft = 0
    $shape.TextFrame2.MarginRight = 0
    $shape.TextFrame2.MarginTop = 0
    $shape.TextFrame2.MarginBottom = 0
    $shape.TextFrame2.TextRange.Font.NameFarEast = "Microsoft YaHei"
    $shape.TextFrame2.TextRange.Font.Name = "Aptos"
    $shape.TextFrame2.TextRange.Font.Size = $fontSize
    $shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Convert-HexColor $hex
    if ($bold -ne 0) {
        $shape.TextFrame2.TextRange.Font.Bold = $msoTrue
    }
}

function Add-Text($slide, [string]$name, [string]$text, [double]$left, [double]$top, [double]$width, [double]$height, [double]$fontSize, [string]$hex, [int]$bold = 0) {
    $shape = $slide.Shapes.AddTextbox($msoTextOrientationHorizontal, $left, $top, $width, $height)
    $shape.Name = $name
    $shape.TextFrame2.TextRange.Text = $text
    Set-TextStyle $shape $fontSize $hex $bold
    return $shape
}

function Add-Panel($slide, [string]$name, [double]$left, [double]$top, [double]$width, [double]$height, [string]$lineHex) {
    $shape = $slide.Shapes.AddShape($msoShapeRoundedRectangle, $left, $top, $width, $height)
    $shape.Name = $name
    Set-NoFill $shape
    Set-Line $shape $lineHex 1.35
    $shape.Shadow.Visible = $msoTrue
    $shape.Shadow.ForeColor.RGB = Convert-HexColor "#9AA7B4"
    $shape.Shadow.Transparency = 0.58
    $shape.Shadow.Blur = 9
    $shape.Shadow.OffsetX = 2
    $shape.Shadow.OffsetY = 3
    $shape.ThreeD.Visible = $msoTrue
    $shape.ThreeD.Depth = 9
    $shape.ThreeD.BevelTopType = [Microsoft.Office.Core.MsoBevelType]::msoBevelSoftRound
    $shape.ThreeD.BevelTopInset = 4
    $shape.ThreeD.BevelTopDepth = 2
    $shape.ThreeD.SetPresetCamera([Microsoft.Office.Core.MsoPresetCamera]::msoCameraLegacyPerspectiveFront)
    return $shape
}

function Add-HRule($slide, [string]$name, [double]$x1, [double]$y, [double]$x2, [string]$hex) {
    $line = $slide.Shapes.AddLine($x1, $y, $x2, $y)
    $line.Name = $name
    $line.Line.ForeColor.RGB = Convert-HexColor $hex
    $line.Line.Weight = 0.75
    return $line
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

function Add-Button($slide, [string]$name, [string]$text, [double]$left, [double]$top, [double]$width, [double]$height) {
    $button = $slide.Shapes.AddShape($msoShapeRoundedRectangle, $left, $top, $width, $height)
    $button.Name = $name
    $button.Fill.Visible = $msoTrue
    $button.Fill.ForeColor.RGB = Convert-HexColor "#EFF5F8"
    $button.Line.Visible = $msoTrue
    $button.Line.ForeColor.RGB = Convert-HexColor "#3E6B82"
    $button.Line.Weight = 1
    $button.TextFrame2.TextRange.Text = $text
    $button.TextFrame2.MarginLeft = 8
    $button.TextFrame2.MarginRight = 8
    $button.TextFrame2.MarginTop = 4
    $button.TextFrame2.MarginBottom = 4
    $button.TextFrame2.TextRange.Font.NameFarEast = "Microsoft YaHei"
    $button.TextFrame2.TextRange.Font.Name = "Aptos"
    $button.TextFrame2.TextRange.Font.Size = 10
    $button.TextFrame2.TextRange.Font.Bold = $msoTrue
    $button.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Convert-HexColor "#3E6B82"
    return $button
}

function Group-Shapes($slide, [string]$groupName, [object[]]$shapes) {
    $names = @()
    foreach ($shape in $shapes) {
        $names += $shape.Name
    }
    $group = $slide.Shapes.Range($names).Group()
    $group.Name = $groupName
    return $group
}

function Write-ObjModel([string]$path) {
    $obj = @(
        "o MVP_Native_Trigger_Metal_Plate",
        "v -2.2 -0.35 -0.08",
        "v  2.2 -0.35 -0.08",
        "v  2.2  0.35 -0.08",
        "v -2.2  0.35 -0.08",
        "v -2.2 -0.35  0.08",
        "v  2.2 -0.35  0.08",
        "v  2.2  0.35  0.08",
        "v -2.2  0.35  0.08",
        "f 1 2 3 4",
        "f 5 8 7 6",
        "f 1 5 6 2",
        "f 2 6 7 3",
        "f 3 7 8 4",
        "f 5 1 4 8"
    )
    Set-Content -LiteralPath $path -Value $obj -Encoding ASCII
}

function Add-Model($slide, [string]$name, [string]$path, [double]$left, [double]$top, [double]$width, [double]$height, [double]$rotX, [double]$rotY) {
    $model = $slide.Shapes.Add3DModel($path, $msoFalse, $msoTrue, $left, $top, $width, $height)
    $model.Name = $name
    $model.Model3D.RotationX = $rotX
    $model.Model3D.RotationY = $rotY
    return $model
}

function Add-FrontCard($slide, [string]$prefix, [double]$left, [double]$top, [double]$width, [double]$height) {
    $items = @()
    $items += Add-Panel $slide "$prefix`_panel" $left $top $width $height "#3E6B82"
    $items += Add-Text $slide "$prefix`_index" "F01" ($left + 34) ($top + 32) 70 28 16 "#3E6B82" 1
    $items += Add-Text $slide "$prefix`_title" (U "5rip5bqm") ($left + 34) ($top + 84) 170 52 32 "#1F2A33" 1
    $items += Add-Text $slide "$prefix`_subtitle" "Temperature" ($left + 36) ($top + 138) 190 28 13 "#5E6B74" 0
    $items += Add-Text $slide "$prefix`_hint" (U "5LiJ57u05qC85byP5Y2h54mH5Y6a5bqm") ($left + 268) ($top + 122) 130 18 9 "#697782" 0
    return Group-Shapes $slide $prefix $items
}

function Add-TrendChart($slide, [string]$prefix, [double]$left, [double]$top, [double]$width, [double]$height) {
    $items = @()
    $items += Add-Text $slide "$prefix`_title" (U "6LaL5Yq/5Zu+") $left ($top - 24) 120 18 12 "#3F4B53" 1
    $x0 = $left + 18
    $y0 = $top + $height - 22
    $x1 = $left + $width - 14
    $y1 = $top + 12
    $axisX = $slide.Shapes.AddLine($x0, $y0, $x1, $y0)
    $axisX.Name = "$prefix`_axis_x"
    $axisX.Line.ForeColor.RGB = Convert-HexColor "#8FA2AF"
    $axisX.Line.Weight = 1
    $items += $axisX
    $axisY = $slide.Shapes.AddLine($x0, $y0, $x0, $y1)
    $axisY.Name = "$prefix`_axis_y"
    $axisY.Line.ForeColor.RGB = Convert-HexColor "#8FA2AF"
    $axisY.Line.Weight = 1
    $items += $axisY
    $points = @(
        @(($x0 + 18), ($top + 28), "0.118"),
        @(($x0 + 78), ($top + 52), "0.103"),
        @(($x0 + 138), ($top + 78), "0.089"),
        @(($x0 + 198), ($top + 102), "0.076")
    )
    for ($i = 0; $i -lt ($points.Count - 1); $i++) {
        $line = $slide.Shapes.AddLine($points[$i][0], $points[$i][1], $points[$i + 1][0], $points[$i + 1][1])
        $line.Name = "$prefix`_line_$i"
        $line.Line.ForeColor.RGB = Convert-HexColor "#3E6B82"
        $line.Line.Weight = 2
        $items += $line
    }
    for ($i = 0; $i -lt $points.Count; $i++) {
        $dot = $slide.Shapes.AddShape($msoShapeOval, $points[$i][0] - 4, $points[$i][1] - 4, 8, 8)
        $dot.Name = "$prefix`_dot_$i"
        $dot.Fill.Visible = $msoTrue
        $dot.Fill.ForeColor.RGB = Convert-HexColor "#3E6B82"
        Set-Line $dot "#FFFFFF" 0.75
        $items += $dot
        $items += Add-Text $slide "$prefix`_label_$i" $points[$i][2] ($points[$i][0] - 14) ($points[$i][1] - 20) 38 12 7 "#3F4B53" 0
    }
    $items += Add-Text $slide "$prefix`_y_high" (U "6auY") ($x0 - 16) ($y1 - 4) 14 14 8 "#697782" 0
    $items += Add-Text $slide "$prefix`_y_low" (U "5L2O") ($x0 - 16) ($y0 - 8) 14 14 8 "#697782" 0
    $items += Add-Text $slide "$prefix`_caption" (U "5rip5bqm5Y2H6auY77yMzrwg5LiL6ZmN") ($left + 34) ($top + $height - 4) 180 16 8 "#697782" 0
    return Group-Shapes $slide $prefix $items
}

function Add-BackAnalysis($slide, [string]$prefix, [double]$left, [double]$top, [double]$width, [double]$height) {
    $items = @()
    $items += Add-Panel $slide "$prefix`_panel" $left $top $width $height "#6B7B3F"
    $items += Add-Text $slide "$prefix`_title" (U "5rip5bqm5Y+Y6YeP5YiG5p6Q") ($left + 30) ($top + 22) 220 30 20 "#1F2A33" 1
    $items += Add-Text $slide "$prefix`_subtitle" "Controlled variable analysis" ($left + 32) ($top + 54) 220 20 9 "#697782" 0
    $items += Add-Text $slide "$prefix`_table_title" (U "5a6e6aqM5rC05bmz6KGo") ($left + 32) ($top + 98) 120 20 12 "#3F4B53" 1
    $tableTop = $top + 128
    $col1 = $left + 34
    $col2 = $left + 98
    $col3 = $left + 202
    $items += Add-HRule $slide "$prefix`_table_top" ($left + 32) $tableTop ($left + 270) "#B7C3CC"
    $items += Add-Text $slide "$prefix`_h1" (U "5rC05bmz") $col1 ($tableTop + 10) 52 18 10 "#3F4B53" 1
    $items += Add-Text $slide "$prefix`_h2" (U "5Y+Y6YeP5Y+W5YC8") $col2 ($tableTop + 10) 86 18 10 "#3F4B53" 1
    $items += Add-Text $slide "$prefix`_h3" "μ" $col3 ($tableTop + 10) 48 18 10 "#3F4B53" 1
    $rows = @(
        @("T1", (U "LTE1IOKEgw=="), "0.118"),
        @("T2", (U "LTEwIOKEgw=="), "0.103"),
        @("T3", (U "LTUg4oSD"), "0.089"),
        @("T4", (U "MCDihIM="), "0.076")
    )
    for ($i = 0; $i -lt $rows.Count; $i++) {
        $y = $tableTop + 36 + ($i * 25)
        $items += Add-HRule $slide "$prefix`_row_$i" ($left + 32) ($y - 5) ($left + 270) "#D6DEE4"
        $items += Add-Text $slide "$prefix`_c1_$i" $rows[$i][0] $col1 $y 52 18 10 "#25313A" 0
        $items += Add-Text $slide "$prefix`_c2_$i" $rows[$i][1] $col2 $y 86 18 10 "#25313A" 0
        $items += Add-Text $slide "$prefix`_c3_$i" $rows[$i][2] $col3 $y 48 18 10 "#25313A" 0
    }
    $chart = Add-TrendChart $slide "$prefix`_trend_chart" ($left + 360) ($top + 140) 250 150
    $items += $chart
    $items += Add-Text $slide "$prefix`_conclusion" (U "57uT6K6677ya5rip5bqm5Y2H6auY5pe277yM5Yaw6KGo6Z2i5ray6Iac5aKe5by677yM5pGp5pOm57O75pWwIM68IOaVtOS9k+S4i+mZjeOAgg==") ($left + 32) ($top + $height - 58) 500 28 11 "#1F2A33" 0
    $items += Add-Text $slide "$prefix`_return" (U "54K55Ye76L+U5Zue") ($left + $width - 96) ($top + $height - 42) 70 18 10 "#3E6B82" 1
    return Group-Shapes $slide $prefix $items
}

function Add-TriggeredEffect($sequence, $shape, $triggerShape, $effectType, [bool]$exit, [double]$duration) {
    $effect = $sequence.AddEffect($shape, $effectType, [Microsoft.Office.Interop.PowerPoint.MsoAnimateByLevel]::msoAnimateLevelNone, [Microsoft.Office.Interop.PowerPoint.MsoAnimTriggerType]::msoAnimTriggerOnShapeClick)
    $effect.Timing.TriggerShape = $triggerShape
    $effect.Timing.Duration = $duration
    if ($exit) {
        $effect.Exit = $msoTrue
    }
    return $effect
}

function Add-WithPreviousEffect($sequence, $shape, $effectType, [bool]$exit, [double]$duration) {
    $effect = $sequence.AddEffect($shape, $effectType, [Microsoft.Office.Interop.PowerPoint.MsoAnimateByLevel]::msoAnimateLevelNone, [Microsoft.Office.Interop.PowerPoint.MsoAnimTriggerType]::msoAnimTriggerWithPrevious)
    $effect.Timing.Duration = $duration
    if ($exit) {
        $effect.Exit = $msoTrue
    }
    return $effect
}

Write-ObjModel $modelPath

$app = New-Object -ComObject PowerPoint.Application
$app.Visible = $msoTrue
$presentation = $null

try {
    if (Test-Path -LiteralPath $outPath) {
        Remove-Item -LiteralPath $outPath -Force
    }

    $presentation = $app.Presentations.Add($msoTrue)
    $presentation.PageSetup.SlideWidth = 960
    $presentation.PageSetup.SlideHeight = 540
    $slide = $presentation.Slides.Add(1, $ppLayoutBlank)
    $slide.Name = "native_trigger_3d_single_page"
    $slide.SlideShowTransition.AdvanceOnClick = $msoFalse
    $slide.SlideShowTransition.AdvanceOnTime = $msoFalse

    $bg = $slide.Shapes.AddShape($msoShapeRectangle, 0, 0, 960, 540)
    $bg.Name = "mvp_background_texture_placeholder"
    Set-NoFill $bg
    Set-Line $bg "#E3E9EE" 0.75
    $bg.ZOrder(1) | Out-Null

    [void](Add-Text $slide "mvp_title" (U "5Y6F55Sf6Kem5Y+RICsgM0Qg5Y2V6aG16aqM6K+B") 54 36 560 32 22 "#1F2A33" 1)
    [void](Add-Text $slide "mvp_subtitle" (U "5ZCM5LiA6aG15YaF77ya54K55Ye75q2j6Z2i57+75Yiw6IOM6Z2i77yM54K55Ye76IOM6Z2i6L+U5Zue5q2j6Z2i44CC") 56 72 760 24 11 "#697782" 0)

    $frontInitial = Add-FrontCard $slide "front_initial" 250 128 460 300
    $backAnalysis = Add-BackAnalysis $slide "back_analysis" 118 116 724 330
    $frontReturn = Add-FrontCard $slide "front_return" 250 128 460 300
    $frontButtonInitial = Add-Button $slide "front_button_initial" (U "54K55Ye757+76L2s") 548 368 96 28
    $backButton = Add-Button $slide "back_button" (U "6L+U5Zue5q2j6Z2i") 736 392 84 28
    $frontButtonReturn = Add-Button $slide "front_button_return" (U "54K55Ye757+76L2s") 548 368 96 28

    $frontInitial.ZOrder(0) | Out-Null
    $backAnalysis.ZOrder(0) | Out-Null
    $frontReturn.ZOrder(0) | Out-Null
    $frontButtonInitial.ZOrder(0) | Out-Null
    $backButton.ZOrder(0) | Out-Null
    $frontButtonReturn.ZOrder(0) | Out-Null

    $frontInitial.ThreeD.RotationY = 12
    $backAnalysis.ThreeD.RotationY = -10
    $frontReturn.ThreeD.RotationY = 12

    $seq1 = $slide.TimeLine.InteractiveSequences.Add()
    [void](Add-TriggeredEffect $seq1 $frontInitial $frontButtonInitial ([Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectFade) $true 0.18)
    [void](Add-WithPreviousEffect $seq1 $frontButtonInitial ([Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectFade) $true 0.05)
    [void](Add-WithPreviousEffect $seq1 $backAnalysis ([Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectWipe) $false 0.34)
    [void](Add-WithPreviousEffect $seq1 $backButton ([Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectFade) $false 0.05)

    $seq2 = $slide.TimeLine.InteractiveSequences.Add()
    [void](Add-TriggeredEffect $seq2 $backAnalysis $backButton ([Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectFade) $true 0.18)
    [void](Add-WithPreviousEffect $seq2 $backButton ([Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectFade) $true 0.05)
    [void](Add-WithPreviousEffect $seq2 $frontReturn ([Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectWipe) $false 0.34)
    [void](Add-WithPreviousEffect $seq2 $frontButtonReturn ([Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectFade) $false 0.05)

    $seq3 = $slide.TimeLine.InteractiveSequences.Add()
    [void](Add-TriggeredEffect $seq3 $frontReturn $frontButtonReturn ([Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectFade) $true 0.18)
    [void](Add-WithPreviousEffect $seq3 $frontButtonReturn ([Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectFade) $true 0.05)
    [void](Add-WithPreviousEffect $seq3 $backAnalysis ([Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectWipe) $false 0.34)
    [void](Add-WithPreviousEffect $seq3 $backButton ([Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectFade) $false 0.05)

    [void](Add-Text $slide "mvp_note" (U "6K+05piO77ya6L+Z5piv5Y6f55SfIFBvd2VyUG9pbnQg6Kem5Y+R5Zmo5Y+KIDNEIOaooeWei+S9k++8jFBBIOS4jeS9nOaSreaUvuerr+S+nei1lui/m+ihjOmqjOivgeOAgg==") 54 474 820 28 9 "#697782" 0)

    $presentation.SaveAs($outPath, $ppSaveAsOpenXMLPresentation)
    Write-Output "Generated $outPath"
}
finally {
    if ($presentation -ne $null) {
        try {
            $presentation.Close()
        }
        catch {
            Write-Output ("Presentation close skipped: " + $_.Exception.Message)
        }
    }
    try {
        $app.Quit()
    }
    catch {
        Write-Output ("PowerPoint quit skipped: " + $_.Exception.Message)
    }
}
