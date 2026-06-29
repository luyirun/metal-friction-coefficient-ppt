$ErrorActionPreference = "Stop"

Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint
Add-Type -AssemblyName Office

$msoTrue = -1
$msoFalse = 0
$msoTextOrientationHorizontal = 1
$msoShapeRectangle = 1
$msoShapeRoundedRectangle = 5
$msoShapeLine = 9
$ppLayoutBlank = 12
$ppSaveAsOpenXMLPresentation = 24

$outDir = Join-Path (Get-Location) "projects\slide09_whitepaper\exports"
$outPath = Join-Path $outDir "slide09_flip_mvp_test.pptx"
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

function Convert-HexColor([string]$hex) {
    $value = $hex.TrimStart("#")
    $r = [Convert]::ToInt32($value.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($value.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($value.Substring(4, 2), 16)
    return $r + ($g * 256) + ($b * 65536)
}

function U([string]$base64) {
    return [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($base64))
}

function Set-NoFill($shape) {
    $shape.Fill.Visible = $msoFalse
}

function Set-Line($shape, [string]$hex, [double]$weight) {
    $shape.Line.Visible = $msoTrue
    $shape.Line.ForeColor.RGB = Convert-HexColor $hex
    $shape.Line.Weight = $weight
}

function Set-TextStyle($shape, [string]$fontName, [double]$fontSize, [string]$hex, [int]$bold = 0) {
    $shape.Fill.Visible = $msoFalse
    $shape.Line.Visible = $msoFalse
    $shape.TextFrame2.MarginLeft = 0
    $shape.TextFrame2.MarginRight = 0
    $shape.TextFrame2.MarginTop = 0
    $shape.TextFrame2.MarginBottom = 0
    $shape.TextFrame2.TextRange.Font.NameFarEast = $fontName
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
    Set-TextStyle $shape "Microsoft YaHei" $fontSize $hex $bold
    return $shape
}

function Add-Panel($slide, [string]$name, [double]$left, [double]$top, [double]$width, [double]$height, [string]$lineHex) {
    $shape = $slide.Shapes.AddShape($msoShapeRoundedRectangle, $left, $top, $width, $height)
    $shape.Name = $name
    Set-NoFill $shape
    Set-Line $shape $lineHex 1.5
    $shape.Shadow.Visible = $msoTrue
    $shape.Shadow.ForeColor.RGB = Convert-HexColor "#9AA7B4"
    $shape.Shadow.Transparency = 0.55
    $shape.Shadow.Blur = 9
    $shape.Shadow.OffsetX = 2
    $shape.Shadow.OffsetY = 3
    $shape.ThreeD.Visible = $msoTrue
    $shape.ThreeD.Depth = 8
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

function Group-Shapes($slide, [string]$groupName, [object[]]$shapes) {
    $names = @()
    foreach ($shape in $shapes) {
        $names += $shape.Name
    }
    $group = $slide.Shapes.Range($names).Group()
    $group.Name = $groupName
    return $group
}

function Add-FrontCard($slide, [string]$prefix, [double]$left, [double]$top, [double]$width, [double]$height) {
    $items = @()
    $items += Add-Panel $slide "$prefix`_panel" $left $top $width $height "#3E6B82"
    $items += Add-Text $slide "$prefix`_index" "F01" ($left + 22) ($top + 20) 70 28 16 "#3E6B82" 1
    $items += Add-Text $slide "$prefix`_title" (U "5rip5bqm") ($left + 22) ($top + 62) ($width - 44) 48 28 "#1F2A33" 1
    $items += Add-Text $slide "$prefix`_subtitle" "Temperature" ($left + 24) ($top + 108) ($width - 48) 28 13 "#5E6B74" 0
    $items += Add-Text $slide "$prefix`_hint" (U "54K55Ye757+76L2s5p+l55yL5Y+Y6YeP5YiG5p6Q") ($left + 24) ($top + $height - 46) ($width - 48) 24 12 "#697782" 0
    $items += Add-HRule $slide "$prefix`_rule" ($left + 24) ($top + $height - 62) ($left + $width - 24) "#B8C7D2"
    return Group-Shapes $slide $prefix $items
}

function Add-BackCard($slide, [string]$prefix, [double]$left, [double]$top, [double]$width, [double]$height) {
    $items = @()
    $items += Add-Panel $slide "$prefix`_panel" $left $top $width $height "#6B7B3F"
    $items += Add-Text $slide "$prefix`_title" (U "5rip5bqm5Y+Y6YeP5YiG5p6Q") ($left + 24) ($top + 20) ($width - 48) 30 20 "#1F2A33" 1
    $items += Add-Text $slide "$prefix`_subtitle" "Controlled variable: metal-ice friction coefficient" ($left + 24) ($top + 52) ($width - 48) 20 9 "#697782" 0

    $tableTop = $top + 88
    $col1 = $left + 26
    $col2 = $left + 96
    $col3 = $left + 206
    $items += Add-HRule $slide "$prefix`_table_top" ($left + 24) $tableTop ($left + $width - 24) "#B7C3CC"
    $items += Add-Text $slide "$prefix`_h1" (U "5rC05bmz") $col1 ($tableTop + 10) 52 18 10 "#3F4B53" 1
    $items += Add-Text $slide "$prefix`_h2" (U "5Y+Y6YeP5Y+W5YC8") $col2 ($tableTop + 10) 92 18 10 "#3F4B53" 1
    $items += Add-Text $slide "$prefix`_h3" "μ" $col3 ($tableTop + 10) 48 18 10 "#3F4B53" 1

    $rows = @(
        @("T1", (U "LTE1IOKEgw=="), "0.118"),
        @("T2", (U "LTEwIOKEgw=="), "0.103"),
        @("T3", (U "LTUg4oSD"), "0.089"),
        @("T4", (U "MCDihIM="), "0.076")
    )
    for ($i = 0; $i -lt $rows.Count; $i++) {
        $y = $tableTop + 36 + ($i * 25)
        $items += Add-HRule $slide "$prefix`_row_$i" ($left + 24) ($y - 5) ($left + $width - 24) "#D6DEE4"
        $items += Add-Text $slide "$prefix`_c1_$i" $rows[$i][0] $col1 $y 52 18 10 "#25313A" 0
        $items += Add-Text $slide "$prefix`_c2_$i" $rows[$i][1] $col2 $y 92 18 10 "#25313A" 0
        $items += Add-Text $slide "$prefix`_c3_$i" $rows[$i][2] $col3 $y 48 18 10 "#25313A" 0
    }

    $items += Add-Text $slide "$prefix`_conclusion" (U "57uT6K6677ya5rip5bqm5Y2H6auY5pe277yM5Yaw6KGo6Z2i5ray6Iac5aKe5by677yM5pGp5pOm57O75pWwIM68IOaVtOS9k+S4i+mZjeOAgg==") ($left + 24) ($top + $height - 76) ($width - 48) 34 11 "#1F2A33" 0
    $items += Add-Text $slide "$prefix`_return" (U "54K55Ye76L+U5Zue") ($left + $width - 92) ($top + $height - 38) 68 18 10 "#3E6B82" 1
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
    $slide.Name = "flip_mvp"

    $bg = $slide.Shapes.AddShape($msoShapeRectangle, 0, 0, 960, 540)
    $bg.Name = "mvp_background_texture_placeholder"
    Set-NoFill $bg
    Set-Line $bg "#E3E9EE" 0.75

    [void](Add-Text $slide "mvp_title" (U "5LiJ57u057+76L2s6Kem5Y+R5ZmoIE1WUA==") 54 36 460 32 22 "#1F2A33" 1)
    [void](Add-Text $slide "mvp_subtitle" (U "6aqM6K+B77ya54K55Ye75q2j6Z2i57+75Yiw6IOM6Z2i77yM54K55Ye76IOM6Z2i6L+U5Zue5q2j6Z2i77yb5paH5a2X5LiO6KGo5qC85L+d5oyB5Y+v57yW6L6R44CC") 56 72 690 24 11 "#697782" 0)

    $frontInitial = Add-FrontCard $slide "front_initial" 310 144 340 238
    $backCard = Add-BackCard $slide "back_card" 310 144 340 238
    $frontReturn = Add-FrontCard $slide "front_return" 310 144 340 238

    $frontInitial.ZOrder(0) | Out-Null
    $backCard.ZOrder(0) | Out-Null
    $frontReturn.ZOrder(0) | Out-Null

    $frontInitial.ThreeD.RotationY = 8
    $backCard.ThreeD.RotationY = -8
    $frontReturn.ThreeD.RotationY = 8

    $seq1 = $slide.TimeLine.InteractiveSequences.Add()
    [void](Add-TriggeredEffect $seq1 $frontInitial $frontInitial ([Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectFlip) $true 0.32)
    [void](Add-TriggeredEffect $seq1 $backCard $frontInitial ([Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectFadedSwivel) $false 0.36)

    $seq2 = $slide.TimeLine.InteractiveSequences.Add()
    [void](Add-TriggeredEffect $seq2 $backCard $backCard ([Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectFlip) $true 0.32)
    [void](Add-TriggeredEffect $seq2 $frontReturn $backCard ([Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectFadedSwivel) $false 0.36)

    [void](Add-Text $slide "mvp_note" (U "5Yik5a6a5Y+j5b6E77ya6Iul5pys5paH5Lu25Zyo5pyq5a6J6KOFIFBBIOeahCBQb3dlclBvaW50IOS4reS7jeWPr+inpuWPkeaSreaUvu+8jOWImeivtOaYjuaSreaUvuerr+S4jeS+nei1liBQQe+8m+iLpeiAgeeJiOacrOmZjee6p++8jOWPquiusOW9lemZjee6p++8jOS4jeaJqeWxleWIsOWujOaVtOmhteOAgg==") 54 466 820 30 9 "#697782" 0)

    $presentation.SaveAs($outPath, $ppSaveAsOpenXMLPresentation)
    Write-Output "Generated $outPath"
}
finally {
    if ($presentation -ne $null) {
        $presentation.Close()
    }
    $app.Quit()
}
