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
$ppActionNextSlide = 1
$ppActionPreviousSlide = 2
$ppEffectMorphByObject = 3954

$root = Get-Location
$outDir = Join-Path $root "projects\slide09_whitepaper\exports"
$assetDir = Join-Path $root "projects\slide09_whitepaper\assets"
$outPath = Join-Path $outDir "slide09_morph_3d_mvp_test.pptx"
$modelPath = Join-Path $assetDir "mvp_metal_plate.obj"
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
    Set-Line $shape $lineHex 1.3
    $shape.Shadow.Visible = $msoTrue
    $shape.Shadow.ForeColor.RGB = Convert-HexColor "#9AA7B4"
    $shape.Shadow.Transparency = 0.58
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

function Write-ObjModel([string]$path) {
    $obj = @(
        "o MVP_Metal_Plate",
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

function Add-ModelIfAvailable($slide, [string]$name, [string]$path, [double]$left, [double]$top, [double]$width, [double]$height) {
    try {
        $model = $slide.Shapes.Add3DModel($path, $msoFalse, $msoTrue, $left, $top, $width, $height)
        $model.Name = $name
        $model.Model3D.RotationX = 330
        $model.Model3D.RotationY = 25
        return $model
    }
    catch {
        return Add-Text $slide "$name`_fallback" ("3D model insert failed: " + $_.Exception.Message) $left $top $width 36 9 "#9B2F2F" 0
    }
}

function Add-FrontSlide($slide, [string]$modelPath) {
    [void](Add-Text $slide "mvp_title" (U "TW9ycGggKyAzRCDmqKHlnovpqozor4E=") 54 36 520 32 22 "#1F2A33" 1)
    [void](Add-Text $slide "mvp_subtitle" (U "54K55Ye75q2j6Z2i5Y2h54mH77ya55SoIE1vcnBoIOi/nua4oOWIsOiDjOmdoi/pgJrov4fliLDkuIvkuIDpobXjgII=") 56 72 760 24 11 "#697782" 0)
    $panel = Add-Panel $slide "morph_card" 285 138 390 260 "#3E6B82"
    $panel.ThreeD.RotationY = 18
    [void](Add-Text $slide "morph_index" "F01" 315 166 70 28 16 "#3E6B82" 1)
    [void](Add-Text $slide "morph_title" (U "5rip5bqm") 315 210 160 48 30 "#1F2A33" 1)
    [void](Add-Text $slide "morph_subtitle" "Temperature" 318 258 170 26 13 "#5E6B74" 0)
    [void](Add-Text $slide "morph_hint" (U "54K55Ye75Y2h54mH77yaTW9ycGgg57+75Yiw5YiG5p6Q6IOM6Z2i") 318 342 260 24 12 "#697782" 0)
    [void](Add-ModelIfAvailable $slide "mvp_3d_model_front" $modelPath 492 178 130 110)
    $panel.ActionSettings(1).Action = $ppActionNextSlide
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

function Add-BackSlide($slide, [string]$modelPath) {
    [void](Add-Text $slide "mvp_title" (U "TW9ycGggKyAzRCDmqKHlnovpqozor4E=") 54 36 520 32 22 "#1F2A33" 1)
    [void](Add-Text $slide "mvp_subtitle" (U "6IOM6Z2i5pi+56S677ya5a6e6aqM5rC05bmz6KGoICsg6LaL5Yq/5Zu+ICsg57uT6K66ICsg55yf5a6eIDNEIOaooeWei+OAgQ==") 56 72 760 24 11 "#697782" 0)
    $panel = Add-Panel $slide "morph_card" 118 122 724 318 "#6B7B3F"
    $panel.ThreeD.RotationY = -18
    [void](Add-Text $slide "back_title" (U "5rip5bqm5Y+Y6YeP5YiG5p6Q") 148 146 220 30 20 "#1F2A33" 1)
    [void](Add-Text $slide "back_tag" "Controlled variable analysis" 150 176 220 20 9 "#697782" 0)
    [void](Add-ModelIfAvailable $slide "mvp_3d_model_back" $modelPath 705 148 88 72)

    [void](Add-Text $slide "table_title" (U "5a6e6aqM5rC05bmz6KGo") 150 212 120 20 12 "#3F4B53" 1)
    $tableTop = 240
    $col1 = 150
    $col2 = 214
    $col3 = 318
    [void](Add-HRule $slide "table_top" 148 $tableTop 386 "#B7C3CC")
    [void](Add-Text $slide "table_h1" (U "5rC05bmz") $col1 ($tableTop + 10) 52 18 10 "#3F4B53" 1)
    [void](Add-Text $slide "table_h2" (U "5Y+Y6YeP5Y+W5YC8") $col2 ($tableTop + 10) 86 18 10 "#3F4B53" 1)
    [void](Add-Text $slide "table_h3" "μ" $col3 ($tableTop + 10) 48 18 10 "#3F4B53" 1)
    $rows = @(
        @("T1", (U "LTE1IOKEgw=="), "0.118"),
        @("T2", (U "LTEwIOKEgw=="), "0.103"),
        @("T3", (U "LTUg4oSD"), "0.089"),
        @("T4", (U "MCDihIM="), "0.076")
    )
    for ($i = 0; $i -lt $rows.Count; $i++) {
        $y = $tableTop + 36 + ($i * 25)
        [void](Add-HRule $slide "table_row_$i" 148 ($y - 5) 386 "#D6DEE4")
        [void](Add-Text $slide "table_c1_$i" $rows[$i][0] $col1 $y 52 18 10 "#25313A" 0)
        [void](Add-Text $slide "table_c2_$i" $rows[$i][1] $col2 $y 86 18 10 "#25313A" 0)
        [void](Add-Text $slide "table_c3_$i" $rows[$i][2] $col3 $y 48 18 10 "#25313A" 0)
    }

    [void](Add-TrendChart $slide "trend_chart" 470 238 250 150)
    [void](Add-Text $slide "conclusion" (U "57uT6K6677ya5rip5bqm5Y2H6auY5pe277yM5Yaw6KGo6Z2i5ray6Iac5aKe5by677yM5pGp5pOm57O75pWwIM68IOaVtOS9k+S4i+mZjeOAgg==") 150 392 510 28 11 "#1F2A33" 0)
    $return = Add-Text $slide "return_button" (U "54K55Ye76L+U5Zue") 730 392 70 18 10 "#3E6B82" 1
    $return.ActionSettings(1).Action = $ppActionPreviousSlide
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

    $slide1 = $presentation.Slides.Add(1, $ppLayoutBlank)
    $slide1.Name = "morph_front"
    $slide2 = $presentation.Slides.Add(2, $ppLayoutBlank)
    $slide2.Name = "morph_back_analysis"

    foreach ($slide in @($slide1, $slide2)) {
        $bg = $slide.Shapes.AddShape($msoShapeRectangle, 0, 0, 960, 540)
        $bg.Name = "mvp_background_texture_placeholder"
        Set-NoFill $bg
        Set-Line $bg "#E3E9EE" 0.75
        $bg.ZOrder(1) | Out-Null
    }

    Add-FrontSlide $slide1 $modelPath
    Add-BackSlide $slide2 $modelPath

    $slide1.SlideShowTransition.EntryEffect = $ppEffectMorphByObject
    $slide1.SlideShowTransition.Duration = 0.75
    $slide2.SlideShowTransition.EntryEffect = $ppEffectMorphByObject
    $slide2.SlideShowTransition.Duration = 0.75

    $presentation.SaveAs($outPath, $ppSaveAsOpenXMLPresentation)
    Write-Output "Generated $outPath"
}
finally {
    if ($presentation -ne $null) {
        $presentation.Close()
    }
    $app.Quit()
}
