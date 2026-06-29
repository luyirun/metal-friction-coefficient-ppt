$ErrorActionPreference = "Stop"

Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint
Add-Type -AssemblyName Office
Add-Type -AssemblyName System.Drawing

$msoTrue = -1
$msoFalse = 0
$msoTextOrientationHorizontal = 1
$msoShapeRectangle = 1
$msoShapeRoundedRectangle = 5
$msoShapeOval = 9
$ppLayoutBlank = 12
$ppSaveAsOpenXMLPresentation = 24

$outDir = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String("Rjpc5qGM6Z2iXEtleW5vdGXnvKnnlaXlm75c5Zug57Sg55+p6Zi1"))
$assetDir = Join-Path $outDir "assets"
$previewDir = Join-Path $outDir "preview"
$outPath = Join-Path $outDir "factor_matrix_interactive_keynote_3x2.pptx"
$bgPath = Join-Path $assetDir "ice_blue_abstract_texture.png"
New-Item -ItemType Directory -Force -Path $outDir, $assetDir, $previewDir | Out-Null

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
    try {
        $shape.Fill.Background()
    }
    catch {
        $shape.Fill.Visible = $msoFalse
    }
    $shape.Fill.Visible = $msoFalse
}

function Set-Line($shape, [string]$hex, [double]$weight, [double]$transparency = 0) {
    $shape.Line.Visible = $msoTrue
    $shape.Line.ForeColor.RGB = Convert-HexColor $hex
    $shape.Line.Weight = $weight
    $shape.Line.Transparency = $transparency
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

function Add-LineShape($slide, [string]$name, [double]$x1, [double]$y1, [double]$x2, [double]$y2, [string]$hex, [double]$weight, [double]$transparency = 0) {
    $line = $slide.Shapes.AddLine($x1, $y1, $x2, $y2)
    $line.Name = $name
    $line.Line.ForeColor.RGB = Convert-HexColor $hex
    $line.Line.Weight = $weight
    $line.Line.Transparency = $transparency
    return $line
}

function Add-Panel($slide, [string]$name, [double]$left, [double]$top, [double]$width, [double]$height, [string]$lineHex, [double]$lineWeight = 1.2) {
    $shape = $slide.Shapes.AddShape($msoShapeRoundedRectangle, $left, $top, $width, $height)
    $shape.Name = $name
    Set-NoFill $shape
    Set-Line $shape $lineHex $lineWeight 0.08
    $shape.Shadow.Visible = $msoTrue
    $shape.Shadow.ForeColor.RGB = Convert-HexColor "#9AB7C9"
    $shape.Shadow.Transparency = 0.72
    $shape.Shadow.Blur = 8
    $shape.Shadow.OffsetX = 1.2
    $shape.Shadow.OffsetY = 1.6
    $shape.ThreeD.Visible = $msoTrue
    $shape.ThreeD.Depth = 4
    $shape.ThreeD.BevelTopType = [Microsoft.Office.Core.MsoBevelType]::msoBevelSoftRound
    $shape.ThreeD.BevelTopInset = 2
    $shape.ThreeD.BevelTopDepth = 1
    $shape.ThreeD.SetPresetCamera([Microsoft.Office.Core.MsoPresetCamera]::msoCameraLegacyPerspectiveFront)
    return $shape
}

function Add-Button($slide, [string]$name, [string]$text, [double]$left, [double]$top, [double]$width, [double]$height) {
    $button = $slide.Shapes.AddShape($msoShapeRoundedRectangle, $left, $top, $width, $height)
    $button.Name = $name
    Set-NoFill $button
    Set-Line $button "#005A99" 1.2 0
    $button.TextFrame2.TextRange.Text = $text
    $button.TextFrame2.MarginLeft = 7
    $button.TextFrame2.MarginRight = 7
    $button.TextFrame2.MarginTop = 3
    $button.TextFrame2.MarginBottom = 3
    $button.TextFrame2.TextRange.Font.NameFarEast = "Microsoft YaHei"
    $button.TextFrame2.TextRange.Font.Name = "Aptos"
    $button.TextFrame2.TextRange.Font.Size = 9.5
    $button.TextFrame2.TextRange.Font.Bold = $msoTrue
    $button.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Convert-HexColor "#00447A"
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

function Add-WipeWithPrevious($sequence, $shape, [double]$duration) {
    $effect = Add-WithPreviousEffect $sequence $shape ([Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectWipe) $false $duration
    try {
        $effect.EffectParameters.Direction = [Microsoft.Office.Interop.PowerPoint.MsoAnimDirection]::msoAnimDirectionRight
    }
    catch {}
    return $effect
}

function New-Texture([string]$path) {
    $width = 1440
    $height = 960
    $bmp = New-Object System.Drawing.Bitmap $width, $height
    $gfx = [System.Drawing.Graphics]::FromImage($bmp)
    $gfx.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $rect = New-Object System.Drawing.Rectangle 0, 0, $width, $height
    $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush $rect, ([System.Drawing.Color]::FromArgb(246,251,255)), ([System.Drawing.Color]::FromArgb(215,235,248)), 35
    $gfx.FillRectangle($brush, $rect)
    $brush.Dispose()

    $pen1 = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(55, 145, 190, 220)), 1.2
    $pen2 = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(36, 255, 122, 26)), 1.0
    for ($i = -220; $i -lt 1550; $i += 75) {
        $gfx.DrawLine($pen1, $i, 0, $i + 520, $height)
    }
    for ($i = 0; $i -lt 12; $i++) {
        $x = 60 + ($i * 118)
        $gfx.DrawLine($pen2, $x, 78, $x + 46, 78)
    }
    $pen1.Dispose()
    $pen2.Dispose()
    $mist = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(72, 255, 255, 255))
    $gfx.FillEllipse($mist, -180, -120, 700, 320)
    $gfx.FillEllipse($mist, 740, 40, 820, 380)
    $mist.Dispose()
    $gfx.Dispose()
    $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
}

function Add-FactorCard($slide, [hashtable]$factor) {
    $l = $factor.Left
    $t = $factor.Top
    $w = 230
    $h = 100
    $items = @()
    $items += Add-Panel $slide ("matrix_panel_" + $factor.Key) $l $t $w $h "#0077A3" 1.1
    $items += Add-Text $slide ("matrix_no_" + $factor.Key) $factor.No ($l + 16) ($t + 14) 38 18 12 "#00447A" 1
    $items += Add-Text $slide ("matrix_title_" + $factor.Key) $factor.Name ($l + 58) ($t + 14) 126 22 15 "#003B73" 1
    $items += Add-Text $slide ("matrix_en_" + $factor.Key) $factor.En ($l + 58) ($t + 40) 142 16 8.5 "#597284" 0
    $items += Add-Text $slide ("matrix_hint_" + $factor.Key) (U "54K55Ye75p+l55yL") ($l + 150) ($t + 71) 60 16 8.5 "#005A99" 1
    return Group-Shapes $slide ("matrix_card_" + $factor.Key) $items
}

function Add-OpenButton($slide, [hashtable]$factor) {
    return Add-Button $slide ("open_button_" + $factor.Key) (U "54K55Ye75p+l55yL") ($factor.Left + 146) ($factor.Top + 68) 68 22
}

function Add-TrendChart($slide, [string]$prefix, [double]$left, [double]$top, [object[]]$rows) {
    $items = @()
    $items += Add-Text $slide "$prefix`_title" (U "6LaL5Yq/5Zu+") $left ($top - 24) 98 18 12 "#003B73" 1
    $x0 = $left + 20
    $y0 = $top + 142
    $x1 = $left + 286
    $y1 = $top + 20
    $items += Add-LineShape $slide "$prefix`_axis_x" $x0 $y0 $x1 $y0 "#8FAFC2" 1 0.08
    $items += Add-LineShape $slide "$prefix`_axis_y" $x0 $y0 $x0 $y1 "#8FAFC2" 1 0.08
    $items += Add-Text $slide "$prefix`_high" (U "6auY") ($x0 - 16) ($y1 - 4) 14 14 8 "#597284" 0
    $items += Add-Text $slide "$prefix`_low" (U "5L2O") ($x0 - 16) ($y0 - 8) 14 14 8 "#597284" 0

    $values = @()
    foreach ($row in $rows) { $values += [double]$row[2] }
    $min = ($values | Measure-Object -Minimum).Minimum
    $max = ($values | Measure-Object -Maximum).Maximum
    if ($max -eq $min) { $max = $min + 0.01 }
    $points = @()
    for ($i = 0; $i -lt $rows.Count; $i++) {
        $x = $x0 + 28 + (($x1 - $x0 - 54) * $i / [Math]::Max(1, $rows.Count - 1))
        $norm = ([double]$rows[$i][2] - $min) / ($max - $min)
        $y = $y0 - 16 - (($y0 - $y1 - 30) * $norm)
        $points += ,@($x, $y, $rows[$i][2])
    }
    for ($i = 0; $i -lt ($points.Count - 1); $i++) {
        $items += Add-LineShape $slide "$prefix`_line_$i" $points[$i][0] $points[$i][1] $points[$i + 1][0] $points[$i + 1][1] "#005A99" 2 0
    }
    for ($i = 0; $i -lt $points.Count; $i++) {
        $dot = $slide.Shapes.AddShape($msoShapeOval, $points[$i][0] - 4, $points[$i][1] - 4, 8, 8)
        $dot.Name = "$prefix`_dot_$i"
        Set-NoFill $dot
        Set-Line $dot "#005A99" 1.6 0
        $items += $dot
        $items += Add-Text $slide "$prefix`_label_$i" ([string]$points[$i][2]) ($points[$i][0] - 15) ($points[$i][1] - 18) 38 10 7 "#435B6B" 0
    }
    return Group-Shapes $slide $prefix $items
}

function Add-AnalysisLayer($slide, [hashtable]$factor) {
    $l = 92
    $t = 154
    $w = 776
    $h = 382
    $items = @()
    $items += Add-Panel $slide ("analysis_panel_" + $factor.Key) $l $t $w $h "#0077A3" 1.2
    $items += Add-Text $slide ("analysis_title_" + $factor.Key) $factor.AnalysisTitle ($l + 30) ($t + 24) 260 26 19 "#003B73" 1
    $items += Add-Text $slide ("analysis_sub_" + $factor.Key) $factor.En ($l + 32) ($t + 54) 230 16 8.5 "#597284" 0
    $items += Add-Text $slide ("analysis_table_title_" + $factor.Key) (U "5a6e6aqM5rC05bmz6KGo") ($l + 34) ($t + 92) 128 20 12 "#003B73" 1
    $tableTop = $t + 124
    $col1 = $l + 36
    $col2 = $l + 104
    $col3 = $l + 226
    $items += Add-LineShape $slide ("analysis_table_top_" + $factor.Key) ($l + 34) $tableTop ($l + 326) $tableTop "#9EBCCE" 0.8 0.1
    $items += Add-Text $slide ("analysis_h1_" + $factor.Key) (U "5rC05bmz") $col1 ($tableTop + 10) 52 16 9 "#294B5E" 1
    $items += Add-Text $slide ("analysis_h2_" + $factor.Key) (U "5Y+Y6YeP5Y+W5YC8") $col2 ($tableTop + 10) 94 16 9 "#294B5E" 1
    $items += Add-Text $slide ("analysis_h3_" + $factor.Key) "μ" $col3 ($tableTop + 10) 42 16 9 "#294B5E" 1
    for ($i = 0; $i -lt $factor.Rows.Count; $i++) {
        $row = $factor.Rows[$i]
        $y = $tableTop + 36 + ($i * 23)
        $items += Add-LineShape $slide ("analysis_row_" + $factor.Key + "_" + $i) ($l + 34) ($y - 5) ($l + 326) ($y - 5) "#C9DCE8" 0.6 0.12
        $items += Add-Text $slide ("analysis_c1_" + $factor.Key + "_" + $i) $row[0] $col1 $y 52 15 8.5 "#263E4E" 0
        $items += Add-Text $slide ("analysis_c2_" + $factor.Key + "_" + $i) $row[1] $col2 $y 108 15 8.5 "#263E4E" 0
        $items += Add-Text $slide ("analysis_c3_" + $factor.Key + "_" + $i) $row[2] $col3 $y 42 15 8.5 "#263E4E" 0
    }
    $chart = Add-TrendChart $slide ("analysis_chart_" + $factor.Key) ($l + 420) ($t + 132) $factor.Rows
    $items += $chart
    $items += Add-Text $slide ("analysis_conclusion_label_" + $factor.Key) (U "57uT6K66") ($l + 34) ($t + 306) 48 18 10 "#003B73" 1
    $items += Add-Text $slide ("analysis_conclusion_" + $factor.Key) $factor.Conclusion ($l + 88) ($t + 306) 560 34 10 "#263E4E" 0
    $items += Add-Text $slide ("analysis_note_" + $factor.Key) (U "5rOo77ya5b2T5YmN5pWw5o2u5Li654mI5byP56S65L6L77yM5ZCO57ut5Y+v5pu/5o2i5Li655yf5a6e5a6e6aqM5pWw5o2u44CC") ($l + 34) ($t + 346) 560 16 7.5 "#6D8291" 0
    return Group-Shapes $slide ("analysis_group_" + $factor.Key) $items
}

function Add-ReturnButton($slide, [hashtable]$factor) {
    return Add-Button $slide ("return_button_" + $factor.Key) (U "6L+U5Zue55+p6Zi1") 736 488 92 26
}

New-Texture $bgPath

$factors = @(
    @{ Key="temperature"; No="F01"; Name=(U "5rip5bqm"); En="Temperature"; Left=86; Top=190; AnalysisTitle=(U "5rip5bqm5Y+Y6YeP5YiG5p6Q"); Rows=@(@("T1",(U "LTE1IOKEgw=="),"0.118"),@("T2",(U "LTEwIOKEgw=="),"0.103"),@("T3",(U "LTUg4oSD"),"0.089"),@("T4",(U "MCDihIM="),"0.076")); Conclusion=(U "5rip5bqm5Y2H6auY5pe277yM5Yaw6KGo6Z2i5ray6Iac5aKe5by677yM5pGp5pOm57O75pWwIM68IOaVtOS9k+S4i+mZjeOAgg==") },
    @{ Key="pressure"; No="F02"; Name=(U "5Y6L5by6"); En="Pressure"; Left=365; Top=190; AnalysisTitle=(U "5Y6L5by65Y+Y6YeP5YiG5p6Q"); Rows=@(@("P1","0.05 MPa","0.082"),@("P2","0.10 MPa","0.096"),@("P3","0.20 MPa","0.091")); Conclusion=(U "5Y6L5by65aKe5aSn5Lya5pS55Y+Y5bGA6YOo5o6l6Kem54q25oCB5LiO5Y6L5Yqb54aU5YyW56iL5bqm77yM5pGp5pOm57O75pWw5Y+v6IO95ZGI546w5YWI5Y2H5ZCO6ZmN6LaL5Yq/44CC") },
    @{ Key="material"; No="F03"; Name=(U "6YeR5bGe5p2Q6LSo"); En="Metal Material"; Left=644; Top=190; AnalysisTitle=(U "6YeR5bGe5p2Q6LSo5Y+Y6YeP5YiG5p6Q"); Rows=@(@("M1",(U "6ZOd5ZCI6YeR"),"0.108"),@("M2",(U "5LiN6ZSI6ZKi"),"0.095"),@("M3",(U "6ZKb5ZCI6YeR"),"0.083")); Conclusion=(U "5LiN5ZCM6YeR5bGe5p2Q5paZ55qE5a+854Ot5oCn44CB6KGo6Z2i6IO95ZKM5b6u6KeC5o6l6Kem54q25oCB5LiN5ZCM77yM5a+86Ie05pGp5pOm57O75pWw5a2Y5Zyo5pi+6JGX5beu5byC44CC") },
    @{ Key="roughness"; No="F04"; Name=(U "6KGo6Z2i57KX57OZ5bqm"); En="Surface Roughness"; Left=86; Top=324; AnalysisTitle=(U "6KGo6Z2i57KX57OZ5bqm5Y+Y6YeP5YiG5p6Q"); Rows=@(@("R1",(U "UmEgMC4yIM68bQ=="),"0.074"),@("R2",(U "UmEgMS4wIM68bQ=="),"0.096"),@("R3",(U "UmEgMy4wIM68bQ=="),"0.124")); Conclusion=(U "6KGo6Z2i57KX57OZ5bqm5aKe5aSn5pe277yM5b6u6KeC5bWM5YWl5ZKM54qB5YmK5pWI5bqU5aKe5by677yM5pGp5pOm57O75pWwIM68IOS4iuWNh+OAgg==") },
    @{ Key="area"; No="F05"; Name=(U "5o6l6Kem6Z2i56ev"); En="Contact Area"; Left=365; Top=324; AnalysisTitle=(U "5o6l6Kem6Z2i56ev5Y+Y6YeP5YiG5p6Q"); Rows=@(@("A1",(U "MSBjbcKy"),"0.112"),@("A2",(U "NCBjbcKy"),"0.096"),@("A3",(U "OSBjbcKy"),"0.087")); Conclusion=(U "5Zyo5YW25LuW5p2h5Lu25LiN5Y+Y5pe277yM5o6l6Kem6Z2i56ev5Y+Y5YyW5Lya5pS55Y+Y5Y2V5L2N5Y6L5by65ZKM5bGA6YOo5ray6Iac54q25oCB77yM5L2/5pGp5pOm57O75pWw5Y+R55Sf5Y+Y5YyW44CC") },
    @{ Key="ice"; No="F06"; Name=(U "5Yaw55qE5oiQ5YiG"); En="Ice Composition"; Left=644; Top=324; AnalysisTitle=(U "5Yaw55qE5oiQ5YiG5Y+Y6YeP5YiG5p6Q"); Rows=@(@("I1",(U "5reh5rC05Yaw"),"0.104"),@("I2",(U "5rW35Yaw"),"0.089"),@("I3",(U "5Lq65bel55uQ5Yaw"),"0.078")); Conclusion=(U "5Yaw5Lit55uQ5YiG5oiW5p2C6LSo5Lya5pS55Y+Y5Yaw54K55LiO6KGo6Z2i5ray6Iac54q25oCB77yM5LuO6ICM5b2x5ZON6YeR5bGeLeWGsOeVjOmdoueahOaRqeaTpuihjOS4uuOAgg==") }
)

$app = New-Object -ComObject PowerPoint.Application
$app.Visible = $msoTrue
$presentation = $null
try {
    if (Test-Path -LiteralPath $outPath) {
        Remove-Item -LiteralPath $outPath -Force
    }
    $presentation = $app.Presentations.Add($msoTrue)
    $presentation.PageSetup.SlideWidth = 960
    $presentation.PageSetup.SlideHeight = 640
    $slide = $presentation.Slides.Add(1, $ppLayoutBlank)
    $slide.Name = "factor_matrix_interactive"
    $slide.SlideShowTransition.AdvanceOnClick = $msoFalse
    $slide.SlideShowTransition.AdvanceOnTime = $msoFalse

    $bg = $slide.Shapes.AddPicture($bgPath, $msoFalse, $msoTrue, 0, 0, 960, 640)
    $bg.Name = "ice_blue_abstract_texture_background"
    $bg.ZOrder(1) | Out-Null

    [void](Add-Text $slide "slide_index" "09" 54 42 94 58 40 "#003B73" 1)
    [void](Add-LineShape $slide "index_rule" 56 103 132 103 "#FF7A1A" 3 0)
    [void](Add-Text $slide "main_title" (U "5Zug57Sg55+p6Zi1") 154 52 260 38 28 "#003B73" 1)
    [void](Add-Text $slide "main_subtitle" (U "5o6n5Yi25Y+Y6YeP5rOV5LiL55qE5pGp5pOm57O75pWw5b2x5ZON5Zug57Sg5oC76KeI") 156 94 560 22 12 "#293E52" 1)
    [void](Add-Text $slide "main_hint" (U "54K55Ye75Lu75LiA5Zug57Sg5p+l55yL5a6e6aqM5rC05bmz44CB6LaL5Yq/5Zu+5LiO54mp55CG57uT6K66") 156 122 560 18 9 "#597284" 0)

    $matrixItems = @()
    $matrixItems += Add-LineShape $slide "matrix_line_h1" 160 306 800 306 "#8DB2C7" 1 0.25
    $matrixItems += Add-LineShape $slide "matrix_line_v1" 305 235 305 420 "#8DB2C7" 1 0.25
    $matrixItems += Add-LineShape $slide "matrix_line_v2" 585 235 585 420 "#8DB2C7" 1 0.25
    $matrixItems += Add-Text $slide "matrix_center_1" (U "5o6n5Yi25Y+Y6YeP") 412 276 120 18 10 "#003B73" 1
    $matrixItems += Add-Text $slide "matrix_center_2" (U "5a6e6aqM5rC05bmz") 412 300 120 18 10 "#003B73" 1
    $matrixItems += Add-Text $slide "matrix_center_3" (U "6LaL5Yq/5bu65qih") 412 324 120 18 10 "#003B73" 1
    $matrixGroup = Group-Shapes $slide "matrix_context_group" $matrixItems

    $cards = @{}
    $buttons = @{}
    foreach ($factor in $factors) {
        $cards[$factor.Key] = Add-FactorCard $slide $factor
        $buttons[$factor.Key] = Add-OpenButton $slide $factor
    }

    $analysis = @{}
    $returns = @{}
    foreach ($factor in $factors) {
        $analysis[$factor.Key] = Add-AnalysisLayer $slide $factor
        $returns[$factor.Key] = Add-ReturnButton $slide $factor
    }

    [void](Add-LineShape $slide "footer_rule" 42 584 918 584 "#8DB2C7" 0.75 0.38)
    [void](Add-Text $slide "footer_title" (U "6YeR5bGe5LiO5Yaw5pGp5pOm57O75pWw5rWL6YeP") 72 598 260 18 10 "#003B73" 1)
    [void](Add-Text $slide "footer_page" "09/10" 842 598 72 20 14 "#003B73" 1)

    foreach ($factor in $factors) {
        $seq = $slide.TimeLine.InteractiveSequences.Add()
        $trigger = $buttons[$factor.Key]
        [void](Add-TriggeredEffect $seq $matrixGroup $trigger ([Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectFade) $true 0.16)
        foreach ($other in $factors) {
            [void](Add-WithPreviousEffect $seq $cards[$other.Key] ([Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectFade) $true 0.14)
            [void](Add-WithPreviousEffect $seq $buttons[$other.Key] ([Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectFade) $true 0.08)
        }
        [void](Add-WipeWithPrevious $seq $analysis[$factor.Key] 0.34)
        [void](Add-WithPreviousEffect $seq $returns[$factor.Key] ([Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectFade) $false 0.08)

        $returnSeq = $slide.TimeLine.InteractiveSequences.Add()
        $returnTrigger = $returns[$factor.Key]
        [void](Add-TriggeredEffect $returnSeq $analysis[$factor.Key] $returnTrigger ([Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectFade) $true 0.16)
        [void](Add-WithPreviousEffect $returnSeq $returns[$factor.Key] ([Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectFade) $true 0.08)
        [void](Add-WipeWithPrevious $returnSeq $matrixGroup 0.24)
        foreach ($other in $factors) {
            [void](Add-WithPreviousEffect $returnSeq $cards[$other.Key] ([Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectFade) $false 0.12)
            [void](Add-WithPreviousEffect $returnSeq $buttons[$other.Key] ([Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectFade) $false 0.08)
        }
    }

    $presentation.SaveAs($outPath, $ppSaveAsOpenXMLPresentation)
    Write-Output "Generated $outPath"
}
finally {
    if ($presentation -ne $null) {
        try { $presentation.Close() } catch {}
    }
    try { $app.Quit() } catch {}
}
