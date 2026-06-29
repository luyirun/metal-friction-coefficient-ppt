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
$ppFixedFormatTypePNG = 2

$outDir = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String("Rjpc5qGM6Z2iXEtleW5vdGXnvKnnlaXlm75c5Zug57Sg55+p6Zi1"))
$assetDir = Join-Path $outDir "assets"
$previewDir = Join-Path $outDir "preview"
$outPath = Join-Path $outDir "factor_matrix_overview_interactive_keynote_3x2.pptx"
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
    try { $shape.Fill.Background() } catch { $shape.Fill.Visible = $msoFalse }
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

function Add-Panel($slide, [string]$name, [double]$left, [double]$top, [double]$width, [double]$height, [string]$lineHex, [double]$lineWeight = 1.1) {
    $shape = $slide.Shapes.AddShape($msoShapeRoundedRectangle, $left, $top, $width, $height)
    $shape.Name = $name
    Set-NoFill $shape
    Set-Line $shape $lineHex $lineWeight 0.04
    $shape.Shadow.Visible = $msoTrue
    $shape.Shadow.ForeColor.RGB = Convert-HexColor "#7FA8C2"
    $shape.Shadow.Transparency = 0.75
    $shape.Shadow.Blur = 8
    $shape.Shadow.OffsetX = 1.2
    $shape.Shadow.OffsetY = 1.6
    return $shape
}

function Add-Button($slide, [string]$name, [string]$text, [double]$left, [double]$top, [double]$width, [double]$height) {
    $button = $slide.Shapes.AddShape($msoShapeRoundedRectangle, $left, $top, $width, $height)
    $button.Name = $name
    Set-NoFill $button
    Set-Line $button "#005A99" 1.05 0
    $button.TextFrame2.TextRange.Text = $text
    $button.TextFrame2.MarginLeft = 6
    $button.TextFrame2.MarginRight = 6
    $button.TextFrame2.MarginTop = 2
    $button.TextFrame2.MarginBottom = 2
    $button.TextFrame2.TextRange.Font.NameFarEast = "Microsoft YaHei"
    $button.TextFrame2.TextRange.Font.Name = "Aptos"
    $button.TextFrame2.TextRange.Font.Size = 8.2
    $button.TextFrame2.TextRange.Font.Bold = $msoTrue
    $button.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Convert-HexColor "#00447A"
    return $button
}

function Group-Shapes($slide, [string]$groupName, [object[]]$shapes) {
    $names = @()
    foreach ($shape in $shapes) { $names += $shape.Name }
    $group = $slide.Shapes.Range($names).Group()
    $group.Name = $groupName
    return $group
}

function Add-WithPreviousEffect($sequence, $shape, $effectType, [bool]$exit, [double]$duration) {
    $effect = $sequence.AddEffect($shape, $effectType, [Microsoft.Office.Interop.PowerPoint.MsoAnimateByLevel]::msoAnimateLevelNone, [Microsoft.Office.Interop.PowerPoint.MsoAnimTriggerType]::msoAnimTriggerWithPrevious)
    $effect.Timing.Duration = $duration
    if ($exit) { $effect.Exit = $msoTrue }
    return $effect
}

function Add-TriggeredEffect($sequence, $shape, $triggerShape, $effectType, [bool]$exit, [double]$duration) {
    $effect = $sequence.AddEffect($shape, $effectType, [Microsoft.Office.Interop.PowerPoint.MsoAnimateByLevel]::msoAnimateLevelNone, [Microsoft.Office.Interop.PowerPoint.MsoAnimTriggerType]::msoAnimTriggerOnShapeClick)
    $effect.Timing.TriggerShape = $triggerShape
    $effect.Timing.Duration = $duration
    if ($exit) { $effect.Exit = $msoTrue }
    return $effect
}

function Add-WipeWithPrevious($sequence, $shape, [double]$duration) {
    $effect = Add-WithPreviousEffect $sequence $shape ([Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectWipe) $false $duration
    try { $effect.EffectParameters.Direction = [Microsoft.Office.Interop.PowerPoint.MsoAnimDirection]::msoAnimDirectionRight } catch {}
    return $effect
}

function New-Texture([string]$path) {
    if (Test-Path -LiteralPath $path) { return }
    $width = 1440
    $height = 960
    $bmp = New-Object System.Drawing.Bitmap $width, $height
    $gfx = [System.Drawing.Graphics]::FromImage($bmp)
    $gfx.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $rect = New-Object System.Drawing.Rectangle 0, 0, $width, $height
    $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush $rect, ([System.Drawing.Color]::FromArgb(247,252,255)), ([System.Drawing.Color]::FromArgb(214,235,248)), 30
    $gfx.FillRectangle($brush, $rect)
    $brush.Dispose()
    $pen1 = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(55, 118, 175, 212)), 1.2
    $pen2 = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(38, 255, 122, 26)), 1.0
    for ($i = -300; $i -lt 1550; $i += 62) { $gfx.DrawLine($pen1, $i, 0, $i + 560, $height) }
    for ($i = 0; $i -lt 13; $i++) {
        $x = 65 + ($i * 112)
        $gfx.DrawLine($pen2, $x, 82, $x + 42, 82)
    }
    $mist = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(78, 255, 255, 255))
    $gfx.FillEllipse($mist, -200, -130, 720, 350)
    $gfx.FillEllipse($mist, 700, 20, 850, 430)
    $mist.Dispose()
    $pen1.Dispose()
    $pen2.Dispose()
    $gfx.Dispose()
    $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
}

function Add-FactorIcon($slide, [string]$prefix, [string]$key, [double]$left, [double]$top) {
    $items = @()
    if ($key -eq "temperature") {
        $items += Add-LineShape $slide "$prefix`_icon_stem" ($left + 18) ($top + 8) ($left + 18) ($top + 42) "#005A99" 1.8 0
        $dot = $slide.Shapes.AddShape($msoShapeOval, $left + 10, $top + 37, 16, 16)
        $dot.Name = "$prefix`_icon_dot"
        Set-NoFill $dot; Set-Line $dot "#005A99" 1.8 0; $items += $dot
        $items += Add-LineShape $slide "$prefix`_icon_snow_1" ($left + 34) ($top + 25) ($left + 48) ($top + 25) "#005A99" 1.2 0.1
        $items += Add-LineShape $slide "$prefix`_icon_snow_2" ($left + 41) ($top + 18) ($left + 41) ($top + 32) "#005A99" 1.2 0.1
    } elseif ($key -eq "pressure") {
        $dial = $slide.Shapes.AddShape($msoShapeOval, $left + 8, $top + 10, 42, 42)
        $dial.Name = "$prefix`_icon_dial"
        Set-NoFill $dial; Set-Line $dial "#005A99" 1.8 0; $items += $dial
        $items += Add-LineShape $slide "$prefix`_icon_needle" ($left + 29) ($top + 31) ($left + 42) ($top + 20) "#005A99" 1.8 0
        $items += Add-LineShape $slide "$prefix`_icon_base" ($left + 18) ($top + 54) ($left + 42) ($top + 54) "#005A99" 1.6 0
    } elseif ($key -eq "roughness") {
        for ($i = 0; $i -lt 7; $i++) {
            $x1 = $left + 8 + ($i * 7)
            $x2 = $x1 + 5
            $y1 = $top + 42 - (($i % 2) * 26)
            $y2 = $top + 16 + (($i % 2) * 24)
            $items += Add-LineShape $slide "$prefix`_icon_wave_$i" $x1 $y1 $x2 $y2 "#005A99" 1.4 0
        }
        $items += Add-LineShape $slide "$prefix`_icon_ground" ($left + 6) ($top + 54) ($left + 58) ($top + 54) "#005A99" 1.5 0
    } elseif ($key -eq "area") {
        $box = $slide.Shapes.AddShape($msoShapeRoundedRectangle, $left + 8, $top + 8, 44, 44)
        $box.Name = "$prefix`_icon_area_box"
        Set-NoFill $box; Set-Line $box "#005A99" 1.6 0; $items += $box
        for ($r = 0; $r -lt 4; $r++) {
            for ($c = 0; $c -lt 4; $c++) {
                $d = $slide.Shapes.AddShape($msoShapeOval, $left + 17 + ($c * 8), $top + 17 + ($r * 8), 2.5, 2.5)
                $d.Name = "$prefix`_icon_area_dot_$r`_$c"
                Set-NoFill $d; Set-Line $d "#005A99" 1.1 0; $items += $d
            }
        }
    } else {
        $items += Add-LineShape $slide "$prefix`_icon_cube_1" ($left + 14) ($top + 44) ($left + 34) ($top + 54) "#005A99" 1.5 0
        $items += Add-LineShape $slide "$prefix`_icon_cube_2" ($left + 34) ($top + 54) ($left + 54) ($top + 42) "#005A99" 1.5 0
        $items += Add-LineShape $slide "$prefix`_icon_cube_3" ($left + 14) ($top + 44) ($left + 34) ($top + 30) "#005A99" 1.5 0
        $items += Add-LineShape $slide "$prefix`_icon_cube_4" ($left + 34) ($top + 30) ($left + 54) ($top + 42) "#005A99" 1.5 0
        $items += Add-LineShape $slide "$prefix`_icon_cube_5" ($left + 34) ($top + 30) ($left + 34) ($top + 54) "#005A99" 1.5 0
    }
    return Group-Shapes $slide "$prefix`_icon" $items
}

function Add-FactorCard($slide, [hashtable]$factor) {
    $l = $factor.Left
    $t = $factor.Top
    $items = @()
    $items += Add-Panel $slide ("matrix_panel_" + $factor.Key) $l $t 142 136 "#3F7EBD" 0.85
    $items += Add-Text $slide ("matrix_no_" + $factor.Key) $factor.No ($l + 14) ($t + 17) 40 17 9 "#00447A" 1
    $items += Add-Text $slide ("matrix_title_" + $factor.Key) $factor.Name ($l + 14) ($t + 44) 88 26 16 "#003B73" 1
    $items += Add-Text $slide ("matrix_en_" + $factor.Key) $factor.En ($l + 14) ($t + 74) 94 16 8.2 "#4A6E87" 0
    $items += Add-FactorIcon $slide ("matrix_" + $factor.Key) $factor.Key ($l + 92) ($t + 44)
    $items += Add-Button $slide ("open_button_" + $factor.Key) (U "54K55Ye75p+l55yL") ($l + 60) ($t + 102) 68 23
    return Group-Shapes $slide ("matrix_card_" + $factor.Key) $items
}

function Add-Highlight($slide, [hashtable]$factor) {
    $items = @()
    $l = $factor.Left
    $t = $factor.Top
    $items += Add-Panel $slide ("highlight_panel_" + $factor.Key) $l $t 142 136 "#00A5FF" 2.2
    $items += Add-LineShape $slide ("highlight_rule_" + $factor.Key) ($l + 12) ($t + 134) ($l + 130) ($t + 134) "#00A5FF" 2.2 0
    return Group-Shapes $slide ("highlight_group_" + $factor.Key) $items
}

function Add-OverviewPanel($slide) {
    $items = @()
    $l = 522; $t = 128; $w = 410; $h = 432
    $items += Add-Panel $slide "overview_panel" $l $t $w $h "#70A8D0" 1.1
    $items += Add-Text $slide "overview_title" (U "5Zug57Sg5oC76KeI") ($l + 28) ($t + 24) 160 26 18 "#003B73" 1
    $items += Add-Text $slide "overview_sub" "Overview" ($l + 30) ($t + 53) 110 16 8.5 "#4A6E87" 0
    $items += Add-Text $slide "overview_mark" "OVERVIEW" ($l + 220) ($t + 18) 150 44 30 "#90B7D6" 1

    $items += Add-Text $slide "overview_path_title" (U "5pa55rOV6Lev5b6E") ($l + 26) ($t + 88) 90 18 12 "#003B73" 1
    $pathY = $t + 122
    $pathLabels = @((U "5o6n5Yi25Y+Y6YeP"), (U "5Y2V5Zug57Sg5Y+Y5YyW"), (U "5pGp5pOm57O75pWwIM68"), (U "5a+55q+U57uT6K66"))
    for ($i = 0; $i -lt 4; $i++) {
        $x = $l + 28 + ($i * 88)
        $items += Add-Panel $slide ("overview_path_box_" + $i) $x $pathY 72 34 "#8FB9D6" 0.8
        $items += Add-Text $slide ("overview_path_text_" + $i) $pathLabels[$i] ($x + 6) ($pathY + 9) 60 14 7.8 "#003B73" 1
        if ($i -lt 3) { $items += Add-LineShape $slide ("overview_path_arrow_" + $i) ($x + 76) ($pathY + 17) ($x + 86) ($pathY + 17) "#0077A3" 1.2 0 }
    }

    $items += Add-Text $slide "overview_group_title" (U "5Y+Y6YeP5YiG57uE") ($l + 26) ($t + 184) 90 18 12 "#003B73" 1
    $groups = @(
        @((U "546v5aKD5p2h5Lu2"), (U "5rip5bqm44CB5Yaw55qE5oiQ5YiG")),
        @((U "5o6l6Kem54q25oCB"), (U "5Y6L5by644CB5o6l6Kem6Z2i56ev")),
        @((U "5p2Q5paZ6KGo6Z2i"), (U "6YeR5bGe5p2Q6LSo44CB6KGo6Z2i57KX57OZ5bqm"))
    )
    for ($i = 0; $i -lt 3; $i++) {
        $y = $t + 214 + ($i * 40)
        $items += Add-LineShape $slide ("overview_group_line_" + $i) ($l + 28) ($y - 7) ($l + 220) ($y - 7) "#B7D1E4" 0.7 0.12
        $items += Add-Text $slide ("overview_group_name_" + $i) $groups[$i][0] ($l + 30) $y 68 16 8.6 "#003B73" 1
        $items += Add-Text $slide ("overview_group_value_" + $i) $groups[$i][1] ($l + 102) $y 180 16 8.2 "#344E61" 0
    }

    $items += Add-Text $slide "overview_radar_title" (U "5b2x5ZON6Zu36L6+") ($l + 258) ($t + 184) 90 18 12 "#003B73" 1
    $cx = $l + 312; $cy = $t + 272
    $rad = 70
    $pts = @()
    for ($i = 0; $i -lt 6; $i++) {
        $ang = (-90 + $i * 60) * [Math]::PI / 180
        $x = $cx + [Math]::Cos($ang) * $rad
        $y = $cy + [Math]::Sin($ang) * $rad
        $pts += ,@($x, $y)
        $items += Add-LineShape $slide ("overview_radar_axis_" + $i) $cx $cy $x $y "#A8C7DB" 0.75 0.16
    }
    for ($i = 0; $i -lt 6; $i++) {
        $j = ($i + 1) % 6
        $items += Add-LineShape $slide ("overview_radar_outer_" + $i) $pts[$i][0] $pts[$i][1] $pts[$j][0] $pts[$j][1] "#A8C7DB" 0.7 0.18
        $ix = $cx + (($pts[$i][0] - $cx) * (0.46 + (($i % 3) * 0.08)))
        $iy = $cy + (($pts[$i][1] - $cy) * (0.46 + (($i % 3) * 0.08)))
        $jx = $cx + (($pts[$j][0] - $cx) * (0.46 + (($j % 3) * 0.08)))
        $jy = $cy + (($pts[$j][1] - $cy) * (0.46 + (($j % 3) * 0.08)))
        $items += Add-LineShape $slide ("overview_radar_value_" + $i) $ix $iy $jx $jy "#0077A3" 1.7 0
    }
    $items += Add-Text $slide "overview_low" (U "5L2O") ($cx - 78) ($cy + 68) 22 12 7.2 "#597284" 0
    $items += Add-Text $slide "overview_high" (U "6auY") ($cx + 62) ($cy - 76) 24 12 7.2 "#597284" 0

    $items += Add-Text $slide "overview_conclusion_label" (U "5oC76KeI5Yik5pat") ($l + 28) ($t + 350) 70 18 10 "#003B73" 1
    $items += Add-Text $slide "overview_conclusion" (U "5pGp5pOm57O75pWw55Sx5p2Q5paZ5bGe5oCn44CB5Yaw6Z2i54q25oCB5LiO5o6l6Kem5p2h5Lu25YWx5ZCM5Yaz5a6a77yM6ZyA6YCQ6aG55o6n5Yi25Y+Y6YeP5YiG5p6Q44CC") ($l + 28) ($t + 376) 350 32 8.6 "#263E4E" 0
    return Group-Shapes $slide "overview_group" $items
}

function Add-TrendChart($slide, [string]$prefix, [double]$left, [double]$top, [object[]]$rows) {
    $items = @()
    $items += Add-Text $slide "$prefix`_title" (U "6LaL5Yq/5Zu+") $left ($top - 22) 96 18 12 "#003B73" 1
    $x0 = $left + 18; $y0 = $top + 136; $x1 = $left + 250; $y1 = $top + 18
    $items += Add-LineShape $slide "$prefix`_axis_x" $x0 $y0 $x1 $y0 "#8FAFC2" 1 0.08
    $items += Add-LineShape $slide "$prefix`_axis_y" $x0 $y0 $x0 $y1 "#8FAFC2" 1 0.08
    $items += Add-Text $slide "$prefix`_high" (U "6auY") ($x0 - 15) ($y1 - 4) 14 12 7.2 "#597284" 0
    $items += Add-Text $slide "$prefix`_low" (U "5L2O") ($x0 - 15) ($y0 - 7) 14 12 7.2 "#597284" 0
    $values = @()
    foreach ($row in $rows) { $values += [double]$row[2] }
    $min = ($values | Measure-Object -Minimum).Minimum
    $max = ($values | Measure-Object -Maximum).Maximum
    if ($max -eq $min) { $max = $min + 0.01 }
    $points = @()
    for ($i = 0; $i -lt $rows.Count; $i++) {
        $x = $x0 + 24 + (($x1 - $x0 - 44) * $i / [Math]::Max(1, $rows.Count - 1))
        $norm = ([double]$rows[$i][2] - $min) / ($max - $min)
        $y = $y0 - 16 - (($y0 - $y1 - 28) * $norm)
        $points += ,@($x, $y, $rows[$i][2], $rows[$i][0])
    }
    for ($i = 0; $i -lt ($points.Count - 1); $i++) {
        $items += Add-LineShape $slide "$prefix`_line_$i" $points[$i][0] $points[$i][1] $points[$i + 1][0] $points[$i + 1][1] "#005A99" 1.8 0
    }
    for ($i = 0; $i -lt $points.Count; $i++) {
        $dot = $slide.Shapes.AddShape($msoShapeOval, $points[$i][0] - 4, $points[$i][1] - 4, 8, 8)
        $dot.Name = "$prefix`_dot_$i"
        Set-NoFill $dot; Set-Line $dot "#005A99" 1.5 0; $items += $dot
        $items += Add-Text $slide "$prefix`_label_$i" ([string]$points[$i][2]) ($points[$i][0] - 14) ($points[$i][1] - 17) 38 10 6.7 "#435B6B" 0
        $items += Add-Text $slide "$prefix`_x_$i" ([string]$points[$i][3]) ($points[$i][0] - 10) ($y0 + 8) 24 10 6.8 "#435B6B" 0
    }
    return Group-Shapes $slide $prefix $items
}

function Add-DetailPanel($slide, [hashtable]$factor) {
    $l = 522; $t = 128; $w = 410; $h = 432
    $items = @()
    $items += Add-Panel $slide ("detail_panel_" + $factor.Key) $l $t $w $h "#70A8D0" 1.1
    $items += Add-Text $slide ("detail_title_" + $factor.Key) $factor.AnalysisTitle ($l + 28) ($t + 24) 210 28 17 "#003B73" 1
    $items += Add-Text $slide ("detail_code_" + $factor.Key) $factor.No ($l + 202) ($t + 27) 36 18 9 "#00447A" 1
    $items += Add-Text $slide ("detail_en_" + $factor.Key) $factor.En ($l + 30) ($t + 54) 150 16 8.4 "#4A6E87" 0
    $items += Add-Text $slide ("detail_mark_" + $factor.Key) $factor.No ($l + 305) ($t + 20) 74 44 30 "#90B7D6" 1

    $items += Add-Text $slide ("detail_table_title_" + $factor.Key) (U "5a6e6aqM5rC05bmz6KGo") ($l + 30) ($t + 92) 120 18 11.5 "#003B73" 1
    $tableTop = $t + 124
    $col1 = $l + 30; $col2 = $l + 82; $col3 = $l + 182
    $items += Add-LineShape $slide ("detail_table_top_" + $factor.Key) ($l + 28) $tableTop ($l + 246) $tableTop "#9EBCCE" 0.8 0.12
    $items += Add-Text $slide ("detail_h1_" + $factor.Key) (U "5rC05bmz") $col1 ($tableTop + 10) 42 15 8.2 "#294B5E" 1
    $items += Add-Text $slide ("detail_h2_" + $factor.Key) (U "5Y+Y6YeP5Y+W5YC8") $col2 ($tableTop + 10) 86 15 8.2 "#294B5E" 1
    $items += Add-Text $slide ("detail_h3_" + $factor.Key) "μ" $col3 ($tableTop + 10) 38 15 8.2 "#294B5E" 1
    for ($i = 0; $i -lt $factor.Rows.Count; $i++) {
        $row = $factor.Rows[$i]
        $y = $tableTop + 35 + ($i * 25)
        $items += Add-LineShape $slide ("detail_row_" + $factor.Key + "_" + $i) ($l + 28) ($y - 5) ($l + 246) ($y - 5) "#C9DCE8" 0.6 0.18
        $items += Add-Text $slide ("detail_c1_" + $factor.Key + "_" + $i) $row[0] $col1 $y 42 15 8.0 "#263E4E" 0
        $items += Add-Text $slide ("detail_c2_" + $factor.Key + "_" + $i) $row[1] $col2 $y 88 15 8.0 "#263E4E" 0
        $items += Add-Text $slide ("detail_c3_" + $factor.Key + "_" + $i) $row[2] $col3 $y 44 15 8.0 "#263E4E" 0
    }
    $items += Add-TrendChart $slide ("detail_chart_" + $factor.Key) ($l + 266) ($t + 142) $factor.Rows
    $items += Add-LineShape $slide ("detail_con_rule_" + $factor.Key) ($l + 28) ($t + 334) ($l + 374) ($t + 334) "#B7D1E4" 0.8 0.14
    $items += Add-Text $slide ("detail_con_label_" + $factor.Key) (U "57uT6K66") ($l + 30) ($t + 350) 42 16 10 "#003B73" 1
    $items += Add-Text $slide ("detail_con_" + $factor.Key) $factor.Conclusion ($l + 76) ($t + 350) 306 32 8.5 "#263E4E" 0
    $items += Add-Text $slide ("detail_note_" + $factor.Key) (U "5rOo77ya5b2T5YmN5pWw5o2u5Li656S65L6L77yM5ZCO57ut5Y+v5pu/5o2i5Li655yf5a6e5a6e6aqM5pWw5o2u44CC") ($l + 30) ($t + 402) 230 14 6.8 "#6D8291" 0
    $items += Add-Button $slide ("return_button_" + $factor.Key) (U "6L+U5Zue5oC76KeI") ($l + 292) ($t + 392) 86 24
    return Group-Shapes $slide ("detail_group_" + $factor.Key) $items
}

function Set-PreviewState($slide, [string]$state, [object[]]$factors) {
    foreach ($factor in $factors) {
        $slide.Shapes.Item("detail_group_" + $factor.Key).Visible = $msoFalse
        $slide.Shapes.Item("highlight_group_" + $factor.Key).Visible = $msoFalse
    }
    $slide.Shapes.Item("overview_group").Visible = $msoTrue
    if ($state -ne "overview") {
        $slide.Shapes.Item("overview_group").Visible = $msoFalse
        $slide.Shapes.Item("detail_group_" + $state).Visible = $msoTrue
        $slide.Shapes.Item("highlight_group_" + $state).Visible = $msoTrue
    }
}

New-Texture $bgPath

$factors = @(
    @{ Key="temperature"; No="F01"; Name=(U "5rip5bqm"); En="Temperature"; Left=42; Top=228; AnalysisTitle=(U "5rip5bqm5Y+Y6YeP5YiG5p6Q"); Rows=@(@("T1",(U "LTE1IOKEgw=="),"0.118"),@("T2",(U "LTEwIOKEgw=="),"0.103"),@("T3",(U "LTUg4oSD"),"0.089"),@("T4",(U "MCDihIM="),"0.076")); Conclusion=(U "5rip5bqm5Y2H6auY5pe277yM5Yaw6KGo6Z2i5ray6Iac5aKe5by677yM5pGp5pOm57O75pWwIM68IOaVtOS9k+S4i+mZjeOAgg==") },
    @{ Key="pressure"; No="F02"; Name=(U "5Y6L5by6"); En="Pressure"; Left=194; Top=228; AnalysisTitle=(U "5Y6L5by65Y+Y6YeP5YiG5p6Q"); Rows=@(@("P1","0.05 MPa","0.082"),@("P2","0.10 MPa","0.096"),@("P3","0.20 MPa","0.091")); Conclusion=(U "5Y6L5by65aKe5aSn5Lya5pS55Y+Y5bGA6YOo5o6l6Kem54q25oCB5LiO5Y6L5Yqb6J6N5YyW56iL5bqm77yM5pGp5pOm57O75pWw5Y+v6IO95ZGI546w5YWI5Y2H5ZCO6ZmN6LaL5Yq/44CC") },
    @{ Key="material"; No="F03"; Name=(U "6YeR5bGe5p2Q6LSo"); En="Metal Material"; Left=346; Top=228; AnalysisTitle=(U "6YeR5bGe5p2Q6LSo5Y+Y6YeP5YiG5p6Q"); Rows=@(@("M1",(U "6ZOd5ZCI6YeR"),"0.118"),@("M2",(U "5LiN6ZSI6ZKi"),"0.103"),@("M3",(U "6ZOc"),"0.089"),@("M4",(U "6ZKb5ZCI6YeR"),"0.076")); Conclusion=(U "5LiN5ZCM6YeR5bGe5p2Q6LSo55qE5a+854Ot5oCn44CB6KGo6Z2i6IO95ZKM5b6u6KeC5o6l6Kem54q25oCB5LiN5ZCM77yM5a+86Ie05pGp5pOm57O75pWw5a2Y5Zyo5pi+6JGX5beu5byC44CC") },
    @{ Key="roughness"; No="F04"; Name=(U "6KGo6Z2i57KX57OZ5bqm"); En="Surface Roughness"; Left=42; Top=384; AnalysisTitle=(U "6KGo6Z2i57KX57OZ5bqm5Y+Y6YeP5YiG5p6Q"); Rows=@(@("R1",(U "UmEgMC4yIM68bQ=="),"0.074"),@("R2",(U "UmEgMS4wIM68bQ=="),"0.096"),@("R3",(U "UmEgMy4wIM68bQ=="),"0.124")); Conclusion=(U "6KGo6Z2i57KX57OZ5bqm5aKe5aSn5pe277yM5b6u6KeC5bWM5YWl5ZKM54qB5YmK5pWI5bqU5aKe5by677yM5pGp5pOm57O75pWwIM68IOS4iuWNh+OAgg==") },
    @{ Key="area"; No="F05"; Name=(U "5o6l6Kem6Z2i56ev"); En="Contact Area"; Left=194; Top=384; AnalysisTitle=(U "5o6l6Kem6Z2i56ev5Y+Y6YeP5YiG5p6Q"); Rows=@(@("A1",(U "MSBjbcKy"),"0.112"),@("A2",(U "NCBjbcKy"),"0.096"),@("A3",(U "OSBjbcKy"),"0.087")); Conclusion=(U "5o6l6Kem6Z2i56ev5Y+Y5YyW5Lya5pS55Y+Y5Y2V5L2N5Y6L5by65ZKM5bGA6YOo5ray6Iac54q25oCB77yM5L2/5pGp5pOm57O75pWw5Y+R55Sf5Y+Y5YyW44CC") },
    @{ Key="ice"; No="F06"; Name=(U "5Yaw55qE5oiQ5YiG"); En="Ice Composition"; Left=346; Top=384; AnalysisTitle=(U "5Yaw55qE5oiQ5YiG5Y+Y6YeP5YiG5p6Q"); Rows=@(@("I1",(U "57qv5rC05Yaw"),"0.104"),@("I2",(U "5rW35Yaw"),"0.089"),@("I3",(U "5Lq65bel55uQ5Yaw"),"0.078")); Conclusion=(U "5Yaw5Lit55uQ5YiG5oiW5p2C6LSo5Lya5pS55Y+Y5Yaw54K55LiO6KGo6Z2i5ray6Iac54q25oCB77yM5LuO6ICM5b2x5ZON6YeR5bGeLeWGsOeVjOmdoueahOaRqeaTpuihjOS4uuOAgg==") }
)

$app = New-Object -ComObject PowerPoint.Application
$app.Visible = $msoTrue
$presentation = $null
try {
    if (Test-Path -LiteralPath $outPath) { Remove-Item -LiteralPath $outPath -Force }
    $presentation = $app.Presentations.Add($msoTrue)
    $presentation.PageSetup.SlideWidth = 960
    $presentation.PageSetup.SlideHeight = 640
    $slide = $presentation.Slides.Add(1, $ppLayoutBlank)
    $slide.Name = "factor_matrix_overview_interactive"
    $slide.SlideShowTransition.AdvanceOnClick = $msoFalse
    $slide.SlideShowTransition.AdvanceOnTime = $msoFalse

    $bg = $slide.Shapes.AddPicture($bgPath, $msoFalse, $msoTrue, 0, 0, 960, 640)
    $bg.Name = "ice_blue_abstract_texture_background"
    $bg.ZOrder(1) | Out-Null

    [void](Add-Text $slide "slide_index" "09" 44 38 86 58 40 "#003B73" 1)
    [void](Add-LineShape $slide "index_rule" 46 94 116 94 "#FF7A1A" 3 0)
    [void](Add-Text $slide "main_title" (U "5Zug57Sg55+p6Zi1") 142 46 240 38 28 "#003B73" 1)
    [void](Add-Text $slide "main_subtitle" (U "5o6n5Yi25Y+Y6YeP5rOV5LiL55qE5pGp5pOm57O75pWw5b2x5ZON5Zug57Sg5oC76KeI") 144 88 520 22 12 "#293E52" 1)
    [void](Add-Text $slide "main_hint" (U "54K55Ye75Lu75LiA5Zug57Sg5p+l55yL5a6e6aqM5rC05bmz44CB6LaL5Yq/5Zu+5LiO54mp55CG57uT6K66") 44 128 480 18 9 "#597284" 0)

    $cards = @{}
    $highlights = @{}
    foreach ($factor in $factors) {
        $cards[$factor.Key] = Add-FactorCard $slide $factor
        $highlights[$factor.Key] = Add-Highlight $slide $factor
    }

    $overview = Add-OverviewPanel $slide
    $details = @{}
    foreach ($factor in $factors) { $details[$factor.Key] = Add-DetailPanel $slide $factor }

    [void](Add-LineShape $slide "footer_rule" 42 592 918 592 "#8DB2C7" 0.75 0.38)
    [void](Add-Text $slide "footer_icon" "*" 44 605 22 20 20 "#0077A3" 1)
    [void](Add-Text $slide "footer_title" (U "6YeR5bGe5LiO5Yaw5pGp5pOm57O75pWw5rWL6YeP") 72 608 260 18 10 "#003B73" 1)
    [void](Add-Text $slide "footer_page" "09/10" 842 604 72 20 14 "#003B73" 1)

    foreach ($factor in $factors) {
        $seq = $slide.TimeLine.InteractiveSequences.Add()
        $trigger = $cards[$factor.Key]
        [void](Add-TriggeredEffect $seq $overview $trigger ([Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectFade) $true 0.10)
        foreach ($other in $factors) {
            [void](Add-WithPreviousEffect $seq $details[$other.Key] ([Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectFade) $true 0.08)
            [void](Add-WithPreviousEffect $seq $highlights[$other.Key] ([Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectFade) $true 0.06)
        }
        [void](Add-WithPreviousEffect $seq $highlights[$factor.Key] ([Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectFade) $false 0.08)
        [void](Add-WipeWithPrevious $seq $details[$factor.Key] 0.28)

        $returnSeq = $slide.TimeLine.InteractiveSequences.Add()
        $returnTrigger = $details[$factor.Key].GroupItems.Item("return_button_" + $factor.Key)
        [void](Add-TriggeredEffect $returnSeq $details[$factor.Key] $returnTrigger ([Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectFade) $true 0.12)
        foreach ($other in $factors) {
            [void](Add-WithPreviousEffect $returnSeq $highlights[$other.Key] ([Microsoft.Office.Interop.PowerPoint.MsoAnimEffect]::msoAnimEffectFade) $true 0.06)
        }
        [void](Add-WipeWithPrevious $returnSeq $overview 0.22)
    }

    Set-PreviewState $slide "overview" $factors
    $presentation.SaveAs($outPath, $ppSaveAsOpenXMLPresentation)

    $initialPreview = Join-Path $previewDir "overview_initial_preview.png"
    $materialPreview = Join-Path $previewDir "overview_material_selected_preview.png"
    Set-PreviewState $slide "overview" $factors
    $slide.Export($initialPreview, "PNG", 1440, 960)
    Set-PreviewState $slide "material" $factors
    $slide.Export($materialPreview, "PNG", 1440, 960)
    Set-PreviewState $slide "overview" $factors
    $presentation.SaveAs($outPath, $ppSaveAsOpenXMLPresentation)

    Copy-Item -LiteralPath $PSCommandPath -Destination (Join-Path $outDir "build_factor_matrix_overview_interactive.ps1") -Force
    Write-Output "Generated $outPath"
    Write-Output "Preview $initialPreview"
    Write-Output "Preview $materialPreview"
}
finally {
    if ($presentation -ne $null) { try { $presentation.Close() } catch {} }
    try { $app.Quit() } catch {}
}
