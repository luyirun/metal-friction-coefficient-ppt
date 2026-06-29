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

$outRoot = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String("Rjpc5qGM6Z2iXEtleW5vdGXnvKnnlaXlm75c5Zug57Sg55+p6Zi1XGVkaXRhYmxlX3N0YXRpY19vdmVydmlldw=="))
$bgPath = Join-Path $outRoot "ice_blue_clean_background.png"
$outPath = Join-Path $outRoot "factor_matrix_static_editable_overview.pptx"
$previewPath = Join-Path $outRoot "factor_matrix_static_editable_overview_preview.png"
New-Item -ItemType Directory -Force -Path $outRoot | Out-Null

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

function Set-NoFill($shape) {
    $shape.Fill.Visible = $msoFalse
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

function Add-Panel($slide, [string]$name, [double]$left, [double]$top, [double]$width, [double]$height, [double]$radius = 5, [string]$lineHex = "#5C95C9", [double]$lineWeight = 0.85) {
    $shape = $slide.Shapes.AddShape($msoShapeRoundedRectangle, $left, $top, $width, $height)
    $shape.Name = $name
    Set-GlassFill $shape "#FFFFFF" 0.24
    Set-Line $shape $lineHex $lineWeight 0.02
    $shape.Shadow.Visible = $msoTrue
    $shape.Shadow.ForeColor.RGB = Convert-HexColor "#6F9FBD"
    $shape.Shadow.Transparency = 0.78
    $shape.Shadow.Blur = 8
    $shape.Shadow.OffsetX = 1.0
    $shape.Shadow.OffsetY = 1.5
    return $shape
}

function Add-Button($slide, [string]$name, [double]$left, [double]$top, [double]$width, [double]$height) {
    $button = $slide.Shapes.AddShape($msoShapeRoundedRectangle, $left, $top, $width, $height)
    $button.Name = $name
    Set-GlassFill $button "#FFFFFF" 0.35
    Set-Line $button "#005A99" 0.9 0
    $button.TextFrame2.TextRange.Text = (U "54K55Ye75p+l55yL")
    $button.TextFrame2.MarginLeft = 6
    $button.TextFrame2.MarginRight = 6
    $button.TextFrame2.MarginTop = 2
    $button.TextFrame2.MarginBottom = 2
    $button.TextFrame2.TextRange.Font.NameFarEast = "Microsoft YaHei"
    $button.TextFrame2.TextRange.Font.Name = "Aptos"
    $button.TextFrame2.TextRange.Font.Size = 8.4
    $button.TextFrame2.TextRange.Font.Bold = $msoTrue
    $button.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Convert-HexColor "#00447A"
    [void](Add-Text $slide ($name + "_arrow") ">" ($left + $width - 17) ($top + 4) 10 12 10 "#005A99" 1)
    return $button
}

function Add-CircleIcon($slide, [string]$name, [double]$left, [double]$top, [string]$label) {
    $circle = $slide.Shapes.AddShape($msoShapeOval, $left, $top, 42, 42)
    $circle.Name = $name
    Set-NoFill $circle
    Set-Line $circle "#2A78C4" 1.0 0.05
    [void](Add-Text $slide ($name + "_label") $label ($left + 8) ($top + 8) 26 22 17 "#005A99" 1)
    return $circle
}

function Add-FactorIcon($slide, [string]$key, [double]$left, [double]$top) {
    if ($key -eq "temperature") {
        [void](Add-LineShape $slide "icon_temp_stem" ($left + 16) ($top + 4) ($left + 16) ($top + 36) "#005A99" 1.8 0)
        $dot = $slide.Shapes.AddShape($msoShapeOval, $left + 8, $top + 31, 16, 16)
        $dot.Name = "icon_temp_dot"; Set-NoFill $dot; Set-Line $dot "#005A99" 1.8 0
        [void](Add-Text $slide "icon_temp_snow" "*" ($left + 28) ($top + 14) 18 18 16 "#005A99" 1)
    } elseif ($key -eq "pressure") {
        $dial = $slide.Shapes.AddShape($msoShapeOval, $left + 5, $top + 5, 42, 42)
        $dial.Name = "icon_pressure_dial"; Set-NoFill $dial; Set-Line $dial "#005A99" 1.6 0
        [void](Add-LineShape $slide "icon_pressure_needle" ($left + 26) ($top + 27) ($left + 39) ($top + 16) "#005A99" 1.7 0)
        [void](Add-LineShape $slide "icon_pressure_base" ($left + 16) ($top + 50) ($left + 38) ($top + 50) "#005A99" 1.4 0)
    } elseif ($key -eq "roughness") {
        for ($i=0; $i -lt 7; $i++) {
            $x1 = $left + 3 + ($i * 6)
            $x2 = $x1 + 4
            $y1 = $top + 38 - (($i % 2) * 24)
            $y2 = $top + 14 + (($i % 2) * 22)
            [void](Add-LineShape $slide ("icon_rough_" + $i) $x1 $y1 $x2 $y2 "#005A99" 1.2 0)
        }
        [void](Add-LineShape $slide "icon_rough_base" ($left + 0) ($top + 50) ($left + 44) ($top + 50) "#005A99" 1.2 0)
    } elseif ($key -eq "area") {
        $box = $slide.Shapes.AddShape($msoShapeRoundedRectangle, $left + 4, $top + 6, 42, 42)
        $box.Name = "icon_area_box"; Set-NoFill $box; Set-Line $box "#005A99" 1.4 0
        for ($r=0; $r -lt 4; $r++) {
            for ($c=0; $c -lt 4; $c++) {
                $d = $slide.Shapes.AddShape($msoShapeOval, $left + 13 + ($c * 8), $top + 15 + ($r * 8), 2.5, 2.5)
                $d.Name = "icon_area_dot_$r`_$c"; Set-NoFill $d; Set-Line $d "#005A99" 1.0 0
            }
        }
    } else {
        [void](Add-LineShape $slide ("icon_cube_" + $key + "_1") ($left + 9) ($top + 38) ($left + 26) ($top + 48) "#005A99" 1.3 0)
        [void](Add-LineShape $slide ("icon_cube_" + $key + "_2") ($left + 26) ($top + 48) ($left + 44) ($top + 38) "#005A99" 1.3 0)
        [void](Add-LineShape $slide ("icon_cube_" + $key + "_3") ($left + 9) ($top + 38) ($left + 26) ($top + 27) "#005A99" 1.3 0)
        [void](Add-LineShape $slide ("icon_cube_" + $key + "_4") ($left + 26) ($top + 27) ($left + 44) ($top + 38) "#005A99" 1.3 0)
        [void](Add-LineShape $slide ("icon_cube_" + $key + "_5") ($left + 26) ($top + 27) ($left + 26) ($top + 48) "#005A99" 1.3 0)
    }
}

function Add-FactorCard($slide, [hashtable]$factor) {
    $l = $factor.Left; $t = $factor.Top
    [void](Add-Panel $slide ("card_" + $factor.Key) $l $t 142 136 5 "#477EB8" 0.85)
    [void](Add-Text $slide ("card_no_" + $factor.Key) $factor.No ($l + 14) ($t + 16) 36 16 8.7 "#00447A" 1)
    [void](Add-Text $slide ("card_title_" + $factor.Key) $factor.Name ($l + 14) ($t + 44) 90 24 15.4 "#003B73" 1)
    [void](Add-Text $slide ("card_en_" + $factor.Key) $factor.En ($l + 14) ($t + 73) 90 14 7.8 "#315C7A" 0)
    Add-FactorIcon $slide $factor.Key ($l + 92) ($t + 45)
    [void](Add-Button $slide ("btn_" + $factor.Key) ($l + 60) ($t + 102) 68 23)
}

function Add-Radar($slide, [double]$cx, [double]$cy, [double]$r) {
    $points = @()
    for ($i=0; $i -lt 6; $i++) {
        $ang = (-90 + $i * 60) * [Math]::PI / 180
        $x = $cx + [Math]::Cos($ang) * $r
        $y = $cy + [Math]::Sin($ang) * $r
        $points += ,@($x,$y)
        [void](Add-LineShape $slide ("radar_axis_" + $i) $cx $cy $x $y "#A8C7DB" 0.7 0.25)
    }
    for ($i=0; $i -lt 6; $i++) {
        $j = ($i + 1) % 6
        [void](Add-LineShape $slide ("radar_outer_" + $i) $points[$i][0] $points[$i][1] $points[$j][0] $points[$j][1] "#A8C7DB" 0.7 0.20)
        $ix = $cx + (($points[$i][0] - $cx) * 0.62)
        $iy = $cy + (($points[$i][1] - $cy) * 0.62)
        $jx = $cx + (($points[$j][0] - $cx) * 0.62)
        $jy = $cy + (($points[$j][1] - $cy) * 0.62)
        [void](Add-LineShape $slide ("radar_value_" + $i) $ix $iy $jx $jy "#0077A3" 1.5 0)
        $dot = $slide.Shapes.AddShape($msoShapeOval, $ix - 3, $iy - 3, 6, 6)
        $dot.Name = "radar_dot_$i"; Set-GlassFill $dot "#EAF6FF" 0.0; Set-Line $dot "#0077A3" 1.3 0
    }
    [void](Add-Text $slide "radar_label_t" ((U "5rip5bqm") + "`nF01") ($cx - 16) ($cy - $r - 32) 50 28 8 "#003B73" 1)
    [void](Add-Text $slide "radar_label_p" ((U "5Y6L5by6") + "`nF02") ($cx + $r + 8) ($cy - 13) 48 28 8 "#003B73" 1)
    [void](Add-Text $slide "radar_label_m" ((U "6YeR5bGe5p2Q6LSo") + "`nF03") ($cx + 48) ($cy + 48) 64 28 8 "#003B73" 1)
    [void](Add-Text $slide "radar_label_r" ((U "6KGo6Z2i57KX57OZ5bqm") + "`nF04") ($cx - 25) ($cy + $r + 6) 72 28 8 "#003B73" 1)
    [void](Add-Text $slide "radar_label_a" ((U "5o6l6Kem6Z2i56ev") + "`nF05") ($cx - $r - 62) ($cy + 45) 62 28 8 "#003B73" 1)
    [void](Add-Text $slide "radar_label_i" ((U "5Yaw55qE5oiQ5YiG") + "`nF06") ($cx - $r - 64) ($cy - 12) 62 28 8 "#003B73" 1)
}

$factors = @(
    @{ Key="temperature"; No="F01"; Name=(U "5rip5bqm"); En="Temperature"; Left=42; Top=218 },
    @{ Key="pressure"; No="F02"; Name=(U "5Y6L5by6"); En="Pressure"; Left=194; Top=218 },
    @{ Key="material"; No="F03"; Name=(U "6YeR5bGe5p2Q6LSo"); En="Metal Material"; Left=346; Top=218 },
    @{ Key="roughness"; No="F04"; Name=(U "6KGo6Z2i57KX57OZ5bqm"); En="Surface Roughness"; Left=42; Top=376 },
    @{ Key="area"; No="F05"; Name=(U "5o6l6Kem6Z2i56ev"); En="Contact Area"; Left=194; Top=376 },
    @{ Key="ice"; No="F06"; Name=(U "5Yaw55qE5oiQ5YiG"); En="Ice Composition"; Left=346; Top=376 }
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
    $slide.Name = "editable_static_overview"

    [void]($slide.Shapes.AddPicture($bgPath, $msoFalse, $msoTrue, 0, 0, 960, 640))

    [void](Add-Text $slide "slide_index" "09" 44 38 86 58 40 "#003B73" 1)
    [void](Add-LineShape $slide "index_rule" 46 94 116 94 "#FF7A1A" 3 0)
    [void](Add-Text $slide "main_title" (U "5Zug57Sg55+p6Zi1") 142 43 238 38 28 "#003B73" 1)
    [void](Add-Text $slide "main_subtitle" (U "5o6n5Yi25Y+Y6YeP5rOV5LiL55qE5pGp5pOm57O75pWw5b2x5ZON5Zug57Sg5oC76KeI") 144 86 520 20 12 "#003B73" 1)
    [void](Add-Text $slide "main_hint" (U "54K55Ye75Lu75LiA5Zug57Sg5p+l55yL5a6e6aqM5rC05bmz44CB6LaL5Yq/5Zu+5LiO54mp55CG57uT6K66") 44 128 480 18 9 "#476378" 0)

    foreach ($factor in $factors) { Add-FactorCard $slide $factor }

    [void](Add-Panel $slide "overview_panel" 502 118 430 448 5 "#70A8D0" 1.1)
    [void](Add-Text $slide "overview_title" (U "5Zug57Sg5oC76KeI") 524 142 146 26 18 "#003B73" 1)
    [void](Add-Text $slide "overview_en" "OVERVIEW" 526 170 90 14 8 "#315C7A" 0)
    [void](Add-Text $slide "overview_big" "OVERVIEW" 746 128 168 42 30 "#99BDD8" 1)

    [void](Add-Panel $slide "flow_box" 518 198 398 104 5 "#B2CCE0" 0.55)
    $flow = @(
        @("flow_1", 548, (U "5o6n5Yi25Y+Y6YeP"), (U "5Zu65a6a5YW25LuW5Zug57Sg"), "="),
        @("flow_2", 650, (U "5Y2V5Zug57Sg5Y+Y5YyW"), (U "5pS55Y+Y5LiA5Liq5Zug57Sg5rC05bmz"), "^"),
        @("flow_3", 752, (U "5pGp5pOm57O75pWwIM68"), (U "5rWL6YeP5ZON5bqU5Y+Y5YyW"), "u"),
        @("flow_4", 854, (U "5a+55q+U57uT6K66"), (U "5YiG5p6Q5b2x5ZON6KeE5b6L"), "[]")
    )
    foreach ($f in $flow) {
        [void](Add-CircleIcon $slide ($f[0] + "_circle") $f[1] 212 $f[4])
        [void](Add-Text $slide ($f[0] + "_title") $f[2] ($f[1] - 12) 260 66 14 8.1 "#003B73" 1)
        [void](Add-Text $slide ($f[0] + "_sub") $f[3] ($f[1] - 13) 280 72 12 7.2 "#456579" 0)
    }
    [void](Add-Text $slide "flow_arrow_1" ">" 616 226 22 18 18 "#8FB9D6" 1)
    [void](Add-Text $slide "flow_arrow_2" ">" 718 226 22 18 18 "#8FB9D6" 1)
    [void](Add-Text $slide "flow_arrow_3" ">" 820 226 22 18 18 "#8FB9D6" 1)

    [void](Add-Text $slide "group_title" (U "5Zug57Sg5YiG57uE5qaC6KeI") 524 324 110 20 11.5 "#003B73" 1)
    [void](Add-Text $slide "radar_title" (U "5YWt5aSn5Zug57Sg5YWz57O75qaC6KeI") 704 324 150 20 11.5 "#003B73" 1)
    $groupData = @(
        @("grp_env", 524, 350, (U "546v5aKD5p2h5Lu2"), (U "5rip5bqm44CB5Yaw55qE5oiQ5YiG"), "temperature"),
        @("grp_contact", 524, 412, (U "5o6l6Kem54q25oCB"), (U "5Y6L5by644CB5o6l6Kem6Z2i56ev"), "pressure"),
        @("grp_material", 524, 474, (U "5p2Q5paZ6KGo6Z2i"), (U "6YeR5bGe5p2Q6LSo44CB6KGo6Z2i57KX57OZ5bqm"), "material")
    )
    foreach ($g in $groupData) {
        [void](Add-Panel $slide ($g[0] + "_box") $g[1] $g[2] 148 54 5 "#98BEDA" 0.65)
        Add-FactorIcon $slide $g[5] ($g[1] + 12) ($g[2] + 10)
        [void](Add-Text $slide ($g[0] + "_title") $g[3] ($g[1] + 54) ($g[2] + 13) 70 14 8.7 "#003B73" 1)
        [void](Add-Text $slide ($g[0] + "_desc") $g[4] ($g[1] + 54) ($g[2] + 32) 86 12 7.0 "#315C7A" 0)
    }

    [void](Add-Panel $slide "radar_box" 686 350 230 178 5 "#B2CCE0" 0.55)
    Add-Radar $slide 800 438 58

    [void](Add-Panel $slide "conclusion_box" 518 536 398 34 5 "#B2CCE0" 0.5)
    [void](Add-Text $slide "conclusion_icon" "!" 534 545 14 16 12 "#005A99" 1)
    [void](Add-Text $slide "conclusion_text" (U "57uT6K6677ya5pGp5pOm57O75pWw55Sx5p2Q5paZ5bGe5oCn44CB5Yaw6Z2i54q25oCB5LiO5o6l6Kem5p2h5Lu25YWx5ZCM5Yaz5a6a77yM6ZyA6YCQ5o6n5Yi25Y+Y6YeP5YiG5p6Q44CC") 556 544 344 18 8.3 "#003B73" 1)

    [void](Add-LineShape $slide "footer_rule" 42 592 918 592 "#8DB2C7" 0.75 0.36)
    [void](Add-Text $slide "footer_icon" "*" 44 604 22 20 20 "#0077A3" 1)
    [void](Add-Text $slide "footer_title" (U "6YeR5bGe5LiO5Yaw5pGp5pOm57O75pWw5rWL6YeP") 72 608 260 18 10 "#003B73" 1)
    [void](Add-Text $slide "footer_page" "09/10" 842 604 72 20 14 "#003B73" 1)

    $presentation.SaveAs($outPath, $ppSaveAsOpenXMLPresentation)
    $slide.Export($previewPath, "PNG", 1536, 1024)
    Copy-Item -LiteralPath $PSCommandPath -Destination (Join-Path $outRoot "build_gorden_editable_static_overview.ps1") -Force
    Write-Output "Generated $outPath"
    Write-Output "Preview $previewPath"
}
finally {
    if ($presentation -ne $null) { try { $presentation.Close() } catch {} }
    try { $app.Quit() } catch {}
}
