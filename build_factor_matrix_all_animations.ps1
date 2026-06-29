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
$msoDirectionRight = [Microsoft.Office.Interop.PowerPoint.MsoAnimDirection]::msoAnimDirectionRight

function U([string]$base64) {
    [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($base64))
}

function Convert-HexColor([string]$hex) {
    $value = $hex.TrimStart("#")
    $r = [Convert]::ToInt32($value.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($value.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($value.Substring(4, 2), 16)
    $r + ($g * 256) + ($b * 65536)
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

function Add-Text($slide, [string]$name, [string]$text, [double]$left, [double]$top, [double]$width, [double]$height, [double]$fontSize, [string]$hex, [int]$bold = 0, [int]$align = 1) {
    $shape = $slide.Shapes.AddTextbox($msoTextOrientationHorizontal, $left, $top, $width, $height)
    $shape.Name = $name
    $shape.TextFrame2.TextRange.Text = $text
    Set-TextStyle $shape $fontSize $hex $bold
    $shape.TextFrame2.TextRange.ParagraphFormat.Alignment = $align
    $shape
}

function Add-LineShape($slide, [string]$name, [double]$x1, [double]$y1, [double]$x2, [double]$y2, [string]$hex, [double]$weight, [double]$transparency = 0) {
    $line = $slide.Shapes.AddLine($x1, $y1, $x2, $y2)
    $line.Name = $name
    $line.Line.ForeColor.RGB = Convert-HexColor $hex
    $line.Line.Weight = $weight
    $line.Line.Transparency = $transparency
    $line
}

function Add-Panel($slide, [string]$name, [double]$left, [double]$top, [double]$width, [double]$height, [string]$lineHex = "#70A8D0", [double]$lineWeight = 0.85, [double]$fillTransparency = 0.24) {
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
    $shape
}

function Add-HitArea($slide, [string]$name, [double]$left, [double]$top, [double]$width, [double]$height) {
    $shape = $slide.Shapes.AddShape($msoShapeRectangle, $left, $top, $width, $height)
    $shape.Name = $name
    $shape.Fill.Visible = $msoTrue
    $shape.Fill.ForeColor.RGB = Convert-HexColor "#FFFFFF"
    $shape.Fill.Transparency = 0.99
    $shape.Line.Visible = $msoFalse
    $shape
}

function Add-Highlight($slide, [string]$name, [object]$card) {
    $shape = $slide.Shapes.AddShape($msoShapeRoundedRectangle, $card[0], $card[1], $card[2], $card[3])
    $shape.Name = $name
    Set-NoFill $shape
    Set-Line $shape "#16A8FF" 1.55 0
    $shape.Shadow.Visible = $msoTrue
    $shape.Shadow.ForeColor.RGB = Convert-HexColor "#16A8FF"
    $shape.Shadow.Transparency = 0.55
    $shape.Shadow.Blur = 10
    $shape.Shadow.OffsetX = 0
    $shape.Shadow.OffsetY = 0
    $shape
}

function Add-Table($slide, [string]$prefix, [object]$factor, [double]$left, [double]$top, [double]$width, [double]$height) {
    $items = @()
    $items += Add-Panel $slide "$prefix`_box" $left $top $width $height "#B2CCE0" 0.55 0.68
    $rows = $factor.labels.Count + 1
    $rowH = $height / $rows
    $c1 = $width * 0.22
    $c2 = $width * 0.39
    $c3 = $width - $c1 - $c2
    $x1 = $left
    $x2 = $left + $c1
    $x3 = $left + $c1 + $c2
    $xs = @($x1, $x2, $x3)
    $ws = @($c1, $c2, $c3)
    for ($i = 1; $i -lt $rows; $i++) {
        $items += Add-LineShape $slide "$prefix`_h_$i" $left ($top + $i * $rowH) ($left + $width) ($top + $i * $rowH) "#B7D4E8" 0.55 0.1
    }
    $items += Add-LineShape $slide "$prefix`_v_1" ($left + $c1) $top ($left + $c1) ($top + $height) "#B7D4E8" 0.55 0.1
    $items += Add-LineShape $slide "$prefix`_v_2" ($left + $c1 + $c2) $top ($left + $c1 + $c2) ($top + $height) "#B7D4E8" 0.55 0.1
    for ($c = 0; $c -lt 3; $c++) {
        $items += Add-Text $slide "$prefix`_header_$c" $factor.columns[$c] ($xs[$c] + 4) ($top + 10) ($ws[$c] - 8) 12 7.6 "#003B73" 1 2
    }
    for ($r = 0; $r -lt $factor.labels.Count; $r++) {
        $texts = @($factor.labels[$r], $factor.values[$r], ("{0:N3}" -f [double]$factor.mu[$r]))
        for ($c = 0; $c -lt 3; $c++) {
            $items += Add-Text $slide "$prefix`_cell_$r`_$c" $texts[$c] ($xs[$c] + 4) ($top + (($r + 1) * $rowH) + 10) ($ws[$c] - 8) 12 7.9 "#003B73" 0 2
        }
    }
    $items
}

function Add-TrendChart($slide, [string]$prefix, [object]$factor, [double]$left, [double]$top, [double]$width, [double]$height) {
    $items = @()
    $x0 = $left + 42
    $y0 = $top + $height - 38
    $x1 = $left + $width - 16
    $y1 = $top + 28
    $maxY = 0.16
    if (($factor.mu | Measure-Object -Maximum).Maximum -le 0.12) { $maxY = 0.14 }
    $items += Add-LineShape $slide "$prefix`_axis_x" $x0 $y0 $x1 $y0 "#7EAAD0" 0.8 0
    $items += Add-LineShape $slide "$prefix`_axis_y" $x0 $y0 $x0 $y1 "#7EAAD0" 0.8 0
    for ($i = 0; $i -le 4; $i++) {
        $v = $maxY * $i / 4
        $ty = $y0 - ($y0 - $y1) * $i / 4
        $items += Add-LineShape $slide "$prefix`_tick_$i" ($x0 - 3) $ty $x0 $ty "#7EAAD0" 0.55 0
        $items += Add-Text $slide "$prefix`_tick_label_$i" ("{0:N2}" -f $v) ($left + 4) ($ty - 5) 32 10 6.8 "#003B73" 0 3
        if ($i -gt 0) {
            $items += Add-LineShape $slide "$prefix`_grid_$i" $x0 $ty $x1 $ty "#D0E2EF" 0.35 0.25
        }
    }
    $points = @()
    $count = $factor.mu.Count
    for ($i = 0; $i -lt $count; $i++) {
        $denom = [Math]::Max(1, $count - 1)
        $px = $x0 + ($x1 - $x0) * $i / $denom
        $py = $y0 - ($y0 - $y1) * ([double]$factor.mu[$i] / $maxY)
        $points += ,@($px, $py)
    }
    for ($i = 0; $i -lt ($count - 1); $i++) {
        $items += Add-LineShape $slide "$prefix`_line_$i" $points[$i][0] $points[$i][1] $points[$i + 1][0] $points[$i + 1][1] "#0077D9" 1.8 0
    }
    for ($i = 0; $i -lt $count; $i++) {
        $dot = $slide.Shapes.AddShape($msoShapeOval, $points[$i][0] - 3.5, $points[$i][1] - 3.5, 7, 7)
        $dot.Name = "$prefix`_dot_$i"
        Set-GlassFill $dot "#EAF6FF" 0
        Set-Line $dot "#0077D9" 1.5 0
        $items += $dot
        $items += Add-Text $slide "$prefix`_value_$i" ("{0:N3}" -f [double]$factor.mu[$i]) ($points[$i][0] - 16) ($points[$i][1] - 22) 34 10 6.7 "#003B73" 0 2
        $items += Add-Text $slide "$prefix`_label_$i" $factor.labels[$i] ($points[$i][0] - 12) ($y0 + 8) 24 10 7.1 "#003B73" 0 2
    }
    $items += Add-Text $slide "$prefix`_xlabel" $factor.xAxis ($x0 + 34) ($y0 + 26) 160 12 7.2 "#003B73" 0 2
    $items += Add-Text $slide "$prefix`_ylabel" (U "5ruR77yI5pGp5pOm57O75pWw77yJ") ($left + 24) ($top + 10) 82 12 7.0 "#003B73" 0 1
    $items += Add-LineShape $slide "$prefix`_legend_line" ($left + $width - 72) ($top + 10) ($left + $width - 58) ($top + 10) "#0077D9" 1.6 0
    $legendDot = $slide.Shapes.AddShape($msoShapeOval, $left + $width - 67, $top + 7.8, 4, 4)
    $legendDot.Name = "$prefix`_legend_dot"
    Set-GlassFill $legendDot "#EAF6FF" 0
    Set-Line $legendDot "#0077D9" 1 0
    $items += $legendDot
    $items += Add-Text $slide "$prefix`_legend" (U "5ruR77yI5pGp5pOm57O75pWw77yJ") ($left + $width - 55) ($top + 5) 72 12 6.7 "#003B73" 0 1
    $items
}

function Group-Items($slide, [string]$groupName, [object[]]$items) {
    $names = @()
    foreach ($item in $items) { $names += $item.Name }
    $group = $slide.Shapes.Range($names).Group()
    $group.Name = $groupName
    $group
}

function Get-OverviewNames($slide) {
    $names = @()
    for ($i = 1; $i -le $slide.Shapes.Count; $i++) {
        $n = $slide.Shapes.Item($i).Name
        if ($n -like "overview*" -or $n -like "flow*" -or $n -like "grp_*" -or $n -like "radar*" -or $n -like "conclusion*" -or $n -eq "group_title") {
            $names += $n
        }
    }
    $names
}

function Add-DetailLayer($slide, [object]$factor) {
    $items = @()
    $id = $factor.id
    $items += Add-Highlight $slide "$id`_active_card_outline" $factor.card
    $arrow = $slide.Shapes.AddShape($msoShapeRightArrow, $factor.arrow[0], $factor.arrow[1], $factor.arrow[2], $factor.arrow[3])
    $arrow.Name = "$id`_active_arrow"
    $arrow.Fill.Visible = $msoTrue
    $arrow.Fill.ForeColor.RGB = Convert-HexColor "#29A8FF"
    $arrow.Fill.Transparency = 0.08
    $arrow.Line.Visible = $msoFalse
    $items += $arrow
    $items += Add-Panel $slide "$id`_detail_panel" 502 118 430 448 "#70A8D0" 1.1 0.24
    $items += Add-Text $slide "$id`_detail_title" $factor.title 524 142 202 26 16.5 "#003B73" 1 1
    $items += Add-Panel $slide "$id`_detail_badge" 680 146 30 18 "#70A8D0" 0.6 0.62
    $items += Add-Text $slide "$id`_detail_badge_text" $factor.code 685 149 22 10 8 "#003B73" 1 2
    $items += Add-Text $slide "$id`_detail_en" $factor.en 526 171 130 14 8.5 "#003B73" 0 1
    $items += Add-Text $slide "$id`_detail_big" $factor.code 812 130 88 42 30 "#99BDD8" 1 1
    $items += Add-Text $slide "$id`_detail_close" "X" 905 134 18 18 14 "#003B73" 1 2
    $items += Add-Panel $slide "$id`_content_box" 518 198 398 332 "#B2CCE0" 0.55 0.62
    $items += Add-Text $slide "$id`_table_title" (U "5a6e6aqM5rC05bmz6KGo") 556 222 106 20 12.5 "#003B73" 1 1
    $items += Add-Text $slide "$id`_chart_title" (U "6LaL5Yq/5Zu+") 716 222 80 20 12.5 "#003B73" 1 1
    $items += Add-Text $slide "$id`_table_icon" "^" 534 222 16 18 15 "#005A99" 1 2
    $items += Add-Text $slide "$id`_chart_icon" (U "4pal") 696 222 16 18 13 "#005A99" 1 2
    $items += Add-Table $slide "$id`_table" $factor 524 258 172 166
    $items += Add-TrendChart $slide "$id`_chart" $factor 704 254 198 174
    $items += Add-Panel $slide "$id`_conclusion_box" 518 454 398 64 "#B2CCE0" 0.55 0.5
    $items += Add-Text $slide "$id`_conclusion_icon" "!" 536 466 16 16 13 "#005A99" 1 2
    $items += Add-Text $slide "$id`_conclusion_title" (U "57uT6K66") 556 465 60 18 13 "#003B73" 1 1
    $items += Add-Text $slide "$id`_conclusion_text" $factor.conclusion 536 491 356 18 8.4 "#003B73" 1 1
    $items += Add-Text $slide "$id`_note" (U "5rOo77ya5b2T5YmN5pWw5o2u5Li656S65L6L77yM5ZCO57ut5Y+v5pu/5o2i5Li655yf5a6e5a6e6aqM5pWw5o2u44CC") 526 544 250 14 7.2 "#476378" 0 1
    $items += Add-Panel $slide "$id`_back_button" 804 532 106 26 "#005A99" 0.9 0.35
    $items += Add-Text $slide "$id`_back_text" (U "PCAg6L+U5Zue55+p6Zi1") 816 540 82 10 9 "#003B73" 1 2
    Group-Items $slide "anim_$id`_detail_group" $items
}

function Add-TriggeredEffect($sequence, $shape, $triggerShape, $effectType, [bool]$exit, [double]$duration) {
    $effect = $sequence.AddEffect($shape, $effectType, $msoAnimateLevelNone, $msoTriggerOnClick)
    $effect.Timing.TriggerShape = $triggerShape
    $effect.Timing.Duration = $duration
    if ($effectType -eq $msoAnimEffectWipe) {
        try { $effect.EffectParameters.Direction = $msoDirectionRight } catch {}
    }
    if ($exit) { $effect.Exit = $msoTrue }
    $effect
}

function Add-WithPreviousEffect($sequence, $shape, $effectType, [bool]$exit, [double]$duration) {
    $effect = $sequence.AddEffect($shape, $effectType, $msoAnimateLevelNone, $msoTriggerWithPrevious)
    $effect.Timing.Duration = $duration
    if ($effectType -eq $msoAnimEffectWipe) {
        try { $effect.EffectParameters.Direction = $msoDirectionRight } catch {}
    }
    if ($exit) { $effect.Exit = $msoTrue }
    $effect
}

$dataJson = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String("ewogICJmYWN0b3JzIjogWwogICAgeyJpZCI6ImYwMSIsImNvZGUiOiJGMDEiLCJ0aXRsZSI6Iua4qeW6puWPmOmHj+WIhuaekCIsImVuIjoiVGVtcGVyYXR1cmUiLCJjYXJkIjpbNDIsMjE4LDE0MiwxMzZdLCJiYWRnZSI6WzQyLDIxOCwxNDIsMTM2XSwiYXJyb3ciOlsxNzgsMjkwLDQ0LDI2XSwiY29sdW1ucyI6WyLmsLTlubMiLCLlj5jph4/lj5blgLwiLCLmu5HvvIjmkanmk6bns7vmlbDvvIkiXSwibGFiZWxzIjpbIlQxIiwiVDIiLCJUMyIsIlQ0Il0sInZhbHVlcyI6WyItMTXCsEMiLCItMTDCsEMiLCItNcKwQyIsIjDCsEMiXSwibXUiOlswLjExOCwwLjEwMywwLjA4OSwwLjA3Nl0sInhBeGlzIjoi5rip5bqm5rC05bmz77yIwrBD77yJIiwiY29uY2x1c2lvbiI6Iua4qeW6puWNh+mrmOaXtu+8jOWGsOihqOmdoua2suiGnOWinuW8uu+8jOaRqeaTpuezu+aVsCDOvCDmlbTkvZPkuIvpmY3jgIIifSwKICAgIHsiaWQiOiJmMDIiLCJjb2RlIjoiRjAyIiwidGl0bGUiOiLljovlvLrlj5jph4/liIbmnpAiLCJlbiI6IlByZXNzdXJlIiwiY2FyZCI6WzE5NCwyMTgsMTQyLDEzNl0sImFycm93IjpbNDg2LDI5MCw0NCwyNl0sImNvbHVtbnMiOlsi5rC05bmzIiwi5Y+Y6YeP5Y+W5YC8Iiwi5ruR77yI5pGp5pOm57O75pWw77yJIl0sImxhYmVscyI6WyJQMSIsIlAyIiwiUDMiXSwidmFsdWVzIjpbIjAuMDUgTVBhIiwiMC4xMCBNUGEiLCIwLjIwIE1QYSJdLCJtdSI6WzAuMDgyLDAuMDk2LDAuMDkxXSwieEF4aXMiOiLljovlvLrmsLTlubMiLCJjb25jbHVzaW9uIjoi5Y6L5by65aKe5aSn5Lya5pS55Y+Y5bGA6YOo5o6l6Kem54q25oCB5LiO5Y6L5Yqb6J6N5YyW56iL5bqm77yM5pGp5pOm57O75pWw5Y+v6IO95ZGI546w5YWI5Y2H5ZCO6ZmN6LaL5Yq/44CCIn0sCiAgICB7ImlkIjoiZjAzIiwiY29kZSI6IkYwMyIsInRpdGxlIjoi6YeR5bGe5p2Q6LSo5Y+Y6YeP5YiG5p6QIiwiZW4iOiJNZXRhbCBNYXRlcmlhbCIsImNhcmQiOlszNDYsMjE4LDE0MiwxMzZdLCJhcnJvdyI6WzQ4NiwyOTAsNDQsMjZdLCJjb2x1bW5zIjpbIuawtOW5syIsIuWPmOmHj+WPluWAvCIsIua7ke+8iOaRqeaTpuezu+aVsO+8iSJdLCJsYWJlbHMiOlsiTTEiLCJNMiIsIk0zIiwiTTQiXSwidmFsdWVzIjpbIumTneWQiOmHkSIsIuS4jemUiOmSoiIsIumTnCIsIumSm+WQiOmHkSJdLCJtdSI6WzAuMTE4LDAuMTAzLDAuMDg5LDAuMDc2XSwieEF4aXMiOiLph5HlsZ7mnZDotKjmsLTlubMiLCJjb25jbHVzaW9uIjoi6YeR5bGe5p2Q6LSo5a+55pGp5pOm57O75pWw5b2x5ZON5pi+6JGX77ya6ZOd5ZCI6YeRID4g5LiN6ZSI6ZKiID4g6ZOcID4g6ZKb5ZCI6YeR44CCIn0sCiAgICB7ImlkIjoiZjA0IiwiY29kZSI6IkYwNCIsInRpdGxlIjoi6KGo6Z2i57KX57OZ5bqm5Y+Y6YeP5YiG5p6QIiwiZW4iOiJTdXJmYWNlIFJvdWdobmVzcyIsImNhcmQiOls0MiwzNzgsMTQyLDEzNl0sImFycm93IjpbMTc4LDQzNSw0NCwyNl0sImNvbHVtbnMiOlsi5rC05bmzIiwi6KGo6Z2i57KX57OZ5bqmIFJhIiwi5ruR77yI5pGp5pOm57O75pWw77yJIl0sImxhYmVscyI6WyJSMSIsIlIyIiwiUjMiXSwidmFsdWVzIjpbIlJhIDAuMiDOvG0iLCJSYSAxLjAgzrxtIiwiUmEgMy4wIM68bSJdLCJtdSI6WzAuMDc0LDAuMDk2LDAuMTI0XSwieEF4aXMiOiLooajpnaLnspfns5nluqbmsLTlubMiLCJjb25jbHVzaW9uIjoi6KGo6Z2i57KX57OZ5bqm5aKe5aSn5pe277yM5b6u6KeC5bWM5YWl5ZKM54qB5YmK5pWI5bqU5aKe5by677yM5pGp5pOm57O75pWwIM68IOS4iuWNh+OAgiJ9LAogICAgeyJpZCI6ImYwNSIsImNvZGUiOiJGMDUiLCJ0aXRsZSI6IuaOpeinpumdouenr+WPmOmHj+WIhuaekCIsImVuIjoiQ29udGFjdCBBcmVhIiwiY2FyZCI6WzE5NCwzNzgsMTQyLDEzNl0sImFycm93IjpbMzMwLDQzNSw0NCwyNl0sImNvbHVtbnMiOlsi5rC05bmzIiwi5Y+Y6YeP5Y+W5YC8Iiwi5ruR77yI5pGp5pOm57O75pWw77yJIl0sImxhYmVscyI6WyJBMSIsIkEyIiwiQTMiXSwidmFsdWVzIjpbIjEgY23CsiIsIjQgY23CsiIsIjkgY23CsiJdLCJtdSI6WzAuMTEyLDAuMDk2LDAuMDg3XSwieEF4aXMiOiLmjqXop6bpnaLnp6/msLTlubMiLCJjb25jbHVzaW9uIjoi5o6l6Kem6Z2i56ev5Y+Y5YyW5Lya5pS55Y+Y5Y2V5L2N5Y6L5by65ZKM5bGA6YOo5ray6Iac54q25oCB77yM5L2/5pGp5pOm57O75pWw5Y+R55Sf5Y+Y5YyW44CCIn0sCiAgICB7ImlkIjoiZjA2IiwiY29kZSI6IkYwNiIsInRpdGxlIjoi5Yaw55qE5oiQ5YiG5Y+Y6YeP5YiG5p6QIiwiZW4iOiJJY2UgQ29tcG9zaXRpb24iLCJjYXJkIjpbMzQ2LDM3OCwxNDIsMTM2XSwiYXJyb3ciOls0ODYsNDM1LDQ0LDI2XSwiY29sdW1ucyI6WyLmsLTlubMiLCLlj5jph4/lj5blgLwiLCLmu5HvvIjmkanmk6bns7vmlbDvvIkiXSwibGFiZWxzIjpbIkkxIiwiSTIiLCJJMyJdLCJ2YWx1ZXMiOlsi57qv5rC05YawIiwi5rW35YawIiwi5Lq65bel55uQ5YawIl0sIm11IjpbMC4xMDQsMC4wODksMC4wNzhdLCJ4QXhpcyI6IuWGsOeahOaIkOWIhuawtOW5syIsImNvbmNsdXNpb24iOiLlhrDkuK3nm5DliIbmiJbmnYLotKjkvJrmlLnlj5jlhrDngrnkuI7ooajpnaLmtrLohpznirbmgIHvvIzku47ogIzlvbHlk43ph5HlsZ4t5Yaw55WM6Z2i55qE5pGp5pOm6KGM5Li644CCIn0KICBdCn0=")) | ConvertFrom-Json

$sourceRoot = U "Rjpc5qGM6Z2iXEtleW5vdGXnvKnnlaXlm75c5Zug57Sg55+p6Zi1XGVkaXRhYmxlX3N0YXRpY19vdmVydmlldw=="
$basePath = Join-Path $sourceRoot "factor_matrix_static_editable_overview.pptx"
$artifactRoot = Join-Path $PSScriptRoot ".codexbridge\turn-artifacts\d95f57b4-1b7f-40f6-8722-6fb574c1651b"
$previewDir = Join-Path $artifactRoot "preview_states"
New-Item -ItemType Directory -Force -Path $artifactRoot | Out-Null
New-Item -ItemType Directory -Force -Path $previewDir | Out-Null
$outPath = Join-Path $artifactRoot "factor_matrix_static_editable_overview_all_factors_animation_v01.pptx"
$notesPath = Join-Path $artifactRoot "factor_matrix_all_factors_animation_notes.md"
if (Test-Path -LiteralPath $outPath) { Remove-Item -LiteralPath $outPath -Force }
Copy-Item -LiteralPath $basePath -Destination $outPath -Force

$app = New-Object -ComObject PowerPoint.Application
$app.Visible = $msoTrue
$presentation = $null

try {
    $presentation = $app.Presentations.Open($outPath, $msoFalse, $msoFalse, $msoTrue)
    $slide = $presentation.Slides.Item(1)
    $slide.Name = "quality_overview_all_factors_animation"
    $slide.SlideShowTransition.AdvanceOnClick = $msoFalse
    $slide.SlideShowTransition.AdvanceOnTime = $msoFalse

    $overviewNames = @(Get-OverviewNames $slide)
    $overviewGroup = $slide.Shapes.Range($overviewNames).Group()
    $overviewGroup.Name = "anim_overview_group"

    $detailGroups = @{}
    foreach ($factor in $dataJson.factors) {
        $detailGroups[$factor.id] = Add-DetailLayer $slide $factor
    }

    $factorTriggers = @{}
    foreach ($factor in $dataJson.factors) {
        $factorTriggers[$factor.id] = Add-HitArea $slide "trigger_factor_$($factor.id)" $factor.card[0] $factor.card[1] $factor.card[2] $factor.card[3]
    }

    $blockers = @{}
    $returnTriggers = @{}
    foreach ($factor in $dataJson.factors) {
        $blockers[$factor.id] = Add-HitArea $slide "block_other_factors_$($factor.id)" 25 185 465 360
        $returnTriggers[$factor.id] = Add-HitArea $slide "trigger_back_$($factor.id)" 804 532 106 26
    }

    foreach ($factor in $dataJson.factors) {
        $seqShow = $slide.TimeLine.InteractiveSequences.Add()
        [void](Add-TriggeredEffect $seqShow $overviewGroup $factorTriggers[$factor.id] $msoAnimEffectFade $true 0.18)
        [void](Add-WithPreviousEffect $seqShow $detailGroups[$factor.id] $msoAnimEffectWipe $false 0.32)
        [void](Add-WithPreviousEffect $seqShow $blockers[$factor.id] $msoAnimEffectFade $false 0.01)
        [void](Add-WithPreviousEffect $seqShow $returnTriggers[$factor.id] $msoAnimEffectFade $false 0.01)

        $seqBack = $slide.TimeLine.InteractiveSequences.Add()
        [void](Add-TriggeredEffect $seqBack $detailGroups[$factor.id] $returnTriggers[$factor.id] $msoAnimEffectFade $true 0.18)
        [void](Add-WithPreviousEffect $seqBack $blockers[$factor.id] $msoAnimEffectFade $true 0.01)
        [void](Add-WithPreviousEffect $seqBack $returnTriggers[$factor.id] $msoAnimEffectFade $true 0.01)
        [void](Add-WithPreviousEffect $seqBack $overviewGroup $msoAnimEffectWipe $false 0.32)
    }

    foreach ($factor in $dataJson.factors) {
        $detailGroups[$factor.id].Visible = $msoFalse
        $blockers[$factor.id].Visible = $msoFalse
        $returnTriggers[$factor.id].Visible = $msoFalse
    }
    $overviewGroup.Visible = $msoTrue
    $slide.Export((Join-Path $previewDir "state_00_overview_initial.png"), "PNG", 1536, 1024)

    foreach ($factor in $dataJson.factors) {
        $overviewGroup.Visible = $msoFalse
        foreach ($f in $dataJson.factors) {
            $detailGroups[$f.id].Visible = $msoFalse
            $blockers[$f.id].Visible = $msoFalse
            $returnTriggers[$f.id].Visible = $msoFalse
        }
        $detailGroups[$factor.id].Visible = $msoTrue
        $blockers[$factor.id].Visible = $msoTrue
        $returnTriggers[$factor.id].Visible = $msoTrue
        $slide.Export((Join-Path $previewDir "state_$($factor.code.ToLower())_$($factor.id)_detail.png"), "PNG", 1536, 1024)
    }

    $overviewGroup.Visible = $msoTrue
    foreach ($factor in $dataJson.factors) {
        $detailGroups[$factor.id].Visible = $msoTrue
        $blockers[$factor.id].Visible = $msoTrue
        $returnTriggers[$factor.id].Visible = $msoTrue
    }

    $presentation.Save()

    @(
        "# Factor matrix all-factors animation notes",
        "",
        "Base PPT: factor_matrix_static_editable_overview.pptx",
        "Output PPT: factor_matrix_static_editable_overview_all_factors_animation_v01.pptx",
        "",
        "Animation scope:",
        "- One slide only.",
        "- Six factor detail groups: anim_f01_detail_group ... anim_f06_detail_group.",
        "- Six factor triggers: trigger_factor_f01 ... trigger_factor_f06.",
        "- Six per-factor return triggers: trigger_back_f01 ... trigger_back_f06.",
        "- No main sequence animations; all interactions use native PowerPoint trigger sequences.",
        "",
        "Interaction:",
        "- Click one factor card: overview fades out, selected detail group wipes in from left to right.",
        "- Click return: selected detail group fades out, overview wipes in from left to right.",
        "- A transparent blocker appears while a detail group is open, preventing other factor card triggers from firing before return."
    ) | Set-Content -LiteralPath $notesPath -Encoding UTF8

    Write-Output "Generated=$outPath"
    Write-Output "PreviewDir=$previewDir"
    Write-Output "Notes=$notesPath"
}
finally {
    if ($presentation -ne $null) { try { $presentation.Close() } catch {} }
    try { $app.Quit() } catch {}
}
