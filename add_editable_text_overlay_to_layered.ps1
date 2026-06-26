$PPT = 'F:\Users\user\Documents\测金属摩擦系数\outputs\metal_ice_high_fidelity_layered.pptx'
$OUT = 'F:\Users\user\Documents\测金属摩擦系数\outputs\metal_ice_high_fidelity_layered_editable_text.pptx'
Copy-Item -LiteralPath $PPT -Destination $OUT -Force

function Add-TextBox($slide, [double]$x, [double]$y, [double]$w, [double]$h, [string]$text, [int]$size, [int]$r, [int]$g, [int]$b, [bool]$bold) {
    $box = $slide.Shapes.AddTextbox(1, $x * 72, $y * 72, $w * 72, $h * 72)
    $box.TextFrame.TextRange.Text = $text
    $box.TextFrame.TextRange.Font.Name = '微软雅黑'
    $box.TextFrame.TextRange.Font.Size = $size
    $box.TextFrame.TextRange.Font.Bold = $(if ($bold) { -1 } else { 0 })
    $box.TextFrame.TextRange.Font.Color.RGB = $r + ($g * 256) + ($b * 65536)
    $box.TextFrame.MarginLeft = 0
    $box.TextFrame.MarginRight = 0
    $box.TextFrame.MarginTop = 0
    $box.TextFrame.MarginBottom = 0
    $box.Fill.Visible = 0
    $box.Line.Visible = 0
    return $box
}

$app = New-Object -ComObject PowerPoint.Application
$app.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$pres = $app.Presentations.Open($OUT, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)

$titles = @(
    "金属与冰`n摩擦系数的测量",
    "选题背景",
    "题目解读",
    "方法选择 —— 动态斜面法",
    "装置设计",
    "变量设计",
    "问题改进",
    "数据处理",
    "结果分析",
    "总结展望"
)

for ($i = 1; $i -le 10; $i++) {
    $slide = $pres.Slides.Item($i)
    if ($i -eq 1) {
        Add-TextBox $slide 0.53 1.38 3.6 1.3 $titles[$i-1] 31 0 74 87 $true | Out-Null
        Add-TextBox $slide 0.35 0.24 0.85 0.42 '01' 34 255 255 255 $true | Out-Null
    } else {
        Add-TextBox $slide 0.78 0.20 5.0 0.42 $titles[$i-1] 22 0 74 87 $true | Out-Null
        Add-TextBox $slide 0.10 0.10 0.50 0.25 ('{0:D2}' -f $i) 18 255 255 255 $true | Out-Null
    }
    Add-TextBox $slide 15.36 8.54 0.35 0.18 ('{0:D2}' -f $i) 13 255 255 255 $true | Out-Null
}

Add-TextBox $pres.Slides.Item(2) 9.48 2.08 1.7 0.35 '25%-55%' 22 232 126 45 $true | Out-Null
Add-TextBox $pres.Slides.Item(2) 9.48 3.43 1.7 0.35 '13%-15%' 22 232 126 45 $true | Out-Null
Add-TextBox $pres.Slides.Item(4) 9.45 2.12 2.7 0.32 'a = g(sinθ - μcosθ)' 18 38 50 56 $false | Out-Null
Add-TextBox $pres.Slides.Item(4) 9.48 3.78 2.7 0.36 'μ = (sinθ - a/g) / cosθ' 18 38 50 56 $false | Out-Null
Add-TextBox $pres.Slides.Item(5) 1.20 1.20 2.3 0.35 '斜面冰面模块' 20 0 74 87 $true | Out-Null
Add-TextBox $pres.Slides.Item(5) 8.28 1.20 3.6 0.35 '垂直冰面模块（压力加载）' 20 0 74 87 $true | Out-Null
Add-TextBox $pres.Slides.Item(10) 2.25 1.02 1.4 0.3 '创新点' 16 0 74 87 $true | Out-Null
Add-TextBox $pres.Slides.Item(10) 9.24 1.02 2.0 0.3 '工程应用前景' 16 0 74 87 $true | Out-Null

$pres.Save()
$pres.Close()
$app.Quit()
Write-Output $OUT
