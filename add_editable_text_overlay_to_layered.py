from pathlib import Path
import win32com.client

ROOT = Path(r"F:\Users\user\Documents\测金属摩擦系数")
PPT = ROOT / "outputs" / "metal_ice_high_fidelity_layered.pptx"
OUT = ROOT / "outputs" / "metal_ice_high_fidelity_layered_editable_text.pptx"

SLIDE_W_PT = 16 * 72
SLIDE_H_PT = 9 * 72


def add_text(slide, x, y, w, h, text, size=20, color=0x004A57, bold=True):
    # x/y/w/h are in inches; PowerPoint COM uses points.
    box = slide.Shapes.AddTextbox(1, x * 72, y * 72, w * 72, h * 72)
    box.TextFrame.TextRange.Text = text
    box.TextFrame.TextRange.Font.Name = "微软雅黑"
    box.TextFrame.TextRange.Font.Size = size
    box.TextFrame.TextRange.Font.Bold = -1 if bold else 0
    # COM color is BGR integer.
    r = (color >> 16) & 255
    g = (color >> 8) & 255
    b = color & 255
    box.TextFrame.TextRange.Font.Color.RGB = r + (g << 8) + (b << 16)
    box.TextFrame.MarginLeft = 0
    box.TextFrame.MarginRight = 0
    box.TextFrame.MarginTop = 0
    box.TextFrame.MarginBottom = 0
    box.Fill.Visible = 0
    box.Line.Visible = 0
    return box


def main():
    app = win32com.client.Dispatch("PowerPoint.Application")
    app.Visible = True
    pres = app.Presentations.Open(str(PPT), False, False, False)
    pres.SaveAs(str(OUT))
    pres.Close()
    pres = app.Presentations.Open(str(OUT), False, False, False)

    titles = [
        "金属与冰\n摩擦系数的测量",
        "选题背景",
        "题目解读",
        "方法选择 —— 动态斜面法",
        "装置设计",
        "变量设计",
        "问题改进",
        "数据处理",
        "结果分析",
        "总结展望",
    ]
    # Add editable titles and page numbers. They are intentionally placed near
    # the rendered title locations; users can edit them later. Font color matches
    # the visual design.
    for i in range(1, 11):
        slide = pres.Slides(i)
        if i == 1:
            add_text(slide, 0.53, 1.38, 3.5, 1.25, titles[i-1], 31, 0x004A57, True)
            add_text(slide, 0.35, 0.24, 0.85, 0.42, "01", 34, 0xFFFFFF, True)
        else:
            add_text(slide, 0.78, 0.20, 5.0, 0.42, titles[i-1], 22, 0x004A57, True)
            add_text(slide, 0.10, 0.10, 0.50, 0.25, f"{i:02d}", 18, 0xFFFFFF, True)
        add_text(slide, 15.36, 8.54, 0.35, 0.18, f"{i:02d}", 13, 0xFFFFFF, True)

    # Key editable numeric/formula overlays on pages where users will likely edit.
    add_text(pres.Slides(2), 9.48, 2.08, 1.7, 0.35, "25%-55%", 22, 0xE87E2D, True)
    add_text(pres.Slides(2), 9.48, 3.43, 1.7, 0.35, "13%-15%", 22, 0xE87E2D, True)
    add_text(pres.Slides(4), 9.45, 2.12, 2.7, 0.32, "a = g(sinθ - μcosθ)", 18, 0x263238, False)
    add_text(pres.Slides(4), 9.48, 3.78, 2.7, 0.36, "μ = (sinθ - a/g) / cosθ", 18, 0x263238, False)
    add_text(pres.Slides(5), 1.20, 1.20, 2.3, 0.35, "斜面冰面模块", 20, 0x004A57, True)
    add_text(pres.Slides(5), 8.28, 1.20, 3.6, 0.35, "垂直冰面模块（压力加载）", 20, 0x004A57, True)
    add_text(pres.Slides(10), 2.25, 1.02, 1.4, 0.3, "创新点", 16, 0x004A57, True)
    add_text(pres.Slides(10), 9.24, 1.02, 2.0, 0.3, "工程应用前景", 16, 0x004A57, True)

    pres.Save()
    pres.Close()
    app.Quit()
    print(OUT)


if __name__ == "__main__":
    main()
