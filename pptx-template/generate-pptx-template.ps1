$ErrorActionPreference = "Stop"

# ============================================================
# PPTX 生成模板 — WPS 演示 COM 自动化
# ============================================================
# 使用方法：
#   1. 修改 $col 颜色配置（换肤）
#   2. 修改 $slides 数据数组（换内容）
#   3. 修改 $outFile 输出路径
#   4. 运行脚本
# ============================================================
# 已固化的经验（避免重复踩坑）：
#   - WPS COM 需要先杀旧进程再启动新实例，等 15 秒
#   - 用 GetActiveObject 连接，不用 New-Object（会报 80010001）
#   - 幻灯片布局用 1=Title（WPS 不支持 12=Blank）
#   - 每页之间加 300ms 延迟，避免 COM 调用被拒绝
#   - 函数名避免 PowerShell 别名（R/FT/NT/WL 等）
#   - 脚本必须用 UTF-8 BOM 编码保存（中文支持）
#   - PageSetup 用 try-catch 包裹（WPS 可能不支持）
#   - 字体大小必须 [float] 强制转换
# ============================================================

# --- 主题颜色（修改此处换肤） ---
$col = @{
    slideBg    = "FAFDFB"   # 幻灯片背景（近白微绿）
    coverBg    = "F0F7F2"   # 封面背景（浅绿）
    sectionBg  = "E8F5E9"   # 章节页背景（淡绿）
    cardBg     = "FFFFFF"   # 卡片背景（白色）
    title      = "1B5E20"   # 标题文字（深绿，白底可读）
    body       = "37474F"   # 正文文字（深灰，最大可读性）
    subtitle   = "388E3C"   # 副标题（中绿）
    muted      = "81C784"   # 次要文字（浅绿）
    white      = "FFFFFF"
    accent     = "66BB6A"   # 强调色（中浅绿）
    accentLight= "A5D6A7"   # 浅强调色
    border     = "C8E6C9"   # 边框色（极浅绿）
    decoCircle = "E8F5E9"   # 装饰圆形色
}

# --- 全局设置 ---
$fontName = "微软雅黑"
$slideW = 960
$slideH = 540
$totalSlides = 0  # 自动计算
$footerText = "培训PPT"
$outFile = "d:\trae\lark-training-ppt-generator\output.pptx"
$wppPath = "C:\Users\gogoj\AppData\Local\kingsoft\WPS Office\12.1.0.28043\office6\wpp.exe"

# ============================================================
# 辅助函数（通用，无需修改）
# ============================================================

function GetRGB($h) {
    return [Convert]::ToInt32($h.Substring(0,2),16) + [Convert]::ToInt32($h.Substring(2,2),16)*256 + [Convert]::ToInt32($h.Substring(4,2),16)*65536
}

function AddText($s, $text, $l, $t, $w, $h, $sz, $c, $b, $a) {
    $tb = $s.Shapes.AddTextbox(1, $l, $t, $w, $h)
    $tr = $tb.TextFrame.TextRange
    $tr.Text = $text
    $tr.Font.Name = $fontName
    $tr.Font.Size = [float]$sz
    $tr.Font.Color.RGB = GetRGB $c
    if ($b) { $tr.Font.Bold = $true }
    if ($a) { $tr.ParagraphFormat.Alignment = $a }
    $tr.ParagraphFormat.SpaceAfter = [float]4
    $tb.TextFrame.WordWrap = $true
    return $tb
}

function AddCard($s, $l, $t, $w, $h, $fillC, $borderC) {
    $card = $s.Shapes.AddShape(5, $l, $t, $w, $h)
    $card.Fill.ForeColor.RGB = GetRGB $fillC
    $card.Line.ForeColor.RGB = GetRGB $borderC
    $card.Line.Weight = [float]1.5
    return $card
}

function AddBar($s, $l, $t, $w, $h, $c) {
    $bar = $s.Shapes.AddShape(1, $l, $t, $w, $h)
    $bar.Fill.ForeColor.RGB = GetRGB $c
    $bar.Line.Visible = $false
    return $bar
}

function AddCircle($s, $l, $t, $d, $c) {
    $o = $s.Shapes.AddShape(9, $l, $t, $d, $d)
    $o.Fill.ForeColor.RGB = GetRGB $c
    $o.Line.Visible = $false
    return $o
}

function SetFooter($s, $n) {
    AddText $s $footerText 40 508 420 20 9 $col.muted $false $null
    AddText $s "$n / $totalSlides" 860 508 60 20 9 $col.muted $false 3
}

function SetNotes($s, $t) {
    $s.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = $t
}

# ============================================================
# 幻灯片渲染器（4 种布局：cover/section/text/2col）
# ============================================================

function RenderSlide($pres, $d, $num) {
    $s = $pres.Slides.Add($num, 1)  # 1=ppLayoutTitle
    foreach ($sh in $s.Shapes) {
        if ($sh.HasTextFrame -eq -1) {
            try { $sh.TextFrame.TextRange.Text = "" } catch {}
        }
    }

    if ($d.type -eq "cover") {
        $s.Background.Fill.ForeColor.RGB = GetRGB $col.coverBg
        AddBar $s 0 0 $slideW 6 $col.accent
        AddCircle $s 700 -100 300 $col.decoCircle
        AddCircle $s 760 360 220 $col.decoCircle
        AddText $s $d.title 60 160 840 100 40 $col.title $true $null
        AddText $s $d.sub 60 290 840 160 15 $col.subtitle $false $null

    } elseif ($d.type -eq "section") {
        $s.Background.Fill.ForeColor.RGB = GetRGB $col.sectionBg
        AddCircle $s 740 -60 260 $col.decoCircle
        AddCircle $s -40 380 200 $col.decoCircle
        AddBar $s 62 245 200 4 $col.accent
        AddText $s $d.title 60 255 840 80 36 $col.title $true $null
        if ($d.sub) {
            AddText $s $d.sub 60 345 840 40 18 $col.subtitle $false $null
        }

    } elseif ($d.type -eq "text") {
        $s.Background.Fill.ForeColor.RGB = GetRGB $col.slideBg
        AddText $s $d.title 60 38 840 50 28 $col.title $true $null
        AddBar $s 62 90 120 3 $col.accent
        AddCard $s 40 108 880 382 $col.cardBg $col.border
        AddText $s $d.content 65 122 830 355 14 $col.body $false $null

    } elseif ($d.type -eq "2col") {
        $s.Background.Fill.ForeColor.RGB = GetRGB $col.slideBg
        AddText $s $d.title 60 38 840 50 28 $col.title $true $null
        AddBar $s 62 90 120 3 $col.accent

        AddCard $s 40 108 430 382 $col.cardBg $col.border
        $leftParts = $d.left -split "`n", 2
        AddText $s $leftParts[0] 65 122 380 30 16 $col.title $true $null
        if ($leftParts.Count -gt 1) {
            AddText $s $leftParts[1] 65 158 380 310 13 $col.body $false $null
        }

        AddCard $s 490 108 430 382 $col.cardBg $col.border
        $rightParts = $d.right -split "`n", 2
        AddText $s $rightParts[0] 515 122 380 30 16 $col.title $true $null
        if ($rightParts.Count -gt 1) {
            AddText $s $rightParts[1] 515 158 380 310 13 $col.body $false $null
        }
    }

    SetFooter $s $num
    SetNotes $s $d.notes
    Start-Sleep -Milliseconds 300
    return $s
}

# ============================================================
# 幻灯片数据（修改此处换内容）
# ============================================================
# 数据格式：
#   cover:   @{type="cover";   title="标题"; sub="副标题"; notes="讲者备注"}
#   section: @{type="section"; title="标题"; sub="副标题"; notes="讲者备注"}
#   text:    @{type="text";    title="标题"; content="内容"; notes="讲者备注"}
#   2col:    @{type="2col";    title="标题"; left="左栏"; right="右栏"; notes="讲者备注"}
# 换行用 `n（反引号+n）
# ============================================================
$slides = @(
    @{type="cover"; title="培训标题"; sub="副标题`n日期 | 对象"; notes="开场白"},
    @{type="text"; title="目录"; content="01 第一部分`n02 第二部分`n03 第三部分"; notes="目录说明"},
    @{type="section"; title="第一部分"; sub="副标题"; notes="过渡说明"},
    @{type="text"; title="内容页"; content="内容文字"; notes="讲解说明"},
    @{type="2col"; title="对比页"; left="左侧标题`n左侧内容"; right="右侧标题`n右侧内容"; notes="对比说明"},
    @{type="cover"; title="谢谢"; sub="感谢参与"; notes="结束语"}
)

# ============================================================
# 主执行（无需修改）
# ============================================================

$totalSlides = $slides.Count
Write-Host "正在启动 WPS 演示..."
Get-Process -Name wpp -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 3
Start-Process $wppPath
Write-Host "  等待加载..."
Start-Sleep -Seconds 15

$ppt = [System.Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application")
Write-Host "  已连接，版本: $($ppt.Version)"
$pres = $ppt.Presentations.Add()
Start-Sleep -Seconds 1
try {
    $pres.PageSetup.SlideWidth = $slideW
    $pres.PageSetup.SlideHeight = $slideH
} catch {
    Write-Host "  PageSetup 跳过"
}

Write-Host "正在生成 $($slides.Count) 页幻灯片..."
for ($i = 0; $i -lt $slides.Count; $i++) {
    RenderSlide $pres $slides[$i] ($i + 1) | Out-Null
    Write-Host "  页 $($i + 1) / $($slides.Count) 完成"
}

Write-Host "正在保存..."
$pres.SaveAs($outFile)
Write-Host "已保存至: $outFile"
Write-Host "完成！"
