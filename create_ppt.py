"""
当着你的面，用 COM 自动化操控 PowerPoint 生成年终总结 PPT
"""
import win32com.client
from win32com.client import constants
import os
import sys
import time

# 解决 Windows 控制台编码问题
sys.stdout.reconfigure(encoding='utf-8')

print("[启动] 正在启动 PowerPoint...")
ppt = win32com.client.Dispatch("PowerPoint.Application")
ppt.Visible = True  # 让 PowerPoint 窗口可见，你看着它操作
ppt.WindowState = 2  # 最大化窗口

# 新建演示文稿
presentation = ppt.Presentations.Add()

# ===== 辅助函数 =====
def add_slide():
    """添加一张空白幻灯片"""
    return presentation.Slides.Add(presentation.Slides.Count + 1, 12)  # 12 = ppLayoutBlank

def add_textbox(slide, left, top, width, height, text, font_size=18, font_name="微软雅黑",
                bold=False, color_rgb=None, alignment=None):
    """在幻灯片上添加文本框"""
    shape = slide.Shapes.AddTextbox(1, left, top, width, height)  # 1 = msoTextOrientationHorizontal
    shape.TextFrame.WordWrap = -1  # True
    tf = shape.TextFrame.TextRange
    tf.Text = text
    tf.Font.Name = font_name
    tf.Font.Size = font_size
    tf.Font.Bold = bold
    if color_rgb:
        tf.Font.Color.RGB = color_rgb
    if alignment is not None:
        tf.ParagraphFormat.Alignment = alignment
    return shape

def add_title(slide, text):
    """标准标题"""
    return add_textbox(slide, 50, 30, 860, 60, text, font_size=36, bold=True,
                       color_rgb=0x1A237E, alignment=2)  # 2 = ppAlignCenter

def add_subtitle(slide, text, top=100):
    """副标题"""
    return add_textbox(slide, 50, top, 860, 40, text, font_size=20,
                       color_rgb=0x546E7A, alignment=2)

def add_body(slide, lines, top=120, left=60, width=850, font_size=16, line_spacing=28):
    """添加正文（多行）"""
    height = len(lines) * line_spacing + 30
    shape = slide.Shapes.AddTextbox(1, left, top, width, height)
    shape.TextFrame.WordWrap = -1
    tf = shape.TextFrame.TextRange
    tf.Text = "\r\n".join(lines)
    tf.Font.Name = "微软雅黑"
    tf.Font.Size = font_size
    tf.Font.Color.RGB = 0x333333
    # 设置行距
    tf.ParagraphFormat.SpaceAfter = 6
    tf.ParagraphFormat.SpaceBefore = 2
    return shape

def add_divider(slide, top):
    """添加分割线"""
    line = slide.Shapes.AddLine(80, top, 880, top)
    line.Line.ForeColor.RGB = 0xE0E0E0
    line.Line.Weight = 1.5
    return line

def add_decorated_box(slide, left, top, width, height, title, value, accent_color):
    """添加卡片式数据框"""
    shape = slide.Shapes.AddShape(1, left, top, width, height)  # 1 = msoShapeRectangle
    shape.Fill.ForeColor.RGB = 0xFFFFFF
    shape.Line.ForeColor.RGB = accent_color
    shape.Line.Weight = 2

    # 数值
    add_textbox(slide, left + 10, top + 15, width - 20, 50, value,
                font_size=32, bold=True, color_rgb=accent_color, alignment=1)
    # 标签
    add_textbox(slide, left + 10, top + 70, width - 20, 30, title,
                font_size=14, color_rgb=0x757575, alignment=1)

print("[创建] 正在创建幻灯片...")

# ============================================================
# 第1页：封面
# ============================================================
print("  [1/8] 封面页...")
slide1 = add_slide()
# 背景色块
bg_shape = slide1.Shapes.AddShape(1, 0, 0, 960, 540)  # msoShapeRectangle
bg_shape.Fill.ForeColor.RGB = 0x1A237E
bg_shape.Line.Visible = 0

add_textbox(slide1, 100, 140, 760, 80, "2025年度工作总结", font_size=48, bold=True,
            color_rgb=0xFFFFFF, alignment=1)
add_textbox(slide1, 100, 240, 760, 50, "ANNUAL WORK SUMMARY", font_size=22,
            color_rgb=0xB0BEC5, alignment=1)
add_divider_line = slide1.Shapes.AddLine(250, 310, 710, 310)
add_divider_line.Line.ForeColor.RGB = 0xFFD54F
add_divider_line.Line.Weight = 3
add_textbox(slide1, 100, 340, 760, 40, "汇报人：张三  |  技术研发部  |  2026年1月", font_size=16,
            color_rgb=0xCFD8DC, alignment=1)

time.sleep(0.8)

# ============================================================
# 第2页：目录
# ============================================================
print("  [2/8] 目录页...")
slide2 = add_slide()
add_title(slide2, "目  录")
add_divider(slide2, 90)

toc_items = [
    ("01", "年度工作概览", "核心数据与关键里程碑"),
    ("02", "重点项目回顾", "主要项目成果与进展"),
    ("03", "工作亮点展示", "创新点与突出贡献"),
    ("04", "不足与反思", "问题分析与改进方向"),
    ("05", "2026年展望", "新年度目标与规划"),
]
for i, (num, title, desc) in enumerate(toc_items):
    y = 140 + i * 70
    # 编号
    add_textbox(slide2, 100, y, 60, 50, num, font_size=32, bold=True,
                color_rgb=0x1A237E, alignment=2)
    # 竖线
    vline = slide2.Shapes.AddLine(175, y + 5, 175, y + 45)
    vline.Line.ForeColor.RGB = 0xFFD54F
    vline.Line.Weight = 2
    # 标题
    add_textbox(slide2, 200, y, 300, 30, title, font_size=22, bold=True,
                color_rgb=0x212121)
    # 描述
    add_textbox(slide2, 200, y + 30, 400, 25, desc, font_size=14,
                color_rgb=0x9E9E9E)

time.sleep(0.8)

# ============================================================
# 第3页：年度工作概览
# ============================================================
print("  [3/8] 年度概览...")
slide3 = add_slide()
add_title(slide3, "年度工作概览")
add_divider(slide3, 90)

# 数据卡片
cards = [
    (60, 130, 240, 110, "完成项目数", "12 个", 0x1A237E),
    (360, 130, 240, 110, "代码提交", "1,847 次", 0x2E7D32),
    (660, 130, 240, 110, "系统上线率", "99.9%", 0xE65100),
]
for left, top, w, h, label, value, color in cards:
    add_decorated_box(slide3, left, top, w, h, label, value, color)

# 正文概述
overview_text = [
    "●  2025年是技术研发部快速发展的一年，团队规模从8人扩展至15人",
    "●  全年主导完成3个核心系统的架构升级，系统性能平均提升40%",
    "●  推动DevOps流程落地，发布效率从每周1次提升至每日3次",
    "●  技术分享与培训共组织24场，团队整体技术水平显著提升",
]
add_body(slide3, overview_text, top=280, font_size=15)

time.sleep(0.8)

# ============================================================
# 第4页：重点项目回顾
# ============================================================
print("  [4/8] 重点项目...")
slide4 = add_slide()
add_title(slide4, "重点项目回顾")
add_divider(slide4, 90)

projects = [
    ("电商平台重构", "2025.03 - 2025.06", "从单体架构迁移至微服务架构，支撑日活用户从10万增长至50万，QPS提升5倍"),
    ("数据中台建设", "2025.05 - 2025.09", "统一数据口径，建设实时数仓，日均数据处理量达10TB，报表产出效率提升80%"),
    ("AI智能客服", "2025.08 - 2025.12", "基于大模型自建智能客服，问题解决率78%，人工客服成本降低35%"),
]

for i, (name, period, desc) in enumerate(projects):
    y = 140 + i * 120
    # 时间线圆点
    circle = slide4.Shapes.AddShape(9, 70, y + 5, 14, 14)  # msoShapeOval
    circle.Fill.ForeColor.RGB = 0x1A237E
    circle.Line.Visible = 0
    # 竖线连接（除了最后一个）
    if i < len(projects) - 1:
        cline = slide4.Shapes.AddLine(77, y + 19, 77, y + 120)
        cline.Line.ForeColor.RGB = 0xE0E0E0
        cline.Line.Weight = 2

    add_textbox(slide4, 110, y, 250, 28, name, font_size=20, bold=True, color_rgb=0x212121)
    add_textbox(slide4, 370, y + 3, 200, 24, period, font_size=14, color_rgb=0x757575)
    add_textbox(slide4, 110, y + 35, 750, 50, desc, font_size=14, color_rgb=0x555555)

time.sleep(0.8)

# ============================================================
# 第5页：数据展示
# ============================================================
print("  [5/8] 数据展示...")
slide5 = add_slide()
add_title(slide5, "关键数据指标")
add_divider(slide5, 90)

# 用矩形做简易柱状图
metrics = [
    ("Q1", 80, "营收增长5%"),
    ("Q2", 140, "营收增长12%"),
    ("Q3", 200, "营收增长18%"),
    ("Q4", 280, "营收增长30%"),
]

bar_left = 100
bar_base = 420
bar_width = 120
max_bar_height = 300

for label, bar_h, note in metrics:
    # 柱状图
    bar_shape = slide5.Shapes.AddShape(1, bar_left, bar_base - bar_h, bar_width, bar_h)
    bar_shape.Fill.ForeColor.RGB = 0x3F51B5
    bar_shape.Line.Visible = 0
    # 标签
    add_textbox(slide5, bar_left, bar_base + 10, bar_width, 30, label,
                font_size=16, bold=True, color_rgb=0x212121, alignment=1)
    # 数值
    add_textbox(slide5, bar_left, bar_base - bar_h - 35, bar_width, 30, note,
                font_size=14, color_rgb=0x3F51B5, alignment=1)
    bar_left += 170

# 数据说明
data_lines = [
    "全年营收同比增长 30%，超额完成年初设定的 20% 目标",
    "用户留存率从 62% 提升至 78%，用户满意度评分 4.8/5.0",
    "系统可用性维持 99.9%，全年仅出现 2 次计划外停机",
]
add_body(slide5, data_lines, top=460, font_size=14, left=50)

time.sleep(0.8)

# ============================================================
# 第6页：工作亮点
# ============================================================
print("  [6/8] 工作亮点...")
slide6 = add_slide()
add_title(slide6, "工作亮点与创新")
add_divider(slide6, 90)

highlights = [
    ("[奖]", "技术创新奖", "自主研发的分布式任务调度框架获得公司年度技术创新奖，已在3个事业部推广使用"),
    ("[源]", "开源贡献", "向Apache/CNCF社区贡献代码累计5000+行，主导的开源项目获GitHub 2K+ Star"),
    ("[人]", "人才培养", "带教3名新人全部通过试用期考核，其中1人获评季度优秀新人"),
    ("[效]", "效率提升", "搭建统一CI/CD平台，将全团队发布流程从2小时缩短至15分钟"),
]

for i, (icon, title, desc) in enumerate(highlights):
    y = 130 + i * 95
    add_textbox(slide6, 70, y, 50, 40, icon, font_size=28, alignment=1)
    add_textbox(slide6, 140, y, 200, 28, title, font_size=18, bold=True, color_rgb=0x1A237E)
    add_textbox(slide6, 140, y + 32, 700, 40, desc, font_size=14, color_rgb=0x555555)

time.sleep(0.8)

# ============================================================
# 第7页：不足与反思
# ============================================================
print("  [7/8] 不足与反思...")
slide7 = add_slide()
add_title(slide7, "不足与反思")
add_divider(slide7, 90)

reflections = [
    ("[!]", "技术债务积累", "部分老旧模块仍使用过时技术栈，重构优先级需要更科学地评估排期"),
    ("[!]", "跨部门协作", "需求传递链路较长，信息衰减导致返工率偏高，需建立更高效的沟通机制"),
    ("[!]", "文档沉淀不足", "部分项目缺少关键的设计文档和运维手册，知识传承存在断层风险"),
    ("[!]", "个人成长瓶颈", "在管理能力上仍需提升，尤其是任务拆解和风险预判能力"),
]

for i, (icon, title, desc) in enumerate(reflections):
    y = 130 + i * 95
    add_textbox(slide7, 70, y, 50, 40, icon, font_size=28, alignment=1)
    add_textbox(slide7, 140, y, 200, 28, title, font_size=18, bold=True, color_rgb=0xE65100)
    add_textbox(slide7, 140, y + 32, 700, 40, desc, font_size=14, color_rgb=0x555555)

time.sleep(0.8)

# ============================================================
# 第8页：2026展望 + 结尾
# ============================================================
print("  [8/8] 2026展望...")
slide8 = add_slide()

# 上半部分：展望
add_title(slide8, "2026年展望")
add_divider(slide8, 90)

goals = [
    ">>  Q1目标：完成遗留系统全面重构，技术栈统一升级",
    ">>  Q2目标：启动国际化项目，支撑海外业务拓展",
    ">>  Q3目标：AI能力深度整合，产品智能化水平达到行业领先",
    ">>  Q4目标：团队规模扩展至25人，建立完善的技术梯队",
]
add_body(slide8, goals, top=120, font_size=16, line_spacing=36)

# 下半部分：感谢语
thank_bg = slide8.Shapes.AddShape(1, 0, 340, 960, 200)
thank_bg.Fill.ForeColor.RGB = 0x1A237E
thank_bg.Line.Visible = 0

add_textbox(slide8, 100, 370, 760, 60, "感谢聆听", font_size=44, bold=True,
            color_rgb=0xFFFFFF, alignment=1)
add_textbox(slide8, 100, 440, 760, 40, "THANK YOU", font_size=22,
            color_rgb=0xB0BEC5, alignment=1)

time.sleep(0.8)

# ===== 保存 =====
save_path = os.path.join(os.path.expanduser("~"), "Desktop", "2025年度工作总结.pptx")
presentation.SaveAs(save_path)
print(f"\n[完成] PPT 已保存到桌面：2025年度工作总结.pptx")
print(f"       共 {presentation.Slides.Count} 页幻灯片")
print("       PowerPoint 保持打开，你可以直接编辑修改！")

# 不关闭 PowerPoint，让用户可以看到最终结果
# ppt.Quit()  # 不退出，让你看看
