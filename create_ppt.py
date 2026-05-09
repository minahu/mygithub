#!/usr/bin/env python3
"""生成提示词工程五大技巧 PPT"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import os

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# ── 颜色方案 ──
BG_DARK   = RGBColor(0x1B, 0x1B, 0x2F)   # 深蓝黑背景
BG_CARD   = RGBColor(0x27, 0x27, 0x44)   # 卡片背景
ACCENT    = RGBColor(0x00, 0xD2, 0xFF)   # 亮青色
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT     = RGBColor(0xCC, 0xCC, 0xCC)
YELLOW    = RGBColor(0xFF, 0xD7, 0x00)
GREEN     = RGBColor(0x00, 0xE6, 0x76)
ORANGE    = RGBColor(0xFF, 0x8C, 0x00)

def set_slide_bg(slide, color):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_shape_bg(slide, left, top, width, height, color, alpha=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    shape.shadow.inherit = False
    return shape

def add_text_box(slide, left, top, width, height, text, font_size=18,
                 color=WHITE, bold=False, alignment=PP_ALIGN.LEFT, font_name="Microsoft YaHei"):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.name = font_name
    p.alignment = alignment
    return txBox

def add_bullet_list(slide, left, top, width, height, items, font_size=16,
                    color=LIGHT, bullet_char="▸"):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = f"{bullet_char} {item}"
        p.font.size = Pt(font_size)
        p.font.color.rgb = color
        p.font.name = "Microsoft YaHei"
        p.space_after = Pt(6)
    return txBox

# ============================================================
# 封面
# ============================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
set_slide_bg(slide, BG_DARK)

add_text_box(slide, Inches(1), Inches(1.5), Inches(11), Inches(1.5),
             "🚀 提示词工程五大技巧", font_size=44, color=ACCENT, bold=True,
             alignment=PP_ALIGN.CENTER)

add_text_box(slide, Inches(1), Inches(3.2), Inches(11), Inches(1),
             "让 AI 更好地理解你、回答你", font_size=28, color=WHITE,
             alignment=PP_ALIGN.CENTER)

add_text_box(slide, Inches(1), Inches(4.5), Inches(11), Inches(0.8),
             "面向初学者  ·  讲解 + 练习  ·  即学即用", font_size=20, color=LIGHT,
             alignment=PP_ALIGN.CENTER)

# 底部五个关键词
keywords = ["清晰指令", "参考文本", "任务拆分", "逐步思考", "测试迭代"]
colors_kw = [ACCENT, GREEN, YELLOW, ORANGE, RGBColor(0xDA, 0x70, 0xD6)]
for i, kw in enumerate(keywords):
    x = Inches(1.5 + i * 2.2)
    add_shape_bg(slide, x, Inches(5.8), Inches(1.8), Inches(0.7), BG_CARD)
    add_text_box(slide, x, Inches(5.85), Inches(1.8), Inches(0.6),
                 kw, font_size=16, color=colors_kw[i], bold=True,
                 alignment=PP_ALIGN.CENTER)

# ============================================================
# 目录页
# ============================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, BG_DARK)

add_text_box(slide, Inches(0.8), Inches(0.5), Inches(11), Inches(1),
             "📋 今天的内容", font_size=36, color=ACCENT, bold=True)

toc_items = [
    ("1", "写清晰的指令", "让模型准确理解你想要什么"),
    ("2", "提供参考文本", "给模型可靠的知识来源"),
    ("3", "将复杂任务拆分为子任务", "降低出错率，逐步解决"),
    ("4", "给模型时间思考", "Chain-of-Thought 让推理更准确"),
    ("5", "系统化测试与迭代", "持续优化你的提示词"),
]

for i, (num, title, desc) in enumerate(toc_items):
    y = Inches(1.8 + i * 1.05)
    add_shape_bg(slide, Inches(1), y, Inches(11), Inches(0.9), BG_CARD)
    add_text_box(slide, Inches(1.3), y + Inches(0.05), Inches(0.5), Inches(0.8),
                 num, font_size=28, color=colors_kw[i], bold=True)
    add_text_box(slide, Inches(2), y + Inches(0.05), Inches(4), Inches(0.4),
                 title, font_size=22, color=WHITE, bold=True)
    add_text_box(slide, Inches(2), y + Inches(0.45), Inches(9), Inches(0.4),
                 desc, font_size=16, color=LIGHT)

# ============================================================
# 技巧模板：每个技巧 = 说明页 + 对比示例页 + 练习页
# ============================================================

techniques = [
    {
        "num": "1",
        "icon": "✏️",
        "title": "写清晰的指令",
        "color": ACCENT,
        "why": "模型无法读心，含糊的提示会得到含糊的回答。越具体、越清晰，结果越好。",
        "tips": [
            "在提示中包含细节，说明你想要的输出格式、长度、风格",
            "用角色扮演指定模型的身份（如「你是一位资深Python工程师」）",
            "用分隔符（如三引号、XML标签）明确区分输入内容",
            "指定完成任务所需的步骤",
            "提供示例（Few-shot），让模型模仿你期望的输出",
        ],
        "bad_prompt":  "帮我写个邮件。",
        "good_prompt": "你是一位专业的商务助理。请帮我写一封正式的英文邮件，\n目的：向客户道歉因发货延迟。\n语气：礼貌、诚恳。\n长度：150词以内。\n结尾需要包含后续补救方案。",
        "exercise": "请改写以下模糊提示，使其更加清晰：\n\n原始提示：「帮我总结一下这篇文章」\n\n要求：加入角色、输出格式、长度限制、语气要求。",
    },
    {
        "num": "2",
        "icon": "📚",
        "title": "提供参考文本",
        "color": GREEN,
        "why": "模型可能会「编造」看似合理但不正确的答案。提供参考文本可以让模型基于可靠信息回答，减少幻觉。",
        "tips": [
            "将参考资料直接粘贴到提示中，要求模型基于此回答",
            "使用引用标注，让模型在回答中注明引用来源",
            "适用于FAQ、文档问答、知识库查询等场景",
            "参考文本过长时，可先做摘要再输入",
        ],
        "bad_prompt":  "量子计算的原理是什么？",
        "good_prompt": "请根据以下参考文本回答问题。如果文本中没有相关信息，\n请回答「根据提供的资料无法回答」。\n\n参考文本：\n\"\"\"\n量子计算利用量子比特（qubit）的叠加态和纠缠态进行计算。\n与经典比特只能是0或1不同，量子比特可以同时处于0和1的\n叠加态，这使得量子计算机在某些问题上具有指数级加速优势。\n\"\"\"\n\n问题：量子比特与经典比特有什么区别？",
        "exercise": "你有一段产品说明书（自选），请写一个提示词：\n1. 把说明书作为参考文本嵌入\n2. 让模型只根据说明书内容回答用户问题\n3. 如果说明书里没有，要求模型说明「资料中未提及」",
    },
    {
        "num": "3",
        "icon": "🧩",
        "title": "将复杂任务拆分为子任务",
        "color": YELLOW,
        "why": "复杂任务一次性完成容易出错。把大任务拆成小步骤，每步都更可控、更准确。",
        "tips": [
            "先分类，再处理：根据问题类型选择不同的处理方式",
            "长对话中定期总结前文，避免模型遗忘上下文",
            "分步输出：先让模型生成大纲，再逐段展开",
            "用管道式处理：上一步的输出作为下一步的输入",
        ],
        "bad_prompt":  "帮我写一篇关于人工智能的3000字论文。",
        "good_prompt": "我需要写一篇关于「人工智能在医疗领域的应用」的论文。\n请分步帮我完成：\n\n第一步：先生成一个论文大纲（包含5个章节）\n第二步：（等我确认大纲后）展开第一章的详细内容\n\n现在请先完成第一步，生成大纲。",
        "exercise": "你需要让 AI 帮你制作一份旅行攻略。\n请把这个大任务拆分为 3-4 个子任务，\n并为第一个子任务写出完整的提示词。\n\n示例拆分方向：目的地调研 → 行程规划 → 预算估算 → 注意事项",
    },
    {
        "num": "4",
        "icon": "🧠",
        "title": "给模型时间思考",
        "color": ORANGE,
        "why": "就像人做数学题需要演算过程，模型在回答前先「想一想」，准确率会显著提升。这就是 Chain-of-Thought（思维链）技术。",
        "tips": [
            "基础：在提示中加一句「请先一步步分析，再给出最终答案」",
            "引导：列出具体的推理步骤让模型逐步执行",
            "结构化：用编号步骤+最终结论的格式要求输出",
            "内心独白：让模型先在隐藏区域推理，再展示最终答案",
            "追问法：先让模型回答，再问「你有没有遗漏什么？」",
        ],
        "bad_prompt":  "17 × 28 + 33 = ？",
        "good_prompt": "请计算 17 × 28 + 33 = ？\n\n要求：\n1. 先列出计算步骤\n2. 每步写出中间结果\n3. 最后给出最终答案\n\n格式：\n步骤1：...\n步骤2：...\n最终答案：...",
        "exercise": "以下是一道逻辑推理题，请用「引导式思维链」写一个提示词：\n\n题目：小明比小红高，小红比小刚高，小刚比小李矮。\n谁最高？谁最矮？\n\n要求：在提示词中明确列出推理步骤（至少3步），\n而不是只说「请一步步思考」。",
    },
    {
        "num": "5",
        "icon": "🔬",
        "title": "系统化测试与迭代",
        "color": RGBColor(0xDA, 0x70, 0xD6),
        "why": "好的提示词不是一次写好的。通过系统化的测试和对比，持续优化提示词效果。",
        "tips": [
            "准备一组测试用例（黄金标准答案），用来评估提示词效果",
            "每次只改一个变量，对比前后效果",
            "用自动化评估：让另一个模型来判断输出质量",
            "记录每次修改和结果，形成提示词优化日志",
            "关注边界情况和异常输入的处理",
        ],
        "bad_prompt":  "（凭感觉反复修改提示词，没有记录）",
        "good_prompt": "测试计划：\n\n提示词版本：v2 — 增加了输出格式要求\n\n测试用例：\n1. 输入：简单问题 → 期望：简短准确回答 ✓\n2. 输入：复杂问题 → 期望：分步解答 ✓\n3. 输入：模糊问题 → 期望：要求澄清 ✗（仍直接回答）\n\n下一步优化：增加指令「如果问题不明确，先要求用户澄清」",
        "exercise": "选择你之前练习中写过的一个提示词：\n1. 设计 3 个不同难度的测试输入\n2. 运行并记录实际输出\n3. 评估哪些地方符合预期，哪些需要改进\n4. 修改提示词，形成 v2 版本",
    },
]

for tech in techniques:
    # ── 说明页 ──
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, BG_DARK)

    # 标题栏
    add_shape_bg(slide, Inches(0), Inches(0), Inches(13.333), Inches(1.2), BG_CARD)
    add_text_box(slide, Inches(0.8), Inches(0.15), Inches(11), Inches(1),
                 f"{tech['icon']}  技巧 {tech['num']}：{tech['title']}",
                 font_size=32, color=tech['color'], bold=True)

    # 为什么重要
    add_text_box(slide, Inches(0.8), Inches(1.6), Inches(5), Inches(0.5),
                 "💡 为什么重要？", font_size=22, color=YELLOW, bold=True)
    add_text_box(slide, Inches(0.8), Inches(2.2), Inches(5.5), Inches(1.2),
                 tech['why'], font_size=17, color=LIGHT)

    # 实用要点
    add_text_box(slide, Inches(0.8), Inches(3.6), Inches(5), Inches(0.5),
                 "📌 实用要点", font_size=22, color=GREEN, bold=True)
    add_bullet_list(slide, Inches(0.8), Inches(4.2), Inches(5.5), Inches(3),
                    tech['tips'], font_size=15)

    # 右侧对比卡片
    # Bad
    add_shape_bg(slide, Inches(7), Inches(1.5), Inches(5.8), Inches(2.4),
                 RGBColor(0x3D, 0x1F, 0x1F))
    add_text_box(slide, Inches(7.3), Inches(1.6), Inches(5), Inches(0.5),
                 "❌ 不够好的提示", font_size=18, color=RGBColor(0xFF, 0x66, 0x66), bold=True)
    add_text_box(slide, Inches(7.3), Inches(2.15), Inches(5.2), Inches(1.6),
                 tech['bad_prompt'], font_size=14, color=LIGHT)

    # Good
    add_shape_bg(slide, Inches(7), Inches(4.2), Inches(5.8), Inches(2.8),
                 RGBColor(0x1F, 0x3D, 0x1F))
    add_text_box(slide, Inches(7.3), Inches(4.3), Inches(5), Inches(0.5),
                 "✅ 优化后的提示", font_size=18, color=GREEN, bold=True)
    add_text_box(slide, Inches(7.3), Inches(4.85), Inches(5.2), Inches(2),
                 tech['good_prompt'], font_size=14, color=LIGHT)

    # ── 练习页 ──
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_slide_bg(slide, BG_DARK)

    add_shape_bg(slide, Inches(0), Inches(0), Inches(13.333), Inches(1.2), BG_CARD)
    add_text_box(slide, Inches(0.8), Inches(0.15), Inches(11), Inches(1),
                 f"🎯  练习：{tech['title']}",
                 font_size=32, color=tech['color'], bold=True)

    # 练习卡片
    add_shape_bg(slide, Inches(1.5), Inches(1.8), Inches(10.3), Inches(4.8),
                 BG_CARD)
    add_text_box(slide, Inches(2), Inches(2), Inches(9), Inches(0.5),
                 "📝 动手练习", font_size=24, color=YELLOW, bold=True)
    add_text_box(slide, Inches(2), Inches(2.7), Inches(9), Inches(3.5),
                 tech['exercise'], font_size=18, color=WHITE)

    add_text_box(slide, Inches(1.5), Inches(6.8), Inches(10), Inches(0.5),
                 "⏱️ 练习时间：3-5 分钟  |  完成后与同伴交流对比",
                 font_size=16, color=LIGHT, alignment=PP_ALIGN.CENTER)

# ============================================================
# 总结页
# ============================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, BG_DARK)

add_text_box(slide, Inches(1), Inches(0.5), Inches(11), Inches(1),
             "🎓 总结：五大技巧速记", font_size=36, color=ACCENT, bold=True,
             alignment=PP_ALIGN.CENTER)

summaries = [
    ("1", "✏️ 清晰指令", "具体、明确、带示例", ACCENT),
    ("2", "📚 参考文本", "提供资料，减少幻觉", GREEN),
    ("3", "🧩 任务拆分", "大任务拆小步，逐步完成", YELLOW),
    ("4", "🧠 逐步思考", "让模型先推理再回答", ORANGE),
    ("5", "🔬 测试迭代", "对比优化，持续改进", RGBColor(0xDA, 0x70, 0xD6)),
]

for i, (num, title, desc, color) in enumerate(summaries):
    y = Inches(1.8 + i * 1.0)
    add_shape_bg(slide, Inches(2), y, Inches(9.3), Inches(0.85), BG_CARD)
    add_text_box(slide, Inches(2.3), y + Inches(0.05), Inches(0.5), Inches(0.7),
                 num, font_size=28, color=color, bold=True)
    add_text_box(slide, Inches(3), y + Inches(0.08), Inches(3), Inches(0.7),
                 title, font_size=22, color=WHITE, bold=True)
    add_text_box(slide, Inches(6.5), y + Inches(0.12), Inches(4.5), Inches(0.7),
                 desc, font_size=18, color=LIGHT)

add_text_box(slide, Inches(1), Inches(6.5), Inches(11), Inches(0.8),
             "记住：好的提示词 = 清晰的意图 + 充分的上下文 + 合理的约束",
             font_size=22, color=YELLOW, bold=True, alignment=PP_ALIGN.CENTER)

# ============================================================
# 结束页
# ============================================================
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, BG_DARK)

add_text_box(slide, Inches(1), Inches(2.5), Inches(11), Inches(1.5),
             "🙏 谢谢！", font_size=54, color=ACCENT, bold=True,
             alignment=PP_ALIGN.CENTER)

add_text_box(slide, Inches(1), Inches(4.2), Inches(11), Inches(1),
             "现在就去试试这些技巧吧！", font_size=28, color=WHITE,
             alignment=PP_ALIGN.CENTER)

add_text_box(slide, Inches(1), Inches(5.5), Inches(11), Inches(0.8),
             "有问题？随时提问 💬", font_size=20, color=LIGHT,
             alignment=PP_ALIGN.CENTER)

# ── 保存 ──
output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                           "提示词工程五大技巧.pptx")
prs.save(output_path)
print(f"✅ PPT 已生成：{output_path}")
print(f"   共 {len(prs.slides)} 页幻灯片")
