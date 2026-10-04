# -*- coding: utf-8 -*-
"""
Antigravity 教師研習第一場 —— 專業大字體動畫簡報生成器（20 頁圖文並茂旗艦版）
風格：溫暖校園風 (Warm Campus) ✕ 16:9 寬螢幕 ✕ 平滑淡出動畫 (Fade) ✕ 大字體設計 (28~44pt)
包含：
  - 單元一：代理 AI 觀念啟蒙 (Slide 1 ~ 5)
  - 單元二：環境建置與安全檢測 (Slide 6 ~ 9) —— 含初學者零代碼一鍵安裝工具庫
  - 單元三：提示語大師 ✕ 基本電學圖文對決 ✕ MCP 磨課師自動化 (Slide 10 ~ 16)
    • Slide 10: 提示語指令結構對比
    • Slide 11: 【圖文實測對決 01】開天窗 vs 串並聯電路圖化簡
    • Slide 12: 【圖文實測對決 02】戴維寧等效電路與最大功率曲線
    • Slide 13: CRIS 萬用架構
    • Slide 14: 連續動作鏈 (Chain of Action)
    • Slide 15: MCP 磨課師人機接力與 1 鍵啟動
    • Slide 16: MCP 巡航與突發急救包
  - 單元四：Excel 算成績全自動化 (Slide 17 ~ 19) —— 含加權、及格標記與學情圖表
  - 收官頁：今日總結與下週精彩預告 (Slide 20)
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls

CIRCUITS_DIR = r'j:\我的雲端硬碟\antigravity\教師研習\circuits'

def create_deck(output_path):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    COLOR_BG_LIGHT       = RGBColor(250, 247, 242)   # #FAF7F2 溫暖米杏底色
    COLOR_PRIMARY_FOREST = RGBColor(31, 68, 53)      # #1F4435 典雅常春藤綠
    COLOR_CARAMEL        = RGBColor(194, 94, 46)     # #C25E2E 焦糖磚紅
    COLOR_AMBER          = RGBColor(224, 159, 62)    # #E09F3E 溫潤琥珀黃
    COLOR_SAGE           = RGBColor(142, 168, 157)   # #8EA89D 鼠尾草灰綠
    COLOR_CARD_BG        = RGBColor(255, 255, 255)   # #FFFFFF 純白卡片底
    COLOR_CARD_BORDER    = RGBColor(230, 224, 214)   # #E6E0D6 邊框灰
    COLOR_TEXT_MAIN      = RGBColor(43, 45, 66)      # #2B2D42 碳黑內文
    COLOR_TEXT_MUTED     = RGBColor(108, 117, 125)   # #6C757D 說明灰
    COLOR_PROMPT_BG      = RGBColor(254, 249, 239)   # #FEF9EF 羊皮紙提示詞底色
    COLOR_PROMPT_BORDER  = RGBColor(217, 138, 42)    # #D98A2A 琥珀金邊框
    COLOR_ALERT_ROSE     = RGBColor(188, 71, 73)     # #BC4749 玫瑰紅
    COLOR_EXCEL_GREEN    = RGBColor(16, 124, 65)     # #107C41 Excel 綠
    COLOR_ZEBRA_ROW      = RGBColor(248, 250, 249)   # 表格斑馬紋
    COLOR_FAIL_PINK      = RGBColor(254, 226, 226)   # 不及格粉紅

    FONT_FAMILY = "微軟正黑體"
    TOTAL_SLIDES = 23

    def apply_slide_transition(slide):
        trans_xml = parse_xml(f'<p:transition {nsdecls("p")} spd="med"><p:fade/></p:transition>')
        slide._element.append(trans_xml)

    def add_slide_scaffolding(slide, title_text, unit_name, slide_idx):
        apply_slide_transition(slide)

        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = COLOR_BG_LIGHT
        bg.line.fill.background()

        top_bar1 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.12))
        top_bar1.fill.solid()
        top_bar1.fill.fore_color.rgb = COLOR_PRIMARY_FOREST
        top_bar1.line.fill.background()

        top_bar2 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, Inches(0.12), Inches(4.2), Inches(0.04))
        top_bar2.fill.solid()
        top_bar2.fill.fore_color.rgb = COLOR_CARAMEL
        top_bar2.line.fill.background()

        crumb_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.35), Inches(8.5), Inches(0.35))
        tf_c = crumb_box.text_frame
        tf_c.margin_left = tf_c.margin_top = tf_c.margin_right = tf_c.margin_bottom = 0
        p_c = tf_c.paragraphs[0]
        p_c.text = f"🌿 ANTIGRAVITY 研習手冊  |  {unit_name}"
        p_c.font.size = Pt(13)
        p_c.font.bold = True
        p_c.font.color.rgb = COLOR_CARAMEL
        p_c.font.name = FONT_FAMILY

        page_box = slide.shapes.add_textbox(Inches(10.2), Inches(0.35), Inches(2.3), Inches(0.35))
        tf_p = page_box.text_frame
        tf_p.margin_left = tf_p.margin_top = tf_p.margin_right = tf_p.margin_bottom = 0
        p_p = tf_p.paragraphs[0]
        p_p.alignment = PP_ALIGN.RIGHT
        p_p.text = f"第 {slide_idx:02d} / {TOTAL_SLIDES:02d} 頁"
        p_p.font.size = Pt(13)
        p_p.font.color.rgb = COLOR_TEXT_MUTED
        p_p.font.name = FONT_FAMILY

        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.72), Inches(11.7), Inches(0.75))
        tf_t = title_box.text_frame
        tf_t.word_wrap = True
        tf_t.margin_left = tf_t.margin_top = tf_t.margin_right = tf_t.margin_bottom = 0
        p_t = tf_t.paragraphs[0]
        p_t.text = title_text
        p_t.font.size = Pt(28)
        p_t.font.bold = True
        p_t.font.color.rgb = COLOR_PRIMARY_FOREST
        p_t.font.name = FONT_FAMILY

    def add_notes(slide, notes_text):
        notes_slide = slide.notes_slide
        tf = notes_slide.notes_text_frame
        tf.text = notes_text

    def add_card(slide, left, top, width, height, bg_color=COLOR_CARD_BG, border_color=COLOR_CARD_BORDER):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = bg_color
        card.line.color.rgb = border_color
        card.line.width = Pt(1.5)
        return card

    def add_prompt_card(slide, left, top, width, height, prompt_title, prompt_content, font_size_pt=13.0):
        card = add_card(slide, left, top, width, height, bg_color=COLOR_PROMPT_BG, border_color=COLOR_PROMPT_BORDER)

        tag = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.2), top + Inches(0.2), Inches(3.8), Inches(0.42))
        tag.fill.solid()
        tag.fill.fore_color.rgb = COLOR_CARAMEL
        tag.line.fill.background()
        tf_tag = tag.text_frame
        tf_tag.margin_left = tf_tag.margin_top = tf_tag.margin_right = tf_tag.margin_bottom = 0
        p_tag = tf_tag.paragraphs[0]
        p_tag.alignment = PP_ALIGN.CENTER
        p_tag.text = f"📜 {prompt_title}"
        p_tag.font.size = Pt(13)
        p_tag.font.bold = True
        p_tag.font.color.rgb = RGBColor(255, 255, 255)
        p_tag.font.name = FONT_FAMILY

        hint_box = slide.shapes.add_textbox(left + width - Inches(2.8), top + Inches(0.22), Inches(2.6), Inches(0.38))
        tf_h = hint_box.text_frame
        tf_h.margin_left = tf_h.margin_top = tf_h.margin_right = tf_h.margin_bottom = 0
        ph = tf_h.paragraphs[0]
        ph.alignment = PP_ALIGN.RIGHT
        ph.text = "💡 可直接反白複製使用"
        ph.font.size = Pt(11)
        ph.font.color.rgb = COLOR_TEXT_MUTED
        ph.font.name = FONT_FAMILY

        tb = slide.shapes.add_textbox(left + Inches(0.3), top + Inches(0.72), width - Inches(0.6), height - Inches(0.9))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        
        lines = prompt_content.strip().split('\n')
        for i, l in enumerate(lines):
            if i == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()
            p.text = l
            p.font.size = Pt(font_size_pt)
            p.font.name = FONT_FAMILY
            p.font.color.rgb = COLOR_TEXT_MAIN
            p.line_spacing = 1.25

    # ==========================================
    # SLIDE 1: 封面頁
    # ==========================================
    s1 = prs.slides.add_slide(blank_layout)
    apply_slide_transition(s1)

    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = COLOR_PRIMARY_FOREST
    bg1.line.fill.background()

    c1 = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(9.5), Inches(-1.5), Inches(5.5), Inches(5.5))
    c1.fill.solid()
    c1.fill.fore_color.rgb = RGBColor(40, 85, 66)
    c1.line.fill.background()

    c2 = s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(-1.0), Inches(4.5), Inches(4.0), Inches(4.0))
    c2.fill.solid()
    c2.fill.fore_color.rgb = RGBColor(25, 55, 43)
    c2.line.fill.background()

    tb1 = s1.shapes.add_textbox(Inches(1.6), Inches(1.8), Inches(10.5), Inches(3.8))
    tf1 = tb1.text_frame
    tf1.word_wrap = True

    p0 = tf1.paragraphs[0]
    p0.text = "🌿 ANTIGRAVITY 代理 AI 於教學上的應用研習（第一場）"
    p0.font.size = Pt(17)
    p0.font.bold = True
    p0.font.color.rgb = COLOR_AMBER
    p0.font.name = FONT_FAMILY
    p0.space_after = Pt(14)

    p1 = tf1.add_paragraph()
    p1.text = "打通任督二脈：\n從對話到代理執行 ✕ MCP ✕ 成績自動化"
    p1.font.size = Pt(44)
    p1.font.bold = True
    p1.font.color.rgb = RGBColor(250, 247, 242)
    p1.font.name = FONT_FAMILY
    p1.space_after = Pt(18)

    p2 = tf1.add_paragraph()
    p2.text = "環境巡檢建置 ✕ 提示語大師（Prompt Master）心法 ✕ 磨課師 MCP 自動化 ✕ Excel 成績加權實戰"
    p2.font.size = Pt(20)
    p2.font.color.rgb = RGBColor(218, 226, 219)
    p2.font.name = FONT_FAMILY

    tb_footer1 = s1.shapes.add_textbox(Inches(1.6), Inches(6.0), Inches(10.5), Inches(0.8))
    tf_f1 = tb_footer1.text_frame
    pf1 = tf_f1.paragraphs[0]
    pf1.text = "🎯 研習對象：中小學與高中職教師（零基礎友善）  |  ⏱️ 研習時長：180 分鐘"
    pf1.font.size = Pt(15)
    pf1.font.color.rgb = RGBColor(180, 195, 185)
    pf1.font.name = FONT_FAMILY

    add_notes(s1, "開場說辭：各位老師好！今天這場研習我們不講生澀難懂的程式碼，重點只有一個：讓 AI 成為每天幫你備課、算加權成績、自動修完研習的實習老師，今天下課前大家都能產出實質成果！")

    # ==========================================
    # SLIDE 2: 研習大綱學習地圖（4 大核心單元）
    # ==========================================
    s2 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s2, "今日研習大綱與學習地圖", "研習導覽", 2)

    agenda_items = [
        ("單元 01", "代理 AI 觀念啟蒙", "顧問 vs. 實習老師，跨時代的能力躍升與三大超能力", COLOR_PRIMARY_FOREST),
        ("單元 02", "環境建置與安全檢測", "AI 數位辦公桌建立、沙盒機制保障、零代碼一鍵安裝工具庫", COLOR_CARAMEL),
        ("單元 03", "提示語大師 ✕ MCP 自動化", "CRIS 架構、電阻網路圖文實測對決、磨課師自動修課助手", COLOR_AMBER),
        ("單元 04", "Excel 算成績全自動化", "紙本隨手記秒轉 Excel、加權算分、及格粉紅標記、學情圖表分析", COLOR_EXCEL_GREEN)
    ]
    for idx, (code, title, desc, col) in enumerate(agenda_items):
        col_idx = idx % 2
        row = idx // 2
        x = Inches(0.8 + col_idx * 6.0)
        y = Inches(1.8 + row * 2.5)
        add_card(s2, x, y, Inches(5.7), Inches(2.25))

        bar = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(5.7), Inches(0.12))
        bar.fill.solid()
        bar.fill.fore_color.rgb = col
        bar.line.fill.background()

        tb = s2.shapes.add_textbox(x + Inches(0.35), y + Inches(0.25), Inches(5.0), Inches(1.85))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.text = code
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = col
        p.font.name = FONT_FAMILY
        
        p_t = tf.add_paragraph()
        p_t.text = title
        p_t.font.size = Pt(20)
        p_t.font.bold = True
        p_t.font.color.rgb = COLOR_TEXT_MAIN
        p_t.font.name = FONT_FAMILY
        p_t.space_before = Pt(4)
        p_t.space_after = Pt(8)

        p_d = tf.add_paragraph()
        p_d.text = desc
        p_d.font.size = Pt(15)
        p_d.font.color.rgb = COLOR_TEXT_MUTED
        p_d.font.name = FONT_FAMILY
        p_d.line_spacing = 1.3

    # ==========================================
    # SLIDE 3: 痛點共鳴
    # ==========================================
    s3 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s3, "老師日常的三大行政與教學痛點", "單元一：觀念啟蒙", 3)
    
    pain_points = [
        ("📝 改考卷與算成績", "手動拉 Excel 函數\n段考平時加權算到眼花\n不及格篩選改色耗費數小時", COLOR_ALERT_ROSE),
        ("🌐 每年必修研習折磨", "磨課師/酷課雲必修幾十小時\n每 20 分鐘彈窗防閒置中斷\n播完不跳下一單元，被迫當點擊工人", COLOR_CARAMEL),
        ("🧠 剛入門的提示語焦慮", "面對空白對話框不知道怎麼說\n隨便問一句回答空洞又無效\n不知道如何下達具體執行規格", COLOR_PRIMARY_FOREST)
    ]
    for idx, (title, desc, color) in enumerate(pain_points):
        x = Inches(0.8 + idx * 4.0)
        card = add_card(s3, x, Inches(1.8), Inches(3.7), Inches(4.5))
        
        bar = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, Inches(1.8), Inches(3.7), Inches(0.16))
        bar.fill.solid()
        bar.fill.fore_color.rgb = color
        bar.line.fill.background()

        tb = s3.shapes.add_textbox(x + Inches(0.3), Inches(2.2), Inches(3.1), Inches(3.8))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(22)
        p.font.bold = True
        p.font.color.rgb = COLOR_TEXT_MAIN
        p.font.name = FONT_FAMILY
        p.space_after = Pt(20)

        p_body = tf.add_paragraph()
        p_body.text = desc
        p_body.font.size = Pt(16)
        p_body.font.color.rgb = COLOR_TEXT_MUTED
        p_body.font.name = FONT_FAMILY
        p_body.line_spacing = 1.4

    # ==========================================
    # SLIDE 4: 觀念革命
    # ==========================================
    s4 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s4, "從「對話問答」到「自主辦事」的跨時代躍升", "單元一：觀念啟蒙", 4)

    col_data = [
        ("傳統對話 AI（ChatGPT 等）", "就像一位【顧問】", [
            "只動嘴巴回答，不動手做事",
            "產出散落對話框，需手動複製貼上",
            "單次問答，無法連續跨檔案操作",
            "需要人盯著螢幕一步一步追問"
        ], COLOR_CARD_BG, COLOR_CARD_BORDER, COLOR_TEXT_MUTED),
        ("Antigravity 代理人 (Agent)", "就像一位【數位實習老師】", [
            "具備虛擬雙手，能直接在電腦建檔運算",
            "直接產出可開啟的 .xlsx 實體檔案",
            "支援「連續動作鏈」，自動執行多步驟",
            "下達目標後自主工作，完成時向您回報！"
        ], RGBColor(245, 249, 246), COLOR_PRIMARY_FOREST, COLOR_PRIMARY_FOREST)
    ]
    for idx, (head, subhead, bullets, bg_col, border_col, title_col) in enumerate(col_data):
        x = Inches(0.8 + idx * 6.0)
        add_card(s4, x, Inches(1.8), Inches(5.7), Inches(4.8), bg_color=bg_col, border_color=border_col)

        tb = s4.shapes.add_textbox(x + Inches(0.35), Inches(2.1), Inches(5.0), Inches(4.2))
        tf = tb.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        p.text = head
        p.font.size = Pt(21)
        p.font.bold = True
        p.font.color.rgb = title_col
        p.font.name = FONT_FAMILY

        p_sub = tf.add_paragraph()
        p_sub.text = subhead
        p_sub.font.size = Pt(17)
        p_sub.font.bold = True
        p_sub.font.color.rgb = COLOR_CARAMEL
        p_sub.font.name = FONT_FAMILY
        p_sub.space_before = Pt(4)
        p_sub.space_after = Pt(14)

        for b in bullets:
            pb = tf.add_paragraph()
            pb.text = "• " + b
            pb.font.size = Pt(15)
            pb.font.color.rgb = COLOR_TEXT_MAIN
            pb.font.name = FONT_FAMILY
            pb.space_after = Pt(10)

    # ==========================================
    # SLIDE 5: 認識三大超能力
    # ==========================================
    s5 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s5, "認識 Antigravity 代理人的「三大超能力」", "單元一：觀念啟蒙", 5)

    powers = [
        ("🧠 自主規劃 (Plan)", "把大任務拆解為執行清單", "給定教學目標，自動啟動深度思考，拆解第一步、第二步、第三步，避免做白工。"),
        ("🛠️ 工具調用 (Tools)", "擁有讀寫檔案與運算的能力", "不只產出文字，能自主編寫並執行腳本，處理 Excel 表格加權運算與圖表繪製。"),
        ("📦 工作區成果 (Artifacts)", "實體檔案直接存在你的硬碟", "所有成績表與報告直接存在電腦資料夾，右側面板隨時即時預覽，一鍵複製使用。")
    ]
    for idx, (title, sub, desc) in enumerate(powers):
        x = Inches(0.8 + idx * 4.0)
        add_card(s5, x, Inches(1.8), Inches(3.7), Inches(4.8))

        tag = s5.shapes.add_shape(MSO_SHAPE.OVAL, x + Inches(0.35), Inches(2.15), Inches(0.65), Inches(0.65))
        tag.fill.solid()
        tag.fill.fore_color.rgb = COLOR_PRIMARY_FOREST
        tag.line.fill.background()
        tf_tag = tag.text_frame
        pt = tf_tag.paragraphs[0]
        pt.alignment = PP_ALIGN.CENTER
        pt.text = str(idx + 1)
        pt.font.size = Pt(16)
        pt.font.bold = True
        pt.font.color.rgb = RGBColor(255, 255, 255)
        pt.font.name = FONT_FAMILY

        tb = s5.shapes.add_textbox(x + Inches(0.3), Inches(2.95), Inches(3.1), Inches(3.4))
        tf = tb.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(19)
        p.font.bold = True
        p.font.color.rgb = COLOR_PRIMARY_FOREST
        p.font.name = FONT_FAMILY

        p_sub = tf.add_paragraph()
        p_sub.text = sub
        p_sub.font.size = Pt(15)
        p_sub.font.bold = True
        p_sub.font.color.rgb = COLOR_CARAMEL
        p_sub.font.name = FONT_FAMILY
        p_sub.space_before = Pt(4)
        p_sub.space_after = Pt(12)

        p_d = tf.add_paragraph()
        p_d.text = desc
        p_d.font.size = Pt(14)
        p_d.font.color.rgb = COLOR_TEXT_MUTED
        p_d.font.name = FONT_FAMILY
        p_d.line_spacing = 1.35

    # ==========================================
    # SLIDE 6: 打造你的 AI 數位辦公桌
    # ==========================================
    s6 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s6, "打造你的「AI 數位辦公桌」：專屬教學工作區", "單元二：環境建置", 6)

    zones = [
        ("📂 左側：專案檔案庫", "左側工作區檔案樹", "打開 Antigravity 點選 Open Workspace，選擇或新建「2026_教學專案」資料夾。所有由 AI 建立的表格都存於此。"),
        ("💬 中央：對話指揮中心", "與 AI 代理人的互動畫布", "如同你的數位助教辦公桌。在這裡下達 CRIS 提示詞，觀察 AI 的思考過程 (Thinking) 與動作執行。"),
        ("📱 右側：成果即時預覽區", "Artifacts 面板", "檔案產生後自動在此彈出預覽！試算表、長條圖與文字直接呈現，免手動開外部軟體即可快速檢查確認。")
    ]
    for idx, (title, sub, sd) in enumerate(zones):
        x = Inches(0.8 + idx * 4.0)
        add_card(s6, x, Inches(1.8), Inches(3.7), Inches(4.8))

        tb = s6.shapes.add_textbox(x + Inches(0.3), Inches(2.1), Inches(3.1), Inches(4.2))
        tf = tb.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(20)
        p.font.bold = True
        p.font.color.rgb = COLOR_PRIMARY_FOREST
        p.font.name = FONT_FAMILY
        p.space_after = Pt(10)

        p2 = tf.add_paragraph()
        p2.text = sd
        p2.font.size = Pt(15)
        p2.font.color.rgb = COLOR_TEXT_MUTED
        p2.font.name = FONT_FAMILY
        p2.line_spacing = 1.35

    # ==========================================
    # SLIDE 7: 沙盒機制
    # ==========================================
    s7 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s7, "安心守則：什麼是「沙盒機制 (Sandbox)」？", "單元二：環境建置", 7)

    add_card(s7, Inches(0.8), Inches(1.8), Inches(11.7), Inches(4.8))
    tb_s7 = s7.shapes.add_textbox(Inches(1.2), Inches(2.1), Inches(10.9), Inches(4.2))
    tf_s7 = tb_s7.text_frame
    tf_s7.word_wrap = True

    p = tf_s7.paragraphs[0]
    p.text = "🛡️ 老師最常擔心的問題：「AI 操作我的電腦，會不會把我原本的重要公文改壞？」"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = COLOR_CARAMEL
    p.font.name = FONT_FAMILY
    p.space_after = Pt(16)

    points = [
        ("專屬隔離區（Sandbox 沙盒）", "就像在操場劃設安全的跑道，Antigravity 的所有自動化執行都被限制在「目前的專案資料夾」，絕對不會跑出去修改你的系統檔或桌面私人檔案。"),
        ("透明可見的執行過程", "AI 的每一步動作、修改了哪個檔案，全部在畫布上記錄得清清楚楚，隨時可以復原。"),
        ("資安與隱私第一準則", "練習時請勿使用真實學生的身分證字號、家長私人電話等敏感個資，使用代號或模擬資料備課最安心。")
    ]
    for t, d in points:
        pt = tf_s7.add_paragraph()
        pt.text = "🌿 " + t
        pt.font.size = Pt(17)
        pt.font.bold = True
        pt.font.color.rgb = COLOR_PRIMARY_FOREST
        pt.font.name = FONT_FAMILY
        pt.space_before = Pt(8)

        pd = tf_s7.add_paragraph()
        pd.text = "    " + d
        pd.font.size = Pt(15)
        pd.font.color.rgb = COLOR_TEXT_MUTED
        pd.font.name = FONT_FAMILY
        pd.space_after = Pt(8)

    # ==========================================
    # SLIDE 8: 【重點強化】初學者環境一鍵安裝與自動巡檢（實測清冊 ✕ 健檢咒語）
    # ==========================================
    s8 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s8, "初學者必備：零代碼環境一鍵健檢與已裝自動化程式清冊", "單元二：環境建置", 8)

    # 左欄：目前電腦已實裝之自動化與繪圖程式清冊 (實測依據)
    add_card(s8, Inches(0.8), Inches(1.8), Inches(5.6), Inches(4.8))
    tb_i8 = s8.shapes.add_textbox(Inches(1.05), Inches(2.0), Inches(5.1), Inches(4.4))
    tf_i8 = tb_i8.text_frame
    tf_i8.word_wrap = True

    p = tf_i8.paragraphs[0]
    p.text = "🖥️ 本機實測已就緒之自動化工具清冊"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_FOREST
    p.font.name = FONT_FAMILY
    p.space_after = Pt(6)

    installed_groups = [
        ("⚡ 基礎環境核心", "Python 3.14 (直譯器) ✕ Node.js v24 (MCP運作) ✕ Git (版本/推播)"),
        ("📐 專業電學與科學繪圖", "schemdraw (IEEE鋸齒電阻/電源/閉迴路) ✕ matplotlib (功率曲線) ✕ Pillow"),
        ("📊 Office 行政自動化", "python-pptx (簡報排版動畫) ✕ openpyxl / pandas (加權算分/學情統計)"),
        ("🌐 瀏覽器巡航與協定", "chrome-devtools-mcp (磨課師人機接力) ✕ selenium ✕ requests")
    ]
    for grp_title, grp_items in installed_groups:
        pt = tf_i8.add_paragraph()
        pt.text = grp_title
        pt.font.size = Pt(14)
        pt.font.bold = True
        pt.font.color.rgb = COLOR_CARAMEL
        pt.font.name = FONT_FAMILY
        
        pi = tf_i8.add_paragraph()
        pi.text = grp_items
        pi.font.size = Pt(12)
        pi.font.color.rgb = COLOR_TEXT_MAIN
        pi.font.name = FONT_FAMILY
        pi.space_after = Pt(4)

    p_bot = tf_i8.add_paragraph()
    p_bot.text = "💡 免敲 CMD 指令！只要下達右側咒語，代理人即在背景自動配齊上述工具。"
    p_bot.font.size = Pt(11)
    p_bot.font.bold = True
    p_bot.font.color.rgb = COLOR_TEXT_MUTED
    p_bot.font.name = FONT_FAMILY

    # 右欄：全自動健檢與安裝咒語 (擴充包含 schemdraw, matplotlib, git)
    prompt8 = """請扮演我的「電腦教學環境裝備醫生」。我是一位剛接觸 AI 的零基礎老師，請幫我檢查並自動裝備好工作環境：

1. 檢查核心環境：Python 與 Node.js 是否已安裝就緒並回報版本？
2. 檢查 Git 工具：是否安裝 Git 版本控制（供後續推播 GitHub 網站使用）？
3. 檢查並自動安裝「教學必備套件庫」（若缺少請背景自動 pip install，免我開終端機）：
   - Office 自動化：openpyxl、pandas、python-pptx、python-docx
   - 專業電學與圖表：schemdraw（標準電路圖）、matplotlib（特性曲線）、pillow
4. 工作區讀寫測試：在當前目錄建立「健檢回報.txt」，並條列檢測狀態。
5. 完成後用親切繁體中文，以打勾 (✅) 方式回報所有裝備就緒清單！"""

    add_prompt_card(s8, Inches(6.6), Inches(1.8), Inches(5.9), Inches(4.8), "實戰提示詞 01：全自動環境健檢與裝備咒語", prompt8, 11.0)

    # ==========================================
    # SLIDE 9: 初試啼聲自主建檔
    # ==========================================
    s9 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s9, "初試啼聲：讓代理人自主在電腦建立第一份檔案", "單元二：環境建置", 9)

    add_card(s9, Inches(0.8), Inches(1.8), Inches(4.2), Inches(4.8))
    tb_i9 = s9.shapes.add_textbox(Inches(1.1), Inches(2.1), Inches(3.6), Inches(4.2))
    tf_i9 = tb_i9.text_frame
    tf_i9.word_wrap = True
    p = tf_i9.paragraphs[0]
    p.text = "🎯 實作目標"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_FOREST
    p.font.name = FONT_FAMILY
    p.space_after = Pt(14)

    tips9 = [
        "體驗代理 AI 的「造物」能力！",
        "觀察送出後：",
        "1. AI 思考中（Thinking）",
        "2. 工具調用（write_to_file）",
        "3. 左側檔案清單跳出「研習簽到與教學備忘.txt」",
        "4. 點開檔案，內容已真實存在硬碟中！"
    ]
    for t in tips9:
        p_t = tf_i9.add_paragraph()
        p_t.text = t
        p_t.font.size = Pt(15)
        p_t.font.color.rgb = COLOR_TEXT_MAIN
        p_t.font.name = FONT_FAMILY
        p_t.space_after = Pt(8)

    prompt9 = """你好！我是 [請填寫學校：例如 大安高工 / 彰師附工] 的 [請填寫學科：例如 電子科 / 資訊科] 老師。
請在我們的工作資料夾中，幫我建立一個名為「研習簽到與教學備忘.txt」的檔案，
內容寫入：
- 今日日期與時間
- 研習主題：Antigravity 代理 AI 教學應用實戰
- 我的教學願景：期許用 AI 幫自己每週省下 3 小時備課時間
- 給老師的一句備課激勵金句
建立完成後，請告訴我檔案已儲存。"""
    add_prompt_card(s9, Inches(5.3), Inches(1.8), Inches(7.2), Inches(4.8), "實戰提示詞範本 02：第一份自主建檔咒語", prompt9, 13.5)

    # ==========================================
    # SLIDE 10: 提示語大師修煉法（基本電學命題結構對比）
    # ==========================================
    s10 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s10, "為什麼別人用 AI 像神仙，我用 AI 像智障？", "單元三：提示語大師", 10)

    add_card(s10, Inches(0.8), Inches(1.8), Inches(5.6), Inches(4.8))
    tb_b10 = s10.shapes.add_textbox(Inches(1.1), Inches(2.1), Inches(5.0), Inches(4.2))
    tf_b10 = tb_b10.text_frame
    tf_b10.word_wrap = True
    p = tf_b10.paragraphs[0]
    p.text = "❌ 失敗指令：隨意說一句話"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = COLOR_ALERT_ROSE
    p.font.name = FONT_FAMILY
    p.space_after = Pt(14)

    p_cmd1 = tf_b10.add_paragraph()
    p_cmd1.text = "「幫我出幾題基本電學題目。」"
    p_cmd1.font.size = Pt(16)
    p_cmd1.font.bold = True
    p_cmd1.font.color.rgb = COLOR_TEXT_MAIN
    p_cmd1.font.name = FONT_FAMILY
    p_cmd1.space_after = Pt(14)

    cons = [
        "AI 不知道給技術高中哪一科、幾年級學生做（高一初學？統測複習？）",
        "題目數值未經設計，常出現除不盡的繁雜小數，無法快速手算",
        "缺乏電路化簡結構，沒有串並聯拆解步驟與戴維寧等效解析",
        "老師還得花大量時間拿計算紙一題一題重新驗算電路"
    ]
    for c in cons:
        p_c = tf_b10.add_paragraph()
        p_c.text = "• " + c
        p_c.font.size = Pt(14)
        p_c.font.color.rgb = COLOR_TEXT_MUTED
        p_c.font.name = FONT_FAMILY
        p_c.space_after = Pt(8)

    add_card(s10, Inches(6.9), Inches(1.8), Inches(5.6), Inches(4.8), border_color=COLOR_PRIMARY_FOREST)
    tb_g10 = s10.shapes.add_textbox(Inches(7.2), Inches(2.1), Inches(5.0), Inches(4.2))
    tf_g10 = tb_g10.text_frame
    tf_g10.word_wrap = True
    p = tf_g10.paragraphs[0]
    p.text = "⭕ 大師級指令：結構化清晰下達"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_FOREST
    p.font.name = FONT_FAMILY
    p.space_after = Pt(14)

    p_cmd2 = tf_g10.add_paragraph()
    p_cmd2.text = "「你是一位技術高中電機電子群名師，正在教高一『基本電學：電阻網路分析』。請針對『直流電阻串並聯化簡與戴維寧定理 (Thevenin)』，設計 3 題統測試題等級的素養選擇題。要求：1. 數值經精心設計（電阻與電壓皆為整數，便於計算） 2. 每題包含清晰電路結構描述、4 個選項與標準答案 3. 詳解需列出完整推導步驟（含等效電阻 Rth 與等效電壓 Vth 之求法，並標註學生常錯的短路/開路盲點），以 Markdown 表格呈現。」"
    p_cmd2.font.size = Pt(13.5)
    p_cmd2.font.bold = True
    p_cmd2.font.color.rgb = COLOR_TEXT_MAIN
    p_cmd2.font.name = FONT_FAMILY
    p_cmd2.space_after = Pt(14)

    pros = [
        "角色清晰：鎖定技高電機電子群高一認知程度",
        "數值友好：整數設計，避免繁雜小數干擾核心電路概念",
        "解析嚴謹：完整列出戴維寧等效電路 (Rth, Vth) 逐步推導",
        "產出即成果：一秒複製即可直接印成課堂隨堂評量卷！"
    ]
    for pr in pros:
        p_pr = tf_g10.add_paragraph()
        p_pr.text = "✔ " + pr
        p_pr.font.size = Pt(14)
        p_pr.font.color.rgb = COLOR_TEXT_MAIN
        p_pr.font.name = FONT_FAMILY
        p_pr.space_after = Pt(8)

    # ==========================================
    # SLIDE 11: 【圖文實測對決 01】隨意指令慘遭開天窗 ✕ 大師指令電路圖化簡
    # ==========================================
    s11 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s11, "實測對決 01：隨意指令慘遭開天窗 ✕ 大師指令電路圖化簡", "單元三：提示語大師", 11)

    # 左側卡片：失敗指令產出慘狀
    add_card(s11, Inches(0.8), Inches(1.8), Inches(5.4), Inches(5.2), border_color=COLOR_ALERT_ROSE)
    tb_fail = s11.shapes.add_textbox(Inches(1.05), Inches(2.0), Inches(4.9), Inches(4.8))
    tf_f = tb_fail.text_frame
    tf_f.word_wrap = True

    p = tf_f.paragraphs[0]
    p.text = "❌ 隨意指令產出（真實車禍現場）"
    p.font.size = Pt(19)
    p.font.bold = True
    p.font.color.rgb = COLOR_ALERT_ROSE
    p.font.name = FONT_FAMILY
    p.space_after = Pt(8)

    fail_items = [
        ("題 1：死背名詞", "「什麼是歐姆定律？請寫出公式。」\n➔ 國中生都知道，完全沒有技高電阻網路評量價值。"),
        ("題 2：數值極度不友善", "「100V 電壓接 33Ω 電阻，求電流？」\n➔ 答案 I = 3.0303...A（無限循環小數，學生算到懷疑人生）"),
        ("題 3：致命傷：電路開天窗！", "「試求下圖電路的戴維寧等效電路。（註：此處無圖）❌」\n➔ AI 憑空說下圖，根本沒有圖！題目直接作廢！"),
        ("老師下場：耗時重工", "無選項、無推導步驟，老師還得花半小時拿計算紙手動重出重算！")
    ]
    for tag, desc in fail_items:
        pt = tf_f.add_paragraph()
        pt.text = "• " + tag
        pt.font.size = Pt(14)
        pt.font.bold = True
        pt.font.color.rgb = COLOR_TEXT_MAIN
        pt.font.name = FONT_FAMILY

        pd = tf_f.add_paragraph()
        pd.text = "   " + desc
        pd.font.size = Pt(12.5)
        pd.font.color.rgb = COLOR_TEXT_MUTED
        pd.font.name = FONT_FAMILY
        pd.space_after = Pt(4)

    # 右側卡片：大師級指令圖文並茂產出
    add_card(s11, Inches(6.5), Inches(1.8), Inches(6.0), Inches(5.2), border_color=COLOR_PRIMARY_FOREST)
    tb_succ1 = s11.shapes.add_textbox(Inches(6.75), Inches(1.95), Inches(5.5), Inches(1.3))
    tf_s1 = tb_succ1.text_frame
    tf_s1.word_wrap = True
    p = tf_s1.paragraphs[0]
    p.text = "⭕ 大師指令產出：題 01 電阻串並聯化簡"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_FOREST
    p.font.name = FONT_FAMILY

    p_body = tf_s1.add_paragraph()
    p_body.text = "【題幹】E = 24V，R1 = 6Ω 與 R2 = 12Ω 並聯後，再與 R3 = 8Ω 串聯。求 RT 與 IT？\n【選項】(A) RT=12Ω, IT=2A   (B) RT=14Ω, IT=1.7A   (C) RT=10Ω, IT=2.4A"
    p_body.font.size = Pt(12)
    p_body.font.color.rgb = COLOR_TEXT_MAIN
    p_body.font.name = FONT_FAMILY

    # 嵌入真實電路圖 1
    c1_path = os.path.join(CIRCUITS_DIR, 'circuit1_series_parallel.png')
    if os.path.exists(c1_path):
        s11.shapes.add_picture(c1_path, Inches(6.75), Inches(3.3), width=Inches(5.5))

    tb_succ1_sol = s11.shapes.add_textbox(Inches(6.75), Inches(5.85), Inches(5.5), Inches(1.1))
    tf_ss1 = tb_succ1_sol.text_frame
    tf_ss1.word_wrap = True
    p_sol = tf_ss1.paragraphs[0]
    p_sol.text = "✔ 正解 (A) | 推導：R1//R2 = 4Ω ➔ RT = 4 + 8 = 12Ω ➔ IT = 24 / 12 = 2A（整數手算超順！）"
    p_sol.font.size = Pt(11.5)
    p_sol.font.bold = True
    p_sol.font.color.rgb = COLOR_PRIMARY_FOREST
    p_sol.font.name = FONT_FAMILY

    p_sol2 = tf_ss1.add_paragraph()
    p_sol2.text = "🚨 常見盲點：學生易誤先算串聯 (6+12=18) 再並聯 8，詳解完整提示避免踩雷！"
    p_sol2.font.size = Pt(11)
    p_sol2.font.color.rgb = COLOR_CARAMEL
    p_sol2.font.name = FONT_FAMILY

    # ==========================================
    # SLIDE 12: 【圖文實測對決 02】戴維寧等效模型 ✕ 最大功率轉移曲線
    # ==========================================
    s12 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s12, "實測對決 02：戴維寧等效模型 ✕ 最大功率轉移曲線", "單元三：提示語大師", 12)

    # 左側卡片：題 02 戴維寧等效模型
    add_card(s12, Inches(0.8), Inches(1.8), Inches(5.7), Inches(5.2))
    tb_t2 = s12.shapes.add_textbox(Inches(1.05), Inches(1.95), Inches(5.2), Inches(1.2))
    tf_t2 = tb_t2.text_frame
    tf_t2.word_wrap = True
    p = tf_t2.paragraphs[0]
    p.text = "⭕ 題 02：戴維寧等效模型與負載分析"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_FOREST
    p.font.name = FONT_FAMILY

    p2_b = tf_t2.add_paragraph()
    p2_b.text = "【題幹】開路電壓 Eth = 18V，等效電阻 Rth = 6Ω。外接負載 RL = 3Ω，求負載電流 IL？"
    p2_b.font.size = Pt(12)
    p2_b.font.color.rgb = COLOR_TEXT_MAIN
    p2_b.font.name = FONT_FAMILY

    c2_path = os.path.join(CIRCUITS_DIR, 'circuit2_thevenin.png')
    if os.path.exists(c2_path):
        s12.shapes.add_picture(c2_path, Inches(1.05), Inches(3.2), width=Inches(5.2))

    tb_t2_sol = s12.shapes.add_textbox(Inches(1.05), Inches(5.8), Inches(5.2), Inches(1.1))
    tf_t2s = tb_t2_sol.text_frame
    tf_t2s.word_wrap = True
    p = tf_t2s.paragraphs[0]
    p.text = "✔ 正解 (B) 2A | 單迴路串聯：IL = 18 / (6 + 3) = 2A（正負整數好算）"
    p.font.size = Pt(11.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_FOREST
    p.font.name = FONT_FAMILY
    p2 = tf_t2s.add_paragraph()
    p2.text = "🚨 防呆叮嚀：求 Rth 時獨立電壓源應「視為短路 (0V)」，學生常誤當開路！"
    p2.font.size = Pt(11)
    p2.font.color.rgb = COLOR_CARAMEL
    p2.font.name = FONT_FAMILY

    # 右側卡片：題 03 最大功率轉移曲線
    add_card(s12, Inches(6.8), Inches(1.8), Inches(5.7), Inches(5.2))
    tb_t3 = s12.shapes.add_textbox(Inches(7.05), Inches(1.95), Inches(5.2), Inches(1.2))
    tf_t3 = tb_t3.text_frame
    tf_t3.word_wrap = True
    p = tf_t3.paragraphs[0]
    p.text = "⭕ 題 03：戴維寧最大功率轉移特性"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_FOREST
    p.font.name = FONT_FAMILY

    p3_b = tf_t3.add_paragraph()
    p3_b.text = "【題幹】Eth = 20V，Rth = 5Ω。求使負載獲最大功率之 RL 阻值與最大功率 PL(max)？"
    p3_b.font.size = Pt(12)
    p3_b.font.color.rgb = COLOR_TEXT_MAIN
    p3_b.font.name = FONT_FAMILY

    c3_path = os.path.join(CIRCUITS_DIR, 'circuit3_max_power.png')
    if os.path.exists(c3_path):
        s12.shapes.add_picture(c3_path, Inches(7.05), Inches(3.2), width=Inches(5.2))

    tb_t3_sol = s12.shapes.add_textbox(Inches(7.05), Inches(5.8), Inches(5.2), Inches(1.1))
    tf_t3s = tb_t3_sol.text_frame
    tf_t3s.word_wrap = True
    p = tf_t3s.paragraphs[0]
    p.text = "✔ 正解 (A) RL=5Ω, PL=20W | 公式：PL(max) = Eth² / (4*Rth) = 400 / 20 = 20W"
    p.font.size = Pt(11.5)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_FOREST
    p.font.name = FONT_FAMILY
    p2 = tf_t3s.add_paragraph()
    p2.text = "🚨 防呆叮嚀：分母常忘乘係數 4，誤選 80W（選項 B 典型誘答陷阱）！"
    p2.font.size = Pt(11)
    p2.font.color.rgb = COLOR_CARAMEL
    p2.font.name = FONT_FAMILY

    # ==========================================
    # SLIDE 13: CRIS 提示語萬用框架
    # ==========================================
    s13 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s13, "提示語大師核心公式：CRIS 萬用架構", "單元三：提示語大師", 13)

    cris_boxes = [
        ("C - Context", "背景與角色", "設定 AI 的專業身分，以及面對的年級、學生程度與教學情境。\n例：『你是技高電機電子群名師，面對高一基本電學學生...』", COLOR_PRIMARY_FOREST),
        ("R - Request", "明確目標任務", "以動詞開頭，直接告訴 AI 最終要完成什麼工作。\n例：『請設計 3 題「電阻串並聯與戴維寧定理」分析試題...』", COLOR_CARAMEL),
        ("I - Instructions", "條件與限制規則", "限制難易度、題型、字數、避免提及的內容。\n例：『阻值與電壓需為整數、詳解含 Rth/Vth 步驟、標註常犯迷思...』", COLOR_AMBER),
        ("S - Style & Format", "風格與輸出格式", "指定輸出外觀，如 Markdown 表格、Word 檔、或 Excel 試算表。\n例：『以 Markdown 表格呈現，含題幹、選項、答案與詳解...』", COLOR_SAGE)
    ]
    for idx, (code, title, desc, col) in enumerate(cris_boxes):
        row = idx // 2
        col_i = idx % 2
        x = Inches(0.8 + col_i * 6.0)
        y = Inches(1.8 + row * 2.5)
        add_card(s13, x, y, Inches(5.7), Inches(2.25))

        tb_c = s13.shapes.add_textbox(x + Inches(0.3), y + Inches(0.2), Inches(5.1), Inches(1.85))
        tf = tb_c.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = f"{code}（{title}）"
        p.font.size = Pt(18)
        p.font.bold = True
        p.font.color.rgb = col
        p.font.name = FONT_FAMILY
        p.space_after = Pt(8)

        p2 = tf.add_paragraph()
        p2.text = desc
        p2.font.size = Pt(14)
        p2.font.color.rgb = COLOR_TEXT_MAIN
        p2.font.name = FONT_FAMILY
        p2.line_spacing = 1.3

    # ==========================================
    # SLIDE 14: 連續動作鏈 (Chain of Action)
    # ==========================================
    s14 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s14, "代理人專用：下達「連續動作鏈 (Chain of Action)」", "單元三：提示語大師", 14)

    steps_chain = [
        ("步驟 1：讀取分析", "「先讀取工作區內的實習成績檔案，檢查各組實測數據是否有缺漏...」"),
        ("步驟 2：計算處理", "「接著依據加權標準計算每位同學總分，並標記需課後補救實習名單...」"),
        ("步驟 3：產出建檔", "「最後將計算後的完整表格存為新檔，並附上一段 200 字實習學情摘要。」")
    ]
    for idx, (st, ex) in enumerate(steps_chain):
        y = Inches(1.8 + idx * 1.6)
        add_card(s14, Inches(0.8), y, Inches(11.7), Inches(1.38))
        
        c_num = s14.shapes.add_shape(MSO_SHAPE.OVAL, Inches(1.1), y + Inches(0.24), Inches(0.9), Inches(0.9))
        c_num.fill.solid()
        c_num.fill.fore_color.rgb = COLOR_CARAMEL
        c_num.line.fill.background()
        tf_n = c_num.text_frame
        pn = tf_n.paragraphs[0]
        pn.alignment = PP_ALIGN.CENTER
        pn.text = str(idx + 1)
        pn.font.size = Pt(20)
        pn.font.bold = True
        pn.font.color.rgb = RGBColor(255, 255, 255)
        pn.font.name = FONT_FAMILY

        tb = s14.shapes.add_textbox(Inches(2.3), y + Inches(0.18), Inches(9.8), Inches(0.98))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = st
        p.font.size = Pt(18)
        p.font.bold = True
        p.font.color.rgb = COLOR_PRIMARY_FOREST
        p.font.name = FONT_FAMILY
        
        p2 = tf.add_paragraph()
        p2.text = "範例指令：" + ex
        p2.font.size = Pt(15)
        p2.font.color.rgb = COLOR_TEXT_MUTED
        p2.font.name = FONT_FAMILY
        p2.space_before = Pt(4)

    # ==========================================
    # SLIDE 15: 神級代理人：MCP 瀏覽器自動化 —— 人機接力與 1 鍵啟動
    # ==========================================
    s15 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s15, "神級代理人：MCP 瀏覽器自動化 —— 人機接力與 1 鍵啟動", "單元三：提示語大師 ✕ MCP", 15)

    add_card(s15, Inches(0.8), Inches(1.8), Inches(5.6), Inches(4.8))
    tb_i15 = s15.shapes.add_textbox(Inches(1.1), Inches(2.05), Inches(5.0), Inches(4.3))
    tf_i15 = tb_i15.text_frame
    tf_i15.word_wrap = True

    p = tf_i15.paragraphs[0]
    p.text = "🤝 30 秒人機接力黃金法則"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_FOREST
    p.font.name = FONT_FAMILY
    p.space_after = Pt(6)

    p_gold = tf_i15.add_paragraph()
    p_gold.text = "💡「資安歸老師，苦工歸 AI！」"
    p_gold.font.size = Pt(16)
    p_gold.font.bold = True
    p_gold.font.color.rgb = COLOR_CARAMEL
    p_gold.font.name = FONT_FAMILY
    p_gold.space_after = Pt(12)

    steps_mcp = [
        ("👤 老師親手做（前 30 秒免踩雷）：", [
            "1. 點兩下「啟動AI伴讀Chrome.bat」（遠端除錯埠 9222）",
            "2. 親自登入教育雲 OpenID（免個資外洩、免驗證碼）",
            "3. 點開課程第一單元，按一次播放鍵！"
        ]),
        ("🤖 AI 代理人接管（後 30 分鐘全代勞）：", [
            "• 0.5 秒自動代點「下一單元」微小按鈕",
            "• 自動按掉「禁止多重視窗」等中斷警告彈窗",
            "• 每 24 分鐘微暫停 5 秒重置防閒置計時"
        ])
    ]
    for title, items in steps_mcp:
        pt = tf_i15.add_paragraph()
        pt.text = title
        pt.font.size = Pt(15)
        pt.font.bold = True
        pt.font.color.rgb = COLOR_TEXT_MAIN
        pt.font.name = FONT_FAMILY
        pt.space_before = Pt(4)

        for it in items:
            pi = tf_i15.add_paragraph()
            pi.text = "   " + it
            pi.font.size = Pt(13)
            pi.font.color.rgb = COLOR_TEXT_MUTED
            pi.font.name = FONT_FAMILY
            pi.space_after = Pt(2)

    prompt15 = """請使用 chrome-devtools-mcp 連接我目前開啟的 Chrome 瀏覽器：
1. 找到我正在播放研習影片的磨課師（或數位學習平臺）分頁。
2. 幫我監控影片進度：當前單元播完時，自動幫我點擊「下一單元」繼續播放。
3. 若畫面彈出「禁止多重視窗」或確認警示窗，請自動點擊「確定」排除。
4. 開始執行，並在每完成一個章節時向我回報！"""
    add_prompt_card(s15, Inches(6.8), Inches(1.8), Inches(5.7), Inches(4.8), "提示詞範本 03-A：生手「一鍵通」秒殺指令", prompt15, 13.0)

    # ==========================================
    # SLIDE 16: 生手一次就成功：全自動巡航 ✕ 突發狀況急救包
    # ==========================================
    s16 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s16, "生手一次就成功：全自動巡航 ✕ 突發狀況急救包", "單元三：提示語大師 ✕ MCP", 16)

    prompt16 = """【角色】請扮演我的「數位研習全自動伴讀導航員」。
【情境】我已登入磨課師研習平臺（Chrome 9222 遠端除錯埠），畫面正停在課程播放頁。
【任務與執行規則】：
1. 鎖定分頁：使用 list_pages 與 select_page 鎖定包含「moocs」或「elearning」課程分頁。
2. 連播機制：以 1.0x 正常速度播放，影片播畢自動點擊「下一單元（nav_next）」。
3. 防閒置破解：每播 24 分鐘自動微暫停 5 秒後恢復，重置 25/30 分鐘防閒置計時器。
4. 彈窗攔截：呼叫 handle_dialog，一旦有警告彈窗 0 秒點確定。
5. 持續巡航，並在右側記錄完成章節數與累積研習時數！"""
    add_prompt_card(s16, Inches(0.8), Inches(1.8), Inches(6.0), Inches(4.8), "提示詞範本 03-B：大師級全自動巡航 CRIS 咒語", prompt16, 12.0)

    add_card(s16, Inches(7.1), Inches(1.8), Inches(5.4), Inches(4.8), border_color=COLOR_CARAMEL)
    tb_i16 = s16.shapes.add_textbox(Inches(7.35), Inches(2.05), Inches(4.9), Inches(4.3))
    tf_i16 = tb_i16.text_frame
    tf_i16.word_wrap = True

    p = tf_i16.paragraphs[0]
    p.text = "🛠️ 研習現場突發狀況「急救包」"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = COLOR_CARAMEL
    p.font.name = FONT_FAMILY
    p.space_after = Pt(4)

    p_sub = tf_i16.add_paragraph()
    p_sub.text = "遇到意外狀況？複製這幾句咒語貼上立即排除："
    p_sub.font.size = Pt(13)
    p_sub.font.color.rgb = COLOR_TEXT_MUTED
    p_sub.font.name = FONT_FAMILY
    p_sub.space_after = Pt(10)

    troubles = [
        ("⚠️ 狀況 1：抓錯分頁", "「請列出 Chrome 所有開啟分頁標題，並將焦點切換到磨課師課程那一頁。」"),
        ("⚠️ 狀況 2：彈窗卡住", "「畫面上跳出警告對話框，請呼叫 handle_dialog 幫我自動按確定關閉。」"),
        ("⚠️ 狀況 3：按鈕微小", "「請在當前網頁尋找包含『下一頁』或『下一單元』按鈕並點擊它。」"),
        ("⚠️ 狀況 4：查詢進度", "「請幫我查看當前課程的累計閱讀時數與完成進度百分比。」")
    ]
    for tag, cmd in troubles:
        pt = tf_i16.add_paragraph()
        pt.text = tag
        pt.font.size = Pt(14)
        pt.font.bold = True
        pt.font.color.rgb = COLOR_PRIMARY_FOREST
        pt.font.name = FONT_FAMILY

        pc = tf_i16.add_paragraph()
        pc.text = "   " + cmd
        pc.font.size = Pt(12)
        pc.font.color.rgb = COLOR_TEXT_MAIN
        pc.font.name = FONT_FAMILY
        pc.space_after = Pt(6)

    # ==========================================
    # SLIDE 17: Office 自動化第一彈
    # ==========================================
    s17 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s17, "Office 自動化第一彈：手動拉公式算成績的終結者", "單元四：Excel 自動化", 17)

    cols17 = [
        ("傳統算成績的痛苦", "• 總分、平均手動打 =SUM、=AVERAGE\n• 不及格要一欄一欄篩選、改粉紅色\n• 等第判定寫巢狀 =IF 寫到眼花\n• 全班最高分、最低分還要再算一次", COLOR_ALERT_ROSE),
        ("Antigravity 代理人解法", "• 把原始資料給它，或直接模擬整班成績\n• AI 在背景自行執行演算法運算\n• 自動標記不及格粉紅色底色\n• 幾秒鐘直接產出格式化完成的 .xlsx", COLOR_EXCEL_GREEN)
    ]
    for idx, (title, content, col) in enumerate(cols17):
        x = Inches(0.8 + idx * 6.0)
        add_card(s17, x, Inches(1.8), Inches(5.7), Inches(4.8))
        tb = s17.shapes.add_textbox(x + Inches(0.3), Inches(2.1), Inches(5.1), Inches(4.2))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(22)
        p.font.bold = True
        p.font.color.rgb = col
        p.font.name = FONT_FAMILY
        p.space_after = Pt(18)

        lines = content.split('\n')
        for l in lines:
            pl = tf.add_paragraph()
            pl.text = l
            pl.font.size = Pt(16)
            pl.font.color.rgb = COLOR_TEXT_MAIN
            pl.font.name = FONT_FAMILY
            pl.space_after = Pt(12)

    # ==========================================
    # SLIDE 18: 實戰演練：全自動班級成績統計與等第運算
    # ==========================================
    s18 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s18, "實戰演練：全自動班級成績統計與等第運算", "單元四：Excel 自動化", 18)

    prompt18 = """請扮演我的「教務與實習成績統計助理」。
任務：在工作區建立「高二電子科專業實習第一次段考成績表.xlsx」。
資料要求：
1. 模擬 15 位技高學生名單（座號 1~15、學生姓名用王小明等假名），包含三項成績：
   - 專業筆試 (占 30%)
   - 示波器實作量測 (占 50%)
   - 平時實習表現與工安規範 (占 20%)
2. 自動計算：
   - 每位學生「學期加權總分（四捨五入至小數第一位）」。
   - 根據加權總分判定「等第」（90優/80甲/70乙/60丙/未滿60丁）。
   - 實作或筆試未滿 60 分用醒目粉紅色底色標記。
3. 表格最下方計算各項「全班平均」、「最高分」、「最低分」。
完成後儲存為 .xlsx 檔案並提供預覽。"""
    add_prompt_card(s18, Inches(0.8), Inches(1.8), Inches(5.9), Inches(4.8), "實戰提示詞範本 04：Excel 自動算成績大師", prompt18, 12.0)

    add_card(s18, Inches(7.0), Inches(1.8), Inches(5.5), Inches(4.8), bg_color=RGBColor(255, 255, 255), border_color=COLOR_EXCEL_GREEN)
    
    excel_bar = s18.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.2), Inches(2.0), Inches(5.1), Inches(0.42))
    excel_bar.fill.solid()
    excel_bar.fill.fore_color.rgb = COLOR_EXCEL_GREEN
    excel_bar.line.fill.background()
    tf_eb = excel_bar.text_frame
    tf_eb.margin_left = tf_eb.margin_top = tf_eb.margin_right = tf_eb.margin_bottom = 0
    p_eb = tf_eb.paragraphs[0]
    p_eb.alignment = PP_ALIGN.CENTER
    p_eb.text = "📊 預期產出：高二電子科專業實習第一次段考成績表.xlsx"
    p_eb.font.size = Pt(12)
    p_eb.font.bold = True
    p_eb.font.color.rgb = RGBColor(255, 255, 255)
    p_eb.font.name = FONT_FAMILY

    table_shape = s18.shapes.add_table(6, 6, Inches(7.2), Inches(2.55), Inches(5.1), Inches(3.8))
    table = table_shape.table
    table.columns[0].width = Inches(0.7)
    table.columns[1].width = Inches(0.9)
    table.columns[2].width = Inches(0.8)
    table.columns[3].width = Inches(0.8)
    table.columns[4].width = Inches(0.8)
    table.columns[5].width = Inches(1.1)

    headers = ["座號", "姓名", "筆試(30%)", "實作(50%)", "平時(20%)", "加權/等第"]
    mock_data = [
        ["01", "王小明", "88", "92", "85", "89.4 (甲)"],
        ["02", "李小華", "55", "74", "60", "65.5 (丙)"],
        ["03", "張大同", "95", "96", "95", "95.5 (優)"],
        ["04", "陳美麗", "72", "48", "75", "60.6 (丙)"],
        ["班平", "全班平均", "77.5", "77.5", "78.8", "77.8"]
    ]

    for col_idx, h in enumerate(headers):
        cell = table.cell(0, col_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = COLOR_EXCEL_GREEN
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.text = h
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.font.name = FONT_FAMILY

    for row_idx, r in enumerate(mock_data):
        for col_idx, val in enumerate(r):
            cell = table.cell(row_idx + 1, col_idx)
            cell.fill.solid()
            if row_idx == 4:
                cell.fill.fore_color.rgb = RGBColor(230, 240, 235)
            elif (row_idx == 1 and col_idx == 2) or (row_idx == 3 and col_idx == 3):
                cell.fill.fore_color.rgb = COLOR_FAIL_PINK
            elif row_idx % 2 == 1:
                cell.fill.fore_color.rgb = COLOR_ZEBRA_ROW
            else:
                cell.fill.fore_color.rgb = RGBColor(255, 255, 255)
            
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER
            p.text = val
            p.font.size = Pt(11)
            p.font.name = FONT_FAMILY
            if (row_idx == 1 and col_idx == 2) or (row_idx == 3 and col_idx == 3):
                p.font.bold = True
                p.font.color.rgb = COLOR_ALERT_ROSE
            elif row_idx == 4:
                p.font.bold = True
                p.font.color.rgb = COLOR_PRIMARY_FOREST
            else:
                p.font.color.rgb = COLOR_TEXT_MAIN

    # ==========================================
    # SLIDE 19: 進階演練：成績深入學情分析與圖表自動生成
    # ==========================================
    s19 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s19, "進階演練：成績深入學情分析與圖表自動生成", "單元四：Excel 自動化", 19)

    prompt19 = """請讀取剛剛建立的「高二電子科專業實習第一次段考成績表.xlsx」，進行深入學情分析：
1. 請繪製一張「實作成績分布級距長條圖」（以 10 分為一個級距），存成圖檔放入工作區。
2. 請列出分析摘要：
   - 班上實作前三名同學是誰？儀器操作亮點在哪？
   - 筆試或實作不及格人數各有幾位？
3. 請幫我寫一段 200 字左右的「技高專業實習學情檢討摘要」，適合直接複製到我的實習教學日誌或教學研究會報告中。"""
    add_prompt_card(s19, Inches(0.8), Inches(1.8), Inches(11.7), Inches(4.8), "實戰提示詞範本 05：學情診斷與視覺化圖表生成", prompt19, 14.0)

    # ==========================================

    # ==========================================
    # SLIDE 20: 成果發布 01：GitHub 帳號註冊 ✕ 建立 www 倉庫
    # ==========================================
    s20 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s20, "成果發布 01：免費帳號開立 ✕ 建立專屬倉庫 (www)", "單元四：成果網頁化與 GitHub 發布", 20)

    # 左欄：帳號開立
    add_card(s20, Inches(0.8), Inches(1.8), Inches(5.6), Inches(4.8))
    tb_l20 = s20.shapes.add_textbox(Inches(1.1), Inches(2.1), Inches(5.0), Inches(4.2))
    tf_l20 = tb_l20.text_frame
    tf_l20.word_wrap = True
    p = tf_l20.paragraphs[0]
    p.text = "📝 步驟一：註冊個人 GitHub 帳號"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_FOREST
    p.font.name = FONT_FAMILY
    p.space_after = Pt(10)

    gh_steps = [
        ("1. 前往官網", "瀏覽器打開 github.com，點右上角綠色「Sign up」。"),
        ("2. 填寫三要素", "輸入信箱（收驗證碼）➔ 設定高強度密碼 ➔ 決定 Username。"),
        ("⭐ 命名即網址！", "Username 會直接變成個人教學網站網址！例如簡老師為 stj580508，日後專屬站點即為 stj580508.github.io。"),
        ("3. 驗證啟用", "完成動物旋轉拼圖防偽驗證，收信填 8 位數驗證碼即可！")
    ]
    for title, desc in gh_steps:
        pt = tf_l20.add_paragraph()
        pt.text = title
        pt.font.size = Pt(15)
        pt.font.bold = True
        pt.font.color.rgb = COLOR_CARAMEL if "命名" in title else COLOR_PRIMARY_FOREST
        pt.font.name = FONT_FAMILY
        pd = tf_l20.add_paragraph()
        pd.text = desc
        pd.font.size = Pt(13)
        pd.font.color.rgb = COLOR_TEXT_MUTED
        pd.font.name = FONT_FAMILY
        pd.space_after = Pt(6)

    # 右欄：建立倉庫
    add_card(s20, Inches(6.9), Inches(1.8), Inches(5.6), Inches(4.8), border_color=COLOR_CARAMEL)
    tb_r20 = s20.shapes.add_textbox(Inches(7.2), Inches(2.1), Inches(5.0), Inches(4.2))
    tf_r20 = tb_r20.text_frame
    tf_r20.word_wrap = True
    p = tf_r20.paragraphs[0]
    p.text = "📂 步驟二：建立名為 www 的儲存庫"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = COLOR_CARAMEL
    p.font.name = FONT_FAMILY
    p.space_after = Pt(10)

    repo_steps = [
        ("1. 點擊右上角「+」號", "在 GitHub 畫面右上角點選「New repository」。"),
        ("2. 倉庫名稱填寫 www", "Repository name 輸入 www，代表個人教學首頁根目錄。"),
        ("3. 務必勾選 Public（公開）", "勾選 Public，其餘 README、.gitignore 保持預設空白免勾。"),
        ("4. 點綠色按鈕建立", "按下「Create repository」，取得網址：github.com/stj580508/www。")
    ]
    for title, desc in repo_steps:
        pt = tf_r20.add_paragraph()
        pt.text = title
        pt.font.size = Pt(15)
        pt.font.bold = True
        pt.font.color.rgb = COLOR_PRIMARY_FOREST
        pt.font.name = FONT_FAMILY
        pd = tf_r20.add_paragraph()
        pd.text = desc
        pd.font.size = Pt(13)
        pd.font.color.rgb = COLOR_TEXT_MUTED
        pd.font.name = FONT_FAMILY
        pd.space_after = Pt(6)

    # ==========================================
    # SLIDE 21: 成果發布 02：產生安全 Token 權杖 (馬賽克保護)
    # ==========================================
    s21 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s21, "成果發布 02：產生個人安全金鑰 (Personal Access Token)", "單元四：成果網頁化與 GitHub 發布", 21)

    # 左欄：為什麼需要金鑰與產生步驟
    add_card(s21, Inches(0.8), Inches(1.8), Inches(5.6), Inches(4.8))
    tb_l21 = s21.shapes.add_textbox(Inches(1.1), Inches(2.1), Inches(5.0), Inches(4.2))
    tf_l21 = tb_l21.text_frame
    tf_l21.word_wrap = True
    p = tf_l21.paragraphs[0]
    p.text = "🔑 步驟三：為什麼不能用登入密碼？"
    p.font.size = Pt(19)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_FOREST
    p.font.name = FONT_FAMILY
    p.space_after = Pt(8)

    token_why = [
        ("🛡️ GitHub 資安規範升級", "為防止密碼遭竊，GitHub 禁止終端機推播使用傳統密碼，必須改用 Personal Access Token (PAT)。"),
        ("1. 前往金鑰設定頁面", "瀏覽器前往 github.com/settings/tokens，點 Generate new token (classic)。"),
        ("2. Note 填寫 antigravity", "有效期限選 No expiration 或 30 天。"),
        ("3. ⭐ 關鍵勾選 repo", "請務必將第一個「repo」（倉庫完整存取）打勾！"),
        ("4. 複製金鑰代碼", "點擊最下方綠色按鈕，複製開頭為 ghp_ 的專屬權杖！")
    ]
    for title, desc in token_why:
        pt = tf_l21.add_paragraph()
        pt.text = title
        pt.font.size = Pt(14)
        pt.font.bold = True
        pt.font.color.rgb = COLOR_CARAMEL if "關鍵" in title else COLOR_PRIMARY_FOREST
        pt.font.name = FONT_FAMILY
        pd = tf_l21.add_paragraph()
        pd.text = desc
        pd.font.size = Pt(12)
        pd.font.color.rgb = COLOR_TEXT_MUTED
        pd.font.name = FONT_FAMILY
        pd.space_after = Pt(4)

    # 右欄：資安守則與馬賽克保護
    add_card(s21, Inches(6.9), Inches(1.8), Inches(5.6), Inches(4.8), border_color=COLOR_CARAMEL)
    tb_r21 = s21.shapes.add_textbox(Inches(7.2), Inches(2.1), Inches(5.0), Inches(4.2))
    tf_r21 = tb_r21.text_frame
    tf_r21.word_wrap = True
    p = tf_r21.paragraphs[0]
    p.text = "🔒 資安第一：金鑰遮罩防洩漏守則"
    p.font.size = Pt(19)
    p.font.bold = True
    p.font.color.rgb = COLOR_CARAMEL
    p.font.name = FONT_FAMILY
    p.space_after = Pt(10)

    security_rules = [
        ("🚨 金鑰等同於您的臨時身分證！", "任何人拿到您的 Token 都可以修改或覆寫您的雲端倉庫。"),
        ("🙈 示範或講義中必須打馬賽克：", "在簡報或公開錄影時，代碼必須做馬賽克遮罩處理："),
        ("   代碼示範：", "token = 'ghp_████████████████████████████████'"),
        ("⚡ 代理人自動化驗證機制：", "只要在當次對話中將 Token 提供給 Antigravity，代理人即自主完成 Git 身分驗證並推播，無需存留硬碟明文檔案！")
    ]
    for title, desc in security_rules:
        pt = tf_r21.add_paragraph()
        pt.text = title
        pt.font.size = Pt(14)
        pt.font.bold = True
        pt.font.color.rgb = COLOR_ALERT_ROSE if "🚨" in title else COLOR_PRIMARY_FOREST
        pt.font.name = FONT_FAMILY
        pd = tf_r21.add_paragraph()
        pd.text = desc
        pd.font.size = Pt(12)
        pd.font.color.rgb = COLOR_TEXT_MAIN
        pd.font.name = FONT_FAMILY
        pd.space_after = Pt(6)

    # ==========================================
    # SLIDE 22: 成果發布 03：代理人推播動作鏈 ✕ 網站上線 (Prompt 06)
    # ==========================================
    s22 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s22, "成果發布 03：代理人一鍵推播 ✕ GitHub Pages 上線 (Prompt 06)", "單元四：成果網頁化與 GitHub 發布", 22)

    # 左欄：大師級提示語 06
    add_card(s22, Inches(0.8), Inches(1.8), Inches(5.6), Inches(4.8))
    tb_l22 = s22.shapes.add_textbox(Inches(1.1), Inches(2.1), Inches(5.0), Inches(4.2))
    tf_l22 = tb_l22.text_frame
    tf_l22.word_wrap = True
    p = tf_l22.paragraphs[0]
    p.text = "📜 實戰提示詞 06：全自動推播咒語"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_FOREST
    p.font.name = FONT_FAMILY
    p.space_after = Pt(8)

    prompt_box22 = [
        "💡 可直接反白複製使用：",
        "我已在 GitHub 建立好公開儲存庫：https://github.com/stj580508/www，",
        "個人金鑰為 ghp_【貼上您的金鑰】。",
        "請幫我執行以下自動化連續動作鏈：",
        "1. 將今日的「基本電學題庫與標準電路圖」做成簡老師個人網站版面 (index.html)。",
        "2. 將 index.html、circuits/ 與 PPTX 投影片 Commit 並推播到 main 分支。",
        "3. 呼叫 GitHub API 為我啟用 GitHub Pages 靜態伺服器！",
        "4. 回報我最終公開上線的個人教學網站網址！"
    ]
    for line in prompt_box22:
        pl = tf_l22.add_paragraph()
        pl.text = line
        if "💡" in line:
            pl.font.size = Pt(12)
            pl.font.bold = True
            pl.font.color.rgb = COLOR_CARAMEL
        else:
            pl.font.size = Pt(12)
            pl.font.color.rgb = COLOR_TEXT_MAIN
        pl.font.name = FONT_FAMILY
        pl.space_after = Pt(4)

    # 右欄：上線成果驗收
    add_card(s22, Inches(6.9), Inches(1.8), Inches(5.6), Inches(4.8), border_color=COLOR_CARAMEL)
    tb_r22 = s22.shapes.add_textbox(Inches(7.2), Inches(2.1), Inches(5.0), Inches(4.2))
    tf_r22 = tb_r22.text_frame
    tf_r22.word_wrap = True
    p = tf_r22.paragraphs[0]
    p.text = "🎉 成果大功告成：全世界隨點隨看！"
    p.font.size = Pt(19)
    p.font.bold = True
    p.font.color.rgb = COLOR_CARAMEL
    p.font.name = FONT_FAMILY
    p.space_after = Pt(10)

    final_achieve = [
        ("🌐 官方專屬上線網址：", "👉 https://stj580508.github.io/www/"),
        ("📘 專屬教學專頁已上線：", "👉 https://stj580508.github.io/www/github_tutorial.html"),
        ("📱 行動學習亮點：", "無論用手機、平板或智慧黑板，掃描 QR Code 就能即時檢視互動題庫與標準電路圖！"),
        ("⚡ 代理 AI 革命性價值：", "老師完全不用敲複雜終端機指令，自然語言即可完成「備課 ➔ 出題 ➔ 繪圖 ➔ 網頁化 ➔ 雲端發布」一條龍！")
    ]
    for title, desc in final_achieve:
        pt = tf_r22.add_paragraph()
        pt.text = title
        pt.font.size = Pt(14)
        pt.font.bold = True
        pt.font.color.rgb = COLOR_PRIMARY_FOREST
        pt.font.name = FONT_FAMILY
        pd = tf_r22.add_paragraph()
        pd.text = desc
        pd.font.size = Pt(12)
        pd.font.color.rgb = COLOR_TEXT_MAIN
        pd.font.name = FONT_FAMILY
        pd.space_after = Pt(6)

    # ==========================================
    # SLIDE 23: 今日總結 ＆ 下週第二場精彩預告！
    # ==========================================
    s23 = prs.slides.add_slide(blank_layout)
    add_slide_scaffolding(s23, "今日收穫總結 ＆ 下週第二場精彩預告！", "研習總結", 23)

    add_card(s23, Inches(0.8), Inches(1.8), Inches(5.6), Inches(4.8))
    tb_l23 = s23.shapes.add_textbox(Inches(1.1), Inches(2.1), Inches(5.0), Inches(4.2))
    tf_l23 = tb_l23.text_frame
    tf_l23.word_wrap = True
    p = tf_l23.paragraphs[0]
    p.text = "🎉 今日已解鎖技能"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = COLOR_PRIMARY_FOREST
    p.font.name = FONT_FAMILY
    p.space_after = Pt(14)

    achieves = [
        "✅ 完成 Antigravity 工作區建置與沙盒安全檢測",
        "✅ 掌握零代碼一鍵自動安裝 Python/Excel 套件",
        "✅ 掌握基本電學電阻網路分析之結構化命題",
        "✅ 親歷電路圖文並茂 vs 開天窗之實測對比震撼",
        "✅ 掌握 MCP 磨課師自動化：人機接力與急救包",
        "✅ 實戰 Excel 自動算技高加權成績與學情分析圖表",
        "✅ 掌握 GitHub 免費註冊 ✕ 成果一鍵轉網頁推播"
    ]
    for a in achieves:
        pa = tf_l23.add_paragraph()
        pa.text = a
        pa.font.size = Pt(13)
        pa.font.bold = True
        pa.font.color.rgb = COLOR_TEXT_MAIN
        pa.font.name = FONT_FAMILY
        pa.space_after = Pt(6)

    add_card(s23, Inches(6.9), Inches(1.8), Inches(5.6), Inches(4.8), border_color=COLOR_CARAMEL)
    tb_r23 = s23.shapes.add_textbox(Inches(7.2), Inches(2.1), Inches(5.0), Inches(4.2))
    tf_r23 = tb_r23.text_frame
    tf_r23.word_wrap = True
    p = tf_r23.paragraphs[0]
    p.text = "🔥 下週第二場重頭戲預告"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = COLOR_CARAMEL
    p.font.name = FONT_FAMILY
    p.space_after = Pt(14)

    teasers = [
        "🌐 零程式碼做網頁：一句話做出「課堂隨機抽籤輪盤」與「自動計分測驗互動網頁」！",
        "🎬 AI 做簡報與做影片：教案一鍵轉 PPT，再自動配語音旁白做成微課教學影片！",
        "👥 多代理人協同備課：讓兩個 AI 一個出題、一個挑錯審查！",
        "🎁 課後任務：試著將自己的教材或題庫推上 GitHub，下週開場互相觀摩！"
    ]
    for t in teasers:
        pt = tf_r23.add_paragraph()
        pt.text = t
        pt.font.size = Pt(14)
        pt.font.color.rgb = COLOR_TEXT_MAIN
        pt.font.name = FONT_FAMILY
        pt.space_after = Pt(10)

    try:
        prs.save(output_path)
        print(f"20-Slide Animated Large Font PowerPoint generated successfully at: {output_path}")
    except PermissionError:
        alt_path = output_path.replace(".pptx", "_v2.pptx")
        prs.save(alt_path)
        print(f"Original locked, saved to: {alt_path}")

if __name__ == "__main__":
    out_file = r"j:\我的雲端硬碟\antigravity\教師研習\第一場_Antigravity代理AI教學實戰_大字動畫版.pptx"
    create_deck(out_file)
