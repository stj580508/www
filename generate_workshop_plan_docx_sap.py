# -*- coding: utf-8 -*-
"""
生成高中職教師研習實施計畫與課程大綱 Word 檔
規範要求：
- 計畫抬頭：115學年度第一學期高中優質化輔助方案（高優計畫）教師實務增能研習
- 研習時段：兩天、每天下午 1 點到 4 點 (13:00 - 16:00)
- 每節課標準 50 分鐘，共 6 節課（中間休息不列）
- 完整涵蓋九大單元內容
- 專業高優計畫公文報部/核銷/公告規格
"""

import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)

def set_table_borders(table, color="D1D5DB", sz="4", val="single"):
    tblPr = table._tbl.tblPr
    tblBorders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(tblBorders)

def create_document():
    doc = Document()
    
    # 邊距設定
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.85)
        section.right_margin = Inches(0.85)
        
        # 頁首與頁尾
        header = section.header
        hp = header.paragraphs[0]
        hp.text = "115學年度第一學期高級中等學校適性學習社區教育資源均質化暨高中優質化輔助方案（高優計畫）教師實務增能研習實施計畫"
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hp.style.font.name = "微軟正黑體"
        hp.style.font.size = Pt(8)
        hp.style.font.color.rgb = RGBColor(148, 163, 184)
        
        footer = section.footer
        fp = footer.paragraphs[0]
        fp.text = "115學年度第一學期高優計畫｜兩天每日 3 節課（每節 50 分鐘）・共 6 節課 6 小時"
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fp.style.font.name = "微軟正黑體"
        fp.style.font.size = Pt(8.5)
        fp.style.font.color.rgb = RGBColor(148, 163, 184)

    # 標準字體
    normal_style = doc.styles['Normal']
    normal_style.font.name = '微軟正黑體'
    normal_style.font.size = Pt(10.5)
    normal_style.font.color.rgb = RGBColor(30, 41, 59)
    normal_style.paragraph_format.line_spacing = 1.25

    # 1. 主標題（高優計畫標準公文抬頭）
    title_p = doc.add_paragraph()
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(4)
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_t1 = title_p.add_run("115學年度第一學期高中優質化輔助方案（高優計畫）\n教師實務增能研習實施計畫\n")
    run_t1.font.size = Pt(13.5)
    run_t1.font.bold = True
    run_t1.font.color.rgb = RGBColor(71, 85, 105)

    run_t2 = title_p.add_run("「AI Agent 代理人賦能技高教學實戰」\n兩天（每日3節／每節50分鐘／共6節課）課程大綱暨實施計畫書")
    run_t2.font.size = Pt(17)
    run_t2.font.bold = True
    run_t2.font.color.rgb = RGBColor(37, 99, 235)

    # 副標題標籤
    tag_p = doc.add_paragraph()
    tag_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tag_p.paragraph_format.space_after = Pt(16)
    run_tag = tag_p.add_run("【以 Antigravity 代理人為核心・高職基本電學、虛擬儀表、手寫成績校正、YOLO視覺AI與ESP32物聯網】")
    run_tag.font.size = Pt(9.5)
    run_tag.font.bold = True
    run_tag.font.color.rgb = RGBColor(13, 148, 136)

    # 2. 研習基本資訊卡片表格
    info_table = doc.add_table(rows=7, cols=2)
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(info_table, color="CBD5E1", sz="6")
    
    headers_data = [
        ("計畫依據", "教育部國民及學前教育署 115 學年度高中優質化輔助方案（高優計畫）實施要點辦理"),
        ("研習主題", "AI Agent 代理人賦能：基本電學、微算機物聯網與邊緣 AI 專題整合教學實戰"),
        ("研習時程", "兩天（第一天 13:00 - 16:00、第二天 13:00 - 16:00，每日 3 節課，每節 50 分鐘，共計 6 節課）"),
        ("研習對象", "技術型高級中等學校（高職）電機與電子群、資訊科、電子科、控制科、電機科等專業群科教師"),
        ("核心工具", "Google Antigravity AI 代理人系統、Python 3.10+、OpenCV、Ultralytics YOLO、ESP32 MicroPython"),
        ("實施方式", "一人一機上機實作、模組化教學演練、專案提示詞 (Prompt) 現場實證產出"),
        ("線上資源", "學員可存取研習專屬入口網站與 GitHub Pages 雲端教材庫 (https://stj580508.github.io/www/)")
    ]
    
    col_widths = [Inches(1.5), Inches(5.3)]
    for i, (label, val) in enumerate(headers_data):
        row = info_table.rows[i]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width, c1.width = col_widths[0], col_widths[1]
        set_cell_background(c0, "F1F5F9")
        set_cell_background(c1, "FFFFFF" if i % 2 == 0 else "F8FAFC")
        set_cell_margins(c0, 60, 60, 100, 100)
        set_cell_margins(c1, 60, 60, 100, 100)
        
        p0 = c0.paragraphs[0]
        r0 = p0.add_run(label)
        r0.font.bold = True
        r0.font.size = Pt(9.5)
        r0.font.color.rgb = RGBColor(51, 65, 85)
        
        p1 = c1.paragraphs[0]
        r1 = p1.add_run(val)
        r1.font.size = Pt(9.5)
        r1.font.color.rgb = RGBColor(15, 23, 42)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 3. 研習宗旨與核心目標
    h1 = doc.add_heading("壹、 研習宗旨與高優目標對應", level=1)
    h1.style.font.name = "微軟正黑體"
    h1.style.font.size = Pt(13)
    h1.style.font.color.rgb = RGBColor(30, 58, 138)

    p_goal = doc.add_paragraph()
    p_goal.add_run(
        "本計畫依據「115學年度高中優質化輔助方案」提升教師專業教學知能與實作創新教學法之核心指標，專為技術型高中專業群科專任教師量身規劃。"
        "為配合高職課務作息、降低全天公假之排課負擔，本研習採「兩天、每日下午 13:00 - 16:00」之課堂節奏，嚴格以「50 分鐘為一節課（每日 3 節課，共 6 節課）」進行標準化教學安排。\n"
        "課程導入次世代「AI Agent 代理人架構（Google Antigravity）」，掌握「自主規劃 ➔ 工具調用 ➔ 自動寫碼 ➔ 自我除錯 ➔ 成果發布」完整工作流。"
        "內容精實涵蓋高職「基本電學」與「微處理機實習」九大核心教材模組：從統測整數題庫與向量電路圖、28頁課堂簡報、微課短片、手寫成績多源校正，"
        "進階至示波器/信號源雙機虛擬量測、MCP 瀏覽器自動化、YOLO 邊緣視覺 AI 與 ESP32 實體單晶片物聯網，達成「課堂立即能用、教務大幅省時、專題戰力升級」之全方位高優效益。"
    )

    # 4. 課程配當表（每節 50 分鐘，共 6 節課，中間休息不列）
    h2 = doc.add_heading("貳、 兩天課程配當表（每節 50 分鐘・每日 3 節課・共 6 節課）", level=1)
    h2.style.font.name = "微軟正黑體"
    h2.style.font.size = Pt(13)
    h2.style.font.color.rgb = RGBColor(30, 58, 138)

    p_time_note = doc.add_paragraph()
    p_time_note.add_run("※ 課程規範：每日下午 13:00 至 16:00 授課，以標準 50 分鐘為一節課，兩天共計 6 節課（核發全國教師在職進修網研習時數 6 小時）。")
    p_time_note.paragraph_format.space_after = Pt(4)
    p_time_note.style.font.size = Pt(9)
    p_time_note.style.font.color.rgb = RGBColor(100, 116, 139)

    table_sched = doc.add_table(rows=1, cols=4)
    table_sched.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table_sched, color="CBD5E1", sz="6")
    
    headers = ["節次與時段", "涵蓋教材單元", "課程核心主題與教學實作內容", "教學方法與產出檢核"]
    hdr_row = table_sched.rows[0]
    col_w = [Inches(1.3), Inches(1.4), Inches(2.9), Inches(1.4)]
    for j, h in enumerate(headers):
        cell = hdr_row.cells[j]
        cell.width = col_w[j]
        set_cell_background(cell, "2563EB")
        set_cell_margins(cell, 80, 80, 80, 80)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(255, 255, 255)

    schedule_data = [
        # 第一天 (Day 1 - 每日 3 節課：AI 代理人賦能、教學備課與教務自動化)
        ("【第一天 第 1 節】\n13:00 - 13:50\n(50 分鐘)", "單元 01\n基本電學題庫與\nIEEE 標準電路圖", 
         "● 統測命題整數解、防呆誘答選項設計原則。\n● Schemdraw 電路圖自動繪製：鋸齒電阻與向量圖渲染。\n● 實作：運用 Prompt 引導 AI 生成附帶詳解之段考題庫。", 
         "實作產出：\n標準電路圖 SVG/PNG 與 Word/HTML 試卷排版。"),

        ("【第一天 第 2 節】\n14:00 - 14:50\n(50 分鐘)", "單元 02 & 03\n28 頁高質感簡報與\n微課短片多媒體合成", 
         "● Python-pptx 模組多版型（極簡瑞士、科技黑曜）自動生成。\n● 單一 Prompt 自動生成 28 頁基本電學課堂教學投影片。\n● 60 秒微課短片分鏡節奏設計與 MoviePy 批次多媒體合成。", 
         "實作產出：\n完整 28 頁教學簡報檔 (.pptx) 與微課短片腳本。"),

        ("【第一天 第 3 節】\n15:00 - 15:50\n(50 分鐘)", "單元 04 & 07\n手寫成績 AI 正規化與\nGitHub Pages 發布", 
         "● 單元 07：解決高職實習多張手寫紙本謄寫痛點，AI 識別「X/V」記號，openpyxl 動態加權核算與粉紅警示底色。\n● 單元 04：GitHub Pages 免費 CDN 託管，PAT 金鑰安全遮罩發布教學成果與密碼查詢網頁。", 
         "實作產出：\n手寫成績正規化 Excel 總表 (.xlsx) 與個人教學成果網站。"),

        # 第二天 (Day 2 - 每日 3 節課：虛擬量測儀表、邊緣 AI 視覺與單晶片實戰)
        ("【第二天 第 4 節】\n13:00 - 13:50\n(50 分鐘)", "單元 05\n示波器與信號產生器\n雙機互動量測系統", 
         "● 解決實習工場儀表不足與操作易燒毀問題。\n● 示波器與信號產生器純網頁 Canvas 動態波形連動。\n● 完整修改後 Master Prompt 架構剖析與學生探究學習單生成。", 
         "實作產出：\n雙機互動網頁模擬器與學生探究量測學習單。"),

        ("【第二天 第 5 節】\n14:00 - 14:50\n(50 分鐘)", "單元 06 & 08\nMCP 數位研習自動化與\nYOLO 視覺 AI 應用", 
         "● 單元 06：Model Context Protocol (MCP) 架構，Chrome DevTools 代理人自主點擊導航與研習巡檢。\n● 單元 08：YOLO 邊緣 AI 視覺四步 SOP（標註 ➔ Nano 模型訓練 ➔ WebCam 推論 ➔ Serial 序列埠警報）。", 
         "實作產出：\nMCP 自動化操作腳本與即時鏡頭 AI 辨識系統。"),

        ("【第二天 第 6 節】\n15:00 - 15:50\n(50 分鐘)", "單元 09\nESP32 實驗板物聯網\n快速上手與綜合發表", 
         "● MicroPython 免編譯優勢：Thonny IDE 快速連接與除錯。\n● GPIO/PWM 呼吸燈/ADC 電壓量測 ➔ 內建 Web Server 獨立儀表板。\n● AI 一鍵產出韌體與 HTML 介面；研習成果綜合發表與 Q&A。", 
         "實作產出：\nESP32 獨立 Web Server 韌體程式與兩天成果驗收。")
    ]

    for row_data in schedule_data:
        row = table_sched.add_row()
        for c_idx, text in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.width = col_w[c_idx]
            set_cell_background(cell, "FFFFFF" if len(table_sched.rows) % 2 == 0 else "F8FAFC")
            set_cell_margins(cell, 60, 60, 80, 80)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 2 else WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(text)
            run.font.size = Pt(8.5)
            if c_idx < 2:
                run.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 5. 九大單元核心教材內容詳解
    h3 = doc.add_heading("參、 研習核心九大單元內容詳解與教學框架", level=1)
    h3.style.font.name = "微軟正黑體"
    h3.style.font.size = Pt(13)
    h3.style.font.color.rgb = RGBColor(30, 58, 138)

    modules_detail = [
        {
            "num": "單元 01",
            "name": "基本電學題庫設計與 IEEE 標準電路圖自動化繪製",
            "section": "第 1 節",
            "core": "電機電子群統測命題邏輯、整數手算設計、Schemdraw 電路圖自動化繪製",
            "content": [
                "1. 高職基本電學命題痛點分析：避免小數點繁複計算，利用「畢氏三元數 (3:4:5)」與「互補電阻」設計整數解電路。",
                "2. 誘答防呆選項生成：透過 AI 分析學生常見盲點（如未開根號、漏算內阻、相角正負顛倒），自動產出高鑑別度選擇題選項。",
                "3. Schemdraw 向量繪圖程式碼實戰：運用 Python 一鍵繪製符合國際 IEEE 標準之鋸齒狀電阻、電感、電容與電源符號，匯出無失真向量圖 (SVG) 與高解析圖檔。"
            ]
        },
        {
            "num": "單元 02",
            "name": "28 頁高質感課堂教學簡報 (PPT) 自動化生成系統",
            "section": "第 2 節 (前半)",
            "core": "Python-pptx 模組、現代科技排版版型、大字多圖與互動式教案投影片",
            "content": [
                "1. 簡報自動化原理：擺脫樣板簡陋限制，利用 Python-pptx 腳本直接控制形狀、顏色、文字階層與邊界。",
                "2. 專屬多風格主題切換：涵蓋「新野獸派」、「瑞士極簡」、「科技黑曜」與「Office 旗艦商務風」。",
                "3. 實作教學：輸入單一教學主題 Prompt，即可在 5 秒內產出具備「章節大綱、學習重點、原理詳解、公式推導卡、課後測驗」之 28 頁完整簡報。"
            ]
        },
        {
            "num": "單元 03",
            "name": "課堂前導與複習微課短片（短影音腳本與多媒體整合）",
            "section": "第 2 節 (後半)",
            "core": "教學短影音分鏡、MoviePy 程式合成、觀念破題與動態波形展示",
            "content": [
                "1. 微課短片架構學：針對現代學生專注力特性，設計「3 秒黃金開場 ➔ 15 秒核心觀念痛點 ➔ 30 秒動態圖解 ➔ 12 秒考題破解」之 60 秒短影音節奏。",
                "2. 腳本與提示詞手冊：AI 自動產出分鏡旁白、配樂情緒建議與動畫提示字元。",
                "3. 實作延伸：結合 MoviePy 腳本批次合成教學影片，快速上傳供學生課前預習。"
            ]
        },
        {
            "num": "單元 04",
            "name": "GitHub Pages 終身免費教學網站發布與資安防護 SOP",
            "section": "第 3 節 (後半)",
            "core": "Git 基礎、GitHub Pages 靜態託管、Personal Access Token (PAT) 金鑰遮罩資安守則",
            "content": [
                "1. 免費託管架構：擺脫學校主機維護難題，利用 GitHub 免費全球 CDN 發布個人專屬教學成果網站與互動教材。",
                "2. Git 推播標準 SOP：工作區初始化、遠端分支綁定、多檔案批次更新與推播流程。",
                "3. 教師必備資安防線：Personal Access Token (PAT) 申請規範、終端機金鑰遮罩技巧與權限最小化原則，杜絕資安外洩事故。"
            ]
        },
        {
            "num": "單元 05",
            "name": "示波器與信號產生器雙機互動教學系統與 Master Prompt 架構",
            "section": "第 4 節",
            "core": "儀器控制面板設計、Canvas 波形即時渲染、狀態機模式 (State Machine)、Master Prompt 架構剖析",
            "content": [
                "1. 實習虛擬化突破：解決實習工場「儀器數量不足、操作生疏容易燒壞」痛點，打造純網頁版雙機互動模擬器。",
                "2. 核心架構剖析：信號產生器輸出頻率/波形/振幅，實時驅動示波器畫面衰減、時基 (Time/Div) 與觸發 (Trigger) 變動。",
                "3. 完整修改後 Master Prompt 解析：結構化提示詞範本（包含角色、UI佈局、事件監聽、防呆檢查），學員可直接套用並衍生客製化其他電子量測儀器。"
            ]
        },
        {
            "num": "單元 06",
            "name": "Model Context Protocol (MCP) 數位研習自動化工作流",
            "section": "第 5 節 (前半)",
            "core": "MCP 協定核心機制、Chrome DevTools 代理人整合、全自動表單填答與數位研習紀錄",
            "content": [
                "1. 次世代 AI 協定：深入了解 Anthropic / Google 開源的 MCP 架構，如何賦予 AI 代理人直接操作本機軟體與瀏覽器的超能力。",
                "2. Chrome DevTools MCP 實機展示：AI 主動開啟瀏覽器、導航至數位學習平臺、辨識 DOM 元素、執行滑鼠點擊與文字輸入。",
                "3. 自動化巡檢實務：示範引導 AI 代理人自主巡檢研習資源、提取測驗題目並自動化整理成教學記錄檔。"
            ]
        },
        {
            "num": "單元 07",
            "name": "紙本手寫成績掃描檔正規化與智慧分析儀表板",
            "section": "第 3 節 (前半)",
            "core": "多源資料對齊、特殊手寫記號識別、openpyxl 自動化試算表、密碼保護成績查詢儀表板",
            "content": [
                "1. 解決實習教師最大痛點：多張分散紙本（隨堂小考、抽測驗收、總考查），筆跡包含「X（缺考）」與「V（通過）」之複雜情境。",
                "2. 自動化正規化：Python 跨表格精準對齊 35 位學生座號，動態試算加權總分（=ROUND(C*0.2+D*0.2+E*0.6, 1)）與等第，自動套用不及格粉紅底色警示。",
                "3. 成績發布雙軌制：一鍵輸出標準 .xlsx 活頁簿，並自動生成具備安全密碼保護的互動網頁儀表板 (grades.html)。"
            ]
        },
        {
            "num": "單元 08",
            "name": "YOLO 視覺 AI 辨識教學應用與邊緣端實作",
            "section": "第 5 節 (後半)",
            "core": "自訂資料集標註 (Roboflow)、Ultralytics YOLO 輕量化訓練、OpenCV 即時串流推論、硬體 Serial 觸發",
            "content": [
                "1. 邊緣 AI 教學四步 SOP：從自訂標註（電路板元件、瑕疵判別、防呆檢驗）到模型匯出標準格式 (YOLO format)。",
                "2. 輕量化模型實作：選用 yolov8n / yolo11n Nano 模型，在一般個人電腦 10 分鐘內完成遷移學習收斂。",
                "3. 即時視訊與硬體聯動：OpenCV 擷取 WebCam 畫面，繪製 Bounding Box 與信心度，當特定瑕疵被辨識時，透過 COM Port 發送訊號驅動實體機構警報。"
            ]
        },
        {
            "num": "單元 09",
            "name": "ESP32 實驗板物聯網快速上手與 AI 韌體輔助開發",
            "section": "第 6 節",
            "core": "MicroPython / Arduino C++ 雙軌開發、GPIO/PWM/ADC 控制、內建 Web Server 獨立儀表板、MQTT 雲端串接",
            "content": [
                "1. 技高微控制器教學新革命：破除傳統 C 語言除錯困難，善用 MicroPython + Thonny IDE「即寫即跑、免編譯」的高效學習優勢。",
                "2. 內建 Web Server 實戰：無需外部伺服器！ESP32 自建 Wi-Fi 伺服器，手機瀏覽器直連即可讀取感測器數值並 AJAX 非同步控制繼電器與燈光。",
                "3. AI 提示詞生成韌體：提供現成 Prompt 讓 AI 自動產出腳位定義表 (Pinout Table)、韌體程式碼與 HTML 前端介面，無縫串聯 YOLO 專題。"
            ]
        }
    ]

    for m in modules_detail:
        m_title = doc.add_paragraph()
        m_title.paragraph_format.space_before = Pt(6)
        m_title.paragraph_format.space_after = Pt(2)
        r_num = m_title.add_run(f"【{m['num']}】{m['name']} ")
        r_num.font.bold = True
        r_num.font.size = Pt(11)
        r_num.font.color.rgb = RGBColor(37, 99, 235)
        
        r_h = m_title.add_run(f"（對應課表：{m['section']}）")
        r_h.font.size = Pt(9.5)
        r_h.font.color.rgb = RGBColor(100, 116, 139)

        p_core = doc.add_paragraph()
        p_core.paragraph_format.space_before = Pt(0)
        p_core.paragraph_format.space_after = Pt(2)
        r_c_lbl = p_core.add_run("核心技術概念：")
        r_c_lbl.font.bold = True
        r_c_lbl.font.size = Pt(9)
        r_c_val = p_core.add_run(m['core'])
        r_c_val.font.size = Pt(9)
        r_c_val.font.color.rgb = RGBColor(71, 85, 105)

        for line in m['content']:
            p_cnt = doc.add_paragraph()
            p_cnt.paragraph_format.space_before = Pt(0)
            p_cnt.paragraph_format.space_after = Pt(2)
            p_cnt.paragraph_format.left_indent = Inches(0.2)
            r_cnt = p_cnt.add_run(line)
            r_cnt.font.size = Pt(9)
            r_cnt.font.color.rgb = RGBColor(30, 41, 59)

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 6. 電腦教室環境設備需求表
    h4 = doc.add_heading("肆、 研習軟硬體環境需求清單（電腦教室整備需求）", level=1)
    h4.style.font.name = "微軟正黑體"
    h4.style.font.size = Pt(13)
    h4.style.font.color.rgb = RGBColor(30, 58, 138)

    table_req = doc.add_table(rows=1, cols=3)
    table_req.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table_req, color="CBD5E1", sz="6")
    
    req_hdrs = ["項目類別", "必備規格與軟體清單", "備註與承辦配合事項"]
    req_w = [Inches(1.5), Inches(3.8), Inches(1.7)]
    for j, h in enumerate(req_hdrs):
        cell = table_req.rows[0].cells[j]
        cell.width = req_w[j]
        set_cell_background(cell, "0F766E")
        set_cell_margins(cell, 80, 80, 80, 80)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(h)
        run.font.bold = True
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(255, 255, 255)

    req_data = [
        ("硬體設備\n(學員機)", "● 個人電腦 (Windows 10/11，建議 16GB RAM)\n● 具備 USB 視訊攝影機 (WebCam，用於第 5 節 YOLO 辨識)\n● 講師自備 ESP32 開發板供現場學員連線實作", "一人一機實作\n教室需配置外網連線"),
        ("網路環境", "● 穩定對外寬頻網路（供存取 Google Antigravity、GitHub）\n● 開放 GitHub (github.com)、Colab 與 Python Package 存取 port", "學校防火牆需放行\n避免擋 GitHub/CDN"),
        ("基礎軟體\n(已預裝或現場安裝)", "● Google Chrome 瀏覽器 (最新版)\n● Python 3.10 或以上環境 (建議包含 pip)\n● Visual Studio Code (VS Code) 編輯器\n● Thonny IDE (MicroPython 燒錄工具)", "可於研習前寄發安裝指引\n或現場由講師帶領配置"),
        ("學員必備帳號", "● 個人 Google 帳號 (使用 Antigravity / Gemini)\n● 個人 GitHub 免費帳號 (使用 GitHub Pages 網頁託管與 Git 推播)", "請承辦單位於行前通知中\n提醒老師先註冊 GitHub")
    ]

    for row_data in req_data:
        row = table_req.add_row()
        for c_idx, text in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.width = req_w[c_idx]
            set_cell_background(cell, "FFFFFF" if len(table_req.rows) % 2 == 0 else "F8FAFC")
            set_cell_margins(cell, 60, 60, 80, 80)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx == 1 else WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(text)
            run.font.size = Pt(8.5)
            if c_idx == 0:
                run.font.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(10)

    # 7. 預期成效與結案效益（對應高優計畫 KPIs）
    h5 = doc.add_heading("伍、 預期高優成效與研習產出（供公文陳報與核銷報告使用）", level=1)
    h5.style.font.name = "微軟正黑體"
    h5.style.font.size = Pt(13)
    h5.style.font.color.rgb = RGBColor(30, 58, 138)

    benefits = [
        ("一、 產出即戰力教學成果", "每位參訓教師完成 6 節課研習後，皆可帶回一份「專業基本電學 28 頁簡報 (.pptx)」、一套「標準題庫與電路圖檔」、以及「已部署於 GitHub Pages 的個人專屬教學網站」，次日即可應用於高職課堂。"),
        ("二、 解決日常教務痛點", "掌握手寫登記成績多源校對自動化技術，徹底消除小考抽測人工謄寫算錯風險，平均可為實習教師節省 80% 以上的學期末成績加權核算時間。"),
        ("三、 升級技高專題實力", "成功打通從「AI 虛擬儀器模擬」跨越至「YOLO 視覺 AI」與「ESP32 實體單晶片控制」的完整技術鏈，有效引導技高學生組隊參加全國專題競賽。"),
        ("四、 達成高優指標與教師共備社群", "完全符合高優計畫教師專業增能與跨領域教學發展目標，透過研習社群與線上開源教材庫 (GitHub)，促進跨校群科教師教材共享與永續共備。")
    ]

    for b_title, b_desc in benefits:
        p_b = doc.add_paragraph()
        p_b.paragraph_format.space_before = Pt(2)
        p_b.paragraph_format.space_after = Pt(3)
        r_bt = p_b.add_run(f"{b_title}：")
        r_bt.font.bold = True
        r_bt.font.size = Pt(9.5)
        r_bt.font.color.rgb = RGBColor(15, 23, 42)
        r_bd = p_b.add_run(b_desc)
        r_bd.font.size = Pt(9.5)
        r_bd.font.color.rgb = RGBColor(71, 85, 105)

    doc.add_paragraph().paragraph_format.space_after = Pt(12)

    # 8. 簽核欄位（高優計畫專用四級簽核）
    p_sign = doc.add_paragraph()
    p_sign.paragraph_format.space_before = Pt(14)
    p_sign.add_run("────────────────────────────────────────────────────────\n").font.color.rgb = RGBColor(203, 213, 225)
    r_sign_text = p_sign.add_run("承辦人：_______________　　高優業務組長：_______________　　實習主任：_______________　　校長：_______________")
    r_sign_text.font.size = Pt(9.5)
    r_sign_text.font.bold = True
    r_sign_text.font.color.rgb = RGBColor(100, 116, 139)

    # 存檔為專用檔名
    output_path = r"j:\我的雲端硬碟\antigravity\教師研習\115學年度上學期高優計畫教師實務增能研習實施計畫書.docx"
    doc.save(output_path)
    print(f"Successfully generated docx at: {output_path}")

if __name__ == "__main__":
    create_document()
