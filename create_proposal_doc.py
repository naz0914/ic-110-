# -*- coding: utf-8 -*-
import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=60, bottom=60, left=100, right=100):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'''
        <w:tcMar {nsdecls("w")}>
            <w:top w:w="{top}" w:type="dxa"/>
            <w:bottom w:w="{bottom}" w:type="dxa"/>
            <w:left w:w="{left}" w:type="dxa"/>
            <w:right w:w="{right}" w:type="dxa"/>
        </w:tcMar>
    ''')
    tcPr.append(tcMar)

def set_table_borders(table, color="B0C4DE", sz="4", val="single"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(f'''
        <w:tblBorders {nsdecls("w")}>
            <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:insideV w:val="none"/>
            <w:left w:val="none"/>
            <w:right w:val="none"/>
        </w:tblBorders>
    ''')
    tblPr.append(borders)

def format_run(run, font_name="微軟正黑體", font_size=10, bold=False, italic=False, color_rgb=(0, 0, 0)):
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = RGBColor(*color_rgb)
    rPr = run._r.get_or_add_rPr()
    rFonts = parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="{font_name}"/>')
    rPr.append(rFonts)

def add_styled_heading(doc, text, level=1):
    p = doc.add_paragraph()
    p.paragraph_format.keep_with_next = True
    if level == 1:
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(text)
        format_run(run, font_name="微軟正黑體", font_size=12, bold=True, color_rgb=(31, 78, 121))
    elif level == 2:
        p.paragraph_format.space_before = Pt(7)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(text)
        format_run(run, font_name="微軟正黑體", font_size=10.5, bold=True, color_rgb=(46, 117, 182))
    elif level == 3:
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(1)
        run = p.add_run(text)
        format_run(run, font_name="微軟正黑體", font_size=10, bold=True, color_rgb=(68, 84, 106))
    return p

def add_body_paragraph(doc, text, space_after=3.5, line_spacing=1.15):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing
    run = p.add_run(text)
    format_run(run, font_name="微軟正黑體", font_size=9.5, color_rgb=(40, 40, 40))
    return p

def add_centered_image(doc, img_path, caption_text, width=Inches(3.2)):
    if os.path.exists(img_path):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(3)
        p_img.paragraph_format.space_after = Pt(1)
        p_img.paragraph_format.keep_with_next = True
        run_img = p_img.add_run()
        run_img.add_picture(img_path, width=width)
        
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_after = Pt(5)
        p_cap.paragraph_format.keep_with_next = False
        rcap = p_cap.add_run(caption_text)
        format_run(rcap, font_name="微軟正黑體", font_size=8.5, bold=True, color_rgb=(90, 90, 90))

def main():
    doc = docx.Document()
    
    # 設置頁邊距 (約 1.8 公分，緊湊而優雅)
    for s in doc.sections:
        s.top_margin = Inches(0.7)
        s.bottom_margin = Inches(0.7)
        s.left_margin = Inches(0.75)
        s.right_margin = Inches(0.75)
        
        header = s.header
        hp = header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hrun = hp.add_run("CAD/VLSI Circuit Design 期末報告構想說明書 | 國立中興大學醫學工程研究所")
        format_run(hrun, font_name="微軟正黑體", font_size=8, color_rgb=(130, 130, 130))

    # --- 文件主標題 ---
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(2)
    p_title.paragraph_format.space_after = Pt(2)
    run_title = p_title.add_run("數位積體電路設計期末報告構想說明書")
    format_run(run_title, font_name="微軟正黑體", font_size=15, bold=True, color_rgb=(31, 78, 121))
    
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(6)
    run_sub = p_sub.add_run("高精度即時心電圖特徵偵測引擎之 ASIC 硬體電路設計 (ECG Detection Engine, EDE)")
    format_run(run_sub, font_name="微軟正黑體", font_size=11, bold=True, color_rgb=(68, 84, 106))

    # --- 學生與課程資訊表格 ---
    info_table = doc.add_table(rows=2, cols=4)
    info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(info_table, color="4F81BD", sz="4", val="single")
    
    row0 = info_table.rows[0]
    row1 = info_table.rows[1]
    
    data_map = [
        (row0.cells[0], "學生姓名", True), (row0.cells[1], "陳沐德", False),
        (row0.cells[2], "學　　號", True), (row0.cells[3], "[請填寫學號]", False),
        (row1.cells[0], "所屬系所", True), (row1.cells[1], "醫學工程研究所", False),
        (row1.cells[2], "指導老師", True), (row1.cells[3], "程德勝 教授", False)
    ]
    
    for cell, txt, is_h in data_map:
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(1.5)
        p.paragraph_format.space_before = Pt(1.5)
        run = p.add_run(txt)
        format_run(run, font_name="微軟正黑體", font_size=9, bold=is_h, color_rgb=(31, 78, 121) if is_h else (0,0,0))
        set_cell_margins(cell, top=40, bottom=40, left=60, right=60)
        if is_h:
            set_cell_background(cell, "EDF2F8")
            
    doc.add_paragraph().paragraph_format.space_after = Pt(3)

    # ==================== 一、 簡介與設計目的 ====================
    add_styled_heading(doc, "一、 簡介與設計動機 (Introduction & Biomedical Motivation)", level=1)
    
    add_body_paragraph(doc, 
        "心電圖 (Electrocardiogram, ECG) 是現代臨床醫學中最關鍵的心臟生理電訊號，用以評估心肌電位之去極化與再極化歷程。一個完整的心電週期（ECG Beat）主要包含 P 波、QRS 複合波與 T 波（如圖一）。其中，QRS 複合波代表心室去極化（Ventricular Depolarization）過程，具有斜率最陡峭與振幅最大的生理特徵。精準擷取 Q、R、S 點之時域座標，不僅為即時心率（Heart Rate）量測與心率變異度（Heart Rate Variability, HRV）分析之基礎，更為診斷心室早期收縮（PVC）、心室頻脈（VT）及急性心肌梗塞等致死性心律不整的重要臨床指標。"
    )
    
    add_body_paragraph(doc,
        "然而，在實際動態心電圖（Holter）或穿戴式醫療監測場域中，受測者呼吸運動、體表晃動與電極阻抗變化將引發嚴重的低頻「基線漂移 (Baseline Wander / DC-bias，約 0.05 ~ 1 Hz)」；此外，電網 60 Hz 諧波及肌肉震顫亦會疊加高頻肌電雜訊（EMG Noise，>35 Hz）。若未經前置濾波處理，細微的 Q、S 波谷極易被雜訊掩蔽，造成特徵點偵測嚴重失真。"
    )

    # 圖一：ECG Beat 示意圖
    add_centered_image(doc, r"c:\Users\milto\Desktop\ic期末報告\fig_p1_1.png", "圖一、 典型心電圖波形 (ECG Beat) 與 Q、R、S 關鍵特徵點示意圖", width=Inches(2.7))

    add_body_paragraph(doc,
        "本專題旨在以 ASIC 數位硬體實現專用「心電圖特徵偵測引擎」(ECG Detection Engine, EDE)。該硬體加速器接收 10-bit 取樣資料，經由管線化濾波架構即時去除 DC-bias 與高頻雜訊，並以專用狀態機在每 120 點之 Beat 區間內自動定位 Q、R、S 座標，於面積與運算延遲之間取得最佳平衡，滿足穿戴式醫療邊緣運算之極低功耗與高即時性需求。"
    )

    # ==================== 二、 預計達到之規格與設計流程 ====================
    add_styled_heading(doc, "二、 預計達到之規格與設計流程 (Specifications & ASIC Flow)", level=1)
    
    add_styled_heading(doc, "2.1 ASIC 實體設計與驗證流程", level=2)
    add_body_paragraph(doc,
        "本專題遵循標準 Cell-Based 數位 ASIC 設計流程，涵蓋前端 RTL 至後端實體驗證完整階段：\n"
        "1. 前端 RTL 設計與行為模擬：以 Verilog HDL 實現流式管線與雙緩衝架構，透過 Synopsys VCS / NC-Verilog 搭配競賽提供之 PAT1 與 PAT2 測試樣本進行模擬比對，驗證功能 100% 正確。\n"
        "2. 邏輯合成 (Logic Synthesis)：使用 Synopsys Design Compiler (DC) 搭配 Slow Library，對 Critical Path 進行時序收斂與邏輯閘優化，產出 Gate-level Netlist 與 SDF 檔。\n"
        "3. 自動佈局與繞線 (APR)：使用 Cadence Innovus (或 Synopsys ICC2) 實現 Macro Layout，依競賽規定完整配置 2um VDD/VSS 電源環 (Power Ring)、2um 垂直電源條 (Power Stripe)、標準單元電源軌 (Power Rail, follow pin) 與 Core Filler，並導出 GDSII 實體佈局檔。\n"
        "4. 實體驗證與後模擬：透過 Mentor Calibre DRC (或 ICV) 達成 DRC = 0 Errors，並以 APR 內建工具達成 LVS = 0 Violations，最後以 post-layout netlist 與 SDF 執行 Post-Sim (+ncmaxdelays)。"
    )

    add_styled_heading(doc, "2.2 預計目標硬體規格指標", level=2)
    
    spec_table = doc.add_table(rows=7, cols=3)
    spec_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(spec_table, color="4F81BD", sz="4", val="single")
    
    spec_data = [
        ("項目 (Parameter)", "設計目標值 (Target Value)", "說明與期末評分關聯"),
        ("製程技術 (Technology)", "TSMC ADFP 虛擬製程 (16nm FinFET) / 輔以 TSMC 0.13 μm", "符合期末報告加分項目；0.13 μm 為競賽基準"),
        ("目標工作時脈 (Clock)", "100 MHz (Clock Period = 10.0 ns)", "預估合成積極探索目標 125 ~ 166 MHz (6~8 ns)"),
        ("輸入訊號格式 (din)", "10-bit 正整數 (0 ~ 1023), 無正負號", "每 Beat 固定 120 筆訊號，總長度 2400 筆樣本"),
        ("特徵輸出規格 (Q, R, S)", "Q[11:0], R[11:0], S[11:0] (12-bit 無正負號)", "每 Beat 三點皆定位後，由 valid=1 觸發同步輸出"),
        ("系統吞吐率 (Throughput)", "1 Sample / Clock Cycle (串流模式)", "設計雙緩衝消除等待，busy 旗標常態維持於 Low"),
        ("評分等級目標 (Target Rank)", "等級 A (Rank A: Complete APR & DRC/LVS Clean)", "極小化 Area × Time (AT 積)，力爭最佳成績")
    ]
    
    for row_idx, row in enumerate(spec_table.rows):
        for col_idx, cell in enumerate(row.cells):
            txt = spec_data[row_idx][col_idx]
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if col_idx < 2 else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(txt)
            is_header = (row_idx == 0)
            format_run(run, font_name="微軟正黑體", font_size=8.5, bold=is_header, color_rgb=(255,255,255) if is_header else (0,0,0))
            set_cell_margins(cell, top=40, bottom=40, left=60, right=60)
            if is_header:
                set_cell_background(cell, "2F5597")
            elif row_idx % 2 == 1:
                set_cell_background(cell, "F2F5F9")

    p_cap_tab1 = doc.add_paragraph()
    p_cap_tab1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap_tab1.paragraph_format.space_after = Pt(4)
    rc1 = p_cap_tab1.add_run("表一、 EDE 晶片目標規格一覽表")
    format_run(rc1, font_name="微軟正黑體", font_size=8.5, bold=True, color_rgb=(90,90,90))

    add_styled_heading(doc, "2.3 訊號介面與腳位定義", level=2)
    
    io_table = doc.add_table(rows=7, cols=4)
    io_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(io_table, color="4F81BD", sz="4", val="single")
    
    io_data = [
        ("訊號名稱", "I/O", "寬度", "時序與電氣行為描述"),
        ("clk", "Input", "1 bit", "系統主時脈輸入訊號 (正緣觸發 Positive Edge Triggered)"),
        ("rst", "Input", "1 bit", "系統重置訊號 (Active-High)，高電位時初始化內部所有管線暫存器"),
        ("din", "Input", "10 bits", "原始 ECG 取樣訊號匯流排 (0 ~ 1023)，每 cycle 輸入一筆樣本"),
        ("busy", "Output", "1 bit", "電路忙碌旗標。Low: 允許 Host 輸入；High: 暫停資料輸入 (din 為 Hi-Z)"),
        ("valid", "Output", "1 bit", "特徵輸出有效旗標。High: 代表當前 Q、R、S 匯流排為有效輸出座標"),
        ("Q, R, S", "Output", "12 bits x 3", "心電圖 Beat 之特徵座標點 (0 ~ 2399)。收集齊全後同步輸出一個 cycle")
    ]
    
    for row_idx, row in enumerate(io_table.rows):
        for col_idx, cell in enumerate(row.cells):
            txt = io_data[row_idx][col_idx]
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if col_idx < 3 else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(txt)
            is_header = (row_idx == 0)
            format_run(run, font_name="微軟正黑體", font_size=8.5, bold=is_header, color_rgb=(255,255,255) if is_header else (0,0,0))
            set_cell_margins(cell, top=40, bottom=40, left=60, right=60)
            if is_header:
                set_cell_background(cell, "2F5597")
            elif row_idx % 2 == 1:
                set_cell_background(cell, "F2F5F9")

    p_cap_tab2 = doc.add_paragraph()
    p_cap_tab2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap_tab2.paragraph_format.space_after = Pt(6)
    rc2 = p_cap_tab2.add_run("表二、 EDE 晶片 I/O 訊號介面定義表")
    format_run(rc2, font_name="微軟正黑體", font_size=8.5, bold=True, color_rgb=(90,90,90))

    # ==================== 三、 電路演算法與定點數無損映射 ====================
    add_styled_heading(doc, "三、 電路演算法與定點數無損映射分析 (Algorithm & Arithmetic Analysis)", level=1)
    
    add_body_paragraph(doc,
        "競賽題目特別提示「運算過程中的小數位數需以最大精確度計算，以避免與標準解答不同」。本設計經數學分析發現：所有除法除數皆為 2 的冪次（16 與 64），本質上皆為純二進制移位運算。因此本架構採用完整無損的定點數硬體映射，達成 100% 絕對零捨入誤差 (Bit-Exact)："
    )

    # 圖二：EDE 電路運作流程
    add_centered_image(doc, r"c:\Users\milto\Desktop\ic期末報告\fig_p3_2.png", "圖二、 EDE 晶片四階段訊號處理與特徵點萃取資料流", width=Inches(3.8))

    add_styled_heading(doc, "3.1 步驟 1：16-tap 移動平均運算 (16-tap Moving Average)", level=2)
    add_body_paragraph(doc,
        "移動平均用以估計低頻基線變化：\n"
        "    MV(n) = [ x(n) + x(n-1) + ... + x(n-15) ] / 16\n"
        "• 邊界處理：超出左邊界（n < 15）之值一律以首筆訊號 x(0) 填補。\n"
        "• 無損映射：16 筆 10-bit 取樣相加之最大值為 1023 × 16 = 16368（需 14-bit Unsigned）。除以 16 即右移 4 bits，若將這 4 bits 作為精確小數保留，則 MV(n) 實質為「10-bit 整數 + 4-bit 小數」，在此運算階完全無任何捨入截斷誤差。"
    )

    add_styled_heading(doc, "3.2 步驟 2：去除基線漂移 (Remove DC-Bias)", level=2)
    add_body_paragraph(doc,
        "扣除基線漂移使訊號中心歸零：\n"
        "    Del_DC(n) = x(n) - MV(n)\n"
        "• 無損映射：將 x(n) 左移 4 bits 對齊小數點後相減：Del_DC_scaled(n) = (x(n) << 4) - Sum16(n)。其動態範圍介於 -16368 ~ +16368，硬體以 16-bit 帶符號數 (Signed [15:0], 2's complement) 暫存，運算精確度達 100%。"
    )

    add_styled_heading(doc, "3.3 步驟 3：15-tap 視窗濾波器 (15-tap Windowing FIR Filter)", level=2)
    add_body_paragraph(doc,
        "透過三角加權視窗濾除高頻雜訊：\n"
        "    WF(n) = [ ∑_{k=-7}^{7} Del_DC(n+k) × C(k+7) ] / 64\n"
        "其中係數 C[0:14] = [1, 2, 3, 4, 5, 6, 7, 8, 7, 6, 5, 4, 3, 2, 1]，加權總和恰為 64。\n"
        "• 邊界填補：超出左邊界以 Del_DC(0) 代替，超出右邊界（n > 2399）以 Del_DC(2399) 代替。\n"
        "• 對稱折疊與無乘法器架構：利用對稱性將 Del_DC(n-j) 與 Del_DC(n+j) 先行相加；係數乘法以移位加法樹 (Shift-Add Tree) 實現，完全免除實體乘法器，大幅縮減面積與 Critical Path。\n"
        "• 無損比較代換定理：由於電路最終目的僅輸出特徵點座標（Index），且除數 64 > 0，故滿足 WF(a) > WF(b) <=> 64×WF(a) > 64×WF(b)。硬體直接以「未除以 64 之 22-bit 帶符號整數」進行極值判讀，徹底消除除法截斷誤差，確保與 Golden Answer 100% 吻合！"
    )

    # 表三：定點數配置表
    fix_table = doc.add_table(rows=5, cols=5)
    fix_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(fix_table, color="4F81BD", sz="4", val="single")
    
    fix_data = [
        ("運算階段變數", "資料型態", "整數位元", "小數位元", "硬體總位元寬度與數值範圍"),
        ("原始訊號 x(n)", "Unsigned", "10 bits", "0 bit", "[9:0], 範圍 0 ~ 1023"),
        ("16 筆累加和 Sum16", "Unsigned", "14 bits", "0 bit", "[13:0], 範圍 0 ~ 16368"),
        ("去除基線 Del_DC", "Signed (2's)", "11 bits + Sign", "4 bits", "[15:0], 範圍 -16368 ~ +16368 (scaled x16)"),
        ("視窗加權和 Conv_WF", "Signed (2's)", "17 bits + Sign", "4 bits", "[21:0], 完整保留無截斷，用於極值判定")
    ]
    
    for row_idx, row in enumerate(fix_table.rows):
        for col_idx, cell in enumerate(row.cells):
            txt = fix_data[row_idx][col_idx]
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(txt)
            is_header = (row_idx == 0)
            format_run(run, font_name="微軟正黑體", font_size=8.5, bold=is_header, color_rgb=(255,255,255) if is_header else (0,0,0))
            set_cell_margins(cell, top=35, bottom=35, left=50, right=50)
            if is_header:
                set_cell_background(cell, "2F5597")
            elif row_idx % 2 == 1:
                set_cell_background(cell, "F2F5F9")

    p_cap_tab3 = doc.add_paragraph()
    p_cap_tab3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap_tab3.paragraph_format.space_after = Pt(4)
    rc3 = p_cap_tab3.add_run("表三、 EDE 演算法各級運算位元寬度與無損定點數規格配置表")
    format_run(rc3, font_name="微軟正黑體", font_size=8.5, bold=True, color_rgb=(90,90,90))

    add_styled_heading(doc, "3.4 步驟 4：Q、R、S 特徵點定位規則與平手仲裁機制", level=2)
    add_body_paragraph(doc,
        "每 120 筆取樣構成一個 ECG Beat（共 20 個 Beats）。特徵定位規則嚴格規範如下：\n"
        "1. R 點定位 (Global Max)：在該 Beat 120 點區間內搜尋 WF(n) 全域最大值，其對應座標即為 R 點。\n"
        "2. Q 點定位 (Local Min 往前搜尋)：自 R 點座標向左（往前）回溯搜尋，當遇到局部反曲波谷點（WF(n-1) > WF(n)）時鎖定為 Q 點。\n"
        "3. S 點定位 (Local Min 往後搜尋)：自 R 點座標向右（往後）延展搜尋，當遇到局部反曲波谷點（WF(n+1) > WF(n)）時鎖定為 S 點。\n"
        "4. 平手仲裁規則 (Tie-Breaking Rule)：若搜尋時遭遇平坦波谷（連續多點數值相同），系統嚴格依題意選取「距離 R 點最近之點」作為最終輸出座標。"
    )

    # 圖三：QRS 搜尋波形圖
    add_centered_image(doc, r"c:\Users\milto\Desktop\ic期末報告\fig_p6_6.png", "圖三、 心電圖 Beat 區間內全域波峰 (R) 與局部波谷 (Q、S) 定位搜尋機制示意圖", width=Inches(3.5))

    # ==================== 四、 電路微架構設計 ====================
    add_styled_heading(doc, "四、 電路微架構設計 (Hardware Microarchitecture)", level=1)
    
    add_body_paragraph(doc,
        "為最佳化期末報告評分之 Area × Time 指標，本設計揚棄傳統儲存 2400 筆之龐大靜態 RAM 架構，創新提出「串流管線 (Streaming Pipeline) 結合雙緩衝 (Ping-Pong Buffer) 延遲隱藏機制」："
    )
    
    add_styled_heading(doc, "4.1 串流濾波前級 (Streaming Filtering Pipeline)", level=2)
    add_body_paragraph(doc,
        "• 16-tap 線性移位暫存器：以 16 級 10-bit 暫存器構成滑動視窗，n < 15 時由多工器載入 x(0) 自動完成邊界填補。\n"
        "• 15-tap 視窗濾波暫存器：WF(n) 需未來取樣 Del_DC(n+7) 參與卷積，僅需 15 級暫存器做中心點延遲，管線延遲僅 7 個 cycles。管線填滿後，每個 cycle 即可穩定流式產出一筆無損 WF(n) 數據。"
    )

    add_styled_heading(doc, "4.2 雙緩衝搜尋引擎與延遲隱藏 (Ping-Pong Search & Latency Hiding)", level=2)
    add_body_paragraph(doc,
        "• 雙緩衝區輪替 (Bank A & Bank B，深度僅 120 words)：當當前 Beat 之 WF 資料流寫入 Bank A 時，比較器即時比對並更新 Global Max 暫存器，在第 120 點寫入瞬間，R 點已底定！\n"
        "• 平行搜尋隱藏延遲 (Latency Hiding)：當下一 Beat 轉向寫入 Bank B 時，專用搜尋狀態機立即啟動對 Bank A 進行 Q 點與 S 點的雙向搜尋。搜尋 Q、S 點至多僅需 30~50 cycles，遠小於下一 Beat 寫入所需的 120 cycles，找點運算得以「完全隱藏於資料輸入時間內」！\n"
        "此架構優勢：使電路之 busy 旗標常態保持為 Low，總模擬週期數逼近理論極限值 2400 cycles，達成最高吞吐量並大幅節省晶片面積。"
    )

    # ==================== 五、 專題時程規劃與里程碑 ====================
    add_styled_heading(doc, "五、 專題時程規劃與里程碑 (Project Schedule & Milestones)", level=1)
    
    plan_table = doc.add_table(rows=6, cols=3)
    plan_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(plan_table, color="4F81BD", sz="4", val="single")
    
    plan_data = [
        ("預計週次與時程", "工作重點與里程碑項目", "預計交付與達成目標"),
        ("第 1 ~ 2 週 (11/03 - 11/16)", "演算法模型建立、定點化無損分析與邊界規則確認", "完成 Python 仿真驗證模型，確認 100% 命中 Golden 答案"),
        ("第 3 ~ 4 週 (11/17 - 11/30)", "Verilog RTL 模組撰寫、串流管線與雙緩衝狀態機整合", "通過 VCS 前模擬 (PAT1 & PAT2)，確認功能與時序完全正確"),
        ("第 5 週 (12/01 - 12/07)", "Synopsys Design Compiler 邏輯合成與時序面積最佳化", "通過 Pre-layout Gate-level 模擬，產出 Netlist 與 SDF 檔"),
        ("第 6 ~ 7 週 (12/08 - 12/21)", "Cadence Innovus / ICC2 自動佈局與繞線 (APR)", "完成 Power Ring/Stripe/Rail/Filler 配置，產出實體 GDSII"),
        ("第 8 週 (12/22 - 12/28)", "Mentor Calibre DRC/LVS 驗證、後模擬與結案總報告書撰寫", "達成 DRC=0 / LVS=0，通過 Post-Sim，於 12/28 期限前完成繳交")
    ]
    
    for row_idx, row in enumerate(plan_table.rows):
        for col_idx, cell in enumerate(row.cells):
            txt = plan_data[row_idx][col_idx]
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if col_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(txt)
            is_header = (row_idx == 0)
            format_run(run, font_name="微軟正黑體", font_size=8.5, bold=is_header, color_rgb=(255,255,255) if is_header else (0,0,0))
            set_cell_margins(cell, top=40, bottom=40, left=60, right=60)
            if is_header:
                set_cell_background(cell, "2F5597")
            elif row_idx % 2 == 1:
                set_cell_background(cell, "F2F5F9")

    p_cap_tab4 = doc.add_paragraph()
    p_cap_tab4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap_tab4.paragraph_format.space_after = Pt(6)
    rc4 = p_cap_tab4.add_run("表四、 數位積體電路設計期末專題預計執行時程規劃表")
    format_run(rc4, font_name="微軟正黑體", font_size=8.5, bold=True, color_rgb=(90,90,90))

    # ==================== 六、 參考文獻 ====================
    add_styled_heading(doc, "六、 參考文獻 (References)", level=1)
    
    refs = [
        "[1] 財團法人國家實驗研究院台灣半導體研究中心 (TSRI), \"2022 IC Design Contest Cell-Based IC Design Category for Graduate Level: ECG Detection Engine (EDE),\" 教育部110學年度全國大學校院積體電路設計競賽決賽題目, 2022.",
        "[2] J. Pan and W. J. Tompkins, \"A real-time QRS detection algorithm,\" IEEE Transactions on Biomedical Engineering, vol. BME-32, no. 3, pp. 230-236, Mar. 1985.",
        "[3] H. C. Chen and S. W. Chen, \"A moving average based filtering and peak detection method for wearable ECG devices,\" IEEE Transactions on Biomedical Circuits and Systems, vol. 12, no. 4, pp. 812-821, Aug. 2018.",
        "[4] V. X. Chouhan and S. S. Mehta, \"Total removal of baseline wander and power line interference from ECG signal using moving average filter,\" International Journal of Computer Applications, vol. 2, no. 5, pp. 24-29, May 2010.",
        "[5] J. M. Rabaey, A. Chandrakasan, and B. Nikolic, Digital Integrated Circuits: A Design Perspective, 2nd ed., Prentice Hall, Upper Saddle River, NJ, 2003."
    ]
    
    for r in refs:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.15
        run = p.add_run(r)
        format_run(run, font_name="微軟正黑體", font_size=8.5, color_rgb=(60, 60, 60))

    # 存檔
    out_docx_path = r"c:\Users\milto\Desktop\ic期末報告\期末報告構想說明書_陳沐德.docx"
    doc.save(out_docx_path)
    print(f"Successfully generated docx at: {out_docx_path}")

if __name__ == "__main__":
    main()
