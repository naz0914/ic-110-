# CAD/VLSI Circuit Design 期末專題：心電圖偵測引擎 (ECG Detection Engine, EDE)

- **學生姓名**：陳沐德
- **所屬系所**：國立成功大學 醫學工程研究所
- **指導老師**：程德勝 教授
- **專題題目**：教育部 110 學年度全國大學校院積體電路設計競賽 (2022 IC Design Contest) 研究所組決賽題目 —— **ECG Detection Engine (EDE)**

---

## 專案內容簡介
本專題為 Cell-Based 數位積體電路 (ASIC) 設計，目標實作一個心電圖特徵偵測硬體加速器。
電路可即時接收 2400 筆 10-bit 取樣點之 ECG 訊號（共 20 個 Beats），經由數位濾波演算法即時去除基線漂移 (DC-bias) 與高頻雜訊，並自動偵測每個 Beat 區間內的 **Q、R、S** 關鍵特徵點座標。

### 核心運算步驟
1. **16-tap Moving Average Filter**：估計低頻基線變化，去除 DC-bias 趨勢。
2. **Remove DC-bias**：扣除移動平均值，訊號水平中心歸零。
3. **15-tap Windowing Filter**：以對稱三角形加權視窗濾除肌肉顫動與高頻雜訊。
4. **Q、R、S Detection Engine**：
   - **R 點**：Beat 區間內全域最大值 (Global Max)。
   - **Q 點**：自 R 點向左回溯搜尋局部最小值 (Local Min)。
   - **S 點**：自 R 點向右延展搜尋局部最小值 (Local Min)。

---

## 目錄架構與重要檔案
```text
├── 110年決賽競賽題目 (研究所組) (1).pdf   # 競賽官方題目與規格書
├── 期末報告說明 (1).pdf                   # 課程期末報告規範與評分要求
├── 期末報告構想說明書_陳沐德.docx         # 期末報告構想說明書 Word 檔 (已排版)
├── 期末報告構想說明書_陳沐德.pdf          # 期末報告構想說明書 PDF 檔 (可直接繳交)
├── create_proposal_doc.py                # 構想說明書生成與排版腳本
└── README.md                             # 專案說明文件
```

---

## 預計硬體目標規格
- **製程技術**：TSMC ADFP 虛擬製程 (16nm FinFET) / 輔以 TSMC 0.13 μm 成熟製程
- **目標工作時脈**：100 MHz (Clock Period = 10.0 ns)
- **硬體架構**：串流管線 (Streaming Pipeline) 結合雙緩衝 (Ping-Pong Buffer) 延遲隱藏機制
- **評分目標**：完成 RTL -> Synthesis -> APR -> DRC/LVS Clean (等級 A，追求最小 Area × Time)
