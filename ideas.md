# SeaRates Excel 網址產生器設計提案

## Approach 1
**Theme Name:** Chartroom Ledger

**Very Brief Intro:** 以航海圖室與資料台帳為靈感，使用深墨藍、紙張米白與航標橙，讓 Excel 轉換流程像一份可靠的航線作業單。介面強調方向、狀態與可追溯性。

**Probability:** 0.03

## Approach 2
**Theme Name:** Quiet Freight Studio

**Very Brief Intro:** 以北歐運輸工具與編輯型工作台為基礎，採霧灰、海玻璃綠與暖白，讓大量欄位資訊保持輕盈、安靜且易掃讀。互動像整理一張清楚的作業桌。

**Probability:** 0.07

## Approach 3
**Theme Name:** Signal Current

**Very Brief Intro:** 以現代航運訊號與雷達介面為靈感，使用午夜藍、青綠訊號色與琥珀警示色，呈現清晰的批次處理節奏。視覺重點放在匯入、匹配、輸出三個關鍵狀態。

**Probability:** 0.02

# Selected Approach: Chartroom Ledger

## Design Movement
現代航海工業設計（nautical industrial editorial）：把航海圖室的工具感、紙張台帳的秩序與數位工作台的即時回饋融合在一起。

## Core Principles
1. **資料先於裝飾：** 每個視覺元素都要幫助使用者確認欄位、狀態或下一步。
2. **航線即結構：** 使用細線、節點、分段與方向箭頭暗示資料從 Excel 流向 SeaRates。
3. **可追溯輸出：** 上傳檔名、列數、成功數、待補數與輸出檔名要始終可見。
4. **穩定而不僵硬：** 用米白底與深藍字維持專業，以航標橙作為少量但明確的行動訊號。

## Color Philosophy
主色深墨海軍藍承擔標題、導覽與主要文字，象徵可信與方向感；霧白與暖米色模擬紙張與工作台，降低長時間操作的疲勞；航標橙只用在主要動作、目前步驟與警示，讓使用者一眼辨認「可以執行」與「需要處理」。

## Layout Paradigm
採左側固定「航線作業欄」與右側寬廣「資料預覽區」的非對稱工作台。左側負責上傳、代碼表狀態與匯出動作；右側展示批次統計、欄位預覽與網址結果。桌面版保持雙欄，窄螢幕改為上下流程。

## Signature Elements
1. 左側垂直航線節點線：Upload → Match → Export。
2. 紙張卡片上的細微等高線／航圖紋理，不使用大面積漸層。
3. 航標橙的短線與小圓點，用於完成狀態與主要 CTA。

## Interaction Philosophy
互動必須回應使用者的工作進度。檔案拖入時顯示「已靠岸」狀態；欄位匹配完成後顯示清楚的成功／待補標籤；匯出按鈕只在有有效資料時突出。每個動作都有短暫、清楚且不打斷工作的回饋。

## Animation
頁面初次載入以 30–60ms 的階梯式淡入呈現統計卡與預覽表；拖曳檔案時只動畫邊框與透明度；成功匯出時航線節點由左至右亮起，持續不超過 260ms。所有動效只使用 opacity 與 transform，並尊重 prefers-reduced-motion。

## Typography System
標題使用 Fraunces 斜體或半斜體，帶出航海日誌的編輯感；內文與表格使用 IBM Plex Sans，確保數字與英文 URL 清楚可讀。H1 42/48、H2 24/30、正文 14/22、表格 12/18；URL 使用等寬字體 IBM Plex Mono。

## Brand Essence
給需要把報單 Excel 快速整理成可追溯 SeaRates 連結的物流與報關人員，一個不需手動複製貼上的航線資料工作台；它把中文地點、英文代碼、運輸方式與輸出檔案一次對齊。

**Personality:** Reliable, precise, composed.

## Brand Voice
標題要像作業標示，簡短而有方向；CTA 使用動詞與結果，不用空泛的「開始使用」。微文案直接說明目前狀態與可處理的資料。

**Example lines:**
- 「把報單帶上船，連結一次整理。」
- 「38 列已對上航線規則，準備匯出。」

## Wordmark & Logo
Logo 使用一個由兩條交會航線與一個方形資料節點組成的幾何符號，不放文字；符號可在側欄、favicon 與匯出狀態中獨立辨識。

## Signature Brand Color
**Beacon Orange — #E36B3D**：像遠距離可辨識的港口航標，僅用於主要行動與需要注意的狀態。

## Style Decisions
- 網頁主體使用「Chartroom Ledger」；不採用紫色漸層、過度圓角或全版置中的 SaaS 模板。
- 介面要同時呈現航線方向與 Excel 批次進度。

## Style Decisions

- 維持使用者指定的簡潔藍白配色；深藍負責結構與文字，航標橙 #E36B3D 只用於目前步驟、主要匯出動作與待處理狀態。
- 以極低對比的紙張網格作為背景，不加入大幅度裝飾，避免影響 Excel 表格閱讀。
- 預覽卡上使用短橙色文件線，作為「結果已整理、可匯出」的視覺訊號；空資料時匯出按鈕保持退後的灰藍狀態。
- 主要文案以繁體中文操作標示為主，英文僅作為次要工作台標籤。
