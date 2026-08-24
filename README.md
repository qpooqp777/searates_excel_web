# SeaRates Excel 網址產生器

一個以瀏覽器與 Python／CLI 為核心的 Excel 資料整理工具。它會讀取報單資料，依照港口與機場代碼對照表判斷運輸方式，產生 SeaRates 查詢網址，並可將結果匯出成新的 Excel 檔案。

本專案的主要設計原則是：**資料留在使用者的裝置內處理、代碼表可維護、網址規則透明、輸出結果容易核對**。

## 功能總覽

| 功能 | 說明 |
|---|---|
| 網頁版 Excel 匯入 | 在瀏覽器載入 `.xlsx` 或 `.xls`，不需要上傳到伺服器。 |
| 貼上文字 | 可貼上三欄資料：運輸方式、出發地英文代碼、目的地英文代碼。 |
| 海運／空運判斷 | 依資料與代碼表產生 `Sea`／`Air` 及 `seaport`／`airport` 參數。 |
| SeaRates 網址產生 | 產生包含起點、終點、運輸方式與路由模式的 Distance & Time 查詢網址。 |
| JSON 代碼表管理 | 可在網頁中搜尋、新增、修改、刪除港口／機場資料，並匯出更新後 JSON。 |
| Python 與 CLI | 可在本機以腳本或命令列批次處理 Excel。 |
| Excel 匯出 | 保留原始欄位，並加入運輸方式、英文代碼、SeaRates 查詢連結與處理狀態。 |

## 操作流程

```mermaid
flowchart LR
    A[Excel 或三欄文字] --> B[讀取資料]
    B --> C[載入 JSON 對照表]
    C --> D{海運或空運}
    D -->|海運| E[seaport 參數]
    D -->|空運| F[airport 參數]
    E --> G[產生 SeaRates URL]
    F --> G
    G --> H[預覽結果]
    H --> I[匯出 Excel]
```

## 網頁版

網頁版提供兩種輸入方式：

1. **上傳 Excel**：適用於已有完整欄位名稱與多欄報單資料的工作流程。
2. **貼上文字**：不需要欄位名稱，每一行只需三欄，欄位順序固定為：

   ```text
   運輸方式<TAB>出發地英文代碼<TAB>目的地英文代碼
   ```

範例：

```text
海運	Valencia,+Valencian+Community,+ES	Keelung,+TW
海運	Mundra,+Gujarat,+IN	Keelung,+TW
空運	Soekarno-Hatta+International+Airport,+Java,+ID	Taoyuan,+Taoyuan+City,+TW
```

貼上資料後，按下「讀取貼上資料」，系統會產生 SeaRates 查詢網址並顯示預覽；按「匯出更新後 Excel」即可下載結果。

> 網頁版的資料處理在瀏覽器內完成。使用者可在「代碼表管理」分頁編輯對照資料，修改內容會留在目前瀏覽器工作階段，必須按「匯出 JSON」才會下載成檔案。

![網頁版介面預覽](docs/images/home-preview.png)

## Excel 輸入欄位

Excel 上傳模式可處理下列報單欄位。原始欄位會保留，程式會在輸出檔加入標準化欄位。

| 欄位 | 用途 |
|---|---|
| 報單號碼 | 保留原始案件識別資訊。 |
| 出口港（裝貨港） | 一般海運資料的出口港名稱。 |
| 出口港位置 | 可作為起點或來源名稱，實際方向請依工作流程確認。 |
| 進口港（卸存地） | 保留原始進口資訊。 |
| 港口位置 | 目的地或卸存地；若包含機場名稱，會依規則使用空運。 |
| 預估運輸距離 | 保留既有距離資料，不由本工具重新計算。 |
| 查詢連結 | 保留原始查詢連結。 |
| 查詢畫面 | 保留原始截圖欄位。 |
| 案件日期 | 保留案件日期。 |
| 運輸項目 | 保留貨物或運輸項目描述。 |
| 總重量 | 保留重量資料。 |

輸出檔案會新增以下欄位：

| 輸出欄位 | 說明 |
|---|---|
| 運輸方式 | `海運` 或 `空運`。 |
| 出發地英文代碼 | 從 JSON 對照表取得的 SeaRates `from` 參數。 |
| 目的地英文代碼 | 從 JSON 對照表取得的 SeaRates `to` 參數。 |
| SeaRates查詢連結 | 產生的完整 SeaRates Distance & Time URL。 |
| 查詢狀態 | `完成` 或顯示待補對照的地點。 |

## JSON 代碼表

`searates_location_codes.json` 是 Python、CLI 與網頁版共用的主要對照資料來源。每個地點可以設定顯示名稱、類型、海運參數、空運參數、別名及代碼。

```json
{
  "schema_version": "1.0",
  "locations": {
    "基隆港": {
      "type": "seaport",
      "sea": "Keelung, TW",
      "aliases": ["Keelung", "KELW"] ,
      "unlocode": "TWKEL"
    }
  }
}
```

欄位規則如下：

| 欄位 | 必要性 | 說明 |
|---|---|---|
| `type` | 建議 | 使用 `seaport` 或 `airport`。 |
| `sea` | 海運必要 | SeaRates 海運的英文 URL 參數。 |
| `air` | 空運必要 | SeaRates 空運的英文 URL 參數。 |
| `aliases` | 建議 | Excel 中可能出現的中文、英文或縮寫名稱。 |
| `unlocode` | 港口建議 | 港口 UN/LOCODE，例如 `TWKEL`。 |
| `iata` | 機場建議 | 機場 IATA 代碼，例如 `TPE`。 |

## Python 版本

Python 版本適合整批處理 Excel。環境需要 Python 3.10 以上，以及 `pandas`、`openpyxl` 和 `xlsxwriter` 等套件。

```bash
pip install pandas openpyxl xlsxwriter
python3 searates_excel_processor.py \
  --input "輸入.xlsx" \
  --output "輸出_SeaRates.xlsx" \
  --codes "searates_location_codes.json"
```

程式會讀取指定輸入檔案，使用 JSON 對照表解析地點，產生海運／空運與 SeaRates URL，最後寫入指定輸出檔案。若 JSON 找不到某個地點，輸出中的「查詢狀態」會標示待補對照，不會默默產生錯誤網址。

## CLI 版本

CLI 版本使用同一套核心處理邏輯，適合整合到批次檔、排程或其他本機工作流程。

```bash
python3 searates_cli.py \
  "輸入.xlsx" \
  --output "輸出_SeaRates.xlsx" \
  --codes "searates_location_codes.json"
```

查看參數說明：

```bash
python3 searates_cli.py --help
```

## SeaRates URL 規則

海運網址使用下列參數：

```text
transportMode=Sea
fromPlaceType=seaport
toPlaceType=seaport
```

空運網址使用下列參數：

```text
transportMode=Air
fromPlaceType=airport
toPlaceType=airport
```

兩種網址都會使用：

```text
routingMode=short
```

本工具只負責產生查詢網址，不會自動繞過 SeaRates 登入、驗證、查詢次數限制或其他網站存取控制。實際距離與航程仍以 SeaRates 開啟網址後顯示的內容為準。

## 專案結構

```text
.
├── client/                         # React 網頁版
├── docs/images/home-preview.png    # README 介面預覽圖
├── searates_location_codes.json    # 港口／機場 JSON 對照表
├── searates_excel_processor.py     # Python 核心處理腳本
├── searates_cli.py                 # CLI 入口
├── SeaRates_Excel_腳本使用說明.md  # 腳本補充說明
└── README.md                       # 本文件
```

## 注意事項

本專案產生的英文地點字串必須符合 SeaRates 當下可識別的地點參數；若 SeaRates 調整地點名稱或查詢頁面規則，應優先更新 JSON 對照表，再重新產生 Excel。對於同一路線的多筆資料，可以共用相同的 SeaRates URL，但每筆案件仍應保留自己的報單號碼與原始欄位。

請避免將含有個人資料、商業機密、完整報單內容或其他敏感資料的 Excel 檔案提交到公開 GitHub 儲存庫。建議使用 `.gitignore` 排除本機輸入檔、輸出檔、截圖與暫存資料。

## 授權與貢獻

本專案目前未指定正式開源授權。若要讓其他人合法重製、修改與散布，請由專案擁有者選擇並加入適合的 LICENSE，例如 MIT License。

歡迎透過 GitHub Issues 提交代碼表修正、網址規則問題或介面改善建議；提交前請先確認沒有包含真實報單資料與敏感資訊。
