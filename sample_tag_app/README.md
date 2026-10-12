# Sample Tag 產生器（D240 / D53）

上傳 PCN → 勾選品項 → 一鍵產出三樣東西：

| 產出 | 說明 |
|---|---|
| **Sample Tag Excel (.xlsx)** | 以 Target 範本的卡片版型產生，每品項一張卡；條碼、縮圖都是**內嵌圖片**（不用裝字型、不用連網、沒有巨集）。IMU 公式保留，可在 Excel 裡改成本／零售價即時重算。另附 `Data` 工作表列出全部品項。 |
| **卡片 PNG（ZIP）** | 每張卡一張圖，檔名 `001_HWD704A.png`。取代原本的 `BatchExportScreenshots_Safe` 截圖巨集：只畫有資料的卡、不會白圖、不需要開 Excel。 |
| **列印 PDF** | A4 直式、每頁 8 張（與 Target 範本列印版面相同）。 |

另有 **Dropdown 更新** 分頁：把 Target 新一季範本的 Dropdowns（Program、IMU 目標、廠商清單）更新進你的工作檔，其他工作表、巨集、格式不動。

---

## 部署 / Deploy

純 Python，不需要任何系統套件（字型已內附在 `fonts/`）。

**本機**
```bash
pip install -r requirements.txt
streamlit run app.py
```

**Streamlit Community Cloud**
1. 把整個資料夾推到 GitHub（private repo）。
2. 在 Streamlit Cloud 建 app，Main file 選 `app.py`。
3. App settings → Secrets 貼上：
   ```toml
   APP_PASSWORD = "團隊密碼"
   ```
   沒設密碼時 App 不會要求登入（方便本機測試）。

---

## 使用流程 / Workflow

### D240
1. 上傳 PCN（例：`2027 C5 D240 Halloween Decor & Night Of.xlsx`）。
2. Program Name 自動依 PCN 的 **Subclass ID** 帶出（`Subclass 27`…），每張卡用自己 Subclass 的 **IMU+Duty Goal** 判斷紅黃綠燈（畫面可切換成 IMU Goal）；也可在上方改成全部同一個 Program。
3. 選填：在「② 產品圖片」放圖（見下方「產品圖片」），帶入卡片左下角的 THUMBNAIL IMAGE。
4. 在表格中勾選要產生的品項，可填 Sample Date、LY DPCI / LY Cost / LY Retail / LY Duty、Comments。
5. 按「產生 N 張 Sample Tag」→ 下載 Excel / PNG / PDF。

### D53
同上，差別只在 Program Name 由 **PCN 檔名**帶出（`… - Boo Basket & Wrap.xlsx` → `Boo Basket & Wrap`），帶不出來時會提醒你手動選。

### 產品圖片 / Product photos

「② 產品圖片」一個上傳框就能放三種來源，可混著放：

| 來源 | 適用時機 | 對應方式 |
|---|---|---|
| **圖檔（可多選）** | 初選階段，手上只有幾張照片 | 依檔名 |
| **ZIP** | 一次很多張（資料夾直接壓縮即可，子資料夾也讀） | 依檔名 |
| **SPK Excel** | 已有 SPK 之後 | 依 DPCI → Spark PID → MFG Style → PCN Item ID |

檔名規則（不分大小寫）：
- 檔名**就是**鍵值：`HWD704A.jpg`、`240-43-5528.png`、`8351342.jpg`
- 檔名**包含**鍵值：`001_HWD704A_front.jpg`、`IMG HWD704A (2).png`（整個字比對，`HWD704` 不會誤配到 `HWD704A`）
- 可用的鍵值：MFG Style #、DPCI、PCN Item ID、Spark PID、Apex ID

對不到的圖片會列在「檔名對不到品項的圖片」，每張旁邊下拉選單選品項即可手動指定。
優先順序：手動指定 ＞ 圖檔／ZIP ＞ SPK。同一品項有多張圖取檔名排序第一張。
所有圖片（圖檔、ZIP、SPK）一律自動裁掉四周白邊／透明邊、轉正手機照片方向，並縮到 700px 以內（Excel 不會變太大）。

### 每季 / Quarterly：Dropdown 更新
1. 「Dropdown 更新」分頁上傳 TEMPLATE（你的工作檔）與 SOURCE（Target 新一季範本）。
2. 檢查差異（新增／移除 Program、IMU 目標變動、廠商、黃燈區間）。
3. 下載更新後的 TEMPLATE；或按「用這份 SOURCE 產生本次 Sample Tag」直接拿新範本產卡。
4. 要讓新範本變成 App 的**預設**：把檔案放進 `templates/`，檔名維持 `D240_Sample_Tag_Template.xlsx` / `D53_Sample_Tag_Template.xlsx`，重新部署即可。

---

## PCN 欄位對應 / Field mapping

欄位一律用**表頭名稱**比對（忽略大小寫與結尾的 `*`），所以 D240、D53 欄位位置不同也能讀。

| 卡片欄位 | PCN 表頭 |
|---|---|
| MFG STYLE # | Manufacturer Style # |
| PRODUCT DESCRIPTION | Vendor Product Description（固定只抓這欄；空白時卡片留白並標黃，表格提醒「缺 Vendor Product Description」） |
| PCN ITEM ID | Pcn Item Id |
| APEX ID | Apex Id |
| SPARK PID | Spark PID |
| FIRST COST | FCA Factory City Unit Cost（空白時改用 FOB Unit Cost，表格會標「成本取 FOB」） |
| PCN RETAIL | Suggested Unit Retail |
| COUNTRY OF EXPORT | Import Country of Origin |
| DUTY RATE | Duty Rate % |
| Program（D240） | Subclass ID → `Subclass N` |

---

## 與原 Excel 版的差異 / What changed vs the macro workbook

- **條碼**：Code 128 圖片（python-barcode 產生編碼）；開發測試時以掃碼解碼器驗證 D240 362 張、D53 19 張全部讀回正確。不再需要 `Code128B`、`code128.ttf` 或 `=IMAGE()`。
- **IMU 公式**加了防呆：成本或零售價空白時顯示 `FIELD WILL AUTO CALC`，不會再出現假的 100% IMU。
- **IMU 紅黃綠燈**改用標準條件式格式、每張卡填自己的目標值（D240 同一份會混多個 Subclass）。Excel、Mac Numbers、Google Sheets 顯示一致。
- **廠商色條**改成固定底色（顏色依範本規則算出，Lucky Star = 洋紅）。
- **品名**自動換行；**Sample Date** 顯示 `yyyy/mm/dd`，且是產生當下的固定日期。
- 所有卡片的框線、合併格都由第一張卡複製，不會再有粗細不一。

---

## 新增 Program / Add a program

在 `core/config.py` 的 `PROGRAMS` 加一筆（範本檔名、Program Name 來源 `subclass` 或 `filename`、預設黃燈區間），並把範本放進 `templates/`。卡片引擎不用改。

## 檔案結構

```
app.py                 Streamlit 介面
core/pcn.py            讀 PCN、Program 對應、資料檢查
core/dropdowns.py      讀 Dropdowns、比對差異、Dropdown 更新（XML surgery）
core/builder.py        產生 Sample Tag Excel（XML surgery，保留範本格式）
core/render.py         卡片 PNG / PDF
core/barcode_img.py    Code 128 條碼圖
core/spk.py            SPK 縮圖擷取
core/images.py         圖檔／ZIP 讀取、裁白邊、依檔名對應
templates/             內建 Target 範本（D240 C5 2026、D53 C5 2027）
fonts/                 Carlito（Calibri 等寬替代，OFL 授權）
```
