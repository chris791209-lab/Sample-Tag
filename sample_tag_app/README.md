# Sample Tag 生成器（D240 / D53）

上傳 PCN ➜ 勾選品項／篩選 Subclass ➜ 一鍵產出三樣東西：

| 產出 | 說明 |
|---|---|
| **Sample Tag Excel (.xlsx)** | 以內建 Target 範本的 Sample Tag 版型生成，每品項一張 Sample Tag，忠實呈現 PCN 的數值；條碼、縮圖都是**內嵌圖片**（不用裝字型、不用連網、沒有巨集）。IMU 公式保留，可在 Excel 裡改成本／零售價即時重算。另附 `Data` 工作表列出全部品項（含該季 IMU GOAL）。 |
| **Sample Tag PNG（ZIP）** | 每張 Sample Tag 一張圖，檔名 `001_HWD704A.png`。取代原本的 `BatchExportScreenshots_Safe` 截圖巨集：只畫有資料的 Sample Tag、不會白圖、不需要開 Excel。 |
| **列印 PDF** | A4 直式、每頁 8 張（與 Target 範本列印版面相同）。 |

每季生成 Sample Tag 的第一步是放入 Target 新一季範本：只取用它的 Dropdowns（該季 IMU Goal，由 Target Sourcing Manager 訂定），Sample Tag 版面、公式一律以內建範本為準。Tag 上的 IMU 欄位依該季 IMU+Duty Goal 上紅黃綠底色；App 畫面不顯示燈號，保持簡潔。

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

D240、D53 分頁操作完全相同：

1. **① Target 新一季範本（更新 IMU Goal）**：放入 Target Sourcing Manager 給的新一季 Sample Tag 範本，App 只取用其中的 Dropdowns（該季 IMU Goal），用於 Tag 上 IMU 的紅黃綠底色，並寫進產出 Excel 的 `Dropdowns` 工作表與 `Data` 工作表的 IMU GOAL 欄。沒放時沿用內建範本的 Dropdowns。
2. **② PCN**（例：`2027 C5 D240 Halloween Decor & Night Of.xlsx`）。
   - Sample Tag 上的 **Subclass** 一律以 PCN 的 Subclass ID 為準（Sample Tag 印 `Subclass 27`，畫面表格與篩選只顯示數字 `27`）。
   - **PCN 沒填 Subclass 的品項**：「待確認」欄顯示「缺 Subclass（可在 Subclass 欄填入）」，直接在 Subclass 欄輸入數字（例如 `27`）。PCN 有值的品項以 PCN 為準，不能改。
   - 畫面會提示「有 N 個品項 PCN 沒有 Subclass，請手動填入，系統將自動抓取該季 IMU」。
3. 選填 **③ 產品圖片**（見下方「產品圖片」），帶入 Sample Tag 左下角的 THUMBNAIL IMAGE。
4. 在表格中勾選要生成的品項（可用「篩選 Subclass」／搜尋縮小範圍；**只會生成目前篩選結果中有勾選的品項**，畫面上看到什麼就生成什麼），可填 Sample Date、LY DPCI / LY Cost / LY Retail / LY Duty、Comments。
5. 按「生成 N 張 Sample Tag」→ 下載 Excel / PNG / PDF。

> 目前內建的 D53 範本 Dropdowns 只有 Program 名稱、沒有 Subclass 目標，D53 需放入 Target 新一季範本，Tag 上的 IMU 才會上底色。

### 產品圖片 / Product photos

「③ 產品圖片」一個上傳框就能放三種來源，可混著放：

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

### 每季 / Quarterly
- 每季新 Program：在 ① 放入 Target 新一季範本即可（只在這次使用）。
- 要讓新範本變成 App 的**預設**：把檔案放進 `templates/`，檔名維持 `D240_Sample_Tag_Template.xlsx` / `D53_Sample_Tag_Template.xlsx`，重新部署即可。
- 只需要把新 Dropdowns 換進既有工作檔（不生成 Sample Tag）時，用「跑 Dropdown 更新」流程處理，不在 App 內。

---

## PCN 欄位對應 / Field mapping

欄位一律用**表頭名稱**比對（忽略大小寫與結尾的 `*`），所以 D240、D53 欄位位置不同也能讀。

| Sample Tag 欄位 | PCN 表頭 |
|---|---|
| MFG STYLE # | Manufacturer Style # |
| PRODUCT DESCRIPTION | Vendor Product Description（固定只抓這欄；空白時 Sample Tag 留白並標黃，「待確認」欄顯示「缺 Vendor Product Description」） |
| PCN ITEM ID | Pcn Item Id |
| APEX ID | Apex Id |
| SPARK PID | Spark PID |
| FIRST COST | FCA Factory City Unit Cost（空白時改用 FOB Unit Cost，表格會標「成本取 FOB」） |
| PCN RETAIL | Suggested Unit Retail |
| COUNTRY OF EXPORT | Import Country of Origin |
| DUTY RATE | Duty Rate % |
| Subclass（Sample Tag 第 2 行） | Subclass ID → `Subclass N`（空白時可在表格手動填） |

---

## 與原 Excel 版的差異 / What changed vs the macro workbook

- **條碼**：Code 128 圖片（python-barcode 生成編碼）；開發測試時以掃碼解碼器驗證 D240 362 張、D53 19 張全部讀回正確。不再需要 `Code128B`、`code128.ttf` 或 `=IMAGE()`。
- **IMU 公式**加了防呆：成本或零售價空白時顯示 `FIELD WILL AUTO CALC`，不會再出現假的 100% IMU。
- **IMU 紅黃綠底色**改用標準條件式格式、每張 Sample Tag 填自己 Subclass 的該季目標（D240 同一份會混多個 Subclass）。Excel、Mac Numbers、Google Sheets 顯示一致。App 畫面上不顯示燈號。
- **廠商色條**改成固定底色（顏色依範本規則算出，Lucky Star = 洋紅）。
- **品名**自動換行；**Sample Date** 顯示 `yyyy/mm/dd`，且是生成當下的固定日期。
- 所有 Sample Tag 的框線、合併格都由第一張複製，不會再有粗細不一。

---

## 新增 Program / Add a program

在 `core/config.py` 的 `PROGRAMS` 加一筆（範本檔名、IMU Goal 對應方式 `subclass`、預設黃燈區間），並把範本放進 `templates/`。Sample Tag 引擎不用改。

## 檔案結構

```
app.py                 Streamlit 介面
core/pcn.py            讀 PCN、Subclass／IMU Goal 對應、資料檢查
core/dropdowns.py      讀 Dropdowns、把新一季 Dropdowns 換進內建範本（XML surgery）
core/builder.py        生成 Sample Tag Excel（XML surgery，保留範本格式）
core/render.py         Sample Tag PNG / PDF
core/barcode_img.py    Code 128 條碼圖
core/spk.py            SPK 縮圖擷取
core/images.py         圖檔／ZIP 讀取、裁白邊、依檔名對應
templates/             內建 Target 範本（D240 C5 2026、D53 C5 2027）
fonts/                 Carlito（Calibri 等寬替代，OFL 授權）
```
