# Google Apps Script (GAS) 開發架構觀念與部署實戰手冊

本手冊由 **指導老師** 親自指導，專為 **天心同仁 (ERP 專業開發團隊)** 設計，旨在協助同仁建立 Google Apps Script (GAS) 的核心系統架構與安全權限觀念，並透過實際環境截圖，完成第一支 Web App 的極簡部署實戰。

> ⚠️ **商業去識別化宣告**：本手冊在網頁範例與截圖中所使用的示範名稱統一為 **「銀行軟體台西分公司」**。
> **版權聲明**：© 2026 **Falo x Force Cheng** (發布於 2026/6/10)。本內容僅供**教學與內部培訓用途**使用。

---

## 目錄
1. [GAS 兩大執行模式：容器綁定型 vs 獨立型](#1-gas-兩大執行模式容器綁定型-vs-獨立型)
2. [安全密鑰管理：試算表 (易管理) vs 指令碼屬性 (安全)](#2-安全密鑰管理試算表-易管理-vs-指令碼屬性-安全)
3. [資料正規化初始化：使用 setup.gs](#3-資料正規化初始化使用-setupgs)
4. [前後端分離思維：GAS HTML vs GitHub HTML + GAS Proxy](#4-前後端分離思維gas-html-vs-github-html--gas-proxy)
5. [實戰演練一：極簡 Web App 部署步驟 (含圖文引導)](#5-實戰演練一極簡-web-app-部署步驟-含圖文引導)
6. [實戰演練二：雙表帳密登入與問答系統 (V2 - GAS 內建網頁版)](#6-實戰演練二雙表帳密登入與問答系統-v2---gas-內建網頁版)
7. [實戰演練三：前後端分離與環境變數安全防護 (V3 - 地端/GitHub HTML + GAS Proxy API)](#7-實戰演練三前後端分離與環境變數安全防護-v3---地端github-html--gas-proxy-api)
8. [常見問題與 AI 互動排錯](#8-常見問題與-ai-互動排錯)

---

## 1. GAS 兩大執行模式：容器綁定型 vs 獨立型

同仁在進行 ERP 系統整合與資料存取時，必須先清楚 GAS 專案的兩種存在形式，這直接決定了程式碼開發入口與存取權限的控制：

* **容器綁定型腳本 (Container-bound Scripts)**：
  * **概念**：類似 Excel VBA。程式碼與特定的 Google 試算表、文件或簡報深度綁定。
  * **特點**：
    * 開啟入口：在 Google 試算表中，點擊頂部選單 **「擴充功能」 > 「Apps Script」**。
    * 當前容器取得：可直接呼叫 `SpreadsheetApp.getActiveSpreadsheet()` 取得當前綁定的試算表。
    * 檔案管理：腳本不會單獨顯示在雲端硬碟中，會隨著試算表複製或刪除而一同複製或刪除。
* **獨立型腳本 (Standalone Scripts)**：
  * **概念**：類似獨立運行的後端 Server。它在雲端硬碟中是一個獨立的檔案。
  * **特點**：
    * 開啟入口：在 Google 雲端硬碟中點選 **「新增」 > 「更多」 > 「Google Apps Script」**。
    * 當前容器取得：必須透過特定的 ID 才能操作試算表，例如：`SpreadsheetApp.openById("試算表_ID")`。
    * 檔案管理：檔案獨立存在，有獨立的版本管理與分享權限，可供多個不同的試算表共用。

---

## 2. 安全密鑰管理：試算表 (易管理) vs 指令碼屬性 (安全)

在 ERP 系統對接中，經常需要串接外部 API（例如 Line Bot API、ERP API 或資料庫金鑰）。我們在管理敏感密碼與參數時，應遵循以下安全原則：

* **非敏感參數 (一般對照表、常數設定)**：
  * **做法**：直接放在 Google Sheets 的某個 `Config` 分頁中。
  * **優點**：易於管理。即使是不懂技術的同仁，也能直接在試算表上調整設定值（如：警報通知人數、資料查詢起訖日），無需修改程式碼。
* **敏感憑證 (API Key、私鑰、重要系統密碼)**：
  * **做法**：**禁止**寫在程式碼中，也**不宜**直接曝露在 Sheets 上。應放入 GAS 專案設定的 **「指令碼屬性 (Script Properties)」** 中。
  * **讀取範例**：
    ```javascript
    // 在程式碼中動態讀取敏感金鑰，避免程式碼外洩時洩漏憑證
    var apiKey = PropertiesService.getScriptProperties().getProperty("ERP_API_KEY");
    ```
  * **設定方法**：在 GAS 編輯器左側點擊 **「專案設定 (齒輪圖標)」** > 下拉至 **「指令碼屬性」** > 點擊 **「新增指令碼屬性」**，將金鑰以 Key-Value 方式儲存於 Google 雲端後台。

---

## 3. 資料正規化初始化：使用 `setup.gs`

在多人協作或將工具部署給其他部門使用時，最常遇到的問題是：使用者複製了您的試算表，但忘記建立對應的 Sheet 分頁（例如 `Logs` 或 `Settings`），或是手動輸入的分頁名稱有空格、錯字，導致程式執行崩潰。

* **解決方案**：在專案中新增一個 `setup.gs` 檔案，專門撰寫環境與資料正規化的初始化函數：
  ```javascript
  function setup() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. 自動檢查並建立 Logs 分頁
    var logSheet = ss.getSheetByName("Logs");
    if (!logSheet) {
      logSheet = ss.insertSheet("Logs");
      // 自動寫入標準欄位標頭 (資料正規化)
      logSheet.appendRow(["時間戳記", "操作人員", "執行動作", "狀態"]);
      logSheet.getRange("A1:D1").setFontWeight("bold").setBackground("#d9ead3");
    }
    
    // 2. 初始化指令碼屬性 (如果尚未設定)
    var scriptProperties = PropertiesService.getScriptProperties();
    if (!scriptProperties.getProperty("SYSTEM_VERSION")) {
      scriptProperties.setProperty("SYSTEM_VERSION", "v1.0.0");
    }
    
    Logger.log("系統初始化與資料正規化設定完成！");
  }
  ```
* **好處**：同仁複製新試算表後，只需要手動執行一次 `setup` 函數，程式就會自動在背景將資料結構、格式與屬性全部建立妥當，達到「開箱即用」並防止人工操作失誤。

---

## 4. 前後端分離思維：GAS HTML vs GitHub HTML + GAS Proxy

GAS 允許我們在專案中建立 HTML 檔案來產出網頁介面，但在進行較具規模的 Web 專案時，架構選擇會直接影響效能與安全性：

1. **GAS_html (內建渲染模式)**：
   * **機制**：在 GAS 內建編輯器中新增 `.html` 檔案，後端呼叫 `HtmlService.createHtmlOutputFromFile()` 來顯示網頁。
   * **缺點**：
     * GAS 的 Web App 網頁會被包裹在 Google 的安全沙盒（iframe）中，導致載入速度慢、效能較差。
     * 對於 CSS/JS 的套件引入（如 Bootstrap、Tailwind、TailwindCSS）及版本控制 (Git) 非常不便。
2. **GitHub HTML + GAS Proxy (推薦的前後端分離架構)**：
   * **機制**：
     * **前端 (Client-side)**：將精美的 HTML/CSS/JS 網頁（例如使用 Cyber Cyan 科技感風格與現代互動效果）撰寫完畢後，發布至 **GitHub Pages** 或其他靜態網站託管服務。
     * **後端 (Server-side)**：GAS 僅作為一個 **API Proxy (代理伺服器)**。前端網頁透過 `fetch()` 向 GAS 的網頁應用程式網址發送 JSON 請求，GAS 接收後負責讀寫 Google Sheet、寄信或處理邏輯，最後再將結果回傳給前端。
   * **優勢**：
     * 前端網頁載入流暢快速，沒有 Google iframe 的遲滯感。
     * 程式碼可以使用 Git 進行完善的版本控制。
     * 後端 GAS 能隱藏所有敏感的資料處理邏輯與 Google 金鑰，避免前端程式碼被瀏覽器「檢視原始碼」而曝露機密。

---

## 5. 實戰演練一：極簡 Web App 部署步驟 (含圖文引導)

現在，我們將實作第一支最簡單的 GAS Web App，目的在於「確認部署成功」並熟悉整個部署與權限設定流程。

> 🔗 **此案例線上正式 Demo 實例連結（紫色醒目框）**：[點擊前往測試](https://script.google.com/macros/s/AKfycbw9xcjQRd1qnfygic4O7fxFpk1TzjQARXBIrftl639C7PzA6n535-o6Csqc7Ji0BesZ/exec)

### 步驟 1：建立 Google 試算表
請進入雲端硬碟的專案共用資料夾（例如：`class3 > [study-gas]`），在空白處按滑鼠右鍵（或點選左上角「新增」），選擇 **「Google 試算表」** 來建立一個新表格。

![在雲端硬碟建立 Google 試算表](./images/gas_step1_create_sheet.png)

### 步驟 2：處理共用資料夾建立提示
因為這是一個共用的專案資料夾，Google 會跳出一個確認視窗，提示此檔案將沿用該資料夾的共用權限。請點選 **「建立並共用」**。

![共用資料夾建立提示](./images/gas_step2_share_dialog.png)

### 步驟 3：進入 Apps Script 編輯器
將新建好的試算表命名（例如：`gas-study-v1`），接著在頂部選單點選 **「擴充功能」 > 「Apps Script」**。

![進入 Apps Script](./images/gas_step3_apps_script_menu.png)

### 步驟 4：認識編輯器初始介面與貼入程式碼
Apps Script 編輯器會開啟一個「未命名的專案」，並在左側檔案列表中自動產生一個 `程式碼.gs`，右側編輯區則顯示一個空的 `myFunction` 函數。

![編輯器初始介面](./images/gas_step4_editor_init.png)

請將 `myFunction` 替換為以下 **「極穩健觸發授權版」** 示範程式碼（可點擊網頁版一鍵複製）：
```javascript
function doGet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = ss.getActiveSheet().getName();
  
  var html = "Deployment success. This is Bank Software Taixi Branch. Current sheet: " + sheetName;
  return HtmlService.createHtmlOutput(html);
}
```
貼上後的畫面如下，確保左側檔案名稱旁的橘色圓點亮起，表示代碼已編輯但尚未儲存，且無任何語法錯誤：

![程式碼貼上成功無語法錯誤](./images/gas_step4_editor_init_success.png)

> 💡 **教學重點提示**：
> 本程式碼採取最穩健的設計：
> 1. **無任何 Emoji 等特殊符號**：完全避免編碼與語法解析的潛在問題。
> 2. **宣告變數 html**：將輸出內容先儲存在變數中，防止代碼單行寬度過長而導致使用者複製或貼上時產生自動折行語法錯誤。
> 3. **保留 SpreadsheetApp 呼叫**：這會強制 Google 在部署時跳出安全授權視窗，方便同仁進行完整的權限授權操作。

### 步驟 5：點擊新增部署作業
在右上角的藍色選單中，點選 **「部署」** > **「新增部署作業」**。

![點選新增部署作業](./images/gas_step5_deploy_menu.png)

### 步驟 6：選取部署類型
在彈出的視窗中，點選左上角「選取類型」旁的 **齒輪圖標**，並在下拉選單中選擇 **「網頁應用程式」**。

![選取部署類型](./images/gas_step6_deploy_type.png)

### 步驟 7：配置部署參數與權限
此時右側會出現部署設定欄位：
1. **說明**：填入說明（例如：`V1.0`）。
2. **將網頁應用程式執行為 (執行身分)**：維持預設的 **「我 (您的 Email)」**。
3. **誰可以存取 (存取權限)**：預設為 **「只有我自己」**。
   * *⚠️ ERP 整合提醒*：若維持預設的「只有我自己」，其他同仁或外部系統呼叫此網址時，會被 Google 拒絕連線。

![預設權限為只有我自己](./images/gas_step7_deploy_settings_private.png)

請點開「誰可以存取」下拉選單，將其修改為 **「所有人」**（以利外部對接或他人存取）。

![修改權限為所有人](./images/gas_step8_deploy_settings_anyone.png)

### 步驟 8：安全授權流程 (首次部署必經)
設定完成並點擊「部署」後，由於我們的程式碼包含了對 `SpreadsheetApp`（試算表）的存取，Google 就會偵測到敏感權限要求，並**強制跳出安全授權視窗**。請同仁依序完成以下放行步驟：
1. 點選 **「授予存取權限 (Authorize Access)」**。

   ![點選授予存取權限](./images/gas_step8_auth_grant.png)
2. 選擇您的 **Google 帳號**。
3. 畫面出現紅色警告「Google hasn't verified this app」。請點擊左下角的 **「進階 (Advanced)」**。

   ![Google尚未驗證警告](./images/gas_step8_auth_unverified.png)
4. 點選下方小字的 **「前往『未命名專案』(Go to 未命名的專案 (unsafe))」**。

   ![確認前往未命名專案](./images/gas_step8_auth_unsafe.png)
5. 在接下來的權限確認畫面中，確認權限範圍（查看、編輯、建立和刪除您在 Google 試算表中的所有試算表），點選右下角的 **「Continue (繼續)」** 完成授權放行。

   ![Google Sheets權限要求](./images/gas_step8_auth_permissions_top.png)
   ![允許授權確認](./images/gas_step8_auth_permissions_bottom.png)

### 步驟 9：取得部署成功網址
部署與授權放行作業完成後，系統會提示「已成功更新部署作業」，並生成專屬的 **「網頁應用程式網址」**。請點擊網址下方的 **「複製」**。

![部署成功獲得網址](./images/gas_step9_deploy_success_v2.png)

### 步驟 10：無痕測試驗證
1. 開啟瀏覽器的 **「無痕視窗」**。
2. 貼上剛才複製的網頁應用程式網址並按下 Enter。
3. 若網頁成功顯示：
   **Deployment success. This is Bank Software Taixi Branch. Current sheet: 工作表1**
   
   ![無痕測試驗證成功](./images/gas_step10_verification_success.png)
   
   即代表您第一版的極簡部署作業完全成功！

### 5.1 重要踩坑提醒：如何正確儲存與更新網頁應用程式 (維持相同網址)

在開發與維護 Web App 時，很多同仁最常遇到的問題是：「修改了 `.gs` 程式碼，但為什麼重新整理網頁後沒看見更新？」或是「每次更新程式，網址就變了，導致外部系統都要重新設定。」

這通常是因為以下兩個關鍵動作沒有正確執行：

#### 避坑步驟一：程式碼必須「儲存變更」

在 Apps Script 中，若程式碼檔案旁邊亮起 **橘色圓點**，代表有變更但尚未存檔。如果此時直接去部署，運行的將會是**舊版程式碼**。請務必使用 `Ctrl + S` 存檔，或點選編輯器上方的「儲存」圖示，確保橘色圓點消失。

![橘色圓點亮起代表尚未儲存，部署選單可見管理部署作業](./images/gas_deploy_pitfall_save.png)

#### 避坑步驟二：必須透過「管理部署」建立新版本，而非「新增部署」

當我們要更新已發布的 Web App 且**維持原本的網頁 URL 網址不變**時，**絕對不要**再次點選「新增部署作業」（這會產生一個全新 ID 的網址），而應依循以下步驟：

1. 點選右上角 **「部署」** > **「管理部署作業」**。
2. 在彈出的管理視窗中，點擊右上角的 **「編輯 (鉛筆圖示)」**：
   ![管理部署作業中點選編輯鉛筆圖示](./images/gas_deploy_pitfall_manage.png)
3. 點開「版本」下拉選單，選擇 **「建立新版本」**：
   ![版本下拉選單選取建立新版本](./images/gas_deploy_pitfall_select_version.png)
4. 點選右下角的 **「部署」** 按鈕完成更新：
   ![選取建立新版本後點選部署](./images/gas_deploy_pitfall_new_version.png)
5. 成功更新後，系統會顯示「已成功更新部署作業」，此時其網網頁應用程式網址依然維持完全相同：
   ![部署更新成功完成畫面](./images/gas_deploy_pitfall_success.png)

---

## 6. 實戰演練二：雙表帳密登入與問答系統 (V2 - GAS 內建網頁版)

在第一個實戰確認部署成功後，我們將進一步模擬真實的 ERP 前後端整合：建立一個具有**帳密登入控制**與**問答撈取**的 Web App 系統。

### 6.1 系統架構設計

本專案使用以下架構：
1. **資料庫（兩張工作表）**：
   * `Passwords`：儲存允許登入的帳號與密碼（包含角色與備註）。
   * `QA`：儲存 5 組問答題目與答案。
2. **初始化腳本 (`setup.gs`)**：
   * 自動檢測是否存在 `Passwords` 與 `QA` 工作表。若不存在則建立，並寫入標頭欄位與初始預設資料（包含預設帳密與 5 組 QA）。
3. **內建呈現網頁的兩種做法**：
   * **做法 A：行內字串版**：直接將網頁 HTML 字串宣告在 `.gs` 程式碼中輸出。適合快速驗證。
   * **做法 B：獨立網頁檔案版 (推薦)**：在 GAS 專案內點選「+」建立一個 `index.html` 檔案，透過 `google.script.run` 進行前後端非同步通訊。這可避開所有 CORS 跨網域限制。

---

### 6.2 雙表資料初始化腳本 (`setup.gs`) 的建立與執行步驟

不論採用做法 A 還是做法 B，我們都需要先建立並初始化 `Passwords` 與 `QA` 資料表。請遵循以下步驟：

#### 步驟 1：建立 `setup.gs` 檔案
1. 在 Apps Script 編輯器左側的「檔案」旁，點選 **「+」** 按鈕。
2. 在下拉選單中選擇 **「指令碼」**：
   ![點選新增指令碼](./images/gas_v2_setup_add_file.png)
3. 將新建的檔案命名為 **`setup`** (系統會自動加上 `.gs` 副檔名)：
   ![新建setup.gs檔案](./images/gas_v2_setup_name_file.png)

#### 步驟 2：貼入初始化程式碼並儲存
1. 清空 `setup.gs` 中的預設代碼，將以下程式碼完整複製並貼上。
2. 點擊編輯器上方的 **「儲存 (磁碟圖示)」** 或使用 `Ctrl + S` 進行存檔：

```javascript
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. 初始化 Passwords 工作表
  var passwordSheet = ss.getSheetByName("Passwords");
  if (!passwordSheet) {
    passwordSheet = ss.insertSheet("Passwords");
    passwordSheet.appendRow(["Username", "Password", "Role", "Note"]);
    passwordSheet.getRange("A1:D1").setFontWeight("bold").setBackground("#cfe2f3");
    passwordSheet.appendRow(["admin", "admin123", "Admin", "Default Admin Account"]);
    passwordSheet.appendRow(["user", "user123", "User", "Default User Account"]);
  }
  
  // 2. 初始化 QA 工作表
  var qaSheet = ss.getSheetByName("QA");
  if (!qaSheet) {
    qaSheet = ss.insertSheet("QA");
    qaSheet.appendRow(["ID", "Question", "Answer"]);
    qaSheet.getRange("A1:C1").setFontWeight("bold").setBackground("#d9ead3");
    
    // 預設 5 組問答
    qaSheet.appendRow(["1", "台西分公司營業時間？", "週一至週五 09:00 - 15:30。"]);
    qaSheet.appendRow(["2", "如何聯絡台西分公司客服？", "請撥打分機 #888。"]);
    qaSheet.appendRow(["3", "銀行軟體系統每日結帳時間？", "每日下午 17:00 進行批次結帳。"]);
    qaSheet.appendRow(["4", "如何申請測試帳號？", "請填寫 ERP 權限申請單送交系統管理員。"]);
    qaSheet.appendRow(["5", "系統發生連線逾時如何處理？", "請確認網路 VPN 狀態，並清除瀏覽器快取。"]);
  }
  
  Logger.log("V2 資料庫初始化與資料正規化設定完成！");
}
```

   ![貼入程式碼並存檔](./images/gas_v2_setup_save_code.png)

#### 步驟 3：執行 `setup` 函數進行授權與建表
1. 確保上方的函數下拉選單中選定為 **`setup`**。
2. 點擊選單列的 **「執行」** 按鈕。
3. 由於本程式碼包含了對 Google 試算表 (SpreadsheetApp) 的存取與編輯，系統會跳出確認視窗，提示需要授權。請點擊 **「審查權限」** 進入授權程序：
   ![點擊審查權限按鈕](./images/gas_v2_setup_auth_required.png)
4. 選擇您的 Google 帳號，並在「這個應用程式未經 Google 驗證」警告畫面中，點開 **「進階」** 隱藏選項，並點選最下方的 **「前往『gas-study-v2』(不安全)」** 連結：
   ![選取前往未命名專案不安全網址](./images/gas_v2_setup_auth_unsafe.png)
5. 在最後的權限確認畫面中，確認允許此專案操作您的試算表，點擊右下角的 **「繼續 (Continue)」** 完成授權放行：
   ![允許授權並點選繼續](./images/gas_v2_setup_auth_continue.png)
6. 授權完畢後，編輯器下方執行記錄將顯示「V2 資料庫初始化與資料正規化設定完成！」及「執行完畢」，這代表您的 `Passwords` 與 `QA` 工作表已成功在雲端試算表中自動建立完成：
   ![執行完畢且資料庫建立成功](./images/gas_v2_setup_run_success.png)
6. 此時點開您的 Google 試算表，您會發現新增了兩個分頁：`Passwords` 工作表（已自動寫入預設帳密與備註）與 `QA` 工作表（已寫入 5 組去識別化的銀行軟體問答）：
   * **Passwords 表實例畫面**：
     ![Passwords工作表](./images/gas_v2_sheet_passwords.png)
   * **QA 表實例畫面**：
     ![QA工作表](./images/gas_v2_sheet_qa.png)

---

### 6.3 做法 A：極簡單檔行內字串版

不建立任何獨立 HTML 檔案，所有的網頁代碼都以「字串」形式寫在 `程式碼.gs` 中：

#### 伺服器端：`程式碼.gs`
```javascript
function doGet(e) {
  var action = e.parameter.action;
  
  // 處理 QA 撈取
  if (action === "getQA") {
    return getQAData();
  }
  
  // 處理登入驗證
  if (action === "login") {
    return handleLogin(e.parameter.username, e.parameter.password);
  }
  
  // 預設輸出網頁 (行內字串版)
  var htmlString = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Taixi Branch QA</title>' +
    '<style>body{font-family:Arial;background:#0b0f19;color:#fff;display:flex;justify-content:center;align-items:center;min-height:100vh;margin:0;}' +
    '.card{background:#171c29;border:1px solid #06b6d4;border-radius:12px;padding:30px;width:300px;}' +
    'input,select,button{width:100%;padding:10px;margin-top:10px;background:#1f2937;color:#fff;border:1px solid #374151;border-radius:6px;box-sizing:border-box;}' +
    'button{background:#06b6d4;font-weight:bold;cursor:pointer;}' +
    '.hidden{display:none;}.error{color:#ef4444;font-size:0.85rem;margin-top:5px; text-align:center;}' +
    '</style></head><body>' +
    '<div id="login" class="card"><h2>系統登入</h2><input type="text" id="user" placeholder="帳號"><input type="password" id="pass" placeholder="密碼"><button onclick="login()">登入</button><div id="err" class="error hidden"></div></div>' +
    '<div id="qa" class="card hidden"><h2>問答查詢系統</h2><select id="sel" onchange="showAns()"><option value="">-- 請選擇 --</option></select><div id="ans" style="margin-top:15px;padding:10px;background:#1e293b;border-left:4px solid #10b981;display:none;"></div></div>' +
    '<script>' +
    'var webUrl = ScriptApp.getService().getUrl();' +
    'var qaList = [];' +
    'function login() {' +
    '  var u = document.getElementById("user").value.trim();' +
    '  var p = document.getElementById("pass").value.trim();' +
    '  if(!u||!p) { alert("請輸入完整帳密！"); return; }' +
    '  fetch(webUrl + "?action=login&username=" + encodeURIComponent(u) + "&password=" + encodeURIComponent(p))' +
    '    .then(r => r.json()).then(res => {' +
    '      if(res.success){' +
    '        document.getElementById("login").classList.add("hidden");' +
    '        document.getElementById("qa").classList.remove("hidden");' +
    '        loadQA();' +
    '      } else { var e=document.getElementById("err"); e.innerText=res.message; e.classList.remove("hidden"); }' +
    '    });' +
    '}' +
    'function loadQA() {' +
    '  fetch(webUrl + "?action=getQA").then(r => r.json()).then(res => {' +
    '    if(res.success){' +
    '      qaList = res.data;' +
    '      var sel = document.getElementById("sel");' +
    '      qaList.forEach(item => {' +
    '        var opt = document.createElement("option");' +
    '        opt.value = item.id; opt.innerText = item.question;' +
    '        sel.appendChild(opt);' +
    '      });' +
    '    }' +
    '  });' +
    '}' +
    'function showAns() {' +
    '  var val = document.getElementById("sel").value;' +
    '  var box = document.getElementById("ans");' +
    '  if(!val){ box.style.display="none"; return; }' +
    '  var found = qaList.find(x => String(x.id) === String(val));' +
    '  if(found){ box.innerText = "答案：" + found.answer; box.style.display="block"; }' +
    '}' +
    '</script></body></html>';
  
  return HtmlService.createHtmlOutput(htmlString);
}

function handleLogin(username, password) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Passwords");
  if (!sheet) return createJsonResponse({ success: false, message: "Passwords sheet not found" });
  
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === username && String(data[i][1]) === String(password)) {
      return createJsonResponse({ success: true, role: data[i][2] });
    }
  }
  return createJsonResponse({ success: false, message: "帳號或密碼錯誤" });
}

function getQAData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("QA");
  if (!sheet) return createJsonResponse({ success: false, message: "QA sheet not found" });
  
  var data = sheet.getDataRange().getValues();
  var qaList = [];
  for (var i = 1; i < data.length; i++) {
    qaList.push({ id: data[i][0], question: data[i][1], answer: data[i][2] });
  }
  return createJsonResponse({ success: true, data: qaList });
}

function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
```

---

### 6.4 做法 B：標準獨立檔案與 `google.script.run` 非同步通訊版 (最推薦)

我們在 Apps Script 專案內建立兩個檔案，利用 Google 內建的非同步通訊管道 `google.script.run` 呼叫後端函數。此做法排版乾淨、支援完整網頁開發且**不需要處理跨網域 CORS 問題**。

#### 伺服器端：`程式碼.gs`
```javascript
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('台西分公司問答系統')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 供前端 google.script.run.handleLogin() 呼叫
function handleLogin(username, password) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Passwords");
  if (!sheet) return { success: false, message: "Passwords sheet not found" };
  
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === username && String(data[i][1]) === String(password)) {
      return { success: true, role: data[i][2] };
    }
  }
  return { success: false, message: "帳號或密碼錯誤" };
}

// 供前端 google.script.run.getQAData() 呼叫
function getQAData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("QA");
  if (!sheet) return { success: false, message: "QA sheet not found" };
  
  var data = sheet.getDataRange().getValues();
  var qaList = [];
  for (var i = 1; i < data.length; i++) {
    qaList.push({ id: data[i][0], question: data[i][1], answer: data[i][2] });
  }
  return { success: true, data: qaList };
}
```

#### 用戶端網頁：[NEW] `index.html`
請在專案中新增一個名為 `index.html` 的 HTML 檔案，並貼入以下程式碼：
```html
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Bank Software Taixi Branch - QA System</title>
  <style>
    body {
      font-family: Arial, sans-serif;
      background-color: #0b0f19;
      color: #f3f4f6;
      display: flex;
      justify-content: center;
      align-items: center;
      min-height: 100vh;
      margin: 0;
    }
    .card {
      background-color: #171c29;
      border: 1px solid #06b6d4;
      border-radius: 12px;
      padding: 30px;
      width: 100%;
      max-width: 400px;
      box-shadow: 0 4px 20px rgba(6, 182, 212, 0.2);
    }
    h2 {
      margin-top: 0;
      color: #06b6d4;
      text-align: center;
    }
    .form-group {
      margin-bottom: 15px;
    }
    label {
      display: block;
      margin-bottom: 5px;
      font-size: 0.9rem;
    }
    input, select {
      width: 100%;
      padding: 10px;
      border: 1px solid #374151;
      background-color: #1f2937;
      color: #ffffff;
      border-radius: 6px;
      box-sizing: border-box;
    }
    button {
      width: 100%;
      padding: 10px;
      background-color: #06b6d4;
      border: none;
      color: white;
      font-weight: bold;
      border-radius: 6px;
      cursor: pointer;
      margin-top: 10px;
    }
    button:hover {
      background-color: #0891b2;
    }
    .hidden {
      display: none;
    }
    .error-msg {
      color: #ef4444;
      font-size: 0.85rem;
      margin-top: 5px;
      text-align: center;
    }
    .answer-box {
      margin-top: 20px;
      padding: 15px;
      background-color: #1e293b;
      border-left: 4px solid #10b981;
      border-radius: 4px;
    }
  </style>
</head>
<body>

  <!-- 1. 登入面板 -->
  <div id="login-panel" class="card">
    <h2>系統登入</h2>
    <div class="form-group">
      <label for="username">帳號</label>
      <input type="text" id="username" placeholder="請輸入帳號">
    </div>
    <div class="form-group">
      <label for="password">密碼</label>
      <input type="password" id="password" placeholder="請輸入密碼">
    </div>
    <button id="login-btn">登入</button>
    <div id="login-error" class="error-msg hidden"></div>
  </div>

  <!-- 2. QA 系統面板 -->
  <div id="qa-panel" class="card hidden">
    <h2>問答查詢系統</h2>
    <div class="form-group">
      <label for="qa-select">請選擇問題</label>
      <select id="qa-select">
        <option value="">-- 請選擇 --</option>
      </select>
    </div>
    <div id="answer-container" class="answer-box hidden">
      <strong style="color: #10b981;">答案：</strong>
      <div id="answer-text" style="margin-top: 5px;"></div>
    </div>
    <button id="logout-btn" style="background-color: #4b5563;">登出</button>
  </div>

  <script>
    const loginPanel = document.getElementById("login-panel");
    const qaPanel = document.getElementById("qa-panel");
    const loginBtn = document.getElementById("login-btn");
    const logoutBtn = document.getElementById("logout-btn");
    const loginError = document.getElementById("login-error");
    const qaSelect = document.getElementById("qa-select");
    const answerContainer = document.getElementById("answer-container");
    const answerText = document.getElementById("answer-text");

    let qaData = [];

    // 處理登入
    loginBtn.addEventListener("click", function() {
      const user = document.getElementById("username").value.trim();
      const pass = document.getElementById("password").value.trim();
      
      if (!user || !pass) {
        showError("請輸入完整帳密！");
        return;
      }

      loginError.classList.add("hidden");
      loginBtn.disabled = true;
      loginBtn.innerText = "驗證中...";

      // 呼叫 GAS 後端 handleLogin 函數
      google.script.run
        .withSuccessHandler(function(result) {
          loginBtn.disabled = false;
          loginBtn.innerText = "登入";
          
          if (result.success) {
            loginPanel.classList.add("hidden");
            qaPanel.classList.remove("hidden");
            loadQA();
          } else {
            showError(result.message);
          }
        })
        .handleLogin(user, pass);
    });

    // 撈取 QA 資料
    function loadQA() {
      google.script.run
        .withSuccessHandler(function(result) {
          if (result.success) {
            qaData = result.data;
            qaSelect.innerHTML = '<option value="">-- 請選擇 --</option>';
            qaData.forEach(item => {
              const opt = document.createElement("option");
              opt.value = item.id;
              opt.innerText = item.question;
              qaSelect.appendChild(opt);
            });
          } else {
            alert("撈取 QA 失敗：" + result.message);
          }
        })
        .getQAData();
    }

    // 選取問題後顯示答案
    qaSelect.addEventListener("change", function() {
      const selectedId = qaSelect.value;
      if (!selectedId) {
        answerContainer.classList.add("hidden");
        return;
      }
      
      const found = qaData.find(item => String(item.id) === String(selectedId));
      if (found) {
        answerText.innerText = found.answer;
        answerContainer.classList.remove("hidden");
      }
    });

    // 登出
    logoutBtn.addEventListener("click", function() {
      document.getElementById("username").value = "";
      document.getElementById("password").value = "";
      qaPanel.classList.add("hidden");
      loginPanel.classList.remove("hidden");
      answerContainer.classList.add("hidden");
    });

    function showError(msg) {
      loginError.innerText = msg;
      loginError.classList.remove("hidden");
    }
  </script>
</body>
</html>
```

---

## 7. 實戰演練三：前後端分離與環境變數安全防護 (V3 - 地端/GitHub HTML + GAS Proxy API)

在掌握了 GAS 內建渲染的 V2 版本後，接下來我們將系統升級為 **前後端分離架構**。前端網頁不再託管於 Google Apps Script 內部，而是放在**同仁的本地端電腦或發布於 GitHub Pages**，僅透過 API 形式呼叫後端 GAS（作為資料 Proxy 代理）。

此外，我們將導入 **環境變數 (Script Properties)** 的安全防護觀念，將敏感帳密從代碼與工作表中抽離。

### 7.1 安全性設計：使用指令碼屬性 (Script Properties)

目前的系統中，密碼是直接以明碼形式寫在 `Passwords` 工作表內。在實際 ERP 整合時，重要的 API Key、外部資料庫密碼或管理員初始憑證絕不可外露。
* **安全變更**：我們將把管理員的初始帳密存放在 GAS 的 **「指令碼屬性 (Script Properties)」** 後台。
* **設定方法**：
  1. 在 Apps Script 編輯器左側點選 ⚙️ **「專案設定」**。
  2. 下拉至「指令碼屬性」區塊，點選「新增指令碼屬性」。
  3. 新增屬性 `ADMIN_PASSWORD`，值設定為 `admin123`；新增屬性 `ADMIN_USERNAME`，值設定為 `admin`。
  4. 回到程式碼中，使用 `PropertiesService.getScriptProperties().getProperty("ADMIN_PASSWORD")` 進行動態讀取。

---

### 7.2 伺服器端 GAS 程式碼

我們在 Apps Script 中更新程式碼，以支援外部 `fetch` JSON 呼叫，並進行跨網域傳輸。

#### 檔案一：`setup.gs` (移除明碼密碼)
```javascript
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. 初始化 Passwords 工作表 (不再直接寫入明碼帳密，工作表僅儲存一般使用者備份)
  var passwordSheet = ss.getSheetByName("Passwords");
  if (!passwordSheet) {
    passwordSheet = ss.insertSheet("Passwords");
    passwordSheet.appendRow(["Username", "Role", "Note"]); // 欄位中移除了 Password
    passwordSheet.getRange("A1:C1").setFontWeight("bold").setBackground("#cfe2f3");
    passwordSheet.appendRow(["admin", "Admin", "Admin account set in Script Properties"]);
  }
  
  // 2. 初始化 QA 工作表
  var qaSheet = ss.getSheetByName("QA");
  if (!qaSheet) {
    qaSheet = ss.insertSheet("QA");
    qaSheet.appendRow(["ID", "Question", "Answer"]);
    qaSheet.getRange("A1:C1").setFontWeight("bold").setBackground("#d9ead3");
    
    // 預設 5 組問答
    qaSheet.appendRow(["1", "台西分公司營業時間？", "週一至週五 09:00 - 15:30。"]);
    qaSheet.appendRow(["2", "如何聯絡台西分公司客服？", "請撥打分機 #888。"]);
    qaSheet.appendRow(["3", "銀行軟體系統每日結帳時間？", "每日下午 17:00 進行批次結帳。"]);
    qaSheet.appendRow(["4", "如何申請測試帳號？", "請填寫 ERP 權限申請單送交系統管理員。"]);
    qaSheet.appendRow(["5", "系統發生連線逾時如何處理？", "請確認網路 VPN 狀態，並清除瀏覽器快取。"]);
  }
  Logger.log("V3 資料庫初始化完成！敏感帳密已從 Sheets 中抽離。");
}
```

#### 檔案二：`程式碼.gs` (改為 API Proxy 路由並啟用安全讀取)
```javascript
function doGet(e) {
  var action = e.parameter.action;
  
  if (action === "getQA") {
    return getQAData();
  }
  
  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    message: "Bank Software Taixi Branch V3 Server is running."
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var postData;
  try {
    postData = JSON.parse(e.postData.contents);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: "Invalid JSON format"
    })).setMimeType(ContentService.MimeType.JSON);
  }
  
  var action = postData.action;
  
  if (action === "login") {
    return handleLogin(postData.username, postData.password);
  }
  
  return ContentService.createTextOutput(JSON.stringify({
    success: false,
    message: "Unknown action"
  })).setMimeType(ContentService.MimeType.JSON);
}

function handleLogin(username, password) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var sysAdminUser = scriptProperties.getProperty("ADMIN_USERNAME");
  var sysAdminPass = scriptProperties.getProperty("ADMIN_PASSWORD");
  
  // 1. 優先比對「指令碼屬性」安全環境變數
  if (username === sysAdminUser && password === sysAdminPass) {
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      role: "Admin",
      message: "Login successful (Admin)"
    })).setMimeType(ContentService.MimeType.JSON);
  }
  
  // 2. 一般使用者驗證（亦可在此串接其他加密機制）
  return ContentService.createTextOutput(JSON.stringify({
    success: false,
    message: "Invalid username or password"
  })).setMimeType(ContentService.MimeType.JSON);
}

function getQAData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("QA");
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: "Database error: QA sheet not found"
    })).setMimeType(ContentService.MimeType.JSON);
  }
  
  var data = sheet.getDataRange().getValues();
  var qaList = [];
  for (var i = 1; i < data.length; i++) {
    qaList.push({ id: data[i][0], question: data[i][1], answer: data[i][2] });
  }
  
  return ContentService.createTextOutput(JSON.stringify({
    success: true,
    data: qaList
  })).setMimeType(ContentService.MimeType.JSON);
}
```

---

### 7.3 用戶端外部網頁程式碼 (`index.html`)

請在您本地電腦上建立一個 `index.html` 檔案（例如放在桌面或發布於 GitHub Pages 專案目錄下），貼入以下程式碼。**並請記得將程式碼中的 `GAS_URL` 換成您 V3 重新部署後所產生的新網址：**

```html
<!DOCTYPE html>
<html lang="zh-TW">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Bank Software Taixi Branch - V3 QA System</title>
  <style>
    body {
      font-family: Arial, sans-serif;
      background-color: #0b0f19;
      color: #f3f4f6;
      display: flex;
      justify-content: center;
      align-items: center;
      min-height: 100vh;
      margin: 0;
    }
    .card {
      background-color: #171c29;
      border: 1px solid #06b6d4;
      border-radius: 12px;
      padding: 30px;
      width: 100%;
      max-width: 400px;
      box-shadow: 0 4px 20px rgba(6, 182, 212, 0.2);
    }
    h2 {
      margin-top: 0;
      color: #06b6d4;
      text-align: center;
    }
    .form-group {
      margin-bottom: 15px;
    }
    label {
      display: block;
      margin-bottom: 5px;
      font-size: 0.9rem;
    }
    input, select {
      width: 100%;
      padding: 10px;
      border: 1px solid #374151;
      background-color: #1f2937;
      color: #ffffff;
      border-radius: 6px;
      box-sizing: border-box;
    }
    button {
      width: 100%;
      padding: 10px;
      background-color: #06b6d4;
      border: none;
      color: white;
      font-weight: bold;
      border-radius: 6px;
      cursor: pointer;
      margin-top: 10px;
    }
    button:hover {
      background-color: #0891b2;
    }
    .hidden {
      display: none;
    }
    .error-msg {
      color: #ef4444;
      font-size: 0.85rem;
      margin-top: 5px;
      text-align: center;
    }
    .answer-box {
      margin-top: 20px;
      padding: 15px;
      background-color: #1e293b;
      border-left: 4px solid #10b981;
      border-radius: 4px;
    }
  </style>
</head>
<body>

  <!-- 1. 登入面板 -->
  <div id="login-panel" class="card">
    <h2>V3 系統登入</h2>
    <div class="form-group">
      <label for="username">帳號</label>
      <input type="text" id="username" placeholder="請輸入帳號">
    </div>
    <div class="form-group">
      <label for="password">密碼</label>
      <input type="password" id="password" placeholder="請輸入密碼">
    </div>
    <button id="login-btn">登入</button>
    <div id="login-error" class="error-msg hidden"></div>
  </div>

  <!-- 2. QA 系統面板 -->
  <div id="qa-panel" class="card hidden">
    <h2>問答查詢系統</h2>
    <div class="form-group">
      <label for="qa-select">請選擇問題</label>
      <select id="qa-select">
        <option value="">-- 請選擇 --</option>
      </select>
    </div>
    <div id="answer-container" class="answer-box hidden">
      <strong style="color: #10b981;">答案：</strong>
      <div id="answer-text" style="margin-top: 5px;"></div>
    </div>
    <button id="logout-btn" style="background-color: #4b5563;">登出</button>
  </div>

  <script>
    // 這裡替換為您 V3 部署後生成的 Web App 網址
    const GAS_URL = "https://script.google.com/macros/s/YOUR_V3_DEPLOYED_ID/exec";

    const loginPanel = document.getElementById("login-panel");
    const qaPanel = document.getElementById("qa-panel");
    const loginBtn = document.getElementById("login-btn");
    const logoutBtn = document.getElementById("logout-btn");
    const loginError = document.getElementById("login-error");
    const qaSelect = document.getElementById("qa-select");
    const answerContainer = document.getElementById("answer-container");
    const answerText = document.getElementById("answer-text");

    let qaData = [];

    // 處理登入
    loginBtn.addEventListener("click", function() {
      const user = document.getElementById("username").value.trim();
      const pass = document.getElementById("password").value.trim();
      
      if (!user || !pass) {
        showError("請輸入完整帳密！");
        return;
      }

      loginError.classList.add("hidden");
      loginBtn.disabled = true;
      loginBtn.innerText = "驗證中...";

      // 使用簡單 POST 請求，避開 Preflight 限制
      fetch(GAS_URL, {
        method: "POST",
        body: JSON.stringify({
          action: "login",
          username: user,
          password: pass
        })
      })
      .then(res => res.json())
      .then(result => {
        loginBtn.disabled = false;
        loginBtn.innerText = "登入";
        
        if (result.success) {
          loginPanel.classList.add("hidden");
          qaPanel.classList.remove("hidden");
          loadQA();
        } else {
          showError(result.message || "登入失敗！");
        }
      })
      .catch(err => {
        loginBtn.disabled = false;
        loginBtn.innerText = "登入";
        showError("連線伺服器失敗，請確認 GAS URL 是否填寫正確！");
        console.error(err);
      });
    });

    // 撈取 QA 資料
    function loadQA() {
      fetch(GAS_URL + "?action=getQA")
        .then(res => res.json())
        .then(result => {
          if (result.success) {
            qaData = result.data;
            qaSelect.innerHTML = '<option value="">-- 請選擇 --</option>';
            qaData.forEach(item => {
              const opt = document.createElement("option");
              opt.value = item.id;
              opt.innerText = item.question;
              qaSelect.appendChild(opt);
            });
          } else {
            alert("撈取 QA 失敗：" + result.message);
          }
        })
        .catch(err => {
          alert("撈取 QA 連線失敗！");
          console.error(err);
        });
    }

    // 選取問題後顯示答案
    qaSelect.addEventListener("change", function() {
      const selectedId = qaSelect.value;
      if (!selectedId) {
        answerContainer.classList.add("hidden");
        return;
      }
      
      const found = qaData.find(item => String(item.id) === String(selectedId));
      if (found) {
        answerText.innerText = found.answer;
        answerContainer.classList.remove("hidden");
      }
    });

    // 登出
    logoutBtn.addEventListener("click", function() {
      document.getElementById("username").value = "";
      document.getElementById("password").value = "";
      qaPanel.classList.add("hidden");
      loginPanel.classList.remove("hidden");
      answerContainer.classList.add("hidden");
    });

    function showError(msg) {
      loginError.innerText = msg;
      loginError.classList.remove("hidden");
    }
  </script>
</body>
</html>
```

---

### 7.4 跨網域對接關鍵：避開 OPTIONS 預檢 (Preflight) 限制

在進行跨網域 (CORS) 對接時，最常見的問題是瀏覽器報出 CORS 錯誤。
1. **什麼是 OPTIONS 預檢？**：當瀏覽器偵測到您向跨網域發送非簡單請求（例如設定了自訂的 Headers，或是將 Content-Type 設為 `application/json`）時，會先自動發送一個 `OPTIONS` 方法的請求到伺服器確認權限。
2. **GAS 的限制**：Google Apps Script 的 Web App **不支援** CORS 預檢（OPTIONS 請求），會直接拒絕並回報失敗。
3. **解決之道（極穩健做法）**：
   * 我們在前端發送 `fetch` 時，**不設定任何自訂的 Headers**。
   * 此時瀏覽器會採用 **「簡單請求 (Simple Request)」** 格式（即預設 `text/plain` 傳送字串），進而**完全避開** OPTIONS 預檢限制。
   * 伺服器端的 GAS 使用 `JSON.parse(e.postData.contents)` 來解析字串，即可完美實現跨網域資料交換！

---

## 8. 常見問題與 AI 互動排錯

* **情境 A：修改了程式碼，重新整理網頁卻沒有任何變化？**
  * **原理解析**：GAS 正式部署網址（以 `/exec` 結尾）是版本鎖定的。當修改了程式碼後，直接重新整理網頁無法更新內容。
  * **與 AI 互動**：
    > *「我改了 `doGet` 的 HTML 字串，但重新整理 Web App 網址卻沒有更新，要怎麼處理？」*
  * **AI 指導方案**：
    1. **開發測試**：請使用「測試部署 (Test deployments)」網址（結尾是 `/dev`），該網址會即時反映儲存後的代碼變更。
    2. **正式發布**：在「管理部署」中編輯現有部署，選擇 **「新增版本」** 後重新部署，舊的 `/exec` 網址才會反映更新。
* **情境 B：其他同仁點開網址，畫面顯示「需要授權」或「無法存取」？**
  * **原理解析**：部署時的「誰有權限存取 (Who has access)」設定有誤，未開放給外部或他人。
  * **與 AI 互動**：
    > *「我部署的 GAS Web App 網址給同仁點開，提示需要登入或無法存取，請問要在哪裡修改部署參數？」*
  * **AI 指導方案**：
    請前往「管理部署」編輯該部署，將 Who has access 修改為 **「所有人」**（在英文介面是 Anyone）並重新部署。
* **情境 C：貼上代碼儲存時提示 `語法錯誤：SyntaxError: Invalid or unexpected token`**
  * **原理解析**：在 JavaScript 中，若使用普通雙引號 `"` 定義字串，該字串**不可**在程式碼中直接按 Enter 鍵換行（折行），否則會導致編譯錯誤。
  * **與 AI 互動**：
    > *「我貼上 `doGet` 代碼時提示 SyntaxError: Invalid or unexpected token，要怎麼修復？」*
  * **AI 指導方案**：
    請確保字串在同一行閉合，或將雙引號 `"` 改為反引號 `` ` ``（Template Literals，支援直接折行）。
