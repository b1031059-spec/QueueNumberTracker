Last Updated: March 26, 2026

---

**Please follow the steps below to complete the environment setup and run the program:**

    1.Download Executable:
        Go to the Releases page and download the latest .exe file and relevant configuration files.

    2.Configure Directory:
        Place the downloaded executable in the same directory as the hospital_doctors.yaml file.

---

**Parameter Settings (hospital_doctors.yaml):**
Before running the program, ensure that hospital_doctors.yaml is correctly formatted. This file defines the initialization behavior:

After configuration, please ensure the hospital_doctors.yaml format is correct. This file defines the following initialization behaviors:

    Language: Set the interface language; supports "zh" (Traditional Chinese) or "en" (English).

    Hospital List: Preset frequently used hospital names, URLs, and doctor names to enable quick switching within the GUI.

    lastBBox: The system automatically records the position and size (Bounding Box) of the last captured area, which will be applied automatically upon the next launch.

---

Running the Program
Important Precautions:

    Screen Sleep: This program acquires data via screen capturing. If you need to run it for a long duration, please adjust your computer's Auto-Sleep settings to prevent capture failure caused by the screen turning off.

    First Launch: Double-click QueueNumberTracker_v1.0.exe to launch. The program requires an internet connection during the first launch to download image recognition models (approx. 30–60 seconds). If the interface appears unresponsive during this time, please wait patiently.

---

GUI Operation Instructions:
This tool provides an intuitive five-page workflow:

    Page 1: Screen Capture Settings

        Operation: Use the sliders or input boxes on the right to adjust the capture area's coordinates and size.

        Preview: The left window displays the selected area in real-time. Ensure it covers the consultation number area.

    Page 2: Recognition & Schedule Settings

        File Naming: Confirm the Excel filename at the top.

        Recognition Test: Verify that the OCR result on the left accurately reads the numbers.

        Parameter Settings: Set the total operating duration and recognition interval on the right.

    Page 3: Operation Log

        Displays real-time running status and recognition records.

        Supports manual stop or automatic shutdown after the schedule completes.

        Data Storage: Upon launch, the program automatically creates a data folder in the current directory. All data will be saved here in Excel format.

    Page 4: Data Post-processing

        Noise Filtering: Check this to generate "clean" data by filtering out abnormal number jumps (creates an additional sheet without modifying original data).

        Threshold Settings: Configure the threshold for the filtering algorithm.

        Visualization: Check this to automatically generate analytical charts based on the data.

        Algorithm details: [Link (Pending Publication)]

    Page 5: Result Display

        Displays the execution results of the filtering algorithm (success status or specific error causes).

---

Technical Details:
    Operating System: Windows
    Programming Language: Python
    Recognition Engine: EasyOCR (Optical Character Recognition)
    Storage Format: Microsoft Excel (.xlsx)

---

最後更新時間: 2026/3/26

**請依照以下步驟完成環境設定與程式執行**

1.下載並解壓縮：
    前往 Releases 頁面，下載壓縮檔並解壓縮。

2.檢查配置目錄：
    確認有三個檔案
        _internal
        .yaml
        .exe
    位於同一個目錄。
    
---

**參數設定 (hospital_doctors.yaml)：**
程式執行依賴 hospital_doctors.yaml 設定檔，設定後可以方便的對儲存的資料進行命名與分類，就算不特別進行設定也可以運作程式，如要進行設定，您可以直接對 hospital_doctors.yaml 進行修改。

設定後，請確保 hospital_doctors.yaml 格式正確。此檔案用於定義初始化行為

    語言：設定介面語言，支援 "zh" (繁體中文) 或 "en" (英文)。

    醫院列表：可依格式預設常用的醫院名稱、網址及醫生姓名，實現在 GUI 內快速切換。

    lastBBox：系統會自動記錄上次截取畫面的位置與尺寸（Bounding Box），下次啟動時將自動帶入，無需重新調整。

---

**執行程式**
使用前注意事項：

    螢幕休眠：本程式透過螢幕擷取獲得數據。若需長時間執行，請調整電腦的自動休眠時間，以免螢幕關閉導致擷取失敗。

    初次啟動：雙擊 .exe 啟動。初次執行需連網下載影像辨識模型，約需 30 秒至 1 分鐘。期間介面若無回應屬正常現象，請耐心等候。

---
**GUI 操作說明：**
本工具提供直覺的五分頁操作流程：

    Page 1：畫面截取設定
        操作：利用右側的拉桿或輸入框來調整擷取區域的坐標與大小。
        預覽：左側視窗會即時顯示目前選定的畫面，請確保覆蓋到看診號碼區域。

    Page 2：辨識與排程設定
        檔案命名：於上方確認 Excel 儲存的檔名。
        辨識測試：確認左側影像辨識結果是否精準讀取數字。
        參數設定：於右側設定程式運作的總時數及辨識間隔。

    Page 3：運作日誌 (Log)
        實時顯示程式的運行狀態與辨識記錄。
        支援手動停止或等待排程結束後自動關閉。
        程式啟動後，會自動在當前目錄建立 data 資料夾。所有的數據將以 Excel 格式儲存於此。

    Page 4：數據後處理
        雜訊過濾：勾選是否產生過濾掉異常跳號後的純淨資料(產生額外頁面，不動原始資料)。
        參數設定：設定過濾演算法的閥值。
        視覺化：勾選是否根據數據自動生成分析圖表。

        關於過濾演算法的詳細邏輯，請參考：[連結(尚未刊登)]

    Page 5：結果展示
        顯示過濾演算法執行結果(是否成功或是發生錯誤原因)。

---
        
技術細節:
    作業系統：Windows 
    開發語言：Python
    辨識引擎：EasyOCR (Optical Character Recognition)
    儲存格式：Microsoft Excel (.xlsx)
