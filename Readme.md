# Channel Check Tool (CCT)

## 專案概覽
Channel Check Tool（CCT）協助進行 DDR 與其他高速通道的電氣驗證。工具會讀取 Touchstone S 參數檔與連接埠中繼資料，判別控制端與記憶體端節點，建立等效電路，並在 Ansys Electronics Desktop（AEDT）Circuit 中執行暫態模擬，以產生波形與訊號完整性統計。

## 主要功能
- 解析 JSON 連接埠中繼資料，辨識單端與差動配對並統一命名。
- 生成相容 Nexxim 的網表，可自訂傳輸端（Tx）與接收端（Rx）終端條件，並呼叫 AEDT Circuit 進行時域分析。
- 若已安裝 scikit-rf，可選擇性剪枝 Touchstone 連接埠，只保留超過臨界值的通道。
- 計算波形積分、ISI 與相關指標，供後續報告或 GUI 使用。
- 內建 PySide GUI（`src/aedb_gui.py`），可選擇輸入檔、調整參數並監控進度。
- 提供前處理範例（`src/1_pre_process.py`），示範如何自 EDB/BRD 設計建立連接埠與頻率掃描。

## 目錄結構
- `src/cct.py`：處理中繼資料、電路生成、模擬與後處理的核心邏輯。
- `src/aedb_gui.py`：PySide GUI 與封裝 CCT 後端的背景工作。 
- `src/run.py`：建立虛擬環境並安裝相依套件的 Python 輔助腳本。

- `run.bat`／`install.bat`：Windows 平台上的安裝與啟動批次檔。
- `data/`：範例資料，包含 `.aedb`、`.sNp` 與 `*_ports.json`。

## 系統需求
- Windows 10 或更新版本，並安裝相容的 AEDT（建議 2024.2 以上）。
- Python 3.9 或更新版本；批次檔會自動建立與管理 `.venv`。
- 主要 Python 套件：`pyedb`、`pyaedt`、`PySide6`（或 PySide2）、`numpy`、`scikit-rf`（可選，用於剪枝）。
- GUI 所需的 Qt 平台外掛（`install.bat` 會協助設定）。

## 安裝步驟
1. 在專案根目錄開啟 PowerShell 或命令提示字元。
2. 執行 `install.bat` 以建立 `.venv` 並安裝相依套件；若在 Linux/WSL，請改執行 `python src/run.py`。
3. 確認 `.venv\Scripts\python.exe` 存在，並確保 pyaedt 可以呼叫到 AEDT。

## GUI 使用方式
1. 執行 `run.bat`，或啟動虛擬環境後執行 `python src/aedb_gui.py`。
2. 選取 `.sNp` 檔案與對應的 `*_ports.json` 中繼資料。
3. 輸入 Tx/Rx 參數，例如驅動電壓、上升時間與終端電阻／電容。
4. 需要剪枝時，可在 Prune 分頁設定臨界值（dB）。
5. 按下 Run 進行完整模擬，或使用 Pre-run 快速取得摘要；結果會顯示於 GUI 並輸出至中繼資料目錄。
6. 可於狀態列與日誌窗格追蹤進度；暫存檔會儲存在中繼資料旁的子資料夾。

## 輸出內容
- 模擬產物會儲存在中繼資料目錄下的 `cct_work/` 等資料夾。
- 啟用剪枝時，篩選後的 Touchstone 會輸出至 `trimmed_touchstone/`。
- 波形統計與 Tx/Rx 對應資訊會以 JSON 格式輸出，供後續分析。



## 開發者備註
- 建議在虛擬環境中執行 Python 指令：`".venv\Scripts\python" some_script.py`。
- 若需新增相依套件，請編輯 `requirements.txt`；`install.bat` 與 `src/run.py` 都會優先讀取該檔案。
- GUI 透過 Qt `QSettings` 儲存使用者偏好，位置依作業系統而異。
- 預設 AEDT 版本為 `2025.1`；可透過中繼資料或 GUI 選項覆寫。

## 疑難排解
- 找不到虛擬環境 Python：重新執行 `install.bat` 或確認 Python 已加入 PATH。
- AEDT 模擬失敗：確認授權狀態，並檢查 AEDT 與 `pyaedt` 版本相容性。
- GUI 無法啟動：確認 Qt 外掛已由 `install.bat` 安裝，必要時手動設定 `QT_PLUGIN_PATH`。

## 授權
本專案採用 [MIT License](LICENSE)。
