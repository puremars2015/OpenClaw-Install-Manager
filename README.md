# OpenClaw自動安裝程式 for Windows

## 小工具使用方法

在dist資料夾內有已經打包好的單一exe檔案，下載後直接執行即可：

如果直接要用載點下載：

[https://github.com/puremars2015/OpenClaw-Install-Manager/raw/refs/heads/main/dist/OpenClawManager.exe](https://github.com/puremars2015/OpenClaw-Install-Manager/raw/refs/heads/main/dist/OpenClawManager.exe)

## 介紹
OpenClaw自動安裝程式是一個用於簡化環境檢查與安裝流程的 Windows GUI 工具。它將介面收斂成兩個部分：

- 檢查 PowerShell 7、Node.js、npm、Git、Python、OpenCode、OpenClaw 是否已安裝。
- 安裝全部尚未有的環境套件，以及安裝指定版本的 OpenClaw 4.1。
- 安裝最新版 OpenClaw，或移除已安裝的 OpenClaw。
- 初始化 OpenClaw，或打開 OpenClaw 交談視窗。
- 設定預設 API Key，並切換預設模型供應商。
- 啟動與停止由本工具建立的 OpenClaw Gateway 行程。

## 功能
- 軟體以 Python 的 GUI 套件 tkinter 開發，提供簡化後的 Windows 安裝介面。
- GUI 提供三類功能：環境檢查、安裝、Gateway 啟停。
- 環境檢查會顯示 PowerShell 7、Node.js、npm、Git、Python、OpenCode、OpenClaw 的安裝狀態、版本與路徑。
- 「安裝全部尚未有的環境套件」會補齊 PowerShell 7、Node.js、npm、Git、Python、OpenCode。
- OpenCode 透過 `npm install -g opencode-ai` 安裝，指令名稱為 `opencode`。
- 「安裝 OpenClaw 4.1」會固定安裝 `openclaw@2026.4.1`。
- 「安裝最新版 OpenClaw」會執行 `npm install -g openclaw`。
- 「移除 OpenClaw」會執行 `npm uninstall -g openclaw`。
- 「OpenClaw初始化」會執行 `openclaw setup`。
- 「打開交談視窗」會執行 `openclaw dashboard`。
- 「設定預設 API Key」可選擇 `OpenRouter`、`OpenAI`、`Anthropic`、`MiniMax`，把 token 寫入 OpenClaw auth store，並切換到該供應商的預設模型。
- 「啟動 OpenClaw Gateway」會執行 `openclaw gateway run --force` 並把輸出顯示在 GUI 日誌區。
- 「停止 OpenClaw Gateway」只會停止由本工具啟動的 Gateway 行程。

## 已實作內容
- `openclaw_manager.py`：以 `tkinter` 開發的 Windows GUI 工具。
- `scripts/openclaw_helper.ps1`：負責檢查環境、安裝缺少的環境套件、安裝指定版本的 OpenClaw。
- `run_openclaw_manager.ps1`：Windows 啟動器，會自動用 `py -3` 或 `python` 啟動 GUI。

## 工具功能
- 檢查 PowerShell 7、Node.js、npm、Git、Python、OpenCode、OpenClaw 是否已安裝。
- 使用 `winget` 補安裝缺少的 PowerShell 7、Node.js LTS、Git、Python 3.12。
- 若缺少 `npm`，會透過安裝 Node.js LTS 一併補齊。
- 使用 `npm install -g opencode-ai` 安裝 OpenCode。
- 使用 `npm install -g openclaw@2026.4.1` 安裝 OpenClaw 4.1。
- 使用 `npm install -g openclaw` 安裝最新版 OpenClaw。
- 使用 `npm uninstall -g openclaw` 移除已安裝的 OpenClaw。
- 使用 `openclaw setup` 初始化 OpenClaw。
- 使用 `openclaw dashboard` 打開 OpenClaw 交談視窗。
- 可透過 GUI 設定 OpenClaw 的預設 API Key 與模型供應商，支援 OpenRouter、OpenAI、Anthropic、MiniMax。
- 透過 GUI 啟動 `openclaw gateway run --force`，並可從 GUI 停止同一個 Gateway 行程。

## 執行方式
1. 先確認本機已安裝 Python 3。
2. 在 PowerShell 執行：

```powershell
.\run_openclaw_manager.ps1
```

3. 進入 GUI 後，建議依序操作：
    - `重新檢查`
    - `安裝全部尚未有的環境套件`
    - `安裝 OpenClaw 4.1`
    - `OpenClaw初始化`
    - `打開交談視窗`
    - `設定預設 API Key`
    - `安裝最新版 OpenClaw` 或 `移除 OpenClaw`
    - `啟動 OpenClaw Gateway`

## 打包成單一 EXE
在專案根目錄執行：

```powershell
.\build_single_exe.ps1
```

完成後會產生單一執行檔：

```text
.\dist\OpenClawManager.exe
```

說明：
- 此打包方式使用 PyInstaller `--onefile`。
- `scripts/openclaw_helper.ps1` 會被一起打包進 EXE，執行時自動解壓到暫存目錄。
- 打包後使用者只需要拿到 `OpenClawManager.exe` 即可執行，不需要另外攜帶 `.ps1` 或 `.py` 檔案。

## 注意事項
- 程式啟動時會先要求管理者權限；若取消 UAC 提示，安裝程式會直接結束，不會以一般權限繼續執行。
- 依賴安裝使用 `winget`，若系統沒有 `winget`，需要先補安裝 App Installer。
- OpenClaw 官方 README 建議 Windows 使用 WSL2；本工具先提供原生 Windows 的檢查與安裝流程。
- 若剛安裝完 Node.js，helper 會優先用常見安裝路徑重新尋找 `node` 與 `npm`，降低 PATH 尚未刷新造成的誤判。
