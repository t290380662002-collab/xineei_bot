# 訂房 Telegram Bot

把一筆訂房資料透過 Telegram 對話填寫，自動套用四間飯店（名匯 / 威尼斯 / 巴黎人 / 倫敦人）
的 Excel 模板，產生一份**獨立的新 Excel 檔**並回傳給你下載。每份產出互不覆蓋。

## 對話流程
`/start` → 選飯店 → 入住日期 → 退房日期 → 房型(代碼或中文) → 房數(件數)
→ 備注(可 /skip) → 是否吸煙 → 逐位入住者(中文 / 英文 / 出生 / 證件)
→ 可加多位 → 「✅ 完成，產生 Excel」→ 收到檔案

## 本機執行
1. 安裝 Python 3.11+
2. `pip install -r requirements.txt`
3. 設定 Token：
   - Windows(PowerShell)：`$env:TELEGRAM_BOT_TOKEN="你的token"`
   - macOS/Linux：`export TELEGRAM_BOT_TOKEN="你的token"`
4. `python bot.py`

## 申請 Telegram Bot Token（@BotFather）
1. 在 Telegram 搜尋 `@BotFather`，對它發 `/newbot`
2. 依指示輸入 bot 名稱與 username（結尾需為 `bot`，如 `MyBookingBot`）
3. 它會給你一組 `123456:ABC...` 的 token，貼到 `.env` 的 `TELEGRAM_BOT_TOKEN`

## 雲端部署（24 小時不斷電，隨時可收發）
Bot 需要常駐環境。推薦方案（費用都很低，多數有免費額度）：
- **Railway**：新建 Project → 上傳本資料夾 → 設定環境變數 `TELEGRAM_BOT_TOKEN` → Start Command 設 `python bot.py`
- **Render**：新建 Background Worker → 連結本資料夾 → 設定環境變數 → Build/Start 都用 `python bot.py`
- **PythonAnywhere / 任意 VPS**：把資料夾傳上去，`pip install -r requirements.txt` 後 `python bot.py` 常駐執行

> 部署時只需保證 `templates/`（四份空白模板）、`config.py`、`fill.py`、`bot.py` 在同一資料夾，
> 並設定環境變數 `TELEGRAM_BOT_TOKEN`。

## 檔案說明
- `config.py`：四間飯店的欄位對應、房型代碼↔方框對照表（要改欄位位置改這裡）
- `fill.py`：把訂房資料填進模板、產生新 Excel 的核心邏輯
- `bot.py`：Telegram 對話機器人
- `templates/`：四份原始空白 Excel（**格式不變，僅填入資料**）
- `test_fill.py`：離線測試，對四間飯店各產生一份範例 Excel 到 `output/`
