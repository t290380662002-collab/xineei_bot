# 部署到 Render（Telegram 訂房 Bot，24h 常駐，免費方案）

Render 的 **Background Worker** 最適合這種輪詢式 Telegram Bot：免費、不會像 Web Service 那樣閒置後休眠。
本專案已經備好 `render.yaml`，照下面步驟即可上線。

---

## 第一步：把程式推到 GitHub

Render 必須從 Git 倉庫部署（沒有「上傳資料夾」功能），所以先要有 GitHub repo。

1. 到 https://github.com/new 開一個新的 **Public 或 Private** 倉庫（例如 `booking-bot`），**不要**勾選 Initialize with README。
2. 在本機專案資料夾執行（我已經幫你 `git init` 並 commit 好，你只需加遠端 + 推送）：

   ```bash
   git remote add origin https://github.com/<你的帳號>/<倉庫名>.git
   git branch -M main
   git push -u origin main
   ```

   > ⚠️ `.env`（含 Bot token）已被 `.gitignore` 排除，不會被上傳。
   > 若你還沒設 git 身份，先執行：
   > `git config --global user.email "you@example.com"` 與 `git config --global user.name "你的名字"`

---

## 第二步：在 Render 建立 Background Worker

1. 登入 https://dashboard.render.com （可用 GitHub 帳號登入）。
2. 右上角 **New +** → **Background Worker**。
3. 連接你的 GitHub 帳號，選剛才的 `booking-bot` 倉庫。
4. Render 會自動讀取 `render.yaml`，設定如下（也可手動確認）：
   - **Name**：`booking-bot`
   - **Environment**：`Python 3`
   - **Build Command**：`pip install -r requirements.txt`
   - **Start Command**：`python bot.py`
   - **Plan**：`Free`
5. 展開 **Advanced → Environment Variables**，新增一筆：
   - **Key**：`TELEGRAM_BOT_TOKEN`
   - **Value**：`8831075226:AAH53OF4qZiF1FhYcxwOIeFCT3ltK1QHLso`（你從 @BotFather 拿到的那組）
6. 點 **Create Worker**，Render 會自動 build 並啟動。

---

## 第三步：確認上線

- 在 Render 的 **Logs** 頁面看到 `Application started` 就成功了。
- 到 Telegram 對 **@xinwea_bot** 發 `/start`，確認能一步步填並收到 Excel。

---

## ⚠️ 重要：避免雙重輪詢衝突

Telegram **同一個 Bot 只能有一個程式在收訊息**。若你之前在本機/沙盒跑過 `bot.py`，部署到 Render 前請先把它停掉，否則會出現：

```
409 Conflict: terminated by other getUpdates request
```

停用方式：
- 本機：在跑程式的終端按 `Ctrl+C`。
- 本專案沙盒（我這邊幫你跑的那個）：告訴我「停掉本機 bot」，我會結束該程序。

確認本機實例已停，Render 的實例才會正常收訊。

---

## 日常管理

- **看日誌**：Render 後台 → 該 Worker → Logs。
- **改程式後上線**：`git commit` + `git push`，Render 會自動重新部署。
- **換 token / 改設定**：Render 後台 → Environment → 編輯後 Save（會自動重啟）。
- **費用**：Free 方案 $0；Worker 不計休眠。若未來需要更多資源再升級付費方案。
