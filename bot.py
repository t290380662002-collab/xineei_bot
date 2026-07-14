# -*- coding: utf-8 -*-
"""
訂房 Telegram Bot
輸入方式（模式 A）：
  貼上訂房文字格式（入住：/退房：/飯店：/房型：/件數：/備注：/是否吸煙：/入住者中文：…）
  -> 自動解析 -> 產生 Excel 回傳
每筆都產生獨立的新 Excel 檔，直接回傳給使用者下載。
（模式 B：/start 逐步對話 已移除，僅保留模式 A。）

運作模式（自動判斷）：
  - 若環境有 WEBHOOK_URL 或 RENDER_EXTERNAL_URL（Render Web Service 自動注入），
    則使用 Webhook 模式：Telegram 主動把訊息推播到本服務網址，
    「同一隻 Bot 只能有一組 webhook」，因此從根本杜絕 polling 的 409 多實例互踢。
  - 否則退回 Polling 模式（本機開發用）。
"""
import os
import logging
from telegram import Update, ReplyKeyboardRemove
from telegram.ext import (
    Application, CommandHandler, MessageHandler, ContextTypes,
    filters,
)
from fill import fill_booking, output_filename
from parse_text import parse_booking_text, looks_like_booking

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

WEBHOOK_PATH = "/webhook"
WEBHOOK_SECRET = os.environ.get("WEBHOOK_SECRET", "xinwea-booking-2026")


def _webhook_base_url():
    """回傳 webhook 基底網址（不含路徑）；無則回 None。"""
    return os.environ.get("WEBHOOK_URL") or os.environ.get("RENDER_EXTERNAL_URL") or None


async def _set_webhook(app):
    base = _webhook_base_url()
    if not base:
        logger.error("未偵測到 WEBHOOK_URL / RENDER_EXTERNAL_URL，無法設定 webhook")
        return
    url = base.rstrip("/") + WEBHOOK_PATH
    await app.bot.set_webhook(url=url, secret_token=WEBHOOK_SECRET,
                              drop_pending_updates=True)
    logger.info("webhook 已設定：%s", url)


async def _clear_webhook(app):
    try:
        await app.bot.delete_webhook(drop_pending_updates=True)
        logger.info("已清除 webhook / 舊 polling 實例")
    except Exception as e:
        logger.warning("清除時發生錯誤（可忽略）：%s", e)


class BookingTextFilter(filters.BaseFilter):
    """判斷訊息是否為「訂房文字格式」（含至少 4 個可辨識欄位）。"""

    def filter(self, update):
        msg = update.effective_message
        if not msg or not msg.text:
            return False
        return looks_like_booking(msg.text)


async def text_entry(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """直接貼上訂房文字 → 解析 → 產生 Excel 回傳。"""
    booking = parse_booking_text(update.message.text)
    try:
        bio = fill_booking(booking)
        fn = output_filename(booking)
        await update.message.reply_document(
            document=bio, filename=fn,
            caption="✅ 已從文字自動填入，訂房 Excel 請下載：")
    except Exception as e:
        logger.exception("文字產檔失敗")
        await update.message.reply_text(f"❌ 產檔失敗：{e}")
    context.user_data.clear()
    return


def _build_application(token):
    """建立 Application：僅保留「貼文字自動產檔」模式（模式 A）。"""
    app = Application.builder().token(token).build()

    async def start_hint(update: Update, context: ContextTypes.DEFAULT_TYPE):
        await update.message.reply_text(
            "請直接貼上訂房文字（含 飯店 / 入住 / 退房 / 房型 / 件數 / 入住者…），"
            "我會自動填入並產生 Excel 回傳。")

    app.add_handler(CommandHandler("start", start_hint))
    app.add_handler(MessageHandler(BookingTextFilter(), text_entry))
    return app


def main():
    # 本機開發時從 .env 讀取 token；正式部署請用環境變數設定
    try:
        from dotenv import load_dotenv
        load_dotenv()
    except Exception:
        pass
    token = os.environ.get("TELEGRAM_BOT_TOKEN")
    if not token:
        raise SystemExit("請先設定環境變數 TELEGRAM_BOT_TOKEN")

    app = _build_application(token)
    base = _webhook_base_url()

    if base:
        # ---- Webhook 模式（Render Web Service）----
        app.post_init = _set_webhook
        app.post_stop = _clear_webhook
        port = int(os.environ.get("PORT", 10000))
        logger.info("Bot 啟動中 (webhook) -> %s%s", base.rstrip("/"), WEBHOOK_PATH)
        app.run_webhook(
            listen="0.0.0.0",
            port=port,
            url_path=WEBHOOK_PATH,
            secret_token=WEBHOOK_SECRET,
            drop_pending_updates=True,
        )
    else:
        # ---- Polling 模式（本機開發備援）----
        app.post_init = _clear_webhook
        logger.info("Bot 啟動中 (polling)...")
        app.run_polling(drop_pending_updates=True, close_loop=False)


if __name__ == "__main__":
    main()
