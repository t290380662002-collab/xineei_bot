# -*- coding: utf-8 -*-
"""
訂房 Telegram Bot
輸入方式（模式 A）：
  貼上訂房文字格式（入住：/退房：/飯店：/房型：/件數：/備注：/是否吸煙：/入住者中文：…）
  -> 自動解析 -> 產生 Excel 回傳
每筆都產生獨立的新 Excel 檔，直接回傳給使用者下載。
（模式 B：/start 逐步對話 已移除，僅保留模式 A。）
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


async def _clear_stale(app):
    """啟動時先清掉任何殘留的 webhook / 舊 polling 實例，避免 Render 重啟時
    新舊兩隻同時 polling 造成的 409 Conflict（terminated by other getUpdates）。"""
    try:
        await app.bot.delete_webhook(drop_pending_updates=True)
        logger.info("已清除可能的殘留 webhook / 舊 polling 實例")
    except Exception as e:
        logger.warning("清除殘留實例時發生錯誤（可忽略）：%s", e)


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
    missing = []
    if not booking.get("飯店"):
        missing.append("飯店")
    if not booking.get("入住"):
        missing.append("入住")
    if not booking.get("退房"):
        missing.append("退房")
    if not booking.get("guests"):
        missing.append("入住者資料(中文/英文/出生/證件)")
    if missing:
        await update.message.reply_text(
            f"✅ 已讀取文字，但缺少必要欄位：{', '.join(missing)}\n"
            f"請確認文字包含「飯店：/入住：/退房：/入住者中文：…」等欄位。",
            reply_markup=ReplyKeyboardRemove())
        return
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
    """建立 Application：僅保留「貼文字自動產檔」模式（模式 A）。
    模式 B（/start 逐步對話）已移除；/start 僅提示直接貼文字。"""
    app = Application.builder().token(token).post_init(_clear_stale).build()

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
    logger.info("Bot 啟動中 (polling)...")
    # 使用官方 run_polling：內建處理 SIGTERM/SIGINT，
    # 部署時舊實例會乾淨退出，避免留下仍佔用 polling 的殭屍程序
    # （這正是長期 409 Conflict「機器人沒反應」的主因）。
    # 短暫的 409（新舊實例並存）會由 PTB 內部自動重試恢復。
    app.run_polling(drop_pending_updates=True, close_loop=False)


if __name__ == "__main__":
    main()
