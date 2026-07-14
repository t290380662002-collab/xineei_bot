# -*- coding: utf-8 -*-
"""
訂房 Telegram Bot
輸入方式（模式 A）：
  貼上訂房文字格式（入住：/退房：/飯店：/房型：/件數：/備注：/是否吸煙：/入住者中文：…）
  -> 自動解析 -> 產生 Excel 回傳
每筆都產生獨立的新 Excel 檔，直接回傳給使用者下載。

運作模式：
  - 若環境有 WEBHOOK_URL 或 RENDER_EXTERNAL_URL（Render Web Service 自動注入），
    則啟用 Webhook 模式：自建 aiohttp 伺服器同時提供
      · POST /webhook  -> 接收 Telegram 推播
      · GET  /         -> 健康檢查（回 200），讓 Render 部署不被判失敗
    「同一隻 Bot 只能有一組 webhook」，因此從根本杜絕 polling 的 409 多實例互踢。
  - 否則退回 Polling 模式（本機開發用）。
"""
import os
import asyncio
import logging
from aiohttp import web
from telegram import Update
from telegram.ext import (
    Application, MessageHandler, ContextTypes,
    filters,
)
from fill import fill_booking, output_filename, verify_booking_names
from parse_text import parse_booking_text, looks_like_booking
import ocr

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

WEBHOOK_PATH = "/webhook"
WEBHOOK_SECRET = os.environ.get("WEBHOOK_SECRET", "xinwea-booking-2026")


def _webhook_base_url():
    """回傳 webhook 基底網址（不含路徑）；無則回 None。"""
    return os.environ.get("WEBHOOK_URL") or os.environ.get("RENDER_EXTERNAL_URL") or None


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
        # 中文 / 英文姓名拼音自動核對：不符時提示使用者確認
        ok, warns = verify_booking_names(booking)
        if not ok:
            await update.message.reply_text("\n".join(warns))
    except Exception:
        logger.exception("文字產檔失敗")
    # 暫存原始文字，讓使用者之後單獨傳證件照片也能核對
    context.user_data["last_text"] = update.message.text
    return


async def _download_photo_source(update):
    """從 photo 或 image document 取圖片 bytes；不支援則回 (None, reason)。"""
    msg = update.effective_message
    src = None
    if msg.photo:
        src = msg.photo[-1]
    elif getattr(msg.document, "mime_type", None) and \
            msg.document.mime_type.startswith("image/"):
        src = msg.document
    if src is None:
        return None, "not_image"
    try:
        data = await src.get_file().download_as_bytearray()
    except Exception:
        logger.exception("下載圖片失敗")
        return None, "download_fail"
    if not data:
        return None, "empty"
    return bytes(data), None


async def photo_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """收到證件照片：OCR 識別，並與訂房文字核對（或僅回顯識別結果）。"""
    data, err = await _download_photo_source(update)
    if err == "not_image":
        await update.effective_message.reply_text(
            "【注意】我看得懂「訂房文字」或直接傳送的「證件照片」。\n"
            "請直接傳送照片（不要用「檔案」方式傳送）；若以檔案傳送，請確認是圖片格式。")
        return
    if err in ("download_fail", "empty"):
        await update.effective_message.reply_text(
            "【注意】證件照片下載失敗，請重試一次，或改為手動輸入。")
        return

    try:
        raw_text = ocr.ocr_image_bytes(data)
    except ocr.TesseractNotInstalled:
        await update.effective_message.reply_text(
            "【注意】證件 OCR 功能需要伺服器安裝 Tesseract，請聯絡管理員設定。")
        return
    except Exception:
        logger.exception("OCR 失敗")
        await update.effective_message.reply_text(
            "【注意】證件照片辨識失敗，請確認照片清晰或改為手動輸入。")
        return

    fields = ocr.extract_fields(raw_text)
    logger.info("OCR 識別欄位：%s | 原文長度 %d", fields, len(raw_text or ""))

    # 文字來源：圖片 caption 優先，否則用上一次貼的訂房文字
    caption = update.message.caption or ""
    booking_text = caption if caption else context.user_data.get("last_text", "")

    if booking_text and looks_like_booking(booking_text):
        booking = parse_booking_text(booking_text)
        try:
            bio = fill_booking(booking)
            fn = output_filename(booking)
            await update.message.reply_document(
                document=bio, filename=fn,
                caption="已從文字自動填入，訂房 Excel 請下載：")
            ok, vwarns = verify_booking_names(booking)
            if not ok:
                await update.message.reply_text("\n".join(vwarns))
        except Exception:
            logger.exception("圖片訊息產檔失敗")
        warns = ocr.verify_ocr_vs_booking(fields, booking)
        if warns:
            await update.message.reply_text("\n".join(warns))
        else:
            await update.message.reply_text("證件資料與填寫內容一致。")
    else:
        await update.message.reply_text(ocr.format_fields(fields))


async def fallback_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """任何未被 text/photo 處理的訊息，給一個指引，避免靜默無反應。"""
    await update.effective_message.reply_text(
        "我看得懂兩種訊息：\n"
        "1. 直接貼上訂房文字（入住/退房/飯店/房型/件數…）\n"
        "2. 傳送證件照片（請用「照片」方式，不要用「檔案」）")


def _build_application(token):
    """建立 Application：僅保留「貼文字自動產檔」模式（模式 A）。
    不註冊任何指令，Telegram 不會顯示指令選單按鈕。"""
    app = Application.builder().token(token).build()
    text_filter = BookingTextFilter()
    photo_filter = filters.PHOTO | filters.Document.IMAGE
    app.add_handler(MessageHandler(text_filter, text_entry))
    app.add_handler(MessageHandler(photo_filter, photo_handler))
    # 兜底：只有非文字、非照片/图片檔案时才提示，避免误触发
    app.add_handler(MessageHandler(
        filters.ChatType.PRIVATE & ~text_filter & ~photo_filter,
        fallback_handler))
    return app


async def _run_webhook_server(app, base):
    """自建 aiohttp 伺服器：同時處理 Telegram webhook 與健康檢查。"""
    port = int(os.environ.get("PORT", 10000))
    url = base.rstrip("/") + WEBHOOK_PATH

    # 設定 webhook（同一隻 Bot 只能有一組，因此不會有多實例互踢）
    await app.bot.set_webhook(url=url, secret_token=WEBHOOK_SECRET,
                              drop_pending_updates=True)
    await app.initialize()
    await app.start()
    logger.info("webhook 已設定：%s", url)

    async def handle_webhook(request):
        if WEBHOOK_SECRET and \
           request.headers.get("X-Telegram-Bot-Api-Secret-Token") != WEBHOOK_SECRET:
            return web.Response(status=403, text="forbidden")
        try:
            data = await request.json()
        except Exception:
            return web.Response(status=400, text="bad json")
        update = Update.de_json(data, app.bot)
        try:
            await app.process_update(update)
        except Exception:
            logger.exception("process_update 失敗")
        return web.Response(text="ok")

    async def handle_health(request):
        status = {"status": "ok"}
        try:
            status["ocr"] = "ready" if ocr.tesseract_ready() else "not_installed"
        except Exception as e:  # noqa
            status["ocr"] = f"error:{e}"
        return web.json_response(status)

    aio_app = web.Application()
    aio_app.router.add_post(WEBHOOK_PATH, handle_webhook)
    aio_app.router.add_get("/", handle_health)
    aio_app.router.add_get("/health", handle_health)

    runner = web.AppRunner(aio_app)
    await runner.setup()
    site = web.TCPSite(runner, "0.0.0.0", port)
    await site.start()
    logger.info("HTTP 服務啟動於 0.0.0.0:%s (health=%s, webhook=%s)", port, "/", WEBHOOK_PATH)

    try:
        while True:
            await asyncio.sleep(3600)
    finally:
        try:
            await app.bot.delete_webhook()
        except Exception:
            pass
        await app.stop()
        await app.shutdown()
        await runner.cleanup()


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
    logger.info("啟動診斷: RENDER_EXTERNAL_URL=%s WEBHOOK_URL=%s PORT=%s",
                os.environ.get("RENDER_EXTERNAL_URL"),
                os.environ.get("WEBHOOK_URL"),
                os.environ.get("PORT"))

    if base:
        # ---- Webhook 模式（Render Web Service）----
        logger.info("Bot 啟動中 (webhook) -> %s%s | PORT=%s",
                    base.rstrip("/"), WEBHOOK_PATH, os.environ.get("PORT"))
        try:
            asyncio.run(_run_webhook_server(app, base))
        except Exception as e:
            logger.exception("webhook 服務啟動失敗：%s", e)
            raise
    else:
        # ---- Polling 模式（本機開發備援）----
        logger.info("Bot 啟動中 (polling)...")
        app.run_polling(drop_pending_updates=True, close_loop=False)


if __name__ == "__main__":
    main()
