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
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, MessageHandler, ContextTypes,
    CallbackQueryHandler, filters,
)
from fill import fill_booking, output_filename, verify_booking_names, verify_guests_age
from parse_text import parse_booking_text, looks_like_booking

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

WEBHOOK_PATH = "/webhook"
WEBHOOK_SECRET = os.environ.get("WEBHOOK_SECRET", "xinwea-booking-2026")


def _webhook_base_url():
    """回傳 webhook 基底網址（不含路徑）；無則回 None。"""
    return os.environ.get("WEBHOOK_URL") or os.environ.get("RENDER_EXTERNAL_URL") or None


class BookingTextFilter(filters.BaseFilter):
    """判斷訊息是否為「訂房文字格式」（含至少 4 個可辨識欄位）。
    顯式排除照片/文件，避免 photo/document 訊息被誤判為文字。"""

    def filter(self, update):
        msg = update.effective_message
        if not msg:
            return False
        if msg.photo or msg.document:
            return False
        if not msg.text:
            return False
        return looks_like_booking(msg.text)


async def text_entry(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """直接貼上訂房文字 → 解析 → 詢問賭廳 → 產生 Excel 回傳。"""
    logger.info("text_entry: msg_id=%s text_len=%s",
                update.message.message_id if update.message else None,
                len(update.message.text or ""))
    booking = parse_booking_text(update.message.text)
    # 先記住解析結果，等待用戶選擇賭廳後才產檔
    context.user_data["pending_booking"] = booking
    g0 = (booking.get("guests") or [{}])[0]
    name = g0.get("cn_name") or g0.get("en_name") or ""
    hotel = booking.get("飯店", "")
    checkin = booking.get("入住", "")
    checkout = booking.get("退房", "")
    summary = f"飯店：{hotel}\n姓名：{name}\n入住：{checkin} 退房：{checkout}\n\n請選擇賭廳："
    keyboard = [
        [
            InlineKeyboardButton("信威-澳門廳", callback_data="junket:信威"),
            InlineKeyboardButton("博樂-澳門廳", callback_data="junket:博樂"),
        ]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    sent = await update.message.reply_text(summary, reply_markup=reply_markup)
    context.user_data["pending_msg_id"] = sent.message_id


async def junket_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """用戶點選賭廳後，設定 junket 並產出 Excel。"""
    query = update.callback_query
    await query.answer()
    data = query.data or ""
    if not data.startswith("junket:"):
        return
    junket = data.split(":", 1)[1]
    booking = context.user_data.get("pending_booking")
    if not booking:
        await query.edit_message_text("⚠️ 找不到待處理的訂房資料，請重新貼上訂房文字。")
        return
    booking["junket"] = junket
    # 移除鍵盤，顯示已選擇
    await query.edit_message_reply_markup(reply_markup=None)
    try:
        bio = fill_booking(booking)
        fn = output_filename(booking)
        await query.message.reply_document(
            document=bio,
            filename=fn,
            caption=f"✅ 已選擇賭廳：{junket}，訂房 Excel 請下載：",
        )
        # 中文 / 英文姓名拼音自動核對：不符時提示使用者確認
        name_ok, name_warns = verify_booking_names(booking)
        # 年齡檢查：未滿 21 歲提示
        age_ok, age_warns = verify_guests_age(booking)
        all_warns = name_warns + age_warns
        if all_warns:
            await query.message.reply_text("\n".join(all_warns))
    except Exception:
        logger.exception("產檔失敗")
        await query.message.reply_text("⚠️ 產檔失敗，請檢查訂房文字格式或稍後再試。")
    finally:
        context.user_data.pop("pending_booking", None)
        context.user_data.pop("pending_msg_id", None)


async def fallback_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """任何未被 text_entry 處理的訊息，給一個指引。"""
    msg = update.effective_message
    logger.info("fallback_handler: msg_id=%s has_photo=%s has_doc=%s text_len=%s",
                msg.message_id if msg else None,
                bool(msg.photo) if msg else None,
                bool(msg.document) if msg else None,
                len(msg.text or "") if msg and msg.text else 0)
    await update.effective_message.reply_text(
        "我只看得懂「訂房文字」格式喔！\n"
        "請直接貼上訂房資訊（入住/退房/飯店/房型/件數/入住者…）")


def _build_application(token):
    """建立 Application：貼文字後先詢問賭廳，再產檔。"""
    app = Application.builder().token(token).build()
    text_filter = BookingTextFilter()
    app.add_handler(CallbackQueryHandler(junket_callback, pattern="^junket:"))
    app.add_handler(MessageHandler(text_filter, text_entry))
    app.add_handler(MessageHandler(
        filters.ChatType.PRIVATE & ~text_filter,
        fallback_handler))
    return app


async def _run_webhook_server(app, base):
    """自建 aiohttp 伺服器：同時處理 Telegram webhook 與健康檢查。
    啟動順序：HTTP 伺服器 → PTB 初始化 → webhook 設定。
    確保健康檢查端點最先可用，避免 Render 判部署失敗。"""
    port = int(os.environ.get("PORT", 10000))
    url = base.rstrip("/") + WEBHOOK_PATH

    # --- 先定義 handler，再啟動 HTTP 伺服器 ---
    _ptb_ready = False  # 標記 PTB 是否已初始化完成

    async def handle_webhook(request):
        if WEBHOOK_SECRET and \
           request.headers.get("X-Telegram-Bot-Api-Secret-Token") != WEBHOOK_SECRET:
            return web.Response(status=403, text="forbidden")
        if not _ptb_ready:
            logger.warning("webhook 收到請求但 PTB 尚未就緒")
            return web.Response(status=503, text="not ready")
        try:
            data = await request.json()
        except Exception:
            return web.Response(status=400, text="bad json")
        update = Update.de_json(data, app.bot)
        msg_type = "unknown"
        if update.message:
            if update.message.photo:
                msg_type = "photo"
            elif update.message.document:
                msg_type = f"document:{update.message.document.mime_type or 'unknown'}"
            elif update.message.text:
                msg_type = "text"
            else:
                msg_type = "other"
        logger.info("handle_webhook: update_id=%s msg_id=%s type=%s",
                    update.update_id,
                    update.message.message_id if update.message else None,
                    msg_type)
        try:
            await app.process_update(update)
        except Exception:
            logger.exception("process_update 失敗")
        return web.Response(text="ok")

    async def handle_health(request):
        return web.json_response({"status": "ok", "ptb": _ptb_ready})

    aio_app = web.Application()
    aio_app.router.add_post(WEBHOOK_PATH, handle_webhook)
    aio_app.router.add_get("/", handle_health)
    aio_app.router.add_get("/health", handle_health)

    runner = web.AppRunner(aio_app)
    await runner.setup()
    site = web.TCPSite(runner, "0.0.0.0", port)
    await site.start()
    logger.info("HTTP 服務啟動於 0.0.0.0:%s (health=/, webhook=%s)", port, WEBHOOK_PATH)

    # --- HTTP 伺服器已啟動，現在初始化 PTB ---
    try:
        await app.initialize()
        await app.start()
        await app.bot.set_webhook(url=url, secret_token=WEBHOOK_SECRET,
                                  drop_pending_updates=True)
        _ptb_ready = True
        logger.info("PTB 初始化完成，webhook 已設定：%s", url)
    except Exception:
        logger.exception("PTB 初始化或 webhook 設定失敗")

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
