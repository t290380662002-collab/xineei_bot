# -*- coding: utf-8 -*-
"""
訂房 Telegram Bot
貼上訂房文字 → 自動解析 → 產生 Excel 回傳。
支援「查」指令查詢/切換賭廳（信威/博樂），預設信威。
"""
import os
import asyncio
import json
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

VERSION = "2026-07-25d"

WEBHOOK_PATH = "/webhook"
WEBHOOK_SECRET = os.environ.get("WEBHOOK_SECRET", "xinwea-booking-2026")

JUNKET_SETTINGS_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "junket_settings.json")


def load_junket_settings() -> dict:
    try:
        with open(JUNKET_SETTINGS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_junket_settings(settings: dict):
    try:
        with open(JUNKET_SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(settings, f, ensure_ascii=False, indent=2)
    except Exception:
        logger.warning("無法寫入 junket 設定檔", exc_info=True)


def _junket_keyboard(current: str):
    """建立賭廳切換按鈕，當前選中的加 ✓。"""
    xw = "信威 ✅" if current == "信威" else "信威"
    bl = "博樂 ✅" if current == "博樂" else "博樂"
    return InlineKeyboardMarkup([[
        InlineKeyboardButton(xw, callback_data="switch:信威"),
        InlineKeyboardButton(bl, callback_data="switch:博樂"),
    ]])


async def _produce_and_reply(target, booking: dict):
    """產出 Excel 並回傳。"""
    try:
        bio = fill_booking(booking)
        fn = output_filename(booking)
        await target.reply_document(
            document=bio, filename=fn,
            caption="✅ 訂房 Excel 請下載：")
        name_ok, name_warns = verify_booking_names(booking)
        age_ok, age_warns = verify_guests_age(booking)
        all_warns = name_warns + age_warns
        if all_warns:
            await target.reply_text("\n".join(all_warns))
    except Exception:
        logger.exception("產檔失敗")


def _webhook_base_url():
    return os.environ.get("WEBHOOK_URL") or os.environ.get("RENDER_EXTERNAL_URL") or None


class QueryFilter(filters.BaseFilter):
    """匹配 /查 或 /查@botname。"""
    def filter(self, update):
        msg = update.effective_message
        if not msg or not msg.text:
            return False
        t = msg.text.strip()
        return t == "/查" or t.startswith("/查@")


class BookingTextFilter(filters.BaseFilter):
    """判斷訊息是否為訂房文字。"""
    def filter(self, update):
        msg = update.effective_message
        if not msg or not msg.text:
            return False
        if msg.photo or msg.document:
            return False
        return looks_like_booking(msg.text)


async def text_entry(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """貼上訂房文字 → 直接產檔。"""
    booking = parse_booking_text(update.message.text)
    chat_id = str(update.effective_chat.id)
    settings = load_junket_settings()
    booking["junket"] = settings.get(chat_id, "信威")
    await _produce_and_reply(update.message, booking)


async def junket_query(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """「查」→ 顯示目前賭廳並提供切換按鈕。"""
    chat_id = str(update.effective_chat.id)
    current = load_junket_settings().get(chat_id, "信威")
    await update.message.reply_text(
        f"目前賭廳：{current}\n點擊下方按鈕可切換：",
        reply_markup=_junket_keyboard(current))


async def junket_switch(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """切換賭廳。"""
    query = update.callback_query
    await query.answer()
    data = (query.data or "")
    if not data.startswith("switch:"):
        return
    new_junket = data.split(":", 1)[1]
    chat_id = str(update.effective_chat.id)
    settings = load_junket_settings()
    settings[chat_id] = new_junket
    save_junket_settings(settings)
    await query.edit_message_text(
        f"✅ 已設定本群賭廳為：{new_junket}（已固定）",
        reply_markup=None)


async def fallback_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = update.effective_message
    logger.info("fallback: msg_id=%s", msg.message_id if msg else None)
    await msg.reply_text(
        "請直接貼上訂房資訊（入住/退房/飯店/房型/件數/入住者…）\n"
        "輸入「/查」可查詢/切換賭廳。")


def _build_application(token):
    app = Application.builder().token(token).build()
    # 點選賭廳切換按鈕
    app.add_handler(CallbackQueryHandler(junket_switch, pattern="^switch:"))
    # 「/查」→ 查詢/切換賭廳
    app.add_handler(MessageHandler(QueryFilter(), junket_query))
    # 貼訂房文字 → 產檔
    app.add_handler(MessageHandler(BookingTextFilter(), text_entry))
    # 其餘
    app.add_handler(MessageHandler(filters.ALL, fallback_handler))
    return app


async def _run_webhook_server(app, base):
    port = int(os.environ.get("PORT", 10000))
    url = base.rstrip("/") + WEBHOOK_PATH
    _ptb_ready = False

    async def handle_webhook(request):
        if WEBHOOK_SECRET and \
           request.headers.get("X-Telegram-Bot-Api-Secret-Token") != WEBHOOK_SECRET:
            return web.Response(status=403, text="forbidden")
        if not _ptb_ready:
            return web.Response(status=503, text="not ready")
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
        return web.json_response(
            {"status": "ok", "ptb": _ptb_ready, "version": VERSION})

    aio_app = web.Application()
    aio_app.router.add_post(WEBHOOK_PATH, handle_webhook)
    aio_app.router.add_get("/", handle_health)
    aio_app.router.add_get("/health", handle_health)

    runner = web.AppRunner(aio_app)
    await runner.setup()
    site = web.TCPSite(runner, "0.0.0.0", port)
    await site.start()
    logger.info("HTTP 啟動於 0.0.0.0:%s", port)

    try:
        await app.initialize()
        await app.start()
        await app.bot.set_webhook(url=url, secret_token=WEBHOOK_SECRET,
                                  drop_pending_updates=True)
        _ptb_ready = True
        logger.info("webhook 已設定：%s", url)
    except Exception:
        logger.exception("PTB 初始化失敗")

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
        logger.info("Bot 啟動 (webhook) -> %s%s", base.rstrip("/"), WEBHOOK_PATH)
        asyncio.run(_run_webhook_server(app, base))
    else:
        logger.info("Bot 啟動 (polling)")
        app.run_polling(drop_pending_updates=True, close_loop=False)


if __name__ == "__main__":
    main()
