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
import json
import logging
from aiohttp import web
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, MessageHandler, ContextTypes,
    CallbackQueryHandler, ChatMemberHandler, filters,
)
from fill import fill_booking, output_filename, verify_booking_names, verify_guests_age
from parse_text import parse_booking_text, looks_like_booking

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

WEBHOOK_PATH = "/webhook"
WEBHOOK_SECRET = os.environ.get("WEBHOOK_SECRET", "xinwea-booking-2026")

# 每個 chat（群/私聊）固定一次的賭廳設定，寫入 JSON 以便重啟後仍記住。
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


def _junket_keyboard():
    return InlineKeyboardMarkup([[
        InlineKeyboardButton("信威", callback_data="junket:信威"),
        InlineKeyboardButton("博樂", callback_data="junket:博樂"),
    ]])


async def _produce_and_reply(target, booking: dict):
    """產出 Excel 並回傳（target 為 Message 物件）。"""
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
    except Exception as e:
        logger.exception("產檔失敗")
        await target.reply_text(f"⚠️ 產檔失敗：{e}")


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
    """直接貼上訂房文字 → 解析 → 若本群尚未選定賭廳則詢問一次，
    選定後固定，之後貼文字直接產檔。"""
    logger.info("text_entry: msg_id=%s text_len=%s",
                update.message.message_id if update.message else None,
                len(update.message.text or ""))
    booking = parse_booking_text(update.message.text)
    chat_id = update.effective_chat.id
    settings = load_junket_settings()
    key = str(chat_id)
    if key in settings:
        # 本群已固定賭廳 → 直接產檔，不再詢問
        booking["junket"] = settings[key]
        await _produce_and_reply(update.message, booking)
        return
    # 尚未設定：暫存訂房資料，彈出按鈕讓用戶選一次賭廳（選後即固定）
    context.chat_data["pending_booking"] = booking
    await update.message.reply_text(
        "本群尚未設定賭廳，請選擇（僅此一次，選後固定）：",
        reply_markup=_junket_keyboard())


async def junket_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """用戶點選按鈕後：寫入本群設定（固定），若有待處理訂房則產出 Excel。"""
    query = update.callback_query
    await query.answer()
    data = query.data or ""
    if not data.startswith("junket:"):
        return
    junket = data.split(":", 1)[1]
    chat_id = update.effective_chat.id
    key = str(chat_id)
    settings = load_junket_settings()
    settings[key] = junket          # 固定到本群（僅一次）
    save_junket_settings(settings)
    booking = context.chat_data.get("pending_booking")
    if booking:
        booking["junket"] = junket
        context.chat_data.pop("pending_booking", None)
        await _produce_and_reply(query.message, booking)
    try:
        await query.edit_message_text(
            f"✅ 已設定本群賭廳為：{junket}（僅此一次，之後貼訂房文字將自動套用）",
            reply_markup=None)
    except Exception:
        pass


async def chat_member_added(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """機器人被加入群組時，詢問一次賭廳（已設定過則不重問）。"""
    member = update.my_chat_member
    if not member:
        return
    chat = update.effective_chat
    if not chat or chat.type not in ("group", "supergroup"):
        return
    if member.new_chat_member.status != "member":
        return
    key = str(chat.id)
    settings = load_junket_settings()
    if key in settings:
        return  # 已固定，不重問
    await context.bot.send_message(
        chat_id=chat.id,
        text="歡迎使用訂房機器人！\n請選擇本群所屬賭廳（僅此一次，選後固定）：",
        reply_markup=_junket_keyboard())


async def fallback_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """任何未被 text_entry 處理的訊息，給一個指引。"""
    msg = update.effective_message
    logger.info("fallback_handler: msg_id=%s has_photo=%s has_doc=%s text_len=%s",
                msg.message_id if msg else None,
                bool(msg.photo) if msg else None,
                bool(msg.document) if msg else None,
                len(msg.text or "") if msg and msg.text else 0)
    await update.effective_message.reply_text(
        "我看不懂這則訊息喔！\n"
        "請直接貼上訂房資訊（入住/退房/飯店/房型/件數/入住者…）")


def _build_application(token):
    """建立 Application：每群首次按鈕選賭廳並固定，之後直接產檔。"""
    app = Application.builder().token(token).build()
    text_filter = BookingTextFilter()
    # 機器人被加入群組 → 提示選一次賭廳
    app.add_handler(ChatMemberHandler(
        chat_member_added, ChatMemberHandler.MY_CHAT_MEMBER))
    # 點選賭廳按鈕
    app.add_handler(CallbackQueryHandler(junket_callback, pattern="^junket:"))
    # 貼訂房文字（已固定��直接產檔，否則彈按鈕）
    app.add_handler(MessageHandler(text_filter, text_entry))
    # 其餘訊息：指引
    app.add_handler(MessageHandler(filters.ALL, fallback_handler))
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
