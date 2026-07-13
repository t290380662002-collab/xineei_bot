# -*- coding: utf-8 -*-
"""
訂房 Telegram Bot
兩種輸入方式（自動判斷）：
  A. 貼上訂房文字格式（入住：/退房：/飯店：/房型：/件數：/備注：/是否吸煙：/入住者中文：…）
     -> 自動解析 -> 產生 Excel 回傳
  B. /start -> 逐步對話：飯店 -> 入住 -> 退房 -> 房型 -> 房數 -> 備注 -> 吸煙
     -> 逐位入住者(中文/英文/出生/證件) -> 可加多位 -> 產生 Excel 回傳
每筆都產生獨立的新 Excel 檔，直接回傳給使用者下載。
"""
import os
import logging
from telegram import Update, ReplyKeyboardRemove
from telegram.ext import (
    Application, CommandHandler, MessageHandler, ContextTypes,
    ConversationHandler, filters,
)
from fill import fill_booking, output_filename
from parse_text import parse_booking_text, looks_like_booking
from config import HOTEL_KEYS

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

HOTEL, CHECKIN, CHECKOUT, ROOMTYPE, ROOMS, REMARK, SMOKING, CN, EN, DOB, IDNO, MORE = range(12)


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    context.user_data["guests"] = []
    await update.message.reply_text(
        "開始新訂房 ✏️\n請輸入飯店名稱（" + " / ".join(HOTEL_KEYS) + "）：",
        reply_markup=ReplyKeyboardRemove())
    return HOTEL


async def hotel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["飯店"] = update.message.text.strip()
    await update.message.reply_text(
        f"飯店：{context.user_data['飯店']}\n請輸入入住日期（格式如 2026/07/20）：")
    return CHECKIN


async def checkin(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["入住"] = update.message.text.strip()
    await update.message.reply_text("請輸入退房日期（如 2026/07/22）：")
    return CHECKOUT


async def checkout(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["退房"] = update.message.text.strip()
    await update.message.reply_text("請輸入房型（可打代碼如 RK，或中文如 豪華大床房）：")
    return ROOMTYPE


async def roomtype(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["房型"] = update.message.text.strip()
    await update.message.reply_text("請輸入房數（件數，如 2）：")
    return ROOMS


async def rooms(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["件數"] = update.message.text.strip()
    await update.message.reply_text("備注（沒有的話請輸入 /skip，或有內容直接傳文字）：")
    return REMARK


async def remark(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["備注"] = update.message.text.strip()
    await update.message.reply_text("是否吸煙？（請填 吸煙 / 不吸煙）")
    return SMOKING


async def skip_remark(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["備注"] = ""
    await update.message.reply_text("是否吸煙？（請填 吸煙 / 不吸煙）")
    return SMOKING


async def smoking(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["是否吸煙"] = update.message.text.strip()
    await update.message.reply_text("填第一位入住者 ▶\n請輸入入住者中文姓名：")
    return CN


async def cn(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["_cn"] = update.message.text.strip()
    await update.message.reply_text("入住者英文姓名（格式如 QU,SHENZHONG）：")
    return EN


async def en(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["_en"] = update.message.text.strip()
    await update.message.reply_text("出生年月日（如 1961/06/11）：")
    return DOB


async def dob(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["_dob"] = update.message.text.strip()
    await update.message.reply_text("證件號碼：")
    return IDNO


async def idno(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["guests"].append({
        "cn_name": context.user_data.pop("_cn", ""),
        "en_name": context.user_data.pop("_en", ""),
        "dob": context.user_data.pop("_dob", ""),
        "idno": update.message.text.strip(),
    })
    n = len(context.user_data["guests"])
    await update.message.reply_text(
        f"✅ 已加入第 {n} 位入住者。\n輸入「完成」產生 Excel，或輸入「下一個」繼續新增入住者：")
    return MORE


async def more(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if "完成" in update.message.text or "✅" in update.message.text:
        return await finish(update, context)
    await update.message.reply_text("▶ 下一位入住者\n請輸入入住者中文姓名：")
    return CN


async def finish(update: Update, context: ContextTypes.DEFAULT_TYPE):
    booking = {k: context.user_data.get(k, "") for k in
               ["飯店", "入住", "退房", "房型", "件數", "備注", "是否吸煙"]}
    booking["guests"] = context.user_data.get("guests", [])
    if not booking["guests"]:
        await update.message.reply_text("還沒有入住者資料，無法產檔。請先填至少一位。")
        return MORE
    try:
        bio = fill_booking(booking)
        fn = output_filename(booking)
        await update.message.reply_document(
            document=bio, filename=fn,
            caption="✅ 訂房 Excel 已產生，請下載：")
    except Exception as e:
        logger.exception("產檔失敗")
        await update.message.reply_text(f"❌ 產檔失敗：{e}")
    context.user_data.clear()
    await update.message.reply_text(
        "如需再填一筆，請輸入 /start", reply_markup=ReplyKeyboardRemove())
    return ConversationHandler.END


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text(
        "已取消。輸入 /start 重新開始。", reply_markup=ReplyKeyboardRemove())
    return ConversationHandler.END


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
            f"請確認文字包含「飯店：/入住：/退房：/入住者中文：…」等欄位，\n"
            f"或直接輸入 /start 逐步填寫。",
            reply_markup=ReplyKeyboardRemove())
        return ConversationHandler.END
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
    await update.message.reply_text(
        "如需再填一筆，請輸入 /start 或再次貼上訂房文字。",
        reply_markup=ReplyKeyboardRemove())
    return ConversationHandler.END


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
    app = Application.builder().token(token).post_init(_clear_stale).build()
    conv = ConversationHandler(
        entry_points=[
            CommandHandler("start", start),
            MessageHandler(BookingTextFilter(), text_entry),
        ],
        states={
            # 每個步驟都先檢查：若這則訊息是「完整訂房文字」，優先自動讀取產檔，
            # 避免使用者貼整段時被誤當成逐步回答（對話狀態殘留也能正確處理）
            HOTEL: [MessageHandler(BookingTextFilter(), text_entry),
                    MessageHandler(filters.TEXT & ~filters.COMMAND, hotel)],
            CHECKIN: [MessageHandler(BookingTextFilter(), text_entry),
                      MessageHandler(filters.TEXT & ~filters.COMMAND, checkin)],
            CHECKOUT: [MessageHandler(BookingTextFilter(), text_entry),
                       MessageHandler(filters.TEXT & ~filters.COMMAND, checkout)],
            ROOMTYPE: [MessageHandler(BookingTextFilter(), text_entry),
                       MessageHandler(filters.TEXT & ~filters.COMMAND, roomtype)],
            ROOMS: [MessageHandler(BookingTextFilter(), text_entry),
                    MessageHandler(filters.TEXT & ~filters.COMMAND, rooms)],
            REMARK: [CommandHandler("skip", skip_remark),
                     MessageHandler(BookingTextFilter(), text_entry),
                     MessageHandler(filters.TEXT & ~filters.COMMAND, remark)],
            SMOKING: [MessageHandler(BookingTextFilter(), text_entry),
                      MessageHandler(filters.TEXT & ~filters.COMMAND, smoking)],
            CN: [MessageHandler(BookingTextFilter(), text_entry),
                 MessageHandler(filters.TEXT & ~filters.COMMAND, cn)],
            EN: [MessageHandler(BookingTextFilter(), text_entry),
                 MessageHandler(filters.TEXT & ~filters.COMMAND, en)],
            DOB: [MessageHandler(BookingTextFilter(), text_entry),
                  MessageHandler(filters.TEXT & ~filters.COMMAND, dob)],
            IDNO: [MessageHandler(BookingTextFilter(), text_entry),
                   MessageHandler(filters.TEXT & ~filters.COMMAND, idno)],
            MORE: [MessageHandler(BookingTextFilter(), text_entry),
                   MessageHandler(filters.TEXT & ~filters.COMMAND, more)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    app.add_handler(conv)
    logger.info("Bot 啟動中 (polling)...")
    _run_with_retry(app)


import asyncio
from telegram.error import Conflict


def _run_with_retry(app, max_retries=50):
    """啟動 polling；若遇到 409 衝突（有其他實例同時 polling），
    等 10 秒後重啟，直到成功或超過重試上限。這讓 Bot 在 Render 重新部署、
    新舊實例短暂並存時能自動恢復，不會永久卡死。"""
    asyncio.run(_poll_loop(app, max_retries))


async def _poll_loop(app, max_retries):
    await app.initialize()
    attempt = 0
    while attempt < max_retries:
        try:
            await app.start()
            await app.updater.start_polling(drop_pending_updates=True)
            logger.info("polling 已啟動，等待訊息…")
            # 持續等待，直到 updater 因停止或衝突而結束
            while app.updater.running:
                await asyncio.sleep(1)
            break  # 正常結束（收到停止信號）
        except Conflict:
            attempt += 1
            logger.warning(
                "偵測到 409 衝突（可能存在其他 Bot 實例），10 秒後重試 (%d/%d)…",
                attempt, max_retries)
            try:
                await app.updater.stop()
            except Exception:
                pass
            try:
                await app.stop()
            except Exception:
                pass
            await asyncio.sleep(10)
    else:
        logger.error("已達重試上限，Bot 停止。請檢查是否有多個實例同時運行。")
    try:
        await app.shutdown()
    except Exception:
        pass


if __name__ == "__main__":
    main()
