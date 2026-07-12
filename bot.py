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
from telegram import Update, ReplyKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler, ContextTypes,
    ConversationHandler, filters,
)
from fill import fill_booking, output_filename
from parse_text import parse_booking_text, looks_like_booking

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

HOTEL, CHECKIN, CHECKOUT, ROOMTYPE, ROOMS, REMARK, SMOKING, CN, EN, DOB, IDNO, MORE = range(12)

HOTEL_KB = ReplyKeyboardMarkup(
    [["名匯", "威尼斯", "巴黎人", "倫敦人"]],
    one_time_keyboard=True, resize_keyboard=True,
)
MORE_KB = ReplyKeyboardMarkup(
    [["➕ 還有下一位", "✅ 完成，產生 Excel"]],
    resize_keyboard=True,
)


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    context.user_data["guests"] = []
    await update.message.reply_text(
        "開始新訂房 ✏️\n請選擇飯店：", reply_markup=HOTEL_KB)
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
        f"✅ 已加入第 {n} 位入住者。\n要再加下一位，還是完成產檔？",
        reply_markup=MORE_KB)
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
    await update.message.reply_text("如需再填一筆，請輸入 /start")
    return ConversationHandler.END


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    await update.message.reply_text("已取消。輸入 /start 重新開始。")
    return ConversationHandler.END


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
            f"或直接輸入 /start 逐步填寫。")
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
    await update.message.reply_text("如需再填一筆，請輸入 /start 或再次貼上訂房文字。")
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
    app = Application.builder().token(token).build()
    conv = ConversationHandler(
        entry_points=[
            CommandHandler("start", start),
            MessageHandler(BookingTextFilter(), text_entry),
        ],
        states={
            HOTEL: [MessageHandler(filters.TEXT & ~filters.COMMAND, hotel)],
            CHECKIN: [MessageHandler(filters.TEXT & ~filters.COMMAND, checkin)],
            CHECKOUT: [MessageHandler(filters.TEXT & ~filters.COMMAND, checkout)],
            ROOMTYPE: [MessageHandler(filters.TEXT & ~filters.COMMAND, roomtype)],
            ROOMS: [MessageHandler(filters.TEXT & ~filters.COMMAND, rooms)],
            REMARK: [CommandHandler("skip", skip_remark),
                     MessageHandler(filters.TEXT & ~filters.COMMAND, remark)],
            SMOKING: [MessageHandler(filters.TEXT & ~filters.COMMAND, smoking)],
            CN: [MessageHandler(filters.TEXT & ~filters.COMMAND, cn)],
            EN: [MessageHandler(filters.TEXT & ~filters.COMMAND, en)],
            DOB: [MessageHandler(filters.TEXT & ~filters.COMMAND, dob)],
            IDNO: [MessageHandler(filters.TEXT & ~filters.COMMAND, idno)],
            MORE: [MessageHandler(filters.TEXT & ~filters.COMMAND, more)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
    )
    app.add_handler(conv)
    logger.info("Bot 啟動中 (polling)...")
    app.run_polling()


if __name__ == "__main__":
    main()
