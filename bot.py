import os
import logging
from datetime import date, datetime
from pathlib import Path

from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook
from telegram import (
    Update, ReplyKeyboardMarkup, KeyboardButton,
    InlineKeyboardButton, InlineKeyboardMarkup,
)
from telegram.ext import (
    Application, CommandHandler, MessageHandler, CallbackQueryHandler,
    ConversationHandler, ContextTypes, filters,
)

load_dotenv()
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()
ADMIN_ID_RAW = os.getenv("ADMIN_ID", "").strip()
if not BOT_TOKEN:
    raise RuntimeError("Environment variable BOT_TOKEN is not set")
if not ADMIN_ID_RAW.isdigit():
    raise RuntimeError("ADMIN_ID must be your numeric Telegram user ID")
ADMIN_ID = int(ADMIN_ID_RAW)

SELLER_IDS_RAW = os.getenv("ALLOWED_SELLER_IDS", "").strip()
ALLOWED_SELLER_IDS = {
    int(value.strip()) for value in SELLER_IDS_RAW.split(",")
    if value.strip().isdigit()
}
# If no seller IDs are configured, any Telegram user can submit an order.
# Admin-only actions remain restricted to ADMIN_ID.

DATA_FILE = Path(os.getenv("ORDERS_FILE", "orders.xlsx"))
HEADERS = [
    "ID", "Дата заказа", "Номер ТТ", "Адрес ТТ", "Серийный номер SIM",
    "Номер телефона SIM", "Сумма заказа", "Продавец", "Telegram ID продавца",
    "Создано",
]

DATE, OUTLET_NUMBER, OUTLET_ADDRESS, SIM_SERIAL, SIM_PHONE, AMOUNT, SELLER = range(7)
CANCEL_TEXT = "❌ Отмена"
ADMIN_MENU_TEXT = "🛠 Панель администратора"
NEW_ORDER_TEXT = "➕ Добавить заказ"


def ensure_workbook():
    if not DATA_FILE.exists():
        wb = Workbook()
        ws = wb.active
        ws.title = "Заказы"
        ws.append(HEADERS)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = "A1:J1"
        wb.save(DATA_FILE)


def get_sheet():
    ensure_workbook()
    wb = load_workbook(DATA_FILE)
    ws = wb["Заказы"] if "Заказы" in wb.sheetnames else wb.active
    return wb, ws


def is_admin(user_id: int) -> bool:
    return user_id == ADMIN_ID


def can_submit(user_id: int) -> bool:
    return is_admin(user_id) or not ALLOWED_SELLER_IDS or user_id in ALLOWED_SELLER_IDS


def main_keyboard(user_id: int):
    rows = [[KeyboardButton(NEW_ORDER_TEXT)]]
    if is_admin(user_id):
        rows.append([KeyboardButton(ADMIN_MENU_TEXT)])
    return ReplyKeyboardMarkup(rows, resize_keyboard=True)


def admin_inline_keyboard():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("📋 Последние заказы", callback_data="admin_recent")],
        [InlineKeyboardButton("📤 Скачать Excel", callback_data="admin_export")],
        [InlineKeyboardButton("🧹 Очистить все заказы", callback_data="admin_clear_ask")],
    ])


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    if not can_submit(user.id):
        await update.message.reply_text("У вас нет доступа к этому боту. Обратитесь к администратору.")
        return
    text = "Привет! Здесь можно зарегистрировать заказ SIM-карты."
    if is_admin(user.id):
        text += "\nУ тебя также есть панель администратора."
    await update.message.reply_text(text, reply_markup=main_keyboard(user.id))


async def admin_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("Эта панель доступна только администратору.")
        return
    await update.message.reply_text("🛠 Панель администратора\nВыбери действие:", reply_markup=admin_inline_keyboard())


async def order_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not can_submit(update.effective_user.id):
        await update.message.reply_text("У вас нет доступа к добавлению заказов.")
        return ConversationHandler.END
    context.user_data["order"] = {"date": date.today().strftime("%d.%m.%Y")}
    keyboard = ReplyKeyboardMarkup(
        [[KeyboardButton("Сегодня"), KeyboardButton("Изменить дату")], [KeyboardButton(CANCEL_TEXT)]],
        resize_keyboard=True,
        one_time_keyboard=True,
    )
    await update.message.reply_text(
        f"Дата заказа: {context.user_data['order']['date']}\nНажми «Сегодня» или измени дату в формате ДД.ММ.ГГГГ.",
        reply_markup=keyboard,
    )
    return DATE


async def receive_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    value = update.message.text.strip()
    if value == CANCEL_TEXT:
        return await cancel(update, context)
    if value == "Сегодня":
        context.user_data["order"]["date"] = date.today().strftime("%d.%m.%Y")
    elif value == "Изменить дату":
        await update.message.reply_text("Введи дату в формате ДД.ММ.ГГГГ, например 20.10.2026.")
        return DATE
    else:
        try:
            parsed = datetime.strptime(value, "%d.%m.%Y").date()
            context.user_data["order"]["date"] = parsed.strftime("%d.%m.%Y")
        except ValueError:
            await update.message.reply_text("Не понял дату. Введи её в формате ДД.ММ.ГГГГ или нажми «Сегодня».")
            return DATE
    await update.message.reply_text("Введи номер торговой точки (ТТ).", reply_markup=ReplyKeyboardMarkup([[KeyboardButton(CANCEL_TEXT)]], resize_keyboard=True))
    return OUTLET_NUMBER


async def receive_outlet_number(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.message.text.strip() == CANCEL_TEXT:
        return await cancel(update, context)
    context.user_data["order"]["outlet_number"] = update.message.text.strip()
    await update.message.reply_text("Введи адрес торговой точки.")
    return OUTLET_ADDRESS


async def receive_outlet_address(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.message.text.strip() == CANCEL_TEXT:
        return await cancel(update, context)
    context.user_data["order"]["outlet_address"] = update.message.text.strip()
    await update.message.reply_text("Введи серийный номер SIM-карты.")
    return SIM_SERIAL


async def receive_sim_serial(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.message.text.strip() == CANCEL_TEXT:
        return await cancel(update, context)
    context.user_data["order"]["sim_serial"] = update.message.text.strip()
    await update.message.reply_text("Введи номер телефона SIM-карты.")
    return SIM_PHONE


async def receive_sim_phone(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.message.text.strip() == CANCEL_TEXT:
        return await cancel(update, context)
    context.user_data["order"]["sim_phone"] = update.message.text.strip()
    await update.message.reply_text("Введи сумму заказа (только число или сумму с валютой).")
    return AMOUNT


async def receive_amount(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.message.text.strip() == CANCEL_TEXT:
        return await cancel(update, context)
    context.user_data["order"]["amount"] = update.message.text.strip()
    await update.message.reply_text("Введи имя продавца.")
    return SELLER


async def receive_seller(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.message.text.strip() == CANCEL_TEXT:
        return await cancel(update, context)
    order = context.user_data["order"]
    order["seller"] = update.message.text.strip()
    user = update.effective_user
    try:
        wb, ws = get_sheet()
        order_id = max(1, ws.max_row)
        ws.append([
            order_id, order["date"], order["outlet_number"], order["outlet_address"],
            order["sim_serial"], order["sim_phone"], order["amount"], order["seller"],
            user.id, datetime.now().strftime("%d.%m.%Y %H:%M:%S"),
        ])
        wb.save(DATA_FILE)
    except Exception:
        logger.exception("Could not save order")
        await update.message.reply_text("Не удалось сохранить заказ. Попробуй позже или сообщи администратору.", reply_markup=main_keyboard(user.id))
        context.user_data.pop("order", None)
        return ConversationHandler.END

    summary = (
        "✅ Заказ сохранён!\n\n"
        f"Дата: {order['date']}\nТТ: {order['outlet_number']}\nАдрес: {order['outlet_address']}\n"
        f"Серийный номер SIM: {order['sim_serial']}\nТелефон SIM: {order['sim_phone']}\n"
        f"Сумма: {order['amount']}\nПродавец: {order['seller']}"
    )
    await update.message.reply_text(summary, reply_markup=main_keyboard(user.id))
    context.user_data.pop("order", None)
    return ConversationHandler.END


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.pop("order", None)
    await update.message.reply_text("Добавление заказа отменено.", reply_markup=main_keyboard(update.effective_user.id))
    return ConversationHandler.END


async def admin_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if not is_admin(query.from_user.id):
        await query.edit_message_text("Недостаточно прав.")
        return

    if query.data == "admin_clear_ask":
        keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton("⚠️ Да, удалить все заказы", callback_data="admin_clear_confirm")],
            [InlineKeyboardButton("↩️ Отмена", callback_data="admin_back")],
        ])
        await query.edit_message_text(
            "⚠️ Ты точно хочешь удалить ВСЕ заказы из Excel-файла?\n\n"
            "Это действие нельзя отменить. Заголовки столбцов останутся.",
            reply_markup=keyboard,
        )
    elif query.data == "admin_clear_confirm":
        try:
            wb, ws = get_sheet()
            if ws.max_row > 1:
                ws.delete_rows(2, ws.max_row - 1)
            ws.auto_filter.ref = "A1:J1"
            wb.save(DATA_FILE)
            await query.edit_message_text("🧹 Все заказы удалены. Заголовки таблицы сохранены.", reply_markup=admin_inline_keyboard())
        except Exception:
            logger.exception("Could not clear orders")
            await query.edit_message_text("Не удалось очистить файл. Проверь логи приложения.", reply_markup=admin_inline_keyboard())
    elif query.data == "admin_recent":
        try:
            _, ws = get_sheet()
            rows = list(ws.iter_rows(min_row=2, values_only=True))
            if not rows:
                text = "Заказов пока нет."
            else:
                lines = []
                for row in rows[-10:][::-1]:
                    # ID, date, outlet, address, serial, phone, amount, seller, telegram ID, created
                    lines.append(f"#{row[0]} · {row[1]} · ТТ {row[2]} · {row[6]} · {row[7]}")
                text = "📋 Последние заказы (до 10):\n\n" + "\n".join(lines)
            await query.edit_message_text(text, reply_markup=admin_inline_keyboard())
        except Exception:
            logger.exception("Could not read orders")
            await query.edit_message_text("Не удалось прочитать заказы.", reply_markup=admin_inline_keyboard())
    elif query.data == "admin_export":
        try:
            ensure_workbook()
            with open(DATA_FILE, "rb") as f:
                await query.message.reply_document(document=f, filename=DATA_FILE.name, caption="📤 Файл с заказами")
            await query.edit_message_text("Файл Excel отправлен выше.", reply_markup=admin_inline_keyboard())
        except Exception:
            logger.exception("Could not export orders")
            await query.edit_message_text("Не удалось отправить Excel-файл.", reply_markup=admin_inline_keyboard())
    elif query.data == "admin_back":
        await query.edit_message_text("🛠 Панель администратора", reply_markup=admin_inline_keyboard())


async def orders_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("Эта команда доступна только администратору.")
        return
    await admin_menu(update, context)


def main():
    ensure_workbook()
    app = Application.builder().token(BOT_TOKEN).build()
    conversation = ConversationHandler(
        entry_points=[
            CommandHandler("neworder", order_start),
            MessageHandler(filters.Regex(f"^{NEW_ORDER_TEXT}$"), order_start),
        ],
        states={
            DATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_date)],
            OUTLET_NUMBER: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_outlet_number)],
            OUTLET_ADDRESS: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_outlet_address)],
            SIM_SERIAL: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_sim_serial)],
            SIM_PHONE: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_sim_phone)],
            AMOUNT: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_amount)],
            SELLER: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_seller)],
        },
        fallbacks=[CommandHandler("cancel", cancel), MessageHandler(filters.Regex(f"^{CANCEL_TEXT}$"), cancel)],
        allow_reentry=True,
    )
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("orders", orders_command))
    app.add_handler(MessageHandler(filters.Regex(f"^{ADMIN_MENU_TEXT}$"), admin_menu))
    app.add_handler(conversation)
    app.add_handler(CallbackQueryHandler(admin_callback, pattern="^admin_"))
    logger.info("Bot started")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
