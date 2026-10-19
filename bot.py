import os
import logging
from datetime import date, datetime
from decimal import Decimal, InvalidOperation

from dotenv import load_dotenv
from openpyxl import Workbook, load_workbook
from telegram import (
    Update,
    ReplyKeyboardMarkup,
    ReplyKeyboardRemove,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
)
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ConversationHandler,
    ContextTypes,
    filters,
)

load_dotenv()
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()
ADMIN_ID_RAW = os.getenv("ADMIN_ID", "").strip()
ADMIN_IDS = {
    int(value.strip())
    for value in ADMIN_ID_RAW.split(",")
    if value.strip().isdigit()
}
SELLER_IDS_RAW = os.getenv("ALLOWED_SELLER_IDS", "").strip()
ALLOWED_SELLER_IDS = {
    int(value.strip())
    for value in SELLER_IDS_RAW.split(",")
    if value.strip().isdigit()
}

if not BOT_TOKEN:
    raise RuntimeError("Environment variable BOT_TOKEN is not set")
if not ADMIN_IDS:
    raise RuntimeError(
        "Укажите хотя бы один числовой Telegram ID в переменной ADMIN_ID. "
        "Несколько ID разделяются запятой."
    )

EXCEL_FILE = os.getenv("EXCEL_FILE", "orders.xlsx")
HEADERS = [
    "ID заказа",
    "Дата заказа",
    "Номер торговой точки",
    "Адрес торговой точки",
    "Серийный номер SIM",
    "Номер телефона SIM",
    "Сумма заказа",
    "Продавец",
    "Telegram ID автора",
    "Время создания",
]

(
    ORDER_DATE,
    OUTLET_NUMBER,
    OUTLET_ADDRESS,
    SIM_SERIAL,
    SIM_PHONE,
    ORDER_AMOUNT,
    SELLER_NAME,
    CONFIRM_ORDER,
) = range(8)

MAIN_KEYBOARD = ReplyKeyboardMarkup(
    [["➕ Добавить заказ"]],
    resize_keyboard=True,
)
ADMIN_KEYBOARD = ReplyKeyboardMarkup(
    [
        ["➕ Добавить заказ"],
        ["📋 Последние заказы", "📥 Скачать Excel"],
        ["🗑 Очистить все заказы"],
    ],
    resize_keyboard=True,
)
CANCEL_KEYBOARD = ReplyKeyboardMarkup([["❌ Отмена"]], resize_keyboard=True)


def is_admin(user_id: int) -> bool:
    return user_id in ADMIN_IDS


def is_allowed_seller(user_id: int) -> bool:
    return is_admin(user_id) or not ALLOWED_SELLER_IDS or user_id in ALLOWED_SELLER_IDS


def get_main_keyboard(user_id: int):
    return ADMIN_KEYBOARD if is_admin(user_id) else MAIN_KEYBOARD


def ensure_workbook():
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = "Заказы"
        ws.append(HEADERS)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:J1"
        wb.save(EXCEL_FILE)


def read_orders():
    ensure_workbook()
    wb = load_workbook(EXCEL_FILE)
    ws = wb.active
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    wb.close()
    return [row for row in rows if any(value is not None for value in row)]


def append_order(values):
    ensure_workbook()
    wb = load_workbook(EXCEL_FILE)
    ws = wb.active
    next_id = max(
        [int(row[0]) for row in ws.iter_rows(min_row=2, max_col=1, values_only=True)
         if row[0] is not None and str(row[0]).isdigit()] or [0]
    ) + 1
    ws.append([next_id] + values)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(EXCEL_FILE)
    wb.close()
    return next_id


def clear_orders():
    ensure_workbook()
    wb = load_workbook(EXCEL_FILE)
    ws = wb.active
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)
    ws.auto_filter.ref = "A1:J1"
    wb.save(EXCEL_FILE)
    wb.close()


def format_order(row):
    # row is a tuple in Excel column order
    return (
        f"Заказ №{row[0]}\n"
        f"Дата: {row[1]}\n"
        f"Точка: {row[2]}\n"
        f"Адрес: {row[3]}\n"
        f"Серийный номер SIM: {row[4]}\n"
        f"Номер телефона SIM: {row[5]}\n"
        f"Сумма: {row[6]}\n"
        f"Продавец: {row[7]}\n"
        f"Telegram ID автора: {row[8]}\n"
        f"Создан: {row[9]}"
    )


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    if not is_allowed_seller(user.id):
        await update.message.reply_text(
            "У вас нет доступа к этому боту. Обратитесь к администратору."
        )
        return
    role = "администратор" if is_admin(user.id) else "продавец"
    await update.message.reply_text(
        f"Привет! Вы вошли как {role}.\n"
        "Нажмите «➕ Добавить заказ», чтобы внести новый заказ.",
        reply_markup=get_main_keyboard(user.id),
    )


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.pop("new_order", None)
    await update.message.reply_text(
        "Ввод заказа отменён.",
        reply_markup=get_main_keyboard(update.effective_user.id),
    )
    return ConversationHandler.END


async def begin_order(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    if not is_allowed_seller(user.id):
        await update.message.reply_text("У вас нет доступа к добавлению заказов.")
        return ConversationHandler.END
    context.user_data["new_order"] = {}
    context.user_data["new_order"]["telegram_id"] = user.id
    context.user_data["new_order"]["created_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    await update.message.reply_text(
        f"Введите дату заказа в формате ДД.ММ.ГГГГ.\n"
        f"Сегодня: {date.today().strftime('%d.%m.%Y')}\n"
        "Можно отправить «сегодня».",
        reply_markup=CANCEL_KEYBOARD,
    )
    return ORDER_DATE


async def order_date(update: Update, context: ContextTypes.DEFAULT_TYPE):
    raw = update.message.text.strip().lower()
    if raw == "сегодня":
        parsed = date.today()
    else:
        try:
            parsed = datetime.strptime(raw, "%d.%m.%Y").date()
        except ValueError:
            await update.message.reply_text("Не понял дату. Введите ДД.ММ.ГГГГ, например 20.10.2026.")
            return ORDER_DATE
    context.user_data["new_order"]["date"] = parsed.strftime("%d.%m.%Y")
    await update.message.reply_text("Введите номер торговой точки:")
    return OUTLET_NUMBER


async def outlet_number(update: Update, context: ContextTypes.DEFAULT_TYPE):
    value = update.message.text.strip()
    if not value:
        await update.message.reply_text("Поле не должно быть пустым. Введите номер торговой точки:")
        return OUTLET_NUMBER
    context.user_data["new_order"]["outlet_number"] = value
    await update.message.reply_text("Введите адрес торговой точки:")
    return OUTLET_ADDRESS


async def outlet_address(update: Update, context: ContextTypes.DEFAULT_TYPE):
    value = update.message.text.strip()
    if not value:
        await update.message.reply_text("Поле не должно быть пустым. Введите адрес торговой точки:")
        return OUTLET_ADDRESS
    context.user_data["new_order"]["outlet_address"] = value
    await update.message.reply_text("Введите серийный номер SIM-карты:")
    return SIM_SERIAL


async def sim_serial(update: Update, context: ContextTypes.DEFAULT_TYPE):
    value = update.message.text.strip()
    if not value:
        await update.message.reply_text("Поле не должно быть пустым. Введите серийный номер SIM:")
        return SIM_SERIAL
    context.user_data["new_order"]["sim_serial"] = value
    await update.message.reply_text("Введите номер телефона SIM-карты:")
    return SIM_PHONE


async def sim_phone(update: Update, context: ContextTypes.DEFAULT_TYPE):
    value = update.message.text.strip()
    if not value:
        await update.message.reply_text("Поле не должно быть пустым. Введите номер телефона SIM:")
        return SIM_PHONE
    context.user_data["new_order"]["sim_phone"] = value
    await update.message.reply_text("Введите сумму заказа (например, 1500 или 1500.50):")
    return ORDER_AMOUNT


async def order_amount(update: Update, context: ContextTypes.DEFAULT_TYPE):
    raw = update.message.text.strip().replace(",", ".").replace(" ", "")
    try:
        amount = Decimal(raw)
        if amount < 0:
            raise InvalidOperation
    except (InvalidOperation, ValueError):
        await update.message.reply_text("Введите сумму числом, например 1500 или 1500.50:")
        return ORDER_AMOUNT
    context.user_data["new_order"]["amount"] = str(amount)
    await update.message.reply_text("Введите имя или фамилию продавца:")
    return SELLER_NAME


async def seller_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    value = update.message.text.strip()
    if not value:
        await update.message.reply_text("Поле не должно быть пустым. Введите имя продавца:")
        return SELLER_NAME
    context.user_data["new_order"]["seller"] = value
    order = context.user_data["new_order"]
    summary = (
        "Проверьте заказ перед сохранением:\n\n"
        f"Дата: {order['date']}\n"
        f"Номер точки: {order['outlet_number']}\n"
        f"Адрес: {order['outlet_address']}\n"
        f"Серийный номер SIM: {order['sim_serial']}\n"
        f"Номер телефона SIM: {order['sim_phone']}\n"
        f"Сумма: {order['amount']}\n"
        f"Продавец: {order['seller']}"
    )
    keyboard = InlineKeyboardMarkup([
        [
            InlineKeyboardButton("✅ Сохранить заказ", callback_data="order:save"),
            InlineKeyboardButton("❌ Отменить", callback_data="order:cancel"),
        ]
    ])
    await update.message.reply_text(summary, reply_markup=ReplyKeyboardRemove())
    await update.message.reply_text("Сохранить этот заказ?", reply_markup=keyboard)
    return CONFIRM_ORDER


async def confirm_order(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.data == "order:cancel":
        context.user_data.pop("new_order", None)
        await query.edit_message_text("Заказ отменён.")
        await query.message.reply_text(
            "Выберите действие:",
            reply_markup=get_main_keyboard(query.from_user.id),
        )
        return ConversationHandler.END

    order = context.user_data.get("new_order")
    if not order:
        await query.edit_message_text("Данные заказа не найдены. Начните заново через «➕ Добавить заказ».")
        return ConversationHandler.END

    values = [
        order["date"],
        order["outlet_number"],
        order["outlet_address"],
        order["sim_serial"],
        order["sim_phone"],
        order["amount"],
        order["seller"],
        order["telegram_id"],
        order["created_at"],
    ]
    order_id = append_order(values)
    await query.edit_message_text(f"✅ Заказ №{order_id} сохранён в Excel.")
    context.user_data.pop("new_order", None)
    await query.message.reply_text(
        "Выберите действие:",
        reply_markup=get_main_keyboard(query.from_user.id),
    )

    for admin_id in ADMIN_IDS:
        if admin_id != query.from_user.id:
            try:
                await context.bot.send_message(
                    chat_id=admin_id,
                    text=f"Добавлен новый заказ №{order_id}.\nПродавец: {order['seller']}\nСумма: {order['amount']}",
                )
            except Exception:
                logger.exception("Не удалось уведомить администратора %s", admin_id)
    return ConversationHandler.END


async def recent_orders(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_admin(user_id):
        await update.message.reply_text("Эта команда доступна только администраторам.")
        return
    rows = read_orders()
    if not rows:
        await update.message.reply_text("Пока заказов нет.", reply_markup=ADMIN_KEYBOARD)
        return
    chunks = []
    for row in rows[-10:][::-1]:
        chunks.append(format_order(row))
    # Telegram message length limit safety
    message = "\n\n──────────\n\n".join(chunks)
    for start_idx in range(0, len(message), 3800):
        await update.message.reply_text(message[start_idx:start_idx + 3800], reply_markup=ADMIN_KEYBOARD)


async def export_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("Скачать Excel могут только администраторы.")
        return
    ensure_workbook()
    with open(EXCEL_FILE, "rb") as file:
        await update.message.reply_document(
            document=file,
            filename=os.path.basename(EXCEL_FILE),
            caption="Файл со всеми заказами.",
        )


async def clear_orders_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if not is_admin(update.effective_user.id):
        await update.message.reply_text("Очищать заказы могут только администраторы.")
        return
    keyboard = InlineKeyboardMarkup([
        [
            InlineKeyboardButton("Да, удалить все", callback_data="clear:yes"),
            InlineKeyboardButton("Нет, отмена", callback_data="clear:no"),
        ]
    ])
    await update.message.reply_text(
        "⚠️ Вы уверены, что хотите удалить ВСЕ заказы из Excel? Это действие нельзя отменить.",
        reply_markup=keyboard,
    )


async def clear_orders_confirm(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if not is_admin(query.from_user.id):
        await query.edit_message_text("Нет доступа.")
        return
    if query.data == "clear:no":
        await query.edit_message_text("Очистка отменена. Заказы сохранены.")
        return
    clear_orders()
    await query.edit_message_text("🗑 Все заказы удалены из Excel. Заголовки таблицы сохранены.")


async def unknown_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = update.message.text.strip()
    if text == "📋 Последние заказы":
        await recent_orders(update, context)
    elif text == "📥 Скачать Excel":
        await export_excel(update, context)
    elif text == "🗑 Очистить все заказы":
        await clear_orders_start(update, context)
    elif text == "➕ Добавить заказ":
        await begin_order(update, context)
    else:
        await update.message.reply_text(
            "Выберите действие на клавиатуре ниже.",
            reply_markup=get_main_keyboard(user_id),
        )


def main():
    ensure_workbook()
    app = Application.builder().token(BOT_TOKEN).build()

    conversation = ConversationHandler(
        entry_points=[
            CommandHandler("add", begin_order),
            MessageHandler(filters.Regex(r"^➕ Добавить заказ$"), begin_order),
        ],
        states={
            ORDER_DATE: [MessageHandler(filters.TEXT & ~filters.COMMAND, order_date)],
            OUTLET_NUMBER: [MessageHandler(filters.TEXT & ~filters.COMMAND, outlet_number)],
            OUTLET_ADDRESS: [MessageHandler(filters.TEXT & ~filters.COMMAND, outlet_address)],
            SIM_SERIAL: [MessageHandler(filters.TEXT & ~filters.COMMAND, sim_serial)],
            SIM_PHONE: [MessageHandler(filters.TEXT & ~filters.COMMAND, sim_phone)],
            ORDER_AMOUNT: [MessageHandler(filters.TEXT & ~filters.COMMAND, order_amount)],
            SELLER_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, seller_name)],
            CONFIRM_ORDER: [CallbackQueryHandler(confirm_order, pattern=r"^order:(save|cancel)$")],
        },
        fallbacks=[CommandHandler("cancel", cancel), MessageHandler(filters.Regex(r"^❌ Отмена$"), cancel)],
        allow_reentry=True,
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("orders", recent_orders))
    app.add_handler(CommandHandler("export", export_excel))
    app.add_handler(CommandHandler("clear", clear_orders_start))
    app.add_handler(CallbackQueryHandler(clear_orders_confirm, pattern=r"^clear:(yes|no)$"))
    app.add_handler(conversation)
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, unknown_text))

    logger.info("Bot started. Admin IDs configured: %s", sorted(ADMIN_IDS))
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
