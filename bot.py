import os
from datetime import datetime
from openpyxl import Workbook, load_workbook

from telegram import Update, ReplyKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    ContextTypes,
    filters,
)

TOKEN = os.getenv("BOT_TOKEN")
FILE_NAME = "construction_progress.xlsx"

user_data = {}

apartments_sections = [
    "Кладка внутр. стены",
    "Кладка наруж. стены",
    "Стяжка",
    "ПГП",
    "Штукатурка гипс",
    "Штукатурка ЦПШ",
    "Сшитый пол",
    "Окна ПВХ",
    "Двери"
]

mop_sections = [
    "Кладка наруж. стены",
    "Кладка коллекторы",
    "Кладка ВШ",
    "Стяжка",
    "Окна алюм.",
    "Гипс МОП",
    "Плитка МОП",
    "Двери МОП",
    "Двери лифт"
]

percent_options = [["0", "10"], ["50", "98"], ["100"]]


def init_excel():
    if not os.path.exists(FILE_NAME):
        wb = Workbook()
        ws = wb.active
        ws.append(["Дата", "Время", "Подъезд", "Этаж", "Тип", "Раздел", "Процент"])
        wb.save(FILE_NAME)


def save_excel(row):
    wb = load_workbook(FILE_NAME)
    ws = wb.active
    ws.append(row)
    wb.save(FILE_NAME)


def main_keyboard():
    return ReplyKeyboardMarkup(
        [
            ["1", "2"],
            ["3", "4"],
            ["5", "6"],
            ["7", "8"],
            ["9", "10"],
            ["📥 Скачать Excel"]
        ],
        resize_keyboard=True
    )


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    user_data[uid] = {"step": "entrance"}

    await update.message.reply_text(
        "Выберите подъезд:",
        reply_markup=main_keyboard()
    )


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    text = update.message.text.strip()

    if text == "📥 Скачать Excel":
        with open(FILE_NAME, "rb") as f:
            await update.message.reply_document(f)
        return

    if uid not in user_data:
        await update.message.reply_text("Напишите /start")
        return

    step = user_data[uid]["step"]

    if step == "entrance":
        user_data[uid]["entrance"] = text
        user_data[uid]["step"] = "floor"

        keyboard = [[str(i), str(i+1)] for i in range(1, 20, 2)]

        await update.message.reply_text(
            "Выберите этаж:",
            reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
        )

    elif step == "floor":
        user_data[uid]["floor"] = text
        user_data[uid]["step"] = "type"

        await update.message.reply_text(
            "Выберите тип:",
            reply_markup=ReplyKeyboardMarkup(
                [["Квартиры"], ["МОП"]],
                resize_keyboard=True
            )
        )

    elif step == "type":
        user_data[uid]["type"] = text
        user_data[uid]["step"] = "section"

        if text == "Квартиры":
            keyboard = [[x] for x in apartments_sections]
        else:
            keyboard = [[x] for x in mop_sections]

        await update.message.reply_text(
            "Выберите раздел:",
            reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
        )

    elif step == "section":
        user_data[uid]["section"] = text
        user_data[uid]["step"] = "percent"

        await update.message.reply_text(
            "Выберите процент:",
            reply_markup=ReplyKeyboardMarkup(percent_options, resize_keyboard=True)
        )

    elif step == "percent":
        now = datetime.now()

        row = [
            now.strftime("%Y-%m-%d"),
            now.strftime("%H:%M:%S"),
            user_data[uid]["entrance"],
            user_data[uid]["floor"],
            user_data[uid]["type"],
            user_data[uid]["section"],
            text
        ]

        save_excel(row)

        user_data[uid]["step"] = "entrance"

        await update.message.reply_text(
            "✅ Сохранено\nВыберите следующий подъезд:",
            reply_markup=main_keyboard()
        )


init_excel()

app = ApplicationBuilder().token(TOKEN).build()

app.add_handler(CommandHandler("start", start))
app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

print("Bot running...")

app.run_polling()
