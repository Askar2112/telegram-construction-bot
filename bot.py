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

apartments = [
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

mop = [
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

percent_keyboard = ReplyKeyboardMarkup(
    [["0", "10"], ["50", "98"], ["100"]],
    resize_keyboard=True
)


def init_excel():
    if not os.path.exists(FILE_NAME):
        wb = Workbook()
        ws = wb.active
        ws.append(["Дата", "Время", "Подъезд", "Этаж", "Тип", "Пункт", "Процент"])
        wb.save(FILE_NAME)


def save_excel(row):
    wb = load_workbook(FILE_NAME)
    ws = wb.active
    ws.append(row)
    wb.save(FILE_NAME)


def entrance_keyboard():
    return ReplyKeyboardMarkup(
        [["1", "2"], ["3", "4"], ["5", "6"], ["7", "8"], ["9", "10"]],
        resize_keyboard=True
    )


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    user_data[uid] = {"step": "entrance"}

    await update.message.reply_text(
        "Выберите подъезд:",
        reply_markup=entrance_keyboard()
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

    data = user_data[uid]
    step = data["step"]

    if step == "entrance":
        data["entrance"] = text
        data["step"] = "floor"

        keyboard = [[str(i), str(i+1)] for i in range(1, 20, 2)]

        await update.message.reply_text(
            "Выберите этаж:",
            reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
        )

    elif step == "floor":
        data["floor"] = text
        data["type"] = "Квартиры"
        data["items"] = apartments
        data["index"] = 0
        data["step"] = "percent"

        await update.message.reply_text(
            f"Квартиры\n{apartments[0]}:",
            reply_markup=percent_keyboard
        )

    elif step == "percent":
        item = data["items"][data["index"]]

        now = datetime.now()

        row = [
            now.strftime("%Y-%m-%d"),
            now.strftime("%H:%M:%S"),
            data["entrance"],
            data["floor"],
            data["type"],
            item,
            text
        ]

        save_excel(row)

        data["index"] += 1

        if data["index"] < len(data["items"]):
            await update.message.reply_text(
                f"{data['type']}\n{data['items'][data['index']]}:",
                reply_markup=percent_keyboard
            )

        else:
            if data["type"] == "Квартиры":
                data["type"] = "МОП"
                data["items"] = mop
                data["index"] = 0

                await update.message.reply_text(
                    f"МОП\n{mop[0]}:",
                    reply_markup=percent_keyboard
                )

            else:
                data["step"] = "next"

                await update.message.reply_text(
                    "✅ Этаж завершён",
                    reply_markup=ReplyKeyboardMarkup(
                        [
                            ["Следующий этаж"],
                            ["Другой подъезд"],
                            ["📥 Скачать Excel"]
                        ],
                        resize_keyboard=True
                    )
                )

    elif step == "next":
        if text == "Следующий этаж":
            data["step"] = "floor"

            keyboard = [[str(i), str(i+1)] for i in range(1, 20, 2)]

            await update.message.reply_text(
                "Выберите этаж:",
                reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
            )

        elif text == "Другой подъезд":
            data["step"] = "entrance"

            await update.message.reply_text(
                "Выберите подъезд:",
                reply_markup=entrance_keyboard()
            )


init_excel()

app = ApplicationBuilder().token(TOKEN).build()

app.add_handler(CommandHandler("start", start))
app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

print("Bot running...")

app.run_polling()
