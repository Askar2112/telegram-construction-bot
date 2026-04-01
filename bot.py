import os
from datetime import datetime
from openpyxl import Workbook, load_workbook

from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    CallbackQueryHandler,
    ContextTypes
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

percent_options = ["0", "10", "50", "98", "100"]


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


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [
        [
            InlineKeyboardButton(str(i), callback_data=f"entrance_{i}"),
            InlineKeyboardButton(str(i+1), callback_data=f"entrance_{i+1}")
        ]
        for i in range(1, 10, 2)
    ]

    await update.message.reply_text(
        "Выберите подъезд:",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )


async def button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    uid = query.from_user.id
    data = query.data

    try:
        if uid not in user_data:
            user_data[uid] = {}

        if data.startswith("entrance_"):
            user_data[uid]["entrance"] = data.split("_")[1]

            keyboard = [
                [
                    InlineKeyboardButton(str(i), callback_data=f"floor_{i}"),
                    InlineKeyboardButton(str(i+1), callback_data=f"floor_{i+1}")
                ]
                for i in range(1, 20, 2)
            ]

            await context.bot.send_message(
                chat_id=query.message.chat_id,
                text="Выберите этаж:",
                reply_markup=InlineKeyboardMarkup(keyboard)
            )

        elif data.startswith("floor_"):
            user_data[uid]["floor"] = data.split("_")[1]

            keyboard = [
                [InlineKeyboardButton("Квартиры", callback_data="apartments")],
                [InlineKeyboardButton("МОП", callback_data="mop")]
            ]

            await context.bot.send_message(
                chat_id=query.message.chat_id,
                text="Выберите тип:",
                reply_markup=InlineKeyboardMarkup(keyboard)
            )

        elif data == "apartments":
            user_data[uid]["type"] = "Квартиры"

            keyboard = [
                [InlineKeyboardButton(section, callback_data=f"section_{section}")]
                for section in apartments_sections
            ]

            await context.bot.send_message(
                chat_id=query.message.chat_id,
                text="Выберите раздел:",
                reply_markup=InlineKeyboardMarkup(keyboard)
            )

        elif data == "mop":
            user_data[uid]["type"] = "МОП"

            keyboard = [
                [InlineKeyboardButton(section, callback_data=f"section_{section}")]
                for section in mop_sections
            ]

            await context.bot.send_message(
                chat_id=query.message.chat_id,
                text="Выберите раздел:",
                reply_markup=InlineKeyboardMarkup(keyboard)
            )

        elif data.startswith("section_"):
            user_data[uid]["section"] = data.replace("section_", "")

            keyboard = [
                [InlineKeyboardButton(p, callback_data=f"percent_{p}")]
                for p in percent_options
            ]

            await context.bot.send_message(
                chat_id=query.message.chat_id,
                text="Выберите процент:",
                reply_markup=InlineKeyboardMarkup(keyboard)
            )

        elif data.startswith("percent_"):
            percent = data.replace("percent_", "")
            now = datetime.now()

            row = [
                now.strftime("%Y-%m-%d"),
                now.strftime("%H:%M:%S"),
                user_data[uid]["entrance"],
                user_data[uid]["floor"],
                user_data[uid]["type"],
                user_data[uid]["section"],
                percent
            ]

            save_excel(row)

            await context.bot.send_message(
                chat_id=query.message.chat_id,
                text="✅ Сохранено"
            )

    except Exception as e:
        print("ERROR:", e)

        await context.bot.send_message(
            chat_id=query.message.chat_id,
            text=f"Ошибка: {e}"
        )


init_excel()

app = ApplicationBuilder().token(TOKEN).build()

app.add_handler(CommandHandler("start", start))
app.add_handler(CallbackQueryHandler(button))

print("Bot running...")

app.run_polling()
