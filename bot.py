import os
import json
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
STATE_FILE = "user_state.json"

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


def load_state():
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_state():
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(user_data, f, ensure_ascii=False)


user_data = load_state()


percent_keyboard = ReplyKeyboardMarkup(
    [
        ["0", "10"],
        ["50", "98"],
        ["100"],
        ["⬅️ Назад"]
    ],
    resize_keyboard=True
)


def init_excel():
    if not os.path.exists(FILE_NAME):
        wb = Workbook()
        ws = wb.active
        ws.append([
            "Дата",
            "Время",
            "Адрес",
            "Подъезд",
            "Этаж",
            "Тип",
            "Пункт",
            "Процент"
        ])
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
    uid = str(update.effective_user.id)

    user_data[uid] = {
        "step": "address",
        "percents": []
    }

    save_state()

    await update.message.reply_text("Введите адрес дома:")


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    uid = str(update.effective_user.id)
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

    if step == "address":
        data["address"] = text
        data["step"] = "entrance"
        save_state()

        await update.message.reply_text(
            "Выберите подъезд:",
            reply_markup=entrance_keyboard()
        )

    elif step == "entrance":
        data["entrance"] = text
        data["step"] = "floor"
        save_state()

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
        data["percents"] = []
        data["step"] = "percent"
        save_state()

        await update.message.reply_text(
            f"Квартиры\n{apartments[0]}:",
            reply_markup=percent_keyboard
        )

    elif step == "percent":

        if text == "⬅️ Назад":
            if data["index"] > 0:
                data["index"] -= 1
                if data["percents"]:
                    data["percents"].pop()

            save_state()

            await update.message.reply_text(
                f"{data['type']}\n{data['items'][data['index']]}:",
                reply_markup=percent_keyboard
            )
            return

        item = data["items"][data["index"]]
        percent = int(text)

        data["percents"].append(percent)

        now = datetime.now()

        row = [
            now.strftime("%Y-%m-%d"),
            now.strftime("%H:%M:%S"),
            data["address"],
            data["entrance"],
            data["floor"],
            data["type"],
            item,
            percent
        ]

        save_excel(row)

        data["index"] += 1
        save_state()

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
                save_state()

                await update.message.reply_text(
                    f"МОП\n{mop[0]}:",
                    reply_markup=percent_keyboard
                )

            else:
                avg = round(sum(data["percents"]) / len(data["percents"]), 1)

                data["step"] = "after_floor"
                save_state()

                await update.message.reply_text(
                    f"✅ Этаж завершён\nСредняя готовность этажа: {avg}%"
                )

                await update.message.reply_text(
                    "Что дальше?",
                    reply_markup=ReplyKeyboardMarkup(
                        [["Следующий этаж"], ["Подъезд завершён"]],
                        resize_keyboard=True
                    )
                )

    elif step == "after_floor":
        if text == "Следующий этаж":
            data["step"] = "floor"
            save_state()

            keyboard = [[str(i), str(i+1)] for i in range(1, 20, 2)]

            await update.message.reply_text(
                "Выберите этаж:",
                reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
            )

        elif text == "Подъезд завершён":
            data["step"] = "after_entrance"
            save_state()

            await update.message.reply_text(
                "Продолжить обход дома?",
                reply_markup=ReplyKeyboardMarkup(
                    [
                        ["Следующий подъезд"],
                        ["Новый адрес"],
                        ["📥 Скачать Excel"]
                    ],
                    resize_keyboard=True
                )
            )

    elif step == "after_entrance":
        if text == "Следующий подъезд":
            data["step"] = "entrance"
            save_state()

            await update.message.reply_text(
                "Выберите следующий подъезд:",
                reply_markup=entrance_keyboard()
            )

        elif text == "Новый адрес":
            data["step"] = "address"
            save_state()

            await update.message.reply_text("Введите новый адрес:")


init_excel()

app = ApplicationBuilder().token(TOKEN).build()

app.add_handler(CommandHandler("start", start))
app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

print("Bot running...")

app.run_polling()
