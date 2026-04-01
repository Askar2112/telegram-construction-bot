import os
import json
from datetime import datetime

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

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

MAX_ENTRANCES = 20
MAX_FLOORS = 20

apartments = [
    "Наруж.стены кладка",
    "Внутр.стены кладка",
    "ПГП",
    "Гипс.штк",
    "Цем.штк",
    "Сшитый пол",
    "Стяжка",
    "Эл.кв+СС",
    "Окна ПВХ",
    "Окна Аллюм",
    "Двери"
]

mop = [
    "Внутр.стены кладка",
    "Коллектор кладка",
    "ВШ кладка",
    "ШТК",
    "Декоративная штк",
    "Сшитый пол",
    "Стяжка",
    "Плитка пол",
    "Плитка настенная",
    "Двери МОП",
    "Двери коллектор",
    "Ограждения ЛК",
    "Разводка ЭЛ+СС",
    "Стояк отопление",
    "Стояк канализация",
    "Стояк ливневка",
    "Стояк вода"
]

if not os.path.exists(STATE_FILE):
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump({}, f)


def load_state():
    with open(STATE_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def save_state():
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump(user_data, f, ensure_ascii=False)


user_data = load_state()

percent_keyboard = ReplyKeyboardMarkup(
    [
        ["0", "10"],
        ["50", "60"],
        ["100"],
        ["Пропустить"],
        ["⬅️ Назад"]
    ],
    resize_keyboard=True
)


def entrance_keyboard():
    keyboard = []
    for i in range(1, MAX_ENTRANCES + 1, 2):
        if i < MAX_ENTRANCES:
            keyboard.append([str(i), str(i + 1)])
        else:
            keyboard.append([str(i)])
    return ReplyKeyboardMarkup(keyboard, resize_keyboard=True)


def floor_keyboard():
    keyboard = []
    for i in range(1, MAX_FLOORS + 1, 2):
        if i < MAX_FLOORS:
            keyboard.append([str(i), str(i + 1)])
        else:
            keyboard.append([str(i)])
    return ReplyKeyboardMarkup(keyboard, resize_keyboard=True)


def init_excel():
    if not os.path.exists(FILE_NAME):
        wb = Workbook()
        ws = wb.active
        ws.append(["Дата", "Время", "Адрес", "Подъезд", "Этаж", "Раздел", "Пункт", "Процент"])

        for cell in ws[1]:
            cell.font = Font(bold=True)

        wb.save(FILE_NAME)


def save_excel(row):
    wb = load_workbook(FILE_NAME)
    ws = wb.active
    ws.append(row)
    wb.save(FILE_NAME)


def get_fill(percent):
    if percent == 0:
        return PatternFill("solid", start_color="FF9999")
    elif percent < 100:
        return PatternFill("solid", start_color="FFF599")
    else:
        return PatternFill("solid", start_color="99FF99")


def format_sheet(ws):
    widths = {
        "A": 16,
        "B": 14,
        "C": 35,
        "D": 12,
        "E": 12,
        "F": 18,
        "G": 40,
        "H": 14
    }

    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    ws.freeze_panes = "A2"

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=True)


def create_floor_excel(data):
    filename = f"{data['address'].replace(' ', '_')}_подъезд_{data['entrance']}_этаж_{data['floor']}.xlsx"

    wb = Workbook()
    ws = wb.active

    headers = ["Дата", "Время", "Адрес", "Подъезд", "Этаж", "Раздел", "Пункт", "Процент"]
    ws.append(headers)

    for cell in ws[1]:
        cell.font = Font(bold=True)

    for row in data["floor_rows"]:
        ws.append(row)

    format_sheet(ws)

    for row_num in range(2, len(data["floor_rows"]) + 2):
        percent = ws[f"H{row_num}"].value
        ws[f"H{row_num}"].fill = get_fill(percent)

    wb.save(filename)
    return filename


def create_full_excel(data):
    filename = f"{data['address'].replace(' ', '_')}_полный_обход.xlsx"

    wb = Workbook()
    ws = wb.active

    headers = ["Дата", "Время", "Адрес", "Подъезд", "Этаж", "Раздел", "Пункт", "Процент"]
    ws.append(headers)

    for cell in ws[1]:
        cell.font = Font(bold=True)

    for row in data["all_rows"]:
        ws.append(row)

    format_sheet(ws)

    for row_num in range(2, len(data["all_rows"]) + 2):
        percent = ws[f"H{row_num}"].value
        ws[f"H{row_num}"].fill = get_fill(percent)

    wb.save(filename)
    return filename


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    uid = str(update.effective_user.id)

    user_data[uid] = {
        "step": "address",
        "percents": [],
        "floor_rows": [],
        "all_rows": []
    }

    save_state()

    await update.message.reply_text("Введите адрес дома:")


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    uid = str(update.effective_user.id)
    text = update.message.text.strip()

    if uid not in user_data:
        await update.message.reply_text("Напишите /start")
        return

    data = user_data[uid]
    step = data["step"]

    if text == "📥 Скачать Excel":
        if data.get("floor_rows"):
            filename = create_floor_excel(data)

            with open(filename, "rb") as f:
                await update.message.reply_document(document=f, filename=filename)

            os.remove(filename)
        return

    if text == "Завершить обход":
        if data.get("all_rows"):
            filename = create_full_excel(data)

            with open(filename, "rb") as f:
                await update.message.reply_document(document=f, filename=filename)

            os.remove(filename)

        user_data[uid] = {
            "step": "address",
            "percents": [],
            "floor_rows": [],
            "all_rows": []
        }

        save_state()

        await update.message.reply_text("✅ Обход завершён")
        await update.message.reply_text("Введите адрес дома:")
        return

    if step == "address":
        data["address"] = text
        data["step"] = "entrance"
        save_state()
        await update.message.reply_text("Выберите подъезд:", reply_markup=entrance_keyboard())

    elif step == "entrance":
        data["entrance"] = text
        data["step"] = "floor"
        save_state()
        await update.message.reply_text("Выберите этаж:", reply_markup=floor_keyboard())

    elif step == "floor":
        data["floor"] = text
        data["type"] = "Кв"
        data["items"] = apartments
        data["index"] = 0
        data["percents"] = []
        data["floor_rows"] = []
        data["step"] = "percent"
        save_state()

        await update.message.reply_text(f"Кв\n{apartments[0]}:", reply_markup=percent_keyboard)

    elif step == "percent":

        if text == "⬅️ Назад":
            if data["index"] > 0:
                data["index"] -= 1
                if data["percents"]:
                    data["percents"].pop()
                if data["floor_rows"]:
                    data["floor_rows"].pop()
                if data["all_rows"]:
                    data["all_rows"].pop()

            save_state()

            await update.message.reply_text(
                f"{data['type']}\n{data['items'][data['index']]}:",
                reply_markup=percent_keyboard
            )
            return

        if text == "Пропустить":
            data["index"] += 1
            save_state()

        else:
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

            data["floor_rows"].append(row)
            data["all_rows"].append(row)
            save_excel(row)

            data["index"] += 1
            save_state()

        if data["index"] < len(data["items"]):
            await update.message.reply_text(
                f"{data['type']}\n{data['items'][data['index']]}:",
                reply_markup=percent_keyboard
            )
            return

        if data["type"] == "Кв":
            data["type"] = "МОП"
            data["items"] = mop
            data["index"] = 0
            save_state()

            await update.message.reply_text(
                f"МОП\n{mop[0]}:",
                reply_markup=percent_keyboard
            )

        else:
            avg = round(sum(data["percents"]) / len(data["percents"]), 1) if data["percents"] else 0

            data["step"] = "after_floor"
            save_state()

            await update.message.reply_text(f"✅ Этаж завершён\nСредняя готовность: {avg}%")

            await update.message.reply_text(
                "Что дальше?",
                reply_markup=ReplyKeyboardMarkup(
                    [
                        ["Следующий этаж"],
                        ["Подъезд завершён"],
                        ["📥 Скачать Excel"],
                        ["Завершить обход"]
                    ],
                    resize_keyboard=True
                )
            )

    elif step == "after_floor":
        if text == "Следующий этаж":
            data["step"] = "floor"
            save_state()
            await update.message.reply_text("Выберите этаж:", reply_markup=floor_keyboard())

        elif text == "Подъезд завершён":
            data["step"] = "after_entrance"
            save_state()
            await update.message.reply_text(
                "Продолжить обход?",
                reply_markup=ReplyKeyboardMarkup(
                    [
                        ["Следующий подъезд"],
                        ["Новый адрес"],
                        ["📥 Скачать Excel"],
                        ["Завершить обход"]
                    ],
                    resize_keyboard=True
                )
            )

    elif step == "after_entrance":
        if text == "Следующий подъезд":
            data["step"] = "entrance"
            save_state()
            await update.message.reply_text("Выберите подъезд:", reply_markup=entrance_keyboard())

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
