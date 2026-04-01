import os
from datetime import datetime
from openpyxl import Workbook, load_workbook

from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.ext import (
ApplicationBuilder,
CommandHandler,
CallbackQueryHandler,
MessageHandler,
filters,
ContextTypes
)

TOKEN = os.getenv("BOT_TOKEN")
FILE_NAME = "construction_progress.xlsx"

apartments_sections = [
"Кладка внутренние стены",
"Кладка наружные стены",
"Стяжка",
"ПГП",
"Штукатурка гипс",
"Штукатурка ЦПШ",
"Сшитый пол",
"Окна ПВХ",
"Двери"
]

mop_sections = [
"Кладка наружные стены",
"Кладка коллекторы",
"Кладка ВШ",
"Стяжка",
"Окна аллюминий",
"Гипс МОП",
"Плитка МОП",
"Двери МОП",
"Двери лифт",
"Ограждения лестниц",
"Двери кровля"
]

percent_options = ["0", "10", "50", "98", "100"]

user_data = {}

def init_excel():
if not os.path.exists(FILE_NAME):
wb = Workbook()
ws = wb.active
ws.append(["Дата","Время","Подъезд","Этаж","Тип","Раздел","Процент"])
wb.save(FILE_NAME)

def save_excel(row):
wb = load_workbook(FILE_NAME)
ws = wb.active
ws.append(row)
wb.save(FILE_NAME)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
keyboard = [[InlineKeyboardButton(f"Подъезд {i}", callback_data=f"entrance_{i}")]
for i in range(1, 11)]

```
await update.message.reply_text(
    "Выберите подъезд:",
    reply_markup=InlineKeyboardMarkup(keyboard)
)
```

async def button(update: Update, context: ContextTypes.DEFAULT_TYPE):
query = update.callback_query
await query.answer()

```
uid = query.from_user.id
data = query.data

if uid not in user_data:
    user_data[uid] = {}

if data.startswith("entrance_"):
    user_data[uid]["entrance"] = data.split("_")[1]

    keyboard = [[InlineKeyboardButton(f"Этаж {i}", callback_data=f"floor_{i}")]
                for i in range(1, 21)]

    await query.edit_message_text(
        "Выберите этаж:",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )

elif data.startswith("floor_"):
    user_data[uid]["floor"] = data.split("_")[1]

    keyboard = [
        [InlineKeyboardButton("Квартиры", callback_data="apartments")],
        [InlineKeyboardButton("МОП", callback_data="mop")]
    ]

    await query.edit_message_text(
        "Выберите раздел:",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )
```

init_excel()

app = ApplicationBuilder().token(TOKEN).build()

app.add_handler(CommandHandler("start", start))
app.add_handler(CallbackQueryHandler(button))

print("Bot running...")

app.run_polling()
