import os
import json
from datetime import datetime

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from telegram import Update, ReplyKeyboardMarkup
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, ContextTypes, filters

TOKEN = os.getenv("BOT_TOKEN")

FILE_NAME = "construction_progress.xlsx"
STATE_FILE = "user_state.json"

MAX_ENTRANCES = 20
MAX_FLOORS = 20

apartments = [
    "Наруж.стены кладка","Внутр.стены кладка","ПГП","Гипс.штк","Цем.штк",
    "Сшитый пол","Стяжка","Эл.кв+СС","Окна ПВХ","Окна Аллюм","Двери"
]

mop = [
    "Внутр.стены кладка","Коллектор кладка","ВШ кладка","ШТК","Декоративная штк",
    "Сшитый пол","Стяжка","Плитка пол","Плитка настенная","Двери МОП",
    "Двери коллектор","Ограждения ЛК","Разводка ЭЛ+СС",
    "Стояк отопление","Стояк канализация","Стояк ливневка","Стояк вода"
]

percent_keyboard = ReplyKeyboardMarkup(
[
["0","10","20"],
["30","40","50"],
["60","70","80"],
["90","98","100"],
["Пропустить"],
["⬅️ Назад"]
], resize_keyboard=True
)

start_keyboard = ReplyKeyboardMarkup([["Start"]], resize_keyboard=True)

def get_fill(value):
    if value == 0:
        return PatternFill("solid", start_color="FF0000")
    elif value == 98:
        return PatternFill("solid", start_color="FFA500")
    elif value == 100:
        return PatternFill("solid", start_color="00FF00")
    else:
        return PatternFill("solid", start_color="FFFF00")

def entrance_keyboard():
    rows=[]
    for i in range(1,MAX_ENTRANCES+1,2):
        if i+1<=MAX_ENTRANCES:
            rows.append([str(i),str(i+1)])
        else:
            rows.append([str(i)])
    return ReplyKeyboardMarkup(rows,resize_keyboard=True)

def floor_keyboard():
    rows=[]
    for i in range(1,MAX_FLOORS+1,2):
        if i+1<=MAX_FLOORS:
            rows.append([str(i),str(i+1)])
        else:
            rows.append([str(i)])
    return ReplyKeyboardMarkup(rows,resize_keyboard=True)

if not os.path.exists(STATE_FILE):
    with open(STATE_FILE,"w") as f:
        json.dump({},f)

def load_state():
    with open(STATE_FILE) as f:
        return json.load(f)

def save_state():
    with open(STATE_FILE,"w") as f:
        json.dump(user_data,f)

user_data=load_state()

def init_excel():
    if not os.path.exists(FILE_NAME):
        wb=Workbook()
        ws=wb.active
        ws.append(["Дата","Время","Адрес","Подъезд","Этаж","Раздел","Пункт","Процент"])
        wb.save(FILE_NAME)

def save_excel(row):
    wb=load_workbook(FILE_NAME)
    ws=wb.active
    ws.append(row)
    last=ws.max_row
    ws[f"H{last}"].fill=get_fill(row[-1])
    wb.save(FILE_NAME)

def create_floor_excel(data):
    name=f"{data['address']}_подъезд_{data['entrance']}_этаж_{data['floor']}.xlsx".replace(" ","_")
    wb=Workbook()
    ws=wb.active
    ws.append(["Дата","Время","Адрес","Подъезд","Этаж","Раздел","Пункт","Процент"])

    for r in data["floor_rows"]:
        ws.append(r)

    for i in range(2,ws.max_row+1):
        ws[f"H{i}"].fill=get_fill(ws[f"H{i}"].value)

    wb.save(name)
    return name

def create_full_excel(data):
    name=f"{data['address']}_полный_обход.xlsx".replace(" ","_")
    wb=Workbook()
    ws=wb.active
    ws.append(["Дата","Время","Адрес","Подъезд","Этаж","Раздел","Пункт","Процент"])

    for r in data["all_rows"]:
        ws.append(r)

    for i in range(2,ws.max_row+1):
        ws[f"H{i}"].fill=get_fill(ws[f"H{i}"].value)

    wb.save(name)
    return name

async def start(update:Update,context:ContextTypes.DEFAULT_TYPE):
    uid=str(update.effective_user.id)
    user_data[uid]={"step":"address","floor_rows":[],"all_rows":[]}
    save_state()
    await update.message.reply_text("Введите адрес дома:",reply_markup=start_keyboard)

async def handle(update:Update,context:ContextTypes.DEFAULT_TYPE):

    uid=str(update.effective_user.id)
    text=update.message.text

    if text=="Start":
        await start(update,context)
        return

    if uid not in user_data:
        await update.message.reply_text("Нажмите Start",reply_markup=start_keyboard)
        return

    data=user_data[uid]
    step=data["step"]

    if text=="Завершить обход":

        if data["all_rows"]:
            file=create_full_excel(data)
            with open(file,"rb") as f:
                await update.message.reply_document(f)
            os.remove(file)

        user_data[uid]={"step":"address","floor_rows":[],"all_rows":[]}
        save_state()

        await update.message.reply_text("Введите адрес дома:",reply_markup=start_keyboard)
        return

    if step=="address":
        data["address"]=text
        data["step"]="entrance"
        save_state()
        await update.message.reply_text("Выберите подъезд:",reply_markup=entrance_keyboard())
        return

    if step=="entrance":
        data["entrance"]=text
        data["step"]="floor"
        save_state()
        await update.message.reply_text("Выберите этаж:",reply_markup=floor_keyboard())
        return

    if step=="floor":
        data["floor"]=text
        data["section"]="Кв"
        data["items"]=apartments
        data["index"]=0
        data["step"]="percent"
        save_state()

        await update.message.reply_text(f"Кв\n{apartments[0]}",reply_markup=percent_keyboard)
        return

    if step=="percent":

        items=data["items"]
        idx=data["index"]

        if text!="Пропустить":
            row=[
                datetime.now().strftime("%Y-%m-%d"),
                datetime.now().strftime("%H:%M"),
                data["address"],
                data["entrance"],
                data["floor"],
                data["section"],
                items[idx],
                int(text)
            ]

            data["floor_rows"].append(row)
            data["all_rows"].append(row)
            save_excel(row)

        data["index"]+=1
        save_state()

        if data["index"]<len(items):
            await update.message.reply_text(
                f"{data['section']}\n{items[data['index']]}",
                reply_markup=percent_keyboard
            )
            return

        if data["section"]=="Кв":
            data["section"]="МОП"
            data["items"]=mop
            data["index"]=0
            save_state()

            await update.message.reply_text(f"МОП\n{mop[0]}",reply_markup=percent_keyboard)
            return

        file=create_floor_excel(data)

        with open(file,"rb") as f:
            await update.message.reply_document(f)

        os.remove(file)

        data["floor_rows"]=[]
        data["step"]="after_floor"
        save_state()

        keyboard=ReplyKeyboardMarkup(
        [["Следующий этаж"],["Следующий подъезд"],["Завершить обход"]],
        resize_keyboard=True)

        await update.message.reply_text("Этаж завершён",reply_markup=keyboard)

    elif step=="after_floor":

        if text=="Следующий этаж":
            data["step"]="floor"
            save_state()
            await update.message.reply_text("Выберите этаж:",reply_markup=floor_keyboard())

        elif text=="Следующий подъезд":
            data["step"]="entrance"
            save_state()
            await update.message.reply_text("Выберите подъезд:",reply_markup=entrance_keyboard())

init_excel()

app=ApplicationBuilder().token(TOKEN).build()

app.add_handler(CommandHandler("start",start))
app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND,handle))

print("BOT STARTED")

app.run_polling()
