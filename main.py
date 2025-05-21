import os
import pandas as pd
from simplekml import Kml
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ApplicationBuilder, MessageHandler, ContextTypes, filters

# --- Функция очистки координат ---
def clean_coord(coord_str):
    return (
        str(coord_str)
        .replace('\n', '')
        .replace('\r', '')
        .replace('"', '')
        .replace("'", '')
        .replace(' ', '')
        .strip()
    )

# --- Функция создания KML-файла из Excel ---
def create_kml_from_excel(file_path):
    df = pd.read_excel(file_path, skiprows=3)

    kml = Kml()

    for index, row in df.iterrows():
        try:
            start_str = clean_coord(row["Начало"])
            end_str = clean_coord(row["Конец"])
            title = str(row["ID ODH"]).strip()[:-2]
            length = str(row["L общая"]).strip()
            name = str(row["Наименование"]).strip()

            lat1, lon1 = map(float, start_str.split(","))
            lat2, lon2 = map(float, end_str.split(","))

            kml.newpoint(
                name=f"{title}, {length}",
                coords=[(lon1, lat1)],
                description=f"Название: {name}"
            )
            kml.newpoint(
                name=f"{title}, {length}",
                coords=[(lon2, lat2)],
                description=f"Название: {name}"
            )
            kml.newlinestring(
                name=title,
                coords=[(lon1, lat1), (lon2, lat2)]
            )
        except Exception as e:
            print(f"[❌] Ошибка в строке {index}: {e}")
            continue

    output_path = file_path.replace(".xlsx", ".kml")
    kml.save(output_path)
    return output_path

# --- Обработчик Excel-файла ---
async def handle_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    document = update.message.document
    file_path = f"./{document.file_name}"

    # Скачиваем файл
    new_file = await context.bot.get_file(document.file_id)
    await new_file.download_to_drive(file_path)

    # Генерация KML
    kml_path = create_kml_from_excel(file_path)

    # Отправляем KML-файл
    await update.message.reply_document(document=open(kml_path, "rb"))

    # Кнопка-ссылка на Яндекс.Конструктор
    await update.message.reply_text(
        "✅ KML-файл готов. Загрузи его сюда, чтобы создать карту в 1 клик:",
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("🗺 Открыть Конструктор Яндекс.Карт", url="https://yandex.ru/maps/tools/constructor/")]
        ])
    )

    os.remove(file_path)
    os.remove(kml_path)

# --- Запуск бота ---
if __name__ == "__main__":
    import asyncio

    TOKEN = "5642326211:AAFY4jNcl-E4OhxEyaEWJ--9tyw4sHagGgk"  # <-- Замени на токен своего бота

    app = ApplicationBuilder().token(TOKEN).build()
    app.add_handler(MessageHandler(filters.Document.ALL, handle_excel))

    print("🤖 Бот запущен...")
    asyncio.run(app.run_polling())
