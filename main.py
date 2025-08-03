import asyncio
import logging
import re
from pymongo import MongoClient
from gridfs import GridFS
from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import Command
from aiogram.filters.state import State, StatesGroup
from aiogram.fsm.context import FSMContext
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton
from playwright.async_api import async_playwright
import nest_asyncio
from io import BytesIO
import aiohttp
import sys
from openpyxl import Workbook
from tempfile import NamedTemporaryFile
from aiogram.types import FSInputFile
import os
from pathlib import Path
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google.auth.transport.requests import Request 
import io 
from aiogram.types import BotCommand
from dotenv import load_dotenv
from aiogram import Bot, Dispatcher
from aiogram.types import Message
import json
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from apscheduler.triggers.cron import CronTrigger
import logging

sys.stdout.reconfigure(encoding='utf-8')

nest_asyncio.apply()

load_dotenv()

BOT_TOKEN = os.getenv("BOT_TOKEN")  # الحصول على توكن البوت
MONGO_URI = os.getenv("MONGO_URI")  # الحصول على URI الخاص بـ MongoDB
GOOGLE_DRIVE_FOLDER_ID = os.getenv("GOOGLE_DRIVE_FOLDER_ID")  

# Настройка бота
bot = Bot(token=BOT_TOKEN)
storage = MemoryStorage()
dp = Dispatcher(storage=storage)
logging.basicConfig(level=logging.INFO)

# Подключение к MongoDB
try:
    client = MongoClient(MONGO_URI)
    client.admin.command('ping')  # Проверка соединения
    print("Подключение к MongoDB успешно!")
    users_collection = client['registration_db']['users']  # Инициализация коллекции
    fs = GridFS(client['registration_db'])  # Инициализация GridFS
except Exception as e:
    print(f"Ошибка подключения: {e}")

# Состояния
class Registration(StatesGroup):
    full_name = State()
    phone = State()
    passport_photo = State()
    medical_book = State()
    fluorography = State()
    inn_check = State()
    inn = State()
    not_self_employed = State()                  
    wait_registration_confirmation = State()     
    finish = State()

def save_to_mongodb(data):
    try:
        users_collection.insert_one(data)
    except Exception as e:
        print(f"Ошибка сохранения в MongoDB: {e}")

async def set_bot_commands(bot: Bot):
    commands = [
        BotCommand(command="start", description="начать сначала")
    ]
    await bot.set_my_commands(commands)

# ---------- إعداد Google Drive OAuth ----------
SCOPES = ['https://www.googleapis.com/auth/drive.file']

def get_gdrive_service():
    creds = None
    creds_json = os.getenv("GOOGLE_CREDENTIALS_JSON")
    if creds_json:
        creds_data = json.loads(creds_json)
        creds = Credentials.from_authorized_user_info(info=creds_data, scopes=SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # هنا ممكن تستعمل ملف client_secrets لو عندك أو تتعامل مع flow بطريقة تانية
            # لكن لو ما عندكش ملف ولا تريد واجهة، استعمل exception أو حلول خاصة
            raise Exception("Credentials not valid and no refresh token available")

    service = build('drive', 'v3', credentials=creds)
    return service

def upload_all_gridfs_images_to_gdrive():
    service = get_gdrive_service()
    for file in fs.find():
        filename = file.filename
        file_data = file.read()
        media = MediaIoBaseUpload(io.BytesIO(file_data), mimetype='image/jpeg', resumable=True)
        file_metadata = {'name': filename}
        uploaded_file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id'
        ).execute()
        print(f"The file has been uploaded: {filename} | ID: {uploaded_file.get('id')}")

async def delete_webhook():
    async with aiohttp.ClientSession() as session:
        async with session.get(f"https://api.telegram.org/bot{BOT_TOKEN}/deleteWebhook") as response:
            result = await response.json()
            if result.get("ok"):
                print("Webhook удалён успешно!")
            else:
                print("Не удалось удалить Webhook:", result)


async def save_photo_to_gridfs_and_disk(file_id: str) -> str:
    try:
        file = await bot.get_file(file_id)

        # تحميل الصورة من تلغرام
        photo_data = BytesIO()
        await bot.download_file(file.file_path, photo_data)
        photo_data.seek(0)

        # حفظ الصورة في مجلد محلي (اختياري)
        local_folder = Path(r"C:\Users\user\Desktop\БОТ\Паспорта")
        local_folder.mkdir(parents=True, exist_ok=True)
        local_path = local_folder / f"passport_{file_id}.jpg"
        with open(local_path, "wb") as f:
            f.write(photo_data.getbuffer())

        # حفظ الصورة في GridFS
        photo_data.seek(0)
        fs.put(photo_data, filename=f"passport_{file_id}.jpg")

        # رفع إلى Google Drive
        photo_data.seek(0)  # مهم جدًا
        service = get_gdrive_service()
        media = MediaIoBaseUpload(photo_data, mimetype='image/jpeg', resumable=True)

        folder_id = "1JnUNMp6c5ulQcrzQZwqng80397RMm-DY"
        file_metadata = {
            'name': f"passport_{file_id}.jpg",
            'parents': [folder_id]
        }

        uploaded_file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id'
        ).execute()

        print(f"✅ The image has been uploaded to Google Drive inside the folder.ID: {uploaded_file.get('id')}")

        return uploaded_file.get('id')  # نرجع Google Drive ID بدلًا من GridFS فقط

    except Exception as e:
        logging.error(f"Ошибка сохранения фото: {e}")
        raise



@dp.message(Command("start"))
async def start_cmd(message: types.Message, state: FSMContext):
    if message.from_user.id == (await bot.get_me()).id:
        await message.answer("⚠️ Error: The bot cannot register itself.")
        return
    await message.answer("Введите ваше ФИО полностью:")
    asyncio.create_task(delayed_notify_user(message.from_user.id))
    await state.set_state(Registration.full_name)


@dp.message(Registration.full_name)
async def process_name(message: types.Message, state: FSMContext):
    full_name = message.text
    await state.update_data(full_name=full_name)
    users_collection.update_one(
        {"_id": message.from_user.id},
        {"$set": {
            "registration_step": "full_name",
            "full_name": full_name
        }},
        upsert=True
    )
    await message.answer("Введите ваш номер телефона (формат: 8XXXXXXXXXX):")
    await state.set_state(Registration.phone)

@dp.message(Registration.phone)
async def process_phone(message: types.Message, state: FSMContext):
    phone = message.text
    if re.match(r'^\d{11}$', phone):
        await state.update_data(phone=phone)
        users_collection.update_one(
            {"_id": message.from_user.id},
            {"$set": {
                "registration_step": "phone",
                "phone": phone
            }},
            upsert=True
        )
        await message.answer("Отправьте фото 1 и 2 страниц паспорта:")
        await state.set_state(Registration.passport_photo)
    else:
        await message.answer("Неверный формат телефона. Введите в формате 8XXXXXXXXXX.")

@dp.message(Registration.passport_photo)
async def process_passport_photo(message: types.Message, state: FSMContext):
    if message.photo:
        try:
            file_id = message.photo[-1].file_id
            # استدعاء الدالة الجديدة التي تحفظ الصورة محلياً و في GridFS
            photo_id = await save_photo_to_gridfs_and_disk(file_id)
            
            await state.update_data(passport_photo=photo_id)
            users_collection.update_one(
                {"_id": message.from_user.id},
                {"$set": {
                    "registration_step": "passport_photo",
                    "passport_photo_id": photo_id
                }},
                upsert=True
            )
            
            keyboard = InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text="Да", callback_data="medical_yes")],
                [InlineKeyboardButton(text="Нет", callback_data="medical_no")]
            ])
            await message.answer("Есть ли у вас медкнижка?", reply_markup=keyboard)
            await state.set_state(Registration.medical_book)
        except Exception as e:
            await message.answer("⚠️ Ошибка сохранения фото. Попробуйте еще раз.")
            logging.error(f"Photo save error: {e}")
    else:
        await message.answer("Пожалуйста, отправьте фото паспорта.")

@dp.callback_query(Registration.medical_book, F.data == "medical_yes")
async def medical_yes(callback: types.CallbackQuery, state: FSMContext):
    await state.update_data(medical_book=True)
    users_collection.update_one(
    {"_id": callback.from_user.id},
    {"$set": {
        "registration_step": "medical_yes",
        "medical_book": True
    }},
    upsert=True
)

    await ask_fluorography(callback.message, state)

@dp.callback_query(Registration.medical_book, F.data == "medical_no")
async def medical_no(callback: types.CallbackQuery, state: FSMContext):
    await state.update_data(medical_book=False)
    users_collection.update_one(
    {"_id": callback.from_user.id},
    {"$set": {
        "registration_step": "medical_no",
        "medical_book": False     
    }},
    upsert=True
)

    await ask_fluorography(callback.message, state)

async def ask_fluorography(message: types.Message, state: FSMContext):
    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="Да", callback_data="fluorography_yes")],
        [InlineKeyboardButton(text="Нет", callback_data="fluorography_no")]
    ])
    await message.answer("Есть ли у вас флюорография?", reply_markup=keyboard)
    await state.set_state(Registration.fluorography)

@dp.callback_query(Registration.fluorography, F.data == "fluorography_yes")
async def fluorography_yes(callback: types.CallbackQuery, state: FSMContext):
    await state.update_data(fluorography=True)
    users_collection.update_one(
    {"_id": callback.from_user.id},
    {"$set": {
        "registration_step": "fluorography_yes",
        "fluorography": True
    }},
    upsert=True
)

    await ask_inn(callback.message, state)

@dp.callback_query(Registration.fluorography, F.data == "fluorography_no")
async def fluorography_no(callback: types.CallbackQuery, state: FSMContext):
    await state.update_data(fluorography=False)
    users_collection.update_one(
    {"_id": callback.from_user.id},
    {"$set": {
        "registration_step": "fluorography_no",
        "fluorography": False
    }},
    upsert=True
)

    await ask_inn(callback.message, state)

async def ask_inn(message: types.Message, state: FSMContext):
    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="Да", callback_data="inn_yes")],
        [InlineKeyboardButton(text="Нет", callback_data="inn_no")]
    ])
    await message.answer("Есть ли у вас ИНН?", reply_markup=keyboard)
    await state.set_state(Registration.inn_check)

@dp.callback_query(Registration.inn_check, F.data == "inn_no")
async def inn_no(callback: types.CallbackQuery, state: FSMContext):
    await callback.message.answer("Спасибо, что выбрали компанию Doker Prof-работа для каждого! Следите за заявками в группе вотсапп (https://chat.whatsapp.com/LKOL6wdu3R7KnLETuG0eiE)")
    users_collection.update_one(
        {"_id": callback.from_user.id},
        {"$set": {
            "registration_step": "inn_no",
            "has_inn":False
        }},
        upsert=True
    )
    await state.set_state(Registration.inn)

@dp.callback_query(Registration.inn_check, F.data == "inn_yes")
async def inn_yes(callback: types.CallbackQuery, state: FSMContext):
    await callback.message.answer("Введите ваш ИНН:")
    users_collection.update_one(
        {"_id": callback.from_user.id},
        {"$set": {
            "registration_step": "inn_yes",
            "has_inn": True
        }},
        upsert=True
    )
    await state.set_state(Registration.inn)

@dp.message(Registration.inn)
async def process_inn(message: types.Message, state: FSMContext):
    inn = message.text.strip()
    if re.match(r'^\d{12}$', inn):
        await state.update_data(inn=inn)
        users_collection.update_one(
            {"_id": message.from_user.id},
            {"$set": {
                "registration_step": "inn",
                "inn": inn
            }},
            upsert=True
        )

        result_text = await check_inn_with_playwright(inn)

        print(f"[process_inn] Received result_text: {repr(result_text)}")
        result_lower = result_text.lower()

        if "не является плательщиком" in result_lower:
            keyboard = InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text="✅ Да, хочу зарегистрироваться", callback_data="register_yes")],
                [InlineKeyboardButton(text="❌ Нет", callback_data="register_no")]
            ])
            await message.answer("❌ ИНН действителен, но не самозанятый. ГОТОВЫ ЗАРЕГИСТРИРОВАТЬСЯ? (ДА/НЕТ)", reply_markup=keyboard)
            await state.set_state(Registration.not_self_employed)
        elif "является плательщиком" in result_lower:
            await message.answer(
                "✅ ИНН подтвержден: самозанятый. Вот инструкция: https://disk.360.yandex.ru/i/yOVc8EApP5-B9w \nОзнакомьтесь с информацией ниже:",
                reply_markup=InlineKeyboardMarkup(inline_keyboard=[
                    [InlineKeyboardButton(text="Ознакомился", callback_data="acknowledged")]
                ])
            )
        elif "указан некорректный инн" in result_lower:
            await message.answer("⚠️ Указан некорректный ИНН. Пожалуйста, проверьте и введите снова.")
        elif "превышено количество запросов" in result_lower:
            await message.answer("⚠️ Слишком много запросов. Попробуйте позже.")
        elif "не удалось проверить" in result_lower or "не найден" in result_lower:
            await message.answer("⚠️ ИНН не найден или недействителен.")
        else:
            await message.answer(f"⚠️ Не удалось определить статус: {result_text}")
    else:
        await message.answer("⚠️ Неверный формат ИНН. Введите 12 цифр.")

@dp.callback_query(F.data == "acknowledged")
async def acknowledged_callback(callback: types.CallbackQuery, state: FSMContext):
    await callback.message.answer("✅ Отлично! Через 10 минут вы получите ссылку на договор.")
    
    # ⏳ انتظار 10 دقائق ثم إرسال العقد
    asyncio.create_task(send_contract_link(callback.from_user.id))

    # ✅ تحديث التسجيل
    users_collection.update_one(
        {"_id": callback.from_user.id},
        {"$set": {
            "registration_step": "self_employed_acknowledged",
            "self_employed": True
        }},
        upsert=True
    )
async def send_contract_link(user_id: int):
    await asyncio.sleep(5)
    await bot.send_message(user_id, "Ссылка на договор с самозанятым: [https://pro.selfwork.ru/connect/B/dokerprof])")
    await bot.send_message(user_id, "Спасибо, что выбрали компанию Doker Prof-работа для каждого! Следите за заявками в группе вотсапп (https://chat.whatsapp.com/LKOL6wdu3R7KnLETuG0eiE)")



@dp.callback_query(Registration.not_self_employed, F.data == "register_no")
async def handle_register_no(callback: types.CallbackQuery, state: FSMContext):
    await callback.message.answer("Спасибо, что выбрали компанию Doker Prof-работа для каждого! Следите за заявками в группе вотсапп (https://chat.whatsapp.com/LKOL6wdu3R7KnLETuG0eiE)")
    users_collection.update_one(
        {"_id": callback.from_user.id},
        {"$set": {
            "registration_step": "inn_no",
            "has_inn": False
        }},
        upsert=True
    )
    await state.clear()

@dp.callback_query(Registration.not_self_employed, F.data == "register_yes")
async def handle_register_yes(callback: types.CallbackQuery, state: FSMContext):
    await callback.message.answer(
        ":📄 ИНСТРУКЦИЯ: КАК СТАТЬ САМОЗАНЯТЫМ ЗА 5 МИНУТ \nчерез смартфон (самый удобный) \n1. Скачай приложение «Мой Налог»: \n• 📱 Для Android: https://play.google.com/store/apps/details?id=ru.fns.taxselfemployed" \
        "\n• 📱 Для iPhone: https://apps.apple.com/ru/app/мой-налог/id1476104524 \n2. Открой приложение и нажмите “Войти” \nМожно авторизоваться через: \n• Госуслуги" \
        "\n• Паспорт РФ + Селфи (если нет Госуслуг) \n3. Подтверди регистрацию как самозанятый \n4. Готово! Ты официально зарегистрирован как самозанятый!",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[
    [InlineKeyboardButton(text="✅ Я зарегистрировался", callback_data="registered_done")],
    [InlineKeyboardButton(text="❌ Не получилось зарегистрироваться?", callback_data="registration_failed")]
    ])
    )
    await state.set_state(Registration.wait_registration_confirmation)

@dp.callback_query(Registration.wait_registration_confirmation, F.data == "registration_failed")
async def handle_registration_failed(callback: types.CallbackQuery, state: FSMContext):
    await callback.message.answer("📲 *Напиши нам в WhatsApp:* 👉 https://wa.me/79920294851 .\n📞Либо приезжайте в офис по адресу : Сони Кривой 83 Бизнес Центр ПОЛЕТ 8 этаж офис 805 с 11:00-15:00 с Собой \n-Паспорт \n-Инн \n-Снилс")
    users_collection.update_one(
        {"_id": callback.from_user.id},
        {"$set": {"registration_step": "registration_failed"}},
        upsert=True
    )

@dp.callback_query(Registration.wait_registration_confirmation, F.data == "registered_done")
async def handle_registered_done(callback: types.CallbackQuery, state: FSMContext):
    await finalize_registration(callback.message, state)  
    await callback.message.answer("Отлично! Теперь введите ваш ИНН заново:")
    await state.set_state(Registration.inn)
    users_collection.update_one(
        {"_id": callback.from_user.id},
        {"$set": {
            "registration_step": "inn_yes",
            "has_inn": True
        }},
        upsert=True
    )

async def save_to_mongodb(data):
    try:
        required_fields = {"full_name", "phone", "medical_book", "fluorography", "inn", "passport_photo"}
        for field in required_fields:
            if field not in data:
                print(f"Missing required field: {field}")
                return

        await users_collection.insert_one(data)
        print("User  data saved successfully:", data)
    except Exception as e:
        print(f"Ошибка сохранения в MongoDB: {e}")

async def finalize_registration(message: types.Message, state: FSMContext):
    data = await state.get_data()
    user_data = {
        "full_name": data.get("full_name"),
        "phone": data.get("phone"),
        "medical_book": data.get("medical_book"),
        "fluorography": data.get("fluorography"),
        "inn": data.get("inn"),
        "passport_photo_id": data.get("passport_photo")
    }
    print("Finalizing registration with data:", user_data)
    users_collection.update_one(
        {"_id": message.from_user.id},
        {"$set": {
            **user_data,
            "registration_step": "finish"
        }},
        upsert=True
    )

    await state.clear()

async def check_inn_with_playwright(inn: str) -> str:
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()
        try:
            print(f"[ИНН CHECK] Checking ИНН: {inn}")
            await page.goto("https://самозанятые.рф/check-inn")
            await page.fill('input[placeholder="12 цифр"]', inn)

            is_disabled = await page.is_disabled('button:has-text("Узнать статус")')
            if is_disabled:
                return "⚠️ ИНН неактивен. Пожалуйста, проверьте корректность."

            await page.click('button:has-text("Узнать статус")')

            # Вместо fixed wait, ждем появления результата
            await page.wait_for_selector('span.style-module--text-result-success--b17b5, span.style-module--text-result-error--d739c, div.result-text', timeout=7000)

            possible_selectors = [
                'span.style-module--text-result-success--b17b5',
                'span.style-module--text-result-error--d739c',
                'div.result-text',
                'span[class*="text-result"]',
            ]

            result_text = ""
            for selector in possible_selectors:
                elements = await page.query_selector_all(selector)
                for el in elements:
                    try:
                        text = await el.inner_text()
                        if text.strip():
                            result_text = text.strip()
                            break
                    except Exception:
                        continue
                if result_text:
                    break

            if not result_text:
                body_text = await page.inner_text("body")
                print("[ИНН CHECK] Body snapshot (first 300 chars):", repr(body_text[:300]))
                return "⚠️ Не удалось определить статус: результат не найден."

            print("[ИНН CHECK] Result text:", repr(result_text))

            return result_text

        except Exception as e:
            print("[ИНН CHECK] Error:", repr(e))
            return "⚠️ Произошла ошибка при проверке ИНН."
        finally:
            await browser.close()

async def delayed_notify_user(user_id: int):
    await asyncio.sleep(3600)  # انتظار ساعة
    try:
        user = users_collection.find_one({"_id": user_id})
        if user and user.get("registration_step") != "finish":
            await bot.send_message(user_id, "⌛ Вы не завершили регистрацию. Мы всё равно добавим вас в список работников. Спасибо!")
    except Exception as e:
        logging.error(f"Ошибка при отправке напоминания пользователю {user_id}: {e}")

# أسماء الخطوات
step_names = {
    "full_name": "📝 Ввод ФИО",
    "phone": "📞 Номер телефона",
    "passport_photo": "📸 Фото паспорта",
    "medical_yes": "✅ Медкнижка: Да",
    "medical_no": "❌ Медкнижка: Нет",
    "fluorography_yes": "✅ Флюорография: Да",
    "fluorography_no": "❌ Флюорография: Нет",
    "inn_yes": "✅ Есть ИНН",
    "inn_no": "❌ Нет ИНН",
    "inn": "💳 Проверка ИНН",
    "register_yes": "✅ Зарегистрируется",
    "register_no": "❌ Не будет регистрироваться",
    "registered_done": "✅ Подтвердил регистрацию",
    "registration_failed": "❌ Не смог зарегистрироваться"
}

# الوظيفة اللي تبعت التقرير
async def send_daily_report():
    try:
        step_counts = {step: users_collection.count_documents({"registration_step": step}) for step in step_names.keys()}

        wb = Workbook()
        ws_summary = wb.active
        ws_summary.title = "Отчет"
        ws_summary.append(["Шаг", "Описание", "Кол-во"])
        for step, desc in step_names.items():
            ws_summary.append([step, desc, step_counts.get(step, 0)])

        ws_detail = wb.create_sheet("Пользователи")
        ws_detail.append(["ID", "Имя", "Телефон", "ИНН", "Медкнижка", "Флюорография", "Шаг регистрации"])

        all_users = users_collection.find()
        for user in all_users:
            ws_detail.append([
                user.get("_id", ""),
                user.get("full_name", ""),
                user.get("phone", ""),
                user.get("inn", ""),
                "Да" if user.get("medical_book") else "Нет",
                "Да" if user.get("fluorography") else "Нет",
                step_names.get(user.get("registration_step", ""), user.get("registration_step", ""))
            ])

        with NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            wb.save(tmp.name)
            tmp.seek(0)

            await bot.send_document(
                chat_id=1085716060,
                document=FSInputFile(tmp.name, filename="full_daily_report.xlsx"),
                caption="📊 Подробный отчет по регистрации (2 листа)"
            )

    except Exception as e:
        logging.error(f"Ошибка при создании Excel: {e}")

scheduler = AsyncIOScheduler()


async def main():
    # إعداد الأوامر
    await set_bot_commands(bot)

    # جدولة إرسال التقرير يوميًا الساعة 22:00
    scheduler.add_job(send_daily_report, CronTrigger(hour=22, minute=0))
    scheduler.start()

    # بدء تشغيل البوت
    await dp.start_polling(bot)

if __name__ == "__main__":
    import asyncio
    asyncio.run(main())
