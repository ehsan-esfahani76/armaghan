import os
import json
import shutil
import datetime
import threading

import pandas as pd
import jdatetime

from zoneinfo import ZoneInfo
from dotenv import load_dotenv
from flask import Flask

from telegram import (
    Update,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
)

from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ContextTypes,
    filters,
)


# =========================================================
# تنظیمات
# =========================================================

load_dotenv()

BOT_TOKEN = os.getenv("BOT_TOKEN")
OWNER_ID_TEXT = os.getenv("OWNER_ID")


if not BOT_TOKEN:
    raise ValueError(
        "BOT_TOKEN در Environment تنظیم نشده است."
    )


if not OWNER_ID_TEXT:
    raise ValueError(
        "OWNER_ID در Environment تنظیم نشده است."
    )


try:
    OWNER_ID = int(OWNER_ID_TEXT)

except ValueError:
    raise ValueError(
        "OWNER_ID باید یک عدد باشد."
    )


# =========================================================
# مسیر فایل‌ها
# =========================================================

DATA_DIR = "data"

EXCEL_FILE = os.path.join(
    DATA_DIR,
    "current_prices.xlsx"
)

BACKUP_FILE = os.path.join(
    DATA_DIR,
    "backup_prices.xlsx"
)

ADMINS_FILE = os.path.join(
    DATA_DIR,
    "admins.json"
)

UPLOAD_INFO_FILE = os.path.join(
    DATA_DIR,
    "upload_info.json"
)


PAGE_SIZE = 10


# ساخت پوشه data
os.makedirs(
    DATA_DIR,
    exist_ok=True
)


# =========================================================
# HTTP Server برای Render
# =========================================================

web_app = Flask(__name__)


@web_app.route("/")
def home():

    return "Bot is running!"


@web_app.route("/health")
def health():

    return "OK"


def run_web_server():

    port = int(
        os.environ.get(
            "PORT",
            10000
        )
    )

    web_app.run(
        host="0.0.0.0",
        port=port
    )


# =========================================================
# اعداد انگلیسی به فارسی
# =========================================================

def to_persian_digits(text):

    english = "0123456789"
    persian = "۰۱۲۳۴۵۶۷۸۹"

    table = str.maketrans(
        english,
        persian
    )

    return str(text).translate(table)


# =========================================================
# مدیریت ادمین‌ها
# =========================================================

def load_admins():

    if not os.path.exists(
        ADMINS_FILE
    ):
        return []

    try:

        with open(
            ADMINS_FILE,
            "r",
            encoding="utf-8"
        ) as file:

            data = json.load(file)

        if not isinstance(
            data,
            list
        ):
            return []

        return [
            int(user_id)
            for user_id in data
        ]

    except Exception as error:

        print(
            "خطا در خواندن admins.json:",
            error
        )

        return []


def save_admins(admins):

    admins = list(
        dict.fromkeys(
            int(user_id)
            for user_id in admins
        )
    )

    with open(
        ADMINS_FILE,
        "w",
        encoding="utf-8"
    ) as file:

        json.dump(
            admins,
            file,
            ensure_ascii=False,
            indent=2
        )


def is_owner(user_id):

    return user_id == OWNER_ID


def is_admin(user_id):

    if user_id == OWNER_ID:
        return True

    return user_id in load_admins()


# =========================================================
# ذخیره تاریخ آخرین آپلود
# =========================================================

def save_upload_info(
    uploader_id=None
):

    iran_time = datetime.datetime.now(
        ZoneInfo("Asia/Tehran")
    )

    jalali = jdatetime.datetime.fromgregorian(
        datetime=iran_time
    )

    upload_info = {

        "date": jalali.strftime(
            "%Y/%m/%d"
        ),

        "time": jalali.strftime(
            "%H:%M"
        ),

        "timestamp": iran_time.isoformat(),

        "uploader_id": uploader_id
    }

    with open(
        UPLOAD_INFO_FILE,
        "w",
        encoding="utf-8"
    ) as file:

        json.dump(
            upload_info,
            file,
            ensure_ascii=False,
            indent=2
        )


# =========================================================
# خواندن تاریخ آخرین آپلود
# =========================================================

def get_upload_info():

    if not os.path.exists(
        UPLOAD_INFO_FILE
    ):
        return None

    try:

        with open(
            UPLOAD_INFO_FILE,
            "r",
            encoding="utf-8"
        ) as file:

            data = json.load(file)

        date = data.get(
            "date"
        )

        time = data.get(
            "time"
        )

        uploader_id = data.get(
            "uploader_id"
        )

        if not date or not time:
            return None

        return {
            "date": to_persian_digits(date),
            "time": to_persian_digits(time),
            "uploader_id": uploader_id
        }

    except Exception as error:

        print(
            "خطا در خواندن upload_info.json:",
            error
        )

        return None


# =========================================================
# فرمت قیمت
# =========================================================

def format_price(value):

    try:

        number = float(value)

        if number.is_integer():

            return f"{int(number):,}"

        return f"{number:,.2f}"

    except Exception:

        return str(value)


# =========================================================
# خواندن Excel
# =========================================================

def load_prices():

    if not os.path.exists(
        EXCEL_FILE
    ):
        return None

    try:

        df = pd.read_excel(
            EXCEL_FILE
        )

        # حذف ستون‌های اضافی Excel
        df = df.loc[
            :,
            ~df.columns.astype(str)
            .str.startswith("Unnamed")
        ]

        # تمیز کردن نام ستون‌ها
        df.columns = [
            str(column).strip()
            for column in df.columns
        ]

        # نام‌های قابل قبول ستون‌ها
        column_mapping = {

            "کد": "code",
            "کد کالا": "code",
            "کدکالا": "code",
            "کد محصول": "code",

            "code": "code",
            "Code": "code",
            "CODE": "code",

            "نام": "name",
            "نام کالا": "name",
            "نامکالا": "name",
            "کالا": "name",
            "نام محصول": "name",

            "name": "name",
            "Name": "name",
            "NAME": "name",

            "قیمت": "price",
            "قیمت کالا": "price",
            "قیمت فروش": "price",
            "قیمت‌فروش": "price",
            "قیمت محصول": "price",

            "price": "price",
            "Price": "price",
            "PRICE": "price",
        }

        df = df.rename(
            columns={
                column:
                column_mapping.get(
                    column,
                    column
                )
                for column in df.columns
            }
        )

        required_columns = [
            "code",
            "name",
            "price"
        ]

        for column in required_columns:

            if column not in df.columns:

                print(
                    f"ستون مورد نیاز پیدا نشد: {column}"
                )

                return None

        df = df[
            [
                "code",
                "name",
                "price"
            ]
        ]

        df = df.fillna("")

        df["code"] = (
            df["code"]
            .astype(str)
            .str.strip()
        )

        df["name"] = (
            df["name"]
            .astype(str)
            .str.strip()
        )

        df = df[
            ~(
                (df["code"] == "")
                &
                (df["name"] == "")
            )
        ]

        return df.reset_index(
            drop=True
        )

    except Exception as error:

        print(
            "خطا در خواندن Excel:",
            error
        )

        return None


# =========================================================
# نمایش یک کالا
# =========================================================

def product_text(
    row,
    number=None
):

    code = str(
        row["code"]
    ).strip()

    name = str(
        row["name"]
    ).strip()

    price = format_price(
        row["price"]
    )

    text = ""

    if number is not None:

        text += (
            f"{number}. "
        )

    text += (
        f"📦 {name}\n"
        f"🔢 کد: {code}\n"
        f"💰 قیمت: {price} تومان"
    )

    return text


# =========================================================
# صفحه‌بندی
# =========================================================

def pagination_keyboard(
    current_page,
    total_pages
):

    buttons = []

    if current_page > 0:

        buttons.append(
            InlineKeyboardButton(
                "◀️ قبلی",
                callback_data=
                f"page:{current_page - 1}"
            )
        )

    if current_page < total_pages - 1:

        buttons.append(
            InlineKeyboardButton(
                "بعدی ▶️",
                callback_data=
                f"page:{current_page + 1}"
            )
        )

    keyboard = []

    if buttons:

        keyboard.append(
            buttons
        )

    keyboard.append(
        [
            InlineKeyboardButton(
                "🔎 جستجوی کالا",
                callback_data="search"
            )
        ]
    )

    keyboard.append(
        [
            InlineKeyboardButton(
                "🏠 منوی اصلی",
                callback_data="home"
            )
        ]
    )

    return InlineKeyboardMarkup(
        keyboard
    )


# =========================================================
# ساخت صفحه لیست قیمت
# =========================================================

def build_price_page(
    page=0
):

    df = load_prices()

    if df is None or df.empty:

        return None, None

    total_products = len(df)

    total_pages = (
        total_products
        + PAGE_SIZE
        - 1
    ) // PAGE_SIZE

    page = max(
        0,
        min(
            page,
            total_pages - 1
        )
    )

    start = (
        page * PAGE_SIZE
    )

    end = (
        start + PAGE_SIZE
    )

    current_products = df.iloc[
        start:end
    ]

    text = (
        "📋 لیست قیمت\n\n"
    )

    upload_info = get_upload_info()

    if upload_info:

        text += (
            "📅 آخرین بروزرسانی: "
            f"{upload_info['date']} "
            f"- "
            f"{upload_info['time']}\n\n"
        )

    text += (
        f"📄 صفحه "
        f"{to_persian_digits(page + 1)} "
        f"از "
        f"{to_persian_digits(total_pages)}\n\n"
    )

    for index, (_, row) in enumerate(
        current_products.iterrows(),
        start=start + 1
    ):

        text += product_text(
            row,
            to_persian_digits(index)
        )

        text += "\n\n"

    return (
        text,
        pagination_keyboard(
            page,
            total_pages
        )
    )


# =========================================================
# منوی اصلی
# =========================================================

def main_menu(
    user_id
):

    keyboard = [

        [
            InlineKeyboardButton(
                "📋 لیست قیمت",
                callback_data="prices"
            )
        ],

        [
            InlineKeyboardButton(
                "🔎 جستجوی کالا",
                callback_data="search"
            )
        ]
    ]

    if is_admin(
        user_id
    ):

        keyboard.append(
            [
                InlineKeyboardButton(
                    "👑 پنل مدیریت",
                    callback_data="admin_panel"
                )
            ]
        )

    keyboard.append(
        [
            InlineKeyboardButton(
                "ℹ️ راهنما",
                callback_data="help"
            )
        ]
    )

    return InlineKeyboardMarkup(
        keyboard
    )


# =========================================================
# پنل مدیریت
# =========================================================

def admin_panel_keyboard(
    user_id
):

    keyboard = [

        [
            InlineKeyboardButton(
                "📤 آپلود Excel",
                callback_data="admin_upload"
            )
        ],

        [
            InlineKeyboardButton(
                "📊 وضعیت لیست قیمت",
                callback_data="status"
            )
        ]
    ]

    if is_owner(
        user_id
    ):

        keyboard.append(
            [
                InlineKeyboardButton(
                    "➕ افزودن ادمین",
                    callback_data="add_admin"
                ),

                InlineKeyboardButton(
                    "➖ حذف ادمین",
                    callback_data="remove_admin"
                )
            ]
        )

        keyboard.append(
            [
                InlineKeyboardButton(
                    "👥 لیست ادمین‌ها",
                    callback_data="list_admins"
                )
            ]
        )

    keyboard.append(
        [
            InlineKeyboardButton(
                "🏠 منوی اصلی",
                callback_data="home"
            )
        ]
    )

    return InlineKeyboardMarkup(
        keyboard
    )


# =========================================================
# /start
# =========================================================

async def start(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    user_id = (
        update.effective_user.id
    )

    await update.message.reply_text(
        "سلام 👋\n\n"
        "به ربات لیست قیمت خوش آمدید.\n\n"
        "لطفاً یکی از گزینه‌های زیر را انتخاب کنید:",
        reply_markup=main_menu(
            user_id
        )
    )


# =========================================================
# /id
# =========================================================

async def my_id(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    user_id = (
        update.effective_user.id
    )

    await update.message.reply_text(
        "🆔 آیدی عددی تلگرام شما:\n\n"
        f"{to_persian_digits(user_id)}"
    )


# =========================================================
# /price
# =========================================================

async def price_command(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    text, keyboard = (
        build_price_page(
            0
        )
    )

    if text is None:

        await update.message.reply_text(
            "❌ هنوز فایل لیست قیمت آپلود نشده است."
        )

        return

    await update.message.reply_text(
        text,
        reply_markup=keyboard
    )


# =========================================================
# /search
# =========================================================

async def search_command(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    context.user_data[
        "searching"
    ] = True

    await update.message.reply_text(
        "🔎 لطفاً نام یا کد کالا را ارسال کنید:"
    )


# =========================================================
# /upload
# =========================================================

async def upload_command(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    user_id = (
        update.effective_user.id
    )

    if not is_admin(
        user_id
    ):

        await update.message.reply_text(
            "⛔ شما اجازه آپلود فایل ندارید."
        )

        return

    context.user_data[
        "waiting_for_excel"
    ] = True

    await update.message.reply_text(
        "📤 لطفاً فایل Excel را ارسال کنید.\n\n"
        "ستون‌های فایل باید شامل این موارد باشند:\n\n"
        "کد | نام | قیمت"
    )


# =========================================================
# /admin
# =========================================================

async def admin_command(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    user_id = (
        update.effective_user.id
    )

    if not is_admin(
        user_id
    ):

        await update.message.reply_text(
            "⛔ شما ادمین نیستید."
        )

        return

    await update.message.reply_text(
        "👑 پنل مدیریت",
        reply_markup=admin_panel_keyboard(
            user_id
        )
    )


# =========================================================
# /help
# =========================================================

async def help_command(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    text = (
        "ℹ️ راهنمای ربات\n\n"
        "/start - منوی اصلی\n"
        "/id - نمایش ID عددی شما\n"
        "/price - نمایش لیست قیمت\n"
        "/search - جستجوی کالا\n"
        "/help - راهنما\n"
    )

    if is_admin(
        update.effective_user.id
    ):

        text += (
            "\n👑 دستورات ادمین:\n"
            "/admin - پنل مدیریت\n"
            "/upload - آپلود Excel\n"
        )

    await update.message.reply_text(
        text
    )


# =========================================================
# دریافت فایل Excel
# =========================================================

async def handle_document(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    user_id = (
        update.effective_user.id
    )

    if not is_admin(
        user_id
    ):

        await update.message.reply_text(
            "⛔ فقط ادمین‌ها می‌توانند فایل Excel آپلود کنند."
        )

        return

    document = (
        update.message.document
    )

    if not document:
        return

    filename = (
        document.file_name
        or ""
    )

    if not filename.lower().endswith(
        ".xlsx"
    ):

        await update.message.reply_text(
            "❌ فقط فایل Excel با فرمت .xlsx قابل قبول است."
        )

        return

    temp_file = os.path.join(
        DATA_DIR,
        "new_prices.xlsx"
    )

    try:

        file = await document.get_file()

        await file.download_to_drive(
            temp_file
        )

        try:

            pd.read_excel(
                temp_file
            )

        except Exception:

            if os.path.exists(
                temp_file
            ):

                os.remove(
                    temp_file
                )

            await update.message.reply_text(
                "❌ فایل Excel قابل خواندن نیست."
            )

            return

        if os.path.exists(
            EXCEL_FILE
        ):

            shutil.copy2(
                EXCEL_FILE,
                BACKUP_FILE
            )

        shutil.move(
            temp_file,
            EXCEL_FILE
        )

        df = load_prices()

        if df is None:

            if os.path.exists(
                BACKUP_FILE
            ):

                shutil.copy2(
                    BACKUP_FILE,
                    EXCEL_FILE
                )

            else:

                if os.path.exists(
                    EXCEL_FILE
                ):

                    os.remove(
                        EXCEL_FILE
                    )

            await update.message.reply_text(
                "❌ ساختار فایل Excel صحیح نیست.\n\n"
                "فایل باید حداقل این سه ستون را داشته باشد:\n\n"
                "کد | نام | قیمت"
            )

            return

        save_upload_info(
            uploader_id=user_id
        )

        context.user_data[
            "waiting_for_excel"
        ] = False

        upload_info = (
            get_upload_info()
        )

        date_text = ""

        if upload_info:

            date_text = (
                f"\n📅 تاریخ بروزرسانی: "
                f"{upload_info['date']} "
                f"- "
                f"{upload_info['time']}"
            )

        await update.message.reply_text(
            "✅ فایل Excel با موفقیت بروزرسانی شد.\n\n"
            f"📦 تعداد کالاها: "
            f"{to_persian_digits(len(df))}"
            f"{date_text}"
        )

    except Exception as error:

        print(
            "Upload error:",
            error
        )

        if os.path.exists(
            temp_file
        ):

            os.remove(
                temp_file
            )

        await update.message.reply_text(
            "❌ هنگام پردازش فایل Excel خطایی رخ داد."
        )


# =========================================================
# جستجوی کالا
# =========================================================

async def search_text(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    if not context.user_data.get(
        "searching",
        False
    ):

        return

    query = (
        update.message.text
        .strip()
        .lower()
    )

    df = load_prices()

    context.user_data[
        "searching"
    ] = False

    if df is None:

        await update.message.reply_text(
            "❌ هنوز فایل قیمت آپلود نشده است."
        )

        return

    results = df[
        df["name"]
        .str.lower()
        .str.contains(
            query,
            na=False,
            regex=False
        )
        |
        df["code"]
        .str.lower()
        .str.contains(
            query,
            na=False,
            regex=False
        )
    ]

    if results.empty:

        await update.message.reply_text(
            "❌ کالایی با این نام یا کد پیدا نشد."
        )

        return

    results = results.head(20)

    text = (
        "🔎 نتایج جستجو\n\n"
    )

    upload_info = (
        get_upload_info()
    )

    if upload_info:

        text += (
            "📅 آخرین بروزرسانی: "
            f"{upload_info['date']} "
            f"- "
            f"{upload_info['time']}\n\n"
        )

    for index, (_, row) in enumerate(
        results.iterrows(),
        start=1
    ):

        text += product_text(
            row,
            to_persian_digits(index)
        )

        text += "\n\n"

    await update.message.reply_text(
        text
    )


# =========================================================
# افزودن / حذف ادمین
# =========================================================

async def handle_admin_input(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    user_id = (
        update.effective_user.id
    )

    if not is_owner(
        user_id
    ):
        return

    action = (
        context.user_data.get(
            "admin_action"
        )
    )

    if not action:
        return

    text = (
        update.message.text.strip()
    )

    # =====================================================
    # افزودن ادمین
    # =====================================================

    if action == "add":

        try:

            new_admin_id = int(
                text
            )

        except ValueError:

            await update.message.reply_text(
                "❌ ID باید عددی باشد.\n\n"
                "مثال:\n"
                "123456789"
            )

            return

        if new_admin_id == OWNER_ID:

            await update.message.reply_text(
                "ℹ️ این ID مربوط به مالک اصلی ربات است."
            )

            context.user_data.pop(
                "admin_action",
                None
            )

            return

        admins = load_admins()

        if new_admin_id in admins:

            await update.message.reply_text(
                "ℹ️ این کاربر قبلاً ادمین است."
            )

            context.user_data.pop(
                "admin_action",
                None
            )

            return

        admins.append(
            new_admin_id
        )

        save_admins(
            admins
        )

        context.user_data.pop(
            "admin_action",
            None
        )

        await update.message.reply_text(
            "✅ ادمین با موفقیت اضافه شد.\n\n"
            f"🆔 ID: "
            f"{to_persian_digits(new_admin_id)}",
            reply_markup=admin_panel_keyboard(
                user_id
            )
        )

    # =====================================================
    # حذف ادمین
    # =====================================================

    elif action == "remove":

        try:

            admin_id = int(
                text
            )

        except ValueError:

            await update.message.reply_text(
                "❌ ID باید عددی باشد."
            )

            return

        if admin_id == OWNER_ID:

            await update.message.reply_text(
                "⛔ مالک اصلی را نمی‌توان حذف کرد."
            )

            return

        admins = load_admins()

        if admin_id not in admins:

            await update.message.reply_text(
                "❌ این ID در لیست ادمین‌ها وجود ندارد."
            )

            return

        admins.remove(
            admin_id
        )

        save_admins(
            admins
        )

        context.user_data.pop(
            "admin_action",
            None
        )

        await update.message.reply_text(
            "✅ ادمین حذف شد.\n\n"
            f"🆔 ID: "
            f"{to_persian_digits(admin_id)}",
            reply_markup=admin_panel_keyboard(
                user_id
            )
        )


# =========================================================
# مدیریت دکمه‌ها
# =========================================================

async def button_handler(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    query = update.callback_query

    await query.answer()

    data = query.data

    user_id = (
        query.from_user.id
    )

    # =====================================================
    # لیست قیمت
    # =====================================================

    if data == "prices":

        text, keyboard = (
            build_price_page(
                0
            )
        )

        if text is None:

            await query.edit_message_text(
                "❌ هنوز فایل لیست قیمت آپلود نشده است.",
                reply_markup=main_menu(
                    user_id
                )
            )

            return

        await query.edit_message_text(
            text,
            reply_markup=keyboard
        )

    # =====================================================
    # صفحه‌بندی
    # =====================================================

    elif data.startswith(
        "page:"
    ):

        try:

            page = int(
                data.split(":")[1]
            )

        except ValueError:

            return

        text, keyboard = (
            build_price_page(
                page
            )
        )

        if text is None:
            return

        await query.edit_message_text(
            text,
            reply_markup=keyboard
        )

    # =====================================================
    # جستجو
    # =====================================================

    elif data == "search":

        context.user_data[
            "searching"
        ] = True

        await query.message.reply_text(
            "🔎 لطفاً نام یا کد کالا را ارسال کنید:"
        )

    # =====================================================
    # خانه
    # =====================================================

    elif data == "home":

        context.user_data.pop(
            "searching",
            None
        )

        context.user_data.pop(
            "admin_action",
            None
        )

        await query.edit_message_text(
            "🏠 منوی اصلی",
            reply_markup=main_menu(
                user_id
            )
        )

    # =====================================================
    # پنل مدیریت
    # =====================================================

    elif data == "admin_panel":

        if not is_admin(
            user_id
        ):
            return

        await query.edit_message_text(
            "👑 پنل مدیریت",
            reply_markup=admin_panel_keyboard(
                user_id
            )
        )

    # =====================================================
    # آپلود Excel
    # =====================================================

    elif data == "admin_upload":

        if not is_admin(
            user_id
        ):
            return

        context.user_data[
            "waiting_for_excel"
        ] = True

        await query.message.reply_text(
            "📤 فایل Excel را ارسال کنید.\n\n"
            "فرمت ستون‌ها:\n\n"
            "کد | نام | قیمت"
        )

    # =====================================================
    # افزودن ادمین
    # =====================================================

    elif data == "add_admin":

        if not is_owner(
            user_id
        ):

            await query.message.reply_text(
                "⛔ فقط OWNER اجازه افزودن ادمین دارد."
            )

            return

        context.user_data[
            "admin_action"
        ] = "add"

        await query.message.reply_text(
            "➕ افزودن ادمین\n\n"
            "لطفاً ID عددی تلگرام شخص را ارسال کنید.\n\n"
            "مثال:\n"
            "123456789\n\n"
            "برای پیدا کردن ID، شخص می‌تواند "
            "داخل همین ربات /id را ارسال کند."
        )

    # =====================================================
    # حذف ادمین
    # =====================================================

    elif data == "remove_admin":

        if not is_owner(
            user_id
        ):
            return

        context.user_data[
            "admin_action"
        ] = "remove"

        await query.message.reply_text(
            "➖ حذف ادمین\n\n"
            "لطفاً ID عددی ادمین را ارسال کنید."
        )

    # =====================================================
    # لیست ادمین‌ها
    # =====================================================

    elif data == "list_admins":

        if not is_owner(
            user_id
        ):
            return

        admins = load_admins()

        text = (
            "👥 لیست ادمین‌ها\n\n"
            "👑 OWNER اصلی:\n"
            f"{to_persian_digits(OWNER_ID)}\n\n"
        )

        if not admins:

            text += (
                "هنوز ادمین دیگری اضافه نشده است."
            )

        else:

            text += (
                "👤 ادمین‌ها:\n\n"
            )

            for index, admin_id in enumerate(
                admins,
                start=1
            ):

                text += (
                    f"{to_persian_digits(index)}. "
                    f"{to_persian_digits(admin_id)}\n"
                )

        await query.message.reply_text(
            text
        )

    # =====================================================
    # وضعیت
    # =====================================================

    elif data == "status":

        if not is_admin(
            user_id
        ):
            return

        df = load_prices()

        if df is None:

            text = (
                "📊 وضعیت لیست قیمت\n\n"
                "❌ فایل Excel آپلود نشده است."
            )

        else:

            upload_info = (
                get_upload_info()
            )

            text = (
                "📊 وضعیت لیست قیمت\n\n"
                f"📦 تعداد کالاها: "
                f"{to_persian_digits(len(df))}\n"
                "✅ وضعیت: فعال\n"
            )

            if upload_info:

                text += (
                    "\n📅 آخرین بروزرسانی: "
                    f"{upload_info['date']} "
                    f"- "
                    f"{upload_info['time']}"
                )

        await query.message.reply_text(
            text
        )

    # =====================================================
    # راهنما
    # =====================================================

    elif data == "help":

        await query.message.reply_text(
            "ℹ️ راهنمای ربات\n\n"
            "📋 لیست قیمت:\n"
            "نمایش کد، نام و قیمت کالاها\n\n"
            "🔎 جستجو:\n"
            "جستجو با کد یا نام کالا\n\n"
            "🆔 /id:\n"
            "نمایش ID عددی تلگرام شما\n\n"
            "👑 پنل مدیریت:\n"
            "مخصوص ادمین‌ها"
        )


# =========================================================
# مدیریت پیام‌های متنی
# =========================================================

async def text_handler(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    user_id = (
        update.effective_user.id
    )

    # =====================================================
    # مدیریت افزودن / حذف ادمین
    # =====================================================

    if (
        is_owner(user_id)
        and context.user_data.get(
            "admin_action"
        )
    ):

        await handle_admin_input(
            update,
            context
        )

        return

    # =====================================================
    # جستجو
    # =====================================================

    if context.user_data.get(
        "searching",
        False
    ):

        await search_text(
            update,
            context
        )


# =========================================================
# اجرای ربات
# =========================================================

def main():

    print(
        "🤖 Starting bot..."
    )

    # =====================================================
    # اجرای HTTP Server
    # =====================================================

    web_thread = threading.Thread(
        target=run_web_server,
        daemon=True
    )

    web_thread.start()

    print(
        "🌐 HTTP server started."
    )

    # =====================================================
    # ساخت Telegram Application
    # =====================================================

    application = (
        Application
        .builder()
        .token(BOT_TOKEN)
        .build()
    )

    # =====================================================
    # Commands
    # =====================================================

    application.add_handler(
        CommandHandler(
            "start",
            start
        )
    )

    application.add_handler(
        CommandHandler(
            "id",
            my_id
        )
    )

    application.add_handler(
        CommandHandler(
            "price",
            price_command
        )
    )

    application.add_handler(
        CommandHandler(
            "search",
            search_command
        )
    )

    application.add_handler(
        CommandHandler(
            "upload",
            upload_command
        )
    )

    application.add_handler(
        CommandHandler(
            "admin",
            admin_command
        )
    )

    application.add_handler(
        CommandHandler(
            "help",
            help_command
        )
    )

    # =====================================================
    # Buttons
    # =====================================================

    application.add_handler(
        CallbackQueryHandler(
            button_handler
        )
    )

    # =====================================================
    # Excel files
    # =====================================================

    application.add_handler(
        MessageHandler(
            filters.Document.ALL,
            handle_document
        )
    )

    # =====================================================
    # Text messages
    # =====================================================

    application.add_handler(
        MessageHandler(
            filters.TEXT
            & ~filters.COMMAND,
            text_handler
        )
    )

    print(
        "✅ Bot is running..."
    )

    # =====================================================
    # Telegram Polling
    # =====================================================

    application.run_polling(
        drop_pending_updates=True
    )


# =========================================================
# Main
# =========================================================

if __name__ == "__main__":

    main()
