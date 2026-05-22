"""
config/settings.py
Настройки проекта: пути, списки стоп-слов, регулярные выражения, имена файлов,
колонки для загрузки из Excel, инициализация логгера для настроек.
"""

import os
import re
import json
from pathlib import Path
from dotenv import load_dotenv
from loguru import logger
from typing import Optional
import smtplib
from email.message import EmailMessage
import ssl

# ===== ВРЕМЕННАЯ НАСТРОЙКА ЛОГГЕРА ДЛЯ ВЫВОДА ИЗ SETTINGS =====
# (основной логгер будет настроен в main.py, а здесь просто дублируем в консоль)
#logger.add(lambda msg: print(msg, end=""), level="INFO")

# Загружаем переменные окружения из файла .env
load_dotenv()

# ===== БАЗОВЫЕ ПУТИ =====
BASE_DIR = Path(__file__).parent.parent               # Корень проекта (PricesForZZap)
SOURCE_FOLDER = os.getenv("SOURCE_FOLDER", r"\\192.168.1.51\public\1s\sourse_prices")  # Сетевая папка с исходными/выходными файлами
ARCHIVE_FOLDER = BASE_DIR / os.getenv("ARCHIVE_FOLDER", "data/archive")                # Локальная папка для архива исходников
REGISTRY_PATH = BASE_DIR / os.getenv("REGISTRY_PATH", "data/registry.json")            # Путь к файлу реестра обработанных файлов
LOG_FOLDER = BASE_DIR / os.getenv("LOG_FOLDER", "logs")                                # Папка для логов
CRITICAL_ERRORS_FOLDER = BASE_DIR / os.getenv("CRITICAL_ERRORS_FOLDER", "logs/critical_errors")  # Папка для JSON-файлов критических ошибок

# ===== ШАБЛОНЫ ИМЁН ВХОДНЫХ ФАЙЛОВ (с датой) =====
PRICES_FILE_PATTERN = "PricesForZZap_{date}.xlsx"            # Шаблон для файла цен (продажи)
SPARE_PARTS_FILE_PATTERN = "SparePartsForZZap_{date}.xlsx"   # Шаблон для файла запчастей (закупки)

# ===== ИМЕНА ВЫХОДНЫХ ФАЙЛОВ (сохраняются в SOURCE_FOLDER) =====
MARKUP_ANALYSIS_FILE = "markup_analysis_for_zzap.xlsx"       # Для анализа наценки (все колонки после обработки)
OUTPUT_ZZAP_FILE = "kamaz_kmv_for_zzap.xlsx"                 # Финальный файл для загрузки на площадку (5 колонок)
MISSING_ASSORTMENT_FILE = "missing_АссортиментZZAP.xlsx"     # Отчёт: строки с пустым полем АссортиментZZAP
MISSING_INVENTORY_FILE = "missing_Инвентаризация.xlsx"       # Отчёт: строки с пустым полем Инвентаризация
MISSING_PRICE_FILE = "missing_Цена.xlsx"                     # Отчёт: строки с пустым полем Цена (из файла продаж)
PRICE_BELOW_PURCHASE_FILE = "price_below_purchase.xlsx"      # Отчёт: строки, где Цена продажи < ЦенаЗакупки

# ===== КОЛОНКИ ДЛЯ ЗАГРУЗКИ ИЗ EXCEL (по ТЗ) =====
PURCHASE_COLUMNS = [
    "КаталожныйНомер", "Номенклатура", "Код", "АссортиментZZAP",
    "КодТовараZZAP", "Производитель", "ТорговаяМаркаZZAP",
    "Инвентаризация", "Состояние", "Количество", "Цена"
]   # Список колонок, которые читаем из файла SparePartsForZZap

SALES_COLUMNS = ["Код", "Цена"]   # Колонки, читаемые из файла PricesForZZap

# ===== СТОП-СЛОВА И ПАТТЕРНЫ ДЛЯ ОЧИСТКИ ПРОИЗВОДИТЕЛЯ (функция get_converted_string) =====
# Географические указатели (удаляемые конструкции)
GEO_INDICATORS = [r"г\.", "город", " г ", "область", "республика", "край", "район"]

# Страны (удаляются целиком как слова)
COUNTRIES = ["Болгария", "Россия", "Украина", "Беларусь", "Казахстан", "Италия"]

# Регионы и их сокращения
REGIONS = [
    "московская обл", "ленинградская обл", "санкт-петербург",
    "Московская", "Московск.", "Ленинградская", "Калужская"
]

# Дополнительные города (ручной список)
EXTRA_CITIES = [
    "Королев", "Н-Челны", "Н-челны", "Н-Челны", "Гродно",
    "Наб Челны", "Московск", "Борисов", "Камида"
]

# Загрузка списка городов из data/russian-cities.json (если файл существует)
cities_json_path = BASE_DIR / "data" / "russian-cities.json"
CITIES = []
if cities_json_path.exists():
    try:
        with open(cities_json_path, 'r', encoding='utf-8') as f:
            cities_data = json.load(f)
            # Предполагается, что JSON содержит массив объектов с полем "name"
            CITIES = [city["name"] for city in cities_data if "name" in city]
        logger.info(f"Загружено {len(CITIES)} городов из JSON")
    except Exception as e:
        logger.error(f"Ошибка загрузки городов из JSON: {e}")
else:
    logger.warning(f"Файл с городами не найден: {cities_json_path}")

# Добавляем ручные города и убираем дубликаты
CITIES.extend(EXTRA_CITIES)
CITIES = list(set(CITIES))

# Объединённый список стоп-слов (все категории) – будет использован в get_converted_string
STOP_WORDS = list(set(GEO_INDICATORS + COUNTRIES + REGIONS + CITIES))

# Регулярные выражения (скомпилированные) для удаления типовых оборотов в названии производителя
PATTERNS = [
    re.compile(r"\bг\.\s*\w+"),                     # "г. Москва"
    re.compile(r"\bобл\b"),                         # "обл"
    re.compile(r"\bобласть\b"),                     # "область"
    re.compile(r"\bреспублика\b"),                  # "республика"
    re.compile(r"\bкрай\b"),                        # "край"
    re.compile(r"\bрайон\b"),                       # "район"
    re.compile(r"\bМосковск\.?\s*обл\b"),           # "Московск. обл"
    re.compile(r"\b[А-Я][а-я]+\.?\s*обл\b"),        # "Калужская обл"
    re.compile(r"г\.\s*[А-Я][а-я\-]+"),             # "г. Город"
]

# ===== ПАТТЕРНЫ ДЛЯ ПОИСКА ФАЙЛОВ В ПАПКЕ (используются модулем monitor.py) =====
PURCHASE_PATTERN_ZZAP = "*SparePartsForZZap_*.xlsx"   # Glob-паттерн для файлов закупки
SALES_PATTERN_ZZAP = "*PricesForZZap_*.xlsx"          # Glob-паттерн для файлов продаж

# ===== ЗАГЛУШКА ДЛЯ ОТПРАВКИ EMAIL (будет вызываться при критических ошибках или отсутствии файлов) =====

# Загружаем из .env
SMTP_HOST = os.getenv("SMTP_HOST", "smtp.yandex.ru")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER = os.getenv("SMTP_USER", "")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD", "")
SMTP_RECIPIENT = os.getenv("SMTP_RECIPIENT", "")


def send_email(subject: str, body: str, attachment_path: Optional[Path] = None) -> None:
    """Отправляет email с правильным типом подключения (TLS или SSL)."""
    if not all([SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASSWORD, SMTP_RECIPIENT]):
        logger.warning("SMTP настройки неполные, письмо не отправлено")
        return

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = SMTP_USER
    msg["To"] = SMTP_RECIPIENT
    msg.set_content(body)

    if attachment_path and attachment_path.exists():
        with open(attachment_path, "rb") as f:
            file_data = f.read()
            file_name = attachment_path.name
        msg.add_attachment(file_data, maintype="application", subtype="octet-stream", filename=file_name)

    try:
        # Выбираем тип подключения в зависимости от порта
        if SMTP_PORT == 465:
            # Для 465 порта используем SMTP_SSL с самого начала
            logger.debug("Используем SMTP_SSL для порта 465")
            with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=30, context=ssl.create_default_context()) as server:
                server.login(SMTP_USER, SMTP_PASSWORD)
                server.send_message(msg)
        else:
            # Для других портов (например, 587) используем стандартный SMTP с STARTTLS
            logger.debug(f"Используем STARTTLS для порта {SMTP_PORT}")
            with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30) as server:
                server.starttls(context=ssl.create_default_context())
                server.login(SMTP_USER, SMTP_PASSWORD)
                server.send_message(msg)
        logger.info(f"Email отправлен: {subject}")
    except Exception as e:
        logger.error(f"Ошибка отправки email: {e}")
        

"""
def send_email(subject: str, body: str, attachment_path: Optional[Path] = None) -> None:
    #Заглушка для отправки email (без реальной рассылки).
    logger.info(f"[EMAIL STUB] Subject: {subject}")
    logger.info(f"[EMAIL STUB] Body: {body}")
    if attachment_path:
        logger.info(f"[EMAIL STUB] Attachment: {attachment_path}")
"""