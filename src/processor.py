"""
src/processor.py
Основная логика обработки Excel-файлов, очистки данных и формирования выходных файлов.
"""

import re
import shutil
from pathlib import Path
from datetime import datetime
from typing import Optional, Tuple
import pandas as pd
from loguru import logger
import shutil
from config.settings import (
    PURCHASE_COLUMNS, SALES_COLUMNS,
    STOP_WORDS, PATTERNS,
    OUTPUT_ZZAP_FILE, MARKUP_ANALYSIS_FILE,
    MISSING_ASSORTMENT_FILE, MISSING_INVENTORY_FILE,
    MISSING_PRICE_FILE, PRICE_BELOW_PURCHASE_FILE ,
    ARCHIVE_FOLDER
)


from config.settings import (
    PURCHASE_COLUMNS, SALES_COLUMNS,
    STOP_WORDS, PATTERNS,
    OUTPUT_ZZAP_FILE, MARKUP_ANALYSIS_FILE,
    MISSING_ASSORTMENT_FILE, MISSING_INVENTORY_FILE,
    MISSING_PRICE_FILE, PRICE_BELOW_PURCHASE_FILE ,
    ARCHIVE_FOLDER,
    send_email  # <-- добавить
)


def archive_original(file_path: Path, archive_folder: Path) -> None:
    """
    Копирует исходный файл в архивную папку.
    Имя файла: исходное_имя_archive.xlsx (дата не добавляется отдельно, т.к. она уже есть в имени файла).
    Если папка архива не существует, создаёт её.
    """
    try:
        archive_folder.mkdir(parents=True, exist_ok=True)
        # Формируем имя: PricesForZZap_2026-05-22_archive.xlsx
        stem = file_path.stem  # без расширения
        archive_name = f"{stem}_archive{file_path.suffix}"
        archive_path = archive_folder / archive_name
        shutil.copy2(file_path, archive_path)
        logger.info(f"Файл архивирован: {file_path} -> {archive_path}")
    except Exception as e:
        logger.error(f"Ошибка архивации {file_path.name}: {e}")
        # Не прерываем выполнение программы, только логируем

def get_converted_string(text: str) -> str:
    """
    Очищает название производителя от географических пометок и стоп-слов.
    Возвращает очищенную строку или "НЕ ОПРЕДЕЛЕН", если результат пуст.
    """
    if not isinstance(text, str) or not text.strip():
        return "НЕ ОПРЕДЕЛЕН"

    original = text
    # Приводим к нижнему регистру для сравнения со стоп-словами
    text_lower = text.lower()

    # Удаляем стоп-слова как целые слова (регулярка)
    if STOP_WORDS:
        # Экранируем спецсимволы и сортируем по длине (сначала длинные)
        escaped = [re.escape(word) for word in STOP_WORDS if word]
        escaped.sort(key=len, reverse=True)
        pattern = r'\b(?:' + '|'.join(escaped) + r')\b'
        text = re.sub(pattern, '', text, flags=re.IGNORECASE)

    # Применяем паттерны из настроек
    for pattern in PATTERNS:
        text = pattern.sub('', text)

    # Удаляем скобки всех видов
    text = re.sub(r'[\[\]\(\)\{\}]', '', text)
    # Убираем лишние пробелы и знаки препинания по краям
    text = re.sub(r'\s+', ' ', text).strip()
    text = text.strip(' ,.;:-')

    if not text:
        logger.warning(f"Производитель после очистки стал пустым: исходный '{original}'")
        return "НЕ ОПРЕДЕЛЕН"
    return text


def load_and_process_purchase(file_path: Path) -> Optional[pd.DataFrame]:
    """
    Загружает и обрабатывает файл SparePartsForZZap.
    Учитывает, что основные заголовки в строке 1, а заголовки 'Цена' и 'Количество' в строке 3.
    Пропускает последнюю строку (итоги).
    """
    logger.info(f"Загрузка файла закупок: {file_path}")
    try:
        # Читаем весь файл без заголовков, чтобы получить все строки
        df_raw = pd.read_excel(file_path, header=None, engine='openpyxl')
        if len(df_raw) < 4:
            logger.error("Файл слишком короткий, ожидается минимум 4 строки")
            return None

        # Строка 0 (индекс 0) – основные заголовки
        headers_row0 = df_raw.iloc[0].fillna('').astype(str)
        # Строка 2 (индекс 2) – заголовки цен и количества
        headers_row2 = df_raw.iloc[2].fillna('').astype(str)

        # Объединяем заголовки:
        # - Берём из row0, если не пусто и не "Итого"
        # - Иначе берём из row2 (для колонок Цена, Количество)
        # - Пропускаем колонки, где оба пусты
        combined_headers = []
        for i in range(len(headers_row0)):
            h0 = headers_row0.iloc[i].strip()
            h2 = headers_row2.iloc[i].strip() if i < len(headers_row2) else ""
            if h0 and h0 != "Итого":
                combined_headers.append(h0)
            elif h2:
                combined_headers.append(h2)
            else:
                # Если оба пусты – колонка не нужна
                continue

        # Данные начинаются со строки 3 (индекс 3) и идут до предпоследней (пропускаем итоги)
        # skipfooter=1 не работает с header=None, поэтому обрежем вручную
        data = df_raw.iloc[3:-1].copy()   # от строки 3 до последней-1
        # Обрезаем колонки по количеству сформированных заголовков
        data = data.iloc[:, :len(combined_headers)]
        data.columns = combined_headers

        # Приводим нужные колонки к строковому типу и очищаем
        for col in ['КаталожныйНомер', 'Номенклатура', 'Код', 'АссортиментZZAP',
                    'КодТовараZZAP', 'Производитель', 'ТорговаяМаркаZZAP',
                    'Инвентаризация', 'Состояние']:
            if col in data.columns:
                data[col] = data[col].fillna('').astype(str).str.strip()
            else:
                logger.warning(f"Колонка {col} не найдена в файле")

        # Переименовываем колонку 'Количество (в ед. хранения)' в 'Количество'
        if 'Количество (в ед. хранения)' in data.columns:
            data = data.rename(columns={'Количество (в ед. хранения)': 'Количество'})

        # Преобразуем числовые колонки
        if 'Цена' in data.columns:
            data['Цена'] = pd.to_numeric(data['Цена'], errors='coerce')
        if 'Количество' in data.columns:
            data['Количество'] = pd.to_numeric(data['Количество'], errors='coerce')

        # Удаляем строки, где все основные поля пусты (опционально)
        data = data.dropna(how='all')
        logger.info(f"Загружено {len(data)} строк из файла закупок")
    except Exception as e:
        error_msg = f"Ошибка загрузки файла закупок {file_path.name}: {e}"
        logger.error(error_msg)
        send_email("Ошибка загрузки файла закупок", error_msg)
        return None

    # ===== Дальнейшая обработка по ТЗ =====
    # Переименовываем колонку Цена в ЦенаЗакупки
    data = data.rename(columns={"Цена": "ЦенаЗакупки"})

    # Добавляем колонку Артикул: сначала из КаталожныйНомер, потом из КодТовараZZAP (если не пусто)
    data["Артикул"] = data["КаталожныйНомер"]
    mask = data["КодТовараZZAP"].notna() & (data["КодТовараZZAP"].astype(str).str.strip() != "")
    data.loc[mask, "Артикул"] = data.loc[mask, "КодТовараZZAP"]

    # Удаляем колонки КаталожныйНомер и КодТовараZZAP
    data = data.drop(columns=["КаталожныйНомер", "КодТовараZZAP"])

    # Создаём Наименование = Номенклатура
    data["Наименование"] = data["Номенклатура"].astype(str)

    # Очищаем Производитель от скобок [ ]
    data["Производитель"] = data["Производитель"].astype(str).str.replace(r'[\[\]]', '', regex=True)

    # Добавляем в Наименование: , (Производитель), затем , [Код]
    data["Наименование"] = data["Наименование"] + ", (" + data["Производитель"] + ")"
    data["Наименование"] = data["Наименование"] + ", [" + data["Код"].astype(str) + "]"

    # Удаляем колонку Номенклатура
    data = data.drop(columns=["Номенклатура"])

    # Заменяем Производитель на ТорговаяМаркаZZAP, если последняя не пуста
    mask_tm = data["ТорговаяМаркаZZAP"].notna() & (data["ТорговаяМаркаZZAP"].astype(str).str.strip() != "")
    data.loc[mask_tm, "Производитель"] = data.loc[mask_tm, "ТорговаяМаркаZZAP"].astype(str)
    data = data.drop(columns=["ТорговаяМаркаZZAP"])

    # Очищаем Производитель через get_converted_string
    data["Производитель"] = data["Производитель"].apply(get_converted_string)

    logger.info(f"Обработка файла закупок завершена, осталось {len(data)} строк")
    return data


def load_and_process_sales(file_path: Path) -> Optional[pd.DataFrame]:
    logger.info(f"Загрузка файла продаж: {file_path}")
    try:
        df = pd.read_excel(file_path, usecols=SALES_COLUMNS, engine='openpyxl')
        logger.info(f"Загружено {len(df)} строк из файла продаж")
    except Exception as e:
        error_msg = f"Ошибка загрузки файла продаж {file_path.name}: {e}"
        logger.error(error_msg)
        send_email("Ошибка загрузки файла продаж", error_msg)
        return None
    # Удаляем строки с пустой ценой
    before = len(df)
    df = df.dropna(subset=['Цена'])
    after = len(df)
    if before != after:
        logger.info(f"Удалено {before - after} строк с пустой ценой")
    df['Код'] = df['Код'].astype(str).str.strip()
    return df


def process_excel_pair(purchase_file: Path, sales_file: Path) -> bool:
    """
    Основная функция обработки пары файлов.
    """
    purchase_df = load_and_process_purchase(purchase_file)
    if purchase_df is None:
        return False
    sales_df = load_and_process_sales(sales_file)
    if sales_df is None:
        return False

    merged = purchase_df.merge(sales_df, on='Код', how='left')

    # 1. Пустой АссортиментZZAP
    missing_assort = merged[merged['АссортиментZZAP'].isna() | (merged['АссортиментZZAP'].astype(str).str.strip() == '')]
    if len(missing_assort) > 0:
        out_path = purchase_file.parent / MISSING_ASSORTMENT_FILE
        missing_assort.to_excel(out_path, index=False)
        logger.info(f"Строк с пустым АссортиментZZAP: {len(missing_assort)} сохранено в {out_path}")
        send_email("Обнаружены пустые АссортиментZZAP", f"Количество: {len(missing_assort)}", attachment_path=out_path)
    else:
        logger.info("Нет строк с пустым АссортиментZZAP")

    # 2. Пустая Инвентаризация
    missing_inv = merged[merged['Инвентаризация'].isna() | (merged['Инвентаризация'].astype(str).str.strip() == '')]
    if len(missing_inv) > 0:
        out_path = purchase_file.parent / MISSING_INVENTORY_FILE
        missing_inv.to_excel(out_path, index=False)
        logger.info(f"Строк с пустой Инвентаризацией: {len(missing_inv)} сохранено в {out_path}")
        send_email("Обнаружены пустые Инвентаризация", f"Количество: {len(missing_inv)}", attachment_path=out_path)
    else:
        logger.info("Нет строк с пустой Инвентаризацией")

    # 3. Пустая Цена
    missing_price = merged[merged['Цена'].isna()]
    if len(missing_price) > 0:
        out_path = purchase_file.parent / MISSING_PRICE_FILE
        missing_price.to_excel(out_path, index=False)
        logger.info(f"Строк с пустой Ценой: {len(missing_price)} сохранено в {out_path}")
        send_email("Обнаружены пустые Цены", f"Количество: {len(missing_price)}", attachment_path=out_path)
    else:
        logger.info("Нет строк с пустой Ценой")

    # 4. Цена < ЦенаЗакупки
    valid_price = merged.dropna(subset=['Цена', 'ЦенаЗакупки'])
    price_below = valid_price[valid_price['Цена'] < valid_price['ЦенаЗакупки']]
    if len(price_below) > 0:
        out_path = purchase_file.parent / PRICE_BELOW_PURCHASE_FILE   # используем новую константу
        price_below.to_excel(out_path, index=False)
        logger.info(f"Строк с Цена < ЦенаЗакупки: {len(price_below)} сохранено в {out_path}")
        send_email("Цена продажи ниже закупочной", f"Количество: {len(price_below)}", attachment_path=out_path)
    else:
        logger.info("Нет строк с Цена < ЦенаЗакупки")

    # Удаление строк с пустыми обязательными полями
    filtered = merged.dropna(subset=['АссортиментZZAP', 'Инвентаризация', 'Цена'])
    filtered = filtered[
        (filtered['АссортиментZZAP'].astype(str).str.strip() != '') &
        (filtered['Инвентаризация'].astype(str).str.strip() != '')
    ]
    logger.info(f"После удаления строк с пустыми обязательными полями осталось {len(filtered)} строк (было {len(merged)})")

    filtered = filtered.drop(columns=['АссортиментZZAP', 'Инвентаризация', 'Состояние'], errors='ignore')

    # Файл анализа
    analysis_path = purchase_file.parent / MARKUP_ANALYSIS_FILE
    filtered.to_excel(analysis_path, index=False)
    logger.info(f"Создан файл анализа наценки: {analysis_path} ({len(filtered)} строк)")

    # Файл для площадки
    output_df = filtered[['Производитель', 'Артикул', 'Наименование', 'Количество', 'Цена']].copy()
    output_path = purchase_file.parent / OUTPUT_ZZAP_FILE
    output_df.to_excel(output_path, index=False)
    logger.info(f"Создан файл для площадки: {output_path} ({len(output_df)} строк)")

    send_email("Обработка завершена успешно", f"Создано {len(output_df)} товаров для выгрузки")
    return True

