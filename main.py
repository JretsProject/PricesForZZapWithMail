"""
main.py - основной модуль для обработки прайс-листов
"""

import sys
from pathlib import Path
from loguru import logger
from config.settings import LOG_FOLDER, SOURCE_FOLDER
from config.settings import PRICES_FILE_PATTERN, SPARE_PARTS_FILE_PATTERN
from src.utils import get_current_date_str, find_file_by_pattern
from src.registry import is_processed, mark_processed
from src.processor import process_excel_pair
from src.error_handler import log_critical_error  # добавлено

# Настройка логгера
LOG_FOLDER.mkdir(parents=True, exist_ok=True)
logger.remove(0)
logger.add(LOG_FOLDER / "app.log", rotation="1 day", retention="7 days", level="DEBUG")
logger.add(sys.stdout, level="INFO")

def main():
    logger.info("=== ЗАПУСК ОБРАБОТКИ ===")
    
    source = Path(SOURCE_FOLDER)
    if not source.exists():
        logger.error(f"Папка не существует: {source}")
        return
    
    today = get_current_date_str()
    
    # Поиск файлов
    prices_file = find_file_by_pattern(source, PRICES_FILE_PATTERN, today)
    spare_file = find_file_by_pattern(source, SPARE_PARTS_FILE_PATTERN, today)
    
    if not prices_file:
        logger.warning(f"Файл цен не найден: {PRICES_FILE_PATTERN.format(date=today)}")
        return
    
    if not spare_file:
        logger.warning(f"Файл запчастей не найден: {SPARE_PARTS_FILE_PATTERN.format(date=today)}")
        return
    
    # Проверка реестра
    prices_mtime = prices_file.stat().st_mtime
    spare_mtime = spare_file.stat().st_mtime
    
    if is_processed(prices_file.name, prices_mtime):
        logger.info(f"Файл {prices_file.name} уже обработан, пропускаем")
        return
    if is_processed(spare_file.name, spare_mtime):
        logger.info(f"Файл {spare_file.name} уже обработан, пропускаем")
        return
    
    # Обработка
    logger.info("Оба файла найдены и не обработаны, запускаем process_excel_pair")
    success = process_excel_pair(spare_file, prices_file)
    
    if success:
        # Помечаем файлы как обработанные
        mark_processed(prices_file.name, prices_mtime)
        mark_processed(spare_file.name, spare_mtime)
        # Архивируем исходные файлы
        from src.processor import archive_original
        from config.settings import ARCHIVE_FOLDER
        archive_original(prices_file, ARCHIVE_FOLDER)
        archive_original(spare_file, ARCHIVE_FOLDER)
        logger.info("Обработка успешно завершена")
    else:
        logger.error("Обработка завершилась с ошибкой")
    
    logger.info("=== ЗАВЕРШЕНО ===")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        logger.critical(f"Необработанная критическая ошибка: {e}")
        log_critical_error(e, context={"script": "main.py", "source_folder": str(SOURCE_FOLDER)})
        sys.exit(1)