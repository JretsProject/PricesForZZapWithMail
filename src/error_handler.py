"""
src/error_handler.py
Обработка критических ошибок: сохранение деталей в JSON-файл.
"""

import json
import traceback
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional
from loguru import logger
from config.settings import CRITICAL_ERRORS_FOLDER


def log_critical_error(exception: Exception, context: Optional[Dict[str, Any]] = None) -> None:
    """
    Сохраняет информацию о критической ошибке в JSON-файл.
    
    Аргументы:
        exception: перехваченное исключение
        context: словарь с дополнительным контекстом (например, имена файлов, путь и т.д.)
    """
    # Убеждаемся, что папка существует
    CRITICAL_ERRORS_FOLDER.mkdir(parents=True, exist_ok=True)
    
    # Формируем запись
    error_record = {
        "timestamp": datetime.now().isoformat(),
        "error_type": type(exception).__name__,
        "error_message": str(exception),
        "traceback": traceback.format_exc(),
        "context": context or {}
    }
    
    # Генерируем имя файла с временной меткой
    timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    filename = f"critical_error_{timestamp_str}.json"
    file_path = CRITICAL_ERRORS_FOLDER / filename
    
    try:
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(error_record, f, ensure_ascii=False, indent=2)
        logger.error(f"Критическая ошибка сохранена в {file_path}")
    except Exception as e:
        logger.error(f"Не удалось сохранить критическую ошибку: {e}")