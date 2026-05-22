"""
src/utils.py
Вспомогательные функции для работы с датами и поиска файлов по шаблону.
"""

from pathlib import Path
from datetime import date
from typing import Optional
from loguru import logger


def get_current_date_str() -> str:
    """
    Возвращает текущую дату в формате YYYY-MM-DD.
    Пример: '2026-05-22'
    """
    return date.today().isoformat()


def find_file_by_pattern(folder: Path, pattern: str, expected_date: Optional[str] = None) -> Optional[Path]:
    """
    Ищет файл в указанной папке по шаблону с подстановкой даты.

    Аргументы:
        folder: путь к папке (Path)
        pattern: шаблон имени файла, содержащий '{date}'.
                 Например: "PricesForZZap_{date}.xlsx"
        expected_date: дата в формате YYYY-MM-DD. Если не указана, используется текущая.

    Возвращает:
        Path к файлу, если он существует, иначе None.
    """
    if expected_date is None:
        expected_date = get_current_date_str()
    
    filename = pattern.format(date=expected_date)
    file_path = folder / filename
    
    if file_path.exists() and file_path.is_file():
        logger.debug(f"Найден файл: {file_path}")
        return file_path
    else:
        logger.debug(f"Файл не найден: {file_path}")
        return None