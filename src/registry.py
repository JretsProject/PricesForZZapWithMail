"""
src/registry.py
Управление реестром обработанных файлов (JSON).
Реестр хранит для каждого файла его имя и время последней модификации (mtime),
чтобы избежать повторной обработки.
"""

import json
from pathlib import Path
from typing import Dict
from loguru import logger
from config.settings import REGISTRY_PATH


def load_registry() -> Dict[str, float]:
    """
    Загружает реестр из JSON-файла.
    Если файл не существует или повреждён, возвращает пустой словарь.
    """
    if not REGISTRY_PATH.exists():
        logger.debug("Файл реестра не найден, будет создан новый.")
        return {}
    try:
        with open(REGISTRY_PATH, 'r', encoding='utf-8') as f:
            registry = json.load(f)
            if not isinstance(registry, dict):
                logger.warning("Реестр имеет неверный формат, сбрасываем.")
                return {}
            # Приводим все ключи к строкам, значения к float
            clean_registry = {str(k): float(v) for k, v in registry.items() if isinstance(v, (int, float))}
            return clean_registry
    except Exception as e:
        logger.error(f"Ошибка загрузки реестра: {e}. Начинаем с пустого реестра.")
        return {}


def save_registry(registry: Dict[str, float]) -> None:
    """Сохраняет реестр в JSON-файл."""
    try:
        REGISTRY_PATH.parent.mkdir(parents=True, exist_ok=True)
        with open(REGISTRY_PATH, 'w', encoding='utf-8') as f:
            json.dump(registry, f, ensure_ascii=False, indent=2)
        logger.debug(f"Реестр сохранён: {len(registry)} записей.")
    except Exception as e:
        logger.error(f"Ошибка сохранения реестра: {e}")


def is_processed(filename: str, mtime: float) -> bool:
    """
    Проверяет, обрабатывался ли уже файл.
    Сравнивает сохранённое время модификации с текущим.
    """
    registry = load_registry()
    stored_mtime = registry.get(filename)
    if stored_mtime is None:
        return False
    return abs(stored_mtime - mtime) < 0.001


def mark_processed(filename: str, mtime: float) -> None:
    """Отмечает файл как обработанный."""
    registry = load_registry()
    registry[filename] = mtime
    save_registry(registry)
    logger.info(f"Файл отмечен как обработанный: {filename} (mtime={mtime})")