"""
logger_cfg.py
全局日志配置
日志文件路径: 项目根目录/logs/excel_analyzer.log
自动探测项目根目录，避免相对路径依赖工作目录
"""
import os
import logging
from pathlib import Path
from logging.handlers import RotatingFileHandler


def _get_project_root() -> Path:
    """自动探测项目根目录，向上搜索包含 tests 文件夹"""
    current = Path(__file__).resolve()
    for _ in range(5):
        if (current / "tests").exists():
            return current
        current = current.parent
    return Path.cwd()


PROJECT_ROOT = _get_project_root()
LOG_DIR = str(PROJECT_ROOT / "logs")
LOG_FILE = os.path.join(LOG_DIR, "excel_analyzer.log")
LOG_LEVEL = logging.INFO
MAX_LOG_SIZE = 5 * 1024 * 1024
BACKUP_COUNT = 3


def setup_logger() -> logging.Logger:
    if not os.path.exists(LOG_DIR):
        os.makedirs(LOG_DIR, exist_ok=True)

    logger = logging.getLogger("excel_analyzer")
    logger.setLevel(LOG_LEVEL)
    if logger.handlers:
        return logger

    log_formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )

    file_handler = RotatingFileHandler(
        LOG_FILE,
        maxBytes=MAX_LOG_SIZE,
        backupCount=BACKUP_COUNT,
        encoding="utf-8"
    )
    file_handler.setFormatter(log_formatter)
    file_handler.setLevel(LOG_LEVEL)

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(log_formatter)
    console_handler.setLevel(LOG_LEVEL)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    return logger


logger = setup_logger()