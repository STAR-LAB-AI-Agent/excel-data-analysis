"""
logger_cfg.py
全局日志配置
日志文件路径: ./logs/excel_analyzer.log
自动创建logs目录；禁止记录完整Excel原始数据，规避低Token风险与敏感数据泄露
"""
import os
import logging
from logging.handlers import RotatingFileHandler

# 日志配置常量
LOG_DIR = "./logs"
LOG_FILE = os.path.join(LOG_DIR, "excel_analyzer.log")
LOG_LEVEL = logging.INFO
# 日志轮转：单文件最大5MB，最多保留3个备份
MAX_LOG_SIZE = 5 * 1024 * 1024
BACKUP_COUNT = 3


def setup_logger() -> logging.Logger:
    # 创建logs文件夹
    if not os.path.exists(LOG_DIR):
        os.makedirs(LOG_DIR, exist_ok=True)

    logger = logging.getLogger("excel_analyzer")
    logger.setLevel(LOG_LEVEL)
    # 防止重复添加handler
    if logger.handlers:
        return logger

    # 日志格式
    log_formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )

    # 文件处理器(轮转日志)
    file_handler = RotatingFileHandler(
        LOG_FILE,
        maxBytes=MAX_LOG_SIZE,
        backupCount=BACKUP_COUNT,
        encoding="utf-8"
    )
    file_handler.setFormatter(log_formatter)
    file_handler.setLevel(LOG_LEVEL)

    # 控制台输出处理器
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(log_formatter)
    console_handler.setLevel(LOG_LEVEL)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    return logger


# 全局logger实例，其他模块直接 import logger
logger = setup_logger()