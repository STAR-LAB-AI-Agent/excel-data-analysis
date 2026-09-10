"""
exceptions.py
自定义业务异常类，所有业务侧错误由此抛出
"""
from src.excel_analyzer.schemas import (
    ERR_FILE_NOT_FOUND,
    ERR_INVALID_SUFFIX,
    ERR_SHEET_NOT_EXIST,
    ERR_INVALID_INTENT,
    ERR_MISSING_ARG,
    ERR_PARSE_EXCEL_FAILED,
    ERR_COLUMN_NOT_FOUND,
    ERR_DATA_TYPE_ERR,
    ERR_INTERNAL_ERROR
)


class BaseBusinessError(Exception):
    """业务异常基类"""
    error_code: str = ERR_INTERNAL_ERROR

    def __init__(self, message: str):
        self.message = message
        super().__init__(self.message)


class FileNotFoundError(BaseBusinessError):
    error_code = ERR_FILE_NOT_FOUND


class InvalidSuffixError(BaseBusinessError):
    error_code = ERR_INVALID_SUFFIX


class SheetNotExistError(BaseBusinessError):
    error_code = ERR_SHEET_NOT_EXIST


class InvalidIntentError(BaseBusinessError):
    error_code = ERR_INVALID_INTENT


class MissingArgumentError(BaseBusinessError):
    error_code = ERR_MISSING_ARG


class ParseExcelFailedError(BaseBusinessError):
    error_code = ERR_PARSE_EXCEL_FAILED


class ColumnNotFoundError(BaseBusinessError):
    error_code = ERR_COLUMN_NOT_FOUND


class DataTypeError(BaseBusinessError):
    error_code = ERR_DATA_TYPE_ERR