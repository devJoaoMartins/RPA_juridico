from __future__ import annotations

import logging
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Tuple

from openpyxl import load_workbook


class ExcelReader:

    def __init__(self, file_path: Path):
        self.file_path = Path(file_path)
        self.wb = None
        self.logger = logging.getLogger(__name__)

    def __enter__(self) -> "ExcelReader":
        try:
            self.wb = load_workbook(self.file_path, data_only=True, read_only=True)
            self.logger.info("Planilha Excel carregada com sucesso")
            return self
        except Exception as e:
            self.logger.error(f"Erro ao abrir Excel: {str(e)}")
            raise

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        if self.wb:
            self.wb.close()

    @staticmethod
    def _format_brl(value: float) -> str:
        v = float(value)
        s = f"{abs(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        sign = "-" if v < 0 else ""
        return f"{sign}R$ {s}"

    @staticmethod
    def _format_percent_br(value: float) -> str:
        pct = float(value) * 100.0
        return f"{pct:.2f}".replace(".", ",") + "%"

    @staticmethod
    def _format_date_br(value: date | datetime) -> str:
        d = value.date() if isinstance(value, datetime) else value
        return d.strftime("%d/%m/%Y")

    def _format_by_number_format(self, value: Any, number_format: str):
        if not isinstance(value, (int, float)):
            return value
        fmt = (number_format or "").lower()
        if "%" in fmt:
            return self._format_percent_br(value)
        if "r$" in fmt or "[$" in fmt:
            return self._format_brl(value)
        return value

    def get_cell_value(self, sheet_name: str, cell_address: str) -> str:
        try:
            sheet = self.wb[sheet_name]  # type: ignore[index]
            cell = sheet[cell_address]
            cell_value = cell.value

            if cell_value is None:
                return ""
            if isinstance(cell_value, (datetime, date)):
                try:
                    return self._format_date_br(cell_value)
                except Exception:
                    return str(cell_value)

            try:
                number_format = getattr(cell, "number_format", "") or ""
                formatted = self._format_by_number_format(cell_value, number_format)
            except Exception:
                formatted = cell_value
            return "" if formatted is None else str(formatted)
        except KeyError:
            self.logger.error(f"Aba '{sheet_name}' não encontrada")
            return ""
        except Exception as e:
            self.logger.error(f"Erro na célula {sheet_name}!{cell_address}: {str(e)}")
            return ""

    def read_mapping(self, mapping: Dict[str, Tuple[str, str]]) -> Dict[str, str]:
        result: Dict[str, str] = {}
        for marker, (sheet, cell) in mapping.items():
            value = self.get_cell_value(sheet, cell)
            result[marker] = "" if value is None else str(value)
        return result

