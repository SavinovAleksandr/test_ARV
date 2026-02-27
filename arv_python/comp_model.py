# -*- coding: utf-8 -*-
"""Модель данных приложения: набор файлов и разобранное задание xlsx (аналог C# compModel)."""
import os
from pathlib import Path
from typing import List, Optional

from models import ComInfo
from xlsx_parser import load_assignment_xlsx


class CompModel:
    """
    Хранит пути к файлам и по запросу отдаёт списки режимов, траекторий, сечений, сценарий, задание xlsx.
    Соответствует C# compModel (regims, ut2s, schs, scns, xlxs).
    """

    def __init__(self):
        self._files: List[str] = []
        self._rgms: List[str] = []
        self._ut2_files: List[str] = []
        self._sch_file: str = ""
        self._scns_file: str = ""
        self._xls_file: Optional[ComInfo] = None
        self._xlsx_path: Optional[str] = None

    @property
    def files(self) -> List[str]:
        return self._files

    @files.setter
    def files(self, value: List[str]):
        self._files = list(value) if value else []
        self._rgms.clear()
        self._ut2_files.clear()
        self._sch_file = ""
        self._scns_file = ""
        self._xls_file = None
        self._xlsx_path = None
        for p in self._files:
            ext = Path(p).suffix.lower()
            if ext == ".rst" and p not in self._rgms:
                self._rgms.append(p)
            elif ext == ".ut2" and p not in self._ut2_files:
                self._ut2_files.append(p)
            elif ext == ".sch":
                self._sch_file = p
            elif ext == ".scn":
                self._scns_file = p

    @property
    def regims(self) -> List[str]:
        """Список путей к .rst. После установки files обновляется; при обращении к xlxs сопоставляются sName."""
        return self._rgms

    def clear_regims(self):
        self._rgms.clear()

    @property
    def ut2s(self) -> List[str]:
        return self._ut2_files

    def clear_ut2s(self):
        self._ut2_files.clear()

    @property
    def schs(self) -> str:
        return self._sch_file

    def clear_schs(self):
        self._sch_file = ""

    @property
    def scns(self) -> str:
        return self._scns_file

    @property
    def xlxs(self) -> Optional[ComInfo]:
        """Файл задания xlsx: разбор листов rgms, rems, ut, arv, grf. Кэш по текущему набору files."""
        for p in self._files:
            if Path(p).suffix.lower() in (".xlsx", ".xls"):
                if p == self._xlsx_path and self._xls_file is not None:
                    return self._xls_file
                info = load_assignment_xlsx(p, self._rgms, self._ut2_files)
                if info is not None:
                    self._xlsx_path = p
                    self._xls_file = info
                    return self._xls_file
        return self._xls_file

    @property
    def xlsx_name(self) -> str:
        """Имя файла задания (без расширения) для отображения."""
        if self._xls_file and self._xls_file.name:
            return self._xls_file.name
        for p in self._files:
            if Path(p).suffix.lower() in (".xlsx", ".xls"):
                return Path(p).stem
        return ""
