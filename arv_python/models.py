# -*- coding: utf-8 -*-
"""Модели данных для Модуля оценки АРВ (аналог C# comInfo, rgmsInfo, remsInfo и т.д.)."""
from dataclasses import dataclass, field
from typing import List, Optional


@dataclass
class ElId:
    ip: int
    iq: int
    np: int


@dataclass
class RgmsInfo:
    num: int
    name: str
    r_id: List[int]  # rId в C#
    s_name: Optional[str] = None  # путь к .rst, заполняется по имени из списка файлов


@dataclass
class RemsInfo:
    id: int
    name: str
    elements: List[ElId]
    mpd_value: float  # -1 если не задано
    sch_id: int  # -1 если не задано
    ut_id: int  # -1 если не задано


@dataclass
class UtInfo:
    id: int
    name: str
    s_name: Optional[str] = None  # путь к .ut2


@dataclass
class PssInfo:
    table: str
    param: str
    id: int
    val: float


@dataclass
class GrfInfo:
    table: str
    param: str
    id: List[int]  # список id (в xlsx через +)
    name: str


@dataclass
class DateValue:
    t: float
    v: float


@dataclass
class GrfData:
    name: str
    pss_on: List[DateValue]
    pss_off: Optional[List[DateValue]] = None
    s_point: float = 0.0


@dataclass
class GrfSize:
    j: int = 0   # следующая строка Excel
    grf_am: int = 0  # количество параметров в блоке


@dataclass
class ComInfo:
    """Данные из файла задания xlsx (листы rgms, rems, ut, arv, grf)."""
    name: str = ""
    rgms: List[RgmsInfo] = field(default_factory=list)
    rems: List[RemsInfo] = field(default_factory=list)
    ut: List[UtInfo] = field(default_factory=list)
    arv: List[PssInfo] = field(default_factory=list)
    grf: List[GrfInfo] = field(default_factory=list)
