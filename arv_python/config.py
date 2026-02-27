# -*- coding: utf-8 -*-
"""Пути к шаблонам Rastr и прочие константы (аналог зашитых путей в C#)."""
import os

# Каталог шаблонов Rastr (как в C#: Personal\RastrWin3\SHABLON)
def get_rastr_shablon_dir() -> str:
    return os.path.join(os.path.expanduser("~"), "Documents", "RastrWin3", "SHABLON")

def get_dynamics_rst() -> str:
    return os.path.join(get_rastr_shablon_dir(), "динамика.rst")

def get_scenario_scn() -> str:
    return os.path.join(get_rastr_shablon_dir(), "сценарий.scn")

def get_avtomatika_dfw() -> str:
    return os.path.join(get_rastr_shablon_dir(), "автоматика.dfw")

def get_sechen_sch() -> str:
    return os.path.join(get_rastr_shablon_dir(), "сечения.sch")

def get_trajectory_ut2() -> str:
    return os.path.join(get_rastr_shablon_dir(), "траектория утяжеления.ut2")

# Имена выходных файлов (в каталоге программы)
RESULT_XLSX = "Приложение. Результат в табличном виде.xlsx"
RESULT_DOCX = "Приложение. Результат в графическом виде.docx"

# CLSID Rastr (ПК RUSTab)
RASTR_CLSID = "{EFC5E4AD-A3DD-11D3-B73F-00500454CF3F}"
# CLSID Word
WORD_CLSID = "{000209FF-0000-0000-C000-000000000046}"
