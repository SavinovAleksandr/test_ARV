# -*- coding: utf-8 -*-
"""
Обёртка над COM Rastr (ПК RUSTab). Используется только под Windows при установленном RUSTab и pywin32.
Без них расчёт динамики и режимов недоступен — приложение покажет ошибку при нажатии «Расчёт».
"""
import sys
from typing import Any, List, Optional

from config import (
    RASTR_CLSID,
    get_avtomatika_dfw,
    get_dynamics_rst,
    get_scenario_scn,
    get_sechen_sch,
    get_trajectory_ut2,
)

# RG_REPL = 1 (константа из ASTRALib.RG_KOD)
RG_REPL = 1

def is_rastr_available() -> bool:
    """Rastr доступен только под Windows при установленном RUSTab."""
    return sys.platform == "win32"


def create_rastr():
    """Создать экземпляр Rastr через COM. Только Windows."""
    if sys.platform != "win32":
        raise RuntimeError("Rastr (ПК RUSTab) доступен только под Windows. Установите RUSTab и запустите на Windows.")
    try:
        import win32com.client
        return win32com.client.Dispatch(RASTR_CLSID)
    except Exception as e:
        raise RuntimeError(
            "Не удалось создать Rastr (ПК RUSTab). Убедитесь, что RUSTab установлен и зарегистрирован. Ошибка: " + str(e)
        ) from e


def load_file(rastr, path: str, template: str, mode: int = RG_REPL):
    """Загрузить файл в шаблон. rastr.Load(mode, path, template)."""
    rastr.Load(mode, path, template)


def save_file(rastr, path: str, template: str):
    rastr.Save(path, template)


def get_table(rastr, name: str):
    return rastr.Tables.Item(name)


def get_col(table, name: str):
    return table.Cols.Item(name)


def col_get_z(col, row: int) -> Any:
    return col.get_Z(row)


def col_set_z(col, row: int, value: Any):
    col.set_Z(row, value)


def col_calc(col, value: str):
    col.Calc(value)


def table_set_sel(table, sel: str):
    table.SetSel(sel)


def table_count(table) -> int:
    return table.Count


def table_find_next_sel(table, from_row: int) -> int:
    return table.FindNextSel(from_row)


def rgm(rastr, arg: str = "") -> int:
    """Расчёт установившегося режима. Возврат: код (0 = AST_OK)."""
    return int(rastr.rgm(arg))


def step_ut(rastr, direction: str):
    return rastr.step_ut(direction)


def fw_dynamic_run(rastr):
    fd = rastr.FWDynamic()
    fd.Run()


def get_chained_graph_snapshot(rastr, table_name: str, param: str, row: int, col: int):
    """Вернуть массив [v, t] для графика по таблице/параметру/строке."""
    return rastr.GetChainedGraphSnapshot(table_name, param, row, col)


def new_file(rastr, path: str):
    rastr.NewFile(path)
