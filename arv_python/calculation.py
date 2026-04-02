# -*- coding: utf-8 -*-
"""
Расчётные функции: DynCalc (динамика PSS вкл/выкл), get_pss_quality (степень демпфирования, эффективность), get_mpd, get_grf.
Соответствует логике MainWindow: get_pss_quality, DynCalc, get_mpd, get_grf.
"""
import math
import sys
from typing import List, Optional, Tuple

from models import DateValue, GrfData, GrfInfo, PssInfo, ComInfo

# Импорт Rastr только на Windows
if sys.platform == "win32":
    try:
        from rastr_client import (
            create_rastr,
            load_file,
            save_file,
            get_table,
            get_col,
            col_get_z,
            col_set_z,
            table_set_sel,
            table_count,
            table_find_next_sel,
            rgm,
            step_ut,
            fw_dynamic_run,
            get_chained_graph_snapshot,
            new_file,
            RG_REPL,
        )
        from config import (
            get_avtomatika_dfw,
            get_dynamics_rst,
            get_scenario_scn,
            get_sechen_sch,
            get_trajectory_ut2,
        )
        _RASTR_AVAILABLE = True
    except Exception:
        _RASTR_AVAILABLE = False
else:
    _RASTR_AVAILABLE = False


def _snapshot_to_series(snapshot, sign: int = 1) -> List[DateValue]:
    """Преобразовать результат GetChainedGraphSnapshot в список (t, v). COM возвращает массив [index, 0]=value, [index, 1]=time."""
    out = []
    try:
        if hasattr(snapshot, "GetDimensions"):
            # COM SafeArray
            dims = snapshot.GetDimensions()
            if len(dims) >= 2:
                for i in range(dims[0]):
                    try:
                        v = snapshot[i, 0]
                        t = snapshot[i, 1]
                        out.append(DateValue(t=float(t), v=float(v) * sign))
                    except (IndexError, TypeError):
                        continue
        else:
            for row in snapshot:
                if isinstance(row, (list, tuple)) and len(row) >= 2:
                    out.append(DateValue(t=float(row[1]), v=float(row[0]) * sign))
    except Exception:
        pass
    return out


def dyn_calc(rastr, com_info: ComInfo, scns_path: str) -> List[GrfData]:
    """
    Расчёт динамики: сначала с PSS включённым (собрать pssOn по каждому grf), затем изменить arv и перезапустить — собрать pssOff.
    """
    shabl_dfw = get_avtomatika_dfw()
    shabl_scn = get_scenario_scn()
    new_file(rastr, shabl_dfw)
    load_file(rastr, scns_path, shabl_scn, RG_REPL)
    fw_dynamic_run(rastr)
    table_act = get_table(rastr, "DFWAutoActionScn")
    col_ts = get_col(table_act, "TimeStart")
    s_point = float(col_get_z(col_ts, 0))
    if s_point == 0.0:
        try:
            table_log = get_table(rastr, "DFWAutoLogicScn")
            if table_count(table_log) > 0:
                col_delay = get_col(table_log, "Delay")
                s_point = float(col_get_z(col_delay, 0))
        except Exception:
            pass

    list_grf_data: List[GrfData] = []

    for g in com_info.grf:
        series_on: List[DateValue] = []
        sign = 1
        tbl = get_table(rastr, g.table)
        col_param = get_col(tbl, g.param)
        col_sta = get_col(tbl, "sta")
        for idx in g.id:
            try:
                if g.table == "vetv" and float(col_get_z(col_param, idx)) < 0:
                    sign = -1
                sta_val = col_get_z(col_sta, idx)
                # В исходном C# обрабатываются только строки, где sta == false.
                # Поэтому пропускаем все "включенные" элементы.
                if sta_val not in (False, 0, "0", "False", "false", None):
                    continue
            except Exception:
                continue
            snap = get_chained_graph_snapshot(rastr, g.table, g.param, idx, 0)
            row_series = _snapshot_to_series(snap, sign)
            if not series_on and row_series:
                series_on = row_series
            elif series_on and row_series:
                for i, r in enumerate(row_series):
                    if i < len(series_on):
                        series_on[i] = DateValue(t=series_on[i].t, v=series_on[i].v + r.v)
        if series_on:
            list_grf_data.append(GrfData(name=g.name, pss_on=series_on, s_point=s_point))

    # Изменить настройки АРВ (выключить PSS)
    for p in com_info.arv:
        tbl = get_table(rastr, p.table)
        col_p = get_col(tbl, p.param)
        col_set_z(col_p, p.id, p.val)
    load_file(rastr, scns_path, shabl_scn, RG_REPL)
    fw_dynamic_run(rastr)

    # Собрать pssOff для каждого grf
    for gi, g in enumerate(com_info.grf):
        series_off: List[DateValue] = []
        sign = 1
        tbl = get_table(rastr, g.table)
        col_param = get_col(tbl, g.param)
        col_sta = get_col(tbl, "sta")
        for idx in g.id:
            try:
                if g.table == "vetv" and float(col_get_z(col_param, idx)) < 0:
                    sign = -1
                sta_val = col_get_z(col_sta, idx)
                if sta_val not in (False, 0, "0", "False", "false", None):
                    continue
            except Exception:
                continue
            snap = get_chained_graph_snapshot(rastr, g.table, g.param, idx, 0)
            row_series = _snapshot_to_series(snap, sign)
            if not series_off and row_series:
                series_off = row_series
            elif series_off and row_series:
                for i, r in enumerate(row_series):
                    if i < len(series_off):
                        series_off[i] = DateValue(t=series_off[i].t, v=series_off[i].v + r.v)
        if gi < len(list_grf_data) and series_off:
            list_grf_data[gi].pss_off = series_off

    return list_grf_data


def get_pss_quality(
    grf_list: List[GrfData],
    com_info: ComInfo,
    s_point: float,
    tras: float,
) -> List[Tuple[str, float, str, str, Optional[str]]]:
    """
    По результатам DynCalc для каждого grf вычислить: степень демпфирования, демпфирование (Да/Нет), эффективность PSS, примечание (рисунок).
    Возвращает список кортежей (name, degree, damping_text, efficiency_text, figure_ref).
    """
    if not grf_list:
        return []
    results = []
    for g in com_info.grf:
        gdata = next((x for x in grf_list if x.name == g.name), None)
        if not gdata or not gdata.pss_on:
            continue
        pss_on = gdata.pss_on
        pss_off = gdata.pss_off or []
        after = [x for x in pss_on if x.t > s_point + 15.0]
        before = [x for x in pss_on if (s_point - 2.0 < x.t < s_point)]
        if not after or not before:
            continue
        num2 = max(x.v for x in after) - min(x.v for x in after)
        num4 = sum(x.v for x in before) / len(before) if before else 0.0
        num3 = max(x.v for x in pss_on if x.t > s_point) - num4
        if num3 == 0:
            continue
        degree = num2 / num3
        if degree > 0.01:
            max1 = max(x.v for x in after)
            t_max1 = next(x.t for x in after if x.v == max1)
            after_t = [x for x in pss_on if x.t > t_max1]
            min1 = min(x.v for x in after_t) if after_t else max1
            t_min1 = next((x.t for x in after_t if x.v == min1), t_max1)
            after_t2 = [x for x in pss_on if x.t > t_min1]
            max2 = max(x.v for x in after_t2) if after_t2 else min1
            t_max2 = next((x.t for x in after_t2 if x.v == max2), t_min1)
            freq = 1.0 / (t_max1 - t_max2) if (t_max1 - t_max2) != 0 else 0
            if 0.6 <= freq <= 2.0:
                damping_text = "Не демпфируется"
            else:
                damping_text = "Демпфируется*"
        else:
            damping_text = "Демпфируется"

        num5 = sum((x.v - num4) ** 2 * x.t for x in pss_on if s_point + 0.06 < x.t < s_point + 25.0)
        num6 = sum((x.v - num4) ** 2 * x.t for x in pss_off if s_point + 0.06 < x.t < s_point + 25.0)
        efficiency_text = "Эффективен" if num5 < num6 else "Не эффективен"
        results.append((g.name, degree, damping_text, efficiency_text, None))
    return results


def get_mpd(rastr, sch_id: int, val: float, ut_id: int, com_info: ComInfo, schs_path: str, ut_sname: str):
    """Ограничение перетоков по траектории утяжеления до заданного МДП по сечению."""
    shabl_ut = get_trajectory_ut2()
    shabl_sch = get_sechen_sch()
    load_file(rastr, ut_sname, shabl_ut, RG_REPL)
    while rgm(rastr, "") != 0:
        step_ut(rastr, "z")
    rgm(rastr, "")
    load_file(rastr, schs_path, shabl_sch, RG_REPL)
    load_file(rastr, ut_sname, shabl_ut, RG_REPL)
    table_ut = get_table(rastr, "ut_common")
    col_kfc = get_col(table_ut, "kfc")
    table_sec = get_table(rastr, "sechen")
    col_psech = get_col(table_sec, "psech")
    table_set_sel(table_sec, f"ns={sch_id}")
    num2 = table_find_next_sel(table_sec, -1)
    if float(col_get_z(col_psech, num2)) <= val:
        return
    num4 = float(col_get_z(col_psech, num2))
    num3 = num4
    while abs(val - num3) > 1.0:
        col_get_z(col_kfc, 0)
        step_ut(rastr, "z")
        rgm(rastr, "")
        num3 = float(col_get_z(col_psech, num2))
        col_set_z(col_kfc, 0, (num3 - val) / (num4 - num3) if (num4 - num3) != 0 else 0)
        num4 = num3


def get_grf(
    pss_on: List[DateValue],
    pss_off: List[DateValue],
    name: str,
    s: float,
    e: float,
    s_pss: float,
    png_path: str,
) -> str:
    """
    Строит два графика (полный и укрупнённый): PSS вкл (красный) и PSS выкл (синий).
    Сохраняет один PNG с двумя подграфиками по вертикали. Возвращает png_path.
    """
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    if not pss_on or not pss_off:
        return png_path
    t_on = [x.t for x in pss_on]
    v_on = [x.v for x in pss_on]
    t_off = [x.t for x in pss_off]
    v_off = [x.v for x in pss_off]
    num3 = min(s_pss * 1.005, s_pss + 0.5)
    num5 = min(s_pss * 0.001, 0.5)
    after_s = [x for x in pss_on if abs(x.v - s_pss) > num5]
    num4 = max(x.t for x in after_s) if after_s else s + 25.0
    num = next((x.t for x in pss_on if x.v < num3 and x.t > s + 0.06), num4)
    num2 = next((x.t for x in pss_off if x.v < num3 and x.t > s + 0.06), num4)
    num6 = min(num, num2)
    on_after = [x for x in pss_on if x.t > num6]
    off_after = [x for x in pss_off if x.t > num6]
    y_min = min(
        min(x.v for x in on_after) if on_after else 0,
        min(x.v for x in off_after) if off_after else 0,
    )
    y_max = max(
        max(x.v for x in on_after) if on_after else 0,
        max(x.v for x in off_after) if off_after else 0,
    )
    x_end = min(num4, s + 25.0, e)

    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 8))
    fig.suptitle(name, fontsize=14)
    ax1.set_xlim(s - 1.0, e)
    ax1.set_xlabel("Время, с")
    ax1.set_ylabel("Переток, МВт")
    ax1.grid(True)
    ax1.plot(t_on, v_on, "r-", linewidth=2, label="PSS введен")
    ax1.plot(t_off, v_off, "b-", linewidth=2, label="PSS выведен")
    ax1.legend(loc="lower center")
    ax2.set_xlim(num6, x_end)
    ax2.set_ylim(y_min, y_max)
    ax2.set_xlabel("Время, с")
    ax2.set_ylabel("Переток, МВт")
    ax2.grid(True)
    ax2.plot(t_on, v_on, "r-", linewidth=2, label="PSS введен")
    ax2.plot(t_off, v_off, "b-", linewidth=2, label="PSS выведен")
    ax2.legend(loc="lower center")
    plt.tight_layout()
    plt.savefig(png_path, dpi=100, bbox_inches="tight")
    plt.close()
    return png_path
