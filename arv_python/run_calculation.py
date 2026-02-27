# -*- coding: utf-8 -*-
"""
Главный цикл расчёта: для каждого базового режима — нормальная и ремонтные схемы, DynCalc, get_pss_quality, отчёты.
"""
import os
import sys
from pathlib import Path
from typing import Callable, List, Optional

from comp_model import CompModel
from models import ComInfo, RemsInfo

from config import RESULT_XLSX, RESULT_DOCX, get_dynamics_rst, get_scenario_scn
from report_excel import write_result_xlsx, build_excel_rows_from_quality, RegimeBlock, SchemeBlock
from report_docx import write_result_docx

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
            col_calc,
            table_set_sel,
            rgm,
            step_ut,
            RG_REPL,
        )
        from calculation import (
            dyn_calc,
            get_pss_quality,
            get_mpd,
            get_grf,
            _RASTR_AVAILABLE,
        )
    except Exception:
        _RASTR_AVAILABLE = False
else:
    _RASTR_AVAILABLE = False


def _disable_elements(rastr, rem: RemsInfo) -> None:
    """Отключить элементы ремонтной схемы: vetv (ip/iq/np) или node (ny)."""
    table_vetv = get_table(rastr, "vetv")
    col_vetv_sta = get_col(table_vetv, "sta")
    table_node = get_table(rastr, "node")
    col_node_sta = get_col(table_node, "sta")
    for el in rem.elements:
        if el.iq != 0:
            table_set_sel(table_vetv, f"ip={el.ip}&iq={el.iq}&np={el.np}")
            col_calc(col_vetv_sta, "1")
        else:
            table_set_sel(table_node, f"ny={el.ip}")
            col_calc(col_node_sta, "1")


def _prepare_regime_file(rastr, scns_path: str, dynamics_rst: str, scenario_scn: str) -> None:
    """Загрузить сценарий, выставить com_dynamics (Tras и т.д.), сохранить не вызываем — вызывающий держит режим в rastr."""
    from config import get_avtomatika_dfw
    from rastr_client import new_file
    # В C# перед циклом по режимам этого нет; настройка com_dynamics делается внутри DynCalc через NewFile(автоматика) и Load(scns).
    # Оставляем без предварительной подготовки; dyn_calc сам делает new_file(автоматика) и load(scns).
    pass


def run_calculation(
    model: CompModel,
    grf_save: bool = True,
    rst_save: bool = False,
    csv_save: bool = False,
    progress_callback: Optional[Callable[[float, float], None]] = None,
) -> None:
    """
    Выполнить расчёт по модели: режимы из xlxs с s_name, для каждого — нормальная схема и ремонтные (по r_id).
    Результаты — Excel и Word в текущем каталоге.
    """
    if not _RASTR_AVAILABLE:
        raise RuntimeError("Rastr (ПК RUSTab) доступен только под Windows с установленным RUSTab и pywin32.")
    xlxs = model.xlxs
    if not xlxs or not xlxs.rgms:
        raise ValueError("Нет задания xlsx или режимов.")
    scns_path = model.scns
    if not scns_path or not os.path.isfile(scns_path):
        raise ValueError("Не найден файл тестового возмущения (.scn).")
    schs_path = model.schs
    dynamics_rst = get_dynamics_rst()
    scenario_scn = get_scenario_scn()
    cwd = os.getcwd()
    # Режимы с привязкой к файлу
    regims_with_file = [r for r in xlxs.rgms if r.s_name]
    if not regims_with_file:
        raise ValueError("Нет режимов с выбранным файлом (.rst).")
    total_steps = 0
    for rg in regims_with_file:
        total_steps += 1
        total_steps += len([rem for rem in xlxs.rems if rem.id in rg.r_id and rem.id != 0])
    step = 0.0
    regime_blocks: List[RegimeBlock] = []
    figures_for_doc: List[tuple] = []  # (caption, png_path)
    rg_num = 0
    for rg in regims_with_file:
        rg_num += 1
        path_reg = os.path.join(cwd, rg.name)
        try:
            os.makedirs(path_reg, exist_ok=True)
        except OSError:
            pass
        schemes: List[SchemeBlock] = []
        # Заголовок режима уже не добавляем отдельно в schemes — добавляем блок (regime_name, schemes)
        # Нормальная схема
        rastr = create_rastr()
        load_file(rastr, rg.s_name, dynamics_rst, RG_REPL)
        if rgm(rastr, "") != 0:
            schemes.append(("Нормальная схема", None))
            step += 1 + len([rem for rem in xlxs.rems if rem.id in rg.r_id and rem.id != 0])
            if progress_callback:
                progress_callback(step, total_steps)
            regime_blocks.append((rg.name, schemes))
            continue
        # Расчёт нормальной схемы
        grf_list = dyn_calc(rastr, xlxs, scns_path)
        if not grf_list:
            s_point = 0.0
            tras = 25.0
        else:
            s_point = grf_list[0].s_point
            try:
                table_com = get_table(rastr, "com_dynamics")
                col_tras = get_col(table_com, "Tras")
                tras = float(col_get_z(col_tras, 0))
            except Exception:
                tras = s_point + 25.0
        quality = get_pss_quality(grf_list, xlxs, s_point, tras)
        note_by_name = {}
        if grf_save:
            fig_num = 0
            for gdata in grf_list:
                if not gdata.pss_on or not gdata.pss_off:
                    continue
                fig_num += 1
                cap = f"Рисунок {rg_num}.0.{fig_num}"
                png_name = f"Рисунок_{rg_num}_0_{fig_num}.png"
                png_path = os.path.join(path_reg, png_name)
                get_grf(gdata.pss_on, gdata.pss_off, gdata.name, s_point - 1.0, tras, s_point, png_path)
                figures_for_doc.append((cap, png_path))
                note_by_name[gdata.name] = cap
        quality_with_notes = [(q[0], q[1], q[2], q[3], note_by_name.get(q[0])) for q in quality]
        excel_rows = build_excel_rows_from_quality(quality_with_notes)
        schemes.append(("Нормальная схема", excel_rows))
        step += 1
        if progress_callback:
            progress_callback(step, total_steps)
        # Ремонтные схемы
        rem_num = 0
        for rem in xlxs.rems:
            if rem.id not in rg.r_id or rem.id == 0:
                continue
            rem_num += 1
            load_file(rastr, rg.s_name, dynamics_rst, RG_REPL)
            _disable_elements(rastr, rem)
            ok = rgm(rastr, "") == 0
            need_mpd = rem.sch_id != -1 and rem.mpd_value != -1.0 and rem.ut_id != -1
            ut_info = next((u for u in xlxs.ut if u.id == rem.ut_id), None)
            ut_sname = ut_info.s_name if ut_info else ""
            if need_mpd and ut_sname and schs_path:
                get_mpd(rastr, rem.sch_id, rem.mpd_value, rem.ut_id, xlxs, schs_path, ut_sname)
            if not ok and not need_mpd:
                schemes.append((rem.name, None))
            else:
                grf_list_rem = dyn_calc(rastr, xlxs, scns_path)
                s_point_rem = grf_list_rem[0].s_point if grf_list_rem else 0.0
                quality_rem = get_pss_quality(grf_list_rem, xlxs, s_point_rem, tras)
                note_by_name_rem = {}
                if grf_save and grf_list_rem:
                    fig_num = 0
                    for gdata in grf_list_rem:
                        if not gdata.pss_on or not gdata.pss_off:
                            continue
                        fig_num += 1
                        cap = f"Рисунок {rg_num}.{rem_num}.{fig_num}"
                        png_name = f"Рисунок_{rg_num}_{rem_num}_{fig_num}.png"
                        png_path = os.path.join(path_reg, png_name)
                        get_grf(gdata.pss_on, gdata.pss_off, gdata.name, gdata.s_point - 1.0, tras, gdata.s_point, png_path)
                        figures_for_doc.append((cap, png_path))
                        note_by_name_rem[gdata.name] = cap
                quality_rem_notes = [(q[0], q[1], q[2], q[3], note_by_name_rem.get(q[0])) for q in quality_rem]
                excel_rows_rem = build_excel_rows_from_quality(quality_rem_notes)
                schemes.append((rem.name, excel_rows_rem))
            if rst_save:
                save_file(rastr, os.path.join(path_reg, f"{rem.name}.rst"), dynamics_rst)
            step += 1
            if progress_callback:
                progress_callback(step, total_steps)
        regime_blocks.append((rg.name, schemes))
    sheet_name = Path(xlxs.name).stem if xlxs.name else "Результат"
    write_result_xlsx(regime_blocks, sheet_name, os.path.join(cwd, RESULT_XLSX))
    write_result_docx(figures_for_doc, os.path.join(cwd, RESULT_DOCX))
