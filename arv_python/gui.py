# -*- coding: utf-8 -*-
"""
Главное окно приложения (аналог MainWindow WPF).
Списки файлов: режимы .rst, траектории .ut2, сечения .sch; поля сценария и задания; чекбоксы; кнопка «Расчёт»; прогресс.
"""
import os
import sys
import tkinter as tk
from pathlib import Path
from tkinter import ttk, messagebox, filedialog

from comp_model import CompModel


class MainWindow:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Модуль оценки АРВ")
        self.root.minsize(500, 450)
        self.model = CompModel()

        # Переменные для полей и чекбоксов
        self.var_tvz = tk.StringVar()
        self.var_task = tk.StringVar()
        self.var_grf_save = tk.BooleanVar(value=True)
        self.var_rst_save = tk.BooleanVar(value=False)
        self.var_csv_save = tk.BooleanVar(value=False)

        self._progress_max = 0
        self._progress_value = tk.DoubleVar(value=0.0)
        self._build_ui()

    def _build_ui(self):
        f = ttk.Frame(self.root, padding=8)
        f.pack(fill=tk.BOTH, expand=True)

        # Блок: Расчётные режимы (*.rst)
        ttk.Label(f, text="Расчётные режимы (*.rst):").grid(row=0, column=0, columnspan=2, sticky=tk.W)
        list_rst = tk.Listbox(f, height=3, selectmode=tk.EXTENDED, font=("TkDefaultFont", 9))
        list_rst.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        list_rst.bind("<Double-1>", lambda e: self._clear_list(list_rst, "rst"))
        self.list_rst = list_rst
        ttk.Button(f, text="Добавить файлы...", command=self._add_files_rst).grid(row=2, column=0, sticky=tk.W, pady=2)

        # Блок: Траектории утяжеления (*.ut2)
        ttk.Label(f, text="Траектории утяжеления (*.ut2):").grid(row=3, column=0, columnspan=2, sticky=tk.W)
        list_ut2 = tk.Listbox(f, height=2, selectmode=tk.EXTENDED, font=("TkDefaultFont", 9))
        list_ut2.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        list_ut2.bind("<Double-1>", lambda e: self._clear_list(list_ut2, "ut2"))
        self.list_ut2 = list_ut2
        ttk.Button(f, text="Добавить файлы...", command=self._add_files_ut2).grid(row=5, column=0, sticky=tk.W, pady=2)

        # Блок: Контролируемые сечения (*.sch)
        ttk.Label(f, text="Контролируемые сечения (*.sch):").grid(row=6, column=0, columnspan=2, sticky=tk.W)
        list_sch = tk.Listbox(f, height=1, selectmode=tk.EXTENDED, font=("TkDefaultFont", 9))
        list_sch.grid(row=7, column=0, columnspan=2, sticky=(tk.W, tk.E))
        list_sch.bind("<Double-1>", lambda e: self._clear_list(list_sch, "sch"))
        self.list_sch = list_sch
        ttk.Button(f, text="Добавить файл...", command=self._add_file_sch).grid(row=8, column=0, sticky=tk.W, pady=2)

        # Тестовое возмущение и Расчётное задание
        ttk.Label(f, text="Тестовое возмущение (.scn):").grid(row=9, column=0, sticky=tk.W)
        self.entry_tvz = ttk.Entry(f, textvariable=self.var_tvz, width=40)
        self.entry_tvz.grid(row=9, column=1, sticky=(tk.W, tk.E))
        ttk.Label(f, text="Расчётное задание (.xlsx):").grid(row=10, column=0, sticky=tk.W)
        self.entry_task = ttk.Entry(f, textvariable=self.var_task, width=40)
        self.entry_task.grid(row=10, column=1, sticky=(tk.W, tk.E))

        # Кнопка «Добавить все файлы» (один каталог или несколько файлов)
        ttk.Button(f, text="Добавить файлы (выбор)...", command=self._add_files_drop).grid(row=11, column=0, columnspan=2, sticky=tk.W, pady=4)

        # Опции
        ttk.Separator(f, orient=tk.HORIZONTAL).grid(row=12, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=6)
        ttk.Checkbutton(f, text="Сохранять графики?", variable=self.var_grf_save).grid(row=13, column=0, columnspan=2, sticky=tk.W)
        ttk.Checkbutton(f, text="Сохранять режимы?", variable=self.var_rst_save).grid(row=14, column=0, columnspan=2, sticky=tk.W)
        ttk.Checkbutton(f, text="Сохранять *.xls?", variable=self.var_csv_save).grid(row=15, column=0, columnspan=2, sticky=tk.W)

        # Кнопка «Расчёт» и прогресс
        self.btn_start = ttk.Button(f, text="Расчёт", command=self._on_start)
        self.btn_start.grid(row=16, column=0, columnspan=2, pady=6)
        self.progress = ttk.Progressbar(f, variable=self._progress_value, maximum=100, mode="determinate")
        self.progress.grid(row=17, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=2)

        f.columnconfigure(1, weight=1)
        f.rowconfigure(1, weight=1)
        f.rowconfigure(4, weight=1)

    def _add_files_drop(self):
        """Выбор нескольких файлов — разнести по спискам по расширению."""
        paths = filedialog.askopenfilenames(
            title="Выберите файлы (.rst, .ut2, .sch, .scn, .xlsx)",
            filetypes=[
                ("Все подходящие", "*.rst *.ut2 *.sch *.scn *.xlsx *.xls"),
                ("Режимы Rastr", "*.rst"),
                ("Траектории", "*.ut2"),
                ("Сечения", "*.sch"),
                ("Сценарий", "*.scn"),
                ("Задание Excel", "*.xlsx *.xls"),
                ("Все файлы", "*.*"),
            ],
        )
        if not paths:
            return
        self.model.files = list(paths)
        self._refresh_lists_from_model()
        self.var_tvz.set(Path(self.model.scns).stem if self.model.scns else "")
        self.var_task.set(self.model.xlsx_name or "")

    def _add_files_rst(self):
        paths = filedialog.askopenfilenames(title="Режимы (*.rst)", filetypes=[("Rastr режимы", "*.rst")])
        if paths:
            self.model._rgms.extend(p for p in paths if p not in self.model._rgms)
            self.model._files = list(set(self.model._files + list(paths)))
            self._refresh_list_rst()

    def _add_files_ut2(self):
        paths = filedialog.askopenfilenames(title="Траектории (*.ut2)", filetypes=[("Траектории утяжеления", "*.ut2")])
        if paths:
            self.model._ut2_files.extend(p for p in paths if p not in self.model._ut2_files)
            self.model._files = list(set(self.model._files + list(paths)))
            self._refresh_list_ut2()

    def _add_file_sch(self):
        path = filedialog.askopenfilename(title="Сечения (*.sch)", filetypes=[("Сечения", "*.sch")])
        if path:
            self.model._sch_file = path
            if path not in self.model._files:
                self.model._files.append(path)
            self._refresh_list_sch()

    def _clear_list(self, listbox: tk.Listbox, kind: str):
        if kind == "rst":
            self.model.clear_regims()
            self._refresh_list_rst()
        elif kind == "ut2":
            self.model.clear_ut2s()
            self._refresh_list_ut2()
        elif kind == "sch":
            self.model.clear_schs()
            self._refresh_list_sch()
        listbox.delete(0, tk.END)

    def _refresh_lists_from_model(self):
        self._refresh_list_rst()
        self._refresh_list_ut2()
        self._refresh_list_sch()

    def _refresh_list_rst(self):
        self.list_rst.delete(0, tk.END)
        for p in self.model.regims:
            self.list_rst.insert(tk.END, Path(p).stem)

    def _refresh_list_ut2(self):
        self.list_ut2.delete(0, tk.END)
        for p in self.model.ut2s:
            self.list_ut2.insert(tk.END, Path(p).stem)

    def _refresh_list_sch(self):
        self.list_sch.delete(0, tk.END)
        if self.model.schs:
            self.list_sch.insert(tk.END, Path(self.model.schs).stem)

    def _on_start(self):
        """Запуск расчёта: проверка данных, диалог, вызов расчётного цикла в фоне."""
        if not self.model.files:
            messagebox.showwarning("Внимание", "Сначала добавьте файлы (.rst, .ut2, .sch, .scn, .xlsx).")
            return
        xlxs = self.model.xlxs
        if not xlxs or not xlxs.rgms:
            messagebox.showwarning("Внимание", "В задании xlsx не найдены режимы или файл задания не выбран.")
            return
        if not self.model.scns:
            messagebox.showwarning("Внимание", "Не найден файл тестового возмущения (.scn).")
            return
        if not messagebox.askyesno("Начало расчета", "Начать расчет?"):
            return
        self.btn_start.state(["disabled"])
        self.root.config(cursor="wait")
        try:
            from run_calculation import run_calculation
            run_calculation(
                self.model,
                grf_save=self.var_grf_save.get(),
                rst_save=self.var_rst_save.get(),
                csv_save=self.var_csv_save.get(),
                progress_callback=self._on_progress,
            )
            messagebox.showinfo(
                "Завершение расчета",
                "Расчет завершен.\nРезультаты сохранены в файлы:\n"
                + os.path.join(os.getcwd(), "Приложение. Результат в табличном виде.xlsx")
                + "\n"
                + os.path.join(os.getcwd(), "Приложение. Результат в графическом виде.docx"),
            )
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка работы алгоритма Модуля:\n{e}")
            import traceback
            traceback.print_exc()
        finally:
            self.root.config(cursor="")
            self.btn_start.state(["!disabled"])

    def _on_progress(self, value: float, maximum: float):
        self._progress_value.set(value)
        self.progress["maximum"] = max(maximum, 1)

    def run(self):
        self.root.mainloop()


def main():
    app = MainWindow()
    app.run()


if __name__ == "__main__":
    main()
