# Функциональные узлы (по метаданным и руководству)

## 1. Импорт файлов (drag&drop)

- **Расширения**: `.rst`, `.ut2`, `.sch`, `.xlsx`, `.scn`.
- **Модель**: `arv_mu_computation.My.compModel` — свойства `files`, `regims` (режимы из .rst), `ut2Files` (ut2s), `schFile` (schs), `scnsFile` (scns), `xlsFile` (xlxs — comInfo из xlsx).
- **Методы MainWindow**: `rgms_add_to_main_win(List<string>)`, `ut2_add_to_main_win(List<string>)`, `sch_add_to_main_win(string)` — добавление списков/файлов в главное окно.
- Обработчики (из strings exe): `start_Click`, `schs_MouseDoubleClick`, `rgms_MouseDoubleClick`, `uts_MouseDoubleClick` — кнопка «Старт» и двойной клик по спискам.

## 2. Отчёты .xlsx и .doc/.docx

- **EPPlus**: ссылка на `OfficeOpenXml.ExcelWorksheet` — формирование/запись xlsx.
- **Word**: в IL есть `Microsoft.Office.Interop.Word.Application`, `DocumentEvents2`, `wdAlignParagraphJustifyMed` — формирование отчётов в Word.
- **Методы MainWindow**: `get_pss_quality(..., ExcelWorksheet x, ..., Application doc, ...)`, `get_grf(..., Application doc)` — расчёт качества и графиков с выводом в Excel/Word.

## 3. Графики (OxyPlot)

- Сборки: `OxyPlot`, `OxyPlot.Wpf`.
- В коде: `grfData`, `datevalue` (t, v), `get_grf` — данные и построение графиков; имена вроде `set_DataFieldX`, `set_DataFieldY`, `get_Z`, `set_Z` указывают на привязку к графикам.

## 4. Прогресс расчёта

- В UI ожидаемо есть индикатор прогресса; в метаданных явно не выделен отдельный класс — искать в XAML/коде MainWindow (прогресс-бар или аналогичный контрол).

## 5. Проверка EPPlus/OxyPlot и автокопирование

- **Ресурсы exe**: встроенные `arv_mu_computation.EPPlus.dll`, `arv_mu_computation.OxyPlot.dll`, `arv_mu_computation.OxyPlot.Wpf.dll` (см. .mresource в IL).
- Класс `arv_mu_computation.My.Resources.Resources`: методы `get_EPPlus()`, `get_OxyPlot()`, `get_OxyPlot_Wpf()` возвращают `byte[]` — при старте, вероятно, проверяется наличие DLL в корне и при отсутствии они копируются из ресурсов.

## 6. Проверка license.lic

- **Класс**: `arv_mu_computation.MainWindow`.
- **Метод**: `license_check()` (instance, void) — метод №123 в таблице методов.
- В восстановленной сборке реализован **обход**: `license_check()` оставлен пустым (ничего не делает, не читает файл и не бросает исключений).
- В оригинале: чтение/проверка `license.lic`; при ошибке возможно `InvalidOperationException` или закрытие приложения — после полной декомпиляции можно увидеть точную логику.

## 7. Точка входа

- **Main**: `Application.Main()` (метод №119) — стандартный вход WPF-приложения.
- **Стартовое окно**: `MainWindow` (StartupUri в App.xaml).

## 8. RUSTab / ASTRALib

- **ASTRALib.Rastr**: используется в `get_pss_quality`, `DynCalc`, `get_mpd`, `get_grf`.
- В решении добавлена заглушка `Stubs/ASTRALib` (класс `Rastr`). Для полной работы расчётов нужна установка RUSTab и ссылка на реальный `Astra.dll` (или аналог).
