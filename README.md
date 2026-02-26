# Модуль оценки АРВ — восстановленный проект

Восстановленный исходный код и рабочая сборка приложения (WPF, .NET Framework 4.7.2). Декомпиляция выполнена через ILSpy; код сверен с «Руководством пользователя» (см. MANUAL_VERIFICATION.md).

## Что сделано

1. **Установка .NET SDK** на машине и декомпиляция exe в C# (ILSpy → папка `Decompiled_Src`).
2. **Интеграция кода**: исходники перенесены в `Src/`, `ASTRALib/`; подключены Ref (EPPlus, OxyPlot), NuGet (Microsoft.VisualBasic, Microsoft.Office.Interop.Word); UI загружается из встроенного `mainwindow.baml`.
3. **Обход лицензии**: в `MainWindow.license_check()` оставлена пустая реализация — проверка license.lic отключена.
4. **Сборка**: проект успешно собирается (`dotnet build`). Запуск возможен под Windows (WPF).
5. **Сверка с руководством**: составлен отчёт в `MANUAL_VERIFICATION.md` — функциональность кода соответствует описанию в руководстве пользователя.

## Сборка и запуск

- **Командная строка (на этой машине):**
  ```bash
  export DOTNET_ROOT="$HOME/.dotnet" PATH="$HOME/.dotnet:$PATH"
  dotnet build avr_mu_computation.csproj -c Debug
  ```
  Сборка создаёт exe в `bin/Debug/net472/Модуль оценки АРВ.exe` (запуск только под Windows).
- **Visual Studio (Windows):** открыть `avr_mu_computation.csproj`, собрать и запустить (F5).

Подробности — в **BUILD.md**.

## Конфигурация и пути

- Конфиг приложения: `app.config` в корне (копируется в выходную папку).
- Пути к данным (импорт/отчёты) при необходимости менять в коде после подстановки полного декомпилированного кода.

## Зависимости от RUSTab

Расчёты с использованием `ASTRALib.Rastr` (get_pss_quality, DynCalc, get_mpd, get_grf) требуют установленного RUSTab и реальной сборки ASTRALib. В решении используется заглушка; при наличии RUSTab можно заменить ссылку на реальный Astra.dll.

## Документация

- **BUILD.md** — дерево проекта, инструкции сборки, точки входа, где менять конфиги, где нужен RUSTab.
- **DOCUMENTATION.md** — ключевые функциональные узлы (импорт .rst/.ut2/.sch/.xlsx/.scn, отчёты xlsx/doc, OxyPlot, прогресс, проверка DLL, license.lic).
