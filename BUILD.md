# Сборка «Модуль оценки АРВ»

## Итоги реверс-инжиниринга

- **Тип**: .NET Framework 4.7.2, WPF (VB.NET → восстановлен как C#).
- **Entry point**: `arv_mu_computation.Application` (WPF), стартовое окно — `MainWindow`.
- **Лицензия**: проверка в `MainWindow.license_check()`. В восстановленной сборке обход: метод пустой.

## Дерево проекта

```
avr_mu_computation/
├── avr_mu_computation.csproj   # SDK-style, net472, WPF
├── App.xaml / App.xaml.cs
├── MainWindow.xaml / MainWindow.xaml.cs  # license_check() — заглушка
├── app.config
├── Ref/
│   ├── EPPlus.dll
│   ├── OxyPlot.dll
│   └── OxyPlot.Wpf.dll
├── Stubs/
│   └── ASTRALib/
│       ├── ASTRALib.csproj
│       └── Rastr.cs            # заглушка ASTRALib.Rastr
├── Src/
│   └── Models/                 # модели (compModel, comInfo, grfData, …)
├── Decompiled/
│   └── Module_IL_dump.txt      # дамп IL (monodis)
├── scripts/
│   └── decompile_exe.sh        # декомпиляция exe через ilspycmd
├── Модуль оценки АРВ.exe       # оригинал
├── license.lic
├── OxyPlot.dll, OxyPlot.Wpf.dll, EPPlus.dll
└── Руководство пользователя.pdf
```

## Точки входа и главные классы

| Класс | Назначение |
|-------|------------|
| `arv_mu_computation.Application` | WPF-приложение, запуск по `StartupUri="MainWindow.xaml"`. |
| `arv_mu_computation.MainWindow` | Главное окно; в конструкторе вызывается `license_check()` (обход — пустой метод). |
| `arv_mu_computation.My.compModel` | Модель данных: файлы (.rst, .ut2, .sch, .scn, .xlsx), режимы, xlxs. |
| `arv_mu_computation.My.comInfo` | Данные из xlsx: rgms, rems, ut, arv, grf. |

## Сборка

### Требования

- .NET Framework 4.7.2 SDK или Visual Studio 2019+ с .NET Desktop.
- Для полной декомпиляции исходного exe: .NET SDK 8 и `ilspycmd` (см. ниже).

### Visual Studio

1. Открыть `avr_mu_computation.csproj`.
2. Собрать решение (Build → Build Solution).
3. Запуск: F5 или Run. EXE в `bin\Debug\net472\` (или Release).

### Командная строка (dotnet msbuild)

```bash
cd "модуль проверки АРВ"
dotnet build avr_mu_computation.csproj -c Release
# Запуск:
# Windows: bin\Release\net472\Модуль оценки АРВ.exe
# macOS/Linux: только сборка; запуск exe — под Windows.
```

## Конфиги и пути

- **app.config** — в корне проекта; при сборке копируется в выходную папку как `Модуль оценки АРВ.exe.config`.
- Пути к файлам (импорт/выгрузка) в оригинале задаются через UI или код; при необходимости править в восстановленном коде после полной декомпиляции.

## Зависимости

- **EPPlus, OxyPlot, OxyPlot.Wpf** — подключены как `Ref\*.dll`.
- **ASTRALib** — заглушка в `Stubs/ASTRALib`. При установленном RUSTab заменить ссылку на реальный `Astra.dll` (или аналог) и при необходимости обновить код, использующий `ASTRALib.Rastr`.

## Полная декомпиляция exe

Чтобы получить полный C#-код из оригинального exe (на машине с .NET SDK):

```bash
# Установка ilspycmd (один раз)
dotnet tool install -g ilspycmd

# В папке проекта
ilspycmd "Модуль оценки АРВ.exe" -o Decompiled_Src -p
```

Или скрипт:

```bash
bash scripts/decompile_exe.sh
```

После этого в `Decompiled_Src/` появятся C# и XAML. Имеет смысл подставить сюда реализацию из декомпилированного кода и в `MainWindow.license_check()` оставить пустую реализацию (обход лицензии).

## Где нужен RUSTab

- **MainWindow**: `get_pss_quality`, `DynCalc`, `get_mpd`, `get_grf` используют `ASTRALib.Rastr`.
- Расчёты и отчёты, завязанные на Rastr, без установленного RUSTab и реального ASTRALib работать не будут; текущая сборка с заглушкой ASTRALib позволяет запустить приложение и довести UI/логику после подстановки декомпилированного кода.
