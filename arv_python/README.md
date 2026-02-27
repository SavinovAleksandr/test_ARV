# Модуль оценки АРВ (Python)

Версия на Python с тем же функционалом, что и исходное приложение на C#/WPF.

## Требования

- Python 3.8+
- **Windows** — для расчётов нужен установленный **ПК RUSTab (Rastr)** и при необходимости **pywin32** (`pip install pywin32`).
- Остальные зависимости: `openpyxl`, `python-docx`, `matplotlib`.

## Установка

```bash
cd arv_python
pip install -r requirements.txt
```

На Windows для доступа к Rastr через COM:

```bash
pip install pywin32
```

## Запуск

```bash
python main.py
```

Откроется окно: добавьте файлы (.rst, .ut2, .sch, .scn, .xlsx), укажите при необходимости сценарий и задание, включите опции и нажмите «Расчёт». Результаты сохраняются в текущий каталог:

- `Приложение. Результат в табличном виде.xlsx`
- `Приложение. Результат в графическом виде.docx`

Шаблоны Rastr должны находиться в каталоге `Documents\RastrWin3\SHABLON\` (как в оригинале).
