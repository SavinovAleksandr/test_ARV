#!/bin/bash
# Декомпиляция exe в C# (требуется .NET SDK: dotnet --version).
# На Windows можно использовать: ilspycmd "Модуль оценки АРВ.exe" -o Decompiled_Src -p
set -e
cd "$(dirname "$0")/.."
EXE="Модуль оценки АРВ.exe"
OUT="Decompiled_Src"
if [[ ! -f "$EXE" ]]; then
  echo "Файл $EXE не найден"
  exit 1
fi
export PATH="$PATH:$HOME/.dotnet/tools"
if ! command -v ilspycmd &>/dev/null; then
  echo "Установка ilspycmd..."
  dotnet tool install -g ilspycmd
fi
echo "Декомпиляция $EXE -> $OUT ..."
ilspycmd "$EXE" -o "$OUT" -p
echo "Готово. Исходники в $OUT/"
