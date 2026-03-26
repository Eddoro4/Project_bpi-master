# Установщик `Project_bpi`

В папке лежат файлы для сборки Windows-установщика через Inno Setup.

## Что внутри

- `Project_bpi.iss` — сценарий установщика
- `BuildInstaller.ps1` — PowerShell-скрипт для запуска сборки установщика

## Что нужно перед сборкой

1. Собрать приложение `Project_bpi`
2. Установить Inno Setup 6

## Быстрый запуск

```powershell
powershell -ExecutionPolicy Bypass -File .\installer\BuildInstaller.ps1
```

Скрипт сначала ищет `Project_bpi.exe` в:

- `Project_bpi\bin\Release`
- `Project_bpi\bin\Debug`

Если нужно указать папку со сборкой вручную:

```powershell
powershell -ExecutionPolicy Bypass -File .\installer\BuildInstaller.ps1 -BuildOutputDir ".\Project_bpi\bin\Release"
```

Готовый установщик будет создан в папке:

```text
installer\output
```
