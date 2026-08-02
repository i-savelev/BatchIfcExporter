# BatchIfcExporter

<img width="2800" height="2000" alt="изображение" src="https://github.com/user-attachments/assets/8695535a-5e83-4a6b-828c-6631213a5ab8" />

Плагин для пакетного экспорта IFC из Revit с гибкой настройкой через Excel и JSON.

## Основные возможности
- 📁 **Экспорт из локальной папки**: пакетная обработка всех `.rvt` файлов с выгрузкой в подпапку `IFC_Output`.
- ☁️ **Экспорт с Revit Server**: выбор и выгрузка моделей напрямую с сервера без локального копирования.

## Настройка через Excel
Плагин читает параметры из Excel-таблицы, сопоставляя `FileName` (подстроку имени модели) с первой подходящей строкой. Поддерживаемые колонки:
- `FileName`: часть имени файла `.rvt` для привязки настроек.
- `ViewName`: имя вида Revit для экспорта (по умолчанию `Navisworks`).
- `JsonConfigPath`: абсолютный путь к JSON-профилю настроек IFC Exporter.
- `MappingFilePath`: путь к пользовательскому mapping-файлу (имеет приоритет над JSON).
- `WorksetExcludePattern`: подстрока для исключения рабочих наборов при открытии (по умолчанию `Связь`), что экономит оперативную память.

| FileName | JsonConfigPath | ViewName | MappingFilePath | WorksetExcludePattern |
| --- | --- | --- | --- | --- |
| `AR_` | `C:\Configs\ifc_ar.json` | `Navisworks` | `C:\Configs\ar_mapping.txt` | `Связь` |
| `KR_` | `C:\Configs\ifc_kr.json` | `IFC Export` |  | `Link` |

## Установка
Поддерживаются версии Revit 2022-2026
1. Скачайте или соберите проект.
2. Скопируйте **все файлы** из папки `output_dll` в папку, настроенную для [PluginsManager](https://github.com/i-savelev/PluginsManager).

## Запуск и использование
1. Откройте Revit и перейдите на вкладку загрузчика (например, `ISTools` в PluginsManager).
2. Выберите команду: **Экспорт IFC из папки** или **Экспорт IFC с Revit сервера**.
3. В появившемся окне настройки:
   - Выберите исходную папку с `.rvt` или подключитесь к Revit Server и отметьте нужные модели в дереве.
   - (Опционально) Укажите путь к Excel-файлу конфигурации или нажмите кнопку сохранения, чтобы создать шаблон.
   - Нажмите кнопку запуска экспорта.
4. Следите за прогрессом в строке состояния формы. По завершении логи будут сохранены рядом с результатом (`export_log.log`).

## Стек и зависимости
Проект собирается как библиотека `.NET Framework 4.8`.
- [RevitLogger](https://github.com/i-savelev/RevitLogger) — внешняя библиотека для подробного логирования в файл и вывода debug-окна.
- [RevitServerBrowser](https://github.com/i-savelev/RevitServerBrowser) — отдельная библиотека для взаимодействия с Revit Server.
- `EPPlus 4.5.3.3`, `Newtonsoft.Json`.
