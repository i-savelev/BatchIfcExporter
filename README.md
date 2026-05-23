# BatchIfcExporter

Плагин для пакетного экспорта IFC из Revit.

Краткое описание возможностей:

- экспорт `.rvt` моделей из локальной папки;
- экспорт выбранных моделей с Revit Server;
- настройка экспорта через Excel-таблицу;
- загрузка IFC-параметров из JSON;
- переопределение mapping-файла на уровне отдельной модели;
- фильтрация рабочих наборов перед открытием модели;
- логирование в файл и вывод диагностического окна;
- workaround для бага `Autodesk IFC Exporter 2026`, который может изменять mapping-файл после первого экспорта.

## Возможности

### 1. Экспорт из локальной папки

Команда `BatchIFCExportCommand` показывает форму `ExportConfigForm`, в которой пользователь:

- выбирает папку с `.rvt` файлами;
- при необходимости указывает Excel-конфиг;
- может сохранить шаблон Excel-конфига;
- запускает пакетный экспорт.

Результат:

- все найденные `.rvt` из выбранной папки обрабатываются по очереди;
- IFC-файлы сохраняются в подпапку `IFC_Output`;
- лог пишется в `export_log.log` рядом с результатами.

### 2. Экспорт с Revit Server

Команда `ExportFromServerCommand` использует `RevitServerBrowser.dll` и позволяет:

- подключиться к доступным Revit Server;
- выбрать одну или несколько моделей;
- выбрать папку выгрузки;
- по желанию подключить Excel-конфиг;
- экспортировать IFC без предварительного копирования моделей локально.

Для этой команды IFC-файлы пишутся напрямую в выбранную пользователем папку.

### 3. Конфигурация по Excel

Excel-файл загружается через `ExcelIfcConfigLoader` и позволяет задавать параметры для отдельных моделей по совпадению части имени файла.

Поддерживаемые колонки:

| Колонка | Назначение |
| --- | --- |
| `FileName` | Подстрока для поиска нужной модели |
| `JsonConfigPath` | Путь к JSON-файлу с IFC-настройками |
| `ViewName` | Имя вида Revit для экспорта |
| `MappingFilePath` | Путь к пользовательскому mapping-файлу |
| `WorksetExcludePattern` | Подстрока для исключения рабочих наборов |

Логика сопоставления:

- для каждой модели берётся первая строка, у которой `FileName` найден в имени `.rvt`;
- если совпадение не найдено, используется конфигурация по умолчанию;
- `ViewName` по умолчанию: `Navisworks`;
- `WorksetExcludePattern` по умолчанию: `Связь`.

Пример:

| FileName | JsonConfigPath | ViewName | MappingFilePath | WorksetExcludePattern |
| --- | --- | --- | --- | --- |
| `AR_` | `C:\Configs\ifc_ar.json` | `Navisworks` | `C:\Configs\ar_mapping.txt` | `Связь` |
| `KR_` | `C:\Configs\ifc_kr.json` | `IFC Export` |  | `Link` |

### 4. Конфигурация IFC через JSON

Класс `IfcJsonConfigLoader` применяет JSON как внешний профиль для `IFCExportConfiguration`.

По сути это конфиг стандартного IFC Exporter for Revit. Такой JSON можно подготовить штатной утилитой экспортера Revit, а затем использовать в пакетном экспорте как готовый профиль настроек.

Особенности:

- если JSON не указан или не найден, применяются дефолтные значения;
- если часть свойств из JSON отсутствует в текущей версии `IFCExportConfiguration`, они пропускаются с записью в лог;
- `MappingFilePath` из Excel имеет приоритет над путём к mapping-файлу, указанному в JSON.

## Как работает открытие моделей

Класс `RvtDocument` открывает модель в два этапа:

1. Сначала документ открывается с `CloseAllWorksets` для анализа списка рабочих наборов.
2. Из списка исключаются наборы, содержащие `WorksetExcludePattern`.
3. Затем документ открывается повторно уже только с нужными рабочими наборами.

Это уменьшает объём загружаемых данных и помогает не тянуть в сессию лишние связи.

Дополнительно в коде есть защитная логика:

- безопасное закрытие документов;
- пропуск `Close()` для linked-документов;
- дополнительный `GC.Collect()` после закрытия;
- расширенная диагностика состояния сессии Revit.

## Настройка экспорта IFC

Класс `IfcExportConfig` отвечает за сборку `IFCExportOptions`.

Что он делает:

- ищет вид по имени (`ViewName`);
- поднимает `Autodesk.IFC.Export.UI.dll`;
- создаёт `IFCExportConfiguration` через рефлексию;
- применяет JSON-конфиг;
- при наличии переопределяет mapping-файл из Excel;
- вызывает `UpdateOptions(...)` с учётом различий API между версиями IFC Exporter.

Если нужный вид не найден, экспорт конкретной модели не выполняется, а ошибка уходит в лог.

## Защита mapping-файла для IFC Exporter 2026

В проекте есть `IfcMappingGuard`.

Назначение:

- перед первым экспортом создаётся резервная копия mapping-файла;
- после первого экспорта считается хеш файла;
- если Autodesk IFC Exporter изменил файл, оригинал восстанавливается;
- первая модель экспортируется повторно уже с восстановленным mapping-файлом.

Эта логика сейчас используется в `ExportFromServerCommand`.

## Логи и диагностика

Используются два механизма:

- `Logger` пишет подробный лог в UTF-8 файл;
- `IsDebugWindow` показывает накопленные сообщения в отдельном окне.

По умолчанию лог пишется во временную папку пользователя:

```text
%LOCALAPPDATA%\Temp\i-savelev\ifc_exporter.log
```

После выбора папки экспорта лог перенаправляется рядом с результатами:

```text
<папка_экспорта>\export_log.log
```

В лог попадают:

- найденные модели;
- найденные и пропущенные конфиги;
- информация о видах и mapping-файлах;
- ошибки открытия, экспорта и закрытия документов;
- итоговая статистика по успехам и ошибкам.

## Стек и зависимости

Проект собирается как `.NET Framework 4.8` class library.

Основные зависимости:

- `RevitAPI.dll`
- `RevitAPIUI.dll`
- `Autodesk.IFC.Export.UI.dll`
- `EPPlus 4.5.3.3`
- `Newtonsoft.Json 13.0.4`
- `RevitServerBrowser.dll` — отдельная библиотека для работы с Revit Server: [i-savelev/RevitServerBrowser](https://github.com/i-savelev/RevitServerBrowser)

`Autodesk.IFC.Export.UI.dll` ожидается в установленном Revit по пути вида:

```text
C:\Program Files\Autodesk\Revit <version>\AddIns\IFCExporterUI\
```

В репозитории также лежат локальные копии Revit API DLL для сборки проекта.

## Запуск

Проект собирается в `BatchIfcExporter.dll`.

Сейчас в репозитории нет `.addin`-манифеста, поэтому типовой сценарий запуска:

- через [PluginsManager](https://github.com/i-savelev/PluginsManager);
- или через [RevitAddInManager](https://github.com/chuongmep/RevitAddInManager).

Доступные команды:

- `Экспорт IFC из папки`
- `Экспорт IFC с Revit сервера`

## Ограничения и замечания

- Плагин работает только внутри процесса Revit.
- Для экспорта нужен корректный вид, указанный в `ViewName`.
- Excel-конфиг опционален, но без него будут использоваться только дефолтные настройки.
- Пути к JSON и mapping-файлам в Excel должны быть абсолютными и доступными с машины, где запущен Revit.
- Экспорт выполняется в транзакции Revit API.

## Структура проекта

Ключевые файлы:

- `BatchIfcExporter/BatchIFCExportCommand.cs` — экспорт локальных `.rvt` из папки;
- `BatchIfcExporter/ExportFromServerCommand.cs` — экспорт моделей с Revit Server;
- `BatchIfcExporter/ExportConfigForm.cs` — форма настройки локального экспорта;
- `BatchIfcExporter/RvtDocument.cs` — двухэтапное открытие моделей и фильтрация ворксетов;
- `BatchIfcExporter/IfcExportConfig.cs` — сборка `IFCExportOptions`;
- `BatchIfcExporter/IfcJsonConfigLoader.cs` — применение JSON-конфига;
- `BatchIfcExporter/ExcelIfcConfigLoader.cs` — загрузка Excel-конфига;
- `BatchIfcExporter/IfcMappingGuard.cs` — защита mapping-файла;
- `BatchIfcExporter/Logger.cs` — файловое логирование;
- `BatchIfcExporter/IsDebugWindow.cs` — окно быстрой диагностики.

## Репозиторий

GitHub: [i-savelev/BatchIfcExporter](https://github.com/i-savelev/BatchIfcExporter)
