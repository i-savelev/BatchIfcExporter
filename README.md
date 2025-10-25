# Пакетный экспорт IFC из Revit: Базовый скрипт

Тут описываю работу скрипта для экспорта файлов IFC из Revit с помощью Revit api. Скрипт простой, не имеет интерфейса и может быть использован как пример или быть расширен.

Как это работает?
- данный способ автоматически вызывает стандартную утилиту экспорта IFC.
- нужно открыть Revit и создать проект, т.к. скрипт может выполняться только из среды Revit
- выбирается папка с файлами Revit и файл мэппинга в двух последовательных диалогах
- происходит экспорт как при ручном запуске, но каждая модель из папки открывается автоматически

Далее подробно опишу алгоритм и основные нюансы с точки зрения кода.

Помимо стандартных RevitAPI.dll и RevitAPIUI.dll необходимо подключить еще Autodesk.IFC.Export.UI.dll. Autodesk.IFC.Export.UI.dll - отвечает за подключение к стандартной утилите экспорта. Эти файлы добавлены в папку RevitAPI репозитория.
Так же эти файлы можно найти в папке `c:\Program Files\Autodesk\Revit 2022\IFCExporterUI`

Репозиторий на github: [GitHub - i-savelev/BatchIfcExporter](https://github.com/i-savelev/BatchIfcExporter)

---

## 1. Открытие RVT-файлов с фильтрацией рабочих наборов

Класс `RvtDocument` отвечает за открытие документа с необходимыми наборами.

Одной из главных особенностей скрипта является **двухэтапное открытие файла**, что позволяет избежать загрузки ненужных данных (например, связанных моделей).

### Этап 1: Анализ рабочих наборов

Файл открывается временно — только для чтения списка рабочих наборов и получения списка id наборов, которые содержат связанные файлы. Для этого документ открывается с закрытыми рабочими наборами:

```csharp
// Открываем документ с закрытыми рабочими наборами — быстро и без нагрузки
OpenOptions openOptionsStep1 = new OpenOptions();
openOptionsStep1.SetOpenWorksetsConfiguration(
    new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets)
);
openOptionsStep1.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets;

Document docTemp = App.OpenDocumentFile(modelPath, openOptionsStep1);
```

Затем фильтруем рабочие наборы, исключая те, что содержат слово «Связь» и получаем список из id:

```csharp
// Получаем все рабочие наборы
var allWorksets = new FilteredWorksetCollector(docTemp).ToWorksets();

// Оставляем ТОЛЬКО наборы, НЕ содержащие слово "Связь"
var worksetIdsToOpen = allWorksets
    .Where(w => !w.Name.Contains(wsExcludeWord)) // например, wsExcludeWord = "Связь"
    .Select(w => w.Id)
    .ToList();
```

После анализа временный документ закрывается:

```csharp
docTemp?.Close(false); // Без сохранения
```

### Этап 2: Окончательное открытие с нужными наборами

Во время повторного открытия файла открываются только те рабочие наборы, которые были отфильтрованы на предыдущем шаге:

```csharp
// Создаём конфигурацию: сначала всё закрыто, затем открываем нужные наборы
WorksetConfiguration finalWorksetConfig = 
    new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets);
finalWorksetConfig.Open(worksetIdsToOpen);

// Открываем документ уже с фильтрацией
OpenOptions openOptionsFinal = new OpenOptions();
openOptionsFinal.SetOpenWorksetsConfiguration(finalWorksetConfig);
openOptionsFinal.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets;

Doc = App.OpenDocumentFile(modelPath, openOptionsFinal);
```

> 💡 **Почему это важно?**  
> В больших проектах связанные модели (архитектура, конструкции, ИТП и т.д.) часто хранятся в отдельных рабочих наборах с пометкой «Связь». Их исключение ускоряет экспорт и уменьшает размер IFC-файла.

---

## 2. Настройка параметров экспорта IFC

Класс `IfcExportConfig` централизует все настройки. Вот как формируется конфигурация:

### Поиск целевого вида

Для экспорта задается вид, с которого будет экспортирована геометрия:
```csharp
// Ищем вид по имени (например, ViewName = "Navisworks")
Autodesk.Revit.DB.View targetView = new FilteredElementCollector(Doc)
    .OfClass(typeof(Autodesk.Revit.DB.View))
    .Cast<Autodesk.Revit.DB.View>()
    .FirstOrDefault(v => 
        v.Name.Equals(ViewName, StringComparison.OrdinalIgnoreCase) && 
        !v.IsTemplate
    );
```

Если вид не найден — скрипт сообщает об этом, но не прерывает выполнение.

### Применение настроек через IFCExportConfiguration

Все поля настрое вынесены в отдельные свойства класса. Они полностью соответствуют параметрам из окна стандартного "IFC Exporter for Revit"

```csharp
// --- Общие настройки ---
public IFCVersion IFCVersion { get; set; } = IFCVersion.IFC4RV;
public int SpaceBoundaries { get; set; } = 0;
public bool SplitWallsAndColumns { get; set; } = false;

// --- Дополнительное содержимое ---
public bool Export2DElements { get; set; } = false;
public bool ExportLinkedFiles { get; set; } = false;

// --- Наборы свойств ---
public bool ExportIFCCommonPropertySets { get; set; } = false;
public bool ExportBaseQuantities { get; set; } = false;

```

Далее эти параметры передаются в конфигурацию `IFCExportConfiguration`
Тут же указывается путь к файлу мэппинга при его наличии.

```csharp
// Создаём стандартную конфигурацию экспорта
BIM.IFC.Export.UI.IFCExportConfiguration config = 
    BIM.IFC.Export.UI.IFCExportConfiguration.CreateDefaultConfiguration();

// Задаём параметры
config.IFCVersion = IFCVersion.IFC4RV;
config.SpaceBoundaries = 0; // без границ помещений
config.ExportLinkedFiles = false; // не экспортировать связи
config.VisibleElementsOfCurrentView = true; // только видимые в виде элементы

// Подключаем файл сопоставления свойств (если существует)
if (File.Exists(MappingFile))
{
    config.ExportUserDefinedPsets = true;
    config.ExportUserDefinedPsetsFileName = MappingFile;
}
```

### Формирование финальных опций экспорта

```csharp
// Создаём объект IFCExportOptions и применяем к нему конфигурацию
IFCExportOptions exportOptions = new IFCExportOptions();
config.UpdateOptions(exportOptions, targetView.Id); // ← важно: указываем ViewId!
```

> ⚠️ **Важно**: метод `UpdateOptions` требует `ElementId` вида — именно поэтому экспорт привязан к конкретному виду.

---

## 3. Пакетная обработка файлов

Основной цикл обработки класса `BatchIFCExportCommand`:

```csharp
foreach (string rvtPath in rvtFiles)
{
    // 1. Открываем документ с фильтрацией
    var rvtDoc = new RvtDocument(app, rvtPath);
    var doc = rvtDoc.Open("Связь"); // исключаем наборы со словом "Связь"

    // 2. Настраиваем экспорт
    var config = new IfcExportConfig(doc, viewName, mappingFilePath);
    var exportOptions = config.GetConfig();

    // 3. Формируем путь для IFC-файла (заменяем расширение)
    string ifcPath = Path.ChangeExtension(rvtPath, ".ifc");
    string folder = Path.GetDirectoryName(ifcPath);
    string filename = Path.GetFileName(ifcPath);

    // 4. Экспортируем в транзакции (обязательно по API Revit)
    using (Transaction tx = new Transaction(doc))
    {
        tx.Start("IFC Export");
        doc.Export(folder, filename, exportOptions);
        tx.Commit();
    }

    // 5. Закрываем документ
    doc.Close(false);
}
```

> 📌 **Примечание**: Хотя экспорт не изменяет модель, Revit API требует выполнения операций в транзакции — даже для чтения/экспорта.

---

## 4. Нюансы и особенности

Для вывода информации я исползую свой класс `IsDebugWindow`, который создает окошко с таблицей.

### Обработка ошибок

Скрипт не падает при отсутствии вида или файла сопоставления:

```csharp
if (targetView == null)
{
    IsDebugWindow.AddRow($"Вид '{ViewName}' не найден в файле {rvtPath}");
    continue; // переходим к следующему файлу
}
```

Аналогично — если файл сопоставления не найден:

```csharp
if (!File.Exists(MappingFile))
{
    IsDebugWindow.AddRow($"Файл сопоставления не найден: {MappingFile}");
    // но экспорт всё равно выполняется без пользовательских Pset
}
```

### Отладка и логирование

Весь процесс логируется в специальное окно:

```csharp
IsDebugWindow.AddRow($"Экспорт завершён: {ifcPath}");
IsDebugWindow.Show(); // показываем лог в конце
```

Это упрощает диагностику при обработке десятков файлов.

## 5. Запуск и установка

Скрипт не имеет интерфейса и не добавляется на панель. Запустить его можно с помощью моего плагина [PluginsManager](https://github.com/i-savelev/PluginsManager). Для этого нужно указать путь до папки с файлом BatchIfcExporter.dll. Подробнее смотри инструкцию к PluginsManager.

Или с помощью [RevitAddInManager](https://github.com/chuongmep/RevitAddInManager?ysclid=mh6lsbgmoo386554988). так же указав путь до файла BatchIfcExporter.dl.
