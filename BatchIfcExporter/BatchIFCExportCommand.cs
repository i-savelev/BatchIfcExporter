using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using BatchIfcExporter;
using ISTools;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace BatchExportIfc
{
    [Autodesk.Revit.Attributes.TransactionAttribute(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class BatchIFCExportCommand : IExternalCommand
    {
        public static string IS_TAB_NAME => "ISTools";
        public static string IS_NAME => "Экспорт IFC из папки";
        public static string IS_IMAGE => "BatchIfcExporter.Resources.ifc_to_folder.png";
        public static string IS_DESCRIPTION => "Автор: https://github.com/i-savelev\r\nПакетный экспорт IFC моделей из выбранной папки в с конфигурацией из Excel";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                Logger.Clear();
                Logger.Info("BatchIFCExportCommand", "▶ Открытие окна экспорта моделей из папки");

                // 🔹 1. Создаём и настраиваем форму
                var form = new ExportConfigForm();
                form.Text = "Пакетный экспорт IFC";

                // 🔹 2. Подписываемся на кнопки — ОБЯЗАТЕЛЬНО до ShowDialog()

                // Кнопка: Выбор папки с моделями
                form.BtnSelectIfcFolder.Click += (s, e) =>
                {
                    using (var dlg = new OpenFileDialog
                    {
                        Title = "Выберите папку с моделями Revit",
                        Filter = "Папки|*",                    // 👈 Хак: показываем только папки
                        CheckFileExists = false,               // 👈 Не требуем файл
                        CheckPathExists = true,                // 👈 Но путь должен существовать
                        ValidateNames = false,                 // 👈 Разрешаем любое имя
                        FileName = "Выберите папку"            // 👈 Подсказка в поле ввода
                    })
                    {
                        if (dlg.ShowDialog(form) == DialogResult.OK)
                        {
                            // 👈 Извлекаем путь из выбранного "файла"
                            string folderPath = Path.GetDirectoryName(dlg.FileName);

                            // Если пользователь ввёл путь вручную и нажал ОК — используем его
                            if (string.IsNullOrEmpty(folderPath) && Directory.Exists(dlg.FileName))
                                folderPath = dlg.FileName;

                            if (!string.IsNullOrEmpty(folderPath) && Directory.Exists(folderPath))
                            {
                                form.TxtIfcFolder.Text = folderPath;
                                form.SetStatus($"Папка: {Path.GetFileName(folderPath)}");
                                Logger.Info("BatchIFCExportCommand", $"Выбрана папка: {folderPath}");
                            }
                            else
                            {
                                form.SetStatus("⚠️ Неверный путь");
                                MessageBox.Show(form, "Выберите корректную папку", "Внимание",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }
                    }
                };

                // Кнопка: Выбор файла конфигурации
                form.BtnSelectConfigFile.Click += (s, e) =>
                {
                    using (var dlg = new OpenFileDialog
                    {
                        Title = "Выберите Excel конфигурацию экспорта",
                        Filter = "Excel Files|*.xlsx;*.xls|All Files|*.*",
                        CheckFileExists = true
                    })
                    {
                        if (dlg.ShowDialog(form) == DialogResult.OK)
                        {
                            form.TxtConfigFile.Text = dlg.FileName;
                            form.SetStatus($"📋 Конфиг: {Path.GetFileName(dlg.FileName)}");
                            Logger.Info("BatchIFCExportCommand", $"Выбран конфиг: {dlg.FileName}");
                        }
                    }
                };

                // Кнопка: Сохранить шаблон конфигурации
                form.BtnSaveTemplate.Click += (s, e) => SaveConfigTemplate(form);

                // Кнопка: Запуск экспорта
                form.BtnRunExport.Click += (s, e) => RunExport(commandData, form);

                // Кнопка: Закрыть (стандартное поведение)
                form.BtnClose.Click += (s, e) => form.Close();

                // 🔹 3. Показываем форму
                form.ShowDialog();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                Logger.Critical("BatchIFCExportCommand", $"❌ Критическая ошибка: {ex}\n{ex.StackTrace}");
                IsDebugWindow.AddRow($"ERROR: {ex.Message}");
                IsDebugWindow.Show();
                return Result.Failed;
            }
        }

        /// <summary>
        /// Создаёт и сохраняет шаблон конфигурации в формате .xlsx через EPPlus.
        /// </summary>
        private void CreateTemplateFile(string path)
        {
            // ⚠️ EPPlus v5+ требует указания контекста лицензии. 
            // Установи один раз при старте плагина или оставь здесь.

            if (!path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                path += ".xlsx";

            using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(path)))
            {
                var ws = package.Workbook.Worksheets.Add("Конфигурация");

                // Заголовки
                string[] headers = { "FileName", "JsonConfigPath", "ViewName", "MappingFilePath", "WorksetExcludePattern" };
                for (int i = 0; i < headers.Length; i++)
                    ws.Cells[1, i + 1].Value = headers[i];

                // Стилизация заголовков
                var headerRange = ws.Cells["A1:E1"];
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                headerRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(242, 242, 242)); // Светло-серый
                headerRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                // Пример строки
                var example = new object[] { "Project_A.rvt", "Navisworks", "", "", "Связь*" };
                for (int i = 0; i < example.Length; i++)
                    ws.Cells[2, i + 1].Value = example[i];

                // Автоподбор ширины колонок
                ws.Cells["A1:E2"].AutoFitColumns();

                package.Save();
            }
        }

        /// <summary>
        /// Диалог сохранения шаблона.
        /// </summary>
        private void SaveConfigTemplate(ExportConfigForm form)
        {
            try
            {
                using (var dlg = new SaveFileDialog
                {
                    Title = "Сохранить шаблон конфигурации",
                    Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                    FileName = "IFC_Export_Config_Template.xlsx",
                    OverwritePrompt = true
                })
                {
                    if (dlg.ShowDialog(form) == DialogResult.OK)
                    {
                        CreateTemplateFile(dlg.FileName);
                        form.SetStatus("✅ Шаблон сохранён");
                        Logger.Info("BatchIFCExportCommand", $"Шаблон сохранён: {dlg.FileName}");
                        TaskDialog.Show("Успех", "Шаблон конфигурации успешно создан!");
                    }
                }
            }
            catch (Exception ex)
            {
                form.SetStatus("❌ Ошибка сохранения");
                Logger.Error("BatchIFCExportCommand", $"Ошибка сохранения шаблона: {ex.Message}");
                TaskDialog.Show("Ошибка", $"Не удалось сохранить шаблон:\n{ex.Message}");
            }
        }

        /// <summary>
        /// Запускает процесс экспорта с валидацией входных данных.
        /// </summary>
        private void RunExport(ExternalCommandData commandData, ExportConfigForm form)
        {
            var folderPath = form.IfcFolderPath;
            var excelConfigPath = form.ConfigFilePath;

            // 🔹 Валидация
            if (string.IsNullOrEmpty(folderPath) || !Directory.Exists(folderPath))
            {
                TaskDialog.Show("Внимание", "Выберите корректную папку с моделями Revit");
                form.SetStatus("⚠️ Требуется папка с моделями");
                return;
            }

            if (!string.IsNullOrEmpty(excelConfigPath) && !File.Exists(excelConfigPath))
            {
                TaskDialog.Show("Внимание", "Указанный файл конфигурации не найден");
                form.SetStatus("⚠️ Конфиг не найден");
                return;
            }

            try
            {
                form.SetStatus("🔄 Инициализация...");
                Logger.SetOutputFolder(folderPath, "export_log.log");
                Logger.Clear();
                Logger.Info("BatchIFCExportCommand", "▶ Запуск команды");

                UIApplication uiapp = commandData.Application;
                Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;

                ISTimer.Start();
                Logger.Debug("BatchIFCExportCommand", $"Revit Version: {app.VersionNumber}");

                // Поиск файлов .rvt
                string[] rvtFiles = Directory.GetFiles(folderPath, "*.rvt");
                Logger.Info("BatchIFCExportCommand", $"Найдено файлов .rvt: {rvtFiles.Length}");
                foreach (var f in rvtFiles)
                    Logger.Debug("BatchIFCExportCommand", $"  • {Path.GetFileName(f)}");

                if (rvtFiles.Length == 0)
                {
                    TaskDialog.Show("Внимание", "В папке не найдено файлов .rvt");
                    form.SetStatus("⚠️ Нет файлов .rvt");
                    return;
                }

                // Загрузка конфигурации
                var excelConfigs = string.IsNullOrEmpty(excelConfigPath)
                    ? new List<IfcModelConfig>()
                    : ExcelIfcConfigLoader.Load(excelConfigPath);

                string defaultView = "Navisworks";

                form.SetStatus($"🚀 Экспорт {rvtFiles.Length} файлов...");

                // 🔹 Запуск основного процесса
                BatchExportIFC(app, new List<string>(rvtFiles), defaultView, excelConfigs, folderPath);

                var time = ISTimer.Stop();
                Logger.Info("BatchIFCExportCommand", $"✅ BatchExportIFC завершён | {time}");

                IsDebugWindow.AddRow(time);
                IsDebugWindow.Show();

                form.SetStatus($"✅ Готово | {time}");
                TaskDialog.Show("Экспорт завершён", $"Обработано файлов: {rvtFiles.Length}\nВремя: {time}");
            }
            catch (Exception ex)
            {
                form.SetStatus("❌ Ошибка экспорта");
                Logger.Critical("BatchIFCExportCommand", $"❌ Ошибка в RunExport: {ex}\n{ex.StackTrace}");
                IsDebugWindow.AddRow($"ERROR: {ex.Message}");
                IsDebugWindow.Show();
                TaskDialog.Show("Ошибка", $"Не удалось выполнить экспорт:\n{ex.Message}");
            }
        }

        /// <summary>
        /// Основной метод пакетного экспорта.
        /// </summary>
        private void BatchExportIFC(
            Autodesk.Revit.ApplicationServices.Application app,
            List<string> rvtFiles,
            string defaultView,
            List<IfcModelConfig> excelConfigs,
            string outputFolder)
        {
            Logger.Debug("BatchIFCExportCommand", $"[Batch] Начало обработки {rvtFiles.Count} файлов");
            int success = 0, fail = 0;

            foreach (string rvtPath in rvtFiles)
            {
                string fileName = Path.GetFileName(rvtPath);
                Logger.Info("BatchIFCExportCommand", $"[{success + fail + 1}/{rvtFiles.Count}] Обработка: {fileName}");

                if (!File.Exists(rvtPath))
                {
                    Logger.Warning("BatchIFCExportCommand", $"Файл не найден: {fileName}");
                    IsDebugWindow.AddRow($"❌ Не найден: {fileName}");
                    fail++;
                    continue;
                }

                try
                {
                    // 🔍 Резолвим конфигурацию для текущей модели
                    var modelConfig = ExcelIfcConfigLoader.Resolve(fileName, excelConfigs);

                    var rvtDoc = new RvtDocument(app, rvtPath);
                    string excludePattern = modelConfig.WorksetExcludePattern ?? "Связь";

                    Logger.Debug("BatchIFCExportCommand",
                        $"Открытие: {fileName} | View: {modelConfig.ViewName ?? defaultView} | Исключение: '{excludePattern}'");

                    var doc = rvtDoc.Open(excludePattern);
                    if (doc == null)
                    {
                        fail++;
                        continue;
                    }

                    // Создаём конфиг с разрешёнными параметрами
                    var ifcCfg = new IfcExportConfig(
                        doc,
                        modelConfig.ViewName ?? defaultView,
                        modelConfig.JsonConfigPath,
                        modelConfig.MappingFilePath // Переопределение мэппинга
                    );

                    var exportOptions = ifcCfg.GetConfig();
                    if (exportOptions == null)
                    {
                        doc.Close(false);
                        fail++;
                        continue;
                    }

                    // Путь для экспорта — в ту же папку или в подпапку Output
                    string ifcOutputFolder = Path.Combine(outputFolder, "IFC_Output");
                    if (!Directory.Exists(ifcOutputFolder))
                        Directory.CreateDirectory(ifcOutputFolder);

                    string ifcFileName = Path.ChangeExtension(fileName, ".ifc");
                    string ifcPath = Path.Combine(ifcOutputFolder, ifcFileName);

                    using (Transaction tx = new Transaction(doc, "BatchIFCExport"))
                    {
                        tx.Start();
                        doc.Export(Path.GetDirectoryName(ifcPath), Path.GetFileName(ifcPath), exportOptions);
                        tx.Commit();
                    }

                    if (File.Exists(ifcPath))
                    {
                        long sizeKb = new FileInfo(ifcPath).Length / 1024;
                        Logger.Info("BatchIFCExportCommand", $"✅ Экспорт: {fileName} ({sizeKb} KB)");
                        IsDebugWindow.AddRow($"✅ {fileName}");
                        success++;
                    }
                    else
                    {
                        Logger.Warning("BatchIFCExportCommand", $"Файл не создан: {fileName}");
                        IsDebugWindow.AddRow($"⚠️ Пусто: {fileName}");
                        fail++;
                    }

                    doc.Close(false);
                }
                catch (Exception ex)
                {
                    Logger.Error("BatchIFCExportCommand", $"❌ Ошибка {fileName}: {ex.Message}");
                    Logger.Debug("BatchIFCExportCommand", $"StackTrace: {ex.StackTrace}");
                    IsDebugWindow.AddRow($"💥 {fileName}: {ex.Message}");
                    fail++;
                }
            }

            Logger.Info("BatchIFCExportCommand", $"[Batch] Итог: Успешно {success}, Ошибок {fail}");
        }
    }
}