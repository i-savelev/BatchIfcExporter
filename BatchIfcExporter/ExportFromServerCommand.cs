using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using BatchIfcExporter;
using ISTools;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Form = System.Windows.Forms.Form;

namespace BatchExportIfc
{
    [Transaction(TransactionMode.Manual)]
    public class ExportFromServerCommand : IExternalCommand
    {
        public static string IS_TAB_NAME => "ISTools";
        public static string IS_NAME => "Экспорт IFC с Revit сервера";
        public static string IS_IMAGE => "BatchIfcExporter.Resources.ifc_to_rs.png";
        public static string IS_DESCRIPTION => "Автор: https://github.com/i-savelev\r\nПакетный экспорт выбранных моделей с Revit Server в IFC с конфигурацией Excel";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                Logger.Clear();
                Logger.Info("=== Запуск ExportFromServerCommand ===");

                var apiYear = int.Parse(commandData.Application.Application.VersionNumber);
                var servers = RevitServerBrowser.RevitServerConfigReader.ReadServers(apiYear);

                var form = new RevitServerBrowser.RevitServerBrowserForm(servers, apiYear);

                // 🔹 Подписываемся на кнопку "Подтвердить"
                form.ConfirmButton.Click += (s, e) =>
                {
                    var selectedPaths = form.SelectedModelPaths;

                    if (!selectedPaths.Any())
                    {
                        MessageBox.Show(form, "Выберите хотя бы одну модель для экспорта", "Внимание",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    // 🔹 Диалог 1: Выбор папки (через OpenFileDialog)
                    string outputFolder = null;
                    using (var dlgFolder = new OpenFileDialog
                    {
                        Title = "📁 Выберите папку для сохранения IFC-файлов",
                        Filter = "Папки|*",
                        CheckFileExists = false,
                        CheckPathExists = true,
                        ValidateNames = false,
                        FileName = "Выберите папку"
                    })
                    {
                        if (dlgFolder.ShowDialog(form) == DialogResult.OK)
                        {
                            outputFolder = Path.GetDirectoryName(dlgFolder.FileName);
                            if (string.IsNullOrEmpty(outputFolder) && Directory.Exists(dlgFolder.FileName))
                                outputFolder = dlgFolder.FileName;
                        }
                    }

                    if (string.IsNullOrEmpty(outputFolder) || !Directory.Exists(outputFolder))
                    {
                        MessageBox.Show(form, "Выберите корректную папку для экспорта", "Внимание",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    Logger.SetOutputFolder(outputFolder, "export_log.log");
                    Logger.Clear();
                    Logger.Info("ExportFromServerCommand", "▶ Запуск команды");
                    Logger.Info("ExportFromServerCommand", $"Список моделей:");
                    int idx = 1;
                    foreach (var path in selectedPaths)
                    {
                        Logger.Info("ExportFromServerCommand", $"[{idx}] {path}");
                        idx++;
                    }

                    // 🔹 Диалог 2: Выбор Excel-конфигурации (опционально)
                    string configPath = null;
                    var dlgResult = MessageBox.Show(form,
                        "Использовать конфигурацию из Excel?\n\n• ViewName\n• JsonConfigPath\n• MappingFilePath\n• WorksetExcludePattern\n\n«Нет» — экспорт с настройками по умолчанию.",
                        "Конфигурация экспорта",
                        MessageBoxButtons.YesNoCancel,
                        MessageBoxIcon.Question);

                    if (dlgResult == DialogResult.Cancel) return;

                    if (dlgResult == DialogResult.Yes)
                    {
                        using (var dlgConfig = new OpenFileDialog
                        {
                            Title = "📋 Выберите файл конфигурации Excel",
                            Filter = "Excel Files|*.xlsx;*.xls|All Files|*.*",
                            CheckFileExists = true
                        })
                        {
                            if (dlgConfig.ShowDialog(form) == DialogResult.OK)
                                configPath = dlgConfig.FileName;
                            else
                                return;
                        }
                    }

                    // 🔹 Запуск экспорта
                    RunExport(commandData, selectedPaths, outputFolder, configPath, form);
                };

                form.ShowDialog();
                Logger.Info("=== Завершение команды ===");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                Logger.Critical($"[CMD] Ошибка: {ex}");
                IsDebugWindow.AddRow($"ERROR: {ex.Message}");
                IsDebugWindow.Show();
                return Result.Failed;
            }
        }

        /// <summary>
        /// Синхронный экспорт моделей с использованием RvtDocument.
        /// </summary>
        private void RunExport(
            ExternalCommandData commandData,
            IReadOnlyList<string> rsnPaths,
            string outputFolder,
            string excelConfigPath,
            Form parentForm)
        {
            try
            {
                UpdateStatus(parentForm, "🔄 Инициализация...");
                Logger.Info($"[EXPORT] Начало: {rsnPaths.Count} моделей");

                var app = commandData.Application.Application;
                int success = 0, fail = 0;

                // Загрузка конфигурации
                var excelConfigs = string.IsNullOrEmpty(excelConfigPath)
                    ? new List<IfcModelConfig>()
                    : ExcelIfcConfigLoader.Load(excelConfigPath);

                string defaultView = "Navisworks";

                if (!Directory.Exists(outputFolder))
                    Directory.CreateDirectory(outputFolder);

                // 🛡️ Инициализация защитника мэппинга (для первой модели с валидным mapping)
                IfcMappingGuard mappingGuard = null;
                string firstModelWithMapping = null;
                bool firstExportProcessed = false;

                foreach (var rsnPath in rsnPaths)
                {
                    var fileName = Path.GetFileName(rsnPath);
                    UpdateStatus(parentForm, $"[{success + fail + 1}/{rsnPaths.Count}] {fileName}");
                    Logger.Info($"[{success + fail + 1}/{rsnPaths.Count}] Обработка: {fileName}");

                    try
                    {
                        var rvtDoc = new RvtDocument(app, rsnPath);

                        // Получаем паттерн исключения из конфига или используем дефолт
                        var modelConfig = ExcelIfcConfigLoader.Resolve(fileName, excelConfigs);
                        string excludePattern = modelConfig.WorksetExcludePattern ?? "Связь";

                        // Открываем документ (RvtDocument.Open() уже делает всю магию с ворксетами)
                        var doc = rvtDoc.Open(excludePattern);
                        if (doc == null)
                            throw new Exception("Не удалось открыть документ");

                        // 🔹 Создаём настройки экспорта
                        var ifcCfg = new IfcExportConfig(
                            doc,
                            modelConfig.ViewName ?? defaultView,
                            modelConfig.JsonConfigPath,
                            modelConfig.MappingFilePath
                        );

                        var exportOptions = ifcCfg.GetConfig();
                        if (exportOptions == null)
                            throw new Exception("Не удалось получить настройки экспорта");

                        var ifcFileName = Path.ChangeExtension(fileName, ".ifc");
                        var ifcPath = Path.Combine(outputFolder, ifcFileName);

                        // 🛡️ Инициализируем защитник для ПЕРВОГО файла с валидным mapping
                        if (!firstExportProcessed &&
                            !string.IsNullOrEmpty(modelConfig.MappingFilePath) &&
                            File.Exists(modelConfig.MappingFilePath) &&
                            mappingGuard == null)
                        {
                            try
                            {
                                mappingGuard = new IfcMappingGuard(modelConfig.MappingFilePath);
                                firstModelWithMapping = rsnPath;
                                Logger.Info("ExportFromServerCommand",
                                    $"🛡️ MappingGuard активирован для: {fileName}");
                            }
                            catch (Exception ex)
                            {
                                Logger.Warning("ExportFromServerCommand",
                                    $"⚠️ Не удалось инициализировать MappingGuard: {ex.Message}");
                            }
                        }

                        // 🔹 Экспорт
                        using (var tx = new Transaction(doc, "ExportIFC"))
                        {
                            tx.Start();
                            doc.Export(outputFolder, ifcFileName, exportOptions);
                            tx.Commit();
                        }

                        // 🛡️ Проверка mapping-файла после ПЕРВОГО экспорта
                        bool needsRetry = false;
                        if (mappingGuard != null && !firstExportProcessed)
                        {
                            firstExportProcessed = true;
                            if (mappingGuard.VerifyAndRestore())
                            {
                                Logger.Info("ExportFromServerCommand",
                                    $"🔄 Mapping изменён — повторный экспорт первой модели: {fileName}");
                                needsRetry = true;
                            }
                        }

                        // 🔁 Защита от бесконечного retry
                        int retryCount = 0;
                        const int MAX_RETRIES = 2;

                        if (needsRetry)
                        {
                            if (retryCount >= MAX_RETRIES)
                            {
                                Logger.Error("ExportFromServerCommand",
                                    $"❌ Превышено число попыток экспорта для {fileName} (max={MAX_RETRIES})");
                                IsDebugWindow.AddRow($"💥 {fileName}: retry limit exceeded");
                                fail++;
                                doc?.Close(false);
                                continue;
                            }

                            retryCount++;
                            Logger.Debug("ExportFromServerCommand",
                                $"🔄 Попытка #{retryCount}/{MAX_RETRIES} для {fileName}");

                            // Закрываем текущий документ перед повторным открытием
                            doc.Close(false);

                            // Переоткрываем документ для повторного экспорта
                            var rvtDocRetry = new RvtDocument(app, rsnPath);
                            var docRetry = rvtDocRetry.Open(excludePattern);
                            if (docRetry == null)
                                throw new Exception("Не удалось переоткрыть документ для retry");

                            // Получаем настройки экспорта заново (mapping уже восстановлен)
                            var ifcCfgRetry = new IfcExportConfig(
                                docRetry,
                                modelConfig.ViewName ?? defaultView,
                                modelConfig.JsonConfigPath,
                                modelConfig.MappingFilePath
                            );
                            var exportOptionsRetry = ifcCfgRetry.GetConfig();
                            if (exportOptionsRetry == null)
                                throw new Exception("Не удалось получить настройки экспорта (retry)");

                            // Повторный экспорт
                            using (var tx = new Transaction(docRetry, "ExportIFC_Retry"))
                            {
                                tx.Start();
                                docRetry.Export(outputFolder, ifcFileName, exportOptionsRetry);
                                tx.Commit();
                            }

                            // Проверяем результат повторного экспорта
                            if (File.Exists(ifcPath))
                            {
                                long sizeKb = new FileInfo(ifcPath).Length / 1024;
                                Logger.Info($"[EXPORT] ✅ RETRY: {fileName} → {ifcFileName} ({sizeKb} KB)");
                                IsDebugWindow.AddRow($"✅ {fileName} (retry)");
                                success++;
                            }
                            else
                            {
                                Logger.Warning($"[EXPORT] ⚠️ Файл не создан после retry: {fileName}");
                                IsDebugWindow.AddRow($"⚠️ Пусто (retry): {fileName}");
                                fail++;
                            }

                            docRetry.Close(false);
                        }
                        else
                        {
                            // Стандартная обработка успеха
                            if (File.Exists(ifcPath))
                            {
                                long sizeKb = new FileInfo(ifcPath).Length / 1024;
                                Logger.Info($"[EXPORT] ✅ {fileName} → {ifcFileName} ({sizeKb} KB)");
                                IsDebugWindow.AddRow($"✅ {fileName}");
                                success++;
                            }
                            else
                            {
                                Logger.Warning($"[EXPORT] ⚠️ Файл не создан: {fileName}");
                                IsDebugWindow.AddRow($"⚠️ Пусто: {fileName}");
                                fail++;
                            }

                            // 🔹 Закрываем документ (без сохранения)
                            doc.Close(false);
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Error($"[EXPORT] ❌ Ошибка {fileName}: {ex.Message}");
                        Logger.Debug($"[EXPORT] Stack: {ex.StackTrace}");
                        IsDebugWindow.AddRow($"💥 {fileName}: {ex.Message}");
                        fail++;
                    }
                }

                // 🧹 Очистка защитника
                mappingGuard?.Dispose();

                UpdateStatus(parentForm, $"✅ Готово: {success} ✅ | {fail} ❌");
                IsDebugWindow.AddRow($"📊 Итог: {success} ✅ | {fail} ❌");
                IsDebugWindow.AddRow($"Лог RevitServerBrowser: {RevitServerBrowser.Logger.GetPath()}");

                IsDebugWindow.Show();

                TaskDialog.Show("Экспорт завершён",
                    $"Обработано: {rsnPaths.Count}\n✅ Успешно: {success}\n❌ Ошибок: {fail}");
            }
            catch (Exception ex)
            {
                Logger.Critical($"[EXPORT] Критическая ошибка: {ex}");
                UpdateStatus(parentForm, "❌ Ошибка экспорта");
                TaskDialog.Show("Ошибка", $"Не удалось выполнить экспорт:\n{ex.Message}");
            }
        }

        /// <summary>
        /// Обновление статуса через заголовок окна (без рефлексии).
        /// </summary>
        private void UpdateStatus(Form form, string text)
        {
            if (form?.IsHandleCreated == true)
            {
                try
                {
                    form.Invoke((MethodInvoker)(() => form.Text = $"Export: {text}"));
                }
                catch { /* Игнорируем ошибки UI */ }
            }
            Logger.Debug($"[STATUS] {text}");
        }
    }
}