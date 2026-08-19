using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using BatchIfcExporter;
using DebugWindow = RevitLogger.DebugWindow;
using Logger = RevitLogger.Logger;
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
        public static string IS_DESCRIPTION => "Автор: https://github.com/i-savelev\r\nРепозиторий: https://github.com/i-savelev/BatchIfcExporter\r\nПакетный экспорт выбранных моделей с Revit Server в IFC с конфигурацией Excel";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var form = new RevitServerBrowser.RevitServerBrowserNativeForm(commandData);

                form.ConfirmButton.Click += (s, e) =>
                {
                    var selectedPaths = form.SelectedModelPaths;

                    if (!selectedPaths.Any())
                    {
                        MessageBox.Show(form, "Выберите хотя бы одну модель для экспорта", "Внимание",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

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
                    Logger.SetLogPath(Path.Combine(outputFolder, "export_log.log"));
                    Logger.Init(hostName: "Autodesk Revit",
                        hostVersionNumber: commandData.Application.Application.VersionNumber,
                        hostBuild: commandData.Application.Application.VersionBuild,
                        hasActiveDocument: commandData.Application.ActiveUIDocument != null);
                    Logger.Info("[ExportFromServerCommand] ▶ Запуск команды");
                    Logger.Info("[ExportFromServerCommand] Список моделей:");
                    int idx = 1;
                    foreach (var path in selectedPaths)
                    {
                        Logger.Info($"[ExportFromServerCommand] [{idx}] {path}");
                        idx++;
                    }

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
                DebugWindow.AddRow($"ERROR: {ex.Message}");
                DebugWindow.Show();
                return Result.Failed;
            }
        }

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

                var excelConfigs = string.IsNullOrEmpty(excelConfigPath)
                    ? new List<IfcModelConfig>()
                    : ExcelIfcConfigLoader.Load(excelConfigPath);

                string defaultView = "Navisworks";

                if (!Directory.Exists(outputFolder))
                    Directory.CreateDirectory(outputFolder);

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
                        var modelConfigs = ExcelIfcConfigLoader.ResolveAll(fileName, excelConfigs);
                        var firstConfig = modelConfigs.First();
                        string excludePattern = firstConfig.WorksetExcludePattern ?? "Связь";

                        var rvtDoc = new RvtDocument(app, rsnPath);
                        var doc = rvtDoc.Open(excludePattern);
                        if (doc == null)
                            throw new Exception("Не удалось открыть документ");

                        if (!firstExportProcessed &&
                            !string.IsNullOrEmpty(firstConfig.MappingFilePath) &&
                            File.Exists(firstConfig.MappingFilePath) &&
                            mappingGuard == null)
                        {
                            try
                            {
                                mappingGuard = new IfcMappingGuard(firstConfig.MappingFilePath);
                                firstModelWithMapping = rsnPath;
                                Logger.Info($"[ExportFromServerCommand] 🛡️ MappingGuard активирован для: {fileName}");
                            }
                            catch (Exception ex)
                            {
                                Logger.Warning($"[ExportFromServerCommand] ⚠️ Не удалось инициализировать MappingGuard: {ex.Message}");
                            }
                        }

                        bool needsRetry = false;
                        int modelSuccess = 0;
                        int modelFail = 0;

                        void ExportAllViews(Document currentDoc)
                        {
                            modelSuccess = 0;
                            modelFail = 0;
                            foreach (var modelConfig in modelConfigs)
                            {
                                var ifcCfg = new IfcExportConfig(currentDoc, modelConfig.ViewName ?? defaultView, modelConfig.JsonConfigPath, modelConfig.MappingFilePath);
                                var exportOptions = ifcCfg.GetConfig();
                                if (exportOptions == null) continue;

                                string suffix = !string.IsNullOrEmpty(modelConfig.FileSuffix)
                                    ? modelConfig.FileSuffix
                                    : (modelConfigs.Count > 1 ? $"_{modelConfig.ViewName}" : "");

                                string ifcFileName = Path.GetFileNameWithoutExtension(fileName) + suffix + ".ifc";
                                var ifcPath = Path.Combine(outputFolder, ifcFileName);

                                using (var tx = new Transaction(currentDoc, $"ExportIFC_{modelConfig.ViewName}"))
                                {
                                    tx.Start();
                                    currentDoc.Export(outputFolder, ifcFileName, exportOptions);
                                    tx.Commit();
                                }

                                if (File.Exists(ifcPath))
                                {
                                    long sizeKb = new FileInfo(ifcPath).Length / 1024;
                                    Logger.Info($"[EXPORT] ✅ {ifcFileName} ({sizeKb} KB)");
                                    DebugWindow.AddRow($"✅ {ifcFileName}");
                                    modelSuccess++;
                                    success++;
                                }
                                else
                                {
                                    Logger.Warning($"[EXPORT] ⚠️ Файл не создан: {ifcFileName}");
                                    DebugWindow.AddRow($"⚠️ Пусто: {ifcFileName}");
                                    modelFail++;
                                    fail++;
                                }
                            }
                        }

                        ExportAllViews(doc);

                        if (mappingGuard != null && !firstExportProcessed)
                        {
                            firstExportProcessed = true;
                            if (mappingGuard.VerifyAndRestore())
                            {
                                Logger.Info($"[ExportFromServerCommand] 🔄 Mapping изменён — повторный экспорт первой модели: {fileName}");
                                needsRetry = true;
                            }
                        }

                        int retryCount = 0;
                        const int MAX_RETRIES = 2;

                        if (needsRetry)
                        {
                            if (retryCount >= MAX_RETRIES)
                            {
                                Logger.Error($"[ExportFromServerCommand] ❌ Превышено число попыток экспорта для {fileName} (max={MAX_RETRIES})");
                                DebugWindow.AddRow($"💥 {fileName}: retry limit exceeded");
                                SafeCloseDocument(doc, fileName, app);
                                continue;
                            }

                            retryCount++;
                            Logger.Debug($"[ExportFromServerCommand] 🔄 Попытка #{retryCount}/{MAX_RETRIES} для {fileName}");

                            success -= modelSuccess;
                            fail -= modelFail;

                            SafeCloseDocument(doc, fileName, app);

                            var rvtDocRetry = new RvtDocument(app, rsnPath);
                            var docRetry = rvtDocRetry.Open(excludePattern);
                            if (docRetry == null)
                                throw new Exception("Не удалось переоткрыть документ для retry");

                            ExportAllViews(docRetry);
                            SafeCloseDocument(docRetry, fileName, app);
                        }
                        else
                        {
                            SafeCloseDocument(doc, fileName, app);
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Error($"[EXPORT] ❌ Ошибка {fileName}: {ex.Message}");
                        Logger.Debug($"[EXPORT] Stack: {ex.StackTrace}");
                        DebugWindow.AddRow($"💥 {fileName}: {ex.Message}");
                        fail++;
                    }
                }

                mappingGuard?.Dispose();

                UpdateStatus(parentForm, $"✅ Готово: {success} ✅ | {fail} ❌");
                DebugWindow.AddRow($"📊 Итог: {success} ✅ | {fail} ❌");

                DebugWindow.Show();

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

        private void SafeCloseDocument(Document doc, string fileName, Autodesk.Revit.ApplicationServices.Application app)
        {
            try
            {
                if (doc == null || !doc.IsValidObject)
                {
                    Logger.Debug($"[ExportFromServerCommand] 🔒 Пропуск закрытия {fileName}: doc=null или невалиден");
                    return;
                }

                Logger.Debug($"[ExportFromServerCommand] 🔒 Закрытие документа: {fileName} | IsLinked={doc.IsLinked}");

                if (doc.IsLinked)
                {
                    Logger.Debug($"[ExportFromServerCommand] ⏭ Пропуск Close() для linked: {fileName}");
                    return;
                }

                doc.Close(false);
                Logger.Debug($"[ExportFromServerCommand] ✅ Документ {fileName} закрыт");

                System.GC.Collect();
                System.GC.WaitForPendingFinalizers();
                Logger.Debug($"[ExportFromServerCommand] 🧹 GC выполнен после закрытия {fileName}");
            }
            catch (Autodesk.Revit.Exceptions.InvalidOperationException ex)
                when (ex.Message.IndexOf("linked", StringComparison.OrdinalIgnoreCase) >= 0 ||
                      ex.Message.IndexOf("Cannot close", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Logger.Warning($"[ExportFromServerCommand] ⚠️ Нельзя закрыть linked-файл {fileName}: {ex.Message}");
            }
            catch (Exception ex)
            {
                Logger.Warning($"[ExportFromServerCommand] ⚠️ Ошибка при закрытии {fileName}: {ex.Message}");
            }
        }

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