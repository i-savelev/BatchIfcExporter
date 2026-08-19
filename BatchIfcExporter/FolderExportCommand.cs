using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using BatchIfcExporter;
using OfficeOpenXml.Style;
using Logger = RevitLogger.Logger;
using DebugWindow = RevitLogger.DebugWindow;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace BatchExportIfc
{
    [Autodesk.Revit.Attributes.TransactionAttribute(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class FolderExportCommand : IExternalCommand
    {
        public static string IS_TAB_NAME => "ISTools";
        public static string IS_NAME => "Экспорт IFC из папки";
        public static string IS_IMAGE => "BatchIfcExporter.Resources.ifc_to_folder.png";
        public static string IS_DESCRIPTION => "Автор: https://github.com/i-savelev\r\nРепозиторий: https://github.com/i-savelev/BatchIfcExporter\r\nПакетный экспорт IFC моделей из выбранной папки в с конфигурацией из Excel";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                Logger.Clear();
                Logger.Info("[BatchIFCExportCommand] ▶ Открытие окна экспорта моделей из папки");
                Logger.Init(hostName: "Autodesk Revit",
                    hostVersionNumber: commandData.Application.Application.VersionNumber,
                    hostBuild: commandData.Application.Application.VersionBuild,
                    hasActiveDocument: commandData.Application.ActiveUIDocument != null);

                var form = new ExportConfigForm();
                form.Text = "Пакетный экспорт IFC";

                form.BtnSelectIfcFolder.Click += (s, e) =>
                {
                    using (var dlg = new OpenFileDialog
                    {
                        Title = "Выберите папку с моделями Revit",
                        Filter = "Папки|*",
                        CheckFileExists = false,
                        CheckPathExists = true,
                        ValidateNames = false,
                        FileName = "Выберите папку"
                    })
                    {
                        if (dlg.ShowDialog(form) == DialogResult.OK)
                        {
                            string folderPath = Path.GetDirectoryName(dlg.FileName);
                            if (string.IsNullOrEmpty(folderPath) && Directory.Exists(dlg.FileName))
                                folderPath = dlg.FileName;

                            if (!string.IsNullOrEmpty(folderPath) && Directory.Exists(folderPath))
                            {
                                form.TxtIfcFolder.Text = folderPath;
                                form.SetStatus($"Папка: {Path.GetFileName(folderPath)}");
                                Logger.Info($"[BatchIFCExportCommand] Выбрана папка: {folderPath}");
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
                            Logger.Info($"[BatchIFCExportCommand] Выбран конфиг: {dlg.FileName}");
                        }
                    }
                };

                form.BtnSaveTemplate.Click += (s, e) => SaveConfigTemplate(form);
                form.BtnRunExport.Click += (s, e) => RunExport(commandData, form);
                form.BtnClose.Click += (s, e) => form.Close();

                form.ShowDialog();
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                Logger.Critical($"[BatchIFCExportCommand] ❌ Критическая ошибка: {ex}\n{ex.StackTrace}");
                DebugWindow.AddRow($"ERROR: {ex.Message}");
                DebugWindow.Show();
                return Result.Failed;
            }
        }

        private void CreateTemplateFile(string path)
        {
            if (!path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                path += ".xlsx";

            using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(path)))
            {
                var ws = package.Workbook.Worksheets.Add("Конфигурация");

                // Добавлена колонка FileSuffix
                string[] headers = { "FileName", "JsonConfigPath", "ViewName", "MappingFilePath", "WorksetExcludePattern", "FileSuffix" };
                for (int i = 0; i < headers.Length; i++)
                    ws.Cells[1, i + 1].Value = headers[i];

                var headerRange = ws.Cells["A1:F1"];
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                headerRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(242, 242, 242));
                headerRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                var example = new object[] { "Project_A.rvt", "", "Navisworks", "", "Связь*", "_AR" };
                for (int i = 0; i < example.Length; i++)
                    ws.Cells[2, i + 1].Value = example[i];

                ws.Cells["A1:F2"].AutoFitColumns();
                package.Save();
            }
        }

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
                        Logger.Info($"[BatchIFCExportCommand] Шаблон сохранён: {dlg.FileName}");
                        TaskDialog.Show("Успех", "Шаблон конфигурации успешно создан!");
                    }
                }
            }
            catch (Exception ex)
            {
                form.SetStatus("❌ Ошибка сохранения");
                Logger.Error($"[BatchIFCExportCommand] Ошибка сохранения шаблона: {ex.Message}");
                TaskDialog.Show("Ошибка", $"Не удалось сохранить шаблон:\n{ex.Message}");
            }
        }

        private void RunExport(ExternalCommandData commandData, ExportConfigForm form)
        {
            var folderPath = form.IfcFolderPath;
            var excelConfigPath = form.ConfigFilePath;

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
                Logger.SetLogPath(Path.Combine(folderPath, "export_log.log"));
                Logger.Clear();
                Logger.Info("[BatchIFCExportCommand] ▶ Запуск команды");

                UIApplication uiapp = commandData.Application;
                Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;

                ISTimer.Start();
                Logger.Debug($"[BatchIFCExportCommand] Revit Version: {app.VersionNumber}");

                string[] rvtFiles = Directory.GetFiles(folderPath, "*.rvt");
                Logger.Info($"[BatchIFCExportCommand] Найдено файлов .rvt: {rvtFiles.Length}");
                foreach (var f in rvtFiles)
                    Logger.Debug($"[BatchIFCExportCommand]   • {Path.GetFileName(f)}");

                if (rvtFiles.Length == 0)
                {
                    TaskDialog.Show("Внимание", "В папке не найдено файлов .rvt");
                    form.SetStatus("⚠️ Нет файлов .rvt");
                    return;
                }

                var excelConfigs = string.IsNullOrEmpty(excelConfigPath)
                    ? new List<IfcModelConfig>()
                    : ExcelIfcConfigLoader.Load(excelConfigPath);

                string defaultView = "Navisworks";
                form.SetStatus($"🚀 Экспорт {rvtFiles.Length} файлов...");

                BatchExportIFC(app, new List<string>(rvtFiles), defaultView, excelConfigs, folderPath);

                var time = ISTimer.Stop();
                Logger.Info($"[BatchIFCExportCommand] ✅ BatchExportIFC завершён | {time}");

                DebugWindow.AddRow(time);
                DebugWindow.Show();

                form.SetStatus($"✅ Готово | {time}");
                TaskDialog.Show("Экспорт завершён", $"Обработано файлов: {rvtFiles.Length}\nВремя: {time}");
            }
            catch (Exception ex)
            {
                form.SetStatus("❌ Ошибка экспорта");
                Logger.Critical($"[BatchIFCExportCommand] ❌ Ошибка в RunExport: {ex}\n{ex.StackTrace}");
                DebugWindow.AddRow($"ERROR: {ex.Message}");
                DebugWindow.Show();
                TaskDialog.Show("Ошибка", $"Не удалось выполнить экспорт:\n{ex.Message}");
            }
        }

        private void BatchExportIFC(
            Autodesk.Revit.ApplicationServices.Application app,
            List<string> rvtFiles,
            string defaultView,
            List<IfcModelConfig> excelConfigs,
            string outputFolder)
        {
            Logger.Debug($"[BatchIFCExportCommand] [Batch] Начало обработки {rvtFiles.Count} файлов");
            int success = 0, fail = 0;

            string ifcOutputFolder = Path.Combine(outputFolder, "IFC_Output");
            if (!Directory.Exists(ifcOutputFolder))
                Directory.CreateDirectory(ifcOutputFolder);

            foreach (string rvtPath in rvtFiles)
            {
                string fileName = Path.GetFileName(rvtPath);
                Logger.Info($"[BatchIFCExportCommand] [{success + fail + 1}/{rvtFiles.Count}] Обработка: {fileName}");

                if (!File.Exists(rvtPath))
                {
                    Logger.Warning($"[BatchIFCExportCommand] Файл не найден: {fileName}");
                    DebugWindow.AddRow($"❌ Не найден: {fileName}");
                    fail++;
                    continue;
                }

                try
                {
                    // Получаем ВСЕ конфигурации для данной модели
                    var modelConfigs = ExcelIfcConfigLoader.ResolveAll(fileName, excelConfigs);
                    var firstConfig = modelConfigs.First();

                    var rvtDoc = new RvtDocument(app, rvtPath);
                    string excludePattern = firstConfig.WorksetExcludePattern ?? "Связь";

                    Logger.Debug($"[BatchIFCExportCommand] Открытие: {fileName} | Исключение: '{excludePattern}' | Конфигов: {modelConfigs.Count}");

                    var doc = rvtDoc.Open(excludePattern);
                    if (doc == null)
                    {
                        fail++;
                        continue;
                    }

                    int modelSuccess = 0;

                    // Цикл по всем видам (конфигурациям) текущей модели
                    foreach (var modelConfig in modelConfigs)
                    {
                        try
                        {
                            var ifcCfg = new IfcExportConfig(
                                doc,
                                modelConfig.ViewName ?? defaultView,
                                modelConfig.JsonConfigPath,
                                modelConfig.MappingFilePath
                            );

                            var exportOptions = ifcCfg.GetConfig();
                            if (exportOptions == null)
                            {
                                Logger.Warning($"[BatchIFCExportCommand] ⚠️ Не удалось получить настройки для вида {modelConfig.ViewName}");
                                continue;
                            }

                            // Формирование уникального имени файла
                            string suffix = !string.IsNullOrEmpty(modelConfig.FileSuffix)
                                ? modelConfig.FileSuffix
                                : (modelConfigs.Count > 1 ? $"_{modelConfig.ViewName}" : "");

                            string ifcFileName = Path.GetFileNameWithoutExtension(fileName) + suffix + ".ifc";
                            string ifcPath = Path.Combine(ifcOutputFolder, ifcFileName);

                            using (Transaction tx = new Transaction(doc, $"BatchIFCExport_{modelConfig.ViewName}"))
                            {
                                tx.Start();
                                doc.Export(ifcOutputFolder, ifcFileName, exportOptions);
                                tx.Commit();
                            }

                            if (File.Exists(ifcPath))
                            {
                                long sizeKb = new FileInfo(ifcPath).Length / 1024;
                                Logger.Info($"[BatchIFCExportCommand] ✅ Экспорт: {ifcFileName} ({sizeKb} KB)");
                                DebugWindow.AddRow($"✅ {ifcFileName}");
                                modelSuccess++;
                                success++;
                            }
                            else
                            {
                                Logger.Warning($"[BatchIFCExportCommand] Файл не создан: {ifcFileName}");
                                DebugWindow.AddRow($"⚠️ Пусто: {ifcFileName}");
                                fail++;
                            }
                        }
                        catch (Exception ex)
                        {
                            Logger.Error($"[BatchIFCExportCommand] ❌ Ошибка экспорта вида {modelConfig.ViewName} для {fileName}: {ex.Message}");
                            DebugWindow.AddRow($"💥 {fileName} ({modelConfig.ViewName}): {ex.Message}");
                            fail++;
                        }
                    }

                    doc.Close(false);
                }
                catch (Exception ex)
                {
                    Logger.Error($"[BatchIFCExportCommand] ❌ Ошибка {fileName}: {ex.Message}");
                    Logger.Debug($"[BatchIFCExportCommand] StackTrace: {ex.StackTrace}");
                    DebugWindow.AddRow($"💥 {fileName}: {ex.Message}");
                    fail++;
                }
            }

            Logger.Info($"[BatchIFCExportCommand] [Batch] Итог: Успешных экспортов {success}, Ошибок {fail}");
        }
    }
}