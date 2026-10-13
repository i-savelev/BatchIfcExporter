using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using Logger = RevitLogger.Logger;

namespace BatchExportIfc
{
    [Transaction(TransactionMode.Manual)]
    public class TestCloseDocumentsCommand : IExternalCommand
    {
        public static string IS_TAB_NAME => "ISTools";
        public static string IS_NAME => "🧪 Тест: Закрыть документы";
        public static string IS_IMAGE => "BatchIfcExporter.Resources.ifc_to_folder.png";
        public static string IS_DESCRIPTION => "Тестовая команда для проверки закрытия открытых документов в сессии Revit и очистки памяти";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            ConfigureLogging(commandData);

            try
            {
                Logger.Info("[TestCloseDocumentsCommand] ▶ Старт команды проверки и закрытия документов");

                var app = commandData.Application.Application;
                var activeDoc = commandData.Application.ActiveUIDocument?.Document;

                Logger.Info($"[TestCloseDocumentsCommand] Активный документ: {activeDoc?.Title ?? "null"}");

                // Собираем список документов для безопасного перебора (избегает изменения коллекции во время итерации)
                var docsToProcess = new List<Document>();
                foreach (Document doc in app.Documents)
                {
                    if (doc.IsValidObject)
                    {
                        docsToProcess.Add(doc);
                    }
                }

                Logger.Info($"[TestCloseDocumentsCommand] Всего валидных документов в сессии: {docsToProcess.Count}");

                int closedCount = 0;
                int skippedCount = 0;
                int errorCount = 0;

                foreach (var doc in docsToProcess)
                {
                    try
                    {
                        string docTitle = doc.Title;
                        bool isLinked = doc.IsLinked;
                        bool isModified = doc.IsModified;
                        string path = doc.PathName;

                        Logger.Debug($"[TestCloseDocumentsCommand] Анализ: '{docTitle}' | Linked={isLinked} | Modified={isModified} | Path={path}");

                        // 1. Пропускаем активный документ, чтобы не мешать пользователю
                        if (doc == activeDoc)
                        {
                            Logger.Info($"[TestCloseDocumentsCommand] ⏭ Пропуск: '{docTitle}' является активным документом.");
                            skippedCount++;
                            continue;
                        }

                        // 2. Пропускаем связанные модели (их нельзя закрыть через doc.Close())
                        if (isLinked)
                        {
                            Logger.Warning($"[TestCloseDocumentsCommand] ⏭ Пропуск: '{docTitle}' является связанной моделью (IsLinked=True).");
                            skippedCount++;
                            continue;
                        }

                        // 3. Предупреждение о несохраненных изменениях
                        if (isModified)
                        {
                            Logger.Warning($"[TestCloseDocumentsCommand] ⚠️ Внимание: '{docTitle}' имеет несохраненные изменения. Будет закрыт без сохранения.");
                        }

                        // 4. Попытка закрытия
                        Logger.Info($"[TestCloseDocumentsCommand] 🔄 Попытка закрытия: '{docTitle}'");
                        doc.Close(false);
                        Logger.Info($"[TestCloseDocumentsCommand] ✅ Успешно закрыт: '{docTitle}'");
                        closedCount++;
                    }
                    catch (Exception ex)
                    {
                        Logger.Exception(ex, $"[TestCloseDocumentsCommand] ❌ Ошибка при закрытии документа '{doc.Title}'");
                        errorCount++;
                    }
                }

                // 5. Очистка памяти
                Logger.Info("[TestCloseDocumentsCommand] 🧹 Запуск сборщика мусора (GC.Collect)");
                GC.Collect();
                GC.WaitForPendingFinalizers();

                Logger.Info($"[TestCloseDocumentsCommand] 🏁 Итог операции: Закрыто={closedCount}, Пропущено={skippedCount}, Ошибок={errorCount}");

                // 6. Финальная проверка остатка
                int remainingCount = 0;
                foreach (Document doc in app.Documents)
                {
                    if (doc.IsValidObject) remainingCount++;
                }
                Logger.Info($"[TestCloseDocumentsCommand] Осталось открытых документов в сессии после очистки: {remainingCount}");

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                Logger.Exception(ex, "[TestCloseDocumentsCommand] Критическая ошибка выполнения команды");
                message = ex.Message;
                return Result.Failed;
            }
        }

        private void ConfigureLogging(ExternalCommandData commandData)
        {
            string logPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Temp",
                "BatchIfcExporter",
                "TestCloseDocuments.log");

            Logger.SetLogPath(logPath);
            Logger.SetLogLevel(Logger.LogLevel.Debug);
            Logger.Init(
                hostName: "Autodesk Revit",
                hostVersionNumber: commandData.Application.Application.VersionNumber,
                hostBuild: commandData.Application.Application.VersionBuild,
                hasActiveDocument: commandData.Application.ActiveUIDocument != null);
        }
    }
}