using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using DebugWindow = RevitLogger.DebugWindow;
using Logger = RevitLogger.Logger;
using System;
using System.Collections.Generic;
using System.Linq;

namespace BatchExportIfc
{
    internal class RvtDocument
    {
        public Application App { get; set; }
        public string Path { get; set; }
        public Document Doc { get; set; }

        public RvtDocument(Application app, string path)
        {
            App = app;
            Path = path;
            Logger.Debug($"[RvtDocument] Инициализация | Путь: {path}");
        }

        public Document Open(string wsExcludeWord)
        {
            Logger.Info($"[RvtDocument] ▶ Открытие документа | Исключаем ворксеты с: '{wsExcludeWord}'");

            // 🔍 Диагностика сессии ПЕРЕД открытием
            LogSessionState("BEFORE_OPEN");

            ModelPath modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(Path);

            // 🔹 ШАГ 1: Быстрое открытие для анализа рабочих наборов
            Logger.Debug("[RvtDocument] [Шаг 1] Открытие для анализа ворксетов (CloseAllWorksets)");
            OpenOptions openOptionsStep1 = new OpenOptions();
            openOptionsStep1.SetOpenWorksetsConfiguration(new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets));
            openOptionsStep1.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets;

            Document docTemp = null;
            WorksetConfiguration finalWorksetConfig = null;
            List<string> closedWorksets = new List<string>();

            try
            {
                Logger.Debug("[RvtDocument] [Шаг 1] Вызов App.OpenDocumentFile()...");
                docTemp = App.OpenDocumentFile(modelPath, openOptionsStep1);

                if (docTemp != null)
                {
                    Logger.Debug($"[RvtDocument] [Шаг 1] ✅ Документ открыт: '{docTemp.Title}'");

                    // 🔥 Ключевая диагностика
                    Logger.Debug($"[RvtDocument] [DIAG] IsLinked={docTemp.IsLinked}, IsModified={docTemp.IsModified}, IsWorkshared={docTemp.IsWorkshared}");
                    Logger.Debug($"[RvtDocument] [DIAG] PathName={docTemp.PathName}, IsValid={docTemp.IsValidObject}");
                }

                string fileName = docTemp?.Title ?? "unknown";
                Logger.Debug($"[RvtDocument] [Шаг 1] Документ временно открыт: {fileName}");

                var allWorksets = new FilteredWorksetCollector(docTemp).ToWorksets().ToList();
                Logger.Info($"[RvtDocument] Всего ворксетов в документе: {allWorksets.Count}");

                var worksetIdsToOpen = allWorksets
                    .Where(w => w.Name.IndexOf(wsExcludeWord, StringComparison.OrdinalIgnoreCase) < 0)
                    .Select(w => w.Id)
                    .ToList();

                foreach (var ws in allWorksets)
                {
                    if (ws.Name.IndexOf(wsExcludeWord, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        closedWorksets.Add($"{ws.Name} (Id:{ws.Id})");
                        Logger.Debug($"[RvtDocument] 🔒 ЗАКРЫВАЕТСЯ: '{ws.Name}'");
                    }
                }

                Logger.Info($"[RvtDocument] [Worksets Summary] Закрывается: {closedWorksets.Count}");

                finalWorksetConfig = new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets);
                finalWorksetConfig.Open(worksetIdsToOpen);

                // 🔥 БЕЗОПАСНОЕ ЗАКРЫТИЕ (только рабочие API)
                SafeCloseDocument(docTemp, "[Шаг 1]");

                // 🔍 Диагностика после закрытия
                LogSessionState("AFTER_STEP1_CLOSE");
            }
            catch (Exception ex)
            {
                Logger.Error($"[RvtDocument] ❌ Ошибка на шаге 1: {ex.Message}");
                DebugWindow.AddRow($"Ошибка анализа ворксетов: {ex.Message}");

                try { SafeCloseDocument(docTemp, "[Шаг 1, catch]"); } catch { }

                // Fallback
                Logger.Warning("[RvtDocument] Fallback: открытие со всеми ворксетами");
                OpenOptions fallbackOptions = new OpenOptions();
                fallbackOptions.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets;
                Doc = App.OpenDocumentFile(modelPath, fallbackOptions);
                return Doc;
            }

            // 🔹 ШАГ 2: Финальное открытие
            Logger.Debug("[RvtDocument] [Шаг 2] Финальное открытие с фильтром ворксетов");
            OpenOptions openOptionsFinal = new OpenOptions();
            openOptionsFinal.SetOpenWorksetsConfiguration(finalWorksetConfig);
            openOptionsFinal.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets;

            try
            {
                Doc = App.OpenDocumentFile(modelPath, openOptionsFinal);

                if (Doc != null)
                {
                    Logger.Debug($"[RvtDocument] [Шаг 2] ✅ Документ открыт: '{Doc.Title}'");
                    Logger.Debug($"[RvtDocument] [DIAG] IsLinked={Doc.IsLinked}, PathName={Doc.PathName}");
                }

                LogSessionState("AFTER_FINAL_OPEN");
                return Doc;
            }
            catch (Exception ex)
            {
                Logger.Critical($"[RvtDocument] ❌ Критическая ошибка при финальном открытии: {ex.Message}");
                DebugWindow.AddRow($"CRITICAL: Не удалось открыть {Path}");
                throw;
            }
        }

        /// <summary>
        /// 🔥 Безопасное закрытие документа (только рабочие API Revit)
        /// </summary>
        private void SafeCloseDocument(Document doc, string context)
        {
            if (doc == null)
            {
                Logger.Debug($"[RvtDocument] {context} doc=null — пропускаем");
                return;
            }

            // 🔍 Проверяем валидность вместо IsClosed
            if (!doc.IsValidObject)
            {
                Logger.Debug($"[RvtDocument] {context} Документ уже невалиден — пропускаем");
                return;
            }

            Logger.Debug($"[RvtDocument] {context} Подготовка к закрытию: IsLinked={doc.IsLinked}, IsValid={doc.IsValidObject}, PathName={doc.PathName}");

            try
            {
                // 🔥 Для linked-файлов НЕ вызываем Close() — это вызывает ошибку
                if (doc.IsLinked)
                {
                    Logger.Warning($"[RvtDocument] {context} ⚠️ Документ является связью — пропускаем Close()");
                    return;
                }

                // Стандартное закрытие
                Logger.Debug($"[RvtDocument] {context} Вызов doc.Close(false)...");
                doc.Close(false);
                Logger.Debug($"[RvtDocument] {context} ✅ Close() выполнен");
            }
            catch (Autodesk.Revit.Exceptions.InvalidOperationException ex)
                when (ex.Message.IndexOf("linked", StringComparison.OrdinalIgnoreCase) >= 0 ||
                      ex.Message.IndexOf("Cannot close", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Logger.Warning($"[RvtDocument] {context} ⚠️ Нельзя закрыть linked-файл: {ex.Message}");
            }
            catch (Exception ex)
            {
                Logger.Warning($"[RvtDocument] {context} ⚠️ Ошибка при Close(): {ex.GetType().Name}: {ex.Message}");
            }

            // 🔹 Принудительный GC для очистки кэша (помогает при "залипании" IsLinked)
            try
            {
                System.GC.Collect();
                System.GC.WaitForPendingFinalizers();
                Logger.Debug($"[RvtDocument] {context} 🧹 GC выполнен");
            }
            catch { }
        }

        /// <summary>
        /// 🔍 Логирование состояния сессии (только рабочие API)
        /// </summary>
        private void LogSessionState(string stage)
        {
            try
            {
                // 🔥 Используем Documents.Count и безопасный перебор
                var docs = new List<Document>();
                foreach (Document d in App.Documents)
                {
                    if (d.IsValidObject) // 🔥 Проверяем валидность вместо IsClosed
                        docs.Add(d);
                }

                Logger.Debug(
                    $"[RvtDocument] [SESSION {stage}] Открыто валидных документов: {docs.Count}");

                foreach (var d in docs)
                {
                    try
                    {
                        Logger.Debug(
                            $"[RvtDocument]   • '{d.Title}' | Linked={d.IsLinked} | Modified={d.IsModified} | Path={d.PathName}");
                    }
                    catch { /* Игнорируем ошибки доступа к свойствам */ }
                }
            }
            catch (Exception ex)
            {
                Logger.Debug($"[RvtDocument] [SESSION {stage}] Ошибка диагностики: {ex.Message}");
            }
        }
    }
}
