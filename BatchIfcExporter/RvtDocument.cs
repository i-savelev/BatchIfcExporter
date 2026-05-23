using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using BatchExportIfc;
using ISTools;
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
            Logger.Debug("RvtDocument", $"Инициализация | Путь: {path}");
        }

        public Document Open(string wsExcludeWord)
        {
            Logger.Info("RvtDocument", $"▶ Открытие документа | Исключаем ворксеты с: '{wsExcludeWord}'");
            ModelPath modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(Path);


            // 🔹 ШАГ 1: Быстрое открытие для анализа рабочих наборов
            Logger.Debug("RvtDocument", "[Шаг 1] Открытие для анализа ворксетов (CloseAllWorksets)");
            OpenOptions openOptionsStep1 = new OpenOptions();
            openOptionsStep1.SetOpenWorksetsConfiguration(new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets));
            openOptionsStep1.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets;

            Document docTemp = null;
            WorksetConfiguration finalWorksetConfig = null;
            List<string> closedWorksets = new List<string>();  // 📋 Список закрытых

            try
            {
                docTemp = App.OpenDocumentFile(modelPath, openOptionsStep1);
                string fileName = docTemp.Title;
                Logger.Debug("RvtDocument", $"[Шаг 1] Документ временно открыт: {docTemp.Title}");

                var allWorksets = new FilteredWorksetCollector(docTemp).ToWorksets().ToList();
                Logger.Info("RvtDocument", $"Всего ворксетов в документе: {allWorksets.Count}");


                // 🔹 Фильтрация: закрываем ворксеты, содержащие ключевое слово
                var worksetIdsToOpen = allWorksets
                    .Where(w => w.Name.IndexOf(wsExcludeWord, StringComparison.OrdinalIgnoreCase) < 0)
                    .Select(w => w.Id)
                    .ToList();

                // 📋 Формируем списки для лога
                foreach (var ws in allWorksets)
                {
                    if (ws.Name.Contains(wsExcludeWord))
                    {
                        closedWorksets.Add($"{ws.Name} (Id:{ws.Id})");
                        Logger.Debug("RvtDocument", $"🔒 ЗАКРЫВАЕТСЯ: '{ws.Name}' (содержит '{wsExcludeWord}')");
                    }
                }

                // 📊 Итоговый отчёт по ворксетам
                Logger.Info("RvtDocument", $"[Worksets Summary] Закрывается: {closedWorksets.Count}");
                if (closedWorksets.Any())
                {
                    Logger.Info("RvtDocument", $"[Closed Worksets] {string.Join("; ", closedWorksets)}");
                    IsDebugWindow.AddRow($"[{fileName}] Закрыто ворксетов: {closedWorksets.Count} | {string.Join(", ", closedWorksets.Take(3))}{(closedWorksets.Count > 3 ? "..." : "")}");
                }

                finalWorksetConfig = new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets);
                finalWorksetConfig.Open(worksetIdsToOpen);

                docTemp?.Close(false);
                Logger.Debug("RvtDocument", "[Шаг 1] Временный документ закрыт");
            }
            catch (Exception ex)
            {
                Logger.Error("RvtDocument", $"❌ Ошибка на шаге 1 (анализ ворксетов): {ex.Message}");
                IsDebugWindow.AddRow($"Ошибка анализа ворксетов: {ex.Message}");
                docTemp?.Close(false);
                // Fallback: открываем со всеми ворксетами
                Logger.Warning("RvtDocument", "Fallback: открытие документа со всеми ворксетами");
                OpenOptions fallbackOptions = new OpenOptions();
                fallbackOptions.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets;
                Doc = App.OpenDocumentFile(modelPath, fallbackOptions);
                return Doc;
            }

            // 🔹 ШАГ 2: Финальное открытие с применённой конфигурацией
            Logger.Debug("RvtDocument", "[Шаг 2] Финальное открытие документа с фильтром ворксетов");
            OpenOptions openOptionsFinal = new OpenOptions();
            openOptionsFinal.SetOpenWorksetsConfiguration(finalWorksetConfig);
            openOptionsFinal.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets;

            try
            {
                Doc = App.OpenDocumentFile(modelPath, openOptionsFinal);

                return Doc;
            }
            catch (Exception ex)
            {
                Logger.Critical("RvtDocument", $"❌ Критическая ошибка при финальном открытии: {ex.Message}\n{ex.StackTrace}");
                IsDebugWindow.AddRow($"CRITICAL: Не удалось открыть {Path}");
                throw;
            }
        }
    }
}
