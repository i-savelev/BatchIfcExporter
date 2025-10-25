using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using ISTools;
using System;
using System.Linq;
using BatchExportIfc;


namespace BatchExportIfc
{
    internal class RvtDocument
    {
        public Application App {  get; set; }
        public string Path { get; set; }
        public Document Doc { get; set; }
        public RvtDocument(Application app, string path) 
        {
            App = app;
            Path = path;
        }
        public Document Open(string wsExcludeWord)
        {
            ModelPath modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(Path);

            // Шаг 1: Быстрое открытие для анализа рабочих наборов
            OpenOptions openOptionsStep1 = new OpenOptions();
            openOptionsStep1.SetOpenWorksetsConfiguration(new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets));
            openOptionsStep1.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets;

            Document docTemp = null;
            WorksetConfiguration finalWorksetConfig = null;

            try
            {
                docTemp = App.OpenDocumentFile(modelPath, openOptionsStep1);

                var allWorksets = new FilteredWorksetCollector(docTemp).ToWorksets();

                // Фильтруем: оставляем открытыми ТОЛЬКО наборы, НЕ содержащие слово "связь"
                var worksetIdsToOpen = allWorksets
                    .Where(w => !w.Name.Contains(wsExcludeWord))
                    .Select(w => w.Id)
                    .ToList();

                finalWorksetConfig = new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets);
                finalWorksetConfig.Open(worksetIdsToOpen);
                docTemp?.Close(false);
            }
            catch (Exception ex)
            {
                docTemp?.Close(false);
                IsDebugWindow.AddRow($"Ошибка при подготовке конфигурации: {ex.Message}");
            }

            // Шаг 2: Открываем с фильтром
            OpenOptions openOptionsFinal = new OpenOptions();
            openOptionsFinal.SetOpenWorksetsConfiguration(finalWorksetConfig);
            openOptionsFinal.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets;
            Doc = App.OpenDocumentFile(modelPath, openOptionsFinal);
            return Doc;
        }
    }
}
