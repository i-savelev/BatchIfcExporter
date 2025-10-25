using Autodesk.Revit.DB;
using System;
using System.IO;
using System.Linq;
using ISTools;

namespace BatchExportIfc
{
    public class IfcExportConfig
    {
        public string ViewName { get; set; }
        public string MappingFile { get; set; }
        public Document Doc { get; set; }
        // --- Общие настройки ---
        public IFCVersion IFCVersion { get; set; } = IFCVersion.IFC4RV;
        public int SpaceBoundaries { get; set; } = 0; // 0 = нет, 1 = 1-го уровня, 2 = 2-го уровня
        public bool SplitWallsAndColumns { get; set; } = false;

        // --- Дополнительное содержимое ---
        public bool VisibleElementsOfCurrentView { get; set; } = true;
        public bool ExportRoomsInView { get; set; } = false;
        public bool Export2DElements { get; set; } = false;
        public bool ExportLinkedFiles { get; set; } = false; // true = экспортировать, false = не экспортировать

        // --- Наборы свойств ---
        public bool ExportIFCCommonPropertySets { get; set; } = false;
        public bool ExportBaseQuantities { get; set; } = false;
        public bool ExportSchedulesAsPsets { get; set; } = false;
        public bool ExportSpecificSchedules { get; set; } = false;

        // --- Уровень детализации ---
        public double TessellationLevelOfDetail { get; set; } = 0.1; // от 0.01 до 1.0

        // --- Расширенные настройки ---
        public bool ExportPartsAsBuildingElements { get; set; } = false;
        public bool ExportSolidModelRep { get; set; } = false;
        public bool UseActiveViewGeometry { get; set; } = true;
        public bool UseFamilyAndTypeNameForReference { get; set; } = false;
        public bool Use2DRoomBoundaryForVolume { get; set; } = false;
        public bool IncludeSiteElevation { get; set; } = false;
        public bool StoreIFCGUID { get; set; } = false;
        public bool ExportBoundingBox { get; set; } = false;


        public IfcExportConfig(
            Document doc,
            string viewName,
            string mappingFile
            )
        {
            Doc = doc;
            ViewName = viewName;
            MappingFile = mappingFile;
        }

        public IFCExportOptions GetConfig()
        {
            Autodesk.Revit.DB.View targetView = new FilteredElementCollector(Doc)
                .OfClass(typeof(Autodesk.Revit.DB.View))
                .Cast<Autodesk.Revit.DB.View>()
                .FirstOrDefault(v => v.Name.Equals(ViewName, StringComparison.OrdinalIgnoreCase) && !v.IsTemplate);

            if (targetView == null)
            {
                IsDebugWindow.AddRow($"Вид '{ViewName}' не найден в файле {Doc.Title}");
              
            }

            BIM.IFC.Export.UI.IFCExportConfiguration config = BIM.IFC.Export.UI.IFCExportConfiguration.CreateDefaultConfiguration();
            // --- Основные настройки ---
            //config.ExportIFCCommonPropertySets = true;
            //config.ExportBaseQuantities = true;
            config.IFCVersion = IFCVersion;
            config.SpaceBoundaries = SpaceBoundaries;
            config.SplitWallsAndColumns = SplitWallsAndColumns;
            config.VisibleElementsOfCurrentView = VisibleElementsOfCurrentView;
            config.ExportRoomsInView = ExportRoomsInView;
            config.Export2DElements = Export2DElements;
            config.ExportLinkedFiles = ExportLinkedFiles;
            config.ExportIFCCommonPropertySets = ExportIFCCommonPropertySets;
            config.ExportBaseQuantities = ExportBaseQuantities;
            config.ExportSchedulesAsPsets = ExportSchedulesAsPsets;
            config.ExportSpecificSchedules = ExportSpecificSchedules;
            if (MappingFile == null)
            {
                if (File.Exists(MappingFile))
                {
                    config.ExportUserDefinedPsets = true;
                    config.ExportUserDefinedPsetsFileName = MappingFile;
                }
                else
                {
                    IsDebugWindow.AddRow($"Файл сопоставления не найден: {MappingFile}");
                }
            }
            
            config.TessellationLevelOfDetail = TessellationLevelOfDetail;
            config.ExportPartsAsBuildingElements = ExportPartsAsBuildingElements;
            config.ExportSolidModelRep = ExportSolidModelRep;
            config.UseActiveViewGeometry = UseActiveViewGeometry;
            config.UseFamilyAndTypeNameForReference = UseFamilyAndTypeNameForReference;
            config.Use2DRoomBoundaryForVolume = Use2DRoomBoundaryForVolume;
            config.IncludeSiteElevation = IncludeSiteElevation;
            config.StoreIFCGUID = StoreIFCGUID;
            config.ExportBoundingBox = ExportBoundingBox;

            // Создаём опции экспорта и применяем конфигурацию
            IFCExportOptions exportOptions = new IFCExportOptions();
            config.UpdateOptions(exportOptions, targetView.Id); // ← Применяем конфиг + указываем ViewId
            return exportOptions;
        }
    }
}
