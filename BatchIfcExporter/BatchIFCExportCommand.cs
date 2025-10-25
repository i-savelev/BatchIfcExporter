using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ISTools;
using BatchExportIfc;


[Autodesk.Revit.Attributes.TransactionAttribute(Autodesk.Revit.Attributes.TransactionMode.Manual)]



public class BatchIFCExportCommand : IExternalCommand
{
    public static string IS_TAB_NAME => "ISTools";
    public static string IS_NAME => "Экспорт IFC";
    public static string IS_IMAGE => "BatchIfcExporter.Resources.IFC.png";
    public static string IS_DESCRIPTION => "";
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        UIApplication uiapp = commandData.Application;
        Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
        ISTimer.Start();
        // Диалог выбора папки с RVT-файлами
        OpenFileDialog folderDialog = new OpenFileDialog();
        folderDialog.Title = "Выберете папку с моделямя Revit";
        folderDialog.Filter = "Папки|\n";
        folderDialog.FileName = "Выбор папки";
        folderDialog.CheckFileExists = false;
        folderDialog.CheckPathExists = true;
        folderDialog.ValidateNames = false;
        if (folderDialog.ShowDialog() != DialogResult.OK) return Result.Cancelled;

        string folderPath = Path.GetDirectoryName(folderDialog.FileName); ;
        string[] rvtFiles = Directory.GetFiles(folderPath, "*.rvt");

        // Диалог выбора файла сопоставления
        string mappingFilePath = null;
        OpenFileDialog mappingDialog = new OpenFileDialog();
        mappingDialog.Title = "Выберете файл сопоставления";
        mappingDialog.Filter = "IFC Mapping Files|*.txt;";
        if (mappingDialog.ShowDialog() != DialogResult.OK) mappingFilePath = null;
        mappingFilePath = mappingDialog.FileName;

        // Запрос имени вида
        string viewName = "Navisworks";
        BatchExportIFC(app, new List<string>(rvtFiles), viewName, mappingFilePath);
        IsDebugWindow.AddRow(ISTimer.Stop());
        IsDebugWindow.Show();
        return Result.Succeeded;
    }

    private void BatchExportIFC(
        Autodesk.Revit.ApplicationServices.Application app, 
        List<string> rvtFiles, 
        string viewName, 
        string mappingFilePath
        )
    {
        foreach (string rvtPath in rvtFiles)
        {
            if (!File.Exists(rvtPath))
            {
                IsDebugWindow.AddRow($"Файл не найден: {rvtPath}");
                continue;
            }
            
            ModelPath modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(rvtPath);
            OpenOptions openOptions = new OpenOptions();

            //doc = app.OpenDocumentFile(modelPath, openOptions);
            var rvtDoc = new RvtDocument(app, rvtPath);
            var doc = rvtDoc.Open("Связь");
            // Находим вид по имени
            Autodesk.Revit.DB.View targetView = new FilteredElementCollector(doc)
                .OfClass(typeof(Autodesk.Revit.DB.View))
                .Cast<Autodesk.Revit.DB.View>()
                .FirstOrDefault(v => v.Name.Equals(viewName, StringComparison.OrdinalIgnoreCase) && !v.IsTemplate);

            if (targetView == null)
            {
                IsDebugWindow.AddRow($"Вид '{viewName}' не найден в файле {rvtPath}");
                continue;
            }
            var config = new IfcExportConfig(doc, viewName, mappingFilePath);
            var exportOptions = config.GetConfig();
            string ifcPath = Path.ChangeExtension(rvtPath, ".ifc");
            string folder = Path.GetDirectoryName(ifcPath);
            string filename = Path.GetFileName(ifcPath);
            ICollection<ElementId> viewIds = new List<ElementId> { targetView.Id };
            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("1");
                doc.Export(folder, filename, exportOptions);
                tx.Commit();
            }
            
            IsDebugWindow.AddRow($"Экспорт завершён: {ifcPath}");


            if (doc != null)
            {
                doc.Close(false); // не сохраняем изменения
            }
        }
    }
}