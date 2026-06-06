using Autodesk.Revit.DB;
using BatchIfcExporter;
using Logger = RevitLogger.Logger;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using Document = Autodesk.Revit.DB.Document;

namespace BatchExportIfc
{
    public class IfcExportConfig
    {
        public Document Doc { get; }
        public string ViewName { get; }
        public string JsonConfigPath { get; }

        public IfcExportConfig(Document doc, string viewName, string jsonConfigPath = null, string mappingOverridePath = null)
        {
            Doc = doc;
            ViewName = viewName;
            JsonConfigPath = jsonConfigPath;
            _mappingOverridePath = mappingOverridePath;
            Logger.Debug($"[IfcExportConfig] Init | View={viewName}, JSON={jsonConfigPath ?? "default"}, MapOverride={mappingOverridePath ?? "null"}");
        }
        private readonly string _mappingOverridePath;

        public IFCExportOptions GetConfig()
        {
            Logger.Info($"[IfcExportConfig] ▶ GetConfig() | Вид: '{ViewName}'");

            var targetView = new FilteredElementCollector(Doc)
                .OfClass(typeof(Autodesk.Revit.DB.View))
                .Cast<Autodesk.Revit.DB.View>()
                .FirstOrDefault(v => v.Name.Equals(ViewName, StringComparison.OrdinalIgnoreCase) && !v.IsTemplate);

            if (targetView == null)
            {
                Logger.Error($"[IfcExportConfig] ❌ Вид '{ViewName}' не найден");
                return null;
            }

            var config = CreateDefaultConfiguration();
            if (config == null) return null;

            // 🔥 Загрузка и применение JSON (включая ExportUserDefinedPsetsFileName)
            var jsonLoader = new IfcJsonConfigLoader();
            jsonLoader.ApplyJsonToConfig(JsonConfigPath, config, config.GetType());
            if (!string.IsNullOrEmpty(_mappingOverridePath) && File.Exists(_mappingOverridePath))
            {
                SetConfigProperty(config, "ExportUserDefinedPsets", true);
                SetConfigProperty(config, "ExportUserDefinedPsetsFileName", _mappingOverridePath);
                Logger.Info($"[IfcExportConfig] 🔄 Мэппинг переопределён из Excel: {Path.GetFileName(_mappingOverridePath)}");
            }
            // Создание IFCExportOptions
            var exportOptions = (IFCExportOptions)Activator.CreateInstance(typeof(IFCExportOptions));
            UpdateOptionsWithContext(Doc, config, exportOptions, targetView.Id, config.GetType());

            Logger.Info("[IfcExportConfig] ✅ GetConfig() завершён");
            return exportOptions;
        }

        #region Вспомогательные методы (без изменений, кроме сигнатур)
        private object CreateDefaultConfiguration()
        {
            string revitPath = Path.GetDirectoryName(typeof(Document).Assembly.Location);
            string expectedPath = Path.Combine(revitPath, @"AddIns\IFCExporterUI\Autodesk.IFC.Export.UI.dll");
            Logger.Debug($"[IfcExportConfig] Поиск сборки: {expectedPath}");

            // 🔍 Безопасный поиск без LINQ (обход NotSupportedException в динамических сборках)
            Assembly loadedAssembly = null;
            foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
            {
                try
                {
                    if (!string.IsNullOrEmpty(asm.Location) &&
                        string.Equals(asm.Location, expectedPath, StringComparison.OrdinalIgnoreCase))
                    {
                        loadedAssembly = asm;
                        Logger.Debug($"[IfcExportConfig] Сборка найдена в AppDomain: {asm.FullName}");
                        break;
                    }
                }
                catch (NotSupportedException) { /* Динамические сборки не поддерживают .Location */ }
                catch { /* Игнорируем другие ошибки доступа к метаданным */ }
            }

            // Fallback: загрузка с диска, если не найдена в памяти
            if (loadedAssembly == null && File.Exists(expectedPath))
            {
                try
                {
                    loadedAssembly = Assembly.LoadFrom(expectedPath);
                    Logger.Info($"[IfcExportConfig] Сборка загружена вручную: {loadedAssembly.FullName}");
                }
                catch (Exception ex)
                {
                    Logger.Error($"[IfcExportConfig] ❌ Ошибка ручной загрузки сборки: {ex.Message}");
                    return null;
                }
            }

            if (loadedAssembly == null)
            {
                Logger.Error("[IfcExportConfig] ❌ Сборка IFCExporterUI не найдена ни в памяти, ни по пути");
                return null;
            }

            Type configType = loadedAssembly.GetType("BIM.IFC.Export.UI.IFCExportConfiguration");
            if (configType == null)
            {
                Logger.Error("[IfcExportConfig] ❌ Тип IFCExportConfiguration не найден в сборке");
                return null;
            }

            var method = configType.GetMethod("CreateDefaultConfiguration", BindingFlags.Public | BindingFlags.Static);
            if (method == null)
            {
                Logger.Error("[IfcExportConfig] ❌ Метод CreateDefaultConfiguration не найден");
                return null;
            }

            return method.Invoke(null, null);
        }

        private void SetConfigProperty(object config, string name, object value)
        {
            if (config == null) return;
            var prop = config.GetType().GetProperty(name, BindingFlags.Public | BindingFlags.Instance);
            if (prop?.CanWrite == true)
            {
                try { prop.SetValue(config, value); }
                catch (Exception ex) { Logger.Warning($"[IfcExportConfig] ⚠️ Ошибка установки '{name}': {ex.Message}"); }
            }
        }

        private void UpdateOptionsWithContext(Document document, object configuration, IFCExportOptions options, ElementId viewId, Type configType)
        {
            try
            {
                Assembly ifcAssembly = configType.Assembly;
                Type commandType = ifcAssembly.GetType("BIM.IFC.Export.UI.IFCCommandOverrideApplication");
                if (commandType == null) { TryUpdateOptions(configType, configuration, options, viewId, document); return; }

                var prop = commandType.GetProperty("TheDocument", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
                if (prop?.GetSetMethod(true) == null || prop.GetGetMethod(true) == null)
                { TryUpdateOptions(configType, configuration, options, viewId, document); return; }

                var getter = prop.GetGetMethod(true);
                var setter = prop.GetSetMethod(true);
                object prev = getter.Invoke(null, null);
                try
                {
                    setter.Invoke(null, new object[] { document });
                    TryUpdateOptions(configType, configuration, options, viewId, document);
                }
                finally { setter.Invoke(null, new object[] { prev }); }
            }
            catch (Exception ex) { Logger.Error($"[IfcExportConfig] ❌ UpdateOptionsWithContext: {ex.Message}"); }
        }

        private void TryUpdateOptions(Type configType, object config, IFCExportOptions options, ElementId viewId, Document document)
        {
            var newMethod = configType.GetMethod("UpdateOptions", new[] { typeof(Document), typeof(IFCExportOptions), typeof(ElementId), typeof(bool) });
            if (newMethod != null) { newMethod.Invoke(config, new object[] { document, options, viewId, false }); return; }

            var legacyMethod = configType.GetMethod("UpdateOptions", new[] { typeof(IFCExportOptions), typeof(ElementId) });
            if (legacyMethod != null) { legacyMethod.Invoke(config, new object[] { options, viewId }); }
        }
        #endregion
    }
}
