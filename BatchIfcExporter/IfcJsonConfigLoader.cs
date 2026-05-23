using Autodesk.Revit.DB;
using BatchIfcExporter;
using ISTools;
using Newtonsoft.Json.Linq;
using System;
using System.IO;
using System.Reflection;
using System.Xml.Linq;

namespace BatchExportIfc
{
    public class IfcJsonConfigLoader
    {
        #region Дефолтные значения
        public IFCVersion IFCVersion { get; set; } = IFCVersion.IFC4RV;
        public int SpaceBoundaries { get; set; } = 0;
        public bool SplitWallsAndColumns { get; set; } = false;
        public bool VisibleElementsOfCurrentView { get; set; } = true;
        public bool ExportRoomsInView { get; set; } = false;
        public bool Export2DElements { get; set; } = false;
        public bool ExportLinkedFiles { get; set; } = false;
        public bool ExportIFCCommonPropertySets { get; set; } = false;
        public bool ExportBaseQuantities { get; set; } = false;
        public bool ExportSchedulesAsPsets { get; set; } = false;
        public bool? ExportSpecificSchedules { get; set; } = false;
        public double TessellationLevelOfDetail { get; set; } = 0.1;
        public bool ExportPartsAsBuildingElements { get; set; } = false;
        public bool ExportSolidModelRep { get; set; } = false;
        public bool UseActiveViewGeometry { get; set; } = true;
        public bool UseFamilyAndTypeNameForReference { get; set; } = false;
        public bool Use2DRoomBoundaryForVolume { get; set; } = false;
        public bool IncludeSiteElevation { get; set; } = false;
        public bool StoreIFCGUID { get; set; } = false;
        public bool ExportBoundingBox { get; set; } = false;
        public string ExportUserDefinedPsetsFileName { get; set; } = null;
        public bool ExportUserDefinedPsets { get; set; } = false;
        #endregion

        public void ApplyJsonToConfig(string jsonPath, object ifcConfig, Type configType)
        {
            if (string.IsNullOrWhiteSpace(jsonPath) || !File.Exists(jsonPath))
            {
                Logger.Warning("IfcJsonLoader", "JSON не указан или не найден. Применяются дефолты.");
                ApplyDefaults(ifcConfig, configType);
                return;
            }

            Logger.Info("IfcJsonLoader", $"📄 Загрузка конфигурации: {Path.GetFileName(jsonPath)}");
            int applied = 0, skipped = 0, uiIgnored = 0, errors = 0;
            string mappingPathFromJson = null;
            try
            {
                JObject json = JObject.Parse(File.ReadAllText(jsonPath));

                foreach (var prop in json.Properties())
                {
                    string key = prop.Name;
                    JToken val = prop.Value;

                    Logger.Debug("IfcJsonLoader", $"🔍 [{key}] JSON тип: {val.Type} | Значение: {val}");

                    // 1. Пропуск null
                    if (val.Type == JTokenType.Null)
                    {
                        Logger.Debug("IfcJsonLoader", $"⏭ [{key}] Пропущено: значение null");
                        continue;
                    }

                    // 2. Пропуск вложенных объектов (UI-параметры)
                    if (val.Type == JTokenType.Object)
                    {
                        Logger.Warning("IfcJsonLoader", $"⏭ [{key}] Пропущено: вложенный объект (UI-параметр, не участвует в экспорте)");
                        uiIgnored++;
                        continue;
                    }

                    // 3. Поиск свойства в IFCExportConfiguration
                    var targetProp = configType.GetProperty(key, BindingFlags.Public | BindingFlags.Instance);
                    if (targetProp == null || !targetProp.CanWrite)
                    {
                        Logger.Warning("IfcJsonLoader", $"⏭ [{key}] Пропущено: свойство отсутствует в IFCExportConfiguration Revit");
                        skipped++;
                        continue;
                    }
                    Logger.Debug("IfcJsonLoader", $"  ↳ API тип свойства: {targetProp.PropertyType.Name}");

                    // 4. Конвертация и установка (изолированный try-catch)
                    try
                    {
                        object converted = ConvertValue(val, targetProp.PropertyType, key);

                        // Обработка bool → Enum (для ExportLinkedFiles)
                        if (converted is bool b && targetProp.PropertyType.IsEnum)
                            converted = ResolveBooleanToEnum(b, targetProp.PropertyType);

                        if (converted == null)
                        {
                            Logger.Warning("IfcJsonLoader", $"⚠️ [{key}] Конвертация вернула null");
                            errors++;
                            continue;
                        }

                        targetProp.SetValue(ifcConfig, converted);
                        if (key == "ExportUserDefinedPsetsFileName") mappingPathFromJson = converted.ToString();

                        applied++;
                        Logger.Debug("IfcJsonLoader", $"✅ [{key}] Успешно применено: {converted}");
                    }
                    catch (Exception ex)
                    {
                        Logger.Error("IfcJsonLoader", $"❌ [{key}] Ошибка применения: {ex.Message}");
                        errors++;
                    }
                }

                // Валидация пути к мэппингу
                if (!string.IsNullOrEmpty(mappingPathFromJson))
                {
                    if (File.Exists(mappingPathFromJson))
                        Logger.Info("IfcJsonLoader", $"📎 Мэппинг найден: {Path.GetFileName(mappingPathFromJson)}");
                    else
                        Logger.Warning("IfcJsonLoader", $"⚠️ Мэппинг не найден на диске: {mappingPathFromJson}");
                }

                Logger.Info("IfcJsonLoader", $"📊 Итог парсинга: Применено={applied} | Пропущено(API)={skipped} | UI/Null={uiIgnored} | Ошибки={errors}");
            }
            catch (Exception ex)
            {
                Logger.Error("IfcJsonLoader", $"❌ Критическая ошибка парсинга JSON: {ex.Message}. Fallback на дефолты.");
                ApplyDefaults(ifcConfig, configType);
            }
        }

        /// <summary>
        /// Безопасная конвертация JToken в целевой тип. Обходит JIT-ошибки Revit 2024+.
        /// </summary>
        private object ConvertValue(JToken token, Type targetType, string propName)
        {
            try
            {
                Logger.Debug("IfcJsonLoader", $"  🔄 Конвертация '{propName}': {token.Type} → {targetType.Name}");

                if (targetType == typeof(bool)) return token.ToObject<bool>();
                if (targetType == typeof(int)) return token.ToObject<int>();
                if (targetType == typeof(long)) return token.ToObject<long>();
                if (targetType == typeof(double)) return token.ToObject<double>();
                if (targetType == typeof(string)) return token.ToString();

                if (targetType == typeof(ElementId))
                {
                    long raw = token.ToObject<long>();
                    if (raw < int.MinValue || raw > int.MaxValue)
                    {
                        Logger.Warning("IfcJsonLoader", $"  ⚠️ '{propName}': значение {raw} выходит за диапазон int");
                        return null;
                    }
                    return SafeCreateElementId((int)raw);
                }

                if (targetType.IsEnum)
                {
                    if (token.Type == JTokenType.Integer) return Enum.ToObject(targetType, token.ToObject<int>());
                    if (token.Type == JTokenType.String) return Enum.Parse(targetType, token.ToString(), true);
                }

                Logger.Warning("IfcJsonLoader", $"  ⏭ '{propName}': тип {targetType.Name} не поддерживается для конвертации");
                return null;
            }
            catch (Exception ex)
            {
                Logger.Warning("IfcJsonLoader", $"  ⚠️ '{propName}': ошибка конвертации → {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Безопасное создание ElementId через рефлексию.
        /// Избегает JIT-ошибки "Method not found: ElementId..ctor(Int32)" в Revit 2024+.
        /// </summary>
        private static object SafeCreateElementId(int value)
        {
            try
            {
                var ctor = typeof(ElementId).GetConstructor(new[] { typeof(int) });
                if (ctor != null) return ctor.Invoke(new object[] { value });

                // Fallback: в некоторых версиях используется статический метод или поле
                Logger.Warning("IfcJsonLoader", $"  ⚠️ Конструктор ElementId(int) недоступен. Пропускаем.");
                return null;
            }
            catch (Exception ex)
            {
                Logger.Error("IfcJsonLoader", $"  ❌ SafeCreateElementId: {ex.Message}");
                return null;
            }
        }

        private void ApplyDefaults(object cfg, Type type)
        {
            ApplyProp(cfg, type, "IFCVersion", IFCVersion); ApplyProp(cfg, type, "SpaceBoundaries", SpaceBoundaries);
            ApplyProp(cfg, type, "SplitWallsAndColumns", SplitWallsAndColumns); ApplyProp(cfg, type, "VisibleElementsOfCurrentView", VisibleElementsOfCurrentView);
            ApplyProp(cfg, type, "ExportRoomsInView", ExportRoomsInView); ApplyProp(cfg, type, "Export2DElements", Export2DElements);
            ApplyProp(cfg, type, "ExportLinkedFiles", ExportLinkedFiles); ApplyProp(cfg, type, "ExportIFCCommonPropertySets", ExportIFCCommonPropertySets);
            ApplyProp(cfg, type, "ExportBaseQuantities", ExportBaseQuantities); ApplyProp(cfg, type, "ExportSchedulesAsPsets", ExportSchedulesAsPsets);
            ApplyProp(cfg, type, "ExportSpecificSchedules", ExportSpecificSchedules); ApplyProp(cfg, type, "TessellationLevelOfDetail", TessellationLevelOfDetail);
            ApplyProp(cfg, type, "ExportPartsAsBuildingElements", ExportPartsAsBuildingElements); ApplyProp(cfg, type, "ExportSolidModelRep", ExportSolidModelRep);
            ApplyProp(cfg, type, "UseActiveViewGeometry", UseActiveViewGeometry); ApplyProp(cfg, type, "UseFamilyAndTypeNameForReference", UseFamilyAndTypeNameForReference);
            ApplyProp(cfg, type, "Use2DRoomBoundaryForVolume", Use2DRoomBoundaryForVolume); ApplyProp(cfg, type, "IncludeSiteElevation", IncludeSiteElevation);
            ApplyProp(cfg, type, "StoreIFCGUID", StoreIFCGUID); ApplyProp(cfg, type, "ExportBoundingBox", ExportBoundingBox);
        }

        private void ApplyProp(object c, Type t, string n, object v)
        {
            var p = t.GetProperty(n, BindingFlags.Public | BindingFlags.Instance);
            if (p?.CanWrite == true) try { p.SetValue(c, v); } catch { }
        }

        private object ResolveBooleanToEnum(bool val, Type enumType)
        {
            foreach (var n in (val ? new[] { "ExportAsSeparate", "Export" } : new[] { "DoNotExport", "DontExport" }))
                if (Enum.IsDefined(enumType, n)) return Enum.Parse(enumType, n);
            var v = Enum.GetValues(enumType);
            return v.Length > 0 ? v.GetValue(val ? 1 : 0) : null;
        }
    }
}