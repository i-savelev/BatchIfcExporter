using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using BatchIfcExporter;
using Logger = RevitLogger.Logger;

namespace BatchExportIfc
{
    public static class ExcelIfcConfigLoader
    {
        public static List<IfcModelConfig> Load(string excelPath)
        {
            var configs = new List<IfcModelConfig>();
            Logger.Info($"[ExcelConfig]  Начало загрузки Excel: {excelPath}");

            if (!File.Exists(excelPath))
            {
                Logger.Error($"[ExcelConfig] ❌ Файл не найден: {excelPath}");
                return configs;
            }

            try
            {
                var fileInfo = new FileInfo(excelPath);
                Logger.Debug($"[ExcelConfig]   ↳ Размер: {fileInfo.Length} байт | Изменён: {fileInfo.LastWriteTime}");

                using (var package = new ExcelPackage(fileInfo))
                {
                    Logger.Debug($"[ExcelConfig]   ↳ Листов в книге: {package.Workbook.Worksheets.Count}");

                    //  Безопасный перебор для логирования (без риска IndexOutOfRangeException)
                    foreach (var ws in package.Workbook.Worksheets)
                    {
                        Logger.Debug($"[ExcelConfig]      Лист: '{ws.Name}' | {ws.Dimension?.Rows} строк x {ws.Dimension?.Columns} столбцов");
                    }

                    if (package.Workbook.Worksheets.Count == 0)
                    {
                        Logger.Error("[ExcelConfig] ❌ В файле нет листов");
                        return configs;
                    }

                    // 🔥 ИСПРАВЛЕНИЕ: EPPlus 4.x использует 1-based индексацию!
                    var worksheet = package.Workbook.Worksheets[1];
                    if (worksheet.Dimension == null)
                    {
                        Logger.Warning("[ExcelConfig] ️ Первый лист пустой (Dimension = null)");
                        return configs;
                    }

                    int rowCount = worksheet.Dimension.Rows;
                    int colCount = worksheet.Dimension.Columns;
                    Logger.Info($"[ExcelConfig] 📊 Таблица: {rowCount} строк x {colCount} столбцов");

                    // Логируем заголовок (1-я строка)
                    if (rowCount >= 1)
                    {
                        string header = "";
                        for (int col = 1; col <= Math.Min(4, colCount); col++)
                            header += $"[{col}]={worksheet.Cells[1, col].Value?.ToString() ?? "null"} | ";
                        Logger.Debug($"[ExcelConfig] 📋 Заголовок: {header}");
                    }

                    // Читаем данные со 2-й строки
                    for (int row = 2; row <= rowCount; row++)
                    {
                        try
                        {
                            var partName = worksheet.Cells[row, 1].Value?.ToString()?.Trim();
                            var jsonPath = worksheet.Cells[row, 2].Value?.ToString()?.Trim();
                            var viewName = worksheet.Cells[row, 3].Value?.ToString()?.Trim();
                            var mappingPath = worksheet.Cells[row, 4].Value?.ToString()?.Trim();
                            var worksetPattern = worksheet.Cells[row, 5].Value?.ToString()?.Trim();
                            Logger.Debug($"[ExcelConfig]   🔍 Строка {row}: Part='{partName}' | JSON='{jsonPath}' | View='{viewName}' | Map='{mappingPath}' | WS='{worksetPattern}'");

                            if (string.IsNullOrEmpty(partName))
                            {
                                Logger.Debug($"[ExcelConfig]     ↳ Пропуск: PartOfModelName пуст");
                                continue;
                            }

                            // Валидация путей к файлам
                            if (!string.IsNullOrEmpty(jsonPath) && !File.Exists(jsonPath))
                                Logger.Warning($"[ExcelConfig]     ↳ ⚠️ JSON не найден на диске: {jsonPath}");
                            if (!string.IsNullOrEmpty(mappingPath) && !File.Exists(mappingPath))
                                Logger.Warning($"[ExcelConfig]     ↳ ⚠️ Mapping не найден на диске: {mappingPath}");

                            configs.Add(new IfcModelConfig
                            {
                                PartOfModelName = partName,
                                JsonConfigPath = jsonPath,
                                ViewName = string.IsNullOrEmpty(viewName) ? "Navisworks" : viewName,
                                MappingFilePath = mappingPath,
                                WorksetExcludePattern = string.IsNullOrEmpty(worksetPattern) ? "Связь" : worksetPattern
                            });
                            Logger.Info($"[ExcelConfig]     ↳ ✅ Добавлена конфигурация: '{partName}'");
                        }
                        catch (Exception rowEx)
                        {
                            Logger.Error($"[ExcelConfig]   ❌ Ошибка чтения строки {row}: {rowEx.Message}");
                        }
                    }
                }
                Logger.Info($"[ExcelConfig]  Итого загружено конфигураций: {configs.Count}");
            }
            catch (Exception ex)
            {
                Logger.Error($"[ExcelConfig] ❌ Критическая ошибка: {ex.GetType().Name}: {ex.Message}");
                Logger.Error($"[ExcelConfig] StackTrace: {ex.StackTrace}");
            }
            return configs;
        }

        public static IfcModelConfig Resolve(string modelFileName, List<IfcModelConfig> configs)
        {
            Logger.Debug($"[ExcelConfig]  Поиск для: {modelFileName} (доступно конфигов: {configs.Count})");

            if (configs == null || configs.Count == 0)
            {
                Logger.Warning("[ExcelConfig] ⚠️ Список пуст → применяются дефолты");
                return new IfcModelConfig { ViewName = "Navisworks" };
            }

            foreach (var cfg in configs)
            {
                if (cfg.IsMatch(modelFileName))
                {
                    Logger.Info($"[ExcelConfig]  MATCH FOUND: '{cfg.PartOfModelName}' → '{modelFileName}'");
                    return cfg;
                }
            }

            Logger.Warning("[ExcelConfig] ⚠️ Нет совпадений → применяются дефолты");
            return new IfcModelConfig { ViewName = "Navisworks" };
        }
    }

    public class IfcModelConfig
    {
        public string PartOfModelName { get; set; }
        public string JsonConfigPath { get; set; }
        public string ViewName { get; set; }
        public string MappingFilePath { get; set; }
        public string WorksetExcludePattern { get; set; } = "Связь";
        public bool IsMatch(string modelFileName)
        {
            if (string.IsNullOrEmpty(PartOfModelName)) return false;
            bool match = modelFileName.IndexOf(PartOfModelName, StringComparison.OrdinalIgnoreCase) >= 0;
            if (match) Logger.Debug($"[ExcelConfig]     ↳ Подстрока '{PartOfModelName}' найдена в '{modelFileName}'");
            return match;
        }
    }
}
