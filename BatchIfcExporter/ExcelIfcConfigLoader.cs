using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
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

                    foreach (var ws in package.Workbook.Worksheets)
                    {
                        Logger.Debug($"[ExcelConfig]      Лист: '{ws.Name}' | {ws.Dimension?.Rows} строк x {ws.Dimension?.Columns} столбцов");
                    }

                    if (package.Workbook.Worksheets.Count == 0)
                    {
                        Logger.Error("[ExcelConfig] ❌ В файле нет листов");
                        return configs;
                    }

                    var worksheet = package.Workbook.Worksheets[1];
                    if (worksheet.Dimension == null)
                    {
                        Logger.Warning("[ExcelConfig] ️ Первый лист пустой (Dimension = null)");
                        return configs;
                    }

                    int rowCount = worksheet.Dimension.Rows;
                    int colCount = worksheet.Dimension.Columns;
                    Logger.Info($"[ExcelConfig] 📊 Таблица: {rowCount} строк x {colCount} столбцов");

                    if (rowCount >= 1)
                    {
                        string header = "";
                        for (int col = 1; col <= Math.Min(6, colCount); col++)
                            header += $"[{col}]={worksheet.Cells[1, col].Value?.ToString() ?? "null"} | ";
                        Logger.Debug($"[ExcelConfig] 📋 Заголовок: {header}");
                    }

                    for (int row = 2; row <= rowCount; row++)
                    {
                        try
                        {
                            var partName = worksheet.Cells[row, 1].Value?.ToString()?.Trim();
                            var jsonPath = worksheet.Cells[row, 2].Value?.ToString()?.Trim();
                            var viewName = worksheet.Cells[row, 3].Value?.ToString()?.Trim();
                            var mappingPath = worksheet.Cells[row, 4].Value?.ToString()?.Trim();
                            var worksetPattern = worksheet.Cells[row, 5].Value?.ToString()?.Trim();
                            var fileSuffix = colCount >= 6 ? worksheet.Cells[row, 6].Value?.ToString()?.Trim() : null;

                            Logger.Debug($"[ExcelConfig]   🔍 Строка {row}: Part='{partName}' | JSON='{jsonPath}' | View='{viewName}' | Map='{mappingPath}' | WS='{worksetPattern}' | Suffix='{fileSuffix}'");

                            if (string.IsNullOrEmpty(partName))
                            {
                                Logger.Debug($"[ExcelConfig]     ↳ Пропуск: PartOfModelName пуст");
                                continue;
                            }

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
                                WorksetExcludePattern = string.IsNullOrEmpty(worksetPattern) ? "Связь" : worksetPattern,
                                FileSuffix = fileSuffix
                            });
                            Logger.Info($"[ExcelConfig]     ↳ ✅ Добавлена конфигурация: '{partName}' (View: {viewName})");
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

        /// <summary>
        /// Возвращает список всех конфигураций, подходящих для данной модели.
        /// </summary>
        public static List<IfcModelConfig> ResolveAll(string modelFileName, List<IfcModelConfig> configs)
        {
            Logger.Debug($"[ExcelConfig]  Поиск всех конфигов для: {modelFileName} (доступно: {configs?.Count ?? 0})");

            var matched = new List<IfcModelConfig>();
            if (configs == null || configs.Count == 0)
            {
                Logger.Warning("[ExcelConfig] ⚠️ Список пуст → применяются дефолты");
                matched.Add(new IfcModelConfig { ViewName = "Navisworks" });
                return matched;
            }

            foreach (var cfg in configs)
            {
                if (cfg.IsMatch(modelFileName))
                {
                    Logger.Info($"[ExcelConfig]  MATCH FOUND: '{cfg.PartOfModelName}' → '{modelFileName}' (View: {cfg.ViewName})");
                    matched.Add(cfg);
                }
            }

            if (matched.Count == 0)
            {
                Logger.Warning("[ExcelConfig] ⚠️ Нет совпадений → применяются дефолты");
                matched.Add(new IfcModelConfig { ViewName = "Navisworks" });
            }
            else
            {
                Logger.Info($"[ExcelConfig] 📊 Найдено конфигураций для модели: {matched.Count}");
            }

            return matched;
        }
    }

    public class IfcModelConfig
    {
        public string PartOfModelName { get; set; }
        public string JsonConfigPath { get; set; }
        public string ViewName { get; set; }
        public string MappingFilePath { get; set; }
        public string WorksetExcludePattern { get; set; } = "Связь";
        public string FileSuffix { get; set; }

        public bool IsMatch(string modelFileName)
        {
            if (string.IsNullOrEmpty(PartOfModelName)) return false;
            bool match = modelFileName.IndexOf(PartOfModelName, StringComparison.OrdinalIgnoreCase) >= 0;
            if (match) Logger.Debug($"[ExcelConfig]     ↳ Подстрока '{PartOfModelName}' найдена в '{modelFileName}'");
            return match;
        }
    }
}