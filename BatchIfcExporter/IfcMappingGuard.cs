using BatchIfcExporter;
using ISTools;
using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace BatchExportIfc
{
    /// <summary>
    /// Защитник файла мэппинга для обхода бага Autodesk IFC Exporter 2026.
    /// Мониторит изменение файла после первого экспорта и при необходимости восстанавливает оригинал.
    /// </summary>
    public class IfcMappingGuard : IDisposable
    {
        private readonly string _mappingFilePath;
        private readonly string _originalHash;
        private readonly string _backupPath;
        private bool _isFirstExport = true;
        private bool _isDisposed;

        public IfcMappingGuard(string mappingFilePath)
        {
            _mappingFilePath = mappingFilePath ?? throw new ArgumentNullException(nameof(mappingFilePath));

            if (!File.Exists(_mappingFilePath))
                throw new FileNotFoundException("Файл мэппинга не найден", _mappingFilePath);

            // 🔐 Вычисляем хеш оригинала
            _originalHash = ComputeFileHash(_mappingFilePath);
            Logger.Debug("IfcMappingGuard", $"🔐 Original hash: {_originalHash}");

            // 💾 Создаём резервную копию (в Temp, чтобы не мешать пользователю)
            _backupPath = Path.Combine(
                Path.GetTempPath(),
                "IfcExportGuard",
                $"mapping_backup_{Guid.NewGuid():N}.txt"
            );
            Directory.CreateDirectory(Path.GetDirectoryName(_backupPath));
            File.Copy(_mappingFilePath, _backupPath, true);
            Logger.Debug("IfcMappingGuard", $"💾 Backup created: {_backupPath}");
        }

        /// <summary>
        /// Проверка и восстановление файла после экспорта.
        /// Возвращает true, если файл был изменён и восстановлен (требуется повторный экспорт).
        /// </summary>
        public bool VerifyAndRestore()
        {
            if (_isDisposed || !File.Exists(_mappingFilePath))
                return false;

            // Проверяем только первый экспорт в сессии
            if (!_isFirstExport)
                return false;

            _isFirstExport = false;

            var currentHash = ComputeFileHash(_mappingFilePath);
            Logger.Debug("IfcMappingGuard", $"🔍 Current hash: {currentHash}");

            if (currentHash != _originalHash)
            {
                Logger.Warning("IfcMappingGuard",
                    $"⚠️ Mapping file was MODIFIED by exporter! Restoring original...");

                RestoreOriginal();
                return true; // Сигнал: нужен повторный экспорт
            }

            Logger.Debug("IfcMappingGuard", "✅ Mapping file unchanged");
            if (currentHash != _originalHash)
            {
                Logger.Warning("IfcMappingGuard", $"⚠️ Mapping file was MODIFIED!");

                // 📋 Дамп первых 5 строк для быстрой диагностики
                try
                {
                    var lines = File.ReadLines(_mappingFilePath).Take(5);
                    Logger.Debug("IfcMappingGuard", "📄 Current file preview:\n" + string.Join("\n", lines));
                }
                catch { }

                RestoreOriginal();
                return true;
            }
            return false;
        }

        /// <summary>
        /// Принудительное восстановление оригинала (для внешнего использования).
        /// </summary>
        public void ForceRestore()
        {
            if (_isDisposed) return;
            RestoreOriginal();
        }

        private void RestoreOriginal()
        {
            try
            {
                if (File.Exists(_backupPath))
                {
                    File.Copy(_backupPath, _mappingFilePath, true);
                    File.SetLastWriteTime(_mappingFilePath, File.GetLastWriteTime(_backupPath));
                    Logger.Info("IfcMappingGuard", "🔄 Mapping file restored from backup");
                }
            }
            catch (Exception ex)
            {
                Logger.Error("IfcMappingGuard", $"❌ Failed to restore mapping: {ex.Message}");
            }
        }

        private static string ComputeFileHash(string path)
        {
            using (var md5 = MD5.Create())
            using (var stream = File.OpenRead(path))
            {
                var hash = md5.ComputeHash(stream);
                return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
            }
        }

        #region IDisposable
        public void Dispose()
        {
            if (_isDisposed) return;

            try
            {
                // 🧹 Удаляем бэкап после завершения работы
                if (File.Exists(_backupPath))
                    File.Delete(_backupPath);
                Logger.Debug("IfcMappingGuard", "🗑️ Cleanup: backup removed");
            }
            catch { /* Игнорируем ошибки очистки */ }

            _isDisposed = true;
        }
        #endregion
    }
}