using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace DaabNavisExport.Utilities
{
    internal sealed class ExportLog
    {
        private readonly string _filePath;
        private readonly List<string> _entries = new();
        private readonly object _syncRoot = new();
        private readonly Stopwatch _stopwatch;

        public ExportLog(string filePath)
        {
            _filePath = filePath;
            _stopwatch = Stopwatch.StartNew();
            Info($"Export session started. Log file: {_filePath}.");
        }

        public string FilePath => _filePath;

        public void Info(string message) => Append("INFO", message);

        public void Warn(string message) => Append("WARN", message);

        public void Error(string message) => Append("ERROR", message);

        public void AppendRaw(string message)
        {
            lock (_syncRoot)
            {
                _entries.Add(message);
            }

            Debug.WriteLine(message);
        }

        private void Append(string level, string message)
        {
            var formatted = $"{DateTime.Now:O} [{level}] {message}";
            AppendRaw(formatted);
        }

        public void Complete()
        {
            _stopwatch.Stop();
            Append("INFO", $"Export session finished in {_stopwatch.Elapsed}.");

            try
            {
                var directory = Path.GetDirectoryName(_filePath);
                if (!string.IsNullOrEmpty(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                lock (_syncRoot)
                {
                    File.WriteAllLines(_filePath, _entries, Encoding.UTF8);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to write export log to {_filePath}: {ex.Message}");
            }
        }
    }
}
