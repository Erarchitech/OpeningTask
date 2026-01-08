using System;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;

namespace OpeningTask.Helpers
{
    internal static class RevitTrace
    {
        private static readonly object _lock = new object();
        private static string _overrideLogPath;

        public static void SetLogPath(string path)
        {
            _overrideLogPath = path;
        }

        public static string LogPath
        {
            get
            {
                if (!string.IsNullOrWhiteSpace(_overrideLogPath))
                    return _overrideLogPath;

                var fixedDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "Autodesk",
                    "Revit",
                    "Addins",
                    "2024");

                if (!string.IsNullOrWhiteSpace(fixedDir))
                    return Path.Combine(fixedDir, "OpeningTask.trace.log");

                try
                {
                    var asm = Assembly.GetExecutingAssembly();
                    var asmPath = asm.Location;
                    var dir = Path.GetDirectoryName(asmPath);
                    if (string.IsNullOrWhiteSpace(dir))
                        dir = Environment.CurrentDirectory;

                    return Path.Combine(dir, "OpeningTask.trace.log");
                }
                catch
                {
                    return Path.Combine(Environment.CurrentDirectory, "OpeningTask.trace.log");
                }
            }
        }

        public static void Info(string message) => Write("INF", message, null);

        public static void Warn(string message) => Write("WRN", message, null);

        public static void Error(string message, Exception ex = null) => Write("ERR", message, ex);

        public static void Flush() { /* no-op for now */ }

        private static void Write(string level, string message, Exception ex)
        {
            var ts = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture);
            var pid = 0;
            try { pid = System.Diagnostics.Process.GetCurrentProcess().Id; } catch { }

            var sb = new StringBuilder();
            sb.Append(ts).Append(" [").Append(level).Append("] ");
            sb.Append("pid=").Append(pid).Append(" ");
            sb.Append(message);

            if (ex != null)
            {
                sb.Append(" | ex=").Append(ex.GetType().FullName).Append(": ").Append(ex.Message);
                if (!string.IsNullOrWhiteSpace(ex.StackTrace))
                    sb.Append(" | st=").Append(ex.StackTrace.Replace(Environment.NewLine, " "));
            }

            var line = sb.ToString();

            lock (_lock)
            {
                try
                {
                    var path = LogPath;
                    TryAppend(path, line);
                }
                catch
                {
                    // ignore any logging failures to avoid affecting Revit
                }
            }
        }

        private static void TryAppend(string path, string line)
        {
            try
            {
                var dir = Path.GetDirectoryName(path);
                if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                File.AppendAllText(path, line + Environment.NewLine, Encoding.UTF8);
                return;
            }
            catch
            {
                // fall through
            }

            // fallback: рядом с dll (или current directory)
            try
            {
                var asmPath = Assembly.GetExecutingAssembly().Location;
                var asmDir = Path.GetDirectoryName(asmPath);
                if (string.IsNullOrWhiteSpace(asmDir))
                    asmDir = Environment.CurrentDirectory;

                var fallbackPath = Path.Combine(asmDir, "OpeningTask.trace.log");
                var fallbackDir = Path.GetDirectoryName(fallbackPath);
                if (!string.IsNullOrWhiteSpace(fallbackDir) && !Directory.Exists(fallbackDir))
                    Directory.CreateDirectory(fallbackDir);

                File.AppendAllText(fallbackPath, line + Environment.NewLine, Encoding.UTF8);
            }
            catch
            {
                // ignore
            }
        }
    }
}
