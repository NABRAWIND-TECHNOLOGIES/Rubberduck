using System;
using System.Diagnostics;
using System.Globalization;
using Path = System.IO.Path;

namespace Rubberduck.Resources
{
    public static class ApplicationConstants
    {
        public static readonly string RUBBERDUCK_FOLDER_PATH = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Rubberduck");
        public static readonly string LOG_FOLDER_PATH = Path.Combine(RUBBERDUCK_FOLDER_PATH, "Logs");

        // Per-process (design D9 item 1): a second Excel instance starting mid-parse must not
        // sweep or collide with the first instance's exported sources. Unconditional for every
        // process, GUI included -- the GUI never observes another process's temp directory.
        public static readonly string RUBBERDUCK_TEMP_PATH = ComposeTempPath(Path.GetTempPath(), Process.GetCurrentProcess().Id);

        /// <summary>
        /// Pure composition of the per-process temp path, so the "one directory per pid" rule is
        /// unit-testable without a real running process.
        /// </summary>
        public static string ComposeTempPath(string tempRoot, int pid) =>
            Path.Combine(tempRoot, "Rubberduck", pid.ToString(CultureInfo.InvariantCulture));
    }
}
