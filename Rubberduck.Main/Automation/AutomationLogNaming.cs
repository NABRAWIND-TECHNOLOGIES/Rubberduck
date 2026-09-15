namespace Rubberduck.Automation
{
    /// <summary>
    /// Pure per-process NLog filename suffixing for automation mode (design D9 item 2). The
    /// interactive GUI resolves its <c>automationSuffix</c> NLog variable to the empty string,
    /// so the stock <c>RubberduckLog.txt</c> filename is produced unchanged - the conditional is
    /// fail-safe toward stock behaviour by construction.
    /// </summary>
    public static class AutomationLogNaming
    {
        /// <summary>
        /// The NLog <c>automationSuffix</c> variable value: <c>".&lt;pid&gt;"</c> while automation
        /// mode is active, or the empty string otherwise.
        /// </summary>
        public static string Suffix(bool isActive, int pid) => isActive ? $".{pid}" : string.Empty;

        /// <summary>
        /// Inserts <paramref name="suffix"/> immediately before the final extension of
        /// <paramref name="fileName"/>. Applies identically to both the live log filename and the
        /// archive filename pattern, matching how the same <c>${var:automationSuffix}</c>
        /// interpolation is used for both targets in <c>NLog.dll.nlog</c>.
        /// </summary>
        public static string Apply(string fileName, string suffix)
        {
            if (string.IsNullOrEmpty(fileName) || string.IsNullOrEmpty(suffix))
            {
                return fileName;
            }

            var extension = System.IO.Path.GetExtension(fileName);
            var nameWithoutExtension = fileName.Substring(0, fileName.Length - extension.Length);
            return nameWithoutExtension + suffix + extension;
        }
    }
}
