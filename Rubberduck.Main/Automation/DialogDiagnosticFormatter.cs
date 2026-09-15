using Rubberduck.UnitTesting;

namespace Rubberduck.Automation
{
    /// <summary>
    /// Pure message/diagnostic formatting for a classified modal dialog (design D13). A compile
    /// error whose source location was successfully captured gets the precise
    /// "at &lt;Module&gt;:&lt;line&gt;:&lt;col&gt;" format; every other case -- a non-compile
    /// dialog, or a compile error whose location could not be read -- fails safe to the plain
    /// caption/text format.
    /// </summary>
    public static class DialogDiagnosticFormatter
    {
        public static string FormatMessage(DialogKind kind, string caption, string text, string module, int? line, int? column, string source)
        {
            var normalizedText = text ?? string.Empty;

            if (kind == DialogKind.CompileError && HasLocation(module, line, column, source))
            {
                return $"VBA compile error: {normalizedText} at {module}:{line.Value}:{column.Value} — {source}";
            }

            return $"VBA dialog: {caption ?? string.Empty} — {normalizedText}";
        }

        public static HeadlessDialogDiagnosticInfo BuildDiagnostic(DialogKind kind, string caption, string text, string module, int? line, int? column, string source)
        {
            var isCompileErrorWithLocation = kind == DialogKind.CompileError && HasLocation(module, line, column, source);
            var diagnosticKind = isCompileErrorWithLocation ? "compileError" : "dialog";

            return new HeadlessDialogDiagnosticInfo(
                diagnosticKind,
                isCompileErrorWithLocation ? module : null,
                isCompileErrorWithLocation ? line : null,
                isCompileErrorWithLocation ? column : null,
                isCompileErrorWithLocation ? source : null,
                caption,
                text);
        }

        private static bool HasLocation(string module, int? line, int? column, string source) =>
            module != null && line.HasValue && column.HasValue && source != null;
    }
}
