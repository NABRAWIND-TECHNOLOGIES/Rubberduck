using System;

namespace Rubberduck.Automation
{
    /// <summary>Kind of modal dialog intercepted under automation mode (design D13).</summary>
    public enum DialogKind
    {
        CompileError,
        RuntimeDebugEndPrompt,
        Unrecognized
    }

    /// <summary>Which control the interceptor must click to dismiss a classified dialog.</summary>
    public enum DialogDismissAction
    {
        /// <summary>OK -- e.g. a "Compile error" dialog.</summary>
        Ok,

        /// <summary>End -- e.g. the classic Run-time error End/Debug/Help prompt.</summary>
        EndOrCancel,

        /// <summary>WM_CLOSE fallback for anything not recognized as the two cases above.</summary>
        Close
    }

    /// <summary>
    /// Pure caption/text classification for a captured <see cref="HeadlessDialogEntry"/>
    /// (design D13). Both VBE dialog kinds recognized here are raised by vbe7.dll on the
    /// Excel STA thread; everything else (a plain <c>MsgBox</c>, an unrelated host prompt)
    /// falls back to the safe <see cref="DialogDismissAction.Close"/> action.
    /// </summary>
    public static class DialogClassifier
    {
        public static DialogKind Classify(string caption, string text)
        {
            var normalizedText = text ?? string.Empty;

            if (normalizedText.IndexOf("Compile error", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return DialogKind.CompileError;
            }

            if (normalizedText.IndexOf("Run-time error", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return DialogKind.RuntimeDebugEndPrompt;
            }

            return DialogKind.Unrecognized;
        }

        public static DialogDismissAction ActionFor(DialogKind kind)
        {
            switch (kind)
            {
                case DialogKind.CompileError:
                    return DialogDismissAction.Ok;
                case DialogKind.RuntimeDebugEndPrompt:
                    return DialogDismissAction.EndOrCancel;
                default:
                    return DialogDismissAction.Close;
            }
        }
    }
}
