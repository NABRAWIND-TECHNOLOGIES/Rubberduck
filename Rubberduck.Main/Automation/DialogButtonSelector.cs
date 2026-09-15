using System.Collections.Generic;

namespace Rubberduck.Automation
{
    /// <summary>A single button child window found on an intercepted modal dialog.</summary>
    public struct DialogButtonCandidate
    {
        public DialogButtonCandidate(int controlId, string text)
        {
            ControlId = controlId;
            Text = text;
        }

        public int ControlId { get; }
        public string Text { get; }
    }

    /// <summary>
    /// Pure button-selection logic for a classified modal dialog (design D13 hotfix). The VBE's
    /// own dialogs are hand-authored resources, not standard <c>MessageBox</c> calls, so their
    /// button control ids are NOT guaranteed to be the well-known <c>IDOK</c>/<c>IDCANCEL</c>
    /// values (1/2) -- the only reliable signal is each button's own caption. Matching is
    /// case-insensitive, strips the Win32 accelerator marker (<c>&amp;</c>), and covers both
    /// English and the Spanish localization observed in this environment.
    /// </summary>
    public static class DialogButtonSelector
    {
        // Ordered by preference: the first matching text wins.
        private static readonly string[] OkTexts = { "ok", "aceptar" };

        // "End"/"Finalizar" only -- a Run-time error prompt's "Debug"/"Depurar" button must
        // never be selected, since clicking it would open the VBE debugger instead of ending
        // the run.
        private static readonly string[] EndTexts = { "end", "finalizar" };

        /// <summary>
        /// Returns the control id of the button to click for <paramref name="action"/>, or
        /// <c>null</c> when no matching button was found (or the action itself, e.g.
        /// <see cref="DialogDismissAction.Close"/>, never clicks a button) -- the caller falls
        /// back to <c>WM_CLOSE</c> in that case.
        /// </summary>
        public static int? SelectButton(DialogDismissAction action, IReadOnlyList<DialogButtonCandidate> candidates)
        {
            if (candidates == null || candidates.Count == 0)
            {
                return null;
            }

            switch (action)
            {
                case DialogDismissAction.Ok:
                    return FindByText(candidates, OkTexts);
                case DialogDismissAction.EndOrCancel:
                    return FindByText(candidates, EndTexts);
                default:
                    return null;
            }
        }

        private static int? FindByText(IReadOnlyList<DialogButtonCandidate> candidates, string[] wantedTextsInPreferenceOrder)
        {
            foreach (var wanted in wantedTextsInPreferenceOrder)
            {
                foreach (var candidate in candidates)
                {
                    if (Normalize(candidate.Text) == wanted)
                    {
                        return candidate.ControlId;
                    }
                }
            }

            return null;
        }

        private static string Normalize(string text) =>
            (text ?? string.Empty).Replace("&", string.Empty).Trim().ToLowerInvariant();
    }
}
