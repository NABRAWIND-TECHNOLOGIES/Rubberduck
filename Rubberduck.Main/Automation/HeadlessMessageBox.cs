using System;
using Rubberduck.Interaction;

namespace Rubberduck.Automation
{
    /// <summary>
    /// Non-interactive <see cref="IMessageBox"/> bound under automation (design D8/"Headless UI
    /// Suppression"): no dialog is ever shown, so a headless Excel process can never block on a
    /// user who is not present. Every call resolves to the caller's own stated default/suggestion
    /// instead of hard-coding a single answer, and is logged so the suppressed interaction is
    /// still auditable in the per-process automation log (design D9 item 2).
    /// </summary>
    public sealed class HeadlessMessageBox : IMessageBox
    {
        private readonly Action<string> _log;

        /// <param name="log">Sink for the suppressed-interaction record; kept injectable so the
        /// decision logic here is testable without NLog or any real logger.</param>
        public HeadlessMessageBox(Action<string> log)
        {
            _log = log ?? throw new ArgumentNullException(nameof(log));
        }

        public void Message(string text) =>
            _log($"[automation] message suppressed: {text}");

        public void NotifyWarn(string text, string caption) =>
            _log($"[automation] warning suppressed ({caption}): {text}");

        public bool Question(string text, string caption)
        {
            // No safe non-reversible default exists for an arbitrary yes/no question, so the
            // suppressed answer is always "No" -- the same fail-safe posture as "leave state
            // unchanged" (design: never show a dialog, never guess a destructive "yes").
            _log($"[automation] question suppressed ({caption}): {text} -> No");
            return false;
        }

        public bool ConfirmYesNo(string text, string caption) => ConfirmYesNo(text, caption, true);

        public bool ConfirmYesNo(string text, string caption, bool suggestion)
        {
            _log($"[automation] confirmation suppressed ({caption}): {text} -> {(suggestion ? "Yes" : "No")}");
            return suggestion;
        }

        public ConfirmationOutcome Confirm(string text, string caption, ConfirmationOutcome suggestion = ConfirmationOutcome.Cancel)
        {
            _log($"[automation] confirmation suppressed ({caption}): {text} -> {suggestion}");
            return suggestion;
        }
    }
}
