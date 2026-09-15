using System;
using NLog;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Automation
{
    /// <summary>
    /// Wiring (design D13 addendum, PR5e hotfix follow-up): after a compile-error dialog is
    /// dismissed, the VBA project is left in a break state that no amount of dialog dismissal
    /// can clear -- the VBE requires an explicit "Reset" command. This executes that command
    /// through Rubberduck's own <see cref="ICommandBars"/> wrappers instead of the live UI,
    /// exactly the way the equivalent hotkey/menu click would.
    /// </summary>
    /// <remarks>
    /// Deliberately logic-free plumbing over <see cref="VbeMenuCaptionMatcher"/> (the actual,
    /// unit-tested decision logic): this class only locates and executes the native control.
    /// The standard VBE "Reset" control id (228) is tried first and its caption verified before
    /// executing it; if that does not resolve to a control whose caption contains "Reset", every
    /// top-level command bar's "Run" popup is enumerated and its children matched by caption
    /// instead. Per the design's own Test Strategy for this kind of native-automation glue, this
    /// class is wiring-exempt and verified only by a real-Excel smoke check.
    /// </remarks>
    public sealed class VbeProjectResetter
    {
        private const int ResetControlId = 228;
        private const string RunMenuCaption = "Run";
        private const string ResetCaption = "Reset";

        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private readonly IVBE _vbe;

        public VbeProjectResetter(IVBE vbe)
        {
            _vbe = vbe ?? throw new ArgumentNullException(nameof(vbe));
        }

        /// <summary>
        /// Executes the VBE "Reset" command if (and only if) the VBE is currently NOT fully in
        /// design mode -- i.e. some project is still in break mode. A no-op otherwise, and fully
        /// exception-safe: a failure here must never escape into the caller (design D13: a
        /// failure to reset must not itself crash or hang the run any worse than it already is).
        /// </summary>
        public void ResetIfInBreakMode()
        {
            try
            {
                if (_vbe.IsInDesignMode)
                {
                    Logger.Trace("VbeProjectResetter: VBE is already fully in design mode; nothing to reset.");
                    return;
                }

                using (var resetControl = FindResetControl())
                {
                    if (resetControl == null || resetControl.IsWrappingNullReference)
                    {
                        Logger.Warn("VbeProjectResetter: could not locate the VBE 'Reset' command; the project may remain in break mode.");
                        return;
                    }

                    Logger.Info("VbeProjectResetter: executing '{0}' (control id {1}) to reset the VBA project.", resetControl.Caption, resetControl.Id);
                    resetControl.Execute();
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(ex, "VbeProjectResetter failed while trying to reset the VBA project; it may remain in break mode.");
            }
        }

        // Deliberately does NOT wrap the collections in `using` blocks that would dispose them
        // on the same return path as a found control: the found control is a distinct COM
        // interface pointer from its parent collection/popup, but disposing the parent on the
        // way out is unnecessary risk in COM interop for zero benefit -- collections/popups that
        // do NOT yield a match are disposed explicitly instead.
        private ICommandBarControl FindResetControl()
        {
            var commandBars = _vbe.CommandBars;

            var byId = commandBars.FindControl(ResetControlId);
            if (!byId.IsWrappingNullReference && VbeMenuCaptionMatcher.ContainsWord(byId.Caption, ResetCaption))
            {
                Logger.Info("VbeProjectResetter: found the Reset control by the standard control id {0} (caption '{1}').", ResetControlId, byId.Caption);
                commandBars.Dispose();
                return byId;
            }

            Logger.Debug(
                "VbeProjectResetter: control id {0} did not resolve to a 'Reset' caption (caption '{1}'); falling back to enumerating the Run menu.",
                ResetControlId, byId.IsWrappingNullReference ? "<none>" : byId.Caption);
            byId.Dispose();

            var found = FindResetControlInRunMenu(commandBars);
            commandBars.Dispose();
            return found;
        }

        private static ICommandBarControl FindResetControlInRunMenu(ICommandBars commandBars)
        {
            foreach (var bar in commandBars)
            {
                var runPopup = FindRunPopup(bar.Controls);
                if (runPopup == null)
                {
                    bar.Dispose();
                    continue;
                }

                foreach (var control in runPopup.Controls)
                {
                    Logger.Debug("VbeProjectResetter: Run menu control id {0}, caption '{1}'.", control.Id, control.Caption);
                    if (VbeMenuCaptionMatcher.Matches(control.Caption, ResetCaption))
                    {
                        Logger.Info("VbeProjectResetter: found the Reset control in the Run menu (id {0}, caption '{1}').", control.Id, control.Caption);
                        runPopup.Dispose();
                        bar.Dispose();
                        return control;
                    }

                    control.Dispose();
                }

                runPopup.Dispose();
                bar.Dispose();
            }

            return null;
        }

        private static ICommandBarPopup FindRunPopup(ICommandBarControls controls)
        {
            foreach (var control in controls)
            {
                if (control.Type == ControlType.Popup && VbeMenuCaptionMatcher.Matches(control.Caption, RunMenuCaption))
                {
                    return control as ICommandBarPopup;
                }

                control.Dispose();
            }

            return null;
        }
    }
}
