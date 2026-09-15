using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using NLog;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Automation
{
    /// <summary>
    /// Thread-scoped WH_CBT hook (design D13) that dismisses a VBE modal dialog ("Compile
    /// error", a <c>MsgBox</c>, or a runtime Debug/End prompt) before it can ever appear on
    /// screen or hang a headless run. Installed by <c>Extension.InitializeAddIn</c> only while
    /// <see cref="Automation.AutomationMode.Current"/>.IsActive is true, and disposed at
    /// shutdown -- a GUI session never installs this hook.
    /// </summary>
    /// <remarks>
    /// Deliberately logic-free plumbing over <see cref="DialogClassifier"/>/
    /// <see cref="DialogButtonSelector"/>/<see cref="DialogDiagnosticFormatter"/>/
    /// <see cref="HeadlessDialogLog"/> (which carry the actual, unit-tested decision logic): this
    /// class only locates/reads/dismisses the native window. Per the design's own Test Strategy,
    /// the hook installation itself is wiring-exempt and verified only by task 9.6's smoke check
    /// against a real modal dialog.
    ///
    /// PR5e hotfix root cause: the original dismissal assumed the VBE's own dialogs use the
    /// well-known <c>IDOK</c>/<c>IDCANCEL</c> control ids (1/2), the way a standard
    /// <c>MessageBox</c> call does. They do not -- the VBE's compile-error and runtime-error
    /// prompts are hand-authored resources in vbe7.dll with their own arbitrary control ids, so
    /// <c>GetDlgItem(hwnd, IDOK)</c> never found a button and the dialog was left on screen.
    /// Dismissal now enumerates every child window and selects the button by its caption via
    /// <see cref="DialogButtonSelector"/> instead.
    /// </remarks>
    public sealed class HeadlessDialogInterceptor : IDisposable
    {
        private const int WH_CBT = 5;
        private const int HCBT_CREATEWND = 3;
        private const int HCBT_ACTIVATE = 5;
        private const string DialogClassName = "#32770";
        private const string ButtonClassName = "Button";
        private const string StaticClassName = "Static";
        private const string EditClassName = "Edit";
        private const int WM_COMMAND = 0x0111;
        private const int WM_CLOSE = 0x0010;
        private const int WM_GETTEXT = 0x000D;
        private const int WM_GETTEXTLENGTH = 0x000E;
        private const int BM_CLICK = 0x00F5;
        private const int IDOK = 1;

        private delegate IntPtr HookProc(int code, IntPtr wParam, IntPtr lParam);
        private delegate bool EnumChildProc(IntPtr hwnd, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, int dwThreadId);

        [DllImport("user32.dll")]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern IntPtr GetDlgItem(IntPtr hDlg, int nIDDlgItem);

        [DllImport("user32.dll")]
        private static extern int GetDlgCtrlID(IntPtr hWndCtl);

        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr hwndParent, EnumChildProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, StringBuilder lParam);

        [DllImport("user32.dll")]
        private static extern bool PostMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private readonly IVBE _vbe;
        private readonly HeadlessDialogLog _log;
        private readonly Func<DateTime> _clock;
        private readonly int _threadId;
        private readonly int _processId;

        // Kept alive for the interceptor's lifetime: the CLR must not collect the delegate
        // while native code still holds a pointer to it (a classic P/Invoke marshaling pitfall).
        private HookProc _hookProc;
        private IntPtr _hookHandle;

        public HeadlessDialogInterceptor(IVBE vbe, HeadlessDialogLog log = null, Func<DateTime> clock = null)
        {
            _vbe = vbe ?? throw new ArgumentNullException(nameof(vbe));
            _log = log ?? HeadlessDialogLog.Shared;
            _clock = clock ?? (() => DateTime.UtcNow);
            _threadId = NativeThreadId();
            _processId = Process.GetCurrentProcess().Id;
        }

        public void Install()
        {
            if (_hookHandle != IntPtr.Zero)
            {
                return;
            }

            _hookProc = HookCallback;
            _hookHandle = SetWindowsHookEx(WH_CBT, _hookProc, IntPtr.Zero, _threadId);

            if (_hookHandle == IntPtr.Zero)
            {
                var error = Marshal.GetLastWin32Error();
                Logger.Warn(
                    "HeadlessDialogInterceptor FAILED to install its WH_CBT hook for thread {0} (process {1}); Win32 error {2}. Modal dialogs will NOT be intercepted.",
                    _threadId, _processId, error);
            }
            else
            {
                Logger.Info(
                    "HeadlessDialogInterceptor installed WH_CBT hook 0x{0:X} for thread {1} (process {2}).",
                    _hookHandle.ToInt64(), _threadId, _processId);
            }
        }

        public void Dispose()
        {
            if (_hookHandle == IntPtr.Zero)
            {
                return;
            }

            UnhookWindowsHookEx(_hookHandle);
            _hookHandle = IntPtr.Zero;
            _hookProc = null;
        }

        private IntPtr HookCallback(int code, IntPtr wParam, IntPtr lParam)
        {
            // The hook callback must never let an exception escape into the message loop
            // (design: "keep the hook callback minimal and exception-free").
            try
            {
                if (code == HCBT_CREATEWND)
                {
                    LogDialogWindowCreated(wParam);
                }
                else if (code == HCBT_ACTIVATE)
                {
                    HandleActivatedWindow(wParam);
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(ex, "HeadlessDialogInterceptor hook callback failed; the dialog, if any, was left untouched.");
            }

            return CallNextHookEx(_hookHandle, code, wParam, lParam);
        }

        private void LogDialogWindowCreated(IntPtr hwnd)
        {
            if (!IsDialogClass(hwnd))
            {
                return;
            }

            GetWindowThreadProcessId(hwnd, out var ownerPid);
            Logger.Debug(
                "HeadlessDialogInterceptor observed HCBT_CREATEWND for a {0} window 0x{1:X} (owner pid {2}).",
                DialogClassName, hwnd.ToInt64(), ownerPid);
        }

        private void HandleActivatedWindow(IntPtr hwnd)
        {
            GetWindowThreadProcessId(hwnd, out var ownerPid);
            if (ownerPid != (uint)_processId)
            {
                Logger.Trace(
                    "HeadlessDialogInterceptor ignored HCBT_ACTIVATE for window 0x{0:X}: owned by pid {1}, not this process ({2}).",
                    hwnd.ToInt64(), ownerPid, _processId);
                return;
            }

            if (!IsDialogClass(hwnd))
            {
                return;
            }

            var caption = GetWindowTextSafe(hwnd);
            var children = EnumerateChildren(hwnd);
            var text = FindDialogText(hwnd, children);
            var buttons = children.Where(c => c.ClassName == ButtonClassName).ToList();

            Logger.Info(
                "HeadlessDialogInterceptor observed a {0} modal dialog 0x{1:X} (caption '{2}', text '{3}', {4} children, buttons: [{5}]).",
                DialogClassName, hwnd.ToInt64(), caption, text, children.Count,
                string.Join(", ", buttons.Select(b => $"{b.ControlId}:'{b.Text}'")));

            var kind = DialogClassifier.Classify(caption, text);
            Logger.Info("HeadlessDialogInterceptor classified dialog 0x{0:X} as {1}.", hwnd.ToInt64(), kind);

            string module = null;
            int? line = null;
            int? column = null;
            string source = null;
            if (kind == DialogKind.CompileError)
            {
                TryReadCompileErrorLocation(out module, out line, out column, out source);
            }

            _log.Append(new HeadlessDialogEntry(caption, text, module, line, column, source, _clock()));

            Dismiss(hwnd, DialogClassifier.ActionFor(kind), buttons);
        }

        private static bool IsDialogClass(IntPtr hwnd)
        {
            var className = new StringBuilder(64);
            GetClassName(hwnd, className, className.Capacity);
            return className.ToString() == DialogClassName;
        }

        private void TryReadCompileErrorLocation(out string module, out int? line, out int? column, out string source)
        {
            module = null;
            line = null;
            column = null;
            source = null;

            try
            {
                using (var codePane = _vbe.ActiveCodePane)
                {
                    if (codePane is null || codePane.IsWrappingNullReference)
                    {
                        return;
                    }

                    var selection = codePane.Selection;
                    using (var codeModule = codePane.CodeModule)
                    using (var component = codeModule.Parent)
                    {
                        module = component.Name;
                        line = selection.StartLine;
                        column = selection.StartColumn;
                        source = codeModule.GetLines(selection.StartLine, 1);
                    }
                }
            }
            catch (Exception ex)
            {
                // Fail safe to "no location" (design D13) -- every COM read here can throw if
                // the VBE is in a state the dialog itself put it in.
                Logger.Trace(ex, "Could not read the active code pane's selection for a compile-error dialog; reporting no location.");
                module = null;
                line = null;
                column = null;
                source = null;
            }
        }

        private void Dismiss(IntPtr hwnd, DialogDismissAction action, IReadOnlyList<ChildWindowInfo> buttons)
        {
            var candidates = buttons
                .Select(b => new DialogButtonCandidate(b.ControlId, b.Text))
                .ToList();

            var selectedControlId = DialogButtonSelector.SelectButton(action, candidates);
            if (selectedControlId.HasValue)
            {
                var button = buttons.FirstOrDefault(b => b.ControlId == selectedControlId.Value);
                if (button.Handle != IntPtr.Zero)
                {
                    Logger.Info(
                        "HeadlessDialogInterceptor dismissing dialog 0x{0:X} via BM_CLICK on control id {1} ('{2}').",
                        hwnd.ToInt64(), button.ControlId, button.Text);
                    SendMessage(button.Handle, BM_CLICK, IntPtr.Zero, IntPtr.Zero);
                    return;
                }
            }

            // Fallback 1: the classic MessageBox IDOK id, in case this dialog variant DOES use
            // the standard id but wasn't picked up as a "Button"-class child for some reason.
            var fallbackButton = GetDlgItem(hwnd, IDOK);
            if (action == DialogDismissAction.Ok && fallbackButton != IntPtr.Zero)
            {
                Logger.Info(
                    "HeadlessDialogInterceptor dismissing dialog 0x{0:X} via WM_COMMAND on fallback IDOK (no captioned button matched).",
                    hwnd.ToInt64());
                PostMessage(hwnd, WM_COMMAND, (IntPtr)IDOK, IntPtr.Zero);
                return;
            }

            // Fallback 2: WM_CLOSE always dismisses a modal dialog, even one with no system menu.
            Logger.Info(
                "HeadlessDialogInterceptor dismissing dialog 0x{0:X} via WM_CLOSE fallback (action {1}, no matching button found).",
                hwnd.ToInt64(), action);
            PostMessage(hwnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
        }

        private struct ChildWindowInfo
        {
            public IntPtr Handle;
            public int ControlId;
            public string ClassName;
            public string Text;
        }

        private static List<ChildWindowInfo> EnumerateChildren(IntPtr hwnd)
        {
            var children = new List<ChildWindowInfo>();

            EnumChildWindows(hwnd, (childHwnd, _) =>
            {
                var classNameBuffer = new StringBuilder(64);
                GetClassName(childHwnd, classNameBuffer, classNameBuffer.Capacity);
                children.Add(new ChildWindowInfo
                {
                    Handle = childHwnd,
                    ControlId = GetDlgCtrlID(childHwnd),
                    ClassName = classNameBuffer.ToString(),
                    Text = GetChildTextSafe(childHwnd)
                });
                return true;
            }, IntPtr.Zero);

            return children;
        }

        private static string FindDialogText(IntPtr hwnd, IReadOnlyList<ChildWindowInfo> children)
        {
            var messageControl = children.FirstOrDefault(c =>
                (c.ClassName == StaticClassName || c.ClassName == EditClassName) && !string.IsNullOrEmpty(c.Text));

            if (!string.IsNullOrEmpty(messageControl.Text))
            {
                return messageControl.Text;
            }

            // Fallback: the fixed id-range read some VBE dialog variants still use.
            return GetDialogTextSafe(hwnd);
        }

        private static string GetChildTextSafe(IntPtr hwnd)
        {
            var length = SendMessage(hwnd, WM_GETTEXTLENGTH, IntPtr.Zero, IntPtr.Zero).ToInt32();
            if (length <= 0)
            {
                return string.Empty;
            }

            var buffer = new StringBuilder(length + 1);
            SendMessage(hwnd, WM_GETTEXT, (IntPtr)buffer.Capacity, buffer);
            return buffer.ToString();
        }

        private static string GetWindowTextSafe(IntPtr hwnd)
        {
            var buffer = new StringBuilder(256);
            GetWindowText(hwnd, buffer, buffer.Capacity);
            return buffer.ToString();
        }

        // Fallback used only when no "Static"/"Edit" child carries the message (see
        // FindDialogText): the VBE's modal dialogs surface their message in a static/edit child
        // control; control id 0xFFFF (-1) is not a reliable constant across dialog variants, so
        // this walks the handful of low child ids the classic VBE dialogs use rather than
        // enumerating every child window for a single line of text.
        private static string GetDialogTextSafe(IntPtr hwnd)
        {
            for (var childId = 0xFFFF; childId >= 0xFFF0; childId--)
            {
                var child = GetDlgItem(hwnd, childId);
                if (child == IntPtr.Zero)
                {
                    continue;
                }

                var text = GetChildTextSafe(child);
                if (text.Length > 0)
                {
                    return text;
                }
            }

            return string.Empty;
        }

        [DllImport("kernel32.dll")]
        private static extern int GetCurrentThreadId();

        private static int NativeThreadId() => GetCurrentThreadId();
    }
}
