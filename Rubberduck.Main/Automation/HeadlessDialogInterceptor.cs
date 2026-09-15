using System;
using System.Diagnostics;
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
    /// <see cref="DialogDiagnosticFormatter"/>/<see cref="HeadlessDialogLog"/> (which carry the
    /// actual, unit-tested decision logic): this class only locates/reads/dismisses the native
    /// window. Per the design's own Test Strategy, the hook installation itself is wiring-exempt
    /// and verified only by task 9.6's smoke check against a real modal dialog.
    /// </remarks>
    public sealed class HeadlessDialogInterceptor : IDisposable
    {
        private const int WH_CBT = 5;
        private const int HCBT_ACTIVATE = 5;
        private const string DialogClassName = "#32770";
        private const int WM_COMMAND = 0x0111;
        private const int WM_CLOSE = 0x0010;
        private const int WM_GETTEXT = 0x000D;
        private const int WM_GETTEXTLENGTH = 0x000E;
        private const int IDOK = 1;
        private const int IDCANCEL = 2;

        private delegate IntPtr HookProc(int code, IntPtr wParam, IntPtr lParam);

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
                if (code == HCBT_ACTIVATE)
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

        private void HandleActivatedWindow(IntPtr hwnd)
        {
            GetWindowThreadProcessId(hwnd, out var ownerPid);
            if (ownerPid != (uint)_processId)
            {
                return;
            }

            var className = new StringBuilder(64);
            GetClassName(hwnd, className, className.Capacity);
            if (className.ToString() != DialogClassName)
            {
                return;
            }

            var caption = GetWindowTextSafe(hwnd);
            var text = GetDialogTextSafe(hwnd);
            var kind = DialogClassifier.Classify(caption, text);

            string module = null;
            int? line = null;
            int? column = null;
            string source = null;
            if (kind == DialogKind.CompileError)
            {
                TryReadCompileErrorLocation(out module, out line, out column, out source);
            }

            _log.Append(new HeadlessDialogEntry(caption, text, module, line, column, source, _clock()));

            Dismiss(hwnd, DialogClassifier.ActionFor(kind));
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

        private void Dismiss(IntPtr hwnd, DialogDismissAction action)
        {
            switch (action)
            {
                case DialogDismissAction.Ok:
                    ClickButtonOrFallBack(hwnd, IDOK);
                    break;
                case DialogDismissAction.EndOrCancel:
                    ClickButtonOrFallBack(hwnd, IDCANCEL);
                    break;
                default:
                    PostMessage(hwnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                    break;
            }
        }

        private void ClickButtonOrFallBack(IntPtr hwnd, int controlId)
        {
            var button = GetDlgItem(hwnd, controlId);
            if (button != IntPtr.Zero)
            {
                PostMessage(hwnd, WM_COMMAND, (IntPtr)controlId, IntPtr.Zero);
            }
            else
            {
                PostMessage(hwnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
            }
        }

        private static string GetWindowTextSafe(IntPtr hwnd)
        {
            var buffer = new StringBuilder(256);
            GetWindowText(hwnd, buffer, buffer.Capacity);
            return buffer.ToString();
        }

        // The VBE's modal dialogs surface their message in a static/edit child control; control
        // id 0xFFFF (-1) is not a reliable constant across dialog variants, so this walks the
        // handful of low child ids the classic VBE dialogs use rather than enumerating every
        // child window for a single line of text.
        private static string GetDialogTextSafe(IntPtr hwnd)
        {
            for (var childId = 0xFFFF; childId >= 0xFFF0; childId--)
            {
                var child = GetDlgItem(hwnd, childId);
                if (child == IntPtr.Zero)
                {
                    continue;
                }

                var length = SendMessage(child, WM_GETTEXTLENGTH, IntPtr.Zero, IntPtr.Zero).ToInt32();
                if (length <= 0)
                {
                    continue;
                }

                var buffer = new StringBuilder(length + 1);
                SendMessage(child, WM_GETTEXT, (IntPtr)buffer.Capacity, buffer);
                if (buffer.Length > 0)
                {
                    return buffer.ToString();
                }
            }

            return string.Empty;
        }

        [DllImport("kernel32.dll")]
        private static extern int GetCurrentThreadId();

        private static int NativeThreadId() => GetCurrentThreadId();
    }
}
