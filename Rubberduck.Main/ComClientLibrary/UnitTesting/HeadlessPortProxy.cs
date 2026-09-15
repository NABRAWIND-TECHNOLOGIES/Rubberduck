using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using Rubberduck.Resources.Registration;

namespace Rubberduck.UnitTesting
{
    /// <summary>
    /// Hotfix (PR5f, design D5/P7 amendment): COM-visible placeholder implementing
    /// <see cref="IRubberduckTestRunner"/>, assigned to the add-in's <c>Object</c> slot inside
    /// <c>_Extension.OnConnection</c> instead of the real <see cref="RubberduckTestRunner"/>.
    /// </summary>
    /// <remarks>
    /// Root cause this class fixes: the VBE only accepts an <c>AddIn.Object</c> assignment while
    /// execution is inside <c>OnConnection</c>. The original design assigned the real runner in
    /// <c>Startup()</c> instead (called from <c>OnStartupComplete</c>, i.e. after
    /// <c>OnConnection</c> has already returned), which throws
    /// <c>COMException E_FAIL</c> from <c>AddIn.set_Object</c> on the normal GUI startup path and
    /// broke every interactive Excel session. The VBE never rejects re-pointing an
    /// <em>already-assigned</em> <c>.Object</c>'s own members, so this proxy is assigned once,
    /// inside <c>OnConnection</c>, and <see cref="Bind"/> is called later from
    /// <c>Startup()</c> to attach the real runner behind it -- no second <c>.Object</c>
    /// assignment ever happens outside <c>OnConnection</c>.
    ///
    /// Before <see cref="Bind"/>, every member returns a safe, non-throwing value so a headless
    /// client polling the port before startup completes gets a well-formed answer instead of a
    /// COM exception or a hang. After <see cref="Bind"/>, every member delegates straight to the
    /// bound target. No locks: like <see cref="RubberduckTestRunner"/>, every call arrives on the
    /// single Excel STA main thread (design "Threading model and thread safety").
    /// </remarks>
    [
        ComVisible(true),
        Guid(RubberduckGuid.TestRunnerProxyGuid),
        ProgId(RubberduckProgId.TestRunnerProxyProgId),
        ClassInterface(ClassInterfaceType.None),
        ComDefaultInterface(typeof(IRubberduckTestRunner)),
        EditorBrowsable(EditorBrowsableState.Always)
    ]
    public class HeadlessPortProxy : IRubberduckTestRunner
    {
        private IRubberduckTestRunner _target;
        private HeadlessErrorInfo? _preBindError;

        /// <summary>
        /// Attaches the real runner once <c>Startup()</c> has resolved it from the IoC
        /// container. From this point on every member delegates to <paramref name="target"/>.
        /// </summary>
        public void Bind(IRubberduckTestRunner target)
        {
            _target = target ?? throw new ArgumentNullException(nameof(target));
        }

        /// <summary>
        /// Records that <c>Startup()</c> threw before <see cref="Bind"/> was ever called, so a
        /// headless client sees a structured <c>STARTUP_FAILED</c> error instead of a proxy that
        /// silently stays in its pre-bind "pending" state forever.
        /// </summary>
        public void NotifyStartupFailed(string detail)
        {
            if (_target != null)
            {
                // The real runner is already bound and owns its own error state; a stale
                // failure notification racing in behind it must not shadow that.
                return;
            }

            _preBindError = new HeadlessErrorInfo("STARTUP_FAILED", "Rubberduck's startup sequence failed.", detail);
        }

        public string Version => _target?.Version ?? HeadlessRunnerVersion.Current;

        public int SchemaVersion => _target?.SchemaVersion ?? TestResultJsonSerializer.SchemaVersion;

        public bool IsReady => _target?.IsReady ?? false;

        public string ParserStatus => _target?.ParserStatus ?? "StartupPending";

        public void RequestParse() => _target?.RequestParse();

        public int DiscoveredTestCount => _target?.DiscoveredTestCount ?? 0;

        public bool IsComplete => _target?.IsComplete ?? false;

        public string StartRun(string ModuleName = "", string MethodName = "")
        {
            if (_target != null)
            {
                return _target.StartRun(ModuleName, MethodName);
            }

            _preBindError = new HeadlessErrorInfo("NOT_READY", "Rubberduck startup has not completed", null);
            return string.Empty;
        }

        public void Cancel() => _target?.Cancel();

        public string GetResults()
        {
            if (_target != null)
            {
                return _target.GetResults();
            }

            // Mirrors RubberduckTestRunner's own pre-run envelope exactly (Idle state, "idle"
            // status, empty selection, zero counts) so a client parsing this before Bind() sees
            // the same shape it would see from the real runner before any StartRun call.
            var run = new HeadlessRunInfo(string.Empty, string.Empty, new HeadlessSelectionInfo("all", null, null), default, default);
            return TestResultJsonSerializer.Serialize(Version, "idle", run, Array.Empty<HeadlessTestResultRecord>(), notRun: 0, error: _preBindError);
        }

        public string LastError => _target != null
            ? _target.LastError
            : (_preBindError.HasValue ? TestResultJsonSerializer.SerializeError(_preBindError.Value) : string.Empty);
    }
}
