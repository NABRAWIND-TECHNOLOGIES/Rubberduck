using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Rubberduck.Automation;
using Rubberduck.Resources.Registration;

namespace Rubberduck.UnitTesting
{
    /// <summary>
    /// COM-visible automation adapter over the unchanged <see cref="ITestEngine"/> (design D5).
    /// Owns only name-based selection resolution and JSON translation, plus a small run state
    /// machine: Idle -&gt; Starting -&gt; Running -&gt; {Complete | Cancelled | Faulted}, with
    /// NoTestsMatched as a terminal shortcut out of Starting.
    /// </summary>
    /// <remarks>
    /// Threading: every port method and every engine event handler run on the Excel STA main
    /// thread by construction -- COM dispatch on an STA object is serialized by the message
    /// loop, and <see cref="TestEngine"/> itself posts all continuations back to that same
    /// thread. No locks are used here, and none should be added: a lock risks deadlocking
    /// against the engine's own message-queue reentrancy (design "Threading model and thread
    /// safety"). <see cref="GetResults"/> therefore returns an immutable snapshot and never
    /// mutates state, and <see cref="IsComplete"/> is a plain field read.
    ///
    /// GUID/ProgId <em>registration and wiring</em> (IoC container, <c>Extension.cs</c> Object
    /// assignment) are added in the wiring slice (PR5, tasks 4.1-4.4). The
    /// <see cref="RubberduckGuid.TestRunnerGuid"/>/<see cref="RubberduckProgId.TestRunnerProgId"/>
    /// constants themselves had to be added alongside this file because the fork's
    /// ComVisibleTypeAnalyzer requires every COM-visible type's Guid/ProgId attribute to
    /// reference those constants by source text -- a literal string does not satisfy it.
    /// </remarks>
    [
        ComVisible(true),
        Guid(RubberduckGuid.TestRunnerGuid),
        ProgId(RubberduckProgId.TestRunnerProgId),
        ClassInterface(ClassInterfaceType.None),
        ComDefaultInterface(typeof(IRubberduckTestRunner)),
        EditorBrowsable(EditorBrowsableState.Always)
    ]
    public class RubberduckTestRunner : IRubberduckTestRunner
    {
        private enum RunState
        {
            Idle,
            // Reserved for a future async hand-off between StartRun validation and the engine
            // actually starting; StartRun currently transitions Idle/terminal -> Running
            // directly (engine.Run() call is synchronous from the caller's point of view), so
            // this value is never assigned today. Kept so IsComplete's guard clause and the
            // RUN_IN_PROGRESS check stay correct if that hand-off is ever introduced.
            Starting,
            Running,
            Complete,
            Cancelled,
            Faulted,
            NoTestsMatched
        }

        private readonly ITestEngine _engine;
        private readonly Func<DateTime> _clock;
        private readonly Func<string> _runIdFactory;
        private readonly Func<string> _parserStatus;
        private readonly Action _requestParse;
        private readonly HeadlessDialogLog _dialogLog;

        private readonly List<HeadlessTestResultRecord> _results = new List<HeadlessTestResultRecord>();
        private readonly List<(DateTime Start, DateTime End)> _consumedDialogWindows = new List<(DateTime Start, DateTime End)>();
        private RunState _state = RunState.Idle;
        private HeadlessSelectionInfo _selection = new HeadlessSelectionInfo("all", null, null);
        private string _runId = string.Empty;
        private DateTime _startedUtc;
        private DateTime _finishedUtc;
        private DateTime _currentTestStartedUtc;
        private HeadlessErrorInfo? _lastError;
        private bool _cancelRequested;

        public RubberduckTestRunner(ITestEngine engine)
            : this(engine, clock: null, runIdFactory: null, parserStatus: null, requestParse: null)
        {
        }

        /// <param name="engine">The unchanged test engine this port adapts (design D1).</param>
        /// <param name="clock">Source for run timestamps; defaults to <see cref="DateTime.UtcNow"/>.</param>
        /// <param name="runIdFactory">Source for run ids; defaults to a new GUID per run.</param>
        /// <param name="parserStatus">Optional source for <see cref="ParserStatus"/>; the engine exposes no parser state directly.</param>
        /// <param name="requestParse">Optional trigger for <see cref="RequestParse"/>; the engine exposes no parse request directly.</param>
        /// <param name="dialogLog">
        /// Source of intercepted modal dialogs (design D13); defaults to the per-process
        /// <see cref="HeadlessDialogLog.Shared"/> instance the interceptor hook writes to.
        /// </param>
        public RubberduckTestRunner(
            ITestEngine engine,
            Func<DateTime> clock,
            Func<string> runIdFactory,
            Func<string> parserStatus,
            Action requestParse,
            HeadlessDialogLog dialogLog = null)
        {
            _engine = engine ?? throw new ArgumentNullException(nameof(engine));
            _clock = clock ?? (() => DateTime.UtcNow);
            _runIdFactory = runIdFactory ?? (() => Guid.NewGuid().ToString());
            _parserStatus = parserStatus;
            _requestParse = requestParse;
            _dialogLog = dialogLog ?? HeadlessDialogLog.Shared;

            // Subscribed exactly once, for the runner's entire lifetime: StartRun never
            // re-subscribes, so a second run cannot double-count events from the first
            // (design "Single subscription to TestCompleted/TestRunCompleted").
            _engine.TestStarted += OnTestStarted;
            _engine.TestCompleted += OnTestCompleted;
            _engine.TestRunCompleted += OnTestRunCompleted;
        }

        public string Version => Assembly.GetExecutingAssembly().GetName().Version.ToString() + "-headless";

        public int SchemaVersion => TestResultJsonSerializer.SchemaVersion;

        public bool IsReady => _engine.CanRun;

        public string ParserStatus => _parserStatus?.Invoke() ?? string.Empty;

        // engine.Tests is null until the parser's first successful Ready transition
        // (TestEngine only assigns its backing field from inside its own StateChangedHandler),
        // so a COM client asking before that point must see 0, not an exception.
        public int DiscoveredTestCount => _engine.Tests?.Count() ?? 0;

        public bool IsComplete => _state != RunState.Starting && _state != RunState.Running;

        public string LastError => _lastError.HasValue
            ? TestResultJsonSerializer.SerializeError(_lastError.Value)
            : string.Empty;

        public void RequestParse() => _requestParse?.Invoke();

        public string StartRun(string ModuleName = "", string MethodName = "")
        {
            if (_state == RunState.Starting || _state == RunState.Running)
            {
                SetError("RUN_IN_PROGRESS", "A test run is already in progress.");
                return string.Empty;
            }

            var discoveredTests = _engine.Tests;
            if (!_engine.CanRun || discoveredTests is null)
            {
                SetError("NOT_READY", "The parser is not in a state that allows a test run.");
                return string.Empty;
            }

            var resolution = TestSelectionResolver.Resolve(
                discoveredTests,
                t => t.Declaration.ComponentName,
                t => t.Declaration.IdentifierName,
                ModuleName,
                MethodName);

            if (resolution.Status == TestSelectionStatus.Invalid)
            {
                SetError("SELECTION_INVALID", "A method name requires a module name.");
                return string.Empty;
            }

            var kind = string.IsNullOrEmpty(ModuleName) ? "all" : string.IsNullOrEmpty(MethodName) ? "module" : "test";
            _selection = new HeadlessSelectionInfo(
                kind,
                string.IsNullOrEmpty(ModuleName) ? null : ModuleName,
                string.IsNullOrEmpty(MethodName) ? null : MethodName);

            if (resolution.Status == TestSelectionStatus.NoMatch)
            {
                _results.Clear();
                _runId = string.Empty;
                _startedUtc = _clock();
                _finishedUtc = _startedUtc;
                _state = RunState.NoTestsMatched;
                SetError("NO_TESTS_MATCHED", "The selection matched no discovered test.");
                return string.Empty;
            }

            _results.Clear();
            _consumedDialogWindows.Clear();
            _lastError = null;
            _cancelRequested = false;
            _runId = _runIdFactory();
            _startedUtc = _clock();
            _currentTestStartedUtc = _startedUtc;
            _state = RunState.Running;

            try
            {
                _engine.Run(resolution.Tests);
            }
            catch (Exception ex)
            {
                _finishedUtc = _clock();
                _state = RunState.Faulted;
                SetError("ENGINE_FAULT", "The test engine threw while running the selected tests.", ex.Message);
            }

            return _runId;
        }

        public void Cancel()
        {
            if (_state != RunState.Starting && _state != RunState.Running)
            {
                return;
            }

            _cancelRequested = true;
            _engine.RequestCancellation();
        }

        public string GetResults()
        {
            // Immutable snapshot: copy the list so a caller, or a concurrent OnTestCompleted
            // reentering via the engine's message-queue flush, can never observe or cause a
            // mutation through this call (design "GetResults must return an immutable
            // snapshot and must never mutate runner state").
            var snapshot = _results.ToList();
            var run = new HeadlessRunInfo(_runId, string.Empty, _selection, _startedUtc, _finishedUtc);

            return TestResultJsonSerializer.Serialize(Version, StatusFor(_state), run, snapshot, notRun: 0, error: _lastError);
        }

        private void OnTestStarted(object sender, TestStartedEventArgs e)
        {
            if (_state != RunState.Running)
            {
                return;
            }

            _currentTestStartedUtc = _clock();
        }

        private void OnTestCompleted(object sender, TestCompletedEventArgs e)
        {
            if (_state != RunState.Running)
            {
                return;
            }

            var completedUtc = _clock();
            var windowStart = _currentTestStartedUtc;

            var outcome = e.Result.Outcome;
            var message = e.Result.Output;
            HeadlessDialogDiagnosticInfo? diagnostic = null;

            // A dialog captured while this test was executing (design D13) overrides the
            // engine's own outcome/message -- the dialog interrupted the test's real
            // execution, so whatever TestOutcome the interrupted COM call produced is not
            // trustworthy on its own (spec "the affected test MUST be reported Inconclusive").
            var dialogEntriesInWindow = _dialogLog.EntriesBetween(windowStart, completedUtc);
            if (dialogEntriesInWindow.Count > 0)
            {
                var dialogEntry = dialogEntriesInWindow[0];
                var kind = DialogClassifier.Classify(dialogEntry.Caption, dialogEntry.Text);
                outcome = TestOutcome.Inconclusive;
                message = DialogDiagnosticFormatter.FormatMessage(kind, dialogEntry.Caption, dialogEntry.Text, dialogEntry.Module, dialogEntry.Line, dialogEntry.Column, dialogEntry.Source);
                diagnostic = DialogDiagnosticFormatter.BuildDiagnostic(kind, dialogEntry.Caption, dialogEntry.Text, dialogEntry.Module, dialogEntry.Line, dialogEntry.Column, dialogEntry.Source);
            }

            _consumedDialogWindows.Add((windowStart, completedUtc));

            _results.Add(new HeadlessTestResultRecord(
                e.Test.Declaration.ProjectName,
                e.Test.Declaration.ComponentName,
                e.Test.Declaration.IdentifierName,
                outcome,
                message,
                e.Result.Duration,
                diagnostic));
        }

        private void OnTestRunCompleted(object sender, TestRunCompletedEventArgs e)
        {
            if (_state != RunState.Running)
            {
                return;
            }

            _finishedUtc = _clock();

            // A dialog captured outside every consumed per-test window (e.g. during module
            // cleanup, or between the last test and TestRunCompleted) cannot be attributed to
            // any specific test, so it surfaces at the run level instead (design D13: "a
            // dialog outside any test window goes to error.detail"). A pre-existing, more
            // specific error (e.g. ENGINE_FAULT) is never overwritten by this.
            if (_lastError is null)
            {
                var strayEntries = _dialogLog
                    .EntriesBetween(_startedUtc, _finishedUtc)
                    .Where(entry => !_consumedDialogWindows.Any(window => entry.TimestampUtc >= window.Start && entry.TimestampUtc < window.End))
                    .ToList();

                if (strayEntries.Count > 0)
                {
                    var strayEntry = strayEntries[0];
                    var kind = DialogClassifier.Classify(strayEntry.Caption, strayEntry.Text);
                    var detail = DialogDiagnosticFormatter.FormatMessage(kind, strayEntry.Caption, strayEntry.Text, strayEntry.Module, strayEntry.Line, strayEntry.Column, strayEntry.Source);
                    SetError("DIALOG_OUTSIDE_TEST_WINDOW", "A VBA dialog appeared outside any test's execution window.", detail);
                }
            }

            _state = _cancelRequested ? RunState.Cancelled : RunState.Complete;
            _cancelRequested = false;
        }

        private void SetError(string code, string message, string detail = null)
        {
            _lastError = new HeadlessErrorInfo(code, message, detail);
        }

        private static string StatusFor(RunState state)
        {
            switch (state)
            {
                case RunState.Complete:
                    return "completed";
                case RunState.Cancelled:
                    return "cancelled";
                case RunState.Faulted:
                    return "faulted";
                case RunState.NoTestsMatched:
                    return "noTestsMatched";
                case RunState.Idle:
                    return "idle";
                default:
                    return "running";
            }
        }
    }
}
