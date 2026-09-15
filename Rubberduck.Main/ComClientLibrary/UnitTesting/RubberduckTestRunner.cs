using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Rubberduck.UnitTesting
{
    /// <summary>
    /// Automation adapter over the unchanged <see cref="ITestEngine"/> (design D5). Owns only
    /// name-based selection resolution and JSON translation, plus a small run state machine:
    /// Idle -&gt; Starting -&gt; Running -&gt; {Complete | Cancelled | Faulted}, with NoTestsMatched
    /// as a terminal shortcut out of Starting.
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
    /// COM visibility (the <c>IRubberduckTestRunner</c> interface implementation and its
    /// ComVisible/Guid/ProgId/ClassInterface/ComDefaultInterface attributes) is added in a
    /// separate follow-up commit/PR (task 3.11), kept out of this review-budget-sized unit.
    /// </remarks>
    public class RubberduckTestRunner
    {
        private enum RunState
        {
            Idle,
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

        private readonly List<HeadlessTestResultRecord> _results = new List<HeadlessTestResultRecord>();
        private RunState _state = RunState.Idle;
        private HeadlessSelectionInfo _selection = new HeadlessSelectionInfo("all", null, null);
        private string _runId = string.Empty;
        private DateTime _startedUtc;
        private DateTime _finishedUtc;
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
        public RubberduckTestRunner(
            ITestEngine engine,
            Func<DateTime> clock,
            Func<string> runIdFactory,
            Func<string> parserStatus,
            Action requestParse)
        {
            _engine = engine ?? throw new ArgumentNullException(nameof(engine));
            _clock = clock ?? (() => DateTime.UtcNow);
            _runIdFactory = runIdFactory ?? (() => Guid.NewGuid().ToString());
            _parserStatus = parserStatus;
            _requestParse = requestParse;

            // Subscribed exactly once, for the runner's entire lifetime: StartRun never
            // re-subscribes, so a second run cannot double-count events from the first
            // (design "Single subscription to TestCompleted/TestRunCompleted").
            _engine.TestCompleted += OnTestCompleted;
            _engine.TestRunCompleted += OnTestRunCompleted;
        }

        public string Version => Assembly.GetExecutingAssembly().GetName().Version.ToString() + "-headless";

        public int SchemaVersion => TestResultJsonSerializer.SchemaVersion;

        public bool IsReady => _engine.CanRun;

        public string ParserStatus => _parserStatus?.Invoke() ?? string.Empty;

        public int DiscoveredTestCount => _engine.Tests.Count();

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

            if (!_engine.CanRun)
            {
                SetError("NOT_READY", "The parser is not in a state that allows a test run.");
                return string.Empty;
            }

            var resolution = TestSelectionResolver.Resolve(
                _engine.Tests,
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
            _lastError = null;
            _cancelRequested = false;
            _runId = _runIdFactory();
            _startedUtc = _clock();
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

        private void OnTestCompleted(object sender, TestCompletedEventArgs e)
        {
            if (_state != RunState.Running)
            {
                return;
            }

            _results.Add(new HeadlessTestResultRecord(
                e.Test.Declaration.ProjectName,
                e.Test.Declaration.ComponentName,
                e.Test.Declaration.IdentifierName,
                e.Result.Outcome,
                e.Result.Output,
                e.Result.Duration));
        }

        private void OnTestRunCompleted(object sender, TestRunCompletedEventArgs e)
        {
            if (_state != RunState.Running)
            {
                return;
            }

            _finishedUtc = _clock();
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
