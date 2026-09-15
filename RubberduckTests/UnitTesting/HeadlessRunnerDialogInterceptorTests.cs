using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;
using NUnit.Framework;
using Rubberduck.Automation;
using Rubberduck.Parsing.VBA;
using Rubberduck.UnitTesting;

namespace RubberduckTests.UnitTesting
{
    // D13 -- compile-error/modal-dialog handling (headless-test-runner, tasks 3.14/3.15).
    // Covers only the pure parts (design's own Test Strategy: the WH_CBT hook installation
    // itself is wiring-exempt, verified only by task 9.6's smoke check).
    [TestFixture]
    [Category("HeadlessRunner")]
    public class HeadlessRunnerDialogInterceptorTests
    {
        private static readonly DateTime Epoch = new DateTime(2026, 9, 15, 10, 0, 0, DateTimeKind.Utc);

        #region DialogClassifier

        [TestCase("Microsoft Visual Basic for Applications", "Compile error:\r\n\r\nVariable not defined", DialogKind.CompileError, DialogDismissAction.Ok)]
        [TestCase("Microsoft Visual Basic for Applications", "COMPILE ERROR: Sub or Function not defined", DialogKind.CompileError, DialogDismissAction.Ok)]
        [TestCase("Microsoft Visual Basic", "Run-time error '91':\r\n\r\nObject variable or With block variable not set", DialogKind.RuntimeDebugEndPrompt, DialogDismissAction.EndOrCancel)]
        [TestCase("Microsoft Excel", "Do you want to save changes?", DialogKind.Unrecognized, DialogDismissAction.Close)]
        [TestCase(null, null, DialogKind.Unrecognized, DialogDismissAction.Close)]
        public void Classify_ByCaptionAndText_ReturnsExpectedKindAndDismissAction(string caption, string text, DialogKind expectedKind, DialogDismissAction expectedAction)
        {
            var kind = DialogClassifier.Classify(caption, text);

            Assert.AreEqual(expectedKind, kind);
            Assert.AreEqual(expectedAction, DialogClassifier.ActionFor(kind));
        }

        #endregion

        #region DialogDiagnosticFormatter

        [Test]
        public void FormatMessage_CompileErrorWithFullLocation_UsesCompileErrorFormat()
        {
            var message = DialogDiagnosticFormatter.FormatMessage(
                DialogKind.CompileError, "Microsoft Visual Basic for Applications", "Compile error:\r\n\r\nVariable not defined",
                module: "HeadlessBrokenTests", line: 3, column: 5, source: "    y = 1");

            Assert.AreEqual("VBA compile error: Compile error:\r\n\r\nVariable not defined at HeadlessBrokenTests:3:5 —     y = 1", message);
        }

        [Test]
        public void FormatMessage_CompileErrorWithoutLocation_FailsSafeToDialogFormat()
        {
            var message = DialogDiagnosticFormatter.FormatMessage(
                DialogKind.CompileError, "Microsoft Visual Basic for Applications", "Compile error:\r\n\r\nVariable not defined",
                module: null, line: null, column: null, source: null);

            Assert.AreEqual("VBA dialog: Microsoft Visual Basic for Applications — Compile error:\r\n\r\nVariable not defined", message);
        }

        [Test]
        public void FormatMessage_RuntimePrompt_UsesDialogFormatRegardlessOfLocation()
        {
            var message = DialogDiagnosticFormatter.FormatMessage(
                DialogKind.RuntimeDebugEndPrompt, "Microsoft Visual Basic", "Run-time error '91':\r\n\r\nObject variable or With block variable not set",
                module: "SomeModule", line: 10, column: 1, source: "x.Foo");

            Assert.AreEqual("VBA dialog: Microsoft Visual Basic — Run-time error '91':\r\n\r\nObject variable or With block variable not set", message);
        }

        [Test]
        public void BuildDiagnostic_CompileErrorWithLocation_ReportsCompileErrorKindAndAllFields()
        {
            var diagnostic = DialogDiagnosticFormatter.BuildDiagnostic(
                DialogKind.CompileError, "Microsoft Visual Basic for Applications", "Compile error:\r\n\r\nVariable not defined",
                module: "HeadlessBrokenTests", line: 3, column: 5, source: "    y = 1");

            Assert.AreEqual("compileError", diagnostic.Kind);
            Assert.AreEqual("HeadlessBrokenTests", diagnostic.Module);
            Assert.AreEqual(3, diagnostic.Line);
            Assert.AreEqual(5, diagnostic.Column);
            Assert.AreEqual("    y = 1", diagnostic.Source);
            Assert.AreEqual("Microsoft Visual Basic for Applications", diagnostic.DialogCaption);
        }

        [Test]
        public void BuildDiagnostic_MissingLocation_ReportsDialogKindWithNullLocationFields()
        {
            var diagnostic = DialogDiagnosticFormatter.BuildDiagnostic(
                DialogKind.Unrecognized, "Microsoft Excel", "Do you want to save changes?",
                module: null, line: null, column: null, source: null);

            Assert.AreEqual("dialog", diagnostic.Kind);
            Assert.IsNull(diagnostic.Module);
            Assert.IsNull(diagnostic.Line);
            Assert.IsNull(diagnostic.Column);
            Assert.IsNull(diagnostic.Source);
        }

        #endregion

        #region HeadlessDialogLog

        [Test]
        public void EntriesBetween_HalfOpenWindow_IncludesStartExcludesEnd()
        {
            var log = new HeadlessDialogLog();
            var atStart = Entry(Epoch);
            var inside = Entry(Epoch.AddSeconds(1));
            var atEnd = Entry(Epoch.AddSeconds(2));
            var before = Entry(Epoch.AddSeconds(-1));
            var after = Entry(Epoch.AddSeconds(3));

            log.Append(before);
            log.Append(atStart);
            log.Append(inside);
            log.Append(atEnd);
            log.Append(after);

            var result = log.EntriesBetween(Epoch, Epoch.AddSeconds(2));

            CollectionAssert.AreEquivalent(new[] { atStart, inside }, result);
        }

        [Test]
        public void EntriesBetween_NoMatchingEntries_ReturnsEmpty()
        {
            var log = new HeadlessDialogLog();
            log.Append(Entry(Epoch.AddHours(-1)));

            var result = log.EntriesBetween(Epoch, Epoch.AddSeconds(1));

            Assert.IsEmpty(result);
        }

        private static HeadlessDialogEntry Entry(DateTime timestampUtc) =>
            new HeadlessDialogEntry("caption", "text", null, null, null, null, timestampUtc);

        #endregion

        #region Runner attachment (RubberduckTestRunner subscribes to TestStarted/TestCompleted)

        private sealed class ManualClock
        {
            public DateTime Current { get; set; }
            public DateTime Tick() => Current;
        }

        // Small controllable ITestEngine double, mirroring HeadlessRunnerStateMachineTests'
        // ControllableFakeEngine (kept local/duplicated per-file per existing precedent, e.g.
        // EngineTests.cs vs HeadlessRunnerStateMachineTests.cs each define their own doubles).
        private sealed class FakeAutomationEngine : ITestEngine
        {
            public event EventHandler<TestRunStartedEventArgs> TestRunStarted;
            public event EventHandler<TestStartedEventArgs> TestStarted;
            public event EventHandler<TestCompletedEventArgs> TestCompleted;
            public event EventHandler<TestRunCompletedEventArgs> TestRunCompleted;
            public event EventHandler TestsRefreshStarted;
            public event EventHandler TestsRefreshed;

            public IEnumerable<TestMethod> Tests { get; set; } = Enumerable.Empty<TestMethod>();
            public IReadOnlyList<TestMethod> LastRunTests => Array.Empty<TestMethod>();
            public bool CanRun { get; set; } = true;
            public bool CanRepeatLastRun => false;

            public void Run(IEnumerable<TestMethod> tests) { }
            public void RunByOutcome(TestOutcome outcome) { }
            public void RepeatLastRun() { }
            public void RequestCancellation() { }

            public void RaiseTestStarted(TestMethod method) => TestStarted?.Invoke(this, new TestStartedEventArgs(method));
            public void RaiseTestCompleted(TestMethod method, TestResult result) => TestCompleted?.Invoke(this, new TestCompletedEventArgs(method, result));
            public void RaiseTestRunCompleted() => TestRunCompleted?.Invoke(this, new TestRunCompletedEventArgs(0));
        }

        private static TestMethod DiscoverSingleTest()
        {
            using (var mocked = new MockedTestEngine(1))
            {
                mocked.ParserState.OnParseRequested(mocked);
                if (mocked.ParserState.Status != ParserState.Ready)
                {
                    Assert.Inconclusive("Parser Error");
                }

                return mocked.TestEngine.Tests.Single();
            }
        }

        [Test]
        public void OnTestCompleted_DialogEntryWithinTestWindow_ForcesInconclusiveWithEnrichedMessageAndDiagnostic()
        {
            var testMethod = DiscoverSingleTest();
            var fake = new FakeAutomationEngine { Tests = new[] { testMethod } };
            var dialogLog = new HeadlessDialogLog();
            var clock = new ManualClock { Current = Epoch };

            var runner = new RubberduckTestRunner(fake, clock: clock.Tick, runIdFactory: () => "run1", parserStatus: null, requestParse: null, dialogLog: dialogLog);

            runner.StartRun();

            clock.Current = Epoch.AddSeconds(1);
            fake.RaiseTestStarted(testMethod);

            dialogLog.Append(new HeadlessDialogEntry(
                "Microsoft Visual Basic for Applications", "Compile error:\r\n\r\nVariable not defined",
                "HeadlessBrokenTests", 3, 5, "    y = 1", Epoch.AddMilliseconds(1500)));

            clock.Current = Epoch.AddSeconds(2);
            // The engine would normally already report Inconclusive here (D12); a Succeeded
            // stub proves the dialog attachment -- not the underlying outcome -- forces the
            // final Inconclusive per spec ("the affected test MUST be reported Inconclusive").
            fake.RaiseTestCompleted(testMethod, new TestResult(TestOutcome.Succeeded, "", 5));

            clock.Current = Epoch.AddSeconds(3);
            fake.RaiseTestRunCompleted();

            var tests = (JArray)JObject.Parse(runner.GetResults())["tests"];
            var first = tests.Single();

            Assert.AreEqual("Inconclusive", (string)first["outcome"]);
            Assert.AreEqual("VBA compile error: Compile error:\r\n\r\nVariable not defined at HeadlessBrokenTests:3:5 —     y = 1", (string)first["message"]);
            Assert.AreEqual("compileError", (string)first["diagnostic"]["kind"]);
            Assert.AreEqual("HeadlessBrokenTests", (string)first["diagnostic"]["module"]);
            Assert.AreEqual(3, (int)first["diagnostic"]["line"]);
            Assert.AreEqual(5, (int)first["diagnostic"]["column"]);
        }

        [Test]
        public void OnTestCompleted_NoDialogEntryInWindow_LeavesOriginalOutcomeAndMessageUntouched()
        {
            var testMethod = DiscoverSingleTest();
            var fake = new FakeAutomationEngine { Tests = new[] { testMethod } };
            var dialogLog = new HeadlessDialogLog();
            var clock = new ManualClock { Current = Epoch };

            var runner = new RubberduckTestRunner(fake, clock: clock.Tick, runIdFactory: () => "run1", parserStatus: null, requestParse: null, dialogLog: dialogLog);

            runner.StartRun();
            clock.Current = Epoch.AddSeconds(1);
            fake.RaiseTestStarted(testMethod);
            clock.Current = Epoch.AddSeconds(2);
            fake.RaiseTestCompleted(testMethod, new TestResult(TestOutcome.Succeeded, "", 5));
            clock.Current = Epoch.AddSeconds(3);
            fake.RaiseTestRunCompleted();

            var tests = (JArray)JObject.Parse(runner.GetResults())["tests"];
            var first = tests.Single();

            Assert.AreEqual("Succeeded", (string)first["outcome"]);
            Assert.IsNull(first["diagnostic"]);
        }

        [Test]
        public void OnTestRunCompleted_DialogEntryOutsideAnyTestWindow_SurfacesInErrorDetail()
        {
            var testMethod = DiscoverSingleTest();
            var fake = new FakeAutomationEngine { Tests = new[] { testMethod } };
            var dialogLog = new HeadlessDialogLog();
            var clock = new ManualClock { Current = Epoch };

            var runner = new RubberduckTestRunner(fake, clock: clock.Tick, runIdFactory: () => "run1", parserStatus: null, requestParse: null, dialogLog: dialogLog);

            runner.StartRun();
            clock.Current = Epoch.AddSeconds(1);
            fake.RaiseTestStarted(testMethod);
            clock.Current = Epoch.AddSeconds(2);
            fake.RaiseTestCompleted(testMethod, new TestResult(TestOutcome.Succeeded, "", 5));

            // Fires after the only test's window [1s, 2s) already closed, but before the run
            // itself is marked complete -- e.g. a stray MsgBox left over from module cleanup.
            dialogLog.Append(new HeadlessDialogEntry("Microsoft Excel", "Do you want to save changes?", null, null, null, null, Epoch.AddMilliseconds(2500)));

            clock.Current = Epoch.AddSeconds(3);
            fake.RaiseTestRunCompleted();

            var envelope = JObject.Parse(runner.GetResults());
            Assert.AreEqual("DIALOG_OUTSIDE_TEST_WINDOW", (string)envelope["error"]["code"]);
            StringAssert.Contains("Do you want to save changes?", (string)envelope["error"]["detail"]);
        }

        #endregion
    }
}
