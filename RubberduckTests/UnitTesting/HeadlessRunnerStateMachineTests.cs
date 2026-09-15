using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;
using NUnit.Framework;
using Rubberduck.Parsing.VBA;
using Rubberduck.UnitTesting;

namespace RubberduckTests.UnitTesting
{
    [TestFixture]
    [Category("HeadlessRunner")]
    public class HeadlessRunnerStateMachineTests
    {
        // Controllable double for ITestEngine: completion is driven manually so the Running
        // window can be inspected (design 3.9 note: "a small ITestEngine fake with manual
        // TestCompleted/TestRunCompleted raising is preferable for state tests" -- unlike
        // MockedTestEngine, which always completes synchronously and cannot be paused).
        private sealed class ControllableFakeEngine : ITestEngine
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
            public int RunCallCount { get; private set; }
            public bool CancellationRequested { get; private set; }
            public bool ThrowsOnRun { get; set; }

            public void Run(IEnumerable<TestMethod> tests)
            {
                RunCallCount++;
                if (ThrowsOnRun)
                {
                    throw new InvalidOperationException("engine boom");
                }
            }

            public void RunByOutcome(TestOutcome outcome) { }
            public void RepeatLastRun() { }
            public void RequestCancellation() => CancellationRequested = true;

            public void RaiseTestCompleted(TestMethod method, TestResult result) =>
                TestCompleted?.Invoke(this, new TestCompletedEventArgs(method, result));

            public void RaiseTestRunCompleted() =>
                TestRunCompleted?.Invoke(this, new TestRunCompletedEventArgs(0));
        }

        // Reuses MockedTestEngine purely for discovery (real Declaration-backed TestMethod
        // instances), then hands the discovered list to a ControllableFakeEngine so the run
        // itself stays fully controllable.
        private static List<TestMethod> DiscoverTests(int methodCount)
        {
            var mocked = new MockedTestEngine(methodCount);
            mocked.ParserState.OnParseRequested(mocked);
            if (mocked.ParserState.Status != ParserState.Ready)
            {
                Assert.Inconclusive("Parser Error");
            }

            return mocked.TestEngine.Tests.ToList();
        }

        private static string StatusOf(string json) => (string)JObject.Parse(json)["status"];

        [Test]
        public void StartRun_OnSynchronousMockedEngine_CompletesBeforeReturning()
        {
            using (var mocked = new MockedTestEngine(1))
            {
                mocked.ParserState.OnParseRequested(mocked);
                if (mocked.ParserState.Status != ParserState.Ready)
                {
                    Assert.Inconclusive("Parser Error");
                }

                var runner = new RubberduckTestRunner(mocked.TestEngine);
                var runId = runner.StartRun();

                Assert.IsNotEmpty(runId);
                Assert.IsTrue(runner.IsComplete);
                Assert.AreEqual("completed", StatusOf(runner.GetResults()));
            }
        }

        [Test]
        public void StartRun_WhileRunInProgress_IsRejectedAndOriginalRunContinues()
        {
            var fake = new ControllableFakeEngine { Tests = DiscoverTests(1) };
            var runner = new RubberduckTestRunner(fake);

            runner.StartRun();
            var second = runner.StartRun();

            Assert.AreEqual(string.Empty, second);
            StringAssert.Contains("RUN_IN_PROGRESS", runner.LastError);
            Assert.AreEqual(1, fake.RunCallCount);
            Assert.IsFalse(runner.IsComplete);
        }

        [Test]
        public void GetResults_MidRun_ReturnsImmutablePartialSnapshot()
        {
            var tests = DiscoverTests(2);
            var fake = new ControllableFakeEngine { Tests = tests };
            var runner = new RubberduckTestRunner(fake);

            runner.StartRun();
            fake.RaiseTestCompleted(tests[0], new TestResult(TestOutcome.Succeeded));

            var first = runner.GetResults();
            var second = runner.GetResults();

            Assert.AreEqual(first, second);
            Assert.AreEqual(1, ((JArray)JObject.Parse(first)["tests"]).Count);
            Assert.IsFalse(runner.IsComplete);
        }

        [Test]
        public void Cancel_RequestsEngineCancellation_AndCompletesAsCancelled()
        {
            var fake = new ControllableFakeEngine { Tests = DiscoverTests(1) };
            var runner = new RubberduckTestRunner(fake);

            runner.StartRun();
            runner.Cancel();

            Assert.IsTrue(fake.CancellationRequested);
            Assert.IsFalse(runner.IsComplete);

            fake.RaiseTestRunCompleted();

            Assert.IsTrue(runner.IsComplete);
            Assert.AreEqual("cancelled", StatusOf(runner.GetResults()));
        }

        [Test]
        public void StartRun_SelectionMatchesNoTest_ReturnsNoTestsMatchedWithoutRunningEngine()
        {
            var fake = new ControllableFakeEngine { Tests = DiscoverTests(1) };
            var runner = new RubberduckTestRunner(fake);

            var runId = runner.StartRun(ModuleName: "DoesNotExist");

            Assert.AreEqual(string.Empty, runId);
            Assert.AreEqual(0, fake.RunCallCount);
            Assert.IsTrue(runner.IsComplete);
            StringAssert.Contains("NO_TESTS_MATCHED", runner.LastError);
            Assert.AreEqual("noTestsMatched", StatusOf(runner.GetResults()));
        }

        [Test]
        public void StartRun_EngineNotReady_ReturnsNotReadyWithoutRunningEngine()
        {
            var fake = new ControllableFakeEngine { Tests = DiscoverTests(1), CanRun = false };
            var runner = new RubberduckTestRunner(fake);

            runner.StartRun();

            Assert.AreEqual(0, fake.RunCallCount);
            StringAssert.Contains("NOT_READY", runner.LastError);
        }

        [Test]
        public void StartRun_MethodWithoutModule_ReturnsSelectionInvalidWithoutRunningEngine()
        {
            var fake = new ControllableFakeEngine { Tests = DiscoverTests(1) };
            var runner = new RubberduckTestRunner(fake);

            runner.StartRun(MethodName: "TestMethod1");

            Assert.AreEqual(0, fake.RunCallCount);
            StringAssert.Contains("SELECTION_INVALID", runner.LastError);
        }

        [Test]
        public void TestCompletedSubscription_IsSingleAcrossMultipleRuns()
        {
            var tests = DiscoverTests(1);
            var fake = new ControllableFakeEngine { Tests = tests };
            var runner = new RubberduckTestRunner(fake);

            runner.StartRun();
            fake.RaiseTestCompleted(tests[0], new TestResult(TestOutcome.Succeeded));
            fake.RaiseTestRunCompleted();

            runner.StartRun();
            fake.RaiseTestCompleted(tests[0], new TestResult(TestOutcome.Succeeded));

            Assert.AreEqual(1, ((JArray)JObject.Parse(runner.GetResults())["tests"]).Count,
                "A second run must not double-count events from a re-subscribed handler.");
        }

        [Test]
        public void StartRun_EngineThrows_TransitionsToFaultedWithEngineFault()
        {
            var fake = new ControllableFakeEngine { Tests = DiscoverTests(1), ThrowsOnRun = true };
            var runner = new RubberduckTestRunner(fake);

            runner.StartRun();

            Assert.IsTrue(runner.IsComplete);
            StringAssert.Contains("ENGINE_FAULT", runner.LastError);
            Assert.AreEqual("faulted", StatusOf(runner.GetResults()));
        }

        // PR5d: a COM client must never receive an exception from a null-guardable member.
        // engine.Tests is null before the parser's first successful Ready transition (TestEngine
        // only assigns its backing field inside StateChangedHandler's else-if branch), so both
        // DiscoveredTestCount and StartRun's selection resolution must tolerate it explicitly
        // instead of relying on CanRun alone.
        [Test]
        public void DiscoveredTestCount_WhenEngineTestsIsNull_ReturnsZeroInsteadOfThrowing()
        {
            var fake = new ControllableFakeEngine { Tests = null };
            var runner = new RubberduckTestRunner(fake);

            Assert.AreEqual(0, runner.DiscoveredTestCount);
        }

        [Test]
        public void StartRun_WhenEngineTestsIsNull_ReturnsNotReadyWithoutRunningEngine()
        {
            var fake = new ControllableFakeEngine { Tests = null };
            var runner = new RubberduckTestRunner(fake);

            var runId = runner.StartRun();

            Assert.AreEqual(string.Empty, runId);
            Assert.AreEqual(0, fake.RunCallCount);
            StringAssert.Contains("NOT_READY", runner.LastError);
            Assert.IsTrue(runner.IsComplete);
        }
    }
}
