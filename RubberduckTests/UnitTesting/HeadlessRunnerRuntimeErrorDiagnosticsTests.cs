using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using Moq;
using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Resources.UnitTesting;
using Rubberduck.UnitTesting;

namespace RubberduckTests.UnitTesting
{
    // D12 — VBA runtime error diagnostics enrichment (headless-test-runner, tasks 3.12/3.13).
    // Uses the same MockedTestEngine/IVBEInteraction mocking pattern as EngineTests.cs.
    [NonParallelizable]
    [TestFixture, Apartment(ApartmentState.STA)]
    public class HeadlessRunnerRuntimeErrorDiagnosticsTests
    {
        [Test]
        [Category("Unit Testing")]
        public void RunTestMethod_ComExceptionWithVbaFacility_ReportsDecodedVbaError()
        {
            // 0x800A005B is the classic VBA runtime error 91 ("Object variable or With block
            // variable not set") re-surfaced as a COM HRESULT under FACILITY_CONTROL (0x800A).
            var comException = new COMException("Object variable or With block variable not set", unchecked((int)0x800A005B));

            var result = RunSingleTestExpectingThrow(comException);

            Assert.AreEqual(TestOutcome.Inconclusive, result.Outcome);
            Assert.AreEqual("VBA runtime error 91: Object variable or With block variable not set", result.Output);
        }

        [Test]
        [Category("Unit Testing")]
        public void RunTestMethod_ComExceptionWithNonVbaFacility_ReportsHresultAndDescription()
        {
            // 0x80070005 is E_ACCESSDENIED (FACILITY_WIN32) — not a VBA runtime error, so the
            // decoder cannot extract a VBA error number, but the HRESULT + description must
            // still surface instead of the old fixed generic text.
            var comException = new COMException("Access is denied.", unchecked((int)0x80070005));

            var result = RunSingleTestExpectingThrow(comException);

            Assert.AreEqual(TestOutcome.Inconclusive, result.Outcome);
            StringAssert.Contains("0x80070005", result.Output);
            StringAssert.Contains("Access is denied.", result.Output);
        }

        [Test]
        [Category("Unit Testing")]
        public void RunTestMethod_GenericException_ReportsUnchangedGenericText()
        {
            var result = RunSingleTestExpectingThrow(new InvalidOperationException("boom"));

            Assert.AreEqual(TestOutcome.Inconclusive, result.Outcome);
            Assert.AreEqual(AssertMessages.TestRunner_ExceptionDuringRun, result.Output);
        }

        [Test]
        [Category("Unit Testing")]
        public void TestInitialize_ComExceptionWithVbaFacility_ReportsDecodedVbaErrorAndSkipsTest()
        {
            var comException = new COMException("Type mismatch", unchecked((int)0x800A000D));

            using (var engine = new MockedTestEngine(1))
            {
                // First RunDeclarations call is @ModuleInitialize (must succeed); the second is
                // @TestInitialize for the single test method, which throws.
                engine.VbeInteraction
                    .SetupSequence(ia => ia.RunDeclarations(engine.TypeLib.Object, It.IsAny<IEnumerable<Declaration>>()))
                    .Pass()
                    .Throws(comException);

                var completionEvents = new List<TestCompletedEventArgs>();
                engine.TestEngine.TestCompleted += (source, args) => completionEvents.Add(args);
                engine.ParserState.OnParseRequested(engine);

                if (engine.ParserState.Status != Rubberduck.Parsing.VBA.ParserState.Ready)
                {
                    Assert.Inconclusive("Parser Error");
                }

                engine.TestEngine.Run(engine.TestEngine.Tests);

                Assert.AreEqual(1, completionEvents.Count);
                var result = completionEvents[0].Result;
                Assert.AreEqual(TestOutcome.Inconclusive, result.Outcome);
                Assert.AreEqual("VBA runtime error 13: Type mismatch", result.Output);
            }
        }

        [Test]
        [Category("Unit Testing")]
        public void ModuleInitialize_ComExceptionWithVbaFacility_ReportsUnknownWithDecodedVbaError()
        {
            var comException = new COMException("Division by zero", unchecked((int)0x800A000B));

            using (var engine = new MockedTestEngine(1))
            {
                // The very first RunDeclarations call is @ModuleInitialize.
                engine.VbeInteraction
                    .Setup(ia => ia.RunDeclarations(engine.TypeLib.Object, It.IsAny<IEnumerable<Declaration>>()))
                    .Throws(comException);

                var completionEvents = new List<TestCompletedEventArgs>();
                engine.TestEngine.TestCompleted += (source, args) => completionEvents.Add(args);
                engine.ParserState.OnParseRequested(engine);

                if (engine.ParserState.Status != Rubberduck.Parsing.VBA.ParserState.Ready)
                {
                    Assert.Inconclusive("Parser Error");
                }

                engine.TestEngine.Run(engine.TestEngine.Tests);

                Assert.AreEqual(1, completionEvents.Count);
                var result = completionEvents[0].Result;
                Assert.AreEqual(TestOutcome.Unknown, result.Outcome);
                Assert.AreEqual("VBA runtime error 11: Division by zero", result.Output);
            }
        }

        private static TestResult RunSingleTestExpectingThrow(Exception exceptionToThrow)
        {
            using (var engine = new MockedTestEngine(1))
            {
                engine.VbeInteraction
                    .Setup(ia => ia.RunTestMethod(engine.TypeLib.Object, It.IsAny<TestMethod>(), It.IsAny<EventHandler<AssertCompletedEventArgs>>(), out It.Ref<long>.IsAny))
                    .Throws(exceptionToThrow);

                var completionEvents = new List<TestCompletedEventArgs>();
                engine.TestEngine.TestCompleted += (source, args) => completionEvents.Add(args);
                engine.ParserState.OnParseRequested(engine);

                if (engine.ParserState.Status != Rubberduck.Parsing.VBA.ParserState.Ready)
                {
                    Assert.Inconclusive("Parser Error");
                }

                engine.TestEngine.Run(engine.TestEngine.Tests);

                Assert.AreEqual(1, completionEvents.Count);
                return completionEvents[0].Result;
            }
        }
    }
}
