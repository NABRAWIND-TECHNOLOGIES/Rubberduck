using System;
using System.Collections.Generic;
using NUnit.Framework;
using Newtonsoft.Json.Linq;
using Rubberduck.UnitTesting;

namespace RubberduckTests.UnitTesting
{
    [TestFixture]
    [Category("HeadlessRunner")]
    public class HeadlessRunnerJsonSerializerTests
    {
        private static HeadlessRunInfo Run(string kind = "all", string module = null, string method = null)
        {
            var selection = new HeadlessSelectionInfo(kind, module, method);
            var started = new DateTime(2026, 9, 15, 10, 0, 0, DateTimeKind.Utc);
            var finished = started.AddMilliseconds(12345);
            return new HeadlessRunInfo("11111111-2222-3333-4444-555555555555", @"C:\path\Book.xlsm", selection, started, finished);
        }

        [Test]
        public void Serialize_FullResultSet_IncludesSchemaVersionRunAndAllTestFields()
        {
            var tests = new[]
            {
                new HeadlessTestResultRecord("VBAProject", "TestModule1", "TestSucceeds", TestOutcome.Succeeded, "", 14),
                new HeadlessTestResultRecord("VBAProject", "TestModule1", "TestFails", TestOutcome.Failed, "Expected 1 but was 2", 8)
            };

            var json = TestResultJsonSerializer.Serialize("2.5.9.1-headless.1", "completed", Run(), tests, notRun: 0, error: null);
            var envelope = JObject.Parse(json);

            Assert.AreEqual(1, (int)envelope["schemaVersion"]);
            Assert.AreEqual("2.5.9.1-headless.1", (string)envelope["port"]["version"]);
            Assert.AreEqual("completed", (string)envelope["status"]);
            Assert.AreEqual("11111111-2222-3333-4444-555555555555", (string)envelope["run"]["runId"]);
            Assert.AreEqual(@"C:\path\Book.xlsm", (string)envelope["run"]["workbook"]);
            Assert.AreEqual(12345, (long)envelope["run"]["durationMs"]);

            var testsArray = (JArray)envelope["tests"];
            Assert.AreEqual(2, testsArray.Count);
            var first = testsArray[0];
            Assert.AreEqual("VBAProject", (string)first["project"]);
            Assert.AreEqual("TestModule1", (string)first["module"]);
            Assert.AreEqual("TestFails", (string)first["method"]);
            Assert.AreEqual("Failed", (string)first["outcome"]);
            Assert.AreEqual("Expected 1 but was 2", (string)first["message"]);
            Assert.AreEqual(8, (long)first["durationMs"]);
        }

        [TestCase(TestOutcome.Succeeded, "Succeeded")]
        [TestCase(TestOutcome.Failed, "Failed")]
        [TestCase(TestOutcome.Inconclusive, "Inconclusive")]
        [TestCase(TestOutcome.Ignored, "Ignored")]
        [TestCase(TestOutcome.Unknown, "Unknown")]
        public void Serialize_OutcomeVocabulary_MapsExactTestOutcomeNames(TestOutcome outcome, string expected)
        {
            var tests = new[] { new HeadlessTestResultRecord("P", "M", "Method1", outcome, "", 1) };

            var json = TestResultJsonSerializer.Serialize("v", "completed", Run(), tests, notRun: 0, error: null);
            var envelope = JObject.Parse(json);

            Assert.AreEqual(expected, (string)envelope["tests"][0]["outcome"]);
        }

        [Test]
        public void Serialize_EmptyMessage_RendersEmptyStringNotNull()
        {
            var tests = new[] { new HeadlessTestResultRecord("P", "M", "Method1", TestOutcome.Succeeded, "", 1) };

            var json = TestResultJsonSerializer.Serialize("v", "completed", Run(), tests, notRun: 0, error: null);
            var envelope = JObject.Parse(json);

            var message = envelope["tests"][0]["message"];
            Assert.AreEqual(JTokenType.String, message.Type);
            Assert.AreEqual(string.Empty, (string)message);
        }

        [Test]
        public void Serialize_OrdersTestsByModuleThenMethod()
        {
            var tests = new[]
            {
                new HeadlessTestResultRecord("P", "ZModule", "AMethod", TestOutcome.Succeeded, "", 1),
                new HeadlessTestResultRecord("P", "AModule", "ZMethod", TestOutcome.Succeeded, "", 1),
                new HeadlessTestResultRecord("P", "AModule", "AMethod", TestOutcome.Succeeded, "", 1)
            };

            var json = TestResultJsonSerializer.Serialize("v", "completed", Run(), tests, notRun: 0, error: null);
            var envelope = JObject.Parse(json);
            var testsArray = (JArray)envelope["tests"];

            Assert.AreEqual("AModule", (string)testsArray[0]["module"]);
            Assert.AreEqual("AMethod", (string)testsArray[0]["method"]);
            Assert.AreEqual("AModule", (string)testsArray[1]["module"]);
            Assert.AreEqual("ZMethod", (string)testsArray[1]["method"]);
            Assert.AreEqual("ZModule", (string)testsArray[2]["module"]);
        }

        [Test]
        public void Serialize_EscapesQuotesBackslashesAndNewlinesInMessage()
        {
            const string rawMessage = "Expected \"1\" but got 2\\3\nnext line";
            var tests = new[] { new HeadlessTestResultRecord("P", "M", "Method1", TestOutcome.Failed, rawMessage, 1) };

            var json = TestResultJsonSerializer.Serialize("v", "completed", Run(), tests, notRun: 0, error: null);

            // Must be well-formed JSON despite the raw quotes/backslashes/newline in the message --
            // a naive concatenation would break parsing here.
            var envelope = JObject.Parse(json);
            Assert.AreEqual(rawMessage, (string)envelope["tests"][0]["message"]);
        }

        [Test]
        public void Serialize_SummaryArithmetic_MatchesOutcomeCountsAndNotRun()
        {
            var tests = new[]
            {
                new HeadlessTestResultRecord("P", "M", "T1", TestOutcome.Succeeded, "", 1),
                new HeadlessTestResultRecord("P", "M", "T2", TestOutcome.Succeeded, "", 1),
                new HeadlessTestResultRecord("P", "M", "T3", TestOutcome.Failed, "boom", 1),
                new HeadlessTestResultRecord("P", "M", "T4", TestOutcome.Inconclusive, "", 1),
                new HeadlessTestResultRecord("P", "M", "T5", TestOutcome.Ignored, "", 1),
                new HeadlessTestResultRecord("P", "M", "T6", TestOutcome.Unknown, "", 1)
            };

            var json = TestResultJsonSerializer.Serialize("v", "timedOut", Run(), tests, notRun: 2, error: null);
            var envelope = JObject.Parse(json);
            var summary = envelope["summary"];

            Assert.AreEqual(8, (int)summary["total"]);
            Assert.AreEqual(2, (int)summary["succeeded"]);
            Assert.AreEqual(1, (int)summary["failed"]);
            Assert.AreEqual(1, (int)summary["inconclusive"]);
            Assert.AreEqual(1, (int)summary["ignored"]);
            Assert.AreEqual(1, (int)summary["unknown"]);
            Assert.AreEqual(2, (int)summary["notRun"]);
        }

        [Test]
        public void Serialize_NullError_RendersJsonNull()
        {
            var json = TestResultJsonSerializer.Serialize("v", "completed", Run(), Array.Empty<HeadlessTestResultRecord>(), notRun: 0, error: null);
            var envelope = JObject.Parse(json);

            Assert.AreEqual(JTokenType.Null, envelope["error"].Type);
        }

        [Test]
        public void Serialize_WithError_IncludesCodeMessageAndDetail()
        {
            var error = new HeadlessErrorInfo("TIMEOUT", "Run did not complete in time", "elapsed=305000ms");

            var json = TestResultJsonSerializer.Serialize("v", "timedOut", Run(), Array.Empty<HeadlessTestResultRecord>(), notRun: 1, error: error);
            var envelope = JObject.Parse(json);

            Assert.AreEqual("TIMEOUT", (string)envelope["error"]["code"]);
            Assert.AreEqual("Run did not complete in time", (string)envelope["error"]["message"]);
            Assert.AreEqual("elapsed=305000ms", (string)envelope["error"]["detail"]);
        }
    }
}
