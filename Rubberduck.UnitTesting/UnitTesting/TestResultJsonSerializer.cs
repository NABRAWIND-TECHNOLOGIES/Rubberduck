using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace Rubberduck.UnitTesting
{
    /// <summary>
    /// A single executed test's outcome, decoupled from <see cref="TestMethod"/> so the JSON
    /// serializer needs no live VBE/parser state to be unit-tested (design D4/D5).
    /// </summary>
    public readonly struct HeadlessTestResultRecord
    {
        public HeadlessTestResultRecord(string project, string module, string method, TestOutcome outcome, string message, long durationMs)
        {
            Project = project ?? string.Empty;
            Module = module ?? string.Empty;
            Method = method ?? string.Empty;
            Outcome = outcome;
            Message = message ?? string.Empty;
            DurationMs = durationMs;
        }

        public string Project { get; }
        public string Module { get; }
        public string Method { get; }
        public TestOutcome Outcome { get; }
        public string Message { get; }
        public long DurationMs { get; }
    }

    /// <summary>
    /// The resolved selection that produced a run, mirrored into the JSON envelope's
    /// <c>run.selection</c> object (design D4).
    /// </summary>
    public readonly struct HeadlessSelectionInfo
    {
        public HeadlessSelectionInfo(string kind, string module, string method)
        {
            Kind = kind ?? string.Empty;
            Module = module;
            Method = method;
        }

        public string Kind { get; }

        /// <summary>Null when the selection did not target a specific module.</summary>
        public string Module { get; }

        /// <summary>Null when the selection did not target a specific method.</summary>
        public string Method { get; }
    }

    /// <summary>
    /// Run-level metadata mirrored into the JSON envelope's <c>run</c> object (design D4).
    /// </summary>
    public readonly struct HeadlessRunInfo
    {
        public HeadlessRunInfo(string runId, string workbook, HeadlessSelectionInfo selection, DateTime startedUtc, DateTime finishedUtc)
        {
            RunId = runId ?? string.Empty;
            Workbook = workbook ?? string.Empty;
            Selection = selection;
            StartedUtc = startedUtc;
            FinishedUtc = finishedUtc;
        }

        public string RunId { get; }
        public string Workbook { get; }
        public HeadlessSelectionInfo Selection { get; }
        public DateTime StartedUtc { get; }
        public DateTime FinishedUtc { get; }
        public long DurationMs => (long)Math.Max(0, (FinishedUtc - StartedUtc).TotalMilliseconds);
    }

    /// <summary>
    /// Structured error state mirrored into the JSON envelope's <c>error</c> object (design D4).
    /// </summary>
    public readonly struct HeadlessErrorInfo
    {
        public HeadlessErrorInfo(string code, string message, string detail)
        {
            Code = code ?? string.Empty;
            Message = message ?? string.Empty;
            Detail = detail;
        }

        public string Code { get; }
        public string Message { get; }

        /// <summary>Null when no further detail is available.</summary>
        public string Detail { get; }
    }

    /// <summary>
    /// Pure serializer for the headless test runner's JSON result envelope (design D4).
    /// Hand-rolled rather than depending on a JSON library: Rubberduck.UnitTesting has no
    /// Newtonsoft.Json reference (only Rubberduck.Core does), and the envelope shape is small
    /// and fixed enough that a minimal, fully-escaped writer is simpler than adding a new
    /// production dependency for it.
    /// </summary>
    internal static class TestResultJsonSerializer
    {
        internal const int SchemaVersion = 1;

        internal static string Serialize(
            string portVersion,
            string status,
            HeadlessRunInfo run,
            IReadOnlyList<HeadlessTestResultRecord> tests,
            int notRun,
            HeadlessErrorInfo? error)
        {
            if (tests is null)
            {
                throw new ArgumentNullException(nameof(tests));
            }

            // Deterministic ordering (module then method) so two runs over the same selection
            // produce byte-identical JSON regardless of the engine's internal enumeration order.
            var ordered = tests
                .OrderBy(t => t.Module, StringComparer.Ordinal)
                .ThenBy(t => t.Method, StringComparer.Ordinal)
                .ToList();

            var succeeded = ordered.Count(t => t.Outcome == TestOutcome.Succeeded);
            var failed = ordered.Count(t => t.Outcome == TestOutcome.Failed);
            var inconclusive = ordered.Count(t => t.Outcome == TestOutcome.Inconclusive);
            var ignored = ordered.Count(t => t.Outcome == TestOutcome.Ignored);
            var unknown = ordered.Count(t => t.Outcome == TestOutcome.Unknown);
            var total = ordered.Count + notRun;

            var testsJson = string.Join(",", ordered.Select(SerializeTest));

            var runJson = "{" +
                $"\"runId\":{JsonString(run.RunId)}," +
                $"\"workbook\":{JsonString(run.Workbook)}," +
                $"\"selection\":{SerializeSelection(run.Selection)}," +
                $"\"startedUtc\":{JsonString(FormatUtc(run.StartedUtc))}," +
                $"\"finishedUtc\":{JsonString(FormatUtc(run.FinishedUtc))}," +
                $"\"durationMs\":{run.DurationMs}" +
                "}";

            var summaryJson = "{" +
                $"\"total\":{total}," +
                $"\"succeeded\":{succeeded}," +
                $"\"failed\":{failed}," +
                $"\"inconclusive\":{inconclusive}," +
                $"\"ignored\":{ignored}," +
                $"\"unknown\":{unknown}," +
                $"\"notRun\":{notRun}" +
                "}";

            var errorJson = error.HasValue ? SerializeError(error.Value) : "null";

            return "{" +
                $"\"schemaVersion\":{SchemaVersion}," +
                $"\"port\":{{\"version\":{JsonString(portVersion)}}}," +
                $"\"status\":{JsonString(status)}," +
                $"\"run\":{runJson}," +
                $"\"summary\":{summaryJson}," +
                $"\"tests\":[{testsJson}]," +
                $"\"error\":{errorJson}" +
                "}";
        }

        private static string SerializeTest(HeadlessTestResultRecord record)
        {
            return "{" +
                $"\"project\":{JsonString(record.Project)}," +
                $"\"module\":{JsonString(record.Module)}," +
                $"\"method\":{JsonString(record.Method)}," +
                $"\"outcome\":{JsonString(record.Outcome.ToString())}," +
                $"\"message\":{JsonString(record.Message)}," +
                $"\"durationMs\":{record.DurationMs}" +
                "}";
        }

        private static string SerializeSelection(HeadlessSelectionInfo selection)
        {
            return "{" +
                $"\"kind\":{JsonString(selection.Kind)}," +
                $"\"module\":{(selection.Module is null ? "null" : JsonString(selection.Module))}," +
                $"\"method\":{(selection.Method is null ? "null" : JsonString(selection.Method))}" +
                "}";
        }

        private static string SerializeError(HeadlessErrorInfo error)
        {
            return "{" +
                $"\"code\":{JsonString(error.Code)}," +
                $"\"message\":{JsonString(error.Message)}," +
                $"\"detail\":{(error.Detail is null ? "null" : JsonString(error.Detail))}" +
                "}";
        }

        private static string FormatUtc(DateTime value)
        {
            return value.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ", CultureInfo.InvariantCulture);
        }

        private static string JsonString(string value)
        {
            var sb = new StringBuilder(value.Length + 2);
            sb.Append('"');
            foreach (var c in value)
            {
                switch (c)
                {
                    case '"':
                        sb.Append("\\\"");
                        break;
                    case '\\':
                        sb.Append("\\\\");
                        break;
                    case '\n':
                        sb.Append("\\n");
                        break;
                    case '\r':
                        sb.Append("\\r");
                        break;
                    case '\t':
                        sb.Append("\\t");
                        break;
                    default:
                        if (c < 0x20)
                        {
                            sb.Append("\\u").Append(((int)c).ToString("x4", CultureInfo.InvariantCulture));
                        }
                        else
                        {
                            sb.Append(c);
                        }
                        break;
                }
            }
            sb.Append('"');
            return sb.ToString();
        }
    }
}
