using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Abstractions;
using System.IO.Abstractions.TestingHelpers;
using NUnit.Framework;
using Rubberduck.Automation;

namespace RubberduckTests.UnitTesting
{
    [TestFixture]
    [Category("HeadlessRunner")]
    public class HeadlessRunnerAutomationModeTests
    {
        private const string BaseDirectory = @"C:\Users\Test\AppData\Local\Rubberduck\automation";
        private const int Pid = 4242;

        private static string MarkerPath(int pid = Pid) => System.IO.Path.Combine(BaseDirectory, $"{pid}.json");

        private static MockFileSystem FileSystemWithMarker(string path, string json) =>
            new MockFileSystem(new Dictionary<string, MockFileData>
            {
                { path, new MockFileData(json) }
            });

        private static AutomationMode CreateMode(MockFileSystem fileSystem, DateTime now, int pid = Pid) =>
            new AutomationMode(fileSystem, BaseDirectory, () => now, pid);

        [Test]
        public void IsActive_MarkerPresentForCurrentPidAndFresh_ReturnsTrue()
        {
            var now = new DateTime(2026, 9, 15, 12, 0, 0, DateTimeKind.Utc);
            var createdUtc = now.AddMinutes(-2);
            var fileSystem = FileSystemWithMarker(MarkerPath(), $"{{\"createdUtc\":\"{createdUtc:O}\",\"runId\":\"abc-123\"}}");

            var mode = CreateMode(fileSystem, now);

            Assert.IsTrue(mode.IsActive);
        }

        [Test]
        public void IsActive_NoMarkerFile_ReturnsFalse()
        {
            var fileSystem = new MockFileSystem();
            var mode = CreateMode(fileSystem, DateTime.UtcNow);

            Assert.IsFalse(mode.IsActive);
        }

        [Test]
        public void IsActive_MarkerOlderThanTenMinutes_ReturnsFalse()
        {
            var now = new DateTime(2026, 9, 15, 12, 0, 0, DateTimeKind.Utc);
            var createdUtc = now.AddMinutes(-11);
            var fileSystem = FileSystemWithMarker(MarkerPath(), $"{{\"createdUtc\":\"{createdUtc:O}\",\"runId\":\"abc-123\"}}");

            var mode = CreateMode(fileSystem, now);

            Assert.IsFalse(mode.IsActive);
        }

        [Test]
        public void IsActive_MarkerJustUnderTenMinutesOld_ReturnsTrue()
        {
            var now = new DateTime(2026, 9, 15, 12, 0, 0, DateTimeKind.Utc);
            var createdUtc = now.AddMinutes(-9).AddSeconds(-59);
            var fileSystem = FileSystemWithMarker(MarkerPath(), $"{{\"createdUtc\":\"{createdUtc:O}\",\"runId\":\"abc-123\"}}");

            var mode = CreateMode(fileSystem, now);

            Assert.IsTrue(mode.IsActive);
        }

        [Test]
        public void IsActive_MarkerContainsMalformedJson_ReturnsFalse()
        {
            var fileSystem = FileSystemWithMarker(MarkerPath(), "{not-valid-json");
            var mode = CreateMode(fileSystem, DateTime.UtcNow);

            Assert.IsFalse(mode.IsActive);
        }

        [Test]
        public void IsActive_MarkerExistsOnlyForDifferentPid_IsIgnored()
        {
            var now = DateTime.UtcNow;
            var fileSystem = FileSystemWithMarker(MarkerPath(pid: 9999), $"{{\"createdUtc\":\"{now:O}\",\"runId\":\"abc-123\"}}");

            var mode = CreateMode(fileSystem, now, pid: Pid);

            Assert.IsFalse(mode.IsActive);
        }

        [Test]
        public void TryRead_MarkerPresent_ReturnsParsedCreatedUtcAndRunId()
        {
            var now = new DateTime(2026, 9, 15, 12, 0, 0, DateTimeKind.Utc);
            var createdUtc = now.AddMinutes(-1);
            var fileSystem = FileSystemWithMarker(MarkerPath(), $"{{\"createdUtc\":\"{createdUtc:O}\",\"runId\":\"run-xyz\"}}");

            var mode = CreateMode(fileSystem, now);

            Assert.IsTrue(mode.TryRead(out var marker));
            Assert.AreEqual(createdUtc, marker.CreatedUtc);
            Assert.AreEqual("run-xyz", marker.RunId);
        }

        // Hardening: DateTime.SpecifyKind on a value Json.NET already zone-adjusted based on the
        // machine's local time zone would silently mislabel it as UTC without converting it --
        // wrong by exactly the local UTC offset. Asserting against a fixed absolute instant
        // (rather than a threshold comparison) makes this deterministic regardless of the
        // machine running the test.
        [Test]
        public void TryRead_MarkerCreatedUtcHasNonZeroOffset_ParsedAsTheExactUtcInstant()
        {
            var expectedUtc = new DateTime(2026, 9, 15, 6, 0, 0, DateTimeKind.Utc);
            var fileSystem = FileSystemWithMarker(
                MarkerPath(),
                "{\"createdUtc\":\"2026-09-15T08:00:00+02:00\",\"runId\":\"offset-case\"}");

            // Fixed just after the true UTC instant, so a correct parse reports "fresh" and an
            // instant shifted by the local machine's UTC offset would very likely report "stale"
            // (or, for a machine whose local offset happens to also be +02:00, would resolve to
            // exactly expectedUtc by coincidence) -- either way TryRead's own out parameter is
            // the deterministic assertion below, not this boolean.
            var mode = CreateMode(fileSystem, expectedUtc.AddMinutes(1));

            Assert.IsTrue(mode.TryRead(out var marker));
            Assert.AreEqual(expectedUtc, marker.CreatedUtc);
            Assert.AreEqual("offset-case", marker.RunId);
        }

        // Hardening: exercise the real IOException catch path (design D9/AutomationMode) with a
        // genuinely locked file on the real file system, rather than a hand-rolled throwing
        // IFileSystem fake -- zero mocks, and the deterministic Windows sharing-violation
        // behaviour of FileShare.None gives a real IOException on the second handle.
        [Test]
        public void TryRead_MarkerFileLockedByAnotherHandle_ReturnsFalseInsteadOfThrowing()
        {
            var baseDirectory = Path.Combine(Path.GetTempPath(), "RubberduckHeadlessTests_" + Guid.NewGuid());
            Directory.CreateDirectory(baseDirectory);
            const int pid = 424242;
            var markerPath = Path.Combine(baseDirectory, $"{pid}.json");
            File.WriteAllText(markerPath, "{\"createdUtc\":\"2026-01-01T00:00:00Z\",\"runId\":\"locked\"}");

            try
            {
                using (new FileStream(markerPath, FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    var mode = new AutomationMode(new FileSystem(), baseDirectory, () => DateTime.UtcNow, pid);

                    Assert.IsFalse(mode.TryRead(out _));
                }
            }
            finally
            {
                Directory.Delete(baseDirectory, true);
            }
        }
    }
}
