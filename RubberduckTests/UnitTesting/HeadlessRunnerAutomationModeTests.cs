using System;
using System.Collections.Generic;
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
    }
}
