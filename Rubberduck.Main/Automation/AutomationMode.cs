using System;
using System.Diagnostics;
using System.Globalization;
using System.IO.Abstractions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Rubberduck.Automation
{
    /// <summary>
    /// A parsed automation marker: the CLI's declared run start time and run id (design D8).
    /// </summary>
    public readonly struct AutomationMarker
    {
        public AutomationMarker(DateTime createdUtc, string runId)
        {
            CreatedUtc = createdUtc;
            RunId = runId;
        }

        public DateTime CreatedUtc { get; }
        public string RunId { get; }
    }

    /// <summary>
    /// Detects whether the current process is being driven by the headless automation CLI, via a
    /// PID-keyed marker file the CLI writes before dispatching Excel and deletes in its
    /// <c>finally</c> block (design D8). With no marker present, every path resolves the same as
    /// stock Rubberduck - the detection is fail-safe toward the interactive GUI by construction.
    /// File system, clock, and PID are all injected so this class needs no real file system,
    /// wall clock, or running process to be unit-tested.
    /// </summary>
    public class AutomationMode
    {
        private static readonly TimeSpan StaleAfter = TimeSpan.FromMinutes(10);

        private readonly IFileSystem _fileSystem;
        private readonly string _baseDirectory;
        private readonly Func<DateTime> _clock;
        private readonly int _pid;

        public AutomationMode(IFileSystem fileSystem, string baseDirectory, Func<DateTime> clock, int pid)
        {
            _fileSystem = fileSystem ?? throw new ArgumentNullException(nameof(fileSystem));
            _baseDirectory = baseDirectory ?? throw new ArgumentNullException(nameof(baseDirectory));
            _clock = clock ?? throw new ArgumentNullException(nameof(clock));
            _pid = pid;
        }

        /// <summary>
        /// Real-environment convenience instance, wired to the actual file system, UTC clock and
        /// current process id. Never needed by unit tests - those construct
        /// <see cref="AutomationMode"/> directly with a <c>MockFileSystem</c> and a fixed clock.
        /// </summary>
        public static AutomationMode Current { get; } = CreateCurrent();

        private static AutomationMode CreateCurrent()
        {
            var fileSystem = new FileSystem();
            var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var baseDirectory = fileSystem.Path.Combine(localAppData, "Rubberduck", "automation");
            return new AutomationMode(fileSystem, baseDirectory, () => DateTime.UtcNow, Process.GetCurrentProcess().Id);
        }

        /// <summary>
        /// <c>true</c> when a fresh (under 10 minutes old), well-formed marker exists for the
        /// current process id.
        /// </summary>
        public bool IsActive => TryRead(out _);

        /// <summary>
        /// Startup failure recorded under automation (design "Headless UI Suppression" / D8):
        /// a condition that would have shown a startup <c>MessageBox</c> instead sets this and
        /// logs <c>Fatal</c>, so the test runner port's <c>LastError</c> can surface
        /// <c>STARTUP_FAILED</c> to the CLI instead of a dialog no one is present to dismiss.
        /// <c>null</c> means startup has not (yet) failed. Process-wide by design: there is
        /// exactly one add-in startup sequence per Excel process.
        /// </summary>
        public static string StartupError { get; set; }

        /// <summary>
        /// Reads and validates this process's marker file, if any. Returns <c>false</c> - never
        /// throws - for an absent, stale, malformed, or unreadable marker, which is the fail-safe
        /// behaviour the port relies on to fall back to stock (interactive) behaviour.
        /// </summary>
        public bool TryRead(out AutomationMarker marker)
        {
            marker = default;

            var path = _fileSystem.Path.Combine(_baseDirectory, $"{_pid}.json");
            if (!_fileSystem.File.Exists(path))
            {
                return false;
            }

            string json;
            try
            {
                json = _fileSystem.File.ReadAllText(path);
            }
            catch (System.IO.IOException)
            {
                return false;
            }
            catch (UnauthorizedAccessException)
            {
                return false;
            }

            if (!TryParse(json, out marker))
            {
                return false;
            }

            var age = _clock() - marker.CreatedUtc;
            return age >= TimeSpan.Zero && age < StaleAfter;
        }

        private static bool TryParse(string json, out AutomationMarker marker)
        {
            marker = default;
            try
            {
                // DateParseHandling.None keeps "createdUtc" as a raw string token instead of
                // letting Json.NET auto-convert it to a DateTime tagged with Local/Unspecified
                // kind based on the machine's time zone. DateTime.SpecifyKind on an
                // already-zone-adjusted value would silently relabel a wall-clock time as UTC
                // without converting it -- wrong by exactly the local UTC offset whenever the
                // marker carries an explicit non-"Z" offset (e.g. "+02:00").
                JObject envelope;
                using (var reader = new JsonTextReader(new System.IO.StringReader(json)) { DateParseHandling = DateParseHandling.None })
                {
                    envelope = JObject.Load(reader);
                }

                var createdUtcText = envelope.Value<string>("createdUtc");
                if (string.IsNullOrEmpty(createdUtcText))
                {
                    return false;
                }

                if (!DateTime.TryParse(
                        createdUtcText,
                        CultureInfo.InvariantCulture,
                        DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal,
                        out var createdUtc))
                {
                    return false;
                }

                var runId = envelope.Value<string>("runId") ?? string.Empty;
                marker = new AutomationMarker(createdUtc, runId);
                return true;
            }
            catch (Exception ex) when (ex is JsonException || ex is FormatException)
            {
                return false;
            }
        }
    }
}
