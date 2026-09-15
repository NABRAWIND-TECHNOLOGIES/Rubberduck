using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Automation
{
    /// <summary>
    /// A single captured modal dialog (design D13), read by the WH_CBT hook before dismissal.
    /// <see cref="Module"/>/<see cref="Line"/>/<see cref="Column"/>/<see cref="Source"/> are only
    /// populated for a compile-error dialog whose source location was successfully read from
    /// the active code pane; every other case fails safe to <c>null</c>.
    /// </summary>
    public readonly struct HeadlessDialogEntry : IEquatable<HeadlessDialogEntry>
    {
        public HeadlessDialogEntry(string caption, string text, string module, int? line, int? column, string source, DateTime timestampUtc)
        {
            Caption = caption ?? string.Empty;
            Text = text ?? string.Empty;
            Module = module;
            Line = line;
            Column = column;
            Source = source;
            TimestampUtc = timestampUtc;
        }

        public string Caption { get; }
        public string Text { get; }

        /// <summary>Null unless this was a compile-error dialog with a successfully read location.</summary>
        public string Module { get; }
        public int? Line { get; }
        public int? Column { get; }
        public string Source { get; }
        public DateTime TimestampUtc { get; }

        public bool Equals(HeadlessDialogEntry other) =>
            Caption == other.Caption && Text == other.Text && Module == other.Module &&
            Line == other.Line && Column == other.Column && Source == other.Source &&
            TimestampUtc == other.TimestampUtc;

        public override bool Equals(object obj) => obj is HeadlessDialogEntry other && Equals(other);
        public override int GetHashCode() => TimestampUtc.GetHashCode() ^ (Caption?.GetHashCode() ?? 0);
    }

    /// <summary>
    /// In-process append-only log of intercepted modal dialogs (design D13). Written by
    /// <see cref="HeadlessDialogInterceptor"/>'s hook callback and read by
    /// <c>RubberduckTestRunner</c> to correlate an entry with the test that was running when it
    /// appeared. A per-process <see cref="Shared"/> instance is used so the interceptor
    /// (installed once in <c>Extension.InitializeAddIn</c>) and the runner (constructed via the
    /// IoC container) observe the same log without new cross-cutting DI plumbing.
    /// </summary>
    public sealed class HeadlessDialogLog
    {
        public static HeadlessDialogLog Shared { get; } = new HeadlessDialogLog();

        private readonly object _sync = new object();
        private readonly List<HeadlessDialogEntry> _entries = new List<HeadlessDialogEntry>();

        public void Append(HeadlessDialogEntry entry)
        {
            lock (_sync)
            {
                _entries.Add(entry);
            }
        }

        /// <summary>
        /// Entries in the half-open window <c>[startUtcInclusive, endUtcExclusive)</c> -- matches
        /// how the runner reads a test's own <c>[TestStarted, TestCompleted)</c> window.
        /// </summary>
        public IReadOnlyList<HeadlessDialogEntry> EntriesBetween(DateTime startUtcInclusive, DateTime endUtcExclusive)
        {
            lock (_sync)
            {
                return _entries
                    .Where(entry => entry.TimestampUtc >= startUtcInclusive && entry.TimestampUtc < endUtcExclusive)
                    .ToList();
            }
        }

        public void Clear()
        {
            lock (_sync)
            {
                _entries.Clear();
            }
        }
    }
}
