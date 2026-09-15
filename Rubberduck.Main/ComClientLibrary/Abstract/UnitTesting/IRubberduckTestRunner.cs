using System.ComponentModel;
using System.Runtime.InteropServices;
using Rubberduck.Resources.Registration;

// ReSharper disable InconsistentNaming
// The parameters on RD's public interfaces follow VBA conventions, not C# conventions --
// see IAssert.cs for the same rationale.

namespace Rubberduck.UnitTesting
{
    /// <summary>
    /// COM-visible headless test-run automation port (design D5). Depends on nothing but
    /// <c>ITestEngine</c> through its implementation; every member here is safe to call from
    /// any run state -- see <c>RubberduckTestRunner</c> for the state machine and threading
    /// contract.
    /// </summary>
    /// <remarks>
    /// GUID/ProgId <em>registration and wiring</em> (IoC container, <c>Extension.cs</c> Object
    /// assignment) are added in the wiring slice (PR5, tasks 4.1-4.4). The
    /// <see cref="RubberduckGuid.ITestRunnerGuid"/> constant itself had to be added alongside
    /// this file because the fork's ComVisibleTypeAnalyzer requires every COM-visible type's
    /// Guid attribute to reference a RubberduckGuid constant by source text -- a literal GUID
    /// string does not satisfy it.
    /// </remarks>
    [
        ComVisible(true),
        Guid(RubberduckGuid.ITestRunnerGuid),
        InterfaceType(ComInterfaceType.InterfaceIsDual),
        EditorBrowsable(EditorBrowsableState.Always)
    ]
    public interface IRubberduckTestRunner
    {
        /// <summary>Assembly version plus the fork's "-headless" suffix.</summary>
        [DispId(1)]
        string Version { get; }

        /// <summary>The JSON result envelope's schema version (design D4).</summary>
        [DispId(2)]
        int SchemaVersion { get; }

        /// <summary>True when the underlying engine can accept a new run (<c>engine.CanRun</c>).</summary>
        [DispId(3)]
        bool IsReady { get; }

        /// <summary>The parser's current state, when a status source was supplied.</summary>
        [DispId(4)]
        string ParserStatus { get; }

        /// <summary>Requests a parser re-run, when a parse trigger was supplied.</summary>
        [DispId(5)]
        void RequestParse();

        /// <summary>The number of tests currently discovered by the engine.</summary>
        [DispId(6)]
        int DiscoveredTestCount { get; }

        /// <summary>
        /// Starts a run for the given exact-name selection and returns immediately with the
        /// run id, or <c>""</c> if the run was rejected or matched no tests -- see
        /// <see cref="LastError"/> for the reason (spec: Asynchronous Run Control).
        /// </summary>
        /// <param name="ModuleName">Exact module name, or "" for the whole suite.</param>
        /// <param name="MethodName">Exact method name; requires <paramref name="ModuleName"/>.</param>
        [DispId(7)]
        string StartRun(string ModuleName = "", string MethodName = "");

        /// <summary>True once the current (or most recent) run has reached a terminal state.</summary>
        [DispId(8)]
        bool IsComplete { get; }

        /// <summary>Requests cancellation of the run in progress; a no-op otherwise.</summary>
        [DispId(9)]
        void Cancel();

        /// <summary>
        /// Returns the JSON result envelope for the current or most recent run. Safe to call
        /// at any time, including mid-run: it returns an immutable snapshot and never mutates
        /// state (design "Threading model and thread safety").
        /// </summary>
        [DispId(10)]
        string GetResults();

        /// <summary>The last structured error as JSON <c>{code,message,detail}</c>, or "".</summary>
        [DispId(11)]
        string LastError { get; }
    }
}
