using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.UnitTesting
{
    /// <summary>
    /// Outcome of resolving a selection against the discovered test set (spec: Exact-Name
    /// Selection, Empty Selection Rejected).
    /// </summary>
    internal enum TestSelectionStatus
    {
        /// <summary>The selection matched one or more tests.</summary>
        Resolved,

        /// <summary>The selection was well-formed but matched no discovered test.</summary>
        NoMatch,

        /// <summary>The selection itself was malformed (a method without a module).</summary>
        Invalid
    }

    /// <summary>Result of <see cref="TestSelectionResolver.Resolve{T}"/>.</summary>
    internal readonly struct TestSelectionResult<T>
    {
        public TestSelectionResult(TestSelectionStatus status, IReadOnlyList<T> tests)
        {
            Status = status;
            Tests = tests ?? Array.Empty<T>();
        }

        public TestSelectionStatus Status { get; }
        public IReadOnlyList<T> Tests { get; }
    }

    /// <summary>
    /// Resolves a module/method selection against a discovered test set by exact,
    /// case-insensitive name match only (spec: Exact-Name Selection -- no substring/regex).
    ///
    /// Generic over <typeparamref name="T"/> and driven by selector delegates rather than
    /// depending on the concrete <see cref="TestMethod"/> type directly: <see cref="TestMethod"/>
    /// requires a live parsed <c>Declaration</c> graph to construct, which would force every
    /// resolver test through the full VBE/parser mock pipeline for what is otherwise pure
    /// filtering logic (design D5 / Test strategy).
    /// </summary>
    internal static class TestSelectionResolver
    {
        internal static TestSelectionResult<T> Resolve<T>(
            IEnumerable<T> discoveredTests,
            Func<T, string> moduleNameSelector,
            Func<T, string> methodNameSelector,
            string module,
            string method)
        {
            if (discoveredTests is null)
            {
                throw new ArgumentNullException(nameof(discoveredTests));
            }

            if (moduleNameSelector is null)
            {
                throw new ArgumentNullException(nameof(moduleNameSelector));
            }

            if (methodNameSelector is null)
            {
                throw new ArgumentNullException(nameof(methodNameSelector));
            }

            var hasModule = !string.IsNullOrEmpty(module);
            var hasMethod = !string.IsNullOrEmpty(method);

            // A method without a module is not a valid selection (spec: "Method without module
            // is invalid" -- there is no "search every module for this method name" mode).
            if (hasMethod && !hasModule)
            {
                return new TestSelectionResult<T>(TestSelectionStatus.Invalid, Array.Empty<T>());
            }

            IReadOnlyList<T> matches;
            if (!hasModule)
            {
                // Whole suite: no filter at all. Ignored tests are included by construction --
                // this resolver never inspects an ignore/skip flag.
                matches = discoveredTests.ToList();
            }
            else if (!hasMethod)
            {
                matches = discoveredTests
                    .Where(t => string.Equals(moduleNameSelector(t), module, StringComparison.OrdinalIgnoreCase))
                    .ToList();
            }
            else
            {
                matches = discoveredTests
                    .Where(t =>
                        string.Equals(moduleNameSelector(t), module, StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(methodNameSelector(t), method, StringComparison.OrdinalIgnoreCase))
                    .ToList();
            }

            var status = matches.Count == 0 ? TestSelectionStatus.NoMatch : TestSelectionStatus.Resolved;
            return new TestSelectionResult<T>(status, matches);
        }
    }
}
