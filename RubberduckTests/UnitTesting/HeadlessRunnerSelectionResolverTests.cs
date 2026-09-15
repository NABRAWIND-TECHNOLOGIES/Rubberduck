using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Rubberduck.UnitTesting;

namespace RubberduckTests.UnitTesting
{
    [TestFixture]
    [Category("HeadlessRunner")]
    public class HeadlessRunnerSelectionResolverTests
    {
        // A minimal fake standing in for TestMethod (spec D5): the resolver is generic over
        // T and driven by module/method selector delegates, so it needs no VBE/parser mock.
        private readonly struct FakeTest
        {
            public FakeTest(string module, string method, bool isIgnored)
            {
                Module = module;
                Method = method;
                IsIgnored = isIgnored;
            }

            public string Module { get; }
            public string Method { get; }
            public bool IsIgnored { get; }
        }

        private static IReadOnlyList<FakeTest> Discovered() => new[]
        {
            new FakeTest("ModuleA", "TestOne", isIgnored: false),
            new FakeTest("ModuleA", "TestTwo", isIgnored: true),
            new FakeTest("ModuleB", "TestOne", isIgnored: false)
        };

        private static TestSelectionResult<FakeTest> Resolve(string module, string method) =>
            TestSelectionResolver.Resolve(Discovered(), t => t.Module, t => t.Method, module, method);

        [Test]
        public void Resolve_WholeSuite_ReturnsAllDiscoveredTests()
        {
            var result = Resolve(module: "", method: "");

            Assert.AreEqual(TestSelectionStatus.Resolved, result.Status);
            Assert.AreEqual(3, result.Tests.Count);
        }

        [Test]
        public void Resolve_OneModule_ReturnsOnlyThatModulesTests_CaseInsensitive()
        {
            var result = Resolve(module: "moduleA", method: "");

            Assert.AreEqual(TestSelectionStatus.Resolved, result.Status);
            Assert.AreEqual(2, result.Tests.Count);
            Assert.IsTrue(result.Tests.All(t => t.Module == "ModuleA"));
        }

        [Test]
        public void Resolve_OneModuleAndMethod_ReturnsOnlyThatSingleTest()
        {
            var result = Resolve(module: "ModuleB", method: "TESTONE");

            Assert.AreEqual(TestSelectionStatus.Resolved, result.Status);
            Assert.AreEqual(1, result.Tests.Count);
            Assert.AreEqual("ModuleB", result.Tests[0].Module);
            Assert.AreEqual("TestOne", result.Tests[0].Method);
        }

        [Test]
        public void Resolve_NoMatch_ReturnsEmptyResultWithNoMatchStatus()
        {
            var result = Resolve(module: "DoesNotExist", method: "");

            Assert.AreEqual(TestSelectionStatus.NoMatch, result.Status);
            Assert.AreEqual(0, result.Tests.Count);
        }

        [Test]
        public void Resolve_IgnoredTestsAreIncludedInModuleSelection()
        {
            var result = Resolve(module: "ModuleA", method: "");

            Assert.IsTrue(result.Tests.Any(t => t.IsIgnored),
                "Ignored tests must be part of the selection (spec: exact-name selection does not filter by ignore status).");
        }

        [Test]
        public void Resolve_MethodWithoutModule_ReturnsInvalidStatus()
        {
            var result = Resolve(module: "", method: "TestOne");

            Assert.AreEqual(TestSelectionStatus.Invalid, result.Status);
            Assert.AreEqual(0, result.Tests.Count);
        }
    }
}
