using System.Reflection;
using NUnit.Framework;
using Rubberduck.UnitTesting;

namespace RubberduckTests.UnitTesting
{
    // Fork version string decision (headless-test-runner): AssemblyInformationalVersion is
    // "2.5.9-headless.1" (set fork-wide via RubberduckBaseProject.csproj), while AssemblyVersion
    // stays purely numeric. HeadlessRunnerVersion.Current is the single shared source both
    // RubberduckTestRunner.Version and HeadlessPortProxy's unbound fallback read from.
    [TestFixture]
    [Category("HeadlessRunner")]
    public class HeadlessRunnerVersionTests
    {
        [Test]
        public void Current_EqualsTheRubberduckAssemblyInformationalVersion()
        {
            var expected = typeof(RubberduckTestRunner).Assembly
                .GetCustomAttribute<AssemblyInformationalVersionAttribute>()
                ?.InformationalVersion;

            Assert.IsNotNull(expected, "The Rubberduck assembly must carry an AssemblyInformationalVersion attribute.");
            Assert.AreEqual(expected, HeadlessRunnerVersion.Current);
        }

        [Test]
        public void Current_ContainsTheHeadlessMarker()
        {
            StringAssert.Contains("-headless", HeadlessRunnerVersion.Current);
        }

        [Test]
        public void RubberduckTestRunner_Version_MatchesHeadlessRunnerVersionCurrent()
        {
            using (var mocked = new MockedTestEngine(0))
            {
                var runner = new RubberduckTestRunner(mocked.TestEngine);

                Assert.AreEqual(HeadlessRunnerVersion.Current, runner.Version);
            }
        }
    }
}
