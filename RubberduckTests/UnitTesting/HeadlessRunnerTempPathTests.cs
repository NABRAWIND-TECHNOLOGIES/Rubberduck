using System.IO;
using NUnit.Framework;
using Rubberduck.Resources;

namespace RubberduckTests.UnitTesting
{
    [TestFixture]
    [Category("HeadlessRunner")]
    public class HeadlessRunnerTempPathTests
    {
        [Test]
        public void ComposeTempPath_GivenRootAndPid_AppendsRubberduckAndThePidAsTwoSegments()
        {
            var result = ApplicationConstants.ComposeTempPath(@"C:\Temp\", 1234);

            Assert.AreEqual(Path.Combine(@"C:\Temp\", "Rubberduck", "1234"), result);
        }

        [Test]
        public void ComposeTempPath_DifferentPids_ProduceDifferentPaths()
        {
            var first = ApplicationConstants.ComposeTempPath(@"C:\Temp\", 111);
            var second = ApplicationConstants.ComposeTempPath(@"C:\Temp\", 222);

            Assert.AreNotEqual(first, second);
        }

        [Test]
        public void RUBBERDUCK_TEMP_PATH_MatchesComposeTempPathForTheCurrentProcess()
        {
            var expected = ApplicationConstants.ComposeTempPath(
                Path.GetTempPath(),
                System.Diagnostics.Process.GetCurrentProcess().Id);

            Assert.AreEqual(expected, ApplicationConstants.RUBBERDUCK_TEMP_PATH);
        }
    }
}
