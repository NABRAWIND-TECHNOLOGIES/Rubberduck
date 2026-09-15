using NUnit.Framework;
using Rubberduck.Automation;

namespace RubberduckTests.UnitTesting
{
    [TestFixture]
    [Category("HeadlessRunner")]
    public class HeadlessRunnerLogNamingTests
    {
        [Test]
        public void Suffix_AutomationActive_ReturnsDotPid()
        {
            Assert.AreEqual(".1234", AutomationLogNaming.Suffix(isActive: true, pid: 1234));
        }

        [Test]
        public void Suffix_AutomationInactive_ReturnsEmpty()
        {
            Assert.AreEqual(string.Empty, AutomationLogNaming.Suffix(isActive: false, pid: 1234));
        }

        [Test]
        public void Apply_FileNameWithSuffix_InsertsBeforeExtension()
        {
            var suffix = AutomationLogNaming.Suffix(isActive: true, pid: 4242);

            Assert.AreEqual("RubberduckLog.4242.txt", AutomationLogNaming.Apply("RubberduckLog.txt", suffix));
        }

        [Test]
        public void Apply_ArchiveFileNameWithSuffix_InsertsBeforeExtensionIdenticallyToFileName()
        {
            var suffix = AutomationLogNaming.Suffix(isActive: true, pid: 4242);

            Assert.AreEqual(@"archives\RubberduckLog.{#}.4242.txt",
                AutomationLogNaming.Apply(@"archives\RubberduckLog.{#}.txt", suffix));
        }

        [Test]
        public void Apply_EmptySuffixFromInactiveMode_ReturnsStockFileNameUnchanged()
        {
            var suffix = AutomationLogNaming.Suffix(isActive: false, pid: 4242);

            Assert.AreEqual("RubberduckLog.txt", AutomationLogNaming.Apply("RubberduckLog.txt", suffix));
        }

        [Test]
        public void Apply_UnsetNLogVariableDefault_ReturnsStockFileName()
        {
            // An NLog variable left at its declared default ("") resolves exactly like the
            // explicit inactive case above - the fail-safe-toward-GUI-behaviour guarantee (D9).
            Assert.AreEqual("RubberduckLog.txt", AutomationLogNaming.Apply("RubberduckLog.txt", string.Empty));
        }
    }
}
