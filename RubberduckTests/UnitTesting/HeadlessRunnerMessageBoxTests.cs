using System;
using System.Collections.Generic;
using NUnit.Framework;
using Rubberduck.Automation;
using Rubberduck.Interaction;

namespace RubberduckTests.UnitTesting
{
    [TestFixture]
    [Category("HeadlessRunner")]
    public class HeadlessRunnerMessageBoxTests
    {
        private static (HeadlessMessageBox box, List<string> logged) CreateBox()
        {
            var logged = new List<string>();
            var box = new HeadlessMessageBox(logged.Add);
            return (box, logged);
        }

        [Test]
        public void Question_AnyText_ReturnsFalseWithoutShowingADialog()
        {
            var (box, logged) = CreateBox();

            var result = box.Question("Load legacy settings?", "Rubberduck");

            Assert.IsFalse(result);
            Assert.That(logged, Has.Exactly(1).Items);
            Assert.That(logged[0], Does.Contain("Load legacy settings?"));
        }

        [Test]
        public void ConfirmYesNo_TwoArgOverload_DefaultsToTrue()
        {
            var (box, _) = CreateBox();

            Assert.IsTrue(box.ConfirmYesNo("Proceed?", "Rubberduck"));
        }

        [TestCase(true)]
        [TestCase(false)]
        public void ConfirmYesNo_ThreeArgOverload_EchoesTheCallersSuggestion(bool suggestion)
        {
            var (box, logged) = CreateBox();

            var result = box.ConfirmYesNo("Proceed?", "Rubberduck", suggestion);

            Assert.AreEqual(suggestion, result);
            Assert.That(logged[0], Does.Contain(suggestion ? "Yes" : "No"));
        }

        [TestCase(ConfirmationOutcome.Yes)]
        [TestCase(ConfirmationOutcome.No)]
        [TestCase(ConfirmationOutcome.Cancel)]
        public void Confirm_EchoesTheCallersSuggestion(ConfirmationOutcome suggestion)
        {
            var (box, logged) = CreateBox();

            var result = box.Confirm("Continue?", "Rubberduck", suggestion);

            Assert.AreEqual(suggestion, result);
            Assert.That(logged[0], Does.Contain(suggestion.ToString()));
        }

        [Test]
        public void Message_LogsTheSuppressedText_InsteadOfShowingADialog()
        {
            var (box, logged) = CreateBox();

            box.Message("Something happened.");

            Assert.That(logged, Has.Exactly(1).Items);
            Assert.That(logged[0], Does.Contain("Something happened."));
        }

        [Test]
        public void NotifyWarn_LogsBothCaptionAndText()
        {
            var (box, logged) = CreateBox();

            box.NotifyWarn("Disk is nearly full.", "Warning");

            Assert.That(logged[0], Does.Contain("Warning"));
            Assert.That(logged[0], Does.Contain("Disk is nearly full."));
        }

        [Test]
        public void Constructor_NullLogSink_Throws()
        {
            Assert.Throws<ArgumentNullException>(() => new HeadlessMessageBox(null));
        }
    }
}
