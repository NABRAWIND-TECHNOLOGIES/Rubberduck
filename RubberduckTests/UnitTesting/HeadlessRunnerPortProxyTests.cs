using System.Linq;
using Newtonsoft.Json.Linq;
using NUnit.Framework;
using Rubberduck.UnitTesting;

namespace RubberduckTests.UnitTesting
{
    [TestFixture]
    [Category("HeadlessRunner")]
    public class HeadlessRunnerPortProxyTests
    {
        // Hotfix PR5f (design D5/P7 amendment): the VBE only accepts an AddIn.Object assignment
        // while inside OnConnection. HeadlessPortProxy is the placeholder object OnConnection
        // assigns instead of the real RubberduckTestRunner; Startup() binds the real runner into
        // it afterwards. Before Bind(), every member must return a safe value instead of
        // throwing or reaching into a null target.
        private sealed class FakeRunner : IRubberduckTestRunner
        {
            public string Version => "9.9.9-fake";
            public int SchemaVersion => 42;
            public bool IsReady => true;
            public string ParserStatus => "Ready";
            public int RequestParseCallCount { get; private set; }
            public void RequestParse() => RequestParseCallCount++;
            public int DiscoveredTestCount => 7;
            public string LastStartRunModule { get; private set; }
            public string LastStartRunMethod { get; private set; }
            public string StartRun(string ModuleName = "", string MethodName = "")
            {
                LastStartRunModule = ModuleName;
                LastStartRunMethod = MethodName;
                return "fake-run-id";
            }
            public bool IsComplete => true;
            public int CancelCallCount { get; private set; }
            public void Cancel() => CancelCallCount++;
            public string GetResults() => "{\"fake\":true}";
            public string LastError => "{\"code\":\"FAKE\",\"message\":\"fake\",\"detail\":null}";
        }

        [Test]
        public void UnboundProxy_ReturnsSafeValuesInstead_OfThrowing()
        {
            var proxy = new HeadlessPortProxy();

            Assert.IsFalse(proxy.IsReady);
            Assert.AreEqual("StartupPending", proxy.ParserStatus);
            Assert.AreEqual(0, proxy.DiscoveredTestCount);
            Assert.IsFalse(proxy.IsComplete);
            // Fork version string decision: the unbound fallback is the shared
            // AssemblyInformationalVersion (e.g. "2.5.9-headless.1"), not a numeric
            // AssemblyVersion + "-headless" suffix.
            StringAssert.Contains("-headless", proxy.Version);
            Assert.AreEqual(1, proxy.SchemaVersion);
        }

        [Test]
        public void UnboundProxy_RequestParseAndCancel_AreNoOps()
        {
            var proxy = new HeadlessPortProxy();

            Assert.DoesNotThrow(() => proxy.RequestParse());
            Assert.DoesNotThrow(() => proxy.Cancel());
        }

        [Test]
        public void UnboundProxy_GetResults_ReturnsValidIdleEnvelope()
        {
            var proxy = new HeadlessPortProxy();

            var json = JObject.Parse(proxy.GetResults());

            Assert.AreEqual(1, json["schemaVersion"].Value<int>());
            Assert.AreEqual("idle", json["status"].Value<string>());
            Assert.AreEqual(0, json["tests"].Count());
            Assert.AreEqual(0, json["summary"]["total"].Value<int>());
        }

        [Test]
        public void UnboundProxy_StartRun_ReturnsEmptyAndSetsNotReadyError()
        {
            var proxy = new HeadlessPortProxy();

            var runId = proxy.StartRun("SomeModule", "SomeMethod");

            Assert.AreEqual(string.Empty, runId);
            var error = JObject.Parse(proxy.LastError);
            Assert.AreEqual("NOT_READY", error["code"].Value<string>());
            Assert.AreEqual("Rubberduck startup has not completed", error["message"].Value<string>());
        }

        [Test]
        public void BoundProxy_DelegatesEveryMember_ToTheTarget()
        {
            var proxy = new HeadlessPortProxy();
            var fake = new FakeRunner();

            proxy.Bind(fake);

            Assert.AreEqual(fake.Version, proxy.Version);
            Assert.AreEqual(fake.SchemaVersion, proxy.SchemaVersion);
            Assert.AreEqual(fake.IsReady, proxy.IsReady);
            Assert.AreEqual(fake.ParserStatus, proxy.ParserStatus);
            Assert.AreEqual(fake.DiscoveredTestCount, proxy.DiscoveredTestCount);
            Assert.AreEqual(fake.IsComplete, proxy.IsComplete);
            Assert.AreEqual(fake.GetResults(), proxy.GetResults());
            Assert.AreEqual(fake.LastError, proxy.LastError);

            proxy.RequestParse();
            Assert.AreEqual(1, fake.RequestParseCallCount);

            proxy.Cancel();
            Assert.AreEqual(1, fake.CancelCallCount);

            var runId = proxy.StartRun("ModuleA", "MethodB");
            Assert.AreEqual("fake-run-id", runId);
            Assert.AreEqual("ModuleA", fake.LastStartRunModule);
            Assert.AreEqual("MethodB", fake.LastStartRunMethod);
        }

        [Test]
        public void UnboundProxy_AfterStartupFailureReported_SurfacesStartupFailedError()
        {
            var proxy = new HeadlessPortProxy();

            proxy.NotifyStartupFailed("Rubberduck's startup sequence threw an unexpected exception.");

            var error = JObject.Parse(proxy.LastError);
            Assert.AreEqual("STARTUP_FAILED", error["code"].Value<string>());
            Assert.AreEqual("Rubberduck's startup sequence threw an unexpected exception.", error["detail"].Value<string>());

            var results = JObject.Parse(proxy.GetResults());
            Assert.AreEqual("STARTUP_FAILED", results["error"]["code"].Value<string>());
        }

        [Test]
        public void BoundProxy_IgnoresStartupFailureNotification_TargetErrorWins()
        {
            var proxy = new HeadlessPortProxy();
            var fake = new FakeRunner();
            proxy.Bind(fake);

            proxy.NotifyStartupFailed("should be ignored once bound");

            Assert.AreEqual(fake.LastError, proxy.LastError);
        }
    }
}
