using Castle.Windsor;
using Extensibility;
using NLog;
using Rubberduck.Automation;
using Rubberduck.Common.WinAPI;
using Rubberduck.Resources;
using Rubberduck.Resources.Registration;
using Rubberduck.Root;
using Rubberduck.Runtime;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.UI;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.ComManagement.TypeLibs;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.VbeRuntime;
using Rubberduck.VersionCheck;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO.Abstractions;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Threading;

namespace Rubberduck
{
    /// <remarks>
    /// Special thanks to Carlos Quintero (MZ-Tools) for providing the general structure here.
    /// </remarks>
    [
        ComVisible(true),
        Guid(RubberduckGuid.ExtensionGuid),
        ProgId(RubberduckProgId.ExtensionProgId),
        ClassInterface(ClassInterfaceType.None),
        ComDefaultInterface(typeof(IDTExtensibility2)),
        EditorBrowsable(EditorBrowsableState.Never)
    ]
    // ReSharper disable once InconsistentNaming // note: underscore prefix hides class from COM API
    public class _Extension : IDTExtensibility2
    {
        private IVBE _vbe;
        private IAddIn _addin;
        private IVbeNativeApi _vbeNativeApi;
        private IBeepInterceptor _beepInterceptor;
        private IFileSystem _fileSystem;
        private bool _isInitialized;
        private bool _isBeginShutdownExecuted;

        private GeneralSettings _initialSettings;

        private IWindsorContainer _container;
        private App _app;
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        // Hotfix (PR5f, design D5/P7 amendment): placeholder assigned to _addin.Object inside
        // OnConnection -- see SetAddInObject and Startup() for why the real runner can no longer
        // be assigned there directly.
        private HeadlessPortProxy _portProxy;
        private bool _portProxyIsAddInObject;

        public void OnAddInsUpdate(ref Array custom) { }

        [SuppressMessage("ReSharper", "InconsistentNaming")]
        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                _vbe = RootComWrapperFactory.GetVbeWrapper(Application);
                _addin = RootComWrapperFactory.GetAddInWrapper(AddInInst);

                // Hotfix (PR5f, design D5/P7 amendment): the VBE only accepts an AddIn.Object
                // assignment while inside OnConnection -- assigning the headless runner later,
                // in Startup(), throws COMException E_FAIL and broke normal GUI startup for
                // every user (confirmed by direct user report, 2026-09-15). The proxy is
                // assigned here, once; Startup() binds the real runner into it afterwards
                // instead of reassigning .Object. If the proxy assignment itself ever fails,
                // fall back to stock's own placeholder so the GUI can never break on this line.
                _portProxy = new HeadlessPortProxy();
                try
                {
                    _addin.Object = _portProxy;
                    _portProxyIsAddInObject = true;
                }
                catch (Exception ex)
                {
                    _logger.Warn(ex, "Failed to assign the headless port proxy as the add-in Object; falling back to the interactive object.");
                    _portProxyIsAddInObject = false;
                    _addin.Object = this;
                }

                _vbeNativeApi = new VbeNativeApiAccessor();
                _beepInterceptor = new BeepInterceptor(_vbeNativeApi);
                _fileSystem = new FileSystem();
                VbeProvider.Initialize(_vbe, _vbeNativeApi, _beepInterceptor);
                VbeNativeServices.HookEvents(_vbe);

                SetAddInObject();

                switch (ConnectMode)
                {
                    case ext_ConnectMode.ext_cm_Startup:
                        // normal execution path - don't initialize just yet, wait for OnStartupComplete to be called by the host.
                        break;
                    case ext_ConnectMode.ext_cm_AfterStartup:
                        _isBeginShutdownExecuted = false;   //When we reconnect after having been unloaded, the variable might no longer have its initial value.
                        InitializeAddIn();
                        break;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        [Conditional("DEBUG")]
        private void SetAddInObject()
        {
            // Hotfix (PR5f, design D5/P7 amendment): a Debug build IS the production install
            // (design D11), so this DEBUG-only convenience must never override the headless port
            // proxy assigned above -- doing so would silently discard the proxy and reintroduce
            // the exact GUI startup break this hotfix exists to fix.
            if (_portProxyIsAddInObject)
            {
                return;
            }

            // FOR DEBUGGING/DEVELOPMENT PURPOSES, ALLOW ACCESS TO SOME VBETypeLibsAPI FEATURES FROM VBA
            _addin.Object = new VBETypeLibsAPI_Object(_vbe);
        }

        private Assembly LoadFromSameFolder(object sender, ResolveEventArgs args)
        {
            var folderPath = _fileSystem.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? string.Empty;
            var assemblyPath = _fileSystem.Path.Combine(folderPath, new AssemblyName(args.Name).Name + ".dll");
            if (!_fileSystem.File.Exists(assemblyPath))
            {
                return null;
            }

            var assembly = Assembly.LoadFile(assemblyPath);
            return assembly;
        }

        public void OnStartupComplete(ref Array custom)
        {
            InitializeAddIn();
        }

        public void OnBeginShutdown(ref Array custom)
        {
            _isBeginShutdownExecuted = true;
            ShutdownAddIn();
        }

        // ReSharper disable InconsistentNaming
        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            switch (RemoveMode)
            {
                case ext_DisconnectMode.ext_dm_UserClosed:
                    ShutdownAddIn();
                    break;

                case ext_DisconnectMode.ext_dm_HostShutdown:
                    if (_isBeginShutdownExecuted)
                    {
                        // this is the normal case: nothing to do here, we already ran ShutdownAddIn.
                    }
                    else
                    {
                        // some hosts do not call OnBeginShutdown: this mitigates it.
                        ShutdownAddIn();
                    }
                    break;
            }
        }

        private void InitializeAddIn()
        {
            Splash2021 splash = null;
            // Read exactly once per call so every decision below (log filename, splash, dialog
            // suppression) agrees on the same snapshot of automation state (design D8/D9).
            // Guarded by its own try/catch (design D8: fail-safe) so a failure while probing
            // automation state can never escape InitializeAddIn's startup guard -- it just
            // resolves to interactive/stock behaviour, the same as "no marker present".
            var isAutomationActive = false;
            try
            {
                isAutomationActive = AutomationMode.Current.IsActive;
            }
            catch (Exception ex)
            {
                _logger.Warn(ex, "Automation mode detection failed; assuming interactive session.");
            }
            try
            {
                if (_isInitialized)
                {
                    // The add-in is already initialized. See:
                    // The strange case of the add-in initialized twice
                    // http://msmvps.com/blogs/carlosq/archive/2013/02/14/the-strange-case-of-the-add-in-initialized-twice.aspx
                    return;
                }

                // Before the first log write (design D9 item 2): under automation, every log
                // line for this process lands in its own RubberduckLog.<pid>.txt instead of the
                // shared RubberduckLog.txt a concurrent instance would otherwise wipe on startup.
                // An untouched GUI session resolves the variable to "" -- byte-identical to stock.
                if (LogManager.Configuration != null)
                {
                    LogManager.Configuration.Variables["automationSuffix"] =
                        AutomationLogNaming.Suffix(isAutomationActive, Process.GetCurrentProcess().Id);
                    LogManager.ReconfigExistingLoggers();
                }

                var pathProvider = PersistencePathProvider.Instance;
                var configLoader = new XmlPersistenceService<GeneralSettings>(pathProvider, _fileSystem);
                var configProvider = new GeneralConfigProvider(configLoader);

                _initialSettings = configProvider.Read();
                if (_initialSettings != null)
                {
                    try
                    {
                        var cultureInfo = CultureInfo.GetCultureInfo(_initialSettings.Language.Code);
                        Dispatcher.CurrentDispatcher.Thread.CurrentUICulture = cultureInfo;
                    }
                    catch (CultureNotFoundException)
                    {
                    }

                    try
                    {
                        if (_initialSettings.SetDpiUnaware)
                        {
                            SHCore.SetProcessDpiAwareness(PROCESS_DPI_AWARENESS.Process_DPI_Unaware);
                        }
                    }
                    catch (Exception)
                    {
                        Debug.Assert(false, "Could not set DPI awareness.");
                    }
                }
                else
                {
                    Debug.Assert(false, "Settings could not be initialized.");
                }

                // No splash under automation (design D8/P6): nothing is present to see or
                // dismiss it, and it costs a real UI window creation on the STA thread the CLI
                // is trying to keep invisible.
                if (!isAutomationActive && (_initialSettings?.CanShowSplash ?? false))
                {
                    splash = new Splash2021(string.Format(RubberduckUI.Rubberduck_AboutBuild, Assembly.GetExecutingAssembly().GetName().Version.ToString(3)));
                    splash.Show();
                    splash.Refresh();
                }

                Startup();
            }
            catch (Win32Exception ex)
            {
                if (isAutomationActive)
                {
                    _logger.Fatal(ex, "Rubberduck reload failed; suppressed under automation.");
                    AutomationMode.StartupError = "STARTUP_FAILED";
                    _portProxy?.NotifyStartupFailed(ex.Message);
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show(Resources.RubberduckUI.RubberduckReloadFailure_Message,
                        RubberduckUI.RubberduckReloadFailure_Title,
                        MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
            }
            catch (Exception exception)
            {
                _logger.Fatal(exception);
                if (isAutomationActive)
                {
                    // No dialog to show, and no one present to dismiss it (design "Headless UI
                    // Suppression" / D8): the failure is already logged above and now also
                    // surfaced through the port's LastError instead.
                    AutomationMode.StartupError = "STARTUP_FAILED";
                    _portProxy?.NotifyStartupFailed(exception.Message);
                }
                else
                {
                    // TODO Use Rubberduck Interaction instead and provide exception stack trace as
                    // an optional "more info" collapsible section to eliminate the conditional.
                    MessageBox.Show(
#if DEBUG
                        exception.ToString(),
#else
                        exception.Message.ToString(),
#endif
                        RubberduckUI.RubberduckLoadFailure, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            finally
            {
                splash?.Dispose();
            }
        }

        private void Startup()
        {
            try
            {
                var currentDomain = AppDomain.CurrentDomain;
                currentDomain.UnhandledException += HandleAppDomainException;
                currentDomain.AssemblyResolve += LoadFromSameFolder;

                _container = new WindsorContainer().Install(new RubberduckIoCInstaller(_vbe, _addin, _initialSettings, _vbeNativeApi, _beepInterceptor));
                _container.Resolve<InstanceProvider>();
                _app = _container.Resolve<App>();
                _app.Startup();

                // Headless automation port (design D5/P7, amended by hotfix PR5f): the real
                // runner is bound into the proxy already assigned to _addin.Object during
                // OnConnection, rather than assigned to .Object here directly -- the VBE rejects
                // an AddIn.Object assignment made outside OnConnection with COMException E_FAIL.
                // A client reading Application.VBE.AddIns(...).Object before this line runs
                // still observes the proxy's own safe pre-bind values, never a half-initialized
                // runner.
                if (_portProxyIsAddInObject)
                {
                    _portProxy.Bind(_container.Resolve<IRubberduckTestRunner>());
                }

                _isInitialized = true;
            }
            catch (Exception e)
            {
                _logger.Fatal(e, "Startup sequence threw an unexpected exception.");
                throw new Exception("Rubberduck's startup sequence threw an unexpected exception. Please check the Rubberduck logs for more information and report an issue if necessary", e);
            }
        }

        private void HandleAppDomainException(object sender, UnhandledExceptionEventArgs e)
        {
            var message = e.IsTerminating
                ? "An unhandled exception occurred. The runtime is shutting down."
                : "An unhandled exception occurred. The runtime continues running.";
            if (e.ExceptionObject is Exception exception)
            {
                _logger.Fatal(exception, message);

            }
            else
            {
                _logger.Fatal(message);
            }
        }

        private void ShutdownAddIn()
        {
            var currentDomain = AppDomain.CurrentDomain;
            try
            {
                _logger.Info("Rubberduck is shutting down.");
                _logger.Trace("Unhooking VBENativeServices events...");
                VbeNativeServices.UnhookEvents();
                VbeProvider.Terminate();

                _logger.Trace("Releasing dockable hosts...");

                using (var windows = _vbe.Windows)
                {
                    windows.ReleaseDockableHosts();
                }

                if (_app != null)
                {
                    _logger.Trace("Initiating App.Shutdown...");
                    _app.Shutdown();
                    _app = null;
                }

                if (_container != null)
                {
                    _logger.Trace("Disposing IoC container...");
                    _container.Dispose();
                    _container = null;
                }
            }
            catch (Exception e)
            {
                _logger.Error(e);
                _logger.Warn("Exception is swallowed.");
                //throw; // <<~ uncomment to crash the process
            }
            finally
            {
                try
                {
                    _logger.Trace("Disposing COM safe...");
                    ComSafeManager.DisposeAndResetComSafe();
                    _addin = null;
                    _vbe = null;

                    _isInitialized = false;
                    _logger.Info("No exceptions were thrown.");
                }
                catch (Exception e)
                {
                    _logger.Error(e);
                    _logger.Warn("Exception disposing the ComSafe has been swallowed.");
                    //throw; // <<~ uncomment to crash the process
                }
                finally
                {
                    _logger.Trace("Unregistering AppDomain handlers....");
                    currentDomain.AssemblyResolve -= LoadFromSameFolder;
                    currentDomain.UnhandledException -= HandleAppDomainException;
                    _logger.Trace("Done. Main Shutdown completed. Toolwindows follow. Quack!");
                    _isInitialized = false;
                }
            }
        }
    }
}
