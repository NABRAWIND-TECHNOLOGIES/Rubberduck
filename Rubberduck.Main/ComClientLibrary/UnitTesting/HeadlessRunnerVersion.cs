using System.Reflection;

namespace Rubberduck.UnitTesting
{
    /// <summary>
    /// Shared version string for the headless automation port (fork version string decision):
    /// the "Rubberduck" assembly's <see cref="AssemblyInformationalVersionAttribute"/>, e.g.
    /// <c>"2.5.9-headless.1"</c>. Set fork-wide via <c>RubberduckBaseProject.csproj</c>'s
    /// <c>InformationalVersion</c> MSBuild property, which leaves the numeric
    /// <c>AssemblyVersion</c>/<c>AssemblyFileVersion</c> (driven by the same project's
    /// <c>Version</c> property) untouched. Both <see cref="RubberduckTestRunner"/> and
    /// <see cref="HeadlessPortProxy"/>'s unbound fallback read this single source so the two
    /// never drift apart.
    /// </summary>
    internal static class HeadlessRunnerVersion
    {
        public static string Current { get; } =
            Assembly.GetExecutingAssembly().GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion
            ?? Assembly.GetExecutingAssembly().GetName().Version.ToString();
    }
}
