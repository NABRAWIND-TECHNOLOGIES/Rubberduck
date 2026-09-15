namespace Rubberduck.Automation
{
    /// <summary>
    /// Pure caption matching for locating a VBE <c>CommandBarControl</c> by its display caption
    /// (design D13 addendum -- locating the VBE "Reset" command). Office/VBE menu captions carry
    /// an accelerator marker (<c>&amp;</c>, e.g. "&amp;Reset") that must be stripped before any
    /// comparison, and localized builds may differ in case.
    /// </summary>
    public static class VbeMenuCaptionMatcher
    {
        /// <summary>The normalized caption equals the normalized wanted text exactly.</summary>
        public static bool Matches(string caption, string wantedCaption) =>
            Normalize(caption) == Normalize(wantedCaption);

        /// <summary>The normalized caption contains the normalized wanted text as a substring.</summary>
        public static bool ContainsWord(string caption, string wantedWord) =>
            Normalize(caption).Contains(Normalize(wantedWord));

        private static string Normalize(string text) =>
            (text ?? string.Empty).Replace("&", string.Empty).Trim().ToLowerInvariant();
    }
}
