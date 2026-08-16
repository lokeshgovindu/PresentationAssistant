namespace PresentationAssistant
{
    /// <summary>
    /// Product identity, kept apart from the package so the parts of the extension that
    /// are plain logic do not have to reference the Visual Studio SDK to know their own
    /// name. That is what lets the test project compile them on their own.
    /// </summary>
    internal static class Product
    {
        public const string Name = "PresentationAssistant";

        /// <summary>Prefix used in the status bar, where space is tight.</summary>
        public const string ShortName = "PA";
    }
}
