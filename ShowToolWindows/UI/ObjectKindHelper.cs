using System;

namespace ShowToolWindows.UI
{
    /// <summary>
    /// Provides helper methods for working with Visual Studio window object kinds.
    /// </summary>
    internal static class ObjectKindHelper
    {
        /// <summary>
        /// Normalizes an object kind string by ensuring GUIDs have enclosing curly braces.
        /// </summary>
        /// <param name="objectKind">The object kind string to normalize.</param>
        /// <returns>The normalized object kind string with curly braces if it's a valid GUID.</returns>
        public static string NormalizeObjectKind(string objectKind)
        {
            if (string.IsNullOrWhiteSpace(objectKind))
            {
                return objectKind;
            }

            // Check if it's already in GUID format with curly braces
            if (objectKind.StartsWith("{") && objectKind.EndsWith("}"))
            {
                return objectKind;
            }

            // Try to parse as GUID and add curly braces if successful
            if (Guid.TryParse(objectKind, out Guid guid))
            {
                return "{" + guid.ToString().ToUpperInvariant() + "}";
            }

            return objectKind;
        }
    }
}
