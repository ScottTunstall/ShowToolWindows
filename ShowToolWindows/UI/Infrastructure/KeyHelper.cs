using System.Windows.Input;

namespace ShowToolWindows.UI.Infrastructure
{
    /// <summary>
    /// Provides helper methods for working with keyboard input.
    /// </summary>
    internal static class KeyHelper
    {
        /// <summary>
        /// Attempts to map a key press to a numeric digit (0-9).
        /// </summary>
        /// <param name="key">The key that was pressed.</param>
        /// <param name="digit">The parsed digit if successful.</param>
        /// <returns>True if the key maps to a digit; otherwise, false.</returns>
        public static bool TryGetDigitFromKey(Key key, out int digit)
        {
            if (key >= Key.D0 && key <= Key.D9)
            {
                digit = key - Key.D0;
                return true;
            }

            if (key >= Key.NumPad0 && key <= Key.NumPad9)
            {
                digit = key - Key.NumPad0;
                return true;
            }

            digit = -1;
            return false;
        }
    }
}