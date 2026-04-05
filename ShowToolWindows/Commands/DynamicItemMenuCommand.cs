using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel.Design;

namespace ShowToolWindows.Commands
{
    /// <summary>
    /// An <see cref="OleMenuCommand"/> subclass that supports dynamic item matching
    /// for menus whose items are generated at run time.
    /// </summary>
    internal class DynamicItemMenuCommand : OleMenuCommand
    {
        private readonly Predicate<int> _matches;

        /// <summary>
        /// Initialises a new instance of the <see cref="DynamicItemMenuCommand"/> class.
        /// </summary>
        /// <param name="rootId">The base command identifier for the dynamic item range.</param>
        /// <param name="invokeHandler">The handler invoked when the user clicks a dynamic item.</param>
        /// <param name="beforeQueryStatusHandler">The handler invoked when Visual Studio queries the item status.</param>
        /// <param name="matches">A predicate that returns <c>true</c> when a command identifier falls within the valid dynamic range.</param>
        public DynamicItemMenuCommand(
            CommandID rootId,
            EventHandler invokeHandler,
            EventHandler beforeQueryStatusHandler,
            Predicate<int> matches)
            : base(invokeHandler, null, beforeQueryStatusHandler, rootId)
        {
            _matches = matches ?? throw new ArgumentNullException(nameof(matches));
        }

        /// <summary>
        /// Determines whether the specified command identifier belongs to this dynamic item range.
        /// </summary>
        /// <param name="cmdId">The command identifier to test.</param>
        /// <returns><c>true</c> if the identifier is within the dynamic range; otherwise, <c>false</c>.</returns>
        public override bool DynamicItemMatch(int cmdId)
        {
            if (_matches(cmdId))
            {
                MatchedCommandId = cmdId;
                return true;
            }

            MatchedCommandId = 0;
            return false;
        }
    }
}