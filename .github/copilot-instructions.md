# Copilot Instructions

## Project Guidelines
- In this codebase, event handlers ALWAYS call Execute methods. Do not put logic in event handlers. event handlers should not enforce thread context.
- Execute methods should always call `ThreadHelper.ThrowIfNotOnUIThread()` first; 
- Execute methods orchestrate smaller focused methods and typically show a status bar notification afterward.
- Execute methods are the main entry points for stash functionality and should validate state before calling lower-level primitive methods.
