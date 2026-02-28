# Copilot Instructions

## Project Guidelines
- In this codebase, event handlers ALWAYS call Execute methods. Do not put logic in event handlers. Event handlers should not enforce thread context.
- Execute methods should always call `ThreadHelper.ThrowIfNotOnUIThread()` first; primitive methods assume UI-thread context was already validated. Therefore, every Execute method must call `ThreadHelper.ThrowIfNotOnUIThread()` before invoking primitive methods.
- Execute methods orchestrate smaller focused methods and typically show a status bar notification afterward.
- Execute methods are the main entry points for stash functionality and should validate state before calling lower-level primitive methods.

## Documentation Standards
- Use UK English when generating XML documentation comments in this repository.
- All public methods must have XML documentation comments.
- Keep comments minimal and follow Clean Code principles in this repository.
