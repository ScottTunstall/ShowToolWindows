
## Documentation

- Whenever `readme.md` in the repository root is updated, also update `marketplace-markdown.md` in the repository root to reflect the same changes. The marketplace version must be tailored for the Visual Studio Marketplace: lead with a developer-focused value proposition, use benefit-led language, include a tick/cross feature comparison table vs Visual Studio built-in layouts, and present commands, modes, keyboard shortcuts, and use cases in a format that attracts developers browsing the Marketplace.

## Visual Studio Extension Development

- OBEY Visual Studio UX patterns (menus, commands, window behaviour): https://learn.microsoft.com/en-us/visualstudio/extensibility/ux-guidelines/application-patterns-for-visual-studio
- Use AsyncPackage, async initialization, and async service retrieval.
- Prefer IAsyncServiceProvider and GetServiceAsync over synchronous calls.
- Use WPF for tool window UI (no WinForms).
- Use ToolWindowPane for tool windows and OleMenuCommandService for commands.
- Tool windows should be shown/hidden, not destroyed.
- Use IVsWindowFrame for visibility control.
- Avoid blocking the UI thread.
- Keep long‑running work off the main thread.

## General Coding Practices

- Use C# compatible with .NET Framework 4.7.2
- OBEY .editorconfig settings.
- Follow .NET naming conventions (PascalCase for types, camelCase for locals).
- Keep methods short and focused.
- Avoid static mutable state.
- Prefer immutability where possible.
- Use XML doc comments for classes. Public properties and methods must have XML docs.
- Comment why, not what.
