# ShowToolWindows — Instant, Stack-Based Tool Window Management

**For Visual Studio 2019, 2022, and 2026**

> Solution Explorer buried again? A dozen tool windows cluttering your screen? You've switched tasks and now need to open the same six debugging panels — again?

**ShowToolWindows** adds a fast, practical tool window management layer on top of Visual Studio. No named layouts to create and maintain. No layout files to manage. No wrestling with **Window → Apply Window Layout**.

---

## What's Added to the Tools Menu

| Command | Description |
|---------|-------------|
| **Show Solution Explorer** | Brings Solution Explorer fully into view — even if it has drifted off-screen |
| **Close All Tool Windows (except Solution Explorer)** | Clears your workspace while keeping navigation accessible |
| **Close All Tool Windows** | Closes every tool window instantly — code editor tabs are never touched |
| **Stash/Restore Tool Windows** | Opens the dedicated stash management window |
| **Apply Stash ▶** | Applies any of your 10 most recent stashes in absolute mode, directly from the menu |

---

## The Stash System

Stashes are lightweight snapshots of which tool windows you want open. They live in a stack — push new stashes to the top, and pop or apply from anywhere in the list.

### Two Apply Modes — Merge and Absolute

| Mode | What Happens | How to Trigger |
|------|-------------|----------------|
| **Merge** | Stashed windows are opened; all existing windows stay open | Double-click a stash · Digit key 0–9 · Pop (Merge) |
| **Absolute** | All current tool windows close, then the stash windows open | Right-click → Apply (Absolute) · Pop (Abs) · Tools → Apply Stash submenu |

This is the key advantage over **Window → Apply Window Layout**: you choose whether to *add to* your workspace or *replace* it.

### Saving a Stash

1. Open the tool windows you want to save via Visual Studio's **View** menu
2. Click **Tools → Stash/Restore Tool Windows**
3. Press **F5** to refresh the list
4. Tick the windows you want to include
5. Click **Stash Checked** — your configuration is saved to the top of the stack

Stashes survive session restarts and are stored in Visual Studio's settings. If a stash with the same combination of tool windows already exists, you are prompted for confirmation before a duplicate is created.

### Quick Apply Without Opening the Window

The **Tools → Apply Stash** submenu shows up to 10 of your most recent stashes by name. Click any entry to apply it immediately in absolute mode — no need to open the Stash/Restore window. Perfect for fast task-switching without leaving your current code file.

---

## The Stash/Restore Window

<img width="535" height="395" alt="Stash/Restore Tool Windows interface" src="https://github.com/user-attachments/assets/df26b936-c891-4eac-88e5-4347a391b091" />

The window gives you full control over your stash stack:

- A **refreshable list** of all currently open tool windows, with tick boxes for selection
- A **stash list** with per-stash operations via right-click context menu
- **Pop (Merge)** and **Pop (Abs)** buttons that apply and delete a stash in a single action
- **Drop All** with confirmation for a clean slate

### Context Menu (Right-click Any Stash)

<img width="730" height="403" alt="Stash context menu" src="https://github.com/user-attachments/assets/d460fc83-594a-4587-8ac0-65fe44602c07" />

| Item | Description |
|------|-------------|
| **Apply (Merge)** | Open stash windows without deleting the stash or closing existing windows |
| **Apply (Absolute)** | Replace the current workspace with the stash windows |
| **Hide All ref'd by Stash** | Close all tool windows referenced in the stash |
| **Drop** | Delete the stash permanently |

---

## Keyboard Shortcuts

### Built into the Stash/Restore Window

| Key | Action |
|-----|--------|
| **F5** | Refresh the tool window list |
| **Ctrl+A** | Tick all tool windows |
| **Delete** | Drop the selected stash |
| **0–9 / NumPad 0–9** | Select the stash at that index and apply it in merge mode |

### Assign Your Own Shortcuts to Menu Commands

1. Open **Tools → Options → Environment → Keyboard**
2. Filter commands by **ScottTunstall**
3. Assign shortcuts to any of:
   - `Tools.ShowSolutionExplorer`
   - `Tools.CloseAllToolWindowsExceptSolutionExplorer`
   - `Tools.CloseAllToolWindows`
   - `Tools.StashRestoreToolWindows`

<img width="743" height="445" alt="Keyboard shortcuts configuration" src="https://github.com/user-attachments/assets/fddec3ae-009e-4a47-8fc9-bd2b450abec6" />

---

## How It Compares to Built-in Window Layouts

| | Window → Apply Window Layout | ShowToolWindows |
|---|:---:|:---:|
| Merge tool windows into existing workspace | ✗ | ✓ |
| Save a snapshot instantly without naming it | ✗ | ✓ |
| Affects tool windows only, not code editor layout | ✗ | ✓ |
| Context menu operations per saved configuration | ✗ | ✓ |
| Apply directly from the menu bar | ✗ | ✓ |
| Multiple saved configurations | ✓ | ✓ |
| Persistent across sessions | ✓ | ✓ |

Both can coexist. Use **Window → Apply Window Layout** for your base workspace arrangement, and this extension for dynamic tool window management within that layout.

---

## Real-World Use Cases

**Debugging session**
Open Output, Watch, Locals, Call Stack, and Diagnostic Tools. Stash them. Apply in merge mode whenever you start a debug session — your existing windows stay open.

**Database work**
Stash SQL Server Object Explorer, Server Explorer, and Data Sources. Apply in absolute mode to get a focused, clean workspace with no distractions.

**Lost Solution Explorer**
Run **Tools → Show Solution Explorer** — it reappears immediately, even if it has drifted off-screen or is buried under other windows.

**Clean slate**
Run **Tools → Close All Tool Windows** to reset your workspace without touching a single code editor tab.

---

## Technical Details

- **Compatibility:** Visual Studio 2019, 2022, and 2026 (x86, amd64, and arm64)
- **Architecture:** Uses Visual Studio's DTE automation layer for tool window management
- **Persistence:** Stashes are stored in Visual Studio's settings and survive session restarts
- **Scope:** Tool windows only — code editor tabs and arrangements are never modified

---

## Source and Licence

Source code: [github.com/ScottTunstall/ShowToolWindows](https://github.com/ScottTunstall/ShowToolWindows)

Developed by Scott Tunstall. Forking is permitted; creating a derivative for sale is not.
