# Show Tool Windows - Advanced Tool Window Management for Visual Studio

**For Visual Studio 2019, 2022, and 2026**

## The Problem

Visual Studio's built-in **Window ? Apply Window Layout** feature has limitations:
- **Replaces all windows** - you cannot merge layouts with your current tool windows
- **Named layouts only** - requires creating and naming layouts before use
- **Affects all windows** - includes code editor layouts, not just tool windows

During a typical development session:
- Solution Explorer disappears offscreen or gets buried under other windows
- Tool windows accumulate until your workspace becomes cluttered
- You repeatedly open the same combinations of tool windows for specific tasks (debugging, profiling, database work, etc.)

## The Solution

This extension provides **flexible, stack-based tool window management**:

### Quick Access Commands
- **Show Solution Explorer** - Instantly bring Solution Explorer fully into view, even if it's offscreen
- **Close All Tool Windows (except Solution Explorer)** - Clean your workspace while preserving navigation
- **Close All Tool Windows** - Nuclear option for complete decluttering (code windows remain untouched)

### Stash/Restore System (The Power Feature)

Unlike **Window ? Apply Window Layout**, the Stash/Restore system provides:

| Feature | Visual Studio Built-in | This Extension |
|---------|----------------------|----------------|
| **Merge tool windows** | No - replaces everything | Yes - add to current workspace |
| **Quick save without naming** | No - must create named layout | Yes - instant stash to stack |
| **Multiple saved configurations** | Yes | Yes |
| **Context menu operations** | No | Yes - apply, hide, drop |
| **Persistent across sessions** | Yes | Yes |
| **Affects code editor layout** | Yes - overwrites everything | No - tool windows only |



## Getting Started

### 1. Access the Commands

After installation, find four new commands on the **Tools** menu:

<img width="442" height="634" alt="Tools menu with Show Tool Windows commands" src="https://github.com/user-attachments/assets/b9f0f04b-ecdc-434e-8d9f-776b39fc17bd" />

### 2. Using the Stash/Restore Tool Window

Click **Tools ? Stash/Restore Tool Windows** to open the management window:

<img width="535" height="395" alt="Stash/Restore Tool Windows interface" src="https://github.com/user-attachments/assets/df26b936-c891-4eac-88e5-4347a391b091" />

## How Stashing Works

### Creating a Stash

1. **Open the tool windows you want to save** using Visual Studio's View menu
2. Click **Refresh (F5)** in the Stash/Restore window to populate the list
3. **Check the tool windows** you want to include
4. Click **Stash Checked** - your configuration is saved to the top of the stack

**Note:** Stashes persist between Visual Studio sessions until you delete them.

### Applying Stashes - Two Modes

The key advantage over Visual Studio's built-in layouts:

#### Merge Mode (Additive)
- **Double-click** a stash, or
- Right-click ? **Apply (Merge)**, or  
- Use **Pop (Merge)** for the top stash

**Result:** Tool windows from the stash are **added** to your current workspace. Existing tool windows remain open.

#### Absolute Mode (Replacement)
- Right-click ? **Apply (Absolute)**, or
- Use **Pop (Abs)** for the top stash

**Result:** Current tool windows are **closed**, then the stashed tool windows are opened. Similar to built-in layouts, but tool-windows-only.

### Managing Stashes

**Pop Operations** (Apply + Delete):
- **Pop (Merge)** - Add stashed tool windows to workspace, then delete the stash
- **Pop (Abs)** - Replace workspace with stashed tool windows, then delete the stash

**Context Menu** (Right-click any stash):

<img width="730" height="403" alt="Stash context menu" src="https://github.com/user-attachments/assets/d460fc83-594a-4587-8ac0-65fe44602c07" />

- **Apply (Merge)** - Add tool windows without deleting stash
- **Apply (Absolute)** - Replace tool windows without deleting stash
- **Hide All ref'd by Stash** - Close all tool windows referenced in the stash
- **Drop** - Delete the stash permanently

**Bulk Operations:**
- **Drop All** - Delete all stashes (confirmation required)

### Keyboard Shortcuts

The Stash/Restore window includes built-in shortcuts:
- **F5** - Refresh tool window list
- **Ctrl+A** - Check all tool windows
- **Delete** - Drop selected stash

## Recommended: Assign Shortcuts to Menu Commands

For maximum productivity, assign keyboard shortcuts to the main commands:

1. Open **Tools ? Options ? Environment ? Keyboard**
2. Search for **ScottTunstall** in the command filter
3. Assign shortcuts to:
   - `Tools.ShowSolutionExplorer`
   - `Tools.CloseAllToolWindowsExceptSolutionExplorer`
   - `Tools.CloseAllToolWindows`
   - `Tools.StashRestoreToolWindows`

<img width="743" height="445" alt="Keyboard shortcuts configuration" src="https://github.com/user-attachments/assets/fddec3ae-009e-4a47-8fc9-bd2b450abec6" />

## Use Cases

### Scenario 1: Debugging Session
**Problem:** You need Output, Watch, Locals, Call Stack, and Diagnostic Tools.  
**Solution:** Open them once, stash them. Apply the stash whenever you start debugging. Use Merge mode to keep your existing workspace intact.

### Scenario 2: Database Work  
**Problem:** You frequently need SQL Server Object Explorer, Server Explorer, and Data Sources.  
**Solution:** Stash this combination. Apply in Absolute mode to clear your workspace and focus on database tasks.

### Scenario 3: Lost Solution Explorer
**Problem:** Solution Explorer has wandered offscreen or is buried.  
**Solution:** Run **Tools ? Show Solution Explorer** - it becomes fully visible immediately.

### Scenario 4: Tool Window Overload
**Problem:** Your workspace has 15 tool windows open and you want a clean slate.  
**Solution:** Run **Tools ? Close All Tool Windows** to reset without affecting your code editor layout.

## Technical Details

- **Persistence:** Stashes are stored in Visual Studio's settings and persist across sessions
- **Scope:** Operations affect tool windows only; code editor tabs and layouts are never modified
- **Architecture:** Uses Visual Studio's DTE automation layer for tool window management
- **Compatibility:** Tested with Visual Studio 2019, 2022, and 2026 (x86 and amd64 architectures)

## Why Not Use Built-in Window Layouts?

Visual Studio's **Window ? Apply Window Layout** has its place, but this extension complements it:

| When to Use Built-in Layouts | When to Use This Extension |
|------------------------------|---------------------------|
| You need named, persistent layouts | You want quick, unnamed stashes |
| You want full window layout control (including code editor arrangement) | You only care about tool windows |
| You're switching between completely different workspace configurations | You want to add tool windows to your current workspace (Merge mode) |
| Your workflow is layout-centric | Your workflow is task-centric |

Both can coexist - use **Window ? Apply Window Layout** for your base workspace arrangement, then use this extension for dynamic tool window management within that layout.

## License

Developed by Scott Tunstall. Forking is allowed; creating derivatives for sale is forbidden.
