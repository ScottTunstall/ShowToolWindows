# Show Tool Window Extension for Visual Studio 2019, 2022 and 2026.

## About

Are you using Visual Studio 2019, 2022 or 2026?

Does your Solution Explorer end up wandering off screen or get stuck under a pile of windows and you have a torrid time trying to find it?
Or do you end up with a load of Tool Windows open and you don't want to reset window layout but would like to close them all?

Fret no more. Reign in naughty tool windows with this extension.


## How to use

After installing the extension, you should see four new commands (circled in image) on the Visual Studio **Tools** menu :

<img width="451" height="269" alt="image" src="https://github.com/user-attachments/assets/c0f7291e-b79b-4197-a9b2-76c77b28dfe3" />



* The first command, **Show Solution Explorer** will make the solution explorer FULLY VISIBLE even if it was partially or mostly offscreen before.
* The second command, **Close all Tool Windows except Solution Explorer** will close all tool windows except the solution explorer - very useful when you're in Tool Window Hell. 
* The third command, **Close all Tool Windows** is the nuclear option and does as you expect. Every tool window goes, but your code windows will be left untouched. 
* The fourth command, **Show/Hide/Stash/Restore tool windows** is for power users. This pops up a floating toolbar window which manages toolbars! you can individually show or hide, stash and restore tool windows configurations at the click of a button! (Yes, kind of like the Git window built into Visual Studio 2022, 2026, except for tool windows). More below.

### The Show/Hide/Stash/Restore Tool Windows floating tool window

Yeah, I know its a mouthful - I'll probably think of a better name for this tool window later!!! 



  - This is super easy to use. First use Visual Studio's **View** menu to view the toolbars you'd like to stash for later use.
  - Click the **Show/Hide/Stash/Restore specific tool windows** menu item on the **Tools** menu. The **Show/Hide specific tool windows** tool window should appear. It looks like this:  

<img width="695" height="361" alt="image" src="https://github.com/user-attachments/assets/d8c33999-dd84-44e1-b62a-34e346ad8bf3" />
  
  - Hit the Refresh button on the **Show/Hide specific tool windows** tool window to sync the visible windows with the toolbar's list  (shortcut is **F5**)
  - Select the tool windows you wish to stash by checking the checkboxes, if not already checked for you.
  - Click the "Stash Selected" button. The visible tool window configuration (excluding the tool management window itself!) will be added to the top of the stash. You have stashed the tool window configuration!
  - The tool window configurations are PERMANENTLY stored in the stash until you **Drop** (delete) them or **Pop** (restore then discard) top item. The stash persists between Visual Studio sessions.
    - Pop has two buttons - **Pop (Merge)** will **merge** the tool windows in the stash with those currently visible. **Pop (Abs)** will **replace** the tool windows with those in the stash.
    - **Only use Pop when you want to discard the top stash.**
  - Tool windows  can be recalled at ANY TIME without popping, simply by double left clicking (**Merge mode**) or double right clicking (**Absolute replace mode**) on a stash list item.
  - You can also right click on a stashed item to get a context menu.
   
  
## Assigning Keyboard Shortcuts (recommended)

I recommend you assign keyboard shortcuts to these commands. To do this, click **Tools | Options** then open **Environment | Keyboard**, like so:

![image](https://github.com/user-attachments/assets/be8c2e2b-e840-452e-9220-e81ea408da81)

Search for **ScottTunstall** in the **Show Commands Containing** textbox. You should see the 3 commands listed on screen. Assign shortcuts to each command as desired.
