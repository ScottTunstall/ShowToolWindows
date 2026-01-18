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
* The fourth command, **Show/Hide/Stash/Restore tool windows** is for power users. This pops up a floating toolbar window which manages toolbars! you can individually show or hide, stash and restore tool windows configurations at the click of a button! More below.

### The Show/Hide/Stash/Restore Tool Windows floating tool window

(Yeah, I know the name is a mouthful - I'll probably think of a better name for this tool window later!!!) 

  - This is super easy to use. First use Visual Studio's **View** menu to view the toolbars you'd like to stash for later use. 
  - Now click the **Show/Hide/Stash/Restore specific tool windows** menu item on the **Tools** menu. The **Show/Hide specific tool windows** tool window should appear. It looks like this (NB: the list contents may vary from what you see):  

<img width="626" height="409" alt="image" src="https://github.com/user-attachments/assets/0c8eb986-5ec4-4d45-9b93-33a8f73ee91f" />

  
  - Hit the Refresh button on the **Show/Hide specific tool windows** tool window to list the visible tool windows. (shortcut is **F5**)
  - Select the tool windows you wish to stash by checking the checkboxes, if not already checked for you.
  - Click the "Stash Selected" button. The visible tool window configuration (excluding the tool management window itself!) will be added to the top of the stash. In future you don't need to use the View menu to show these toolbars - simply double click the stash and they open for you.
  - Tool windows  can be recalled at ANY TIME without popping, simply by double left clicking (**Merge mode**) on a stash list item. 
  - You can also right click on a stashed item to get a context menu.   
  - The stash configurations are PERMANENTLY stored in the stash, and persist between Visual Studio sessions, until you **Drop** (delete) them or **Pop** (restore then discard) top item.  
    - **Pop** has two buttons - **Pop (Merge)** will **merge** the tool windows in the stash with those currently visible. **Pop (Abs)** will **replace** the tool windows with those in the stash.
    - **Only use Pop when you want to discard a stash after use.**
    - **Drop All** deletes all stashes. You will be asked to confirm if you are sure, as deletion cannot be undone.
 
  
## Assigning Keyboard Shortcuts (recommended)

I recommend you assign keyboard shortcuts to these commands. To do this, click **Tools | Options** then open **Environment | Keyboard**, like so:

<img width="872" height="653" alt="image" src="https://github.com/user-attachments/assets/b1b55c27-b7e4-4e6a-a60d-3999a72a0df0" />

Search for **ScottTunstall** in the **Show Commands Containing** textbox. You should see the four commands listed on screen. Assign shortcuts to each command as desired.
