# Show Tool Window Extension for Visual Studio 2019, 2022 and 2026.

## About

Are you using Visual Studio 2019, 2022 or 2026?

Does your Solution Explorer end up wandering off screen or get stuck under a pile of windows and you have a torrid time trying to find it?
Or do you end up with a load of Tool Windows open and you don't want to reset window layout but would like to close them all?

Fret no more. Reign in naughty tool windows with this extension.


## How to use

After installing the extension, you should see four new commands (circled in image) on the Visual Studio **Tools** menu :

<img width="442" height="634" alt="image" src="https://github.com/user-attachments/assets/b9f0f04b-ecdc-434e-8d9f-776b39fc17bd" />



* The first command, **Show Solution Explorer** will make the solution explorer FULLY VISIBLE even if it was partially or mostly offscreen before.
* The second command, **Close all Tool Windows except Solution Explorer** will close all tool windows except the solution explorer - very useful when you're in Tool Window Hell. 
* The third command, **Close all Tool Windows** is the nuclear option and does as you expect. Every tool window goes, but your code windows will be left untouched. 
* The fourth command, **Stash/Restore tool windows** is for power users. This pops up a floating toolbar window which allows you to stash (save) and restore tool windows configurations at the click of a button! More below.

### The Stash/Restore Tool Windows floating tool window

  - **"_What's a stash?_"** you ask? Its a collection of tool windows that you saved for later retrieval.

#### View the tool windows you wish to stash (do this first!) 
  - First use Visual Studio's **View** menu to view the tool windows you'd like to stash. This brings them into the **Stash/Restore Tool Window**'s "view".   
  - Now click the **Stash/Restore tool windows** menu item on the **Tools** menu. The **Stash/Restore Tool Windows** modeless tool window should appear. It looks like this (NB: the list contents may vary from what you see):  

<img width="535" height="395" alt="image" src="https://github.com/user-attachments/assets/df26b936-c891-4eac-88e5-4347a391b091" />

#### Creating a stash (do this second!)

  - Hit the **Refresh** button to list the available tool windows (shortcut is **F5**).
    - Note the use of the word _available_ rather than "visible" - once Visual Studio opens a tool window, even if its invisible, its usually still lurking.
  - Select the tool windows you wish to stash by checking the respective checkboxes, if not already checked for you.
  - Click the **Stash Checked** button. The checked tool windows (excluding the tool management window itself!) will be stashed at the top of the stash list - yes, just like a LIFO stack. **You can now use the stash functions to show the stashed tool windows in one go, rather than the chore of doing it with the View Menu!**
  - There's no limit to the amount of stashes you can add, although I wouldn't recommend adding hundreds.
  - Tool windows are PERMANENTLY "stored" in the stash, and persist between Visual Studio sessions, until you **Drop** (delete) them or **Pop** (restore then delete) top item.

#### Applying a stash  
  - Tool windows in the stash can be applied (shown) at ANY TIME, simply by double left clicking (**Merge mode**) on a stash list item. You can also right click on a stashed item to get a context menu (see below).

#### Popping a stash  
  - If you want to pop the top item from the list, you can use either the **Pop (Merge)** or the **Pop (Abs)** button.    
    - **Pop (Merge)** will **merge** the tool windows in the popped stash with those currently visible.
    - **Pop (Abs)** , where Abs is short for **Absolute**, will **replace** currently visible tool windows with those in the popped stash.
    - The Pop operation permanently deletes the stash item. If you don't want to delete the stash, use the context menu (see below)

#### Dropping a stash

  - **Drop All** permanently deletes all stashes. You will be asked to confirm if you are sure, as deletion cannot be undone.

#### The Stash item context menu

When you right click on a stash in the list, the context menu below will appear:

<img width="730" height="403" alt="image" src="https://github.com/user-attachments/assets/d460fc83-594a-4587-8ac0-65fe44602c07" />
   
  - **Apply (Merge)** will **merge** the tool windows in the selected stash with those currently visible.
  - **Apply (Abs)** will **replace** the currently visible tool windows with those in the selected stash.


  
## Assigning Keyboard Shortcuts (recommended)

I recommend you assign keyboard shortcuts to these commands. To do this, click **Tools | Options** then open **Environment | Keyboard**, like so:

<img width="872" height="653" alt="image" src="https://github.com/user-attachments/assets/b1b55c27-b7e4-4e6a-a60d-3999a72a0df0" />

Search for **ScottTunstall** in the **Show Commands Containing** textbox. You should see the four commands listed on screen. Assign shortcuts to each command as desired.
