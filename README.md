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
* The fourth command, **Stash/Restore tool windows** is for power users. This pops up a **floating tool window** which allows you to stash (save) and restore tool windows configurations at the click of a button! More below.

### The Stash/Restore Tool Windows floating tool window

**"_What's a stash?_"** you ask? **It's a collection of tool windows that you save for later retrieval**. (Or rather, _references_ to tool windows)


#### View the tool windows you wish to stash (do this first!) 
  - First use Visual Studio's **View** menu to view the tool windows you'd like to stash. This brings them into the **Stash/Restore Tool Window**'s "view".
    - If you can't see all of the tool windows in the list because the listbox is too small, there's a **horizontal splitter bar** just above the **Stashes** collapsible panel.   
  - Now click the **Stash/Restore tool windows** menu item on the **Tools** menu. The **Stash/Restore Tool Windows** modeless tool window should appear. It looks like this (NB: the list contents will differ from what you see):  

<img width="535" height="395" alt="image" src="https://github.com/user-attachments/assets/df26b936-c891-4eac-88e5-4347a391b091" />

#### Creating a stash (do this second!)

  - Hit the **Refresh** button to list the available tool windows (shortcut is **F5**).
    - Note the use of the word _available_ rather than "visible" - once Visual Studio opens a tool window, even if its invisible, its usually still lurking.
  - Select the tool windows you wish to stash by checking the respective checkboxes, if not already checked for you.
  - Click the **Stash Checked** button. A stash for the checked tool windows will be _pushed_ to the top of the stash list - yes, just like a LIFO stack. **You can now use the Pop/Apply stash functions to show the stashed tool windows in one go, rather than the chore of doing it with the View Menu!**
  - There's no limit to the amount of stashes you can add, although I wouldn't recommend adding hundreds.
    - If you _do_ add a lot of stashes and screen real estate is getting tight, remember there's a horizontal splitter bar above the **Stashes** button which you can use to size the stashes listbox. And of course, you can size the **Stash/Restore** tool window as well. 
  - Stashes persist between Visual Studio sessions, until you **Drop** (delete) or **Pop** (restore then delete) them.

#### Applying a stash  
  - Tool windows referred to by a stash can be applied (shown) at ANY TIME, simply by double left clicking (**Merge mode**) on a stash list item. You can also right click on a stashed item to get a context menu (see below).

#### Popping a stash  
  - If you want to pop the top item from the list, you can use either the **Pop (Merge)** or the **Pop (Abs)** button.    
    - **Pop (Merge)** will **merge** the tool windows in the popped stash with those currently visible.
    - **Pop (Abs)** , where Abs is short for **Absolute**, will **replace** currently visible tool windows with those referred to by the popped stash.
    - The Pop operation **permanently deletes** the stash item. If you don't want to delete the stash, use the **context menu** (see below)

#### Dropping a stash

  - **Drop All** permanently deletes all stashes. You will be asked to confirm if you are sure, as deletion cannot be undone.

#### The Stash item context menu

When you **right click** on a stash item, the item will be selected and a context menu will appear:

<img width="730" height="403" alt="image" src="https://github.com/user-attachments/assets/d460fc83-594a-4587-8ac0-65fe44602c07" />
   
  - **Apply (Merge)** will **merge** the tool windows in the selected stash with those currently visible.
  - **Apply (Absolute)** will **replace** the currently visible tool windows with those in the selected stash.
  - **Hide All ref'd by Stash** will hide all of the toolbar windows referenced by the stash.
  - **Drop** will delete the stash permanently.


  
## Assigning Keyboard Shortcuts (recommended)

I recommend you assign keyboard shortcuts to these commands. To do this, click **Tools | Options** then open **Environment | Keyboard**, like so:

<img width="743" height="445" alt="image" src="https://github.com/user-attachments/assets/fddec3ae-009e-4a47-8fc9-bd2b450abec6" />

Search for **ScottTunstall** in the **Show Commands Containing** textbox. You should see the four commands listed on screen. Assign shortcuts to each command as desired.
