



OpenSoldat Polyworks 1.7.0.0


updated 2022-07-28




Instructions

Unzipping creates a PolyWorks folder.
If PolyWorks does not detect your OpenSoldat directory on startup, go to Edit -> Preferences and set it there.
If you don't have OpenSoldat installed you can create a folder with the Maps, Scenery-gfx, and Textures directories.

Requirements: Windows 98/Me/2000/XP/Vista/7/8/10/11; DirectX 8.1




FAQ

01. What is PolyWorks?
PolyWorks is a map editor for OpenSoldat.

02. Where can I get the latest version of PolyWorks?
The latest test version can be found on discord in the #polyworks channel: https://discord.gg/VWfkqMEkMh

03. Why is there no help file?
There is now :D

05. Can you make polys go behind the player?
No, this isn't possible in OpenSoldat.

06. The Palette/Scenery/Display window doesn't show, and it's checked off in the Window Menu.
Try Window -> Workspace -> Reset Window Locations.

07. Can you make the controls like in MapMaker?
No. You'll get used to it.

08. How do I put multiple textures in my map?
Put two or more textures into a bitmap and use the texture tool to manipulate the texture coordinates.

09. I get this error when I try to run PolyWorks: System Error &H8007007E (-2147024770). The specified module could not be found. (Or any 'missing file' type errors.)
Make sure these files exist in your Windows\System32 folder: MBMouse.ocx, COMDLG32.OCX, mscomctl.ocx, msvbvm60.dll, dx8vb.dll, scrrun.dll. The first three are included in the PolyWorks zip, the others can be found with google. If putting them in the Windows\System32 folder doesn't work try resistering the missing files: start -> run, type in regsvr32 name_of_file.dll_or_ocx.

10. I get a runtime error when I start PolyWorks.
Try changing the Dir entry in the Preferences section in polyworks.ini to your OpenSoldat directory.

11. Are there poly bugs in maps made with PolyWorks?
In PolyWorks the "bouncy poly" bug has been eliminated (where polys would randomly turn bouncy along the edge after compiling). Poly bugs associated with vertices still occur, but they are easy to prevent with correct poly placement. Read the Map Maker Manual for more info.

12. I get a Direct3D initialization error when I start PolyWorks.
Make sure your color setting is either 16-bit or 32-bit (control panel -> display -> settings tab).

13. There is no scenery in my scenery window!
Right click with the scenery tool to bring up the main scenery list.



Changes in v1.7.0.0
- added support for resetting zoom by clicking zoom label
- added optional skinning support for resize button
- added custom MinZoom and MaxZoom (via ini settings)
- added reset zoom ini-setting
- added Min/Max/Reset Zoom inputs in preferences window
- added MinZoom/MaxZoom/ResetZoom default ini settings to polyworks.ini
- added default shortcut key info to help page
- added support loading of jpeg and jpg textures in map dialog
- added snapped tool windows stay snapped to window border on resize
- added * to filename label if there will be a prompt on close/load
- added default values for when polyworks.ini is missing to errors on startup
- modified use I-Beam cursor for zoom percentage input
- modified align preference labels consistently
- modified use large titlebar for up to 5k resolutions
- modified only hide tool windows when moving editor window instead on mouse down
- modified increase height of scenery list
- modified *clr ini settings into *color
- modified improve readability of default tools window buttons
- modified change skin folder to <polyworks_dir>\skins\*
- modified move default "gfx" skin into skin folder under new name "default"
- modified sync help tools naming with current naming in program
- modified make help html pass html validator checks
- modified rename clrpicker.cur into colorpicker.cur
- modified center manual and limit width to improve readability
- modified center copyright information in help file
- modified adjust preferences positioning and labels for consistency
- modified don't lowercase folder selection result in preferences dialog
- modified use consistent error messages
- modified remove layout jump due to missing image size info in help file
- modified save poly type color config on close in polyworks.ini
- fixed skin colors not applied to all controls
- fixed preference window other section "use 4 verts for scenery" text cut off
- fixed Maximized window overlaps taskbar
- fixed uninstaller doesn't delete new resize.bmp file
- fixed preferences dialog other section frame border skinning wasn't applied
- fixed issue error terminating issue when colorpicker is open
- fixed zoom info label cut off on large zoom numbers
- fixed broken grid rendering for narrow window sizes
- fixed menu setting is off initially even if grid is on
- fixed initial grid rendering for narrow windows is broken before resize
- fixed programm doesn't open in maximized mode
- fixed error message when selecting using vertex selection tool while zoomed in
- fixed maximized main window gets resized after applying preferences
- fixed point render issues in Wine (and WineD3D for Windows)
- fixed render error when loading workspace in windowed mode
- fixed background color assignment for options not applied for second background color
- fixed make tool windows all snap after workspace reset in non maximized state
- fixed reset workspace doesn't handle different taskbar size/position and 16:9 screens correctly
- fixed getting rgb color for AARRGGBB returns wrong value
- fixed wrong maximize button icon if main window is maximized on start up
- fixed loading/saving zoom* config setting should use "##.#" format only to work with initial settings
- fixed clear sketch menu doesn't rerender canvas after clear

Changes in v1.6.0.1
- added resizing support for main window (bottom right corner)
- added ctrl+shift+o hotkey for opening compiled maps
- modified newly created ini files are easier to read due to space between sections
- modified minimal main window width and height to can be as low as 300x200
- modified used "Load Compiled Map" title for open-compiled-map dialog
- fixed broken "°" label for Hue in Color Picker
- fixed broken Soldat directory check in Preferences Dialog
- fixed missing gif.tga file prevents texture loading
- fixed saving Preferences shuffles around scenery textures resulting in wrong texture positions
- fixed mouse move error message while refreshing preferences (after saving Preferences)
- fixed issue with overlapping hidden progressbar control on small windows sizes
- fixed side-window content flickers white on show after dragging main window
- fixed wrong initial input handling due to non default window size
- fixed broken undos when undo folder exists but is empty

Changes in v1.6.0.0
- added remember sticky state of tools windows after reopening workspace/pw
- added snapped subwindows stay by the main window if it moved
- added support for collapsing Tools Window
- modified execute form snapping for palette wndow like for other windows
- modified use Arial in Display form
- modified switch from colour to color in filenames and files
- modified switch to Arial as default font
- modified use lowercase pms file extension
- modified replaced pwlib.dll with vb6 gif loading
- modified refresh scenery reloads the complete scenery list not just the selected scenery
- modified sort scenery in 'Scenery' panel
- fixed make text in extended mode in preference form readable with arial font
- fixed preferences dialog disappears and cannot be opened after save error popup
- fixed error on startup due to uninitialized variables
- fixed skipping detailed warning for invalid soldat directory setting/registry key
- fixed position of texture window close button not all the way to the right
- fixed missing textures reset polygons on maximize
- fixed polyworks can't find Soldat directory
- fixed text selection on focus doesn't work
- fixed texture loading errors on startup
- fixed error on missing undo folder

Changes in v1.5.0.13
- added flag collides, background, and background transition polygon types

Changes in v1.5.0.12
- added flagger collides and non-flagger collides polygon types
- modified extension for saving from uppercase to lowercase

Changes in v1.5.0.11
- fixed bouncy polygons not being compiled correctly
- fixed saving of waypoints to prefabs not working

Changes in v1.5.0.10
- fixed maps being buggy in soldat 1.5.0

Changes in v1.5.0.9
- fixed styling of bounciness label
- fixed buggy polygon points
- fixed fixed texture

Changes in v1.5.0.8
- added polygon bounciness option for the bouncy polygon type
- added counter for the amount of *different* sceneries in the map
- fixed movement of spawns/objects while zoomed
- fixed texture window now closes properly
- fixed user defined x/y modifying the behavior of fixed texture
- fixed saving of user defined x/y

Changes in v1.5.0.7
- added mouse position label in status bar
- added right click menu on selection tools
- modified keyboard shortcuts deselect (escape) and duplicate (ctrl+d)
- modified selecting a single collider now gets its radius
- fixed the cutting of file names in recent files menu
- fixed a small copy/paste bug

Changes in v1.5.0.6
- added copy and paste (ctrl+c and ctrl+v, duplicate is now ctrl+shift+v)
- added invert selection (ctrl+i)
- added new polygon types for Soldat 1.5.1
- modified manual type in transform works on everything
- modified collider radius can be changed after it's placed
- fixed a crash when loading corrupt scenery
- fixed selection bug with hidden scenery layers
- fixed keyboard shortcut for save as (ctrl+shift+s)

Changes in v1.5.0.5
- added new arrangement of the main menu
- added basic texture transformation functions
- added menu item to reset the view
- added possibility to show and hide individual scenery layers
- fixed jpg sceneries and textures are not selectable (nvidia card compatibility errors ingame)
- fixed problems with drag and drop
- fixed polygon blend enable/disable

Changes in v1.5.0.4
- added associate pms files with polyworks on installation
- added icons for pms and pfb files (Created by VirtualTT)
- added more settings in preferences
- added selection for all corners of scenery
- modified how gif files are loaded
- fixed undo selection
- fixed saving of light and sketch display options
- fixed selection rectangle bug
- fixed command line argument bugs
- fixed transform tool

Changes in v1.5.0.3
- added new icon is now visible in taskbar
- added clear sketch function
- modified wider scenery menu
- fixed window state errors
- fixed background color in preferences
- fixed black trails in vista

Changes in v1.5.0.2
- added new icon (created by VirtualTT)
- fixed light bugs
- fixed avarage vertex colors not saving correctly
- fixed opacity for the 4 first polys
- fixed wireframe opacity bug

Changes in v1.5.0.1
- modified moved help button a bit away from minimize button
- fixed error when switching to/from windowed mode

Changes in v1.5.0.0
- added snap selected vertices function
- fixed gif files now working correctly

Changes in v1.4.0.17
- fixed connection severing
- fixed vertex alpha saving in prefabs

Changes in v1.4.0.16
- more descriptive error codes

Changes in v1.4.0.10
- fixed objects texture size bug
- disabled undo after clear unused scenery

Changes in v1.4.0.9
- extended path lengths from 80 to 260
- suppressed "not acquired" DirectInput error
- fixed loading maps with duplicates of the first scenery in the scenery list

Changes in v1.4.0.8
- rewrote part of compile code
- fixed scenery with wrong case not loading
- added lights range

Changes in v1.4.0.6
- opacity applied to polys on creation
- added vertex alpha control in properties window

Changes in v1.4.0.5
- fixed red/blue components of poly colors switched on export
- fixed scenery filter bug

Changes in v1.4.0.4
- changed ini loading code back to how it was before 1.4
- fixed tool hotkey/circle drawing bug

Changes in v1.4.0.2
- depth map works with opacity
- fixed depth map tool causing crash if there are any invisible polygons
- fixed right click menu not showing after quad checked

Changes in v1.4.0.1
- max radius changed to 128
- reload scenery bug fixed
- scenery in master list sorted alphabetically
- more error trapping

Changes in v1.4
- sketch tool
- lights tool
- depthmap tool
- big scenery list on right click with scenery tool
- average vertex colors function
- some changes to the ini file
- select skin from preferences window
- snap radius in ini file
- fixed various bugs
- introduced various new bugs
- other stuff i am too lazy to list

Changes in v1.3
- customizable fonts
- texture window
- changes to texture panel in properties window
- gdi+ for png support
- fixed scenery timestamp
- fixed scenery file name (case sensitive)
- horizontal flip on waypoints changes left<->right
- colors picked from map are selected in the palette
- save/load workspace
- other bug fixes

Changes in v1.2
- create poly with selected vertices function
- better skin support
- import function
- keyboard input with directinput
- loads png and jpg scenery and textures
- load compiled map from Soldat Maps folder
- select vertices by color function
- waypoint support
- fixed some maps causing errors when loading
- choose uncompiled maps dir and prefabs dir
- check boxes and option boxes have clickable labels
- hotkeys for waypoints
- hotkeys for display layers
- changed prefabs format and extension to .PFB
- fixed undo making invisible polys visible
- properties window shows number of polys/scenery/spawns/colliders/waypoints/connections
- properties window shows element name of scenery when single scenery is selected
- disabled recent files when empty
- creating something sets that layer to visible if it is not
- can only color polys/scenery when visible
- experimental textured quad function
- pressing F1 or the ? button opens the help file
- turned off vsync
- various bug fixes

Changes in v1.0
- fast compile, no bouncy polys
- [ and ] cycle through tools
- custom hotkeys
- gfx.bmp split up into two files
- automatic directory detection
- color picker tool
- show vertex color radius
- compile progress bar
- type in zoom level
- map options
- spawn points
- colliders
- scaling
- rotation
- scenery
- constrained transformations (hold shift)
- flip/rotate
- scenery/polys snap to grid while creating/scaling
- tab cycles through vertices/polys/scenery
- gostek object
- actual size function
- fit to screen function
- properties window
- recent files
- ini files
- hex code in color picker window
- arrow keys move texture coords with texture tool
- undo/redo
- run soldat with last compiled map
- move/texture tools work when no vertices are selected
- fixed lighten, darken, difference blend modes
- help file
- lots of bug fixes and improvements




Credits:

created by Anna Zajaczkowski in Visual Basic 6
updated version by Jacob Lindberg (Fryer) and Gregor A. Cieslak (Shoozza)
original PolyWorks concept and ideas by Michal Zajaczkowski
graphics by Michal Zajaczkowski and Anna Zajaczkowski, based on Soldat style
new icon by VirtualTT
thanks to everyone at #soldat.polyworks :)
thanks to Michal Marcinkowski for releasing the Map Maker source (and for making Soldat ;])



