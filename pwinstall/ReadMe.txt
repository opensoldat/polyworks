



Soldat Polyworks 1.6.0.1


updated 2022-02-04




Instructions

Unzipping creates a PolyWorks folder.
If PolyWorks does not detect your Soldat directory on startup, go to Edit -> Preferences and set it there.
If you don't have Soldat installed you can create a folder with the Maps, Scenery-gfx, and Textures directories.

Requirements: Windows 98/Me/2000/XP; DirectX 8.1




FAQ

01. What is PolyWorks?
PolyWorks is a map editor for Soldat.

02. Where can I get the latest version of PolyWorks?
The latest test version can be found on irc in the #soldat.polyworks channel on quakenet.

03. Why is there no help file?
There is now :D

05. Can you make polys go behind the player?
No, this isn't possible in Soldat.

06. The Palette/Scenery/Display window doesn't show, and it's checked off in the Window Menu.
Try Window -> Workspace -> Reset Window Locations.

07. Can you make the controls like in MapMaker?
No. You'll get used to it.

08. How do I put multiple textures in my map?
Put two or more textures into a bitmap and use the texture tool to manipulate the texture coordinates.

09. I get this error when I try to run PolyWorks: System Error &H8007007E (-2147024770). The specified module could not be found. (Or any 'missing file' type errors.)
Make sure these files exist in your Windows\System32 folder: MBMouse.ocx, COMDLG32.OCX, mscomctl.ocx, msvbvm60.dll, dx8vb.dll, scrrun.dll. The first three are included in the PolyWorks zip, the others can be found with google. If putting them in the Windows\System32 folder doesn't work try resistering the missing files: start -> run, type in regsvr32 name_of_file.dll_or_ocx.

10. I get a runtime error when I start PolyWorks.
Try changing the Dir entry in the Preferences section in polyworks.ini to your Soldat directory.

11. Are there poly bugs in maps made with PolyWorks?
In PolyWorks the "bouncy poly" bug has been eliminated (where polys would randomly turn bouncy along the edge after compiling). Poly bugs associated with vertices still occur, but they are easy to prevent with correct poly placement. Read the Map Maker Manual for more info.

12. I get a Direct3D initialization error when I start PolyWorks.
Make sure your color setting is either 16-bit or 32-bit (control panel -> display -> settings tab).

13. There is no scenery in my scenery window!
Right click with the scenery tool to bring up the main scenery list.



Changes in v1.6.0.1
- Added resizing support for main window (bottom right corner)
- Added ctrl+shift+o hotkey for opening compiled maps
- Modified newly created ini files are easier to read due to space between sections
- Modified minimal main window width and height to can be as low as 300x200
- Modified used "Load Compiled Map" title for open-compiled-map dialog
- Fixed broken "°" label for Hue in Color Picker
- Fixed broken Soldat directory check in Preferences Dialog
- Fixed missing gif.tga file prevents texture loading
- Fixed saving Preferences shuffles around scenery textures resulting in wrong texture positions
- Fixed mouse move error message while refreshing preferences (after saving Preferences)
- Fixed issue with overlapping hidden progressbar control on small windows sizes
- Fixed side-window content flickers white on show after dragging main window

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
- changed extension for saving from uppercase to lowercase

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
- fixed a small copy/paste bug
- added right click menu on selection tools
- changed keyboard shortcuts deselect (escape) and duplicate (ctrl+d)
- fixed the cutting of file names in recent files menu
- selecting a single collider now gets its radius

Changes in v1.5.0.6
- copy and paste (ctrl+c and ctrl+v, duplicate is now ctrl+shift+v)
- invert selection (ctrl+i)
- new polygon types for Soldat 1.5.1
- manual type in transform works on everything
- collider radius can be changed after it's placed
- fixed a crash when loading corrupt scenery
- fixed selection bug with hidden scenery layers
- fixed keyboard shortcut for save as (ctrl+shift+s)

Changes in v1.5.0.5
- jpg sceneries and textures are not selectable (nvidia card compatibility errors ingame)
- fixed problems with drag and drop
- fixed polygon blend enable/disable
- new arrangement of the main menu
- added basic texture transformation functions
- added menu item to reset the view
- possibility to show and hide individual scenery layers

Changes in v1.5.0.4
- change how gif files are loaded
- fixed undo selection
- fixed saving of light and sketch display options
- fixed selection rectangle bug
- fixed command line argument bugs
- associate pms files with polyworks on installation
- icons for pms and pfb files (Created by VirtualTT)
- more settings in preferences
- selection for all corners of scenery
- fixed transform tool

Changes in v1.5.0.3
- fixed window state errors
- new icon is now visible in taskbar
- fixed background color in preferences
- wider scenery menu
- added clear sketch function
- fixed black trails in vista

Changes in v1.5.0.2
- fixed light bugs
- fixed avarage vertex colors not saving correctly
- fixed opacity for the 4 first polys
- fixed wireframe opacity bug
- new icon (created by VirtualTT)

Changes in v1.5.0.1
- fixed error when switching to/from windowed mode
- moved help button a bit away from minimize button

changes in v1.5.0.0
- added snap selected vertices function
- fixed gif files now working correctly

changes in v1.4.0.17
- fixed connection severing
- fixed vertex alpha saving in prefabs

Changes in v1.4.0.16
- more descriptive error codes

changes in v1.4.0.10
- fixed objects texture size bug
- disabled undo after clear unused scenery

changes in v1.4.0.9
- extended path lengths from 80 to 260
- suppressed "not acquired" DirectInput error
- fixed loading maps with duplicates of the first scenery in the scenery list

changes in v1.4.0.8
- rewrote part of compile code
- fixed scenery with wrong case not loading
- added lights range

changes in v1.4.0.6
- opacity applied to polys on creation
- added vertex alpha control in properties window

changes in v1.4.0.5
- fixed red/blue components of poly colors switched on export
- fixed scenery filter bug

changes in v1.4.0.4
- changed ini loading code back to how it was before 1.4
- fixed tool hotkey/circle drawing bug

changes in v1.4.0.2
- depth map works with opacity
- fixed depth map tool causing crash if there are any invisible polygons
- fixed right click menu not showing after quad checked

changes in v1.4.0.1
- max radius changed to 128
- reload scenery bug fixed
- scenery in master list sorted alphabetically
- more error trapping

changes in v1.4
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

changes in v1.3
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

changes in v1.2
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

changes in v1.0
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

programmed by Anna Zajaczkowski in Visual Basic 6
updated version by Jacob Lindberg (Fryer)
original PolyWorks concept and ideas by Michal Zajaczkowski
graphics by Michal Zajaczkowski and Anna Zajaczkowski, based on Soldat style
new icon by VirtualTT
thanks to everyone at #soldat.polyworks :)
thanks to Michal Marcinkowski for releasing the Map Maker source (and for making Soldat ;])




IRC: #soldat.polyworks on quakenet
email: soldat.polyworks@gmail.com



