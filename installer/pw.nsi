; Script generated by the HM NIS Edit Script Wizard.
; HM NIS Edit Wizard helper defines
!define PRODUCT_NAME "OpenSoldat PolyWorks"
!define PRODUCT_VERSION "1.7.0.0"
!define PRODUCT_PUBLISHER "Copyright Anna Zajaczkowski"
!define PRODUCT_WEB_SITE "http://forums.soldat.pl"
!define PRODUCT_DIR_REGKEY "Software\Microsoft\Windows\CurrentVersion\App Paths\OpenSoldat PolyWorks.exe"
!define PRODUCT_UNINST_KEY "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}"
!define PRODUCT_UNINST_ROOT_KEY "HKLM"


; MUI 1.67 compatible ------
!include "MUI.nsh"


; MUI Settings
!define MUI_ABORTWARNING
!define MUI_ICON "PW.ico"
!define MUI_UNICON "PW.ico"


; Welcome page
!insertmacro MUI_PAGE_WELCOME

; License page
!insertmacro MUI_PAGE_LICENSE "../LICENSE"

; Directory page
!insertmacro MUI_PAGE_DIRECTORY

; Instfiles page
!insertmacro MUI_PAGE_INSTFILES

; Finish page
!insertmacro MUI_PAGE_FINISH

; Uninstaller pages
!insertmacro MUI_UNPAGE_INSTFILES

; Language files
!insertmacro MUI_LANGUAGE "English"


; MUI end ------

!define SHCNE_ASSOCCHANGED 0x8000000
!define SHCNF_IDLIST 0

Name "${PRODUCT_NAME} ${PRODUCT_VERSION}"
OutFile "PWSetup.exe"
InstallDir "$PROGRAMFILES\OpenSoldat PolyWorks"
InstallDirRegKey HKLM "${PRODUCT_DIR_REGKEY}" ""
ShowInstDetails show
ShowUnInstDetails show

BrandingText " "


; On Install

Section "MainSection" SEC01
  SetOutPath "$INSTDIR"
  SetOverwrite ifnewer
  File "OpenSoldat PolyWorks.exe"
  File "COMDLG32.OCX"
  File "MBMouse.ocx"
  File "mscomctl.ocx"
  File "polyworks.ini"
  File "PolyWorks Help.html"
  File "ReadMe.txt"
  File "PMS.ico"
  File "PFB.ico"

  SetOutPath "$SYSDIR"
  File "dx8vb.dll"
  File "MBMouse.ocx"
  File "mscomctl.ocx"
  File "COMDLG32.OCX"
  RegDLL "$SYSDIR\dx8vb.dll"
  RegDLL "$SYSDIR\MBMouse.ocx"
  RegDLL "$SYSDIR\mscomctl.ocx"
  RegDLL "$SYSDIR\COMDLG32.OCX"

  SetOutPath "$INSTDIR\BMPtoCUR"
  File "BMPtoCUR\BMP to CUR.exe"
  File "BMPtoCUR\ReadMe.txt"

  SetOutPath "$INSTDIR\skins\default"
  File "skins\default\titlebar_palette.bmp"
  File "skins\default\tool_gfx.bmp"
  File "skins\default\titlebar_texture.bmp"
  File "skins\default\slider_arrow.bmp"
  File "skins\default\pattern.bmp"
  File "skins\default\tools.bmp"
  File "skins\default\path.png"
  File "skins\default\titlebar_preferences.bmp"
  File "skins\default\titlebar_colorpicker.bmp"
  File "skins\default\vertex8x8.bmp"
  File "skins\default\titlebar_tools.bmp"
  File "skins\default\titlebar_map.bmp"
  File "skins\default\titlebar_main.bmp"
  File "skins\default\color_picker.bmp"
  File "skins\default\titlebar_properties.bmp"
  File "skins\default\button_gfx.bmp"
  File "skins\default\notfound.bmp"
  File "skins\default\titlebar_display.bmp"
  File "skins\default\objects.bmp"
  File "skins\default\lines.bmp"
  File "skins\default\colors.ini"
  File "skins\default\sketch.bmp"
  File "skins\default\rcenter.bmp"
  File "skins\default\titlebar_waypoints.bmp"
  File "skins\default\titlebar_scenery.bmp"
  File "skins\default\resize.bmp"

  SetOutPath "$INSTDIR\skins\default\cursors"
  File "skins\default\cursors\pselect.cur"
  File "skins\default\cursors\waypoint.cur"
  File "skins\default\cursors\create.cur"
  File "skins\default\cursors\pselsub.cur"
  File "skins\default\cursors\smudge.cur"
  File "skins\default\cursors\pcolor.cur"
  File "skins\default\cursors\rotate.cur"
  File "skins\default\cursors\hand.cur"
  File "skins\default\cursors\scale.cur"
  File "skins\default\cursors\pseladd.cur"
  File "skins\default\cursors\vselect.cur"
  File "skins\default\cursors\colorpicker.cur"
  File "skins\default\cursors\move.cur"
  File "skins\default\cursors\scenery.cur"
  File "skins\default\cursors\color_picker.cur"
  File "skins\default\cursors\quad.cur"
  File "skins\default\cursors\depthmap.cur"
  File "skins\default\cursors\texture.cur"
  File "skins\default\cursors\vselsub.cur"
  File "skins\default\cursors\objects.cur"
  File "skins\default\cursors\pixpicker.cur"
  File "skins\default\cursors\vcolor.cur"
  File "skins\default\cursors\light.cur"
  File "skins\default\cursors\sketch.cur"
  File "skins\default\cursors\connect.cur"
  File "skins\default\cursors\eraser.cur"
  File "skins\default\cursors\vseladd.cur"
  File "skins\default\cursors\litpicker.cur"

  SetOutPath "$INSTDIR\Help"
  File "Help\tool_sketch.gif"
  File "Help\tool_pcolor.gif"
  File "Help\tool_vselect.gif"
  File "Help\tool_scenery.gif"
  File "Help\tool_texture.gif"
  File "Help\tool_objects.gif"
  File "Help\tool_create.gif"
  File "Help\tool_vcolor.gif"
  File "Help\tool_depthmap.gif"
  File "Help\tool_colorpicker.gif"
  File "Help\tool_lights.gif"
  File "Help\tool_move.gif"
  File "Help\tool_pselect.gif"
  File "Help\tool_waypoint.gif"

  SetOutPath "$INSTDIR\lists"
  File "lists\defaults.txt"
  CreateDirectory "$INSTDIR\Maps"

  SetOutPath "$INSTDIR\palettes"
  File "palettes\current.txt"
  File "palettes\MZpalette.txt"
  File "palettes\palette.txt"
  CreateDirectory "$INSTDIR\Prefabs"
  CreateDirectory "$INSTDIR\undo"
  CreateDirectory "$INSTDIR\Temp"

  SetOutPath "$INSTDIR\Workspace"
  File "Workspace\current.ini"

  CreateDirectory "$SMPROGRAMS\OpenSoldat PolyWorks"
  CreateShortCut "$SMPROGRAMS\OpenSoldat PolyWorks\OpenSoldat PolyWorks.lnk" "$INSTDIR\OpenSoldat PolyWorks.exe"
  CreateShortCut "$SMPROGRAMS\OpenSoldat PolyWorks\Uninstall OpenSoldat PolyWorks.lnk" "$INSTDIR\uninst.exe"

  WriteRegStr HKCR ".pms" "" "OpenSoldat PolyWorks Map"
  WriteRegStr HKCR "OpenSoldat PolyWorks Map" "" "OpenSoldat PolyWorks Map"
  WriteRegStr HKCR "OpenSoldat PolyWorks Map\shell" "" "open"
  WriteRegStr HKCR "OpenSoldat PolyWorks Map\DefaultIcon" "" "$INSTDIR\PMS.ico"
  WriteRegStr HKCR "OpenSoldat PolyWorks Map\shell\open\command" "" '"$INSTDIR\OpenSoldat PolyWorks.exe" "%1"'
  WriteRegStr HKCR ".pfb" "" "OpenSoldat PolyWorks Prefab"
  WriteRegStr HKCR "OpenSoldat PolyWorks Prefab" "" "OpenSoldat PolyWorks Prefab"
  WriteRegStr HKCR "OpenSoldat PolyWorks Prefab\DefaultIcon" "" "$INSTDIR\PFB.ico"
  System::Call 'Shell32::SHChangeNotify(i ${SHCNE_ASSOCCHANGED}, i ${SHCNF_IDLIST}, i 0, i 0)'
SectionEnd

Section -Post
  WriteUninstaller "$INSTDIR\uninst.exe"
  WriteRegStr HKLM "${PRODUCT_DIR_REGKEY}" "" "$INSTDIR\OpenSoldat PolyWorks.exe"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayName" "$(^Name)"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "UninstallString" "$INSTDIR\uninst.exe"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayIcon" "$INSTDIR\OpenSoldat PolyWorks.exe"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "DisplayVersion" "${PRODUCT_VERSION}"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "URLInfoAbout" "${PRODUCT_WEB_SITE}"
  WriteRegStr ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}" "Publisher" "${PRODUCT_PUBLISHER}"
SectionEnd


; On Uninstall

Function un.onUninstSuccess
  HideWindow
  MessageBox MB_ICONINFORMATION|MB_OK "$(^Name) was successfully removed from your computer."
FunctionEnd

Function un.onInit
  MessageBox MB_ICONQUESTION|MB_YESNO|MB_DEFBUTTON2 "Are you sure you want to completely remove $(^Name) and all of its components?" IDYES +2

  Abort
FunctionEnd

Section Uninstall
  Delete /REBOOTOK "$INSTDIR\uninst.exe"
  Delete /REBOOTOK "$INSTDIR\OpenSoldat PolyWorks.exe"
  Delete /REBOOTOK "$INSTDIR\bgothl.ttf"
  Delete /REBOOTOK "$INSTDIR\COMDLG32.OCX"
  Delete /REBOOTOK "$INSTDIR\lucon.ttf"
  Delete /REBOOTOK "$INSTDIR\MBMouse.ocx"
  Delete /REBOOTOK "$INSTDIR\mscomctl.ocx"
  Delete /REBOOTOK "$INSTDIR\polyworks.ini"
  Delete /REBOOTOK "$INSTDIR\PolyWorks Help.html"
  Delete /REBOOTOK "$INSTDIR\ReadMe.txt"
  Delete /REBOOTOK "$INSTDIR\dx8vb.dll"
  Delete /REBOOTOK "$INSTDIR\PMS.ico"
  Delete /REBOOTOK "$INSTDIR\PFB.ico"

  Delete /REBOOTOK "$INSTDIR\BMPtoCUR\BMP to CUR.exe"
  Delete /REBOOTOK "$INSTDIR\BMPtoCUR\ReadMe.txt"
  RMDir /REBOOTOK "$INSTDIR\BMPtoCUR"

  Delete /REBOOTOK "$INSTDIR\skins\default\Thumbs.db"
  Delete /REBOOTOK "$INSTDIR\skins\default\titlebar_palette.bmp"
  Delete /REBOOTOK "$INSTDIR\skins\default\tool_gfx.bmp"
  Delete /REBOOTOK "$INSTDIR\skins\default\titlebar_texture.bmp"
  Delete /REBOOTOK "$INSTDIR\skins\default\slider_arrow.bmp"
  Delete /REBOOTOK "$INSTDIR\skins\default\pattern.bmp"
  Delete /REBOOTOK "$INSTDIR\skins\default\tools.bmp"
  Delete /REBOOTOK "$INSTDIR\skins\default\path.png"
  Delete /REBOOTOK "$INSTDIR\skins\default\titlebar_preferences.bmp"
  Delete /REBOOTOK "$INSTDIR\skins\default\titlebar_colorpicker.bmp"
  Delete /REBOOTOK "$INSTDIR\skins\default\vertex8x8.bmp"
  Delete /REBOOTOK "$INSTDIR\skins\default\titlebar_tools.bmp"
  Delete /REBOOTOK "$INSTDIR\skins\default\titlebar_map.bmp"
  Delete /REBOOTOK "$INSTDIR\skins\default\titlebar_main.bmp"
  Delete /REBOOTOK "$INSTDIR\skins\default\color_picker.bmp"
  Delete /REBOOTOK "$INSTDIR\skins\default\titlebar_properties.bmp"
  Delete /REBOOTOK "$INSTDIR\skins\default\button_gfx.bmp"
  Delete /REBOOTOK "$INSTDIR\skins\default\notfound.bmp"
  Delete /REBOOTOK "$INSTDIR\skins\default\titlebar_display.bmp"
  Delete /REBOOTOK "$INSTDIR\skins\default\objects.bmp"
  Delete /REBOOTOK "$INSTDIR\skins\default\lines.bmp"
  Delete /REBOOTOK "$INSTDIR\skins\default\colors.ini"
  Delete /REBOOTOK "$INSTDIR\skins\default\sketch.bmp"
  Delete /REBOOTOK "$INSTDIR\skins\default\rcenter.bmp"
  Delete /REBOOTOK "$INSTDIR\skins\default\titlebar_waypoints.bmp"
  Delete /REBOOTOK "$INSTDIR\skins\default\titlebar_scenery.bmp"
  Delete /REBOOTOK "$INSTDIR\skins\default\resize.bmp"

  Delete /REBOOTOK "$INSTDIR\skins\default\cursors\pselect.cur"
  Delete /REBOOTOK "$INSTDIR\skins\default\cursors\waypoint.cur"
  Delete /REBOOTOK "$INSTDIR\skins\default\cursors\create.cur"
  Delete /REBOOTOK "$INSTDIR\skins\default\cursors\pselsub.cur"
  Delete /REBOOTOK "$INSTDIR\skins\default\cursors\smudge.cur"
  Delete /REBOOTOK "$INSTDIR\skins\default\cursors\pcolor.cur"
  Delete /REBOOTOK "$INSTDIR\skins\default\cursors\rotate.cur"
  Delete /REBOOTOK "$INSTDIR\skins\default\cursors\hand.cur"
  Delete /REBOOTOK "$INSTDIR\skins\default\cursors\scale.cur"
  Delete /REBOOTOK "$INSTDIR\skins\default\cursors\pseladd.cur"
  Delete /REBOOTOK "$INSTDIR\skins\default\cursors\vselect.cur"
  Delete /REBOOTOK "$INSTDIR\skins\default\cursors\colorpicker.cur"
  Delete /REBOOTOK "$INSTDIR\skins\default\cursors\move.cur"
  Delete /REBOOTOK "$INSTDIR\skins\default\cursors\scenery.cur"
  Delete /REBOOTOK "$INSTDIR\skins\default\cursors\color_picker.cur"
  Delete /REBOOTOK "$INSTDIR\skins\default\cursors\quad.cur"
  Delete /REBOOTOK "$INSTDIR\skins\default\cursors\depthmap.cur"
  Delete /REBOOTOK "$INSTDIR\skins\default\cursors\texture.cur"
  Delete /REBOOTOK "$INSTDIR\skins\default\cursors\vselsub.cur"
  Delete /REBOOTOK "$INSTDIR\skins\default\cursors\objects.cur"
  Delete /REBOOTOK "$INSTDIR\skins\default\cursors\pixpicker.cur"
  Delete /REBOOTOK "$INSTDIR\skins\default\cursors\vcolor.cur"
  Delete /REBOOTOK "$INSTDIR\skins\default\cursors\light.cur"
  Delete /REBOOTOK "$INSTDIR\skins\default\cursors\sketch.cur"
  Delete /REBOOTOK "$INSTDIR\skins\default\cursors\connect.cur"
  Delete /REBOOTOK "$INSTDIR\skins\default\cursors\eraser.cur"
  Delete /REBOOTOK "$INSTDIR\skins\default\cursors\vseladd.cur"
  Delete /REBOOTOK "$INSTDIR\skins\default\cursors\litpicker.cur"
  RMDir /REBOOTOK "$INSTDIR\skins\default\cursors"
  RMDir /REBOOTOK "$INSTDIR\skins\default"
  RMDir /REBOOTOK "$INSTDIR\skins"

  Delete /REBOOTOK "$INSTDIR\Help\Thumbs.db"
  Delete /REBOOTOK "$INSTDIR\Help\tool_sketch.gif"
  Delete /REBOOTOK "$INSTDIR\Help\tool_pcolor.gif"
  Delete /REBOOTOK "$INSTDIR\Help\tool_vselect.gif"
  Delete /REBOOTOK "$INSTDIR\Help\tool_scenery.gif"
  Delete /REBOOTOK "$INSTDIR\Help\tool_texture.gif"
  Delete /REBOOTOK "$INSTDIR\Help\tool_objects.gif"
  Delete /REBOOTOK "$INSTDIR\Help\tool_create.gif"
  Delete /REBOOTOK "$INSTDIR\Help\tool_vcolor.gif"
  Delete /REBOOTOK "$INSTDIR\Help\tool_depthmap.gif"
  Delete /REBOOTOK "$INSTDIR\Help\tool_colorpicker.gif"
  Delete /REBOOTOK "$INSTDIR\Help\tool_lights.gif"
  Delete /REBOOTOK "$INSTDIR\Help\tool_move.gif"
  Delete /REBOOTOK "$INSTDIR\Help\tool_pselect.gif"
  Delete /REBOOTOK "$INSTDIR\Help\tool_waypoint.gif"
  RMDir /REBOOTOK "$INSTDIR\Help"

  Delete /REBOOTOK "$INSTDIR\lists\defaults.txt"
  RMDir /REBOOTOK "$INSTDIR\lists"
  RMDir /REBOOTOK "$INSTDIR\Maps"

  Delete /REBOOTOK "$INSTDIR\palettes\current.txt"
  Delete /REBOOTOK "$INSTDIR\palettes\MZpalette.txt"
  Delete /REBOOTOK "$INSTDIR\palettes\palette.txt"
  RMDir /REBOOTOK "$INSTDIR\palettes"

  RMDir /REBOOTOK "$INSTDIR\Prefabs"

  Delete /REBOOTOK "$INSTDIR\undo\*.pwn"
  RMDir /REBOOTOK "$INSTDIR\undo"

  Delete /REBOOTOK "$INSTDIR\Temp\Thumbs.db"
  Delete /REBOOTOK "$INSTDIR\Temp\gif.tga"
  Delete /REBOOTOK "$INSTDIR\Temp\copy.PFB"
  RMDir /REBOOTOK "$INSTDIR\Temp"

  Delete /REBOOTOK "$INSTDIR\Workspace\current.ini"
  RMDir /REBOOTOK "$INSTDIR\Workspace"
  RMDir /REBOOTOK "$INSTDIR"

  Delete /REBOOTOK "$SMPROGRAMS\OpenSoldat Polyworks\OpenSoldat PolyWorks.lnk"
  Delete /REBOOTOK "$SMPROGRAMS\OpenSoldat Polyworks\Uninstall OpenSoldat PolyWorks.lnk"
  RMDir /REBOOTOK "$SMPROGRAMS\OpenSoldat Polyworks"

  DeleteRegKey HKCR ".pms"
  DeleteRegKey HKCR "OpenSoldat PolyWorks Map"
  DeleteRegKey HKCR ".pfb"
  DeleteRegKey HKCR "OpenSoldat PolyWorks Prefab"
  System::Call 'Shell32::SHChangeNotify(i ${SHCNE_ASSOCCHANGED}, i ${SHCNF_IDLIST}, i 0, i 0)'

  DeleteRegKey ${PRODUCT_UNINST_ROOT_KEY} "${PRODUCT_UNINST_KEY}"
  DeleteRegKey HKLM "${PRODUCT_DIR_REGKEY}"
  SetAutoClose true
SectionEnd