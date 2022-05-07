Attribute VB_Name = "modConfig"
Option Explicit

' loading and saving config files, workspaces, skins goes here


' Fix vb6 ide casing changes
#If False Then
    Public FileName, color, token, A, R, G, B, commonDialog, value, Val, X, Y, Z, Left, hWnd, Mid, Right, BackColor
    'Public FileName, color, token, A, R, G, B, commonDialog, value, Val, X, Y, Z, Left, hWnd, Mid, Right, BackColor
#End If


' vars - public

Public gPolyColor As TColor

Public gPolyTypeColors(0 To 25) As Long


' vars - private

Private Const DEFAULT_MAX_ZOOM As Single = 512
Private Const DEFAULT_MIN_ZOOM As Single = 0.03125
Private Const DEFAULT_RESET_ZOOM As Single = 1


' functions - public

Public Sub LoadSettings()

    On Error GoTo ErrorHandler

    appPath = App.path

    Debug.Assert SetIdePath  ' workaround for debugging with ide

    Dim i As Integer
    Dim numRecent As Integer
    Dim strTemp As String
    Dim sgnTemp As Single
    Dim errVal As String

    errVal = "1"

    frmSoldatMapEditor.soldatDir = LoadString("Preferences", "Dir", , 1024)
    frmSoldatMapEditor.uncompDir = LoadString("Preferences", "Uncompiled", , 1024)
    frmSoldatMapEditor.prefabDir = LoadString("Preferences", "Prefabs", , 1024)

    frmSoldatMapEditor.gridSpacing = LoadInt("Preferences", "GridSpacing", , 32)
    frmSoldatMapEditor.gridDivisions = LoadInt("Preferences", "GridDiv", , 4)
    frmSoldatMapEditor.gridColor1 = HexToLong(LoadString("Preferences", "GridColor1", , 1024, "000000"))
    frmSoldatMapEditor.gridColor2 = HexToLong(LoadString("Preferences", "GridColor2", , 1024, "000000"))
    frmSoldatMapEditor.gridOp1 = LoadInt("Preferences", "GridAlpha1", , 255)
    frmSoldatMapEditor.gridOp2 = LoadInt("Preferences", "GridAlpha2", , 51)
    frmSoldatMapEditor.polyBlendSrc = LoadInt("Preferences", "PolySrc", , 5)
    frmSoldatMapEditor.polyBlendDest = LoadInt("Preferences", "PolyDest", , 6)
    frmSoldatMapEditor.wireBlendSrc = LoadInt("Preferences", "WireSrc", , 2)
    frmSoldatMapEditor.wireBlendDest = LoadInt("Preferences", "WireDest", , 2)
    frmSoldatMapEditor.pointColor = HexToLong(LoadString("Preferences", "PointColor", , 1024, "CE4D4A"))
    frmSoldatMapEditor.selectionColor = HexToLong(LoadString("Preferences", "SelectionColor", , 1024, "CE4D4A"))
    frmSoldatMapEditor.backgroundColor = HexToLong(LoadString("Preferences", "BackColor", , 1024, "555555"))
    frmSoldatMapEditor.maxUndo = LoadInt("Preferences", "MaxUndo", , 16)
    frmSoldatMapEditor.sceneryVerts = LoadBoolean("Preferences", "SceneryVerts", , False)
    frmSoldatMapEditor.topmost = LoadBoolean("Preferences", "Topmost", , True)

    strTemp = LoadString("Preferences", "MinZoom", , 1024, "51200")
    strTemp = Replace(strTemp, ",", ".", 1, -1, vbTextCompare)
    If IsNumeric(strTemp) Then
        frmSoldatMapEditor.gMinZoom = Val(strTemp) / 100
    Else
        frmSoldatMapEditor.gMinZoom = DEFAULT_MIN_ZOOM
    End If

    strTemp = LoadString("Preferences", "MaxZoom", , 1024, "3.125")
    strTemp = Replace(strTemp, ",", ".", 1, -1, vbTextCompare)
    If IsNumeric(strTemp) Then
        frmSoldatMapEditor.gMaxZoom = Val(strTemp) / 100
    Else
        frmSoldatMapEditor.gMaxZoom = DEFAULT_MAX_ZOOM
    End If

    If frmSoldatMapEditor.gMinZoom = frmSoldatMapEditor.gMaxZoom Then
        frmSoldatMapEditor.gMinZoom = DEFAULT_MIN_ZOOM
        frmSoldatMapEditor.gMaxZoom = DEFAULT_MAX_ZOOM
    ElseIf frmSoldatMapEditor.gMinZoom > frmSoldatMapEditor.gMaxZoom Then
        sgnTemp = frmSoldatMapEditor.gMaxZoom
        frmSoldatMapEditor.gMaxZoom = frmSoldatMapEditor.gMinZoom
        frmSoldatMapEditor.gMinZoom = sgnTemp
    End If

    strTemp = LoadString("Preferences", "ResetZoom", , 1024, "100")
    strTemp = Replace(strTemp, ",", ".", 1, -1, vbTextCompare)
    If IsNumeric(strTemp) Then
        frmSoldatMapEditor.gResetZoom = Val(strTemp) / 100
    Else
        frmSoldatMapEditor.gResetZoom = DEFAULT_RESET_ZOOM
    End If

    If frmSoldatMapEditor.gResetZoom > frmSoldatMapEditor.gMaxZoom Then
        frmSoldatMapEditor.gResetZoom = frmSoldatMapEditor.gMaxZoom
    ElseIf frmSoldatMapEditor.gResetZoom < frmSoldatMapEditor.gMinZoom Then
        frmSoldatMapEditor.gResetZoom = frmSoldatMapEditor.gMinZoom
    End If

    errVal = "2"

    frmSoldatMapEditor.showBG = LoadString("Display", "Background")
    frmSoldatMapEditor.showPolys = LoadString("Display", "Polys")
    frmSoldatMapEditor.showTexture = LoadString("Display", "Texture")
    frmSoldatMapEditor.showWireframe = LoadString("Display", "Wireframe")
    frmSoldatMapEditor.showPoints = LoadString("Display", "Points")
    frmSoldatMapEditor.showScenery = LoadString("Display", "Scenery")
    frmSoldatMapEditor.showObjects = LoadString("Display", "Objects")
    frmSoldatMapEditor.showWaypoints = LoadString("Display", "Waypoints")
    frmSoldatMapEditor.showGrid = LoadString("Display", "Grid")
    frmSoldatMapEditor.showLights = LoadString("Display", "Lights")
    frmSoldatMapEditor.showSketch = LoadString("Display", "Sketch")

    frmSoldatMapEditor.mnuGrid.Checked = frmSoldatMapEditor.showGrid

    errVal = "3"

    frmSoldatMapEditor.currentTool = LoadInt("ToolSettings", "CurrentTool")
    frmSoldatMapEditor.ohSnap = LoadString("ToolSettings", "SnapVertices")
    frmSoldatMapEditor.snapToGrid = LoadString("ToolSettings", "SnapToGrid")
    frmSoldatMapEditor.fixedTexture = LoadString("ToolSettings", "FixedTexture")
    frmSoldatMapEditor.opacity = LoadInt("ToolSettings", "Opacity") / 100
    frmSoldatMapEditor.colorRadius = LoadInt("ToolSettings", "ColorRadius")
    gPolyColor = GetRGB(HexToLong(LoadString("ToolSettings", "CurrentColor")))
    frmSoldatMapEditor.colorMode = LoadInt("ToolSettings", "ColorMode")
    frmSoldatMapEditor.blendMode = LoadInt("ToolSettings", "BlendMode")
    frmSoldatMapEditor.snapRadius = LoadInt("ToolSettings", "SnapRadius")
    frmScenery.rotateScenery = LoadString("ToolSettings", "RotateScenery")
    frmScenery.scaleScenery = LoadString("ToolSettings", "ScaleScenery")
    frmSoldatMapEditor.xTexture = LoadInt("ToolSettings", "TextureWidth")
    frmSoldatMapEditor.yTexture = LoadInt("ToolSettings", "TextureHeight")
    frmSoldatMapEditor.gTextureFile = LoadString("ToolSettings", "Texture", , 1024)
    frmSoldatMapEditor.mnuCustomX.Checked = LoadString("ToolSettings", "CustomX")
    frmSoldatMapEditor.mnuCustomY.Checked = LoadString("ToolSettings", "CustomY")

    errVal = "4"

    frmTools.SetHotKey 0, LoadByte("HotKeys", "Move", , 30)
    frmTools.SetHotKey 1, LoadByte("HotKeys", "Create", , 16)
    frmTools.SetHotKey 2, LoadByte("HotKeys", "VertexSelection", , 31)
    frmTools.SetHotKey 3, LoadByte("HotKeys", "PolySelection", , 17)
    frmTools.SetHotKey 4, LoadByte("HotKeys", "VertexColor", , 32)
    frmTools.SetHotKey 5, LoadByte("HotKeys", "PolyColor", , 18)
    frmTools.SetHotKey 6, LoadByte("HotKeys", "Texture", , 33)
    frmTools.SetHotKey 7, LoadByte("HotKeys", "Scenery", , 19)
    frmTools.SetHotKey 8, LoadByte("HotKeys", "Waypoints", , 34)
    frmTools.SetHotKey 9, LoadByte("HotKeys", "Objects", , 20)
    frmTools.SetHotKey 10, LoadByte("HotKeys", "ColorPicker", , 35)
    frmTools.SetHotKey 11, LoadByte("HotKeys", "Sketch", , 21)
    frmTools.SetHotKey 12, LoadByte("HotKeys", "Lights", , 36)
    frmTools.SetHotKey 13, LoadByte("HotKeys", "DepthMap", , 22)

    errVal = "5"

    frmWaypoints.SetWaypointKey 0, LoadByte("WaypointKeys", "Left", , 36)
    frmWaypoints.SetWaypointKey 1, LoadByte("WaypointKeys", "Right", , 37)
    frmWaypoints.SetWaypointKey 2, LoadByte("WaypointKeys", "Up", , 23)
    frmWaypoints.SetWaypointKey 3, LoadByte("WaypointKeys", "Down", , 50)
    frmWaypoints.SetWaypointKey 4, LoadByte("WaypointKeys", "Fly", , 49)

    errVal = "6"

    frmDisplay.SetLayerKey 0, LoadByte("LayerKeys", "Background", , 79)
    frmDisplay.SetLayerKey 1, LoadByte("LayerKeys", "Polys", , 80)
    frmDisplay.SetLayerKey 2, LoadByte("LayerKeys", "Texture", , 81)
    frmDisplay.SetLayerKey 3, LoadByte("LayerKeys", "Wireframe", , 75)
    frmDisplay.SetLayerKey 4, LoadByte("LayerKeys", "Points", , 76)
    frmDisplay.SetLayerKey 5, LoadByte("LayerKeys", "Scenery", , 77)
    frmDisplay.SetLayerKey 6, LoadByte("LayerKeys", "Objects", , 71)
    frmDisplay.SetLayerKey 7, LoadByte("LayerKeys", "Waypoints", , 72)

    errVal = "7"

    frmSoldatMapEditor.mnuRecent(0).Caption = LoadString("RecentFiles", "01", , 1024)
    frmSoldatMapEditor.mnuRecent(1).Caption = LoadString("RecentFiles", "02", , 1024)
    frmSoldatMapEditor.mnuRecent(2).Caption = LoadString("RecentFiles", "03", , 1024)
    frmSoldatMapEditor.mnuRecent(3).Caption = LoadString("RecentFiles", "04", , 1024)
    frmSoldatMapEditor.mnuRecent(4).Caption = LoadString("RecentFiles", "05", , 1024)
    frmSoldatMapEditor.mnuRecent(5).Caption = LoadString("RecentFiles", "06", , 1024)
    frmSoldatMapEditor.mnuRecent(6).Caption = LoadString("RecentFiles", "07", , 1024)
    frmSoldatMapEditor.mnuRecent(7).Caption = LoadString("RecentFiles", "08", , 1024)
    frmSoldatMapEditor.mnuRecent(8).Caption = LoadString("RecentFiles", "09", , 1024)
    frmSoldatMapEditor.mnuRecent(9).Caption = LoadString("RecentFiles", "10", , 1024)

    errVal = "8"

    gPolyTypeColors(1) = HexToLong(LoadString("PolyTypeColors", "OnlyBullets", , , "7ACC29"))
    gPolyTypeColors(2) = HexToLong(LoadString("PolyTypeColors", "OnlyPlayer", , , "CCCC29"))
    gPolyTypeColors(3) = HexToLong(LoadString("PolyTypeColors", "DoesntCollide", , , "29CC29"))
    gPolyTypeColors(4) = HexToLong(LoadString("PolyTypeColors", "Ice", , , "29CCCC"))
    gPolyTypeColors(5) = HexToLong(LoadString("PolyTypeColors", "Deadly", , , "CC297A"))
    gPolyTypeColors(6) = HexToLong(LoadString("PolyTypeColors", "BloodyDeadly", , , "CC29CC"))
    gPolyTypeColors(7) = HexToLong(LoadString("PolyTypeColors", "Hurts", , , "CC2929"))
    gPolyTypeColors(8) = HexToLong(LoadString("PolyTypeColors", "Regenerates", , , "2929CC"))
    gPolyTypeColors(9) = HexToLong(LoadString("PolyTypeColors", "Lava", , , "CC7A29"))
    gPolyTypeColors(10) = HexToLong(LoadString("PolyTypeColors", "TeamBullets", , , "7A7A29"))
    gPolyTypeColors(11) = HexToLong(LoadString("PolyTypeColors", "TeamPlayers", , , "7A2929"))
    gPolyTypeColors(12) = gPolyTypeColors(10)
    gPolyTypeColors(13) = gPolyTypeColors(11)
    gPolyTypeColors(14) = gPolyTypeColors(10)
    gPolyTypeColors(15) = gPolyTypeColors(11)
    gPolyTypeColors(16) = gPolyTypeColors(10)
    gPolyTypeColors(17) = gPolyTypeColors(11)
    gPolyTypeColors(18) = HexToLong(LoadString("PolyTypeColors", "Bouncy", , , "297ACC"))
    gPolyTypeColors(19) = HexToLong(LoadString("PolyTypeColors", "Explosive", , , "CCCCCC"))
    gPolyTypeColors(20) = HexToLong(LoadString("PolyTypeColors", "HurtFlaggers", , , "CCCC7A"))
    gPolyTypeColors(21) = HexToLong(LoadString("PolyTypeColors", "OnlyFlagger", , , "7A7ACC"))
    gPolyTypeColors(22) = HexToLong(LoadString("PolyTypeColors", "NonFlagger", , , "7A29CC"))
    gPolyTypeColors(23) = HexToLong(LoadString("PolyTypeColors", "FlagCollides", , , "29297A"))
    gPolyTypeColors(24) = HexToLong(LoadString("PolyTypeColors", "Back", , , "292929"))
    gPolyTypeColors(25) = HexToLong(LoadString("PolyTypeColors", "BackTransition", , , "7A7A7A"))

    errVal = "9"

    gfxDir = LoadString("gfx", "Dir", , 1024, "default")

    If gfxDir = "" Then gfxDir = "default"

    errVal = "10"

    For i = frmSoldatMapEditor.mnuRecent.LBound + 1 To frmSoldatMapEditor.mnuRecent.UBound
        If frmSoldatMapEditor.mnuRecent(i).Caption = "" Then
            numRecent = numRecent + 1
            frmSoldatMapEditor.mnuRecent(i).Visible = False
        Else
            frmSoldatMapEditor.mnuRecent(i).Visible = True
        End If
    Next
    If numRecent = frmSoldatMapEditor.mnuRecent.Count - 1 And frmSoldatMapEditor.mnuRecent(frmSoldatMapEditor.mnuRecent.LBound).Caption = "" Then
        frmSoldatMapEditor.mnuRecentFiles.Enabled = False
    End If

    Exit Sub

ErrorHandler:

    MsgBox "Error loading ini file" & vbNewLine & Error & vbNewLine & errVal

End Sub

Public Sub SaveSettings()

    Dim X As Integer
    Dim Y As Integer
    Dim i As Integer
    Dim KeyCode As Byte

    Dim iniString As String
    Dim currentColor As Long
    Dim sNull As String
    sNull = Chr(0)
    Dim isNewFile As Boolean
    
    isNewFile = Not FileExists(appPath & "\polyworks.ini")

    ' preferences
    iniString = _
        "Dir=" & frmSoldatMapEditor.soldatDir & sNull & _
        "Uncompiled=" & frmSoldatMapEditor.uncompDir & sNull & _
        "Prefabs=" & frmSoldatMapEditor.prefabDir & sNull & _
        "GridSpacing=" & frmSoldatMapEditor.gridSpacing & sNull & _
        "GridDiv=" & frmSoldatMapEditor.gridDivisions & sNull & _
        "GridColor1=" & RGBtoHex(frmSoldatMapEditor.gridColor1) & sNull & _
        "GridColor2=" & RGBtoHex(frmSoldatMapEditor.gridColor2) & sNull & _
        "GridAlpha1=" & frmSoldatMapEditor.gridOp1 & sNull & _
        "GridAlpha2=" & frmSoldatMapEditor.gridOp2 & sNull & _
        "PolySrc=" & frmSoldatMapEditor.polyBlendSrc & sNull & _
        "PolyDest=" & frmSoldatMapEditor.polyBlendDest & sNull & _
        "WireSrc=" & frmSoldatMapEditor.wireBlendSrc & sNull & _
        "WireDest=" & frmSoldatMapEditor.wireBlendDest & sNull & _
        "PointColor=" & RGBtoHex(frmSoldatMapEditor.pointColor) & sNull & _
        "SelectionColor=" & RGBtoHex(frmSoldatMapEditor.selectionColor) & sNull & _
        "BackColor=" & RGBtoHex(frmSoldatMapEditor.backgroundColor) & sNull & _
        "MaxUndo=" & frmSoldatMapEditor.maxUndo & sNull & _
        "SceneryVerts=" & frmSoldatMapEditor.sceneryVerts & sNull & _
        "Topmost=" & frmSoldatMapEditor.topmost & sNull & _
        "MinZoom=" & Trim(Str(frmSoldatMapEditor.gMaxZoom * 100)) & sNull & _
        "MaxZoom=" & Trim(Str(frmSoldatMapEditor.gMinZoom * 100)) & sNull & _
        "ResetZoom=" & Trim(Str(frmSoldatMapEditor.gResetZoom * 100)) & _
        IIf(isNewFile, vbNewLine, "") & sNull & sNull
    SaveSection "Preferences", iniString

    ' display
    iniString = _
        "Background=" & frmSoldatMapEditor.showBG & sNull & _
        "Polys=" & frmSoldatMapEditor.showPolys & sNull & _
        "Texture=" & frmSoldatMapEditor.showTexture & sNull & _
        "Wireframe=" & frmSoldatMapEditor.showWireframe & sNull & _
        "Points=" & frmSoldatMapEditor.showPoints & sNull & _
        "Scenery=" & frmSoldatMapEditor.showScenery & sNull & _
        "Objects=" & frmSoldatMapEditor.showObjects & sNull & _
        "Waypoints=" & frmSoldatMapEditor.showWaypoints & sNull & _
        "Grid=" & frmSoldatMapEditor.showGrid & sNull & _
        "Lights=" & frmSoldatMapEditor.showLights & sNull & _
        "Sketch=" & frmSoldatMapEditor.showSketch & _
        IIf(isNewFile, vbNewLine, "") & sNull & sNull
    SaveSection "Display", iniString

    ' tool settings
    currentColor = RGB(gPolyColor.blue, gPolyColor.green, gPolyColor.red)
    iniString = _
        "CurrentTool=" & frmSoldatMapEditor.currentTool & sNull & _
        "SnapVertices=" & frmSoldatMapEditor.ohSnap & sNull & _
        "SnapToGrid=" & frmSoldatMapEditor.snapToGrid & sNull & _
        "FixedTexture=" & frmSoldatMapEditor.fixedTexture & sNull & _
        "Opacity=" & (frmSoldatMapEditor.opacity * 100) & sNull & _
        "ColorRadius=" & frmSoldatMapEditor.colorRadius & sNull & _
        "CurrentColor=" & RGBtoHex(currentColor) & sNull & _
        "ColorMode=" & frmSoldatMapEditor.colorMode & sNull & _
        "BlendMode=" & frmSoldatMapEditor.blendMode & sNull & _
        "SnapRadius=" & frmSoldatMapEditor.snapRadius & sNull & _
        "RotateScenery=" & frmScenery.rotateScenery & sNull & _
        "ScaleScenery=" & frmScenery.scaleScenery & sNull & _
        "TextureWidth=" & frmSoldatMapEditor.xTexture & sNull & _
        "TextureHeight=" & frmSoldatMapEditor.yTexture & sNull & _
        "Texture=" & frmSoldatMapEditor.gTextureFile & sNull & _
        "CustomX=" & frmSoldatMapEditor.mnuCustomX.Checked & sNull & _
        "CustomY=" & frmSoldatMapEditor.mnuCustomY.Checked & _
        IIf(isNewFile, vbNewLine, "") & sNull & sNull
    SaveSection "ToolSettings", iniString

    ' hotkeys
    iniString = _
        "Move=" & frmTools.GetHotKey(0) & sNull & _
        "Create=" & frmTools.GetHotKey(1) & sNull & _
        "VertexSelection=" & frmTools.GetHotKey(2) & sNull & _
        "PolySelection=" & frmTools.GetHotKey(3) & sNull & _
        "VertexColor=" & frmTools.GetHotKey(4) & sNull & _
        "PolyColor=" & frmTools.GetHotKey(5) & sNull & _
        "Texture=" & frmTools.GetHotKey(6) & sNull & _
        "Scenery=" & frmTools.GetHotKey(7) & sNull & _
        "Waypoints=" & frmTools.GetHotKey(8) & sNull & _
        "Objects=" & frmTools.GetHotKey(9) & sNull & _
        "ColorPicker=" & frmTools.GetHotKey(10) & sNull & _
        "Sketch=" & frmTools.GetHotKey(11) & sNull & _
        "Lights=" & frmTools.GetHotKey(12) & sNull & _
        "Depthmap=" & frmTools.GetHotKey(13) & _
        IIf(isNewFile, vbNewLine, "") & sNull & sNull
    SaveSection "HotKeys", iniString

    ' waypoint keys
    iniString = _
        "Left=" & frmWaypoints.GetWaypointKey(0) & sNull & _
        "Right=" & frmWaypoints.GetWaypointKey(1) & sNull & _
        "Up=" & frmWaypoints.GetWaypointKey(2) & sNull & _
        "Down=" & frmWaypoints.GetWaypointKey(3) & sNull & _
        "Fly=" & frmWaypoints.GetWaypointKey(4) & _
        IIf(isNewFile, vbNewLine, "") & sNull & sNull
    SaveSection "WaypointKeys", iniString

    ' layer keys
    iniString = _
        "Background=" & frmDisplay.GetLayerKey(0) & sNull & _
        "Polys=" & frmDisplay.GetLayerKey(1) & sNull & _
        "Texture=" & frmDisplay.GetLayerKey(2) & sNull & _
        "Wireframe=" & frmDisplay.GetLayerKey(3) & sNull & _
        "Points=" & frmDisplay.GetLayerKey(4) & sNull & _
        "Scenery=" & frmDisplay.GetLayerKey(5) & sNull & _
        "Objects=" & frmDisplay.GetLayerKey(6) & sNull & _
        "Waypoints=" & frmDisplay.GetLayerKey(7) & _
        IIf(isNewFile, vbNewLine, "") & sNull & sNull
    SaveSection "LayerKeys", iniString

    ' palette
    frmPalette.SavePalette appPath & "\palettes\current.txt"

    ' recent files
    iniString = _
        "01=" & frmSoldatMapEditor.mnuRecent(0).Caption & sNull & _
        "02=" & frmSoldatMapEditor.mnuRecent(1).Caption & sNull & _
        "03=" & frmSoldatMapEditor.mnuRecent(2).Caption & sNull & _
        "04=" & frmSoldatMapEditor.mnuRecent(3).Caption & sNull & _
        "05=" & frmSoldatMapEditor.mnuRecent(4).Caption & sNull & _
        "06=" & frmSoldatMapEditor.mnuRecent(5).Caption & sNull & _
        "07=" & frmSoldatMapEditor.mnuRecent(6).Caption & sNull & _
        "08=" & frmSoldatMapEditor.mnuRecent(7).Caption & sNull & _
        "09=" & frmSoldatMapEditor.mnuRecent(8).Caption & sNull & _
        "10=" & frmSoldatMapEditor.mnuRecent(9).Caption & _
        IIf(isNewFile, vbNewLine, "") & sNull & sNull
    SaveSection "RecentFiles", iniString

    ' gfx dir
    iniString = "Dir=" & gfxDir & sNull & sNull
    SaveSection "gfx", iniString

    ' workspace
    isNewFile = Not FileExists(appPath & "\workspace\current.ini")

    iniString = _
        "WindowState=" & frmSoldatMapEditor.Tag & sNull & _
        "Width=" & frmSoldatMapEditor.formWidth & sNull & _
        "Height=" & frmSoldatMapEditor.formHeight & sNull & _
        "Left=" & frmSoldatMapEditor.formLeft & sNull & _
        "Top=" & frmSoldatMapEditor.formTop & _
        IIf(isNewFile, vbNewLine, "") & sNull & sNull
    SaveSection "Main", iniString, appPath & "\workspace\current.ini"

    SaveWindow "Tools", frmTools, False, isNewFile
    SaveWindow "Display", frmDisplay, frmDisplay.collapsed, isNewFile
    SaveWindow "Properties", frmInfo, frmInfo.collapsed, isNewFile
    SaveWindow "Palette", frmPalette, frmPalette.collapsed, isNewFile
    SaveWindow "Scenery", frmScenery, frmScenery.collapsed, isNewFile
    SaveWindow "Waypoints", frmWaypoints, frmWaypoints.collapsed, isNewFile
    SaveWindow "Texture", frmTexture, frmTexture.collapsed, isNewFile

End Sub

Public Sub LoadWorkspace(Optional theFileName As String = "current.ini", Optional bSkipScenery As Boolean = False)

    On Error GoTo ErrorHandler

    frmSoldatMapEditor.Tag = LoadInt("Main", "WindowState", appPath & "\workspace\" & theFileName)
    frmSoldatMapEditor.formWidth = LoadInt("Main", "Width", appPath & "\workspace\" & theFileName)
    frmSoldatMapEditor.formHeight = LoadInt("Main", "Height", appPath & "\workspace\" & theFileName)
    frmSoldatMapEditor.formLeft = LoadInt("Main", "Left", appPath & "\workspace\" & theFileName)
    frmSoldatMapEditor.formTop = LoadInt("Main", "Top", appPath & "\workspace\" & theFileName)

    frmSoldatMapEditor.picResize.Top = frmSoldatMapEditor.formHeight - frmSoldatMapEditor.picResize.Height
    frmSoldatMapEditor.picResize.Left = frmSoldatMapEditor.formWidth - frmSoldatMapEditor.picResize.Width

    If frmSoldatMapEditor.Tag = vbNormal Then
        frmSoldatMapEditor.Width = frmSoldatMapEditor.formWidth * Screen.TwipsPerPixelX
        frmSoldatMapEditor.Height = frmSoldatMapEditor.formHeight * Screen.TwipsPerPixelY
        frmSoldatMapEditor.Left = frmSoldatMapEditor.formLeft * Screen.TwipsPerPixelX
        frmSoldatMapEditor.Top = frmSoldatMapEditor.formTop * Screen.TwipsPerPixelY
    Else
        frmSoldatMapEditor.Tag = vbNormal
        frmSoldatMapEditor.Width = frmSoldatMapEditor.formWidth * Screen.TwipsPerPixelX
        frmSoldatMapEditor.Height = frmSoldatMapEditor.formHeight * Screen.TwipsPerPixelY
        frmSoldatMapEditor.Left = frmSoldatMapEditor.formLeft * Screen.TwipsPerPixelX
        frmSoldatMapEditor.Top = frmSoldatMapEditor.formTop * Screen.TwipsPerPixelY
        frmSoldatMapEditor.MaximizeBorderLessForm
        frmSoldatMapEditor.picResize.Visible = False
    End If

    frmSoldatMapEditor.tvwScenery.Height = frmSoldatMapEditor.formHeight - 41 - 20

    frmSoldatMapEditor.mnuTools.Checked = LoadString("Tools", "Visible", appPath & "\workspace\" & theFileName)
    frmTools.xPos = LoadInt("Tools", "Left", appPath & "\workspace\" & theFileName)
    frmTools.yPos = LoadInt("Tools", "Top", appPath & "\workspace\" & theFileName)
    frmTools.collapsed = LoadString("Tools", "Collapsed", appPath & "\workspace\" & theFileName)
    frmTools.Tag = IIf(LoadString("Tools", "Snapped", appPath & "\workspace\" & theFileName) = "True", "snap", "")

    frmSoldatMapEditor.mnuDisplay.Checked = LoadString("Display", "Visible", appPath & "\workspace\" & theFileName)
    frmDisplay.xPos = LoadInt("Display", "Left", appPath & "\workspace\" & theFileName)
    frmDisplay.yPos = LoadInt("Display", "Top", appPath & "\workspace\" & theFileName)
    frmDisplay.collapsed = LoadString("Display", "Collapsed", appPath & "\workspace\" & theFileName)
    frmDisplay.Tag = IIf(LoadString("Display", "Snapped", appPath & "\workspace\" & theFileName) = "True", "snap", "")

    frmSoldatMapEditor.mnuInfo.Checked = LoadString("Properties", "Visible", appPath & "\workspace\" & theFileName)
    frmInfo.xPos = LoadInt("Properties", "Left", appPath & "\workspace\" & theFileName)
    frmInfo.yPos = LoadInt("Properties", "Top", appPath & "\workspace\" & theFileName)
    frmInfo.collapsed = LoadString("Properties", "Collapsed", appPath & "\workspace\" & theFileName)
    frmInfo.Tag = IIf(LoadString("Properties", "Snapped", appPath & "\workspace\" & theFileName) = "True", "snap", "")

    frmSoldatMapEditor.mnuPalette.Checked = LoadString("Palette", "Visible", appPath & "\workspace\" & theFileName)
    frmPalette.xPos = LoadInt("Palette", "Left", appPath & "\workspace\" & theFileName)
    frmPalette.yPos = LoadInt("Palette", "Top", appPath & "\workspace\" & theFileName)
    frmPalette.collapsed = LoadString("Palette", "Collapsed", appPath & "\workspace\" & theFileName)
    frmPalette.Tag = IIf(LoadString("Palette", "Snapped", appPath & "\workspace\" & theFileName) = "True", "snap", "")

    If Not bSkipScenery Then
        frmSoldatMapEditor.mnuScenery.Checked = LoadString("Scenery", "Visible", appPath & "\workspace\" & theFileName)
        frmScenery.xPos = LoadInt("Scenery", "Left", appPath & "\workspace\" & theFileName)
        frmScenery.yPos = LoadInt("Scenery", "Top", appPath & "\workspace\" & theFileName)
        frmScenery.collapsed = LoadString("Scenery", "Collapsed", appPath & "\workspace\" & theFileName)
        frmScenery.Tag = IIf(LoadString("Scenery", "Snapped", appPath & "\workspace\" & theFileName) = "True", "snap", "")
    End If

    frmSoldatMapEditor.mnuWaypoints.Checked = LoadString("Waypoints", "Visible", appPath & "\workspace\" & theFileName)
    frmWaypoints.xPos = LoadInt("Waypoints", "Left", appPath & "\workspace\" & theFileName)
    frmWaypoints.yPos = LoadInt("Waypoints", "Top", appPath & "\workspace\" & theFileName)
    frmWaypoints.collapsed = LoadString("Waypoints", "Collapsed", appPath & "\workspace\" & theFileName)
    frmWaypoints.Tag = IIf(LoadString("Waypoints", "Snapped", appPath & "\workspace\" & theFileName) = "True", "snap", "")

    frmSoldatMapEditor.mnuTexture.Checked = LoadString("Texture", "Visible", appPath & "\workspace\" & theFileName)
    frmTexture.xPos = LoadInt("Texture", "Left", appPath & "\workspace\" & theFileName)
    frmTexture.yPos = LoadInt("Texture", "Top", appPath & "\workspace\" & theFileName)
    frmTexture.collapsed = LoadString("Texture", "Collapsed", appPath & "\workspace\" & theFileName)
    frmTexture.Tag = IIf(LoadString("Texture", "Snapped", appPath & "\workspace\" & theFileName) = "True", "snap", "")

    Exit Sub

ErrorHandler:

    MsgBox "Error loading workspace" & vbNewLine & Error

End Sub

Public Sub SaveWindow(sectionName As String, window As Form, collapsed As Boolean, isNewFile As Boolean, Optional theFileName As String = "current.ini")

    Dim leftVal As Integer
    Dim topVal As Integer
    Dim iniString As String
    Dim sNull As String
    sNull = Chr(0)

    leftVal = window.Left / Screen.TwipsPerPixelX
    topVal = window.Top / Screen.TwipsPerPixelY

    iniString = _
        "Visible=" & window.Visible & sNull & _
        "Left=" & leftVal & sNull & _
        "Top=" & topVal & sNull & _
        "Collapsed=" & collapsed & sNull & _
        "Snapped=" & IIf(Len(window.Tag) > 0, "True", "False") & _
        IIf(isNewFile, vbNewLine, "") & sNull & sNull

    SaveSection sectionName, iniString, appPath & "\workspace\" & theFileName

End Sub


' functions - private

Private Function SetIdePath() As Boolean

    appPath = appPath & "\pwinstall"
    SetIdePath = True

End Function
