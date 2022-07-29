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

    frmOpensoldatMapEditor.opensoldatDir = LoadString("Preferences", "Dir", , 1024)
    frmOpensoldatMapEditor.uncompDir = LoadString("Preferences", "Uncompiled", , 1024)
    frmOpensoldatMapEditor.prefabDir = LoadString("Preferences", "Prefabs", , 1024)

    frmOpensoldatMapEditor.gridSpacing = LoadInt("Preferences", "GridSpacing", , 32)
    frmOpensoldatMapEditor.gridDivisions = LoadInt("Preferences", "GridDiv", , 4)
    frmOpensoldatMapEditor.gridColor1 = HexToLong(LoadString("Preferences", "GridColor1", , 1024, "000000"))
    frmOpensoldatMapEditor.gridColor2 = HexToLong(LoadString("Preferences", "GridColor2", , 1024, "000000"))
    frmOpensoldatMapEditor.gridOp1 = LoadInt("Preferences", "GridAlpha1", , 255)
    frmOpensoldatMapEditor.gridOp2 = LoadInt("Preferences", "GridAlpha2", , 51)
    frmOpensoldatMapEditor.polyBlendSrc = LoadInt("Preferences", "PolySrc", , 5)
    frmOpensoldatMapEditor.polyBlendDest = LoadInt("Preferences", "PolyDest", , 6)
    frmOpensoldatMapEditor.wireBlendSrc = LoadInt("Preferences", "WireSrc", , 2)
    frmOpensoldatMapEditor.wireBlendDest = LoadInt("Preferences", "WireDest", , 2)
    frmOpensoldatMapEditor.pointColor = HexToLong(LoadString("Preferences", "PointColor", , 1024, "CE4D4A"))
    frmOpensoldatMapEditor.selectionColor = HexToLong(LoadString("Preferences", "SelectionColor", , 1024, "CE4D4A"))
    frmOpensoldatMapEditor.backgroundColor = HexToLong(LoadString("Preferences", "BackColor", , 1024, "555555"))
    frmOpensoldatMapEditor.maxUndo = LoadInt("Preferences", "MaxUndo", , 16)
    frmOpensoldatMapEditor.sceneryVerts = LoadBoolean("Preferences", "SceneryVerts", , False)
    frmOpensoldatMapEditor.topmost = LoadBoolean("Preferences", "Topmost", , True)

    strTemp = LoadString("Preferences", "MinZoom", , 1024, "51200")
    strTemp = Replace(strTemp, ",", ".", 1, -1, vbTextCompare)
    If IsNumeric(strTemp) Then
        frmOpensoldatMapEditor.gMinZoom = Val(strTemp) / 100
    Else
        frmOpensoldatMapEditor.gMinZoom = DEFAULT_MIN_ZOOM
    End If

    strTemp = LoadString("Preferences", "MaxZoom", , 1024, "3.125")
    strTemp = Replace(strTemp, ",", ".", 1, -1, vbTextCompare)
    If IsNumeric(strTemp) Then
        frmOpensoldatMapEditor.gMaxZoom = Val(strTemp) / 100
    Else
        frmOpensoldatMapEditor.gMaxZoom = DEFAULT_MAX_ZOOM
    End If

    If frmOpensoldatMapEditor.gMinZoom = frmOpensoldatMapEditor.gMaxZoom Then
        frmOpensoldatMapEditor.gMinZoom = DEFAULT_MIN_ZOOM
        frmOpensoldatMapEditor.gMaxZoom = DEFAULT_MAX_ZOOM
    ElseIf frmOpensoldatMapEditor.gMinZoom > frmOpensoldatMapEditor.gMaxZoom Then
        sgnTemp = frmOpensoldatMapEditor.gMaxZoom
        frmOpensoldatMapEditor.gMaxZoom = frmOpensoldatMapEditor.gMinZoom
        frmOpensoldatMapEditor.gMinZoom = sgnTemp
    End If

    strTemp = LoadString("Preferences", "ResetZoom", , 1024, "100")
    strTemp = Replace(strTemp, ",", ".", 1, -1, vbTextCompare)
    If IsNumeric(strTemp) Then
        frmOpensoldatMapEditor.gResetZoom = Val(strTemp) / 100
    Else
        frmOpensoldatMapEditor.gResetZoom = DEFAULT_RESET_ZOOM
    End If

    If frmOpensoldatMapEditor.gResetZoom > frmOpensoldatMapEditor.gMaxZoom Then
        frmOpensoldatMapEditor.gResetZoom = frmOpensoldatMapEditor.gMaxZoom
    ElseIf frmOpensoldatMapEditor.gResetZoom < frmOpensoldatMapEditor.gMinZoom Then
        frmOpensoldatMapEditor.gResetZoom = frmOpensoldatMapEditor.gMinZoom
    End If

    errVal = "2"

    frmOpensoldatMapEditor.showBG = LoadBoolean("Display", "Background", , True)
    frmOpensoldatMapEditor.showPolys = LoadBoolean("Display", "Polys", , True)
    frmOpensoldatMapEditor.showTexture = LoadBoolean("Display", "Texture", , True)
    frmOpensoldatMapEditor.showWireframe = LoadBoolean("Display", "Wireframe", , False)
    frmOpensoldatMapEditor.showPoints = LoadBoolean("Display", "Points", , False)
    frmOpensoldatMapEditor.showScenery = LoadBoolean("Display", "Scenery", , True)
    frmOpensoldatMapEditor.showObjects = LoadBoolean("Display", "Objects", , True)
    frmOpensoldatMapEditor.showWaypoints = LoadBoolean("Display", "Waypoints", , False)
    frmOpensoldatMapEditor.showGrid = LoadBoolean("Display", "Grid", , False)
    frmOpensoldatMapEditor.showLights = LoadBoolean("Display", "Lights", , True)
    frmOpensoldatMapEditor.showSketch = LoadBoolean("Display", "Sketch", , True)

    frmOpensoldatMapEditor.mnuGrid.Checked = frmOpensoldatMapEditor.showGrid

    errVal = "3"

    frmOpensoldatMapEditor.currentTool = LoadByte("ToolSettings", "CurrentTool", , 1)
    frmOpensoldatMapEditor.ohSnap = LoadBoolean("ToolSettings", "SnapVertices", , True)
    frmOpensoldatMapEditor.snapToGrid = LoadBoolean("ToolSettings", "SnapToGrid", , True)
    frmOpensoldatMapEditor.fixedTexture = LoadBoolean("ToolSettings", "FixedTexture", , False)
    frmOpensoldatMapEditor.opacity = LoadInt("ToolSettings", "Opacity", , 100) / 100
    frmOpensoldatMapEditor.colorRadius = LoadInt("ToolSettings", "ColorRadius", , 16)
    gPolyColor = GetRGB(HexToLong(LoadString("ToolSettings", "CurrentColor", , , "FFFFFF")))
    frmOpensoldatMapEditor.colorMode = LoadInt("ToolSettings", "ColorMode", , 1)
    frmOpensoldatMapEditor.blendMode = LoadInt("ToolSettings", "BlendMode", , 0)
    frmOpensoldatMapEditor.snapRadius = LoadInt("ToolSettings", "SnapRadius", , 8)
    frmScenery.rotateScenery = LoadBoolean("ToolSettings", "RotateScenery", , False)
    frmScenery.scaleScenery = LoadBoolean("ToolSettings", "ScaleScenery", , False)
    frmOpensoldatMapEditor.xTexture = LoadInt("ToolSettings", "TextureWidth", , 128)
    frmOpensoldatMapEditor.yTexture = LoadInt("ToolSettings", "TextureHeight", , 128)
    frmOpensoldatMapEditor.gTextureFile = LoadString("ToolSettings", "Texture", , 1024, "banana.bmp")
    frmOpensoldatMapEditor.mnuCustomX.Checked = LoadBoolean("ToolSettings", "CustomX", , False)
    frmOpensoldatMapEditor.mnuCustomY.Checked = LoadBoolean("ToolSettings", "CustomY", , False)

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

    frmOpensoldatMapEditor.mnuRecent(0).Caption = LoadString("RecentFiles", "01", , 1024)
    frmOpensoldatMapEditor.mnuRecent(1).Caption = LoadString("RecentFiles", "02", , 1024)
    frmOpensoldatMapEditor.mnuRecent(2).Caption = LoadString("RecentFiles", "03", , 1024)
    frmOpensoldatMapEditor.mnuRecent(3).Caption = LoadString("RecentFiles", "04", , 1024)
    frmOpensoldatMapEditor.mnuRecent(4).Caption = LoadString("RecentFiles", "05", , 1024)
    frmOpensoldatMapEditor.mnuRecent(5).Caption = LoadString("RecentFiles", "06", , 1024)
    frmOpensoldatMapEditor.mnuRecent(6).Caption = LoadString("RecentFiles", "07", , 1024)
    frmOpensoldatMapEditor.mnuRecent(7).Caption = LoadString("RecentFiles", "08", , 1024)
    frmOpensoldatMapEditor.mnuRecent(8).Caption = LoadString("RecentFiles", "09", , 1024)
    frmOpensoldatMapEditor.mnuRecent(9).Caption = LoadString("RecentFiles", "10", , 1024)

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

    For i = frmOpensoldatMapEditor.mnuRecent.LBound + 1 To frmOpensoldatMapEditor.mnuRecent.UBound
        If frmOpensoldatMapEditor.mnuRecent(i).Caption = "" Then
            numRecent = numRecent + 1
            frmOpensoldatMapEditor.mnuRecent(i).Visible = False
        Else
            frmOpensoldatMapEditor.mnuRecent(i).Visible = True
        End If
    Next
    If numRecent = frmOpensoldatMapEditor.mnuRecent.Count - 1 And frmOpensoldatMapEditor.mnuRecent(frmOpensoldatMapEditor.mnuRecent.LBound).Caption = "" Then
        frmOpensoldatMapEditor.mnuRecentFiles.Enabled = False
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
    Dim isNewFile As Boolean
    Dim sNull As String

    sNull = Chr(0)
    isNewFile = Not FileExists(appPath & "\polyworks.ini")

    ' preferences
    iniString = _
        "Dir=" & frmOpensoldatMapEditor.opensoldatDir & sNull & _
        "Uncompiled=" & frmOpensoldatMapEditor.uncompDir & sNull & _
        "Prefabs=" & frmOpensoldatMapEditor.prefabDir & sNull & _
        "GridSpacing=" & frmOpensoldatMapEditor.gridSpacing & sNull & _
        "GridDiv=" & frmOpensoldatMapEditor.gridDivisions & sNull & _
        "GridColor1=" & RGBtoHex(frmOpensoldatMapEditor.gridColor1) & sNull & _
        "GridColor2=" & RGBtoHex(frmOpensoldatMapEditor.gridColor2) & sNull & _
        "GridAlpha1=" & frmOpensoldatMapEditor.gridOp1 & sNull & _
        "GridAlpha2=" & frmOpensoldatMapEditor.gridOp2 & sNull & _
        "PolySrc=" & frmOpensoldatMapEditor.polyBlendSrc & sNull & _
        "PolyDest=" & frmOpensoldatMapEditor.polyBlendDest & sNull & _
        "WireSrc=" & frmOpensoldatMapEditor.wireBlendSrc & sNull & _
        "WireDest=" & frmOpensoldatMapEditor.wireBlendDest & sNull & _
        "PointColor=" & RGBtoHex(frmOpensoldatMapEditor.pointColor) & sNull & _
        "SelectionColor=" & RGBtoHex(frmOpensoldatMapEditor.selectionColor) & sNull & _
        "BackColor=" & RGBtoHex(frmOpensoldatMapEditor.backgroundColor) & sNull & _
        "MaxUndo=" & frmOpensoldatMapEditor.maxUndo & sNull & _
        "SceneryVerts=" & CStr(frmOpensoldatMapEditor.sceneryVerts) & sNull & _
        "Topmost=" & CStr(frmOpensoldatMapEditor.topmost) & sNull & _
        "MinZoom=" & Trim(Str(frmOpensoldatMapEditor.gMaxZoom * 100)) & sNull & _
        "MaxZoom=" & Trim(Str(frmOpensoldatMapEditor.gMinZoom * 100)) & sNull & _
        "ResetZoom=" & Trim(Str(frmOpensoldatMapEditor.gResetZoom * 100)) & _
        IIf(isNewFile, vbNewLine, "") & sNull & sNull
    SaveSection "Preferences", iniString

    ' display
    iniString = _
        "Background=" & CStr(frmOpensoldatMapEditor.showBG) & sNull & _
        "Polys=" & CStr(frmOpensoldatMapEditor.showPolys) & sNull & _
        "Texture=" & CStr(frmOpensoldatMapEditor.showTexture) & sNull & _
        "Wireframe=" & CStr(frmOpensoldatMapEditor.showWireframe) & sNull & _
        "Points=" & CStr(frmOpensoldatMapEditor.showPoints) & sNull & _
        "Scenery=" & CStr(frmOpensoldatMapEditor.showScenery) & sNull & _
        "Objects=" & CStr(frmOpensoldatMapEditor.showObjects) & sNull & _
        "Waypoints=" & CStr(frmOpensoldatMapEditor.showWaypoints) & sNull & _
        "Grid=" & CStr(frmOpensoldatMapEditor.showGrid) & sNull & _
        "Lights=" & CStr(frmOpensoldatMapEditor.showLights) & sNull & _
        "Sketch=" & CStr(frmOpensoldatMapEditor.showSketch) & _
        IIf(isNewFile, vbNewLine, "") & sNull & sNull
    SaveSection "Display", iniString

    ' tool settings
    currentColor = RGB(gPolyColor.blue, gPolyColor.green, gPolyColor.red)
    iniString = _
        "CurrentTool=" & frmOpensoldatMapEditor.currentTool & sNull & _
        "SnapVertices=" & CStr(frmOpensoldatMapEditor.ohSnap) & sNull & _
        "SnapToGrid=" & CStr(frmOpensoldatMapEditor.snapToGrid) & sNull & _
        "FixedTexture=" & CStr(frmOpensoldatMapEditor.fixedTexture) & sNull & _
        "Opacity=" & (frmOpensoldatMapEditor.opacity * 100) & sNull & _
        "ColorRadius=" & frmOpensoldatMapEditor.colorRadius & sNull & _
        "CurrentColor=" & RGBtoHex(currentColor) & sNull & _
        "ColorMode=" & frmOpensoldatMapEditor.colorMode & sNull & _
        "BlendMode=" & frmOpensoldatMapEditor.blendMode & sNull & _
        "SnapRadius=" & frmOpensoldatMapEditor.snapRadius & sNull & _
        "RotateScenery=" & CStr(frmScenery.rotateScenery) & sNull & _
        "ScaleScenery=" & CStr(frmScenery.scaleScenery) & sNull & _
        "TextureWidth=" & frmOpensoldatMapEditor.xTexture & sNull & _
        "TextureHeight=" & frmOpensoldatMapEditor.yTexture & sNull & _
        "Texture=" & frmOpensoldatMapEditor.gTextureFile & sNull & _
        "CustomX=" & CStr(frmOpensoldatMapEditor.mnuCustomX.Checked) & sNull & _
        "CustomY=" & CStr(frmOpensoldatMapEditor.mnuCustomY.Checked) & _
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
        "01=" & frmOpensoldatMapEditor.mnuRecent(0).Caption & sNull & _
        "02=" & frmOpensoldatMapEditor.mnuRecent(1).Caption & sNull & _
        "03=" & frmOpensoldatMapEditor.mnuRecent(2).Caption & sNull & _
        "04=" & frmOpensoldatMapEditor.mnuRecent(3).Caption & sNull & _
        "05=" & frmOpensoldatMapEditor.mnuRecent(4).Caption & sNull & _
        "06=" & frmOpensoldatMapEditor.mnuRecent(5).Caption & sNull & _
        "07=" & frmOpensoldatMapEditor.mnuRecent(6).Caption & sNull & _
        "08=" & frmOpensoldatMapEditor.mnuRecent(7).Caption & sNull & _
        "09=" & frmOpensoldatMapEditor.mnuRecent(8).Caption & sNull & _
        "10=" & frmOpensoldatMapEditor.mnuRecent(9).Caption & _
        IIf(isNewFile, vbNewLine, "") & sNull & sNull
    SaveSection "RecentFiles", iniString

    ' polytype colors
    iniString = _
        "OnlyBullets=" & RGBtoHex(gPolyTypeColors(1)) & sNull & _
        "OnlyPlayer=" & RGBtoHex(gPolyTypeColors(2)) & sNull & _
        "DoesntCollide=" & RGBtoHex(gPolyTypeColors(3)) & sNull & _
        "Ice=" & RGBtoHex(gPolyTypeColors(4)) & sNull & _
        "Deadly=" & RGBtoHex(gPolyTypeColors(5)) & sNull & _
        "BloodyDeadly=" & RGBtoHex(gPolyTypeColors(6)) & sNull & _
        "Hurts=" & RGBtoHex(gPolyTypeColors(7)) & sNull & _
        "Regenerates=" & RGBtoHex(gPolyTypeColors(8)) & sNull & _
        "Lava=" & RGBtoHex(gPolyTypeColors(9)) & sNull & _
        "TeamBullets=" & RGBtoHex(gPolyTypeColors(10)) & sNull & _
        "TeamPlayers=" & RGBtoHex(gPolyTypeColors(11)) & sNull & _
        "Bouncy=" & RGBtoHex(gPolyTypeColors(18)) & sNull & _
        "Explosive=" & RGBtoHex(gPolyTypeColors(19)) & sNull & _
        "HurtFlaggers=" & RGBtoHex(gPolyTypeColors(20)) & sNull & _
        "OnlyFlagger=" & RGBtoHex(gPolyTypeColors(21)) & sNull & _
        "NonFlagger=" & RGBtoHex(gPolyTypeColors(22)) & sNull & _
        "FlagCollides=" & RGBtoHex(gPolyTypeColors(23)) & sNull & _
        "Back=" & RGBtoHex(gPolyTypeColors(24)) & sNull & _
        "BackTransition=" & RGBtoHex(gPolyTypeColors(25)) & _
        IIf(isNewFile, vbNewLine, "") & sNull & sNull
    SaveSection "PolyTypeColors", iniString

    ' gfx dir
    iniString = "Dir=" & gfxDir & sNull & sNull
    SaveSection "gfx", iniString

    ' workspace
    isNewFile = Not FileExists(appPath & "\workspace\current.ini")

    iniString = _
        "WindowState=" & frmOpensoldatMapEditor.Tag & sNull & _
        "Width=" & frmOpensoldatMapEditor.formWidth & sNull & _
        "Height=" & frmOpensoldatMapEditor.formHeight & sNull & _
        "Left=" & frmOpensoldatMapEditor.formLeft & sNull & _
        "Top=" & frmOpensoldatMapEditor.formTop & _
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

    frmOpensoldatMapEditor.Tag = LoadInt("Main", "WindowState", appPath & "\workspace\" & theFileName)
    frmOpensoldatMapEditor.formWidth = LoadInt("Main", "Width", appPath & "\workspace\" & theFileName)
    frmOpensoldatMapEditor.formHeight = LoadInt("Main", "Height", appPath & "\workspace\" & theFileName)
    frmOpensoldatMapEditor.formLeft = LoadInt("Main", "Left", appPath & "\workspace\" & theFileName)
    frmOpensoldatMapEditor.formTop = LoadInt("Main", "Top", appPath & "\workspace\" & theFileName)

    frmOpensoldatMapEditor.picResize.Top = frmOpensoldatMapEditor.formHeight - frmOpensoldatMapEditor.picResize.Height
    frmOpensoldatMapEditor.picResize.Left = frmOpensoldatMapEditor.formWidth - frmOpensoldatMapEditor.picResize.Width

    If frmOpensoldatMapEditor.Tag = vbNormal Then
        frmOpensoldatMapEditor.Width = frmOpensoldatMapEditor.formWidth * Screen.TwipsPerPixelX
        frmOpensoldatMapEditor.Height = frmOpensoldatMapEditor.formHeight * Screen.TwipsPerPixelY
        frmOpensoldatMapEditor.Left = frmOpensoldatMapEditor.formLeft * Screen.TwipsPerPixelX
        frmOpensoldatMapEditor.Top = frmOpensoldatMapEditor.formTop * Screen.TwipsPerPixelY
    Else
        frmOpensoldatMapEditor.Tag = vbNormal
        frmOpensoldatMapEditor.Width = frmOpensoldatMapEditor.formWidth * Screen.TwipsPerPixelX
        frmOpensoldatMapEditor.Height = frmOpensoldatMapEditor.formHeight * Screen.TwipsPerPixelY
        frmOpensoldatMapEditor.Left = frmOpensoldatMapEditor.formLeft * Screen.TwipsPerPixelX
        frmOpensoldatMapEditor.Top = frmOpensoldatMapEditor.formTop * Screen.TwipsPerPixelY
        frmOpensoldatMapEditor.MaximizeBorderLessForm
        frmOpensoldatMapEditor.picResize.Visible = False
    End If

    frmOpensoldatMapEditor.tvwScenery.Height = frmOpensoldatMapEditor.formHeight - 41 - 20

    frmOpensoldatMapEditor.mnuTools.Checked = LoadString("Tools", "Visible", appPath & "\workspace\" & theFileName)
    frmTools.xPos = LoadInt("Tools", "Left", appPath & "\workspace\" & theFileName)
    frmTools.yPos = LoadInt("Tools", "Top", appPath & "\workspace\" & theFileName)
    frmTools.collapsed = LoadString("Tools", "Collapsed", appPath & "\workspace\" & theFileName)
    frmTools.Tag = IIf(LoadString("Tools", "Snapped", appPath & "\workspace\" & theFileName) = "True", "snap", "")

    frmOpensoldatMapEditor.mnuDisplay.Checked = LoadString("Display", "Visible", appPath & "\workspace\" & theFileName)
    frmDisplay.xPos = LoadInt("Display", "Left", appPath & "\workspace\" & theFileName)
    frmDisplay.yPos = LoadInt("Display", "Top", appPath & "\workspace\" & theFileName)
    frmDisplay.collapsed = LoadString("Display", "Collapsed", appPath & "\workspace\" & theFileName)
    frmDisplay.Tag = IIf(LoadString("Display", "Snapped", appPath & "\workspace\" & theFileName) = "True", "snap", "")

    frmOpensoldatMapEditor.mnuInfo.Checked = LoadString("Properties", "Visible", appPath & "\workspace\" & theFileName)
    frmInfo.xPos = LoadInt("Properties", "Left", appPath & "\workspace\" & theFileName)
    frmInfo.yPos = LoadInt("Properties", "Top", appPath & "\workspace\" & theFileName)
    frmInfo.collapsed = LoadString("Properties", "Collapsed", appPath & "\workspace\" & theFileName)
    frmInfo.Tag = IIf(LoadString("Properties", "Snapped", appPath & "\workspace\" & theFileName) = "True", "snap", "")

    frmOpensoldatMapEditor.mnuPalette.Checked = LoadString("Palette", "Visible", appPath & "\workspace\" & theFileName)
    frmPalette.xPos = LoadInt("Palette", "Left", appPath & "\workspace\" & theFileName)
    frmPalette.yPos = LoadInt("Palette", "Top", appPath & "\workspace\" & theFileName)
    frmPalette.collapsed = LoadString("Palette", "Collapsed", appPath & "\workspace\" & theFileName)
    frmPalette.Tag = IIf(LoadString("Palette", "Snapped", appPath & "\workspace\" & theFileName) = "True", "snap", "")

    If Not bSkipScenery Then
        frmOpensoldatMapEditor.mnuScenery.Checked = LoadString("Scenery", "Visible", appPath & "\workspace\" & theFileName)
        frmScenery.xPos = LoadInt("Scenery", "Left", appPath & "\workspace\" & theFileName)
        frmScenery.yPos = LoadInt("Scenery", "Top", appPath & "\workspace\" & theFileName)
        frmScenery.collapsed = LoadString("Scenery", "Collapsed", appPath & "\workspace\" & theFileName)
        frmScenery.Tag = IIf(LoadString("Scenery", "Snapped", appPath & "\workspace\" & theFileName) = "True", "snap", "")
    End If

    frmOpensoldatMapEditor.mnuWaypoints.Checked = LoadString("Waypoints", "Visible", appPath & "\workspace\" & theFileName)
    frmWaypoints.xPos = LoadInt("Waypoints", "Left", appPath & "\workspace\" & theFileName)
    frmWaypoints.yPos = LoadInt("Waypoints", "Top", appPath & "\workspace\" & theFileName)
    frmWaypoints.collapsed = LoadString("Waypoints", "Collapsed", appPath & "\workspace\" & theFileName)
    frmWaypoints.Tag = IIf(LoadString("Waypoints", "Snapped", appPath & "\workspace\" & theFileName) = "True", "snap", "")

    frmOpensoldatMapEditor.mnuTexture.Checked = LoadString("Texture", "Visible", appPath & "\workspace\" & theFileName)
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
        "Visible=" & CStr(window.Visible) & sNull & _
        "Left=" & leftVal & sNull & _
        "Top=" & topVal & sNull & _
        "Collapsed=" & CStr(collapsed) & sNull & _
        "Snapped=" & CStr(Len(window.Tag) > 0) & _
        IIf(isNewFile, vbNewLine, "") & sNull & sNull

    SaveSection sectionName, iniString, appPath & "\workspace\" & theFileName

End Sub


' functions - private

Private Function SetIdePath() As Boolean

    appPath = appPath & "\pwinstall"
    SetIdePath = True

End Function
