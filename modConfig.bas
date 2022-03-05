Attribute VB_Name = "modConfig"
Option Explicit

' Fix vb6 ide casing changes
#If False Then
    Public FileName, color, token, A, R, G, B, commonDialog, value, Val, X, Y, Z, Left, hWnd, Mid, Right, BackColor
    'Public FileName, color, token, A, R, G, B, commonDialog, value, Val, X, Y, Z, Left, hWnd, Mid, Right, BackColor
#End If

' loading and saving config files, workspaces, skins goes here


Public gPolyClr As TColor

Public gPolyTypeClrs(0 To 25) As Long


Private Const DEFAULT_MAX_ZOOM As Single = 512
Private Const DEFAULT_MIN_ZOOM As Single = 0.03125
Private Const DEFAULT_RESET_ZOOM As Single = 1


Public Sub loadINI()

    On Error GoTo ErrorHandler

    appPath = App.path

    Debug.Assert SetIdePath  ' workaround for debugging with ide

    Dim i As Integer
    Dim numRecent As Integer
    Dim strTemp As String
    Dim sgnTemp As Single
    Dim errVal As String

    errVal = "1"

    frmSoldatMapEditor.soldatDir = loadString("Preferences", "Dir", , 1024)
    frmSoldatMapEditor.uncompDir = loadString("Preferences", "Uncompiled", , 1024)
    frmSoldatMapEditor.prefabDir = loadString("Preferences", "Prefabs", , 1024)

    frmSoldatMapEditor.gridSpacing = loadInt("Preferences", "GridSpacing")
    frmSoldatMapEditor.gridDivisions = loadInt("Preferences", "GridDiv")
    frmSoldatMapEditor.gridColor1 = HexToLong(loadString("Preferences", "GridColor1"))
    frmSoldatMapEditor.gridColor2 = HexToLong(loadString("Preferences", "GridColor2"))
    frmSoldatMapEditor.gridOp1 = loadInt("Preferences", "GridAlpha1")
    frmSoldatMapEditor.gridOp2 = loadInt("Preferences", "GridAlpha2")
    frmSoldatMapEditor.polyBlendSrc = loadInt("Preferences", "PolySrc")
    frmSoldatMapEditor.polyBlendDest = loadInt("Preferences", "PolyDest")
    frmSoldatMapEditor.wireBlendSrc = loadInt("Preferences", "WireSrc")
    frmSoldatMapEditor.wireBlendDest = loadInt("Preferences", "WireDest")
    frmSoldatMapEditor.pointColor = HexToLong(loadString("Preferences", "PointColor"))
    frmSoldatMapEditor.selectionColor = HexToLong(loadString("Preferences", "SelectionColor"))
    frmSoldatMapEditor.backClr = HexToLong(loadString("Preferences", "BackColor"))
    frmSoldatMapEditor.max_undo = loadInt("Preferences", "MaxUndo")
    frmSoldatMapEditor.sceneryVerts = loadString("Preferences", "SceneryVerts")
    frmSoldatMapEditor.topmost = loadString("Preferences", "Topmost")

    strTemp = loadString("Preferences", "MinZoom")
    If IsNumeric(strTemp) Then
        frmSoldatMapEditor.gMinZoom = CSng(strTemp) / 100
    Else
       frmSoldatMapEditor.gMinZoom = DEFAULT_MIN_ZOOM
    End If

    strTemp = loadString("Preferences", "MaxZoom")
    If IsNumeric(strTemp) Then
        frmSoldatMapEditor.gMaxZoom = CSng(strTemp) / 100
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

    strTemp = loadString("Preferences", "ResetZoom")
    If IsNumeric(strTemp) Then
        frmSoldatMapEditor.gResetZoom = CSng(strTemp) / 100
    Else
        frmSoldatMapEditor.gResetZoom = DEFAULT_RESET_ZOOM
    End If

    If frmSoldatMapEditor.gResetZoom > frmSoldatMapEditor.gMaxZoom Then
        frmSoldatMapEditor.gResetZoom = frmSoldatMapEditor.gMaxZoom
    ElseIf frmSoldatMapEditor.gResetZoom < frmSoldatMapEditor.gMinZoom Then
        frmSoldatMapEditor.gResetZoom = frmSoldatMapEditor.gMinZoom
    End If

    errVal = "2"

    frmSoldatMapEditor.showBG = loadString("Display", "Background")
    frmSoldatMapEditor.showPolys = loadString("Display", "Polys")
    frmSoldatMapEditor.showTexture = loadString("Display", "Texture")
    frmSoldatMapEditor.showWireframe = loadString("Display", "Wireframe")
    frmSoldatMapEditor.showPoints = loadString("Display", "Points")
    frmSoldatMapEditor.showScenery = loadString("Display", "Scenery")
    frmSoldatMapEditor.showObjects = loadString("Display", "Objects")
    frmSoldatMapEditor.showWaypoints = loadString("Display", "Waypoints")
    frmSoldatMapEditor.showGrid = loadString("Display", "Grid")
    frmSoldatMapEditor.showLights = loadString("Display", "Lights")
    frmSoldatMapEditor.showSketch = loadString("Display", "Sketch")
    
    frmSoldatMapEditor.mnuGrid.Checked = frmSoldatMapEditor.showGrid

    errVal = "3"

    frmSoldatMapEditor.currentTool = loadInt("ToolSettings", "CurrentTool")
    frmSoldatMapEditor.ohSnap = loadString("ToolSettings", "SnapVertices")
    frmSoldatMapEditor.snapToGrid = loadString("ToolSettings", "SnapToGrid")
    frmSoldatMapEditor.fixedTexture = loadString("ToolSettings", "FixedTexture")
    frmSoldatMapEditor.opacity = loadInt("ToolSettings", "Opacity") / 100
    frmSoldatMapEditor.clrRadius = loadInt("ToolSettings", "ColorRadius")
    gPolyClr = getRGB(HexToLong(loadString("ToolSettings", "CurrentColor")))
    frmSoldatMapEditor.colorMode = loadInt("ToolSettings", "ColorMode")
    frmSoldatMapEditor.blendMode = loadInt("ToolSettings", "BlendMode")
    frmSoldatMapEditor.snapRadius = loadInt("ToolSettings", "SnapRadius")
    frmScenery.rotateScenery = loadString("ToolSettings", "RotateScenery")
    frmScenery.scaleScenery = loadString("ToolSettings", "ScaleScenery")
    frmSoldatMapEditor.xTexture = loadInt("ToolSettings", "TextureWidth")
    frmSoldatMapEditor.yTexture = loadInt("ToolSettings", "TextureHeight")
    frmSoldatMapEditor.gTextureFile = loadString("ToolSettings", "Texture", , 1024)
    frmSoldatMapEditor.mnuCustomX.Checked = loadString("ToolSettings", "CustomX")
    frmSoldatMapEditor.mnuCustomY.Checked = loadString("ToolSettings", "CustomY")

    errVal = "4"

    frmTools.setHotKey 0, loadInt("HotKeys", "Move")
    frmTools.setHotKey 1, loadInt("HotKeys", "Create")
    frmTools.setHotKey 2, loadInt("HotKeys", "VertexSelection")
    frmTools.setHotKey 3, loadInt("HotKeys", "PolySelection")
    frmTools.setHotKey 4, loadInt("HotKeys", "VertexColor")
    frmTools.setHotKey 5, loadInt("HotKeys", "PolyColor")
    frmTools.setHotKey 6, loadInt("HotKeys", "Texture")
    frmTools.setHotKey 7, loadInt("HotKeys", "Scenery")
    frmTools.setHotKey 8, loadInt("HotKeys", "Waypoints")
    frmTools.setHotKey 9, loadInt("HotKeys", "Objects")
    frmTools.setHotKey 10, loadInt("HotKeys", "ColorPicker")
    frmTools.setHotKey 11, loadInt("HotKeys", "Sketch")
    frmTools.setHotKey 12, loadInt("HotKeys", "Lights")
    frmTools.setHotKey 13, loadInt("HotKeys", "DepthMap")

    errVal = "5"

    frmWaypoints.setWayptKey 0, loadInt("WaypointKeys", "Left")
    frmWaypoints.setWayptKey 1, loadInt("WaypointKeys", "Right")
    frmWaypoints.setWayptKey 2, loadInt("WaypointKeys", "Up")
    frmWaypoints.setWayptKey 3, loadInt("WaypointKeys", "Down")
    frmWaypoints.setWayptKey 4, loadInt("WaypointKeys", "Fly")

    errVal = "6"

    frmDisplay.setLayerKey 0, loadInt("LayerKeys", "Background")
    frmDisplay.setLayerKey 1, loadInt("LayerKeys", "Polys")
    frmDisplay.setLayerKey 2, loadInt("LayerKeys", "Texture")
    frmDisplay.setLayerKey 3, loadInt("LayerKeys", "Wireframe")
    frmDisplay.setLayerKey 4, loadInt("LayerKeys", "Points")
    frmDisplay.setLayerKey 5, loadInt("LayerKeys", "Scenery")
    frmDisplay.setLayerKey 6, loadInt("LayerKeys", "Objects")
    frmDisplay.setLayerKey 7, loadInt("LayerKeys", "Waypoints")

    errVal = "7"

    frmSoldatMapEditor.mnuRecent(0).Caption = loadString("RecentFiles", "01", , 1024)
    frmSoldatMapEditor.mnuRecent(1).Caption = loadString("RecentFiles", "02", , 1024)
    frmSoldatMapEditor.mnuRecent(2).Caption = loadString("RecentFiles", "03", , 1024)
    frmSoldatMapEditor.mnuRecent(3).Caption = loadString("RecentFiles", "04", , 1024)
    frmSoldatMapEditor.mnuRecent(4).Caption = loadString("RecentFiles", "05", , 1024)
    frmSoldatMapEditor.mnuRecent(5).Caption = loadString("RecentFiles", "06", , 1024)
    frmSoldatMapEditor.mnuRecent(6).Caption = loadString("RecentFiles", "07", , 1024)
    frmSoldatMapEditor.mnuRecent(7).Caption = loadString("RecentFiles", "08", , 1024)
    frmSoldatMapEditor.mnuRecent(8).Caption = loadString("RecentFiles", "09", , 1024)
    frmSoldatMapEditor.mnuRecent(9).Caption = loadString("RecentFiles", "10", , 1024)

    errVal = "8"

    gPolyTypeClrs(1) = CLng("&H" + (loadString("PolyTypeColors", "OnlyBullets")))
    gPolyTypeClrs(2) = CLng("&H" + (loadString("PolyTypeColors", "OnlyPlayer")))
    gPolyTypeClrs(3) = CLng("&H" + (loadString("PolyTypeColors", "DoesntCollide")))
    gPolyTypeClrs(4) = CLng("&H" + (loadString("PolyTypeColors", "Ice")))
    gPolyTypeClrs(5) = CLng("&H" + (loadString("PolyTypeColors", "Deadly")))
    gPolyTypeClrs(6) = CLng("&H" + (loadString("PolyTypeColors", "BloodyDeadly")))
    gPolyTypeClrs(7) = CLng("&H" + (loadString("PolyTypeColors", "Hurts")))
    gPolyTypeClrs(8) = CLng("&H" + (loadString("PolyTypeColors", "Regenerates")))
    gPolyTypeClrs(9) = CLng("&H" + (loadString("PolyTypeColors", "Lava")))
    gPolyTypeClrs(10) = CLng("&H" + (loadString("PolyTypeColors", "TeamBullets")))
    gPolyTypeClrs(11) = CLng("&H" + (loadString("PolyTypeColors", "TeamPlayers")))
    gPolyTypeClrs(12) = gPolyTypeClrs(10)
    gPolyTypeClrs(13) = gPolyTypeClrs(11)
    gPolyTypeClrs(14) = gPolyTypeClrs(10)
    gPolyTypeClrs(15) = gPolyTypeClrs(11)
    gPolyTypeClrs(16) = gPolyTypeClrs(10)
    gPolyTypeClrs(17) = gPolyTypeClrs(11)
    gPolyTypeClrs(18) = CLng("&H" + (loadString("PolyTypeColors", "Bouncy")))
    gPolyTypeClrs(19) = CLng("&H" + (loadString("PolyTypeColors", "Explosive")))
    gPolyTypeClrs(20) = CLng("&H" + (loadString("PolyTypeColors", "HurtFlaggers")))
    gPolyTypeClrs(21) = CLng("&H" + (loadString("PolyTypeColors", "OnlyFlagger")))
    gPolyTypeClrs(22) = CLng("&H" + (loadString("PolyTypeColors", "NonFlagger")))
    gPolyTypeClrs(23) = CLng("&H" + (loadString("PolyTypeColors", "FlagCollides")))
    gPolyTypeClrs(24) = CLng("&H" + (loadString("PolyTypeColors", "Back")))
    gPolyTypeClrs(25) = CLng("&H" + (loadString("PolyTypeColors", "BackTransition")))

    errVal = "9"

    gfxDir = loadString("gfx", "Dir", , 1024)

    If gfxDir = "" Then gfxDir = "gfx"

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

    MsgBox "Error loading ini file" & vbNewLine & Error$ & vbNewLine & errVal

End Sub

Public Sub saveSettings()

    Dim X As Integer
    Dim Y As Integer
    Dim i As Integer
    Dim KeyCode As Byte

    Dim iniString As String
    Dim currentColor As Long
    Dim sNull As String
    sNull = Chr$(0)
    Dim isNewFile As Boolean
    isNewFile = False

    ' preferences
    iniString = _
        "Dir=" & soldatDir & sNull & _
        "Uncompiled=" & uncompDir & sNull & _
        "Prefabs=" & prefabDir & sNull & _
        "GridSpacing=" & gridSpacing & sNull & _
        "GridDiv=" & gridDivisions & sNull & _
        "GridColor1=" & RGBtoHex(gridColor1) & sNull & _
        "GridColor2=" & RGBtoHex(gridColor2) & sNull & _
        "GridAlpha1=" & gridOp1 & sNull & _
        "GridAlpha2=" & gridOp2 & sNull & _
        "PolySrc=" & polyBlendSrc & sNull & _
        "PolyDest=" & polyBlendDest & sNull & _
        "WireSrc=" & wireBlendSrc & sNull & _
        "WireDest=" & wireBlendDest & sNull & _
        "PointColor=" & RGBtoHex(pointColor) & sNull & _
        "SelectionColor=" & RGBtoHex(selectionColor) & sNull & _
        "BackColor=" & RGBtoHex(backClr) & sNull & _
        "MaxUndo=" & max_undo & sNull & _
        "SceneryVerts=" & sceneryVerts & sNull & _
        "Topmost=" & topmost & sNull & _
        "MinZoom=" & gMaxZoom * 100 & sNull & _
        "MaxZoom=" & gMinZoom * 100 & sNull & _
        "ResetZoom=" & gResetZoom * 100 & sNull & sNull
    saveSection "Preferences", iniString

    ' display
    iniString = _
        "Background=" & showBG & sNull & _
        "Polys=" & showPolys & sNull & _
        "Texture=" & showTexture & sNull & _
        "Wireframe=" & showWireframe & sNull & _
        "Points=" & showPoints & sNull & _
        "Scenery=" & showScenery & sNull & _
        "Objects=" & showObjects & sNull & _
        "Waypoints=" & showWaypoints & sNull & _
        "Grid=" & showGrid & sNull & _
        "Lights=" & showLights & sNull & _
        "Sketch=" & showSketch & sNull & sNull
    saveSection "Display", iniString

    ' tool settings
    currentColor = RGB(gPolyClr.blue, gPolyClr.green, gPolyClr.red)
    iniString = _
        "CurrentTool=" & currentTool & sNull & _
        "SnapVertices=" & ohSnap & sNull & _
        "SnapToGrid=" & snapToGrid & sNull & _
        "FixedTexture=" & fixedTexture & sNull & _
        "Opacity=" & (opacity * 100) & sNull & _
        "ColorRadius=" & clrRadius & sNull & _
        "CurrentColor=" & RGBtoHex(currentColor) & sNull & _
        "ColorMode=" & colorMode & sNull & _
        "BlendMode=" & blendMode & sNull & _
        "SnapRadius=" & snapRadius & sNull & _
        "RotateScenery=" & frmScenery.rotateScenery & sNull & _
        "ScaleScenery=" & frmScenery.scaleScenery & sNull & _
        "TextureWidth=" & xTexture & sNull & _
        "TextureHeight=" & yTexture & sNull & _
        "Texture=" & gTextureFile & sNull & _
        "CustomX=" & mnuCustomX.Checked & sNull & _
        "CustomY=" & mnuCustomY.Checked & sNull & sNull
    saveSection "ToolSettings", iniString

    ' hotkeys
    iniString = _
        "Move=" & frmTools.getHotKey(0) & sNull & _
        "Create=" & frmTools.getHotKey(1) & sNull & _
        "VertexSelection=" & frmTools.getHotKey(2) & sNull & _
        "PolySelection=" & frmTools.getHotKey(3) & sNull & _
        "VertexColor=" & frmTools.getHotKey(4) & sNull & _
        "PolyColor=" & frmTools.getHotKey(5) & sNull & _
        "Texture=" & frmTools.getHotKey(6) & sNull & _
        "Scenery=" & frmTools.getHotKey(7) & sNull & _
        "Waypoints=" & frmTools.getHotKey(8) & sNull & _
        "Objects=" & frmTools.getHotKey(9) & sNull & _
        "ColorPicker=" & frmTools.getHotKey(10) & sNull & _
        "Sketch=" & frmTools.getHotKey(11) & sNull & _
        "Lights=" & frmTools.getHotKey(12) & sNull & _
        "Depthmap=" & frmTools.getHotKey(13) & sNull & sNull
    saveSection "HotKeys", iniString

    ' waypoint keys
    iniString = _
        "Left=" & frmWaypoints.getWayptKey(0) & sNull & _
        "Right=" & frmWaypoints.getWayptKey(1) & sNull & _
        "Up=" & frmWaypoints.getWayptKey(2) & sNull & _
        "Down=" & frmWaypoints.getWayptKey(3) & sNull & _
        "Fly=" & frmWaypoints.getWayptKey(4) & sNull & sNull
    saveSection "WaypointKeys", iniString

    ' layer keys
    iniString = _
        "Background=" & frmDisplay.getLayerKey(0) & sNull & _
        "Polys=" & frmDisplay.getLayerKey(1) & sNull & _
        "Texture=" & frmDisplay.getLayerKey(2) & sNull & _
        "Wireframe=" & frmDisplay.getLayerKey(3) & sNull & _
        "Points=" & frmDisplay.getLayerKey(4) & sNull & _
        "Scenery=" & frmDisplay.getLayerKey(5) & sNull & _
        "Objects=" & frmDisplay.getLayerKey(6) & sNull & _
        "Waypoints=" & frmDisplay.getLayerKey(7) & sNull & sNull
    saveSection "LayerKeys", iniString

    ' palette
    frmPalette.savePalette appPath & "\palettes\current.txt"

    ' workspace
    isNewFile = Not FileExists(appPath & "\workspace\current.ini")

    iniString = _
        "WindowState=" & Me.Tag & sNull & _
        "Width=" & formWidth & sNull & _
        "Height=" & formHeight & sNull & _
        "Left=" & formLeft & sNull & _
        "Top=" & formTop & _
        IIf(isNewFile, vbNewLine, "") & sNull & sNull
    saveSection "Main", iniString, appPath & "\workspace\current.ini"

    saveWindow "Tools", frmTools, False, isNewFile
    saveWindow "Display", frmDisplay, frmDisplay.collapsed, isNewFile
    saveWindow "Properties", frmInfo, frmInfo.collapsed, isNewFile
    saveWindow "Palette", frmPalette, frmPalette.collapsed, isNewFile
    saveWindow "Scenery", frmScenery, frmScenery.collapsed, isNewFile
    saveWindow "Waypoints", frmWaypoints, frmWaypoints.collapsed, isNewFile
    saveWindow "Texture", frmTexture, frmTexture.collapsed, isNewFile

    ' recent files
    iniString = _
        "01=" & mnuRecent(0).Caption & sNull & _
        "02=" & mnuRecent(1).Caption & sNull & _
        "03=" & mnuRecent(2).Caption & sNull & _
        "04=" & mnuRecent(3).Caption & sNull & _
        "05=" & mnuRecent(4).Caption & sNull & _
        "06=" & mnuRecent(5).Caption & sNull & _
        "07=" & mnuRecent(6).Caption & sNull & _
        "08=" & mnuRecent(7).Caption & sNull & _
        "09=" & mnuRecent(8).Caption & sNull & _
        "10=" & mnuRecent(9).Caption & sNull & sNull
    saveSection "RecentFiles", iniString

    ' gfx dir
    iniString = "Dir=" & gfxDir & sNull & sNull
    saveSection "gfx", iniString

End Sub

Private Function SetIdePath() As Boolean

  appPath = appPath & "\pwinstall"
  SetIdePath = True

End Function
