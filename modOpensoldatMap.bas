Attribute VB_Name = "modOpensoldatMap"
Option Explicit

' loading and saving opensoldat maps


' Fix vb6 ide casing changes
#If False Then
    Public FileName, color, token, A, R, G, B, commonDialog, value, Val, X, Y, Z, Left, hWnd, Mid, Right, BackColor
    'Public FileName, color, token, A, R, G, B, commonDialog, value, Val, X, Y, Z, Left, hWnd, Mid, Right, BackColor
#End If

' types

Public Type TImageInfo
    Width       As Integer
    miplevels   As Integer
    Height      As Integer
    depth       As Integer
End Type

Public Type TVertexData
    vertex(1 To 3)   As Byte
    polyType         As Byte
    color(1 To 3)    As TColor
End Type

Public Type TTriangle
    vertex(1 To 3) As D3DVECTOR2
End Type

Public Type TLightSource
    selected    As Byte
    color       As TColor
    intensity   As Single
    range       As Integer
    X           As Single
    Y           As Single
    Z           As Single
End Type


' map types

Public Type TCustomVertex
    X       As Single
    Y       As Single
    Z       As Single
    rhw     As Single
    color   As Long
    tu      As Single
    tv      As Single
End Type

Public Type TSketchVertex
    X As Single
    Y As Single
    Z As Single
End Type

Public Type TSketchLine
    vertex(1 To 2) As TSketchVertex
End Type

Public Type TVertexHit
    X As Single ' sin of angle
    Y As Single ' cos of angle
    Z As Single ' 0
End Type

Public Type TPolyHit
    vertex(1 To 3) As TVertexHit
End Type

Public Type TPolygon
    vertex(1 To 3) As TCustomVertex
    Perp           As TPolyHit
End Type

Public Type TLine
    vertex(1 To 2) As TCustomVertex
End Type

Public Type TProp
    active      As Boolean
    Style       As Integer
    Width       As Long
    Height      As Long
    X           As Single
    Y           As Single
    rotation    As Single
    ScaleX      As Single
    ScaleY      As Single
    alpha       As Long
    color       As Long
    level       As Long
End Type

Public Type TScenery
    Style       As Integer
    Translation As D3DVECTOR2
    rotation    As Single
    Scaling     As D3DVECTOR2
    alpha       As Byte
    color       As Long
    level       As Byte
    selected    As Byte
    screenTr    As D3DVECTOR2
End Type

Public Type TSpawnPoint
    active  As Long ' Boolean
    X       As Single
    Y       As Single
    Team    As Long
End Type

Public Type TSaveSpawnPoint
    active  As Long ' Boolean
    X       As Long
    Y       As Long
    Team    As Long
End Type

Public Type TCollider
    active  As Long    ' Boolean
    X       As Single
    Y       As Single
    radius  As Single
End Type

Public Type TOptions
    mapName(0 To 38)        As Byte ' String * 39
    textureName(0 To 24)    As Byte ' String * 25
    backgroundColor1        As Long
    backgroundColor2        As Long
    StartJet                As Long
    GrenadePacks            As Byte
    Medikits                As Byte
    Weather                 As Byte
    Steps                   As Byte
    MapRandomID             As Long ' Integer
End Type

Public Type TMapFile_Polygon
   Poly     As TPolygon
   polyType As Byte
End Type

Public Type TMapFile_Scenery
   sceneryName(0 To 50) As Byte
   Date                 As Long
End Type

Public Type TextureData
    Width   As Integer
    Height  As Integer
    reScale As D3DVECTOR2
    Texture As Direct3DTexture8
End Type

Public Type TNewWaypoint
    active                  As Long
    id                      As Long
    X                       As Long
    Y                       As Long
    Left                    As Byte
    Right                   As Byte
    up                      As Byte
    down                    As Byte
    m2                      As Byte
    pathNum                 As Byte
    special                 As Byte
    crap(1 To 5)            As Byte
    connectionsNum          As Long
    Connections(1 To 20)    As Long
End Type

Public Type TWaypoint
    tempIndex       As Integer
    selected        As Boolean
    X               As Single
    Y               As Single
    wayType(0 To 4) As Boolean
    special         As Byte
    pathNum         As Byte
    numConnections  As Byte
End Type

Public Type TConnection
    point1 As Integer
    point2 As Integer
End Type
