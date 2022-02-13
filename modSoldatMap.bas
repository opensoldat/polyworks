Attribute VB_Name = "modSoldatMap"
Option Explicit

' Fix vb6 ide casing changes
#If False Then
    Public FileName, color, token, A, R, G, B, commonDialog, value, Val, X, Y, Z, Left, hWnd, Mid, Right
    'Public FileName, color, token, A, R, G, B, commonDialog, value, Val, X, Y, Z, Left, hWnd, Mid, Right
#End If

' loading and saving soldat maps

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
