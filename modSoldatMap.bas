Attribute VB_Name = "modSoldatMap"
Option Explicit

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
    z           As Single
End Type
