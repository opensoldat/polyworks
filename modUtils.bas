Attribute VB_Name = "modUtils"
Option Explicit

Public Type TColor
    red     As Byte
    green   As Byte
    blue    As Byte
End Type

Public Function getRGB(DecValue As Long) As TColor

    Dim hexValue As String

    hexValue = Hex$(Val(DecValue))

    If Len(hexValue) < 6 Then
        hexValue = String$(6 - Len(hexValue), "0") + hexValue
    End If

    getRGB.blue = CLng("&H" + right$(hexValue, 2))
    hexValue = left$(hexValue, Len(hexValue) - 2)
    getRGB.green = CLng("&H" + right$(hexValue, 2))
    hexValue = left$(hexValue, Len(hexValue) - 2)
    getRGB.red = CLng("&H" + right$(hexValue, 2))

End Function

Public Function getAlpha(tehColor As Long) As Byte

    Dim hexValue As String

    hexValue = Hex$(Val(tehColor))

    If Len(hexValue) <= 6 Then
        getAlpha = 0
    Else
        If Len(hexValue) < 8 Then
            hexValue = String$(8 - Len(hexValue), "0") + hexValue
        End If
        getAlpha = CLng("&H" + left$(hexValue, 2))
    End If

End Function

Public Function ARGB(ByVal alphaVal As Byte, clrVal As Long) As Long

    Dim clrString As String

    clrString = Hex$(clrVal)
    If Len(clrString) < 6 Then
        clrString = String$(6 - Len(clrString), "0") & clrString
    ElseIf Len(clrString) > 6 Then
        clrString = right$(clrString, 6)
    End If
    If Len(Hex$(alphaVal)) = 1 Then
        clrString = "0" + Hex$(alphaVal) & clrString
    ElseIf Len(Hex$(alphaVal)) = 2 Then
        clrString = Hex$(alphaVal) & clrString
    End If
    ARGB = CLng("&H" & clrString)

End Function

Public Function makeColor(red As Byte, green As Byte, blue As Byte) As TColor

    makeColor.red = red
    makeColor.green = green
    makeColor.blue = blue

End Function


Public Function diffVal(val1 As Byte, val2 As Byte) As Byte

    If val1 > val2 Then
        diffVal = val1 - val2
    Else
        diffVal = val2 - val1
    End If

End Function

Public Function lowerVal(val1 As Byte, val2 As Byte) As Byte

    If val1 < val2 Then
        lowerVal = val1
    Else
        lowerVal = val2
    End If

End Function

Public Function higherVal(val1 As Byte, val2 As Byte) As Byte

    If val1 > val2 Then
        higherVal = val1
    Else
        higherVal = val2
    End If

End Function

