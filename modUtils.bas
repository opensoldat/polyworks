Attribute VB_Name = "modUtils"
Option Explicit

' Fix vb6 ide casing changes
#If False Then
    Public FileName, color, token, A, R, G, B, commonDialog, value, Val, X, Y, Z, Left, hWnd, Mid, Right, BackColor
    'Public FileName, color, token, A, R, G, B, commonDialog, value, Val, X, Y, Z, Left, hWnd, Mid, Right, BackColor
#End If


Public Type TColor
    red     As Byte
    green   As Byte
    blue    As Byte
End Type


Public Function GetRGB(DecValue As Long) As TColor

    Dim hexValue As String

    hexValue = Hex$(Val(DecValue))

    If Len(hexValue) < 6 Then
        hexValue = String$(6 - Len(hexValue), "0") + hexValue
    End If

    GetRGB.blue = CLng("&H" + Right$(hexValue, 2))
    hexValue = Left$(hexValue, Len(hexValue) - 2)
    GetRGB.green = CLng("&H" + Right$(hexValue, 2))
    hexValue = Left$(hexValue, Len(hexValue) - 2)
    GetRGB.red = CLng("&H" + Right$(hexValue, 2))

End Function

Public Function GetAlpha(tehColor As Long) As Byte

    Dim hexValue As String

    hexValue = Hex$(Val(tehColor))

    If Len(hexValue) <= 6 Then
        GetAlpha = 0
    Else
        If Len(hexValue) < 8 Then
            hexValue = String$(8 - Len(hexValue), "0") + hexValue
        End If
        GetAlpha = CLng("&H" + Left$(hexValue, 2))
    End If

End Function

Public Function ARGB(ByVal alphaVal As Byte, clrVal As Long) As Long

    Dim clrString As String

    clrString = Hex$(clrVal)

    If Len(clrString) < 6 Then
        clrString = String$(6 - Len(clrString), "0") & clrString
    ElseIf Len(clrString) > 6 Then
        clrString = Right$(clrString, 6)
    End If

    If Len(Hex$(alphaVal)) = 1 Then
        clrString = "0" + Hex$(alphaVal) & clrString
    ElseIf Len(Hex$(alphaVal)) = 2 Then
        clrString = Hex$(alphaVal) & clrString
    End If

    ARGB = CLng("&H" & clrString)

End Function

Public Function MakeColor(red As Byte, green As Byte, blue As Byte) As TColor

    MakeColor.red = red
    MakeColor.green = green
    MakeColor.blue = blue

End Function


Public Function DiffVal(val1 As Byte, val2 As Byte) As Byte

    If val1 > val2 Then
        DiffVal = val1 - val2
    Else
        DiffVal = val2 - val1
    End If

End Function

Public Function LowerVal(val1 As Byte, val2 As Byte) As Byte

    If val1 < val2 Then
        LowerVal = val1
    Else
        LowerVal = val2
    End If

End Function

Public Function HigherVal(val1 As Byte, val2 As Byte) As Byte

    If val1 > val2 Then
        HigherVal = val1
    Else
        HigherVal = val2
    End If

End Function

Public Function FileExists(theFileName As String) As Boolean

    On Error GoTo ErrorHandler

    FileExists = FileLen(theFileName) > 0

ErrorHandler:

End Function

Public Function Clamp(value As Single, min As Single, max As Single) As Single

    If value < min Then
        Clamp = min
    ElseIf value > max Then
        Clamp = max
    Else
        Clamp = value
    End If

End Function

Public Sub SetFormFonts(theForm As Form)
    Dim c As Control

    For Each c In theForm.Controls
        If c.Tag = "font1" Then
            c.Font.Name = font1
        ElseIf c.Tag = "font2" Then
            c.Font.Name = font2
        End If
    Next
End Sub

Public Function GetAngle(ByVal xVal As Single, ByVal yVal As Single) As Single

    If xVal < 0 Then
        GetAngle = PI - Atn(yVal / xVal)
    ElseIf xVal > 0 Then
        If Atn(yVal / xVal) > 0 Then
            GetAngle = 2 * PI - Atn(yVal / xVal)
        Else
            GetAngle = -Atn(yVal / xVal)
        End If
    Else
        If yVal > 0 Then
            GetAngle = 3 * PI / 2
        Else
            GetAngle = PI / 2
        End If
    End If

End Function

Public Function Midpoint(ByVal p1 As Single, ByVal p2 As Single) As Single

    If p1 < p2 Then
        Midpoint = p1 + (p2 - p1) / 2
    Else
        Midpoint = p2 + (p1 - p2) / 2
    End If

End Function

Public Function IsBetween(p1, p2, p3) As Boolean

    IsBetween = False

    If (p1 >= p2 And p2 >= p3) Or (p3 >= p2 And p2 >= p1) Then
        IsBetween = True
    End If

End Function

Public Function DirExists(DirName As String) As Boolean

    On Error GoTo ErrorHandler
    DirExists = GetAttr(DirName) And vbDirectory

ErrorHandler:

End Function
