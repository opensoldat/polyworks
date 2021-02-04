VERSION 5.00
Begin VB.Form frmTools 
   Appearance      =   0  'Flat
   BackColor       =   &H004A3C31&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   960
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   64
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picTools 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   13
      Left            =   480
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   15
      Tag             =   "Depth Map"
      ToolTipText     =   "(.)"
      Top             =   3120
      Width           =   480
   End
   Begin VB.PictureBox picTools 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   12
      Left            =   0
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   14
      Tag             =   "Lights"
      ToolTipText     =   "(.)"
      Top             =   3120
      Width           =   480
   End
   Begin VB.PictureBox picTools 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   11
      Left            =   480
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   13
      Tag             =   "Sketch"
      ToolTipText     =   "(.)"
      Top             =   2640
      Width           =   480
   End
   Begin VB.PictureBox picTools 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   10
      Left            =   0
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   12
      Tag             =   "Color Picker"
      ToolTipText     =   "Color Picker (,)"
      Top             =   2640
      Width           =   480
   End
   Begin VB.PictureBox picTools 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   9
      Left            =   480
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   11
      Tag             =   "Objects"
      ToolTipText     =   "Objects (T)"
      Top             =   2160
      Width           =   480
   End
   Begin VB.PictureBox picTools 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   8
      Left            =   0
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   10
      Tag             =   "Waypoints"
      ToolTipText     =   "Waypoints (T)"
      Top             =   2160
      Width           =   480
   End
   Begin VB.PictureBox picTools 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   7
      Left            =   480
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   9
      Tag             =   "Scenery"
      ToolTipText     =   "Scenery (Y)"
      Top             =   1680
      Width           =   480
   End
   Begin VB.PictureBox picTools 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   0
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   8
      Tag             =   "Transform"
      ToolTipText     =   "Transform (M)"
      Top             =   240
      Width           =   480
   End
   Begin VB.PictureBox picTools 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   1
      Left            =   480
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   7
      Tag             =   "Poly Creation"
      ToolTipText     =   "Poly Creation (C)"
      Top             =   240
      Width           =   480
   End
   Begin VB.PictureBox picTools 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   2
      Left            =   0
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   6
      Tag             =   "Vertex Selection"
      ToolTipText     =   "Vertex Selection (V)"
      Top             =   720
      Width           =   480
   End
   Begin VB.PictureBox picTools 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   3
      Left            =   480
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   5
      Tag             =   "Poly Selection"
      ToolTipText     =   "Poly Selection (P)"
      Top             =   720
      Width           =   480
   End
   Begin VB.PictureBox picTools 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   4
      Left            =   0
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   4
      Tag             =   "Vertex Color"
      ToolTipText     =   "Vertex Color (E)"
      Top             =   1200
      Width           =   480
   End
   Begin VB.PictureBox picTools 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   5
      Left            =   480
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   3
      Tag             =   "Poly Color"
      ToolTipText     =   "Poly Color (R)"
      Top             =   1200
      Width           =   480
   End
   Begin VB.PictureBox picTools 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   6
      Left            =   0
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   2
      Tag             =   "Texture"
      ToolTipText     =   "Texture (T)"
      Top             =   1680
      Width           =   480
   End
   Begin VB.PictureBox picTitle 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   960
      Begin VB.PictureBox picHide 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H004A3C31&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   720
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "3"
         Top             =   0
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim curTool As Byte
Dim curButton As Byte
Public xPos As Integer, yPos  As Integer
Dim formHeight As Integer
Public collapsed As Boolean
Dim hotKeys(0 To 13) As Byte

Public Function getHotKey(ByVal Index As Byte) As Byte

    getHotKey = hotKeys(Index)

End Function

Public Sub setHotKey(Index As Integer, ByVal value As Byte)

    If value > 0 Then
        hotKeys(Index) = value
    End If

End Sub

Public Sub initTool(value As Byte)

    curTool = value

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    MsgBox KeyCode

End Sub

Private Sub Form_Load()

    On Error GoTo ErrorHandler

    SetColors
    
    formHeight = Me.ScaleHeight

    setForm

    Exit Sub

ErrorHandler:

    MsgBox Error$ & vbNewLine & "Error loading Tools form"

End Sub

Public Sub setForm()

    Me.left = xPos * Screen.TwipsPerPixelX
    Me.Top = yPos * Screen.TwipsPerPixelY
    If collapsed Then
        Me.Height = 19 * Screen.TwipsPerPixelY
    Else
        Me.Height = formHeight * Screen.TwipsPerPixelY
    End If

End Sub

Private Sub picTitle_DblClick()

    collapsed = Not collapsed
    If collapsed Then
        Me.Height = 19 * Screen.TwipsPerPixelY
    Else
        Me.Height = formHeight * Screen.TwipsPerPixelY
    End If

End Sub

Private Sub picTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ReleaseCapture
    SendMessage Me.hWnd, WM_NCLBUTTONDOWN, 2, 0&

    snapForm Me, frmPalette
    snapForm Me, frmWaypoints
    snapForm Me, frmDisplay
    snapForm Me, frmScenery
    snapForm Me, frmInfo
    snapForm Me, frmTexture
    Me.Tag = snapForm(Me, frmSoldatMapEditor)

    xPos = Me.left / Screen.TwipsPerPixelX
    yPos = Me.Top / Screen.TwipsPerPixelY

End Sub

Public Sub picTools_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim i As Integer

    If curTool <> Index Then
        For i = 0 To 13
            BitBlt picTools(i).hDC, 0, 0, 32, 32, frmSoldatMapEditor.picGfx.hDC, 0, i * 32, vbSrcCopy
            picTools(i).Refresh
        Next
        BitBlt picTools(Index).hDC, 0, 0, 32, 32, frmSoldatMapEditor.picGfx.hDC, 64, Index * 32, vbSrcCopy
        picTools(Index).Refresh
    End If
    curTool = Index

End Sub

Private Sub picTools_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If curTool <> Index Then
        mouseEvent picTools(Index), (picTools(Index).ScaleWidth - X), (picTools(Index).ScaleHeight - Y), 0, Index * 32, 32, 32
    End If

End Sub

Private Sub picTools_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    frmSoldatMapEditor.setCurrentTool curTool
    frmSoldatMapEditor.MouseIcon = frmSoldatMapEditor.ImageList.ListImages(curTool + 1).Picture
    frmSoldatMapEditor.RegainFocus

End Sub

Private Sub picHide_Click()

    Me.Hide
    frmSoldatMapEditor.mnuTools.Checked = False

End Sub

Private Sub picHide_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picHide, X, Y, BUTTON_SMALL, 0, BUTTON_DOWN

End Sub

Private Sub picHide_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picHide, X, Y, BUTTON_SMALL, 0, BUTTON_MOVE

End Sub

Private Sub picHide_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picHide, X, Y, BUTTON_SMALL, 0, BUTTON_UP

End Sub

Public Sub SetColors()

    On Error Resume Next

    Dim i As Integer

    picTitle.Picture = LoadPicture(appPath & "\" & gfxDir & "\titlebar_tools.bmp")

    mouseEvent2 picHide, 0, 0, BUTTON_SMALL, 0, BUTTON_UP

    For i = 0 To 13
        BitBlt picTools(i).hDC, 0, 0, 32, 32, frmSoldatMapEditor.picGfx.hDC, 0, i * 32, vbSrcCopy
        picTools(i).Refresh
        frmTools.picTools(i).ToolTipText = frmTools.picTools(i).Tag & " (" & Chr$(MapVirtualKey(hotKeys(i), 1)) & ")"
    Next
    BitBlt picTools(curTool).hDC, 0, 0, 32, 32, frmSoldatMapEditor.picGfx.hDC, 64, curTool * 32, vbSrcCopy
    picTools(curTool).Refresh

End Sub
