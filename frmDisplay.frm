VERSION 5.00
Begin VB.Form frmDisplay 
   Appearance      =   0  'Flat
   BackColor       =   &H004A3C31&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2400
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3120
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   160
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   208
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picLayer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   10
      Left            =   1800
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   22
      Tag             =   "4"
      Top             =   840
      Width           =   240
   End
   Begin VB.PictureBox picLayer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   9
      Left            =   1800
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   20
      Tag             =   "4"
      Top             =   600
      Width           =   240
   End
   Begin VB.PictureBox picLayer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   8
      Left            =   1800
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   17
      Tag             =   "4"
      Top             =   360
      Width           =   240
   End
   Begin VB.PictureBox picLayer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   5
      Left            =   120
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   15
      Tag             =   "4"
      Top             =   840
      Width           =   240
   End
   Begin VB.PictureBox picLayer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   120
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   11
      Tag             =   "4"
      Top             =   2040
      Width           =   240
   End
   Begin VB.PictureBox picLayer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   120
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   10
      Tag             =   "4"
      Top             =   1800
      Width           =   240
   End
   Begin VB.PictureBox picLayer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   120
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   9
      Tag             =   "4"
      Top             =   1560
      Width           =   240
   End
   Begin VB.PictureBox picLayer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   3
      Left            =   120
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   8
      Tag             =   "4"
      Top             =   1320
      Width           =   240
   End
   Begin VB.PictureBox picLayer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   4
      Left            =   120
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   7
      Tag             =   "4"
      Top             =   1080
      Width           =   240
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
      ScaleWidth      =   208
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   3120
      Begin VB.PictureBox picHide 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H004A3C31&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   2880
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
   Begin VB.PictureBox picLayer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   6
      Left            =   120
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   16
      Tag             =   "4"
      Top             =   600
      Width           =   240
   End
   Begin VB.PictureBox picLayer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   7
      Left            =   120
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   18
      Tag             =   "4"
      Top             =   360
      Width           =   240
   End
   Begin VB.Label lblLayer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " Sketch"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   2040
      TabIndex        =   23
      Tag             =   "font2"
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblLayer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " Lights"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   2040
      TabIndex        =   21
      Tag             =   "font2"
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblLayer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " Waypoints"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   19
      Tag             =   "font2"
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblLayer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " Scenery"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   14
      Tag             =   "font2"
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblLayer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " Objects"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   13
      Tag             =   "font2"
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblLayer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " Grid"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   2040
      TabIndex        =   12
      Tag             =   "font2"
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lblLayer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " Background"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Tag             =   "font2"
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label lblLayer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " Polygons"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Tag             =   "font2"
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label lblLayer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " Texture"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Tag             =   "font2"
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblLayer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " Wireframe"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Tag             =   "font2"
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblLayer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " Points"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   2
      Tag             =   "font2"
      Top             =   1080
      Width           =   1455
   End
End
Attribute VB_Name = "frmDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Fix vb6 ide casing changes
#If False Then
    Public FileName, color, token, A, R, G, B, commonDialog, value, Val, X, Y, Z, Left, hWnd, Mid, Right, BackColor
    'Public FileName, color, token, A, R, G, B, commonDialog, value, Val, X, Y, Z, Left, hWnd, Mid, Right, BackColor
#End If


Public xPos As Integer
Public yPos As Integer
Public collapsed As Boolean


Private Const LAYER_BG As Byte = 0
Private Const LAYER_POLYS As Byte = 1
Private Const LAYER_TEXTURE As Byte = 2
Private Const LAYER_WIREFRAME As Byte = 3
Private Const LAYER_POINTS As Byte = 4
Private Const LAYER_SCENERY As Byte = 5
Private Const LAYER_OBJECTS As Byte = 6
Private Const LAYER_WAYPOINTS As Byte = 7
Private Const LAYER_GRID As Byte = 8

Private layers(0 To 10) As Boolean
Private layerKeys(0 To 7) As Byte

Private formHeight As Integer


Public Function GetLayerKey(ByVal Index As Byte) As Byte

    GetLayerKey = layerKeys(Index)

End Function

Public Sub setLayerKey(Index As Integer, ByVal value As Byte)

    If value > 0 Then
        layerKeys(Index) = value
    End If

End Sub

Private Sub Form_GotFocus()

    Beep

End Sub

Private Sub Form_Load()

    Dim i As Integer

    On Error GoTo ErrorHandler

    Me.SetColors
    formHeight = Me.ScaleHeight
    setForm

    Exit Sub

ErrorHandler:

    MsgBox Error$ & vbNewLine & "Error loading Display form"

End Sub

Public Sub setForm()

    Me.Left = xPos * Screen.TwipsPerPixelX
    Me.Top = yPos * Screen.TwipsPerPixelY

    If collapsed Then
        Me.Height = 19 * Screen.TwipsPerPixelY
    Else
        Me.Height = formHeight * Screen.TwipsPerPixelY
    End If

End Sub

Public Sub setLayer(Index As Integer, value As Boolean)

    layers(Index) = value
    MouseEvent2 picLayer(Index), 0, 0, BUTTON_SMALL, layers(Index), BUTTON_UP

End Sub

Private Sub lblLayer_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    picLayer_MouseMove Index, Button, 0, 0, 0

End Sub

Public Sub picLayer_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picLayer(Index), X, Y, BUTTON_SMALL, layers(Index), BUTTON_DOWN

End Sub

Private Sub picLayer_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picLayer(Index), X, Y, BUTTON_SMALL, layers(Index), BUTTON_MOVE, lblLayer(Index).Width + 16

End Sub

Public Sub picLayer_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    layers(Index) = Not layers(Index)
    frmSoldatMapEditor.SetDispOptions Index, layers(Index)
    MouseEvent2 frmDisplay.picLayer(Index), 0, 0, BUTTON_SMALL, layers(Index), BUTTON_UP

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

    SnapForm Me, frmPalette
    SnapForm Me, frmWaypoints
    SnapForm Me, frmTools
    SnapForm Me, frmScenery
    SnapForm Me, frmInfo
    SnapForm Me, frmTexture
    Me.Tag = SnapForm(Me, frmSoldatMapEditor)

    xPos = Me.Left / Screen.TwipsPerPixelX
    yPos = Me.Top / Screen.TwipsPerPixelY

End Sub

Private Sub picHide_Click()

    Me.Hide
    frmSoldatMapEditor.mnuDisplay.Checked = False

End Sub

Private Sub picHide_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picHide, X, Y, BUTTON_SMALL, 0, BUTTON_DOWN

End Sub

Private Sub picHide_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picHide, X, Y, BUTTON_SMALL, 0, BUTTON_MOVE

End Sub

Private Sub picHide_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picHide, X, Y, BUTTON_SMALL, 0, BUTTON_UP

End Sub

Public Sub RefreshButtons()

    Dim i As Integer

    Debug.Assert picLayer.LBound = LBound(layers)
    Debug.Assert picLayer.UBound = UBound(layers)

    For i = picLayer.LBound To picLayer.UBound
        MouseEvent2 picLayer(i), 0, 0, BUTTON_SMALL, layers(i), BUTTON_UP
    Next

End Sub

Public Sub SetColors()

    On Error Resume Next

    Dim i As Integer
    Dim c As Control

    picTitle.Picture = LoadPicture(appPath & "\skins\" & gfxDir & "\titlebar_display.bmp")
    MouseEvent2 picHide, 0, 0, BUTTON_SMALL, 0, BUTTON_UP

    Me.BackColor = bgColor

    For Each c In lblLayer
        c.BackColor = lblBackColor
        c.ForeColor = lblTextColor
    Next

    SetFormFonts Me

End Sub
