VERSION 5.00
Begin VB.Form frmWaypoints 
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
   Begin VB.PictureBox picShow 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   1920
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   21
      Tag             =   "6"
      Top             =   1920
      Width           =   240
   End
   Begin VB.PictureBox picShow 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   1920
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   19
      Tag             =   "6"
      Top             =   1680
      Width           =   240
   End
   Begin VB.PictureBox picShow 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   1920
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   17
      Tag             =   "6"
      Top             =   1440
      Width           =   240
   End
   Begin VB.ComboBox cboSpecial 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      ItemData        =   "frmWaypoints.frx":0000
      Left            =   120
      List            =   "frmWaypoints.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Tag             =   "font1"
      ToolTipText     =   "Special"
      Top             =   1200
      Width           =   1455
   End
   Begin VB.PictureBox picType 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   4
      Left            =   1920
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   8
      Tag             =   "4"
      Top             =   840
      Width           =   240
   End
   Begin VB.PictureBox picType 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   3
      Left            =   720
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   7
      Tag             =   "4"
      Top             =   840
      Width           =   240
   End
   Begin VB.PictureBox picType 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   720
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   6
      Tag             =   "4"
      Top             =   360
      Width           =   240
   End
   Begin VB.PictureBox picType 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   1320
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   5
      Tag             =   "4"
      Top             =   600
      Width           =   240
   End
   Begin VB.PictureBox picType 
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
      TabIndex        =   4
      Tag             =   "4"
      Top             =   600
      Width           =   240
   End
   Begin VB.PictureBox picPath 
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
      TabIndex        =   3
      Tag             =   "6"
      Top             =   1920
      Width           =   240
   End
   Begin VB.PictureBox picPath 
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
      TabIndex        =   2
      Tag             =   "6"
      Top             =   1680
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
   Begin VB.Label lblNumCon 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H007B614A&
      Height          =   255
      Left            =   2520
      TabIndex        =   24
      Top             =   360
      Width           =   375
   End
   Begin VB.Label lblWaypoints 
      BackStyle       =   0  'Transparent
      Caption         =   "Show:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   23
      Tag             =   "font2"
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblShow 
      BackStyle       =   0  'Transparent
      Caption         =   " Path2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   22
      Tag             =   "font2"
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label lblShow 
      BackStyle       =   0  'Transparent
      Caption         =   " Path1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   2160
      TabIndex        =   20
      Tag             =   "font2"
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label lblShow 
      BackStyle       =   0  'Transparent
      Caption         =   " All"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   2160
      TabIndex        =   18
      Tag             =   "font2"
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lblPath 
      BackStyle       =   0  'Transparent
      Caption         =   " Path 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   16
      Tag             =   "font2"
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lblPath 
      BackStyle       =   0  'Transparent
      Caption         =   " Path 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   15
      Tag             =   "font2"
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      Caption         =   " Fly"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   2160
      TabIndex        =   14
      Tag             =   "font2"
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      Caption         =   " Left"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   13
      Tag             =   "font2"
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      Caption         =   " Down"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   960
      TabIndex        =   12
      Tag             =   "font2"
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      Caption         =   " Right"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   11
      Tag             =   "font2"
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      Caption         =   " Up"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   10
      Tag             =   "font2"
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "frmWaypoints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim formHeight As Integer
Public collapsed As Boolean
Const COLLAPSED_HEIGHT = 19

Public xPos As Integer, yPos  As Integer

Dim wayptType(0 To 4) As Boolean
Public wayptPath As Byte
Public showPaths As Byte

Dim wayptKeys(0 To 4) As Byte

Public noChange As Boolean

Public Function getWayptKey(ByVal Index As Byte) As Byte

    getWayptKey = wayptKeys(Index)

End Function

Public Sub setWayptKey(Index As Integer, ByVal value As Byte)

    If value > 0 Then
        wayptKeys(Index) = value
    End If

End Sub

Private Sub Form_Load()

    Dim i As Integer

    On Error GoTo ErrorHandler

    Me.SetColors

    formHeight = Me.ScaleHeight

    setForm

    Exit Sub

ErrorHandler:

    MsgBox Error$ & vbNewLine & "Error loading Waypoints form"

End Sub

Public Sub setForm()

    Me.left = xPos * Screen.TwipsPerPixelX
    Me.Top = yPos * Screen.TwipsPerPixelY
    If collapsed Then
        Me.Height = COLLAPSED_HEIGHT * Screen.TwipsPerPixelY
    Else
        Me.Height = formHeight * Screen.TwipsPerPixelY
    End If

End Sub

Private Sub cboSpecial_Click()

    If noChange = False And cboSpecial.ListIndex > -1 Then
        If Not frmSoldatMapEditor.setSpecial(cboSpecial.ListIndex) Then
            cboSpecial.ListIndex = -1
        End If
    End If

End Sub

Public Sub getPathNum(tehValue As Byte)

    mouseEvent2 picPath(0), 0, 0, BUTTON_SMALL, tehValue = 1, BUTTON_UP
    mouseEvent2 picPath(1), 0, 0, BUTTON_SMALL, tehValue = 2, BUTTON_UP
    wayptPath = tehValue - 1

End Sub

Public Sub getWayType(Index As Integer, tehValue As Boolean)

    wayptType(Index) = tehValue
    mouseEvent2 picType(Index), 0, 0, BUTTON_SMALL, tehValue, BUTTON_UP

End Sub

Private Sub lblPath_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    picPath_MouseMove Index, 1, 0, 0, 0

End Sub

Private Sub lblShow_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    picShow_MouseMove Index, Button, 0, 0, 0

End Sub

Private Sub lblType_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    picType_MouseMove Index, Button, 0, 0, 0

End Sub

Public Sub ClearWaypt()

    Dim i As Integer

    For i = 0 To 4
        mouseEvent2 picType(i), 0, 0, BUTTON_SMALL, 0, BUTTON_UP
        wayptType(i) = False
    Next

    cboSpecial.ListIndex = -1
    lblNumCon.Caption = ""

End Sub


Private Sub picTitle_DblClick()

    collapsed = Not collapsed
    If collapsed Then
        Me.Height = COLLAPSED_HEIGHT * Screen.TwipsPerPixelY
    Else
        Me.Height = formHeight * Screen.TwipsPerPixelY
    End If

End Sub

Private Sub picTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ReleaseCapture
    SendMessage Me.hWnd, WM_NCLBUTTONDOWN, 2, 0&

    snapForm Me, frmPalette
    snapForm Me, frmInfo
    snapForm Me, frmTools
    snapForm Me, frmScenery
    snapForm Me, frmDisplay
    snapForm Me, frmTexture
    Me.Tag = snapForm(Me, frmSoldatMapEditor)

    xPos = Me.left / Screen.TwipsPerPixelX
    yPos = Me.Top / Screen.TwipsPerPixelY

End Sub

Private Sub picHide_Click()

    Me.Hide
    frmSoldatMapEditor.mnuWaypoints.Checked = False

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

Private Sub picPath_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picPath(Index), X, Y, BUTTON_SMALL, (Index = wayptPath), BUTTON_DOWN

End Sub

Private Sub picPath_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picPath(Index), X, Y, BUTTON_SMALL, (Index = wayptPath), BUTTON_MOVE, lblPath(Index).Width + 16

End Sub

Private Sub picPath_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim i As Integer

    wayptPath = Index

    For i = 0 To 1
        If i <> Index Then
            mouseEvent2 picPath(i), X, Y, BUTTON_SMALL, (i = wayptPath), BUTTON_UP
        End If
    Next

    frmSoldatMapEditor.setPathNum Index + 1

End Sub

Public Sub picType_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picType(Index), X, Y, BUTTON_SMALL, wayptType(Index), BUTTON_DOWN

End Sub

Private Sub picType_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picType(Index), X, Y, BUTTON_SMALL, wayptType(Index), BUTTON_MOVE, lblType(Index).Width + 16

End Sub

Public Sub picType_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Not frmSoldatMapEditor.setWayType(Index, Not wayptType(Index)) Then Exit Sub

    wayptType(Index) = Not wayptType(Index)
    mouseEvent2 picType(Index), 0, 0, BUTTON_SMALL, wayptType(Index), BUTTON_UP
    If Index = 0 Then
        wayptType(1) = False
        mouseEvent2 picType(1), 0, 0, BUTTON_SMALL, 0, BUTTON_UP
    ElseIf Index = 1 Then
        wayptType(0) = False
        mouseEvent2 picType(0), 0, 0, BUTTON_SMALL, 0, BUTTON_UP
    ElseIf Index = 2 Then
        wayptType(3) = False
        mouseEvent2 picType(3), 0, 0, BUTTON_SMALL, 0, BUTTON_UP
    ElseIf Index = 3 Then
        wayptType(2) = False
        mouseEvent2 picType(2), 0, 0, BUTTON_SMALL, 0, BUTTON_UP
    End If

End Sub

Private Sub picShow_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picShow(Index), X, Y, BUTTON_SMALL, (Index = showPaths), BUTTON_DOWN

End Sub

Private Sub picShow_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picShow(Index), X, Y, BUTTON_SMALL, (Index = showPaths), BUTTON_MOVE, lblShow(Index).Width + 16

End Sub

Public Sub picShow_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim i As Integer

    showPaths = Index

    For i = 0 To 2
        If i <> Index Then
            mouseEvent2 picShow(i), X, Y, BUTTON_SMALL, (i = showPaths), BUTTON_UP
        End If
    Next

    frmSoldatMapEditor.setShowPaths

End Sub

Public Sub SetColors()

    On Error Resume Next

    Dim i As Integer
    Dim c As Control


    picTitle.Picture = LoadPicture(appPath & "\" & gfxDir & "\titlebar_waypoints.bmp")
    mouseEvent2 picHide, 0, 0, BUTTON_SMALL, 0, BUTTON_UP

    mouseEvent2 picPath(0), 0, 0, BUTTON_SMALL, True, BUTTON_UP
    mouseEvent2 picPath(1), 0, 0, BUTTON_SMALL, False, BUTTON_UP

    For i = 0 To 4
        mouseEvent2 picType(i), 0, 0, BUTTON_SMALL, 0, BUTTON_UP
    Next

    For i = 0 To 2
        mouseEvent2 picShow(i), 0, 0, BUTTON_SMALL, i = showPaths, BUTTON_UP
    Next


    Me.BackColor = bgClr

    For i = 0 To 4
        lblType(i).BackColor = lblBackClr
        lblType(i).ForeColor = lblTextClr
    Next

    For i = 0 To 1
        lblPath(i).BackColor = lblBackClr
        lblPath(i).ForeColor = lblTextClr
    Next

    For i = 0 To 2
        lblShow(i).BackColor = lblBackClr
        lblShow(i).ForeColor = lblTextClr
    Next

    lblWaypoints.BackColor = lblBackClr
    lblWaypoints.ForeColor = lblTextClr

    cboSpecial.BackColor = txtBackClr
    cboSpecial.ForeColor = txtTextClr

    lblNumCon.BackColor = lblBackClr
    lblNumCon.ForeColor = lblTextClr

    For Each c In Me.Controls
        If c.Tag = "font1" Then
            c.Font.Name = font1
        ElseIf c.Tag = "font2" Then
            c.Font.Name = font2
        End If
    Next

End Sub
