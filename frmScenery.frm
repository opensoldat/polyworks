VERSION 5.00
Begin VB.Form frmScenery 
   Appearance      =   0  'Flat
   BackColor       =   &H004A3C31&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2550
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   3120
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   170
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   208
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picRotate 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1320
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   12
      Tag             =   "4"
      Top             =   1920
      Width           =   240
   End
   Begin VB.PictureBox picScale 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1320
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   11
      Tag             =   "4"
      Top             =   2160
      Width           =   240
   End
   Begin VB.PictureBox picLevel 
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
      TabIndex        =   10
      Tag             =   "6"
      Top             =   2160
      Width           =   240
   End
   Begin VB.PictureBox picLevel 
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
      TabIndex        =   6
      Tag             =   "6"
      Top             =   1920
      Width           =   240
   End
   Begin VB.PictureBox picLevel 
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
      TabIndex        =   5
      Tag             =   "6"
      Top             =   1680
      Width           =   240
   End
   Begin VB.PictureBox picScenery 
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   4
      Top             =   360
      Width           =   975
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
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   3120
      Begin VB.PictureBox picSceneryMenu 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H004A3C31&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   2640
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "8"
         Top             =   0
         Width           =   240
      End
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
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "3"
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.ListBox lstScenery 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   1320
      TabIndex        =   0
      Tag             =   "font1"
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label lblLvl 
      BackStyle       =   0  'Transparent
      Caption         =   "Level:"
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
      Left            =   120
      TabIndex        =   15
      Tag             =   "font2"
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label lblRotate 
      BackStyle       =   0  'Transparent
      Caption         =   "Rotate"
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
      Left            =   1680
      TabIndex        =   14
      Tag             =   "font2"
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label lblScale 
      BackStyle       =   0  'Transparent
      Caption         =   "Scale"
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
      Left            =   1680
      TabIndex        =   13
      Tag             =   "font2"
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lblLevel 
      BackStyle       =   0  'Transparent
      Caption         =   "Front"
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
      Left            =   480
      TabIndex        =   9
      Tag             =   "font2"
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblLevel 
      BackStyle       =   0  'Transparent
      Caption         =   "Middle"
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
      Left            =   480
      TabIndex        =   8
      Tag             =   "font2"
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label lblLevel 
      BackStyle       =   0  'Transparent
      Caption         =   "Back"
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
      Left            =   480
      TabIndex        =   7
      Tag             =   "font2"
      Top             =   1680
      Width           =   735
   End
   Begin VB.Image imgScenery 
      Appearance      =   0  'Flat
      Height          =   975
      Left            =   120
      Top             =   360
      Width           =   975
   End
   Begin VB.Menu mnuScenery 
      Caption         =   "Scenery"
      Visible         =   0   'False
      Begin VB.Menu mnuClearUnused 
         Caption         =   "Clear Unused"
      End
      Begin VB.Menu mnuReload 
         Caption         =   "Reload Scenery List"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh Scenery"
      End
   End
End
Attribute VB_Name = "frmScenery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' scenery dialog


' Fix vb6 ide casing changes
#If False Then
    Public FileName, color, token, A, R, G, B, commonDialog, value, Val, X, Y, Z, Left, hWnd, Mid, Right, BackColor
    'Public FileName, color, token, A, R, G, B, commonDialog, value, Val, X, Y, Z, Left, hWnd, Mid, Right, BackColor
#End If


' vars - public

Public xPos As Integer
Public yPos As Integer
Public collapsed As Boolean

Public level As Byte
Public rotateScenery As Boolean
Public scaleScenery As Boolean
Public notClicked As Boolean


' vars - private

Private formHeight As Integer
Private checkVal As Boolean
Private selNode As Node


' functions - public

Public Sub ListScenery()

    On Error GoTo ErrorHandler

    Dim file As Variant
    Dim Index As Integer
    Dim i As Integer
    Dim sceneryName As String
    Dim fileOpen As Boolean
    Dim tempNode As Node

    frmSoldatMapEditor.tvwScenery.Nodes.Clear

    frmSoldatMapEditor.tvwScenery.Nodes.Add , , "In Use", "In Use"

    ' load all scenery
    frmSoldatMapEditor.tvwScenery.Nodes.Add , , "Master List", "Master List"

    file = Dir(frmSoldatMapEditor.soldatDir & "Scenery-gfx\" & "*.bmp", vbDirectory)
    Do While Len(file)
        frmSoldatMapEditor.tvwScenery.Nodes.Add "Master List", tvwChild, , file
        file = Dir
    Loop

    file = Dir(frmSoldatMapEditor.soldatDir & "Scenery-gfx\" & "*.png", vbDirectory)
    Do While Len(file)
        frmSoldatMapEditor.tvwScenery.Nodes.Add "Master List", tvwChild, , file
        file = Dir
    Loop

    file = Dir(frmSoldatMapEditor.soldatDir & "Scenery-gfx\" & "*.tga", vbDirectory)
    Do While Len(file)
        frmSoldatMapEditor.tvwScenery.Nodes.Add "Master List", tvwChild, , file
        file = Dir
    Loop

    file = Dir(frmSoldatMapEditor.soldatDir & "Scenery-gfx\" & "*.gif", vbDirectory)
    Do While Len(file)
        frmSoldatMapEditor.tvwScenery.Nodes.Add "Master List", tvwChild, , file
        file = Dir
    Loop

    frmSoldatMapEditor.tvwScenery.Nodes("Master List").Sorted = True
    frmSoldatMapEditor.tvwScenery.Nodes("Master List").Sorted = False

    frmSoldatMapEditor.tvwScenery.Nodes("Master List").Child.selected = True
    frmSoldatMapEditor.tvwScenery_NodeClick frmSoldatMapEditor.tvwScenery.SelectedItem

    ' load lists

    file = Dir(appPath & "\lists\" & "*.txt", vbDirectory)
    Do While Len(file)  ' for every txt file in lists
        file = Left(file, Len(file) - 4)
        frmSoldatMapEditor.tvwScenery.Nodes.Add , , file, file
        fileOpen = True
        Open appPath & "\lists\" & file & ".txt" For Input As #1

            Do Until EOF(1)
                Input #1, sceneryName
                frmSoldatMapEditor.tvwScenery.Nodes.Add file, tvwChild, , sceneryName
            Loop

        Close #1

        fileOpen = False
        file = Dir
    Loop

    Exit Sub

ErrorHandler:

    MsgBox "Error loading scenery tree" & vbNewLine & Error & vbNewLine & sceneryName
    If fileOpen Then
        Close #1
    End If

End Sub

Public Sub SetForm()

    Me.Left = xPos * Screen.TwipsPerPixelX
    Me.Top = yPos * Screen.TwipsPerPixelY
    If collapsed Then
        Me.Height = 19 * Screen.TwipsPerPixelY
    Else
        Me.Height = formHeight * Screen.TwipsPerPixelY
    End If

End Sub

Public Sub SetColors()

    On Error Resume Next

    Dim i As Integer
    Dim c As Control


    picTitle.Picture = LoadPicture(appPath & "\skins\" & gfxDir & "\titlebar_scenery.bmp")

    MouseEvent2 picHide, 0, 0, BUTTON_SMALL, 0, BUTTON_UP
    MouseEvent2 picSceneryMenu, 0, 0, BUTTON_SMALL, 0, BUTTON_UP

    For i = picLevel.LBound To picLevel.UBound
        MouseEvent2 picLevel(i), 0, 0, BUTTON_SMALL, (i = level), BUTTON_UP
    Next

    MouseEvent2 picScale, 0, 0, BUTTON_SMALL, scaleScenery, BUTTON_UP
    MouseEvent2 picRotate, 0, 0, BUTTON_SMALL, rotateScenery, BUTTON_UP


    Me.BackColor = bgColor
    lblLvl.BackColor = lblBackColor
    lblLvl.ForeColor = lblTextColor
    For Each c In lblLevel
        c.BackColor = lblBackColor
        c.ForeColor = lblTextColor
    Next
    lblRotate.BackColor = lblBackColor
    lblRotate.ForeColor = lblTextColor
    lblScale.BackColor = lblBackColor
    lblScale.ForeColor = lblTextColor
    lstScenery.BackColor = txtBackColor
    lstScenery.ForeColor = txtTextColor
    picScenery.BackColor = bgColor

    SetFormFonts Me

End Sub


' functions - private


' events - public

Public Sub lstScenery_Click()

    Dim token As Long

    On Error GoTo ErrorHandler

    If lstScenery.List(lstScenery.ListIndex) = "" Then
        lstScenery.ListIndex = -1
        Exit Sub
    End If

    If Len(Dir(frmSoldatMapEditor.soldatDir & "Scenery-gfx\" & lstScenery.List(lstScenery.ListIndex))) <> 0 Then
        token = InitGDIPlus
        picScenery.Picture = LoadPictureGDIPlus(frmSoldatMapEditor.soldatDir & "Scenery-gfx\" & lstScenery.List(lstScenery.ListIndex), , , RGB(0, 255, 0))
        FreeGDIPlus token
        frmSoldatMapEditor.SetCurrentScenery lstScenery.ListIndex + 1, lstScenery.List(lstScenery.ListIndex)
    Else
        frmSoldatMapEditor.SetCurrentScenery lstScenery.ListIndex + 1, "notfound.bmp"
        picScenery.Picture = LoadPicture(appPath & "\skins\" & gfxDir & "\notfound.bmp")
        frmSoldatMapEditor.tvwScenery.SelectedItem = Nothing
    End If

    lstScenery.ToolTipText = lstScenery.List(lstScenery.ListIndex)
    frmSoldatMapEditor.tvwScenery.Nodes(lstScenery.List(lstScenery.ListIndex)).selected = True

    Exit Sub

ErrorHandler:

    MsgBox "Error clicking scenery" & vbNewLine & Error

End Sub

Public Sub picLevel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picLevel(Index), X, Y, BUTTON_SMALL, (Index = level), BUTTON_DOWN

End Sub


' events - private

Private Sub Form_Load()

    Dim i As Integer

    On Error GoTo ErrorHandler

    Me.SetColors
    formHeight = Me.ScaleHeight
    SetForm
    ListScenery

    Exit Sub

ErrorHandler:

    MsgBox "Error loading Scenery form" & vbNewLine & Error

End Sub

Private Sub lblLevel_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    picLevel_MouseMove Index, 1, 0, 0, 0

End Sub

Private Sub lblRotate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    picRotate_MouseMove 1, 0, 0, 0

End Sub

Private Sub lblScale_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    picScale_MouseMove 1, 0, 0, 0

End Sub

Private Sub mnuClearUnused_Click()

    frmSoldatMapEditor.ClearUnused

End Sub

Private Sub mnuReload_Click()

    Dim i As Integer

    ListScenery

    For i = 0 To lstScenery.ListCount - 1
        frmSoldatMapEditor.tvwScenery.Nodes.Add "In Use", tvwChild, lstScenery.List(i), lstScenery.List(i)
    Next

End Sub

Private Sub mnuRefresh_Click()

    Dim Index As Integer

    For Index = 1 To lstScenery.ListCount
        frmSoldatMapEditor.RefreshSceneryTextures Index
    Next
    frmSoldatMapEditor.Render

End Sub

Private Sub picSceneryMenu_Click()

    PopupMenu mnuScenery, , picHide.Left + picHide.ScaleWidth, picSceneryMenu.Top + picSceneryMenu.ScaleHeight

End Sub

Private Sub picSceneryMenu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picSceneryMenu, X, Y, BUTTON_SMALL, 0, BUTTON_DOWN

End Sub

Private Sub picSceneryMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picSceneryMenu, X, Y, BUTTON_SMALL, 0, BUTTON_MOVE

End Sub

Private Sub picSceneryMenu_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picSceneryMenu, X, Y, BUTTON_SMALL, 0, BUTTON_UP

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
    SnapForm Me, frmDisplay
    SnapForm Me, frmTools
    SnapForm Me, frmInfo
    SnapForm Me, frmTexture
    Me.Tag = SnapForm(Me, frmSoldatMapEditor)

    xPos = Me.Left / Screen.TwipsPerPixelX
    yPos = Me.Top / Screen.TwipsPerPixelY

End Sub

Private Sub picHide_Click()

    Me.Hide
    frmSoldatMapEditor.mnuScenery.Checked = False

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

Private Sub picRotate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picRotate, X, Y, BUTTON_SMALL, rotateScenery, BUTTON_DOWN

End Sub

Private Sub picRotate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picRotate, X, Y, BUTTON_SMALL, rotateScenery, BUTTON_MOVE, lblRotate.Width + 16

End Sub

Private Sub picRotate_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    rotateScenery = Not rotateScenery

End Sub

Private Sub picScale_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picScale, X, Y, BUTTON_SMALL, scaleScenery, BUTTON_DOWN

End Sub

Private Sub picScale_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picScale, X, Y, BUTTON_SMALL, scaleScenery, BUTTON_MOVE, lblScale.Width + 16

End Sub

Private Sub picScale_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    scaleScenery = Not scaleScenery

End Sub

Private Sub picLevel_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picLevel(Index), X, Y, BUTTON_SMALL, (Index = level), BUTTON_MOVE, lblLevel(Index).Width + 16

End Sub

Private Sub picLevel_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim i As Integer

    level = Index

    For i = picLevel.LBound To picLevel.UBound
        If i <> Index Then
            MouseEvent2 picLevel(i), X, Y, BUTTON_SMALL, (i = level), BUTTON_UP
        End If
    Next

    frmSoldatMapEditor.SetSceneryLevel level
    frmSoldatMapEditor.RegainFocus

End Sub
