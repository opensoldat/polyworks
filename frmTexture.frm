VERSION 5.00
Begin VB.Form frmTexture 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H004A3C31&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4320
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   1200
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   288
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   80
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTexture 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DrawStyle       =   2  'Dot
      DrawWidth       =   2
      ForeColor       =   &H00FF0000&
      Height          =   3840
      Left            =   120
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   2
      Top             =   360
      Width           =   960
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
      ScaleWidth      =   80
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   1200
      Begin VB.PictureBox picHide 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H004A3C31&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   960
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
Attribute VB_Name = "frmTexture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


' Fix vb6 ide casing changes
#If False Then
    Private Token
    'Private Token
#End If


Public xPos As Integer
Public yPos  As Integer
Public collapsed As Boolean
Public x1tex As Single
Public x2tex As Single
Public y1tex As Single
Public y2tex As Single


Private formHeight As Integer


Private Sub Form_Load()

    On Error GoTo ErrorHandler

    Me.SetColors
    formHeight = Me.ScaleHeight
    setForm

    Exit Sub

ErrorHandler:

    MsgBox Error$ & vbNewLine & "Error loading texture form"

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

Public Sub setTexCoords(tehValue As Single, Index As Integer)

    picTexture.Line (x1tex, y1tex)-(x2tex, y2tex), RGB(255, 255, 255), B
    If Index = 0 Then
        x1tex = tehValue / 2
    ElseIf Index = 1 Then
        x2tex = tehValue / 2
    ElseIf Index = 2 Then
        y1tex = tehValue / 2
    ElseIf Index = 3 Then
        y2tex = tehValue / 2
    End If
    picTexture.Line (x1tex, y1tex)-(x2tex, y2tex), RGB(255, 255, 255), B

End Sub

Private Sub picTexture_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button <> 0 Then
        picTexture.DrawMode = 6
        picTexture.Line (x1tex, y1tex)-(x2tex, y2tex), RGB(255, 255, 255), B
        x1tex = Int((X + 0) / 16) * 16
        y1tex = Int((Y + 0) / 16) * 16
        x2tex = x1tex + 16
        y2tex = y1tex + 16
        picTexture.Line (x1tex, y1tex)-(x2tex, y2tex), RGB(255, 255, 255), B
    End If

End Sub

Private Sub picTexture_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim drawBox As Boolean

    If Button <> 0 Then
        If X + 16 > frmSoldatMapEditor.xTexture / 2 Then
            X = frmSoldatMapEditor.xTexture / 2 - 16
        ElseIf X + 16 < 0 Then
            X = -16
        End If
        If Y + 16 > frmSoldatMapEditor.yTexture / 2 Then
            Y = frmSoldatMapEditor.yTexture / 2 - 16
        ElseIf Y + 16 < 0 Then
            Y = -16
        End If
        If Int((X + 16) / 16) * 16 <> x2tex Then
            drawBox = True
        End If
        If Int((Y + 16) / 16) * 16 <> y2tex Then
            drawBox = True
        End If
        If drawBox Then
            picTexture.Line (x1tex, y1tex)-(x2tex, y2tex), RGB(255, 255, 255), B
            x2tex = Int((X + 16) / 16) * 16
            y2tex = Int((Y + 16) / 16) * 16
            picTexture.Line (x1tex, y1tex)-(x2tex, y2tex), RGB(255, 255, 255), B
        End If
    End If

End Sub

Private Sub picTexture_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button <> 0 Then
        If X + 16 > frmSoldatMapEditor.xTexture / 2 Then
            X = frmSoldatMapEditor.xTexture / 2 - 16
        ElseIf X + 16 < 0 Then
            X = -16
        End If
        If Y + 16 > frmSoldatMapEditor.yTexture / 2 Then
            Y = frmSoldatMapEditor.yTexture / 2 - 16
        ElseIf Y + 16 < 0 Then
            Y = -16
        End If

        x2tex = Int((X + 16) / 16) * 16
        y2tex = Int((Y + 16) / 16) * 16

        frmInfo.txtQuadX(0).Text = x1tex * 2
        frmInfo.txtQuadY(0).Text = y1tex * 2
        frmInfo.txtQuadX(1).Text = x2tex * 2
        frmInfo.txtQuadY(1).Text = y2tex * 2
    End If

End Sub

Public Sub setTexture(texturePath As String)

    On Error GoTo ErrorHandler

    Dim texWidth As Integer
    Dim texHeight As Integer
    Dim X As Integer
    Dim Y As Integer

    texWidth = frmSoldatMapEditor.xTexture
    texHeight = frmSoldatMapEditor.yTexture

    picTexture.DrawMode = 13

    picTexture.Width = texWidth / 2
    picTexture.Height = texHeight / 2
    frmTexture.Width = (texWidth / 2 + 2 + 16) * Screen.TwipsPerPixelX
    formHeight = texHeight / 2 + 18 + 16
    frmTexture.Height = formHeight * Screen.TwipsPerPixelY
    picHide.left = frmTexture.Width / Screen.TwipsPerPixelX - 17

    Dim Token As Long
    Token = InitGDIPlus
    picTexture.Picture = LoadPictureGDIPlus(frmSoldatMapEditor.soldatDir & "textures\" & texturePath, texWidth / 2, texHeight / 2)
    FreeGDIPlus Token

    For Y = 0 To (texHeight / 32)
        If Y Mod 4 = 0 Then
            picTexture.DrawWidth = 2
        Else
            picTexture.DrawWidth = 1
        End If
        picTexture.Line (0, Y * 16)-(texWidth / 2, Y * 16), RGB(0, 0, 0)
    Next

    For X = 0 To (texWidth / 32)
        If X Mod 4 = 0 Then
            picTexture.DrawWidth = 2
        Else
            picTexture.DrawWidth = 1
        End If
        picTexture.Line (X * 16, 0)-(X * 16, texHeight), RGB(0, 0, 0)
    Next

    x1tex = 0
    y1tex = 0
    x2tex = texWidth / 2
    y2tex = texHeight / 2
    picTexture.DrawMode = 6
    picTexture.Line (x1tex, y1tex)-(x2tex, y2tex), RGB(255, 255, 255), B

    Exit Sub

ErrorHandler:

    MsgBox "Error setting texture" & vbNewLine & Error$

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

    snapForm Me, frmTools
    snapForm Me, frmPalette
    snapForm Me, frmWaypoints
    snapForm Me, frmDisplay
    snapForm Me, frmScenery
    snapForm Me, frmInfo
    Me.Tag = snapForm(Me, frmSoldatMapEditor)

    xPos = Me.left / Screen.TwipsPerPixelX
    yPos = Me.Top / Screen.TwipsPerPixelY

End Sub

Private Sub picHide_Click()

    Me.Hide
    frmSoldatMapEditor.mnuTexture.Checked = False

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

    picTitle.Picture = LoadPicture(appPath & "\" & gfxDir & "\titlebar_texture.bmp")
    mouseEvent2 picHide, 0, 0, BUTTON_SMALL, 0, BUTTON_UP

    Me.BackColor = bgClr

End Sub
