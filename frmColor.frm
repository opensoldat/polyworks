VERSION 5.00
Begin VB.Form frmColor 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H004A3C31&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5640
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7080
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   376
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   472
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtHexCode 
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
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
      Height          =   285
      Left            =   4440
      TabIndex        =   27
      Tag             =   "font1"
      Top             =   5160
      Width           =   855
   End
   Begin VB.PictureBox picSpectrum 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      DrawMode        =   6  'Mask Pen Not
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   240
      MousePointer    =   99  'Custom
      ScaleHeight     =   255
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   255
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   480
      Width           =   3855
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   240
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   62
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   4440
      Width           =   960
   End
   Begin VB.TextBox txtRGB 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   3600
      MaxLength       =   3
      TabIndex        =   0
      Tag             =   "font1"
      Top             =   4440
      Width           =   480
   End
   Begin VB.TextBox txtRGB 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   3600
      MaxLength       =   3
      TabIndex        =   1
      Tag             =   "font1"
      Top             =   4800
      Width           =   480
   End
   Begin VB.TextBox txtRGB 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   3600
      MaxLength       =   3
      TabIndex        =   2
      Tag             =   "font1"
      Top             =   5160
      Width           =   480
   End
   Begin VB.TextBox txtHue 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      MaxLength       =   3
      TabIndex        =   3
      Tag             =   "font1"
      Top             =   4440
      Width           =   480
   End
   Begin VB.TextBox txtSat 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      MaxLength       =   3
      TabIndex        =   4
      Tag             =   "font1"
      Top             =   4800
      Width           =   480
   End
   Begin VB.TextBox txtBright 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      MaxLength       =   3
      TabIndex        =   5
      Tag             =   "font1"
      Top             =   5160
      Width           =   480
   End
   Begin VB.PictureBox picHue 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   5160
      ScaleHeight     =   255
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   480
      Width           =   255
   End
   Begin VB.PictureBox picSat 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   4680
      ScaleHeight     =   255
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   480
      Width           =   255
   End
   Begin VB.PictureBox picBright 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   4200
      ScaleHeight     =   255
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   480
      Width           =   255
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
      ScaleWidth      =   472
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   7080
      Begin VB.PictureBox picHide 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H004A3C31&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   6840
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   12
         TabStop         =   0   'False
         Tag             =   "3"
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.PictureBox picRGB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      ForeColor       =   &H80000008&
      Height          =   3855
      Index           =   2
      Left            =   6600
      ScaleHeight     =   255
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   480
      Width           =   255
   End
   Begin VB.PictureBox picRGB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      ForeColor       =   &H80000008&
      Height          =   3855
      Index           =   1
      Left            =   6120
      ScaleHeight     =   255
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   480
      Width           =   255
   End
   Begin VB.PictureBox picRGB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      ForeColor       =   &H80000008&
      Height          =   3855
      Index           =   0
      Left            =   5640
      ScaleHeight     =   255
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   480
      Width           =   255
   End
   Begin VB.PictureBox picCancel 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
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
      Height          =   360
      Left            =   5880
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   7
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   5040
      Width           =   960
   End
   Begin VB.PictureBox picOK 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
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
      Height          =   360
      Left            =   5880
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   6
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   4560
      Width           =   960
   End
   Begin VB.Label lblClr 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
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
      Left            =   2640
      TabIndex        =   26
      Tag             =   "font2"
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label lblClr 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
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
      Left            =   2640
      TabIndex        =   25
      Tag             =   "font2"
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label lblClr 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "°"
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
      Left            =   2640
      TabIndex        =   24
      Tag             =   "font2"
      Top             =   4440
      Width           =   135
   End
   Begin VB.Image imgBright 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   4440
      Top             =   480
      Width           =   225
   End
   Begin VB.Image imgSat 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   4920
      Top             =   480
      Width           =   225
   End
   Begin VB.Image imgHue 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   5400
      Top             =   480
      Width           =   225
   End
   Begin VB.Image imgRGB 
      Appearance      =   0  'Flat
      Height          =   225
      Index           =   2
      Left            =   6840
      Top             =   480
      Width           =   225
   End
   Begin VB.Image imgRGB 
      Appearance      =   0  'Flat
      Height          =   225
      Index           =   1
      Left            =   6360
      Top             =   480
      Width           =   225
   End
   Begin VB.Image imgRGB 
      Appearance      =   0  'Flat
      Height          =   225
      Index           =   0
      Left            =   5880
      Top             =   480
      Width           =   225
   End
   Begin VB.Label lblClr 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00614B3D&
      BackStyle       =   0  'Transparent
      Caption         =   "R"
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
      Left            =   3240
      TabIndex        =   21
      Tag             =   "font2"
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label lblClr 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00614B3D&
      BackStyle       =   0  'Transparent
      Caption         =   "G"
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
      Left            =   3240
      TabIndex        =   20
      Tag             =   "font2"
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label lblClr 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00614B3D&
      BackStyle       =   0  'Transparent
      Caption         =   "B"
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
      Left            =   3240
      TabIndex        =   19
      Tag             =   "font2"
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label lblClr 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00614B3D&
      BackStyle       =   0  'Transparent
      Caption         =   "H"
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
      Left            =   1800
      TabIndex        =   18
      Tag             =   "font2"
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label lblClr 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00614B3D&
      BackStyle       =   0  'Transparent
      Caption         =   "S"
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
      Left            =   1800
      TabIndex        =   17
      Tag             =   "font2"
      Top             =   4800
      Width           =   255
   End
   Begin VB.Label lblClr 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00614B3D&
      BackStyle       =   0  'Transparent
      Caption         =   "B"
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
      Left            =   1800
      TabIndex        =   16
      Tag             =   "font2"
      Top             =   5160
      Width           =   255
   End
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Fix vb6 ide casing changes
#If False Then
    Public FileName, color, token, A, R, G, B, commonDialog, value, Val, X, Y, Z, Left, hWnd, Mid
    'Public FileName, color, token, A, R, G, B, commonDialog, value, Val, X, Y, Z, Left, hWnd, Mid
#End If


Public red As Byte
Public green As Byte
Public blue As Byte

Public ok As Boolean


Private mHue As Single
Private mSat As Single
Private mBright As Single

Private mLow As Byte
Private mMid As Byte
Private mHigh As Byte

Private mColor(0 To 2) As Byte
Private mPureColor(0 To 2) As Byte

Private mOldX As Integer
Private mOldY As Integer

Private Const R As Byte = 0
Private Const G As Byte = 1
Private Const B As Byte = 2

Private mHexValue As String

Private mNonModal As Boolean

Private mLastTool As Byte


Public Sub InitColor(initRed As Byte, initGreen As Byte, initBlue As Byte)

    On Error GoTo ErrorHandler

    mColor(R) = initRed
    mColor(G) = initGreen
    mColor(B) = initBlue
    red = mColor(R)
    green = mColor(G)
    blue = mColor(B)

    changeRGB

    picSpectrum.Cls
    mOldX = (mHue / 360 * 256)
    mOldY = 255 - (mSat * 255)
    picSpectrum.Circle (mOldX, mOldY), 5.5, RGB(0, 0, 0)

    picColor.BackColor = RGB(mColor(R), mColor(B), mColor(G))

    updateAll
    updateRGB
    updateHSB
    updateHex

    Exit Sub

ErrorHandler:

    MsgBox "Error initializing color picker" & vbNewLine & Error$

End Sub

Public Sub ChangeColor(ByRef pic As PictureBox, ByRef rVal As Byte, ByRef gVal As Byte, ByRef bVal As Byte, ByVal cTool As Byte)

    mNonModal = True

    mLastTool = frmSoldatMapEditor.setTempTool(10)
    frmSoldatMapEditor.setCurrentTool 10

    frmSoldatMapEditor.picMenuBar.Enabled = False
    frmTools.Enabled = False
    frmPalette.Enabled = False
    frmScenery.Enabled = False
    frmInfo.Enabled = False
    frmWaypoints.Enabled = False
    frmDisplay.picTitle.Enabled = False

    Me.Show , frmSoldatMapEditor

End Sub


Private Sub HideColor(apply As Boolean)

    On Error GoTo ErrorHandler

    If mNonModal Then
        If apply Then
            frmPalette.setValues red, green, blue
            frmPalette.checkPalette red, green, blue
        End If

        mNonModal = False

        frmSoldatMapEditor.picMenuBar.Enabled = True

        frmTools.Enabled = True
        frmPalette.Enabled = True
        frmScenery.Enabled = True
        frmInfo.Enabled = True
        frmWaypoints.Enabled = True
        frmDisplay.picTitle.Enabled = True

        frmSoldatMapEditor.setCurrentTool mLastTool
        mLastTool = 0
    End If

    Me.Hide
    frmSoldatMapEditor.RegainFocus

    Exit Sub

ErrorHandler:

    MsgBox Error$

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    Const ESCAPE = 27
    Const ENTER = 13

    If KeyAscii = ESCAPE Then
        picColor.SetFocus
        picCancel_Click
    ElseIf KeyAscii = ENTER Then
        picColor.SetFocus
        picOK_Click
    End If

End Sub

Private Sub Form_Load()

    On Error GoTo ErrorHandler

    Me.SetColors

    mOldX = -16
    mOldY = -16
    ok = False
    mHue = 0
    mSat = 0
    mBright = 0
    mLow = B
    mMid = G
    mHigh = R
    mPureColor(0) = 255
    mPureColor(1) = 255
    mPureColor(2) = 255

    Exit Sub

ErrorHandler:

    MsgBox Error$ & vbNewLine & "Error loading Color Picker form"

End Sub

Private Sub lblRGB_Click(Index As Integer)

End Sub

Private Sub picSpectrum_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    picSpectrum_MouseMove Button, Shift, X, Y

End Sub

Private Sub picSpectrum_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        X = Clamp(X, 0, 255)
        Y = Clamp(Y, 0, 255)
        mSat = (255 - Y) / 255
        mHue = X / 255 * 359
        calculateHue
        changeRGB
        txtSat.Text = Int(mSat * 100 + 0.5)
        txtHue.Text = Int(mHue + 0.5)
        updateAll
        updateRGB
        updateHex

        picSpectrum.Circle (mOldX, mOldY), 5.5, RGB(0, 0, 0)
        mOldX = X
        mOldY = Y
        picSpectrum.Circle (X, Y), 5.5, RGB(0, 0, 0)
    End If

End Sub

Private Sub picRGB_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    picRGB_MouseMove Index, Button, Shift, X, Y

End Sub

Private Sub picRGB_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        X = 255 - Clamp(Y, 0, 255) ' grab y pos as it's a vertical bar
        mColor(Index) = X
        changeRGB
        txtRGB(Index).Text = mColor(Index)
        updateAll
        updateHSB
        updateHex

        picSpectrum.Circle (mOldX, mOldY), 5.5, RGB(0, 0, 0)
        mOldX = mHue / 360 * 255
        mOldY = 255 - mSat * 255
        picSpectrum.Circle (mOldX, mOldY), 5.5, RGB(0, 0, 0)
    End If

End Sub

Private Sub picHue_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    picHue_MouseMove Button, Shift, X, Y

End Sub

Private Sub picHue_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        X = 255 - Clamp(Y, 0, 255) ' grab y pos as it's a vertical bar
        mHue = X / 255 * 359

        calculateHue
        changeHue

        txtHue.Text = Int(mHue + 0.5)
        updateAll
        updateRGB
        updateHex

        picSpectrum.Circle (mOldX, mOldY), 5.5, RGB(0, 0, 0)
        mOldX = X
        picSpectrum.Circle (mOldX, mOldY), 5.5, RGB(0, 0, 0)
    End If

End Sub

Private Sub calculateHue()

    On Error GoTo ErrorHandler

    If mHue < 60 Then
        mColor(R) = mBright * 255
        mColor(G) = ((255 - (mHue / 60 * 255)) * (1 - mSat) + (mHue / 60 * 255)) * mBright
        mColor(B) = 255 * (1 - mSat) * mBright
    ElseIf mHue < 120 Then
        mColor(R) = ((255 - ((120 - mHue) / 60 * 255)) * (1 - mSat) + ((120 - mHue) / 60 * 255)) * mBright
        mColor(G) = mBright * 255
        mColor(B) = 255 * (1 - mSat) * mBright
    ElseIf mHue < 180 Then
        mColor(R) = 255 * (1 - mSat) * mBright
        mColor(G) = mBright * 255
        mColor(B) = ((255 - ((mHue - 120) / 60 * 255)) * (1 - mSat) + ((mHue - 120) / 60 * 255)) * mBright
    ElseIf mHue < 240 Then
        mColor(R) = 255 * (1 - mSat) * mBright
        mColor(G) = ((255 - ((240 - mHue) / 60 * 255)) * (1 - mSat) + ((240 - mHue) / 60 * 255)) * mBright
        mColor(B) = mBright * 255
    ElseIf mHue < 300 Then
        mColor(R) = ((255 - ((mHue - 240) / 60 * 255)) * (1 - mSat) + ((mHue - 240) / 60 * 255)) * mBright
        mColor(G) = 255 * (1 - mSat) * mBright
        mColor(B) = mBright * 255
    ElseIf mHue < 360 Then
        mColor(R) = mBright * 255
        mColor(G) = 255 * (1 - mSat) * mBright
        mColor(B) = ((255 - ((360 - mHue) / 60 * 255)) * (1 - mSat) + ((360 - mHue) / 60 * 255)) * mBright
    End If

    Exit Sub

ErrorHandler:

    MsgBox Error$

End Sub

Private Sub picSat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    picSat_MouseMove Button, Shift, X, Y

End Sub

Private Sub picSat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        X = 255 - Clamp(Y, 0, 255) ' grab y pos as it's a vertical bar
        mSat = X / 255
        If mColor(R) = mColor(G) And mColor(R) = mColor(B) And mSat > 0 Then 'determine rgb based on hue
            calculateHue
        Else
            mColor(mLow) = ((1 - mSat) * 255) * mBright
            mColor(mMid) = ((255 - mPureColor(mMid)) * (1 - mSat) + mPureColor(mMid)) * mBright
            mColor(mHigh) = mPureColor(mHigh) * mBright
        End If
        updateAll
        txtSat.Text = Int(mSat * 100 + 0.5)
        updateRGB
        updateHex

        picSpectrum.Circle (mOldX, mOldY), 5.5, RGB(0, 0, 0)
        mOldY = 255 - X
        picSpectrum.Circle (mOldX, mOldY), 5.5, RGB(0, 0, 0)
    End If

End Sub

Private Sub picBright_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    picBright_MouseMove Button, Shift, X, Y

End Sub

Private Sub picBright_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        X = 255 - Clamp(Y, 0, 255) ' grab y pos as it's a vertical bar
        mBright = X / 255
        mColor(mLow) = ((1 - mSat) * 255) * mBright
        mColor(mMid) = ((255 - mPureColor(mMid)) * (1 - mSat) + mPureColor(mMid)) * mBright
        mColor(mHigh) = mPureColor(mHigh) * mBright
        updateAll
        txtBright.Text = Int(mBright * 100 + 0.5)
        updateRGB
        updateHex
    End If

End Sub

Private Sub changeRGB() ' when rgb modified by user

    If mColor(R) = mColor(G) And mColor(R) = mColor(B) Then
        mBright = mColor(R) / 255
        mSat = 0
        If mHue < 60 Then
            mPureColor(R) = 255
            mPureColor(G) = (mHue / 60) * 255
            mPureColor(B) = 0
        ElseIf mHue < 120 Then
            mPureColor(R) = ((120 - mHue) / 60) * 255
            mPureColor(G) = 255
            mPureColor(B) = 0
        ElseIf mHue < 180 Then
            mPureColor(R) = 0
            mPureColor(G) = 255
            mPureColor(B) = ((mHue - 120) / 60) * 255
        ElseIf mHue < 240 Then
            mPureColor(R) = 0
            mPureColor(G) = ((240 - mHue) / 60) * 255
            mPureColor(B) = 255
        ElseIf mHue < 300 Then
            mPureColor(R) = ((mHue - 240) / 60) * 255
            mPureColor(G) = 0
            mPureColor(B) = 255
        ElseIf mHue < 360 Then
            mPureColor(R) = 255
            mPureColor(G) = 0
            mPureColor(B) = ((360 - mHue) / 60) * 255
        End If
        Exit Sub
    End If

    ' get hue from rgb
    If mColor(R) >= mColor(G) And mColor(R) >= mColor(B) Then
        If mColor(G) >= mColor(B) Then
            mHue = (mColor(G) - mColor(B)) / (mColor(R) - mColor(B)) * 60
        Else
            mHue = 360 - (mColor(B) - mColor(G)) / (mColor(R) - mColor(G)) * 60
        End If
    ElseIf mColor(G) >= mColor(R) And mColor(G) >= mColor(B) Then
        If mColor(R) >= mColor(B) Then
            mHue = 120 - (mColor(R) - mColor(B)) / (mColor(G) - mColor(B)) * 60
        Else
            mHue = (mColor(B) - mColor(R)) / (mColor(G) - mColor(R)) * 60 + 120
        End If
    ElseIf mColor(B) >= mColor(R) And mColor(B) >= mColor(G) Then
        If mColor(R) >= mColor(G) Then
            mHue = (mColor(R) - mColor(G)) / (mColor(B) - mColor(G)) * 60 + 240
        Else
            mHue = 240 - (mColor(G) - mColor(R)) / (mColor(B) - mColor(R)) * 60
        End If
    End If

    changeHue

    mSat = 1 - (mColor(mLow) / mColor(mHigh))
    mBright = mColor(mHigh) / 255

End Sub

Private Sub changeHue()

    If mHue < 60 Then
        mHigh = R
        mMid = G
        mLow = B
        mPureColor(R) = 255
        mPureColor(G) = (mHue / 60) * 255
        mPureColor(B) = 0
    ElseIf mHue < 120 Then
        mHigh = G
        mMid = R
        mLow = B
        mPureColor(R) = ((120 - mHue) / 60) * 255
        mPureColor(G) = 255
        mPureColor(B) = 0
    ElseIf mHue < 180 Then
        mHigh = G
        mMid = B
        mLow = R
        mPureColor(R) = 0
        mPureColor(G) = 255
        mPureColor(B) = ((mHue - 120) / 60) * 255
    ElseIf mHue < 240 Then
        mHigh = B
        mMid = G
        mLow = R
        mPureColor(R) = 0
        mPureColor(G) = ((240 - mHue) / 60) * 255
        mPureColor(B) = 255
    ElseIf mHue < 300 Then
        mHigh = B
        mMid = R
        mLow = G
        mPureColor(R) = ((mHue - 240) / 60) * 255
        mPureColor(G) = 0
        mPureColor(B) = 255
    ElseIf mHue < 360 Then
        mHigh = R
        mMid = B
        mLow = G
        mPureColor(R) = 255
        mPureColor(G) = 0
        mPureColor(B) = ((360 - mHue) / 60) * 255
    End If

End Sub

Private Sub updateAll()

    picColor.BackColor = RGB(mColor(R), mColor(G), mColor(B))

    imgRGB(R).Top = picRGB(R).Top + 255 - mColor(R) - 7
    imgRGB(G).Top = picRGB(G).Top + 255 - mColor(G) - 7
    imgRGB(B).Top = picRGB(B).Top + 255 - mColor(B) - 7

    imgHue.Top = picHue.Top + 255 - Int(mHue * 256 / 360) - 7
    imgSat.Top = picSat.Top + 255 - Int(mSat * 255) - 7
    imgBright.Top = picBright.Top + 255 - Int(mBright * 255) - 7

    Render

End Sub

Private Sub updateRGB()

    txtRGB(R).Text = mColor(R)
    txtRGB(G).Text = mColor(G)
    txtRGB(B).Text = mColor(B)

End Sub

Private Sub updateHSB()

    txtHue.Text = Int(mHue + 0.5)
    txtSat.Text = Int(mSat * 100 + 0.5)
    txtBright.Text = Int(mBright * 100 + 0.5)

End Sub

Private Sub updateHex()

    mHexValue = RGBtoHex(RGB(mColor(B), mColor(G), mColor(R)))
    txtHexCode.Text = RGBtoHex(RGB(mColor(B), mColor(G), mColor(R)))

End Sub

Private Sub Render()

    Dim i As Integer
    Dim redVal As Byte
    Dim greenVal As Byte
    Dim blueVal As Byte

    For i = 0 To 255
        picRGB(R).Line (0, 255 - i)-(16, 255 - i), RGB(i, mColor(G), mColor(B))
        picRGB(G).Line (0, 255 - i)-(16, 255 - i), RGB(mColor(R), i, mColor(B))
        picRGB(B).Line (0, 255 - i)-(16, 255 - i), RGB(mColor(R), mColor(G), i)

        redVal = ((255 - mPureColor(R)) * (1 - i / 255) + mPureColor(R)) * mBright
        greenVal = ((255 - mPureColor(G)) * (1 - i / 255) + mPureColor(G)) * mBright
        blueVal = ((255 - mPureColor(B)) * (1 - i / 255) + mPureColor(B)) * mBright
        picSat.Line (0, 255 - i)-(16, 255 - i), RGB(redVal, greenVal, blueVal)

        redVal = ((255 - mPureColor(R)) * (1 - mSat) + mPureColor(R)) * (i / 255)
        greenVal = ((255 - mPureColor(G)) * (1 - mSat) + mPureColor(G)) * (i / 255)
        blueVal = ((255 - mPureColor(B)) * (1 - mSat) + mPureColor(B)) * (i / 255)
        picBright.Line (0, 255 - i)-(16, 255 - i), RGB(redVal, greenVal, blueVal)

        If i <= (255 / 6) Then
            redVal = mBright * 255
            greenVal = ((255 - (i * 6)) * (1 - mSat) + (i * 6)) * mBright
            blueVal = 255 * (1 - mSat) * mBright
        ElseIf i <= (255 / 3) Then
            redVal = ((255 - ((255 / 3 - i) * 6)) * (1 - mSat) + ((255 / 3 - i) * 6)) * mBright
            greenVal = mBright * 255
            blueVal = 255 * (1 - mSat) * mBright
        ElseIf i <= (255 / 2) Then
            redVal = 255 * (1 - mSat) * mBright
            greenVal = mBright * 255
            blueVal = ((255 - ((i - 255 / 3) * 6)) * (1 - mSat) + ((i - 255 / 3) * 6)) * mBright
        ElseIf i <= (255 / 3 * 2) Then
            redVal = 255 * (1 - mSat) * mBright
            greenVal = ((255 - ((255 / 3 * 2 - i) * 6)) * (1 - mSat) + ((255 / 3 * 2 - i) * 6)) * mBright
            blueVal = mBright * 255
        ElseIf i <= (255 / 6 * 5) Then
            redVal = ((255 - ((i - 255 / 3 * 2) * 6)) * (1 - mSat) + ((i - 255 / 3 * 2) * 6)) * mBright
            greenVal = 255 * (1 - mSat) * mBright
            blueVal = mBright * 255
        ElseIf i <= 255 Then
            redVal = mBright * 255
            greenVal = 255 * (1 - mSat) * mBright
            blueVal = ((255 - ((255 - i) * 6)) * (1 - mSat) + ((255 - i) * 6)) * mBright
        End If

        picHue.Line (0, 255 - i)-(16, 255 - i), RGB(redVal, greenVal, blueVal)
    Next

    picRGB(R).Refresh
    picRGB(G).Refresh
    picRGB(B).Refresh
    picHue.Refresh
    picSat.Refresh
    picBright.Refresh

End Sub

Private Sub txtHexCode_Change()

    Dim tempHexVal As String

    If HexToLong(txtHexCode.Text) = -1 Then
        ' no-op
    ElseIf mHexValue <> txtHexCode.Text Then
        If Len(txtHexCode.Text) < 6 Then
            tempHexVal = String$(6 - Len(txtHexCode.Text), "0") & txtHexCode.Text
        ElseIf Len(txtHexCode.Text) > 6 Then
            tempHexVal = right(txtHexCode.Text, 6)
        Else
            tempHexVal = txtHexCode.Text
        End If
        mColor(B) = CLng("&H" + right(tempHexVal, 2))
        tempHexVal = Left(tempHexVal, Len(tempHexVal) - 2)
        mColor(G) = CLng("&H" + right(tempHexVal, 2))
        tempHexVal = Left(tempHexVal, Len(tempHexVal) - 2)
        mColor(R) = CLng("&H" + right(tempHexVal, 2))
        changeRGB
        updateAll
        updateRGB
        updateHSB
    End If

End Sub

Private Sub txtHexCode_LostFocus()

    If HexToLong(txtHexCode.Text) = -1 Then
        txtHexCode.Text = mHexValue
        mColor(B) = CLng("&H" + right(mHexValue, 2))
        mHexValue = Left(mHexValue, Len(mHexValue) - 2)
        mColor(G) = CLng("&H" + right(mHexValue, 2))
        mHexValue = Left(mHexValue, Len(mHexValue) - 2)
        mColor(R) = CLng("&H" + right(mHexValue, 2))
        changeRGB
        updateAll
        updateRGB
        updateHSB
    Else
        If Len(txtHexCode.Text) > 6 Then
            txtHexCode.Text = right(txtHexCode.Text, 6)
        ElseIf Len(txtHexCode.Text) < 6 Then
            txtHexCode = String$(6 - Len(txtHexCode.Text), "0") & txtHexCode.Text
        End If
        mHexValue = txtHexCode.Text
    End If

End Sub

Private Sub txtRGB_Change(Index As Integer)

    If IsNumeric(txtRGB(Index).Text) = False And txtRGB(Index).Text <> "" Then
        txtRGB(Index).Text = mColor(Index)
    ElseIf txtRGB(Index).Text = "" Then
        ' no-op
    ElseIf txtRGB(Index).Text >= 0 And txtRGB(Index).Text <= 255 Then
        If mColor(Index) <> txtRGB(Index).Text Then
            mColor(Index) = txtRGB(Index).Text
            changeRGB
            updateAll
            updateHSB
            updateHex
        End If
    End If

End Sub

Private Sub txtRGB_GotFocus(Index As Integer)

    SelectAllText txtRGB(Index)

End Sub

Private Sub txtRGB_LostFocus(Index As Integer)

    txtRGB(Index).Text = mColor(Index)

End Sub

Private Sub txtHue_Change()

    If IsNumeric(txtHue.Text) = False And txtHue.Text <> "" Then
        txtHue.Text = Int(mHue + 0.5)
    ElseIf txtHue.Text = "" Then
        ' no-op
    ElseIf txtHue.Text >= 0 And txtHue.Text <= 359 Then
        If Int(mHue + 0.5) <> txtHue.Text Then
            mHue = txtHue.Text
            If Not (mColor(R) = mColor(G) And mColor(R) = mColor(B)) Then
                calculateHue
            Else

            End If
            changeHue
            updateAll
            updateRGB
            updateHex

            picSpectrum.Circle (mOldX, mOldY), 5.5, RGB(0, 0, 0)
            mOldX = mHue / 360 * 256
            picSpectrum.Circle (mOldX, mOldY), 5.5, RGB(0, 0, 0)
        End If
    End If

End Sub

Private Sub txtHue_GotFocus()

    SelectAllText txtHue

End Sub

Private Sub txtHue_LostFocus()

    txtHue.Text = Int(mHue + 0.5)

End Sub

Private Sub txtSat_Change()

    If IsNumeric(txtSat.Text) = False And txtSat.Text <> "" Then
        txtSat.Text = Int(mSat * 100 + 0.5)
    ElseIf txtSat.Text = "" Then
        ' no-op
    ElseIf txtSat.Text >= 0 And txtSat.Text <= 100 Then
        If Int(mSat * 100 + 0.5) <> txtSat.Text Then
            mSat = txtSat.Text / 100
            mColor(mLow) = ((1 - mSat) * 255) * mBright
            mColor(mMid) = ((255 - mPureColor(mMid)) * (1 - mSat) + mPureColor(mMid)) * mBright
            mColor(mHigh) = mPureColor(mHigh) * mBright
            updateAll
            updateRGB
            updateHex

            picSpectrum.Circle (mOldX, mOldY), 5.5, RGB(0, 0, 0)
            mOldY = 255 - mSat * 255
            picSpectrum.Circle (mOldX, mOldY), 5.5, RGB(0, 0, 0)
        End If
    End If

End Sub

Private Sub txtSat_LostFocus()

    txtSat.Text = Int(mSat * 100 + 0.5)

End Sub

Private Sub txtSat_GotFocus()

    SelectAllText txtSat

End Sub

Private Sub txtBright_Change()

    If IsNumeric(txtBright.Text) = False And txtBright.Text <> "" Then
        txtBright.Text = Int(mBright * 100 + 0.5)
    ElseIf txtBright.Text = "" Then
        ' no-op
    ElseIf txtBright.Text >= 0 And txtBright.Text <= 100 Then
        If Int(mBright * 100 + 0.5) <> txtBright.Text Then
            mBright = txtBright.Text / 100
            mColor(mLow) = ((1 - mSat) * 255) * mBright
            mColor(mMid) = ((255 - mPureColor(mMid)) * (1 - mSat) + mPureColor(mMid)) * mBright
            mColor(mHigh) = mPureColor(mHigh) * mBright
            updateAll
            updateRGB
            updateHex
        End If
    End If

End Sub

Private Sub txtBright_LostFocus()

    txtBright.Text = Int(mBright * 100 + 0.5)

End Sub

Private Sub txtBright_GotFocus()

    SelectAllText txtBright

End Sub

Private Sub picTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ReleaseCapture
    SendMessage Me.hWnd, WM_NCLBUTTONDOWN, 2, 0&

End Sub

Private Sub picHide_Click()

    HideColor False

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

Private Sub picCancel_Click()

    ok = False
    HideColor False

End Sub

Private Sub picCancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picCancel, X, Y, BUTTON_LARGE, 0, BUTTON_DOWN

End Sub

Private Sub picCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picCancel, X, Y, BUTTON_LARGE, 0, BUTTON_MOVE

End Sub

Private Sub picCancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picCancel, 0, 0, BUTTON_LARGE, 0, BUTTON_UP

End Sub

Private Sub picOK_Click()

    ok = True
    red = mColor(R)
    green = mColor(G)
    blue = mColor(B)

    HideColor True

End Sub

Private Sub picOK_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picOK, X, Y, BUTTON_LARGE, 0, BUTTON_DOWN

End Sub

Private Sub picOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picOK, X, Y, BUTTON_LARGE, 0, BUTTON_MOVE

End Sub

Private Sub picOK_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picOK, 0, 0, BUTTON_LARGE, 0, BUTTON_UP

End Sub

Public Sub SetColors()

    On Error Resume Next

    Dim c As Control

    picTitle.Picture = LoadPicture(appPath & "\" & gfxDir & "\titlebar_colorpicker.bmp")
    picSpectrum.Picture = LoadPicture(appPath & "\" & gfxDir & "\color_picker.bmp")

    mouseEvent2 picHide, 0, 0, BUTTON_SMALL, 0, BUTTON_UP
    mouseEvent2 picOK, 0, 0, BUTTON_LARGE, 0, BUTTON_UP
    mouseEvent2 picCancel, 0, 0, BUTTON_LARGE, 0, BUTTON_UP

    For Each c In imgRGB
        c.Picture = LoadPicture(appPath & "\" & gfxDir & "\slider_arrow.bmp")
    Next
    imgHue.Picture = LoadPicture(appPath & "\" & gfxDir & "\slider_arrow.bmp")
    imgBright.Picture = LoadPicture(appPath & "\" & gfxDir & "\slider_arrow.bmp")
    imgSat.Picture = LoadPicture(appPath & "\" & gfxDir & "\slider_arrow.bmp")

    picSpectrum.MouseIcon = LoadPicture(appPath & "\" & gfxDir & "\cursors\color_picker.cur")


    Me.BackColor = bgClr

    For Each c In lblClr
        c.BackColor = lblBackClr
        c.ForeColor = lblTextClr
    Next

    For Each c In txtRGB
        c.BackColor = txtBackClr
        c.ForeColor = txtTextClr
    Next

    txtHue.BackColor = txtBackClr
    txtHue.ForeColor = txtTextClr

    txtSat.BackColor = txtBackClr
    txtSat.ForeColor = txtTextClr

    txtBright.BackColor = txtBackClr
    txtBright.ForeColor = txtTextClr

    txtHexCode.BackColor = bgClr
    txtHexCode.ForeColor = lblTextClr

    For Each c In Me.Controls
        If c.Tag = "font1" Then
            c.Font.Name = font1
        ElseIf c.Tag = "font2" Then
            c.Font.Name = font2
        End If
    Next

End Sub
