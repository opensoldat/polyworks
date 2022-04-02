VERSION 5.00
Begin VB.Form frmMap 
   Appearance      =   0  'Flat
   BackColor       =   &H004A3C31&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5040
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   5400
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   336
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Height          =   375
      Left            =   4320
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   13
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   4560
      Width           =   975
   End
   Begin VB.TextBox txtJet 
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
      Left            =   3120
      TabIndex        =   12
      Tag             =   "font1"
      Top             =   1680
      Width           =   855
   End
   Begin VB.PictureBox picTexture 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   3240
      ScaleHeight     =   127
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   127
      TabIndex        =   11
      ToolTipText     =   "Map Texture"
      Top             =   2400
      Width           =   1935
   End
   Begin VB.ComboBox cboTexture 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      ItemData        =   "frmMap.frx":0000
      Left            =   360
      List            =   "frmMap.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Tag             =   "font1"
      ToolTipText     =   "Map Texture"
      Top             =   2760
      Width           =   2655
   End
   Begin VB.TextBox txtDesc 
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
      Left            =   1920
      MaxLength       =   38
      TabIndex        =   0
      Tag             =   "font1"
      Text            =   "New Soldat Map"
      ToolTipText     =   "Map Description"
      Top             =   480
      Width           =   3135
   End
   Begin VB.PictureBox picBackColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   1920
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   9
      ToolTipText     =   "Top Background Color"
      Top             =   3240
      Width           =   495
   End
   Begin VB.PictureBox picBackColor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   1920
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   8
      ToolTipText     =   "Bottom Background Color"
      Top             =   3840
      Width           =   495
   End
   Begin VB.ComboBox cboWeather 
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
      Height          =   285
      ItemData        =   "frmMap.frx":0004
      Left            =   1560
      List            =   "frmMap.frx":0014
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Tag             =   "font1"
      ToolTipText     =   "Weather"
      Top             =   960
      Width           =   1335
   End
   Begin VB.ComboBox cboJet 
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
      Height          =   285
      ItemData        =   "frmMap.frx":0035
      Left            =   1560
      List            =   "frmMap.frx":0054
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Tag             =   "font1"
      ToolTipText     =   "Jets"
      Top             =   1680
      Width           =   1335
   End
   Begin VB.ComboBox cboGrenades 
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
      Height          =   285
      ItemData        =   "frmMap.frx":009F
      Left            =   4440
      List            =   "frmMap.frx":00CA
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Tag             =   "font1"
      ToolTipText     =   "Grenade Kits"
      Top             =   960
      Width           =   615
   End
   Begin VB.ComboBox cboMedikits 
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
      Height          =   285
      ItemData        =   "frmMap.frx":00F8
      Left            =   4440
      List            =   "frmMap.frx":0123
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Tag             =   "font1"
      ToolTipText     =   "Medikits"
      Top             =   1320
      Width           =   615
   End
   Begin VB.ComboBox cboSteps 
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
      Height          =   285
      ItemData        =   "frmMap.frx":0151
      Left            =   1560
      List            =   "frmMap.frx":015E
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Tag             =   "font1"
      ToolTipText     =   "Steps"
      Top             =   1320
      Width           =   1335
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
      ScaleWidth      =   360
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   5400
      Begin VB.PictureBox picHide 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H004A3C31&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   5160
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   "3"
         Top             =   0
         Width           =   240
      End
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
      Height          =   375
      Left            =   3960
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   14
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   3840
      Width           =   975
   End
   Begin VB.Shape fraMap 
      BorderColor     =   &H000B3C0D&
      Height          =   2175
      Index           =   1
      Left            =   120
      Top             =   2280
      Width           =   5175
   End
   Begin VB.Shape fraMap 
      BorderColor     =   &H000B3C0D&
      Height          =   1815
      Index           =   0
      Left            =   120
      Top             =   360
      Width           =   5175
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00614B3D&
      Caption         =   "Texture:"
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
      TabIndex        =   22
      Tag             =   "font2"
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00614B3D&
      Caption         =   "Background:"
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
      TabIndex        =   21
      Tag             =   "font2"
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00614B3D&
      Caption         =   "Description:"
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
      TabIndex        =   20
      Tag             =   "font2"
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00614B3D&
      Caption         =   "Medikits:"
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
      Left            =   3120
      TabIndex        =   19
      Tag             =   "font2"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00614B3D&
      Caption         =   "Grenades:"
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
      Left            =   3120
      TabIndex        =   18
      Tag             =   "font2"
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00614B3D&
      Caption         =   "Jet Fuel:"
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
      TabIndex        =   17
      Tag             =   "font2"
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00614B3D&
      Caption         =   "Steps:"
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
      TabIndex        =   16
      Tag             =   "font2"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblMap 
      BackColor       =   &H00614B3D&
      Caption         =   "Weather:"
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
      TabIndex        =   15
      Tag             =   "font2"
      Top             =   960
      Width           =   1095
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' maps dialog


' Fix vb6 ide casing changes
#If False Then
    Public FileName, color, token, A, R, G, B, commonDialog, value, Val, X, Y, Z, Left, hWnd, Mid, Right, BackColor
    'Public FileName, color, token, A, R, G, B, commonDialog, value, Val, X, Y, Z, Left, hWnd, Mid, Right, BackColor
#End If


' vars - public


' vars - private


' functions - public

Public Sub LoadTextures()

    On Error GoTo ErrorHandler

    Dim strParent As String
    Dim strPath As String

    Dim objFSO As FileSystemObject
    Dim objFiles As Files
    Dim objFile As file

    cboTexture.Clear

    strParent = frmSoldatMapEditor.soldatDir
    strPath = frmSoldatMapEditor.soldatDir & "textures\"

    Set objFSO = New FileSystemObject

    If Not objFSO.FolderExists(strPath) Then Exit Sub

    Set objFiles = objFSO.GetFolder(strPath).Files

    For Each objFile In objFiles
        If Right(objFile.Name, 3) = "bmp" Then
            cboTexture.AddItem objFile.Name
        End If
    Next

    Exit Sub

ErrorHandler:

    MsgBox "loading textures failed" & vbNewLine & Error$

End Sub

Public Sub LoadTextures2()

    On Error GoTo ErrorHandler

    Dim file As String

    cboTexture.Clear

    file = Dir$(frmSoldatMapEditor.soldatDir & "textures\" & "*.bmp", vbDirectory)
    Do While Len(file)
        cboTexture.AddItem file
        file = Dir$
    Loop

    file = Dir$(frmSoldatMapEditor.soldatDir & "textures\" & "*.png", vbDirectory)
    Do While Len(file)
        cboTexture.AddItem file
        file = Dir$
    Loop

    Exit Sub

ErrorHandler:

    MsgBox "loading textures failed" & vbNewLine & Error$

End Sub

Public Sub LoadFromList()  ' unused?

    On Error GoTo ErrorHandler

    Dim textureName As String

    cboTexture.Clear

    Open appPath & "\texture_list.txt" For Input As #1

        Do While Not EOF(1)
            Input #1, textureName
            cboTexture.AddItem textureName
        Loop

    Close #1

    Exit Sub

ErrorHandler:

    MsgBox Error$

End Sub

Public Sub SetColors()

    On Error Resume Next

    Dim i As Integer
    Dim c As Control

    picTitle.Picture = LoadPicture(appPath & "\skins\" & gfxDir & "\titlebar_map.bmp")

    MouseEvent2 picHide, 0, 0, BUTTON_SMALL, 0, BUTTON_UP
    MouseEvent2 picOK, 0, 0, BUTTON_LARGE, 0, BUTTON_UP
    MouseEvent2 picCancel, 0, 0, BUTTON_LARGE, 0, BUTTON_UP

    Me.BackColor = bgColor

    For Each c In lblMap
        c.BackColor = lblBackColor
        c.ForeColor = lblTextColor
    Next

    txtDesc.BackColor = txtBackColor
    txtDesc.ForeColor = txtTextColor
    txtJet.BackColor = txtBackColor
    txtJet.ForeColor = txtTextColor

    cboWeather.BackColor = txtBackColor
    cboWeather.ForeColor = txtTextColor
    cboSteps.BackColor = txtBackColor
    cboSteps.ForeColor = txtTextColor
    cboJet.BackColor = txtBackColor
    cboJet.ForeColor = txtTextColor
    cboGrenades.BackColor = txtBackColor
    cboGrenades.ForeColor = txtTextColor
    cboMedikits.BackColor = txtBackColor
    cboMedikits.ForeColor = txtTextColor
    cboTexture.BackColor = txtBackColor
    cboTexture.ForeColor = txtTextColor

    For Each c In fraMap
        c.BorderColor = frameColor
    Next

    SetFormFonts Me

End Sub


' functions - private


' events - public

Public Sub Form_Load()

    On Error GoTo ErrorHandler

    Me.SetColors
    LoadTextures2
    frmSoldatMapEditor.GetOptions
    GetJets

    Exit Sub

ErrorHandler:

    MsgBox Error$ & vbNewLine & "Error loading Map form"

End Sub

Public Sub mnuRefresh_Click()

    Dim i As Integer

    LoadTextures2

    For i = 0 To cboTexture.ListCount - 1
        If cboTexture.List(i) = frmSoldatMapEditor.gTextureFile And cboTexture.List(i) <> "" Then
            cboTexture.ListIndex = i
        End If
    Next

End Sub


' events - private

Private Sub cboJet_Click()

    Select Case cboJet.ListIndex
        Case 0  ' none
            txtJet.Text = "0"
        Case 1  ' minimal
            txtJet.Text = "12"
        Case 2  ' very low
            txtJet.Text = "45"
        Case 3  ' low
            txtJet.Text = "95"
        Case 4  ' normal
            txtJet.Text = "190"
        Case 5  ' high
            txtJet.Text = "320"
        Case 6  ' maximum
            txtJet.Text = "800"
        Case 7  ' infinite
            txtJet.Text = "32766"
        Case 8  ' custom
    End Select

    If cboJet.ListIndex <> 8 Then
        txtJet.Enabled = False
    Else
        txtJet.Enabled = True
    End If

End Sub

Private Sub GetJets()

    Select Case txtJet.Text
        Case 0  ' none
            cboJet.ListIndex = 0
        Case 12  ' minimal
            cboJet.ListIndex = 1
        Case 45  ' very low
            cboJet.ListIndex = 2
        Case 95  ' low
            cboJet.ListIndex = 3
        Case 190  ' normal
            cboJet.ListIndex = 4
        Case 320  ' high
            cboJet.ListIndex = 5
        Case 800  ' maximum
            cboJet.ListIndex = 6
        Case 32766  ' infinite
            cboJet.ListIndex = 7
        Case Else  ' custom
            cboJet.ListIndex = 8
    End Select

End Sub

Private Sub cboTexture_Click()

    On Error GoTo ErrorHandler

    If cboTexture.List(cboTexture.ListIndex) <> "" Then
        frmSoldatMapEditor.SetMapTexture cboTexture.List(cboTexture.ListIndex)
        frmTexture.SetTexture cboTexture.List(cboTexture.ListIndex)

        Dim token As Long
        token = InitGDIPlus
        picTexture.Picture = LoadPictureGDIPlus(frmSoldatMapEditor.soldatDir & "textures\" & cboTexture.List(cboTexture.ListIndex), 128, 128)
        FreeGDIPlus token
    End If

    Exit Sub

ErrorHandler:

    MsgBox "Error showing texture" & vbNewLine & Error$

End Sub

Private Sub picBackColor_Click(Index As Integer)

    picBackColor(Index).BackColor = frmSoldatMapEditor.SetBGColor(Index + 1)

End Sub

Private Sub picCancel_Click()

    Unload Me

End Sub

Private Sub picOK_Click()

    frmSoldatMapEditor.SetOptions
    Unload Me
    frmSoldatMapEditor.RegainFocus

End Sub

Private Sub picTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ReleaseCapture
    SendMessage Me.hWnd, WM_NCLBUTTONDOWN, 2, 0&

End Sub

Private Sub picCancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picCancel, X, Y, BUTTON_LARGE, 0, BUTTON_DOWN

End Sub

Private Sub picCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picCancel, X, Y, BUTTON_LARGE, 0, BUTTON_MOVE

End Sub

Private Sub picOK_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picOK, X, Y, BUTTON_LARGE, 0, BUTTON_DOWN

End Sub

Private Sub picOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picOK, X, Y, BUTTON_LARGE, 0, BUTTON_MOVE

End Sub

Private Sub picHide_Click()

    frmSoldatMapEditor.SetOptions
    frmSoldatMapEditor.mnuMap.Checked = False
    Unload Me

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

Private Sub txtJet_KeyPress(KeyAscii As Integer)

    If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
    Else
        KeyAscii = 0
    End If

End Sub
