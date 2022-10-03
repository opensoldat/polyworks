VERSION 5.00
Begin VB.Form frmTaskBar 
   Caption         =   "OpenSoldat PolyWorks"
   ClientHeight    =   1920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1920
   Icon            =   "frmTaskBar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   128
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   128
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmTaskBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' taskbar emulation - show taskbar button with behavior


' Fix vb6 ide casing changes
#If False Then
    Public FileName, color, token, A, R, G, B, commonDialog, value, Val, X, Y, Z, Left, hWnd, Mid, Right, BackColor
    'Public FileName, color, token, A, R, G, B, commonDialog, value, Val, X, Y, Z, Left, hWnd, Mid, Right, BackColor
#End If


' vars - public


' vars - private


' functions - public


' functions - private


' events - public


' events - private

Private Sub Form_GotFocus()

    If frmOpenSoldatMapEditor.Visible Then
        frmOpenSoldatMapEditor.RegainFocus
    End If

End Sub

Private Sub Form_Load()

    Me.Left = 2000 * Screen.TwipsPerPixelX
    Me.Top = 2000 * Screen.TwipsPerPixelY

End Sub

Private Sub Form_Resize()

    If Not frmOpenSoldatMapEditor.Visible And Me.WindowState = vbNormal Then  ' show when it gets restored
        frmOpenSoldatMapEditor.Show
        If frmOpenSoldatMapEditor.mnuDisplay.Checked Then frmDisplay.Show
        If frmOpenSoldatMapEditor.mnuWaypoints.Checked Then frmWaypoints.Show
        If frmOpenSoldatMapEditor.mnuTools.Checked Then frmTools.Show
        If frmOpenSoldatMapEditor.mnuPalette.Checked Then frmPalette.Show
        If frmOpenSoldatMapEditor.mnuScenery.Checked Then frmScenery.Show
        If frmOpenSoldatMapEditor.mnuInfo.Checked Then frmInfo.Show
        If frmOpenSoldatMapEditor.mnuTexture.Checked Then frmTexture.Show
        If frmOpenSoldatMapEditor.Tag = vbNormal Then
            frmOpenSoldatMapEditor.Left = frmOpenSoldatMapEditor.formLeft * Screen.TwipsPerPixelX
            frmOpenSoldatMapEditor.Top = frmOpenSoldatMapEditor.formTop * Screen.TwipsPerPixelY
            frmOpenSoldatMapEditor.ScaleWidth = frmOpenSoldatMapEditor.formWidth
            frmOpenSoldatMapEditor.ScaleHeight = frmOpenSoldatMapEditor.formHeight
        End If
        frmOpenSoldatMapEditor.RegainFocus
    ElseIf Not frmOpenSoldatMapEditor.Visible And Me.WindowState = vbMinimized Then
        '  no-op
    ElseIf frmOpenSoldatMapEditor.Visible And Me.WindowState = vbNormal Then
        frmOpenSoldatMapEditor.RegainFocus
    ElseIf frmOpenSoldatMapEditor.Visible And Me.WindowState = vbMinimized Then
        frmOpenSoldatMapEditor.RegainFocus
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If frmColor.Visible Then
        frmColor.Hide
    End If

    frmOpenSoldatMapEditor.Terminate

End Sub
