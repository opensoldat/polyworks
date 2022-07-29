VERSION 5.00
Begin VB.Form frmTaskBar 
   Caption         =   "opensoldat PolyWorks"
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

    If frmOpensoldatMapEditor.Visible Then
        frmOpensoldatMapEditor.RegainFocus
    End If

End Sub

Private Sub Form_Load()

    Me.Left = 2000 * Screen.TwipsPerPixelX
    Me.Top = 2000 * Screen.TwipsPerPixelY

End Sub

Private Sub Form_Resize()

    If Not frmOpensoldatMapEditor.Visible And Me.WindowState = vbNormal Then  ' show when it gets restored
        frmOpensoldatMapEditor.Show
        If frmOpensoldatMapEditor.mnuDisplay.Checked Then frmDisplay.Show
        If frmOpensoldatMapEditor.mnuWaypoints.Checked Then frmWaypoints.Show
        If frmOpensoldatMapEditor.mnuTools.Checked Then frmTools.Show
        If frmOpensoldatMapEditor.mnuPalette.Checked Then frmPalette.Show
        If frmOpensoldatMapEditor.mnuScenery.Checked Then frmScenery.Show
        If frmOpensoldatMapEditor.mnuInfo.Checked Then frmInfo.Show
        If frmOpensoldatMapEditor.mnuTexture.Checked Then frmTexture.Show
        If frmOpensoldatMapEditor.Tag = vbNormal Then
            frmOpensoldatMapEditor.Left = frmOpensoldatMapEditor.formLeft * Screen.TwipsPerPixelX
            frmOpensoldatMapEditor.Top = frmOpensoldatMapEditor.formTop * Screen.TwipsPerPixelY
            frmOpensoldatMapEditor.ScaleWidth = frmOpensoldatMapEditor.formWidth
            frmOpensoldatMapEditor.ScaleHeight = frmOpensoldatMapEditor.formHeight
        End If
        frmOpensoldatMapEditor.RegainFocus
    ElseIf Not frmOpensoldatMapEditor.Visible And Me.WindowState = vbMinimized Then
        '  no-op
    ElseIf frmOpensoldatMapEditor.Visible And Me.WindowState = vbNormal Then
        frmOpensoldatMapEditor.RegainFocus
    ElseIf frmOpensoldatMapEditor.Visible And Me.WindowState = vbMinimized Then
        frmOpensoldatMapEditor.RegainFocus
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If frmColor.Visible Then
        frmColor.Hide
    End If

    frmOpensoldatMapEditor.Terminate

End Sub
