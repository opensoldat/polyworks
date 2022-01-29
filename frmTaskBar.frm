VERSION 5.00
Begin VB.Form frmTaskBar 
   Caption         =   "Soldat PolyWorks"
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

Dim minimized As Boolean

Private Sub Form_GotFocus()

    If frmSoldatMapEditor.Visible Then
        frmSoldatMapEditor.RegainFocus
    End If

End Sub

Private Sub Form_Load()

    Me.left = 2000 * Screen.TwipsPerPixelX
    Me.Top = 2000 * Screen.TwipsPerPixelY

End Sub

Private Sub Form_Resize()

    If Not frmSoldatMapEditor.Visible And Me.WindowState = vbNormal Then 'show when it gets restored
        frmSoldatMapEditor.Show
        If frmSoldatMapEditor.mnuDisplay.Checked Then frmDisplay.Show
        If frmSoldatMapEditor.mnuWaypoints.Checked Then frmWaypoints.Show
        If frmSoldatMapEditor.mnuTools.Checked Then frmTools.Show
        If frmSoldatMapEditor.mnuPalette.Checked Then frmPalette.Show
        If frmSoldatMapEditor.mnuScenery.Checked Then frmScenery.Show
        If frmSoldatMapEditor.mnuInfo.Checked Then frmInfo.Show
        If frmSoldatMapEditor.mnuTexture.Checked Then frmTexture.Show
        If frmSoldatMapEditor.WindowState = vbNormal Then
            frmSoldatMapEditor.left = frmSoldatMapEditor.formLeft * Screen.TwipsPerPixelX
            frmSoldatMapEditor.Top = frmSoldatMapEditor.formTop * Screen.TwipsPerPixelY
            frmSoldatMapEditor.ScaleWidth = frmSoldatMapEditor.formWidth
            frmSoldatMapEditor.ScaleHeight = frmSoldatMapEditor.formHeight
        End If
        frmSoldatMapEditor.RegainFocus
    ElseIf Not frmSoldatMapEditor.Visible And Me.WindowState = vbMinimized Then
        ' no-op
    ElseIf frmSoldatMapEditor.Visible And Me.WindowState = vbNormal Then
        frmSoldatMapEditor.RegainFocus
    ElseIf frmSoldatMapEditor.Visible And Me.WindowState = vbMinimized Then
        frmSoldatMapEditor.RegainFocus
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    frmSoldatMapEditor.Terminate

End Sub
