VERSION 5.00
Object = "{DDA53BD0-2CD0-11D4-8ED4-00E07D815373}#1.0#0"; "MBMouse.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSoldatMapEditor 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9000
   ClientLeft      =   3600
   ClientTop       =   3180
   ClientWidth     =   12000
   ControlBox      =   0   'False
   DrawMode        =   6  'Mask Pen Not
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSoldatMapEditor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picResize 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   11700
      MousePointer    =   8  'Size NW SE
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   8715
      Width           =   300
   End
   Begin VB.PictureBox picButtonGfx 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   4080
      ScaleHeight     =   193
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   241
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   3615
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   4080
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      _Version        =   393216
   End
   Begin VB.PictureBox picStatus 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   0
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   8730
      Width           =   12000
      Begin VB.TextBox txtZoom 
         Appearance      =   0  'Flat
         BackColor       =   &H004A3C31&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   3000
         MousePointer    =   3  'I-Beam
         TabIndex        =   14
         TabStop         =   0   'False
         Tag             =   "font1"
         Top             =   45
         Width           =   975
      End
      Begin VB.Label lblMousePosition 
         BackStyle       =   0  'Transparent
         Caption         =   "Position:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   8160
         TabIndex        =   20
         Tag             =   "font2"
         Top             =   45
         Width           =   3735
      End
      Begin VB.Label lblFileName 
         BackColor       =   &H004A3C31&
         BackStyle       =   0  'Transparent
         Caption         =   "Untitled.pms"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Tag             =   "font2"
         Top             =   45
         Width           =   2055
      End
      Begin VB.Label lblZoom 
         BackStyle       =   0  'Transparent
         Caption         =   "Zoom:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2400
         MousePointer    =   1  'Arrow
         TabIndex        =   12
         Tag             =   "font2"
         Top             =   45
         Width           =   615
      End
      Begin VB.Label lblCurrentTool 
         BackStyle       =   0  'Transparent
         Caption         =   "Current Tool:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4080
         TabIndex        =   11
         Tag             =   "font2"
         Top             =   45
         Width           =   3855
      End
   End
   Begin VB.PictureBox picMenuBar 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   0
      MousePointer    =   1  'Arrow
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   375
      Width           =   12000
      Begin VB.PictureBox picMenu 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H004A3C31&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   3840
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   0
         Width           =   960
      End
      Begin VB.PictureBox picMenu 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H004A3C31&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   1920
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   0
         Width           =   960
      End
      Begin VB.PictureBox picMenu 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H004A3C31&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   960
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         Width           =   960
      End
      Begin VB.PictureBox picMenu 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H004A3C31&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   2880
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Width           =   960
      End
      Begin VB.PictureBox picMenu 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H004A3C31&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   0
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   960
      End
      Begin VB.PictureBox picProgress 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H004A3C31&
         BorderStyle     =   0  'None
         FillColor       =   &H007B614A&
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   9960
         ScaleHeight     =   11
         ScaleMode       =   0  'User
         ScaleWidth      =   128
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   30
         Width           =   1920
      End
   End
   Begin VB.PictureBox picTitle 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      MousePointer    =   1  'Arrow
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   0
      Top             =   0
      Width           =   12000
      Begin VB.PictureBox picHelp 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H004A3C31&
         BorderStyle     =   0  'None
         FillColor       =   &H80000008&
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   10800
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   17
         TabStop         =   0   'False
         Tag             =   "9"
         ToolTipText     =   "Help"
         Top             =   0
         Width           =   240
      End
      Begin VB.PictureBox picMinimize 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H004A3C31&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   11280
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   8
         TabStop         =   0   'False
         Tag             =   "0"
         ToolTipText     =   "Minimize"
         Top             =   0
         Width           =   240
      End
      Begin VB.PictureBox picMaximize 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H004A3C31&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   11520
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   "1"
         ToolTipText     =   "Restore Down/Maximize"
         Top             =   0
         Width           =   240
      End
      Begin VB.PictureBox picExit 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H004A3C31&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   11760
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   6
         TabStop         =   0   'False
         Tag             =   "3"
         ToolTipText     =   "Close"
         Top             =   0
         Width           =   240
      End
   End
   Begin MSComDlg.CommonDialog commonDialog 
      Left            =   3120
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox picGfx 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   2520
      ScaleHeight     =   449
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSComctlLib.TreeView tvwScenery 
      Height          =   8085
      Left            =   0
      TabIndex        =   18
      Tag             =   "font1"
      Top             =   600
      Visible         =   0   'False
      Width           =   5730
      _ExtentX        =   10107
      _ExtentY        =   14261
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   423
      LabelEdit       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      Appearance      =   0
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
   End
   Begin MBMouseHelper.MouseHelper MouseHelper 
      Left            =   2520
      Top             =   600
      _ExtentX        =   900
      _ExtentY        =   900
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "&File"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuOpenCompiled 
         Caption         =   "O&pen Compiled..."
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "Open &Recent"
         Begin VB.Menu mnuRecent 
            Caption         =   ""
            Index           =   0
         End
         Begin VB.Menu mnuRecent 
            Caption         =   ""
            Index           =   1
         End
         Begin VB.Menu mnuRecent 
            Caption         =   ""
            Index           =   2
         End
         Begin VB.Menu mnuRecent 
            Caption         =   ""
            Index           =   3
         End
         Begin VB.Menu mnuRecent 
            Caption         =   ""
            Index           =   4
         End
         Begin VB.Menu mnuRecent 
            Caption         =   ""
            Index           =   5
         End
         Begin VB.Menu mnuRecent 
            Caption         =   ""
            Index           =   6
         End
         Begin VB.Menu mnuRecent 
            Caption         =   ""
            Index           =   7
         End
         Begin VB.Menu mnuRecent 
            Caption         =   ""
            Index           =   8
         End
         Begin VB.Menu mnuRecent 
            Caption         =   ""
            Index           =   9
         End
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuSep15 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCompile 
         Caption         =   "&Compile to pms"
      End
      Begin VB.Menu mnuCompileAs 
         Caption         =   "Compile to &pms As..."
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuSep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "&Export..."
      End
      Begin VB.Menu mnuImport 
         Caption         =   "&Import..."
      End
      Begin VB.Menu mnuSep18 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRunSoldat 
         Caption         =   "&Run Soldat"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Edit"
      Index           =   1
      Visible         =   0   'False
      Begin VB.Menu mnuUndo 
         Caption         =   "Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "Redo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuSep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDuplicate 
         Caption         =   "Duplicate"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clear"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuSep32 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuInvertSel 
         Caption         =   "Invert Selection"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuDeselect 
         Caption         =   "Deselect"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuSelColor 
         Caption         =   "Select by Color"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArrange 
         Caption         =   "Arrange"
         Begin VB.Menu mnuBringToFront 
            Caption         =   "Bring to Front"
         End
         Begin VB.Menu mnuBringForward 
            Caption         =   "Bring Forward"
         End
         Begin VB.Menu mnuSendBackward 
            Caption         =   "Send Backward"
         End
         Begin VB.Menu mnuSendToBack 
            Caption         =   "Send to Back"
         End
      End
      Begin VB.Menu mnuSep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSplit 
         Caption         =   "Split at Vertex"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuJoinVertices 
         Caption         =   "Join Vertices"
         Shortcut        =   ^J
      End
      Begin VB.Menu mnuSnapSelected 
         Caption         =   "Snap Selected Vertices"
      End
      Begin VB.Menu mnuCreate 
         Caption         =   "Create with Selected"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuTransform 
         Caption         =   "Transform"
         Begin VB.Menu mnuRotate 
            Caption         =   "Rotate 180°"
            Index           =   0
         End
         Begin VB.Menu mnuRotate 
            Caption         =   "Rotate 90° CW"
            Index           =   1
         End
         Begin VB.Menu mnuRotate 
            Caption         =   "Rotate 90° CCW"
            Index           =   2
         End
         Begin VB.Menu mnuSep6 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFlip 
            Caption         =   "Flip Horizontal"
            Index           =   0
         End
         Begin VB.Menu mnuFlip 
            Caption         =   "Flip Vertical"
            Index           =   1
         End
      End
      Begin VB.Menu mnuSep20 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSever 
         Caption         =   "Sever Connections"
      End
      Begin VB.Menu mnuSep16 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClrSketch 
         Caption         =   "Clear sketch"
      End
      Begin VB.Menu mnuSep30 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMap 
         Caption         =   "Map Settings..."
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuSep21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPreferences 
         Caption         =   "Preferences..."
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Texture"
      Index           =   2
      Visible         =   0   'False
      Begin VB.Menu mnuFixTexture 
         Caption         =   "Fix Texture"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuUntexture 
         Caption         =   "Untexture"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuTransformTexture 
         Caption         =   "Transform Texture"
         Begin VB.Menu mnuRotateTexture 
            Caption         =   "Rotate 180°"
            Index           =   0
         End
         Begin VB.Menu mnuRotateTexture 
            Caption         =   "Rotate 90° CW"
            Index           =   1
         End
         Begin VB.Menu mnuRotateTexture 
            Caption         =   "Rotate 90° CCW"
            Index           =   2
         End
         Begin VB.Menu mnuSep31 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFlipTexture 
            Caption         =   "Flip Horizontal"
            Index           =   0
         End
         Begin VB.Menu mnuFlipTexture 
            Caption         =   "Flip Vertical"
            Index           =   1
         End
      End
      Begin VB.Menu mnuSep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAverage 
         Caption         =   "Average Vertex Colors"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuApplyLight 
         Caption         =   "Apply Light to Vertices"
      End
      Begin VB.Menu mnuSep17 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFixedTexture 
         Caption         =   "Fixed Texture"
      End
      Begin VB.Menu mnuCustomX 
         Caption         =   "User Defined X"
      End
      Begin VB.Menu mnuCustomY 
         Caption         =   "User Defined Y"
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "View"
      Index           =   3
      Visible         =   0   'False
      Begin VB.Menu mnuZoomIn 
         Caption         =   "Zoom In"
      End
      Begin VB.Menu mnuZoomOut 
         Caption         =   "Zoom Out"
      End
      Begin VB.Menu mnuFitOnScreen 
         Caption         =   "Fit on Screen"
      End
      Begin VB.Menu mnuActualPixels 
         Caption         =   "Actual Size"
      End
      Begin VB.Menu mnuResetView 
         Caption         =   "Reset View"
      End
      Begin VB.Menu mnuSep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGrid 
         Caption         =   "Show Grid"
      End
      Begin VB.Menu mnuSnapToGrid 
         Caption         =   "Snap to Grid"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSnapToVerts 
         Caption         =   "Snap to Vertices"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSep13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBlendWireframe 
         Caption         =   "Blend Wireframe"
      End
      Begin VB.Menu mnuBlendPolys 
         Caption         =   "Blend Polys"
      End
      Begin VB.Menu mnuShowSceneryLayers 
         Caption         =   "Show Scenery Layers"
         Begin VB.Menu mnuShowSceneryLayer 
            Caption         =   "Back"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuShowSceneryLayer 
            Caption         =   "Middle"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu mnuShowSceneryLayer 
            Caption         =   "Front"
            Checked         =   -1  'True
            Index           =   2
         End
      End
      Begin VB.Menu mnuSep14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefreshBG 
         Caption         =   "Refresh"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Window"
      Index           =   4
      Visible         =   0   'False
      Begin VB.Menu mnuWorkspace 
         Caption         =   "Workspace"
         Begin VB.Menu mnuLoadSpace 
            Caption         =   "Load Workspace..."
         End
         Begin VB.Menu mnuSaveSpace 
            Caption         =   "Save Workspace..."
         End
         Begin VB.Menu mnuResetWindows 
            Caption         =   "Reset Window Locations"
         End
      End
      Begin VB.Menu mnuShowAll 
         Caption         =   "Show All"
      End
      Begin VB.Menu mnuHideAll 
         Caption         =   "Hide All"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTools 
         Caption         =   "Tools"
      End
      Begin VB.Menu mnuDisplay 
         Caption         =   "Display"
      End
      Begin VB.Menu mnuPalette 
         Caption         =   "Palette"
      End
      Begin VB.Menu mnuWaypoints 
         Caption         =   "Waypoints"
      End
      Begin VB.Menu mnuScenery 
         Caption         =   "Scenery"
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "Properties"
      End
      Begin VB.Menu mnuTexture 
         Caption         =   "Texture"
      End
   End
   Begin VB.Menu mnuObjects 
      Caption         =   "Objects"
      Visible         =   0   'False
      Begin VB.Menu mnuSpawn 
         Caption         =   "Player Spawn"
         Index           =   0
      End
      Begin VB.Menu mnuSpawn 
         Caption         =   "Alpha Team"
         Index           =   1
      End
      Begin VB.Menu mnuSpawn 
         Caption         =   "Bravo Team"
         Index           =   2
      End
      Begin VB.Menu mnuSpawn 
         Caption         =   "Charlie Team"
         Index           =   3
      End
      Begin VB.Menu mnuSpawn 
         Caption         =   "Delta Team"
         Index           =   4
      End
      Begin VB.Menu mnuSpawn 
         Caption         =   "Alpha Flag"
         Index           =   5
      End
      Begin VB.Menu mnuSpawn 
         Caption         =   "Bravo Flag"
         Index           =   6
      End
      Begin VB.Menu mnuSpawn 
         Caption         =   "Grenade Kit"
         Index           =   7
      End
      Begin VB.Menu mnuSpawn 
         Caption         =   "Medikit"
         Index           =   8
      End
      Begin VB.Menu mnuSpawn 
         Caption         =   "Cluster Grenades"
         Index           =   9
      End
      Begin VB.Menu mnuSpawn 
         Caption         =   "Vest"
         Index           =   10
      End
      Begin VB.Menu mnuSpawn 
         Caption         =   "Flame"
         Index           =   11
      End
      Begin VB.Menu mnuSpawn 
         Caption         =   "Berserker"
         Index           =   12
      End
      Begin VB.Menu mnuSpawn 
         Caption         =   "Predator"
         Index           =   13
      End
      Begin VB.Menu mnuSpawn 
         Caption         =   "Point Match Flag"
         Index           =   14
      End
      Begin VB.Menu mnuSpawn 
         Caption         =   "Rambo Bow"
         Index           =   15
      End
      Begin VB.Menu mnuSpawn 
         Caption         =   "Stat Gun"
         Index           =   16
      End
      Begin VB.Menu mnuSepObj 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCollider 
         Caption         =   "Collider"
      End
      Begin VB.Menu mnuSepObj2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGostek 
         Caption         =   "Gostek"
      End
   End
   Begin VB.Menu mnuPolyTypes 
      Caption         =   "Polygon Types"
      Visible         =   0   'False
      Begin VB.Menu mnuPolyType 
         Caption         =   "Normal"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuPolyType 
         Caption         =   "Only Bullets Collide"
         Index           =   1
      End
      Begin VB.Menu mnuPolyType 
         Caption         =   "Only Player Collides"
         Index           =   2
      End
      Begin VB.Menu mnuPolyType 
         Caption         =   "Doesn't Collide"
         Index           =   3
      End
      Begin VB.Menu mnuPolyType 
         Caption         =   "Ice"
         Index           =   4
      End
      Begin VB.Menu mnuPolyType 
         Caption         =   "Deadly"
         Index           =   5
      End
      Begin VB.Menu mnuPolyType 
         Caption         =   "Bloody Deadly"
         Index           =   6
      End
      Begin VB.Menu mnuPolyType 
         Caption         =   "Hurts"
         Index           =   7
      End
      Begin VB.Menu mnuPolyType 
         Caption         =   "Regenerates"
         Index           =   8
      End
      Begin VB.Menu mnuPolyType 
         Caption         =   "Lava"
         Index           =   9
      End
      Begin VB.Menu mnuPolyType 
         Caption         =   "Red Bullets Collides"
         Index           =   10
      End
      Begin VB.Menu mnuPolyType 
         Caption         =   "Red Players Collide"
         Index           =   11
      End
      Begin VB.Menu mnuPolyType 
         Caption         =   "Blue Bullets Collides"
         Index           =   12
      End
      Begin VB.Menu mnuPolyType 
         Caption         =   "Blue Players Collide"
         Index           =   13
      End
      Begin VB.Menu mnuPolyType 
         Caption         =   "Yellow Bullets Collides"
         Index           =   14
      End
      Begin VB.Menu mnuPolyType 
         Caption         =   "Yellow Players Collide"
         Index           =   15
      End
      Begin VB.Menu mnuPolyType 
         Caption         =   "Green Bullets Collides"
         Index           =   16
      End
      Begin VB.Menu mnuPolyType 
         Caption         =   "Green Players Collide"
         Index           =   17
      End
      Begin VB.Menu mnuPolyType 
         Caption         =   "Bouncy"
         Index           =   18
      End
      Begin VB.Menu mnuPolyType 
         Caption         =   "Explosive"
         Index           =   19
      End
      Begin VB.Menu mnuPolyType 
         Caption         =   "Hurts Flaggers"
         Index           =   20
      End
      Begin VB.Menu mnuPolyType 
         Caption         =   "Flagger Collides"
         Index           =   21
      End
      Begin VB.Menu mnuPolyType 
         Caption         =   "Non-Flagger Collides"
         Index           =   22
      End
      Begin VB.Menu mnuPolyType 
         Caption         =   "Flag Collides"
         Index           =   23
      End
      Begin VB.Menu mnuPolyType 
         Caption         =   "Background"
         Index           =   24
      End
      Begin VB.Menu mnuPolyType 
         Caption         =   "Background Transition"
         Index           =   25
      End
      Begin VB.Menu mnuSep19 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuad 
         Caption         =   "Textured Quad"
      End
   End
   Begin VB.Menu mnuMove 
      Caption         =   "Move"
      Visible         =   0   'False
      Begin VB.Menu mnuSetRCenter 
         Caption         =   "Set Reference Point"
      End
      Begin VB.Menu mnuCenterRCenter 
         Caption         =   "Center Reference Point"
      End
      Begin VB.Menu mnuFixedRCenter 
         Caption         =   "Fixed Reference Point"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuWaypoint 
      Caption         =   "Waypoint"
      Visible         =   0   'False
      Begin VB.Menu mnuWayType 
         Caption         =   "Left"
         Index           =   0
      End
      Begin VB.Menu mnuWayType 
         Caption         =   "Right"
         Index           =   1
      End
      Begin VB.Menu mnuWayType 
         Caption         =   "Up"
         Index           =   2
      End
      Begin VB.Menu mnuWayType 
         Caption         =   "Down"
         Index           =   3
      End
      Begin VB.Menu mnuWayType 
         Caption         =   "Fly"
         Index           =   4
      End
   End
   Begin VB.Menu mnuScen 
      Caption         =   "Scenery"
      Visible         =   0   'False
      Begin VB.Menu mnuScenTrans 
         Caption         =   "Rotate"
         Index           =   0
      End
      Begin VB.Menu mnuScenTrans 
         Caption         =   "Scale"
         Index           =   1
      End
      Begin VB.Menu mnuScenSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuScenLevel 
         Caption         =   "Back"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuScenLevel 
         Caption         =   "Middle"
         Index           =   1
      End
      Begin VB.Menu mnuScenLevel 
         Caption         =   "Front"
         Index           =   2
      End
   End
   Begin VB.Menu mnuScenTree 
      Caption         =   "Scenery Tree"
      Visible         =   0   'False
      Begin VB.Menu mnuScenList 
         Caption         =   "<list name>"
      End
      Begin VB.Menu mnuScenRemove 
         Caption         =   "Remove from List"
      End
   End
   Begin VB.Menu mnuVertexSelect 
      Caption         =   "VertexSelect"
      Visible         =   0   'False
      Begin VB.Menu mnuVSelDuplicate 
         Caption         =   "Duplicate"
      End
      Begin VB.Menu mnuVSelCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuVSelPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnuVSelClear 
         Caption         =   "Clear"
      End
      Begin VB.Menu mnuVSel0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVSelArrange 
         Caption         =   "Arrange"
         Begin VB.Menu mnuVSelBringToFront 
            Caption         =   "Bring To Front"
         End
         Begin VB.Menu mnuVSelBringForward 
            Caption         =   "Bring Forward"
         End
         Begin VB.Menu mnuVSelSendBackward 
            Caption         =   "Send Backward"
         End
         Begin VB.Menu mnuVSelSendToBack 
            Caption         =   "Send To Back"
         End
      End
      Begin VB.Menu mnuVSelTransform 
         Caption         =   "Transform"
         Begin VB.Menu mnuVSelRotate 
            Caption         =   "Rotate 180°"
            Index           =   0
         End
         Begin VB.Menu mnuVSelRotate 
            Caption         =   "Rotate 90° CW"
            Index           =   1
         End
         Begin VB.Menu mnuVSelRotate 
            Caption         =   "Rotate 90° CCW"
            Index           =   2
         End
         Begin VB.Menu mnuVSelSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuVSelFlip 
            Caption         =   "Flip Horizontal"
            Index           =   0
         End
         Begin VB.Menu mnuVSelFlip 
            Caption         =   "Flip Vertical"
            Index           =   1
         End
      End
   End
End
Attribute VB_Name = "frmSoldatMapEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Fix vb6 ide casing changes
#If False Then
    Public FileName, color, token, A, R, G, B, commonDialog, value, Val, X, Y, Z, Left, hWnd, Mid, Right, BackColor, bottom
    'Public FileName, color, token, A, R, G, B, commonDialog, value, Val, X, Y, Z, Left, hWnd, Mid, Right, BackColor
#End If


Public backClr As Long
Public pointColor As Long
Public selectionColor As Long
Public gridColor1 As Long
Public gridColor2 As Long
Public polyBlendSrc As Long
Public polyBlendDest As Long
Public wireBlendSrc As Long
Public wireBlendDest As Long
Public soldatDir As String
Public uncompDir As String
Public prefabDir As String
Public gridSpacing As Integer
Public gridDivisions As Integer
Public gridOp1 As Byte
Public gridOp2 As Byte

Public sceneryVerts As Boolean
Public topmost As Boolean

Public formHeight As Integer
Public formWidth As Integer
Public formLeft As Integer
Public formTop As Integer

Public gMaxZoom As Single
Public gMinZoom As Single
Public gResetZoom As Single

Public gTextureFile As String

Public xTexture As Integer
Public yTexture As Integer

Public shiftDown As Boolean
Public ctrlDown As Boolean
Public altDown As Boolean


Private noRedraw As Boolean

Private DX As DirectX8
Private D3D As Direct3D8
Private D3DDevice As Direct3DDevice8
Private DI As DirectInput8
Private DIDevice As DirectInputDevice8
Private DIState As DIKEYBOARDSTATE

Private Const BUFFER_SIZE As Long = 10

Private hEvent As Long
Implements DirectXEvent8

Private D3DX As D3DX8
Private mapTexture As Direct3DTexture8
Private particleTexture As Direct3DTexture8
Private patternTexture As Direct3DTexture8
Private objectsTexture As Direct3DTexture8
Private lineTexture As Direct3DTexture8
Private pathTexture As Direct3DTexture8
Private rCenterTexture As Direct3DTexture8
Private sketchTexture As Direct3DTexture8

Private renderTarget As Direct3DTexture8
Private renderSurface As Direct3DSurface8
Private backBuffer As Direct3DSurface8

Private scenerySprite As D3DXSprite

Private Const COLOR_KEY As Long = &HFF00FF00

Private Const FVF As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE
Private Const FVF2 As Long = D3DFVF_XYZ

Private Version As Long
Private Polys() As TPolygon
Private PolyCoords() As TTriangle

Private Scenery() As TScenery
Private SceneryTextures() As TextureData

Private Spawns() As TSpawnPoint
Private Colliders() As TCollider
Private Waypoints() As TWaypoint
Private Connections() As TConnection

Private Lights() As TLightSource

Private Options As TOptions
Private mPolyCount As Long

Private sceneryCount As Long
Private sceneryElements As Long
Private spawnPoints As Long
Private colliderCount As Long
Private waypointCount As Long
Private conCount As Integer
Private lightCount As Integer

Private mapTitle As String
Private maxX As Single
Private maxY As Single
Private minX As Single
Private minY As Single

Private bgPolys(1 To 4) As TCustomVertex
Private bgPolyCoords(1 To 4) As D3DVECTOR2
Private bgColors(1 To 2) As TColor

Private Const MAX_POLYS  As Integer = 4000

Private Const TOOL_MOVE As Byte = 0
Private Const TOOL_CREATE As Byte = 1
Private Const TOOL_VSELECT As Byte = 2
Private Const TOOL_PSELECT As Byte = 3
Private Const TOOL_VCOLOR As Byte = 4
Private Const TOOL_PCOLOR As Byte = 5
Private Const TOOL_TEXTURE As Byte = 6
Private Const TOOL_SCENERY As Byte = 7
Private Const TOOL_WAYPOINT As Byte = 8
Private Const TOOL_OBJECTS As Byte = 9
Private Const TOOL_COLORPICKER As Byte = 10
Private Const TOOL_SKETCH As Byte = 11
Private Const TOOL_LIGHTS As Byte = 12
Private Const TOOL_DEPTHMAP As Byte = 13

Private Const TOOL_HAND As Byte = 14
Private Const TOOL_VSELADD As Byte = 15
Private Const TOOL_VSELSUB As Byte = 16
Private Const TOOL_PSELADD As Byte = 17
Private Const TOOL_PSELSUB As Byte = 18
Private Const TOOL_SCALE As Byte = 19
Private Const TOOL_ROTATE As Byte = 20
Private Const TOOL_CONNECT As Byte = 21
Private Const TOOL_QUAD As Byte = 22
Private Const TOOL_PIXPICKER As Byte = 23
Private Const TOOL_LITPICKER As Byte = 24
Private Const TOOL_ERASER As Byte = 25
Private Const TOOL_SMUDGE As Byte = 26
Private Const TOOL_NULL As Byte = 255

Private Const KEY_SHIFT As Byte = 1
Private Const KEY_CTRL As Byte = 2
Private Const KEY_ALT As Byte = 4

Private sketch() As TSketchLine
Private sketchLines As Integer
Private selectedSketch(1 To 2) As Integer

Private circleOn As Boolean
Private leftMouseDown As Boolean

Private initialized As Boolean
Private initialized2 As Boolean
Private acquired As Boolean
Private selectionChanged As Boolean

Private clrPolys As Boolean
Private clrWireframe As Boolean
Private sslBack As Boolean
Private sslMid As Boolean
Private sslFront As Boolean

Public opacity As Single
Public blendMode As Integer

Private scrollCoords(1 To 2) As D3DVECTOR2    ' coordinates for scrolling
Private mouseCoords As D3DVECTOR2             ' coordinates of mouse
Private moveCoords(1 To 2) As D3DVECTOR2      ' coordinates for moving vertices
Private selectedCoords(1 To 2) As D3DVECTOR2  ' coordinates of selected area
Private selectedPolys() As Integer            ' list of selected polys and verts
Private vertexList() As TVertexData           ' list of polys with selected verts
Private numVerts As Integer                   ' number of current vertex being created
Private numCorners As Integer                 ' number of corner of scenery being created

Private numSelectedPolys As Integer
Private numSelectedScenery As Integer         ' number of currently selected scenery
Private numSelColliders As Integer
Private numSelSpawns As Integer
Private numSelWaypoints As Integer
Private numSelLights As Integer

Private creatingQuad As Boolean

Private currentFileName As String
Private prompt As Boolean

Private toolAction As Boolean
Private spaceDown As Boolean

Private currentScenery As String

Private zoomFactor As Single
Private pointRadius As Integer
Public snapRadius As Integer
Public clrRadius As Integer
Public ohSnap As Boolean
Public snapToGrid As Boolean
Public fixedTexture As Boolean
Public showBG As Boolean
Public showPolys As Boolean
Public showTexture As Boolean
Public showWireframe As Boolean
Public showPoints As Boolean
Public showScenery As Boolean
Public showObjects As Boolean
Public showGrid As Boolean
Public showWaypoints As Boolean
Private showPath1 As Boolean
Private showPath2 As Boolean
Public showSketch As Boolean
Public showLights As Boolean
Public currentTool As Byte
Private currentFunction As Byte
Private particleSize As Single
Public colorMode As Byte
Private eraseCircle As Boolean
Private eraseLines As Boolean

Private polyType As Byte

Private rCenter As D3DVECTOR2
Private selRect(3) As D3DVECTOR2 ' RECT

Private xGridLines() As TLine
Private yGridLines() As TLine
Private inc As Single

Private scaleDiff As D3DVECTOR2
Private rDiff As Single

Private gostek As D3DVECTOR2

Private imageInfo As TImageInfo
Private textureDesc As D3DSURFACE_DESC

Private noneSelected As Boolean

Private currentUndo As Integer
Private numUndo As Integer
Private numRedo As Integer
Public maxUndo As Integer
Private lastCompiled As String

Private currentWaypoint As Integer

Private objTexSize As D3DVECTOR2

Private mIsResizingWindow As Boolean
Private mMouseStartPosX As Long
Private mMouseStartPosY As Long
Private mInitialWindowWidth As Single
Private mInitialWindowHeight As Single

Private mPrevWidth As Long
Private mPrevHeight As Long
Private mPrevLeft As Long
Private mPrevTop As Long

Private Const QUICK_MOVE_DELTA = 90000

Private Declare Function MoveWindow Lib "user32" ( _
    ByVal hWnd As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal bRepaint As Long) As Long

Private Const SPI_GETWORKAREA = 48
Private Declare Function SystemParametersInfo& Lib "user32" Alias "SystemParametersInfoA" ( _
    ByVal uAction As Long, _
    ByVal uParam As Long, lpvParam As Any, _
    ByVal fuWinIni As Long)

' Global constants
Property Get MIN_FORM_WIDTH() As Integer
    MIN_FORM_WIDTH = 300
End Property

Property Get MIN_FORM_HEIGHT() As Integer
    MIN_FORM_HEIGHT = 200
End Property


Private Function QuickHide(ByRef myWindow As Form)
    MoveWindow myWindow.hWnd, _
        (myWindow.Left - QUICK_MOVE_DELTA) / Screen.TwipsPerPixelX, _
        (myWindow.Top - QUICK_MOVE_DELTA) / Screen.TwipsPerPixelY, _
        myWindow.Width / Screen.TwipsPerPixelX, _
        myWindow.Height / Screen.TwipsPerPixelY, _
        False
End Function

Private Function QuickMoveAndShow(ByRef myWindow As Form, nLeft, nTop)
    myWindow.Move nLeft + QUICK_MOVE_DELTA, nTop + QUICK_MOVE_DELTA
End Function

Private Sub Form_Load()

    On Error GoTo ErrorHandler

    Dim i As Integer
    Dim temp As String
    Dim err As String
    
    Dim prevMousePointer As Integer

    initialized = False

    modConfig.loadSettings
    loadColors


    err = "Error setting colors"
    Me.SetColors
    Me.Show
    Me.Tag = vbNormal


    err = "Error setting directories"
    If Len(Dir$(uncompDir, vbDirectory)) = 0 Or uncompDir = "" Then
        uncompDir = appPath & "\Maps\"
    End If

    If Len(Dir$(prefabDir, vbDirectory)) = 0 Or prefabDir = "" Then
        prefabDir = appPath & "\Prefabs\"
    End If

    ' if given directory doesn't exist, change to default
    If Len(Dir$(soldatDir & "Textures\", vbDirectory)) = 0 Or soldatDir = "" Then
        temp = GetSoldatDir
        If temp <> "" Then
            soldatDir = temp
            temp = ""
        End If
    End If

    frmTools.initTool currentTool

    initGfx


    err = "Error loading cursors"
    loadCursors


    err = "Error initializing values"

    ' init values
    scrollCoords(2).X = -Me.ScaleWidth / 2
    scrollCoords(2).Y = -Me.ScaleHeight / 2
    pointRadius = 4
    particleSize = pointRadius * 2
    zoomFactor = 1
    scaleDiff.X = 1
    scaleDiff.Y = 1
    sslBack = True
    sslMid = True
    sslFront = True

    gPolyTypeClrs(0) = selectionColor

    ReDim Scenery(0)
    ReDim Preserve SceneryTextures(0)
    ReDim Spawns(0)
    ReDim Colliders(0)

    ReDim sketch(0)

    sketch(0).vertex(1).Z = 1
    sketch(0).vertex(2).Z = 1

    Colliders(0).radius = clrRadius


    err = "Error initializing color picker"

    frmColor.picSpectrum.Cls
    frmColor.InitColor gPolyClr.red, gPolyClr.green, gPolyClr.blue


    err = "Error setting current tool icon (" & currentTool & ")"

    currentFunction = currentTool


    err = "Error initializing grid"
    initGrid


    err = "Error initializing D3D"
    initialized2 = False
    loadWorkspace "current.ini", True
    Init


    err = "Error initializing DInput"
    InitDInput


    err = "Error setting up palette windows"

    loadWorkspace
    initGrid

    ' show windows
    frmTaskBar.Show
    frmTools.Show 0, frmSoldatMapEditor
    frmPalette.Show 0, frmSoldatMapEditor
    frmDisplay.Show 0, frmSoldatMapEditor
    frmWaypoints.Show 0, frmSoldatMapEditor
    frmScenery.Show 0, frmSoldatMapEditor
    frmInfo.Show 0, frmSoldatMapEditor
    frmTexture.Show 0, frmSoldatMapEditor

    ' set window settings
    frmDisplay.Visible = mnuDisplay.Checked
    frmWaypoints.Visible = mnuWaypoints.Checked
    frmPalette.Visible = mnuPalette.Checked
    frmTools.Visible = mnuTools.Checked
    frmScenery.Visible = mnuScenery.Checked
    frmInfo.Visible = mnuInfo.Checked
    frmTexture.Visible = mnuTexture.Checked

    frmPalette.refreshPalette clrRadius, opacity, blendMode, colorMode
    frmPalette.setValues gPolyClr.red, gPolyClr.green, gPolyClr.blue
    frmDisplay.setLayer 0, showBG
    frmDisplay.setLayer 1, showPolys
    frmDisplay.setLayer 2, showTexture
    frmDisplay.setLayer 3, showWireframe
    frmDisplay.setLayer 4, showPoints
    frmDisplay.setLayer 5, showScenery
    frmDisplay.setLayer 6, showObjects
    frmDisplay.setLayer 7, showWaypoints
    frmDisplay.setLayer 8, showGrid
    frmDisplay.setLayer 9, showLights
    frmDisplay.setLayer 10, showSketch

    mnuFixedTexture.Checked = fixedTexture
    mnuSnapToGrid.Checked = snapToGrid
    mnuSnapToVerts.Checked = ohSnap

    lblCurrentTool.Caption = frmSoldatMapEditor.ImageList.ListImages(currentFunction + 1).Tag

    frmSoldatMapEditor.commonDialog.Filter = "Map File (*.pms)|*.pms"
    commonDialog.Flags = cdlOFNOverwritePrompt Or cdlOFNPathMustExist Or cdlOFNFileMustExist


    err = "Error parsing command line args"

    temp = Command$
    If Right(temp, 1) = """" Then
        temp = Left(temp, Len(temp) - 1)
        temp = Right(temp, Len(temp) - 1)
    End If

    newMap
    If LCase$(Right(temp, 4)) = ".pms" Then
    
        prevMousePointer = Me.MousePointer
        Me.MousePointer = vbHourglass

        If Dir$(temp) <> "" Then
            LoadFile temp
        ElseIf Dir$(appPath & "\Maps\" & temp) <> "" Then
            LoadFile appPath & "\Maps\" & temp
        ElseIf Dir$(soldatDir & "Maps\" & temp) <> "" Then
            LoadFile soldatDir & "Maps\" & temp
        End If

        Me.MousePointer = prevMousePointer
    End If


    err = "Error acquiring input device"

    Me.SetFocus
    DIDevice.Acquire
    acquired = True

    Exit Sub

ErrorHandler:

    MsgBox "Error loading" & vbNewLine & err & vbNewLine & Error$

End Sub

Public Sub RestoreBorderLessForm()

    If Me.Tag = vbNormal Then Exit Sub

    Me.Tag = vbNormal
    Me.Move mPrevLeft, mPrevTop, mPrevWidth, mPrevHeight

End Sub

Public Sub MaximizeBorderLessForm()

    If Me.Tag = vbMaximized Then Exit Sub

    Dim ScreenWidth&, ScreenHeight&, ScreenLeft&, ScreenTop&
    Dim DesktopArea As RECT
    Call SystemParametersInfo(SPI_GETWORKAREA, 0, DesktopArea, 0)

    ScreenHeight = (DesktopArea.bottom - DesktopArea.Top) * Screen.TwipsPerPixelY
    ScreenWidth = (DesktopArea.Right - DesktopArea.Left) * Screen.TwipsPerPixelX
    ScreenLeft = DesktopArea.Left * Screen.TwipsPerPixelX
    ScreenTop = DesktopArea.Top * Screen.TwipsPerPixelY

    mPrevLeft = Me.Left
    mPrevTop = Me.Top
    mPrevWidth = Me.Width
    mPrevHeight = Me.Height

    Me.Tag = vbMaximized
    Me.Move ScreenLeft, ScreenTop, ScreenWidth, ScreenHeight

End Sub

Private Sub SetCursor(Index As Integer)

    On Error GoTo ErrorHandler

    Me.MouseIcon = frmSoldatMapEditor.ImageList.ListImages(Index).Picture

    Exit Sub

ErrorHandler:

    MsgBox "Error setting cursor" & vbNewLine & Error$

End Sub

Public Sub loadCursors()

    On Error GoTo ErrorHandler

    ImageList.ListImages.Clear

    ' load cursors
    ImageList.ListImages.Add TOOL_MOVE + 1, "move", LoadPicture(appPath & "\skins\" & gfxDir & "\cursors\move.cur")
    ImageList.ListImages.Add TOOL_CREATE + 1, "create", LoadPicture(appPath & "\skins\" & gfxDir & "\cursors\create.cur")
    ImageList.ListImages.Add TOOL_VSELECT + 1, "vselect", LoadPicture(appPath & "\skins\" & gfxDir & "\cursors\vselect.cur")
    ImageList.ListImages.Add TOOL_PSELECT + 1, "pselect", LoadPicture(appPath & "\skins\" & gfxDir & "\cursors\pselect.cur")
    ImageList.ListImages.Add TOOL_VCOLOR + 1, "vcolor", LoadPicture(appPath & "\skins\" & gfxDir & "\cursors\vcolor.cur")
    ImageList.ListImages.Add TOOL_PCOLOR + 1, "pcolor", LoadPicture(appPath & "\skins\" & gfxDir & "\cursors\pcolor.cur")
    ImageList.ListImages.Add TOOL_TEXTURE + 1, "texture", LoadPicture(appPath & "\skins\" & gfxDir & "\cursors\texture.cur")
    ImageList.ListImages.Add TOOL_SCENERY + 1, "scenery", LoadPicture(appPath & "\skins\" & gfxDir & "\cursors\scenery.cur")
    ImageList.ListImages.Add TOOL_WAYPOINT + 1, "waypoint", LoadPicture(appPath & "\skins\" & gfxDir & "\cursors\waypoint.cur")
    ImageList.ListImages.Add TOOL_OBJECTS + 1, "objects", LoadPicture(appPath & "\skins\" & gfxDir & "\cursors\objects.cur")
    ImageList.ListImages.Add TOOL_COLORPICKER + 1, "clrpicker", LoadPicture(appPath & "\skins\" & gfxDir & "\cursors\clrpicker.cur")
    ImageList.ListImages.Add TOOL_SKETCH + 1, "sketch", LoadPicture(appPath & "\skins\" & gfxDir & "\cursors\sketch.cur")
    ImageList.ListImages.Add TOOL_LIGHTS + 1, "lights", LoadPicture(appPath & "\skins\" & gfxDir & "\cursors\light.cur")
    ImageList.ListImages.Add TOOL_DEPTHMAP + 1, "depthmap", LoadPicture(appPath & "\skins\" & gfxDir & "\cursors\depthmap.cur")

    ImageList.ListImages.Add TOOL_HAND + 1, "hand", LoadPicture(appPath & "\skins\" & gfxDir & "\cursors\hand.cur")
    ImageList.ListImages.Add TOOL_VSELADD + 1, "vseladd", LoadPicture(appPath & "\skins\" & gfxDir & "\cursors\vseladd.cur")
    ImageList.ListImages.Add TOOL_VSELSUB + 1, "vselsub", LoadPicture(appPath & "\skins\" & gfxDir & "\cursors\vselsub.cur")
    ImageList.ListImages.Add TOOL_PSELADD + 1, "pseladd", LoadPicture(appPath & "\skins\" & gfxDir & "\cursors\pseladd.cur")
    ImageList.ListImages.Add TOOL_PSELSUB + 1, "pselsub", LoadPicture(appPath & "\skins\" & gfxDir & "\cursors\pselsub.cur")
    ImageList.ListImages.Add TOOL_SCALE + 1, "scale", LoadPicture(appPath & "\skins\" & gfxDir & "\cursors\scale.cur")
    ImageList.ListImages.Add TOOL_ROTATE + 1, "rotate", LoadPicture(appPath & "\skins\" & gfxDir & "\cursors\rotate.cur")
    ImageList.ListImages.Add TOOL_CONNECT + 1, "connect", LoadPicture(appPath & "\skins\" & gfxDir & "\cursors\connect.cur")
    ImageList.ListImages.Add TOOL_QUAD + 1, "quad", LoadPicture(appPath & "\skins\" & gfxDir & "\cursors\quad.cur")
    ImageList.ListImages.Add TOOL_PIXPICKER + 1, "pixpicker", LoadPicture(appPath & "\skins\" & gfxDir & "\cursors\pixpicker.cur")
    ImageList.ListImages.Add TOOL_LITPICKER + 1, "litpicker", LoadPicture(appPath & "\skins\" & gfxDir & "\cursors\litpicker.cur")
    ImageList.ListImages.Add TOOL_ERASER + 1, "eraser", LoadPicture(appPath & "\skins\" & gfxDir & "\cursors\eraser.cur")
    ImageList.ListImages.Add TOOL_SMUDGE + 1, "smudge", LoadPicture(appPath & "\skins\" & gfxDir & "\cursors\smudge.cur")

    ImageList.ListImages.Item(TOOL_MOVE + 1).Tag = "Move Selection"
    ImageList.ListImages.Item(TOOL_CREATE + 1).Tag = "Create Polygons"
    ImageList.ListImages.Item(TOOL_VSELECT + 1).Tag = "Select Vertices"
    ImageList.ListImages.Item(TOOL_PSELECT + 1).Tag = "Select Polygons"
    ImageList.ListImages.Item(TOOL_VCOLOR + 1).Tag = "Color Vertices"
    ImageList.ListImages.Item(TOOL_PCOLOR + 1).Tag = "Color Polygons"
    ImageList.ListImages.Item(TOOL_TEXTURE + 1).Tag = "Transform Texture"
    ImageList.ListImages.Item(TOOL_SCENERY + 1).Tag = "Create Scenery"
    ImageList.ListImages.Item(TOOL_WAYPOINT + 1).Tag = "Create Waypoints"
    ImageList.ListImages.Item(TOOL_OBJECTS + 1).Tag = "Place Spawn Points or Colliders"
    ImageList.ListImages.Item(TOOL_COLORPICKER + 1).Tag = "Pick a Vertex Color"
    ImageList.ListImages.Item(TOOL_SKETCH + 1).Tag = "Sketch"
    ImageList.ListImages.Item(TOOL_LIGHTS + 1).Tag = "Create Lights"
    ImageList.ListImages.Item(TOOL_DEPTHMAP + 1).Tag = "Edit Depth Map"

    ImageList.ListImages.Item(TOOL_HAND + 1).Tag = "Scroll Map"
    ImageList.ListImages.Item(TOOL_VSELADD + 1).Tag = "Add to Selection"
    ImageList.ListImages.Item(TOOL_VSELSUB + 1).Tag = "Subtract from Selection"
    ImageList.ListImages.Item(TOOL_PSELADD + 1).Tag = "Add to Selection"
    ImageList.ListImages.Item(TOOL_PSELSUB + 1).Tag = "Subtract from Selection"
    ImageList.ListImages.Item(TOOL_SCALE + 1).Tag = "Scale Selection"
    ImageList.ListImages.Item(TOOL_ROTATE + 1).Tag = "Rotate Selection"
    ImageList.ListImages.Item(TOOL_CONNECT + 1).Tag = "Connect Waypoints"
    ImageList.ListImages.Item(TOOL_QUAD + 1).Tag = "Create Quad"
    ImageList.ListImages.Item(TOOL_PIXPICKER + 1).Tag = "Pick a pixel color"
    ImageList.ListImages.Item(TOOL_LITPICKER + 1).Tag = "Pick a Lit Vertex Color"
    ImageList.ListImages.Item(TOOL_ERASER + 1).Tag = "Erase Lines"
    ImageList.ListImages.Item(TOOL_SMUDGE + 1).Tag = "Move Lines"

    Exit Sub

ErrorHandler:

    MsgBox "Error loading cursors" & vbNewLine & Error$

End Sub

Public Sub initGfx()

    Dim i As Integer
    Dim c As Control

    picTitle.Picture = LoadPicture(appPath & "\skins\" & gfxDir & "\titlebar_main.bmp")
    If FileExists(appPath & "\skins\" & gfxDir & "\resize.bmp") Then
        picResize.Picture = LoadPicture(appPath & "\skins\" & gfxDir & "\resize.bmp")
    Else
        picResize.Picture = Nothing
    End If

    picGfx.Picture = LoadPicture(appPath & "\skins\" & gfxDir & "\tool_gfx.bmp")
    picButtonGfx.Picture = LoadPicture(appPath & "\skins\" & gfxDir & "\button_gfx.bmp")

    ' draw control box buttons
    mouseEvent2 picExit, 0, 0, BUTTON_SMALL, 0, BUTTON_UP
    mouseEvent2 picMaximize, 0, 0, BUTTON_SMALL, (Me.Tag = vbNormal), BUTTON_UP
    mouseEvent2 picMinimize, 0, 0, BUTTON_SMALL, 0, BUTTON_UP
    mouseEvent2 picHelp, 0, 0, BUTTON_SMALL, 0, BUTTON_UP

    ' draw menu buttons
    For Each c In picMenu
        mouseEvent2 c, 0, 0, BUTTON_MENU, 0, BUTTON_UP
    Next

End Sub

Private Sub centerView()

    Dim i As Integer

    If mPolyCount > 0 Then
        For i = 1 To mPolyCount
            Polys(i).vertex(1).X = (PolyCoords(i).vertex(1).X - scrollCoords(2).X) * zoomFactor
            Polys(i).vertex(1).Y = (PolyCoords(i).vertex(1).Y - scrollCoords(2).Y) * zoomFactor
            Polys(i).vertex(2).X = (PolyCoords(i).vertex(2).X - scrollCoords(2).X) * zoomFactor
            Polys(i).vertex(2).Y = (PolyCoords(i).vertex(2).Y - scrollCoords(2).Y) * zoomFactor
            Polys(i).vertex(3).X = (PolyCoords(i).vertex(3).X - scrollCoords(2).X) * zoomFactor
            Polys(i).vertex(3).Y = (PolyCoords(i).vertex(3).Y - scrollCoords(2).Y) * zoomFactor
        Next
    End If

    For i = 1 To 4
        bgPolys(i).X = bgPolyCoords(i).X - scrollCoords(2).X * zoomFactor
        bgPolys(i).Y = bgPolyCoords(i).Y - scrollCoords(2).Y * zoomFactor
    Next

End Sub

Public Sub Init()

    On Error GoTo ErrorHandler

    initialized = False
    noRedraw = False
    selectionChanged = False

    Dim DispMode As D3DDISPLAYMODE
    Dim D3DWindow As D3DPRESENT_PARAMETERS
    Dim debugVal As String


    debugVal = "Error creating Direct3D objects"

    If Not initialized2 Then
        Set D3DX = New D3DX8
        Set DX = New DirectX8
        Set D3D = DX.Direct3DCreate()
        initialized2 = True
    End If


    debugVal = "Error getting display mode"

    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    D3DWindow.Windowed = 1
    D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY
    D3DWindow.BackBufferFormat = D3DFMT_A8R8G8B8


    debugVal = "Error creating D3D device"

    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Me.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow) ' Main screen turn on.


    debugVal = "Error setting render states"

    D3DDevice.SetVertexShader FVF
    D3DDevice.SetRenderState D3DRS_LIGHTING, False

    D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE  ' polys that are ccw

    D3DDevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
    D3DDevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    D3DDevice.SetRenderState D3DRS_POINTSIZE, FtoDW(particleSize)

    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG2, D3DTA_DIFFUSE

    Set renderTarget = D3DX.CreateTexture(D3DDevice, 256, 256, D3DX_DEFAULT, D3DUSAGE_RENDERTARGET, D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT)
    Set renderSurface = renderTarget.GetSurfaceLevel(0)
    Set backBuffer = D3DDevice.GetBackBuffer(0, D3DBACKBUFFER_TYPE_MONO)


    debugVal = "Error creating pattern texture"

    Set patternTexture = D3DX.CreateTextureFromFile(D3DDevice, appPath & "\skins\" & gfxDir & "\pattern.bmp")


    debugVal = "Error creating objects texture"

    Set objectsTexture = D3DX.CreateTextureFromFileEx(D3DDevice, appPath & "\skins\" & gfxDir & "\objects.bmp", D3DX_DEFAULT, D3DX_DEFAULT, _
            D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, _
            D3DX_FILTER_POINT, COLOR_KEY, ByVal 0, ByVal 0)

    objectsTexture.GetLevelDesc 0, textureDesc

    objTexSize.X = textureDesc.Width
    objTexSize.Y = textureDesc.Height


    debugVal = "Error creating scenery not found texture"

    Set SceneryTextures(0).Texture = D3DX.CreateTextureFromFileEx(D3DDevice, appPath & "\skins\" & gfxDir & "\notfound.bmp", D3DX_DEFAULT, D3DX_DEFAULT, _
            D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
            D3DX_FILTER_POINT, COLOR_KEY, ByVal 0, ByVal 0)

    SceneryTextures(0).Texture.GetLevelDesc 0, textureDesc

    SceneryTextures(0).Width = imageInfo.Width
    SceneryTextures(0).Height = imageInfo.Height

    SceneryTextures(0).reScale.X = SceneryTextures(0).Width / textureDesc.Width
    SceneryTextures(0).reScale.Y = SceneryTextures(0).Height / textureDesc.Height

    If SceneryTextures(0).reScale.X = 0 Or SceneryTextures(0).reScale.Y = 0 Then
        SceneryTextures(0).reScale.X = 1
        SceneryTextures(0).reScale.Y = 1
    End If


    debugVal = "Error creating line texture"

    Set lineTexture = D3DX.CreateTextureFromFileEx(D3DDevice, appPath & "\skins\" & gfxDir & "\lines.bmp", D3DX_DEFAULT, D3DX_DEFAULT, _
            D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
            D3DX_FILTER_POINT, COLOR_KEY, ByVal 0, ByVal 0)


    debugVal = "Error creating path texture"

    Set pathTexture = D3DX.CreateTextureFromFileEx(D3DDevice, appPath & "\skins\" & gfxDir & "\path.png", D3DX_DEFAULT, D3DX_DEFAULT, _
            D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
            D3DX_FILTER_POINT, COLOR_KEY, ByVal 0, ByVal 0)


    debugVal = "Error creating rotation center texture"

    Set rCenterTexture = D3DX.CreateTextureFromFileEx(D3DDevice, appPath & "\skins\" & gfxDir & "\rcenter.bmp", D3DX_DEFAULT, D3DX_DEFAULT, _
            D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
            D3DX_FILTER_POINT, COLOR_KEY, ByVal 0, ByVal 0)


    debugVal = "Error creating sketch texture"

    Set sketchTexture = D3DX.CreateTextureFromFileEx(D3DDevice, appPath & "\skins\" & gfxDir & "\sketch.bmp", D3DX_DEFAULT, D3DX_DEFAULT, _
            D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
            D3DX_FILTER_POINT, COLOR_KEY, ByVal 0, ByVal 0)


    debugVal = "Error creating scenery sprite"

    Set scenerySprite = D3DX.CreateSprite(D3DDevice)


    debugVal = "Error creating particle texture"

    Set particleTexture = D3DX.CreateTextureFromFileEx(D3DDevice, appPath & "\skins\" & gfxDir & "\vertex8x8.bmp", 8, 8, _
            D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
            D3DX_FILTER_POINT, COLOR_KEY, ByVal 0, ByVal 0)

    initialized = True

    Exit Sub

ErrorHandler:

    If D3DX Is Nothing Then
        MsgBox "Direct3D initialization failed" & vbNewLine & debugVal & vbNewLine & Error$
    Else
        MsgBox "Direct3D initialization failed" & vbNewLine & D3DX.GetErrorString(err.Number) & vbNewLine & debugVal
    End If

End Sub

Private Sub InitDInput()

    On Error GoTo ErrorHandler

    Dim debugVal As String

    Dim i As Long
    Dim DevProp As DIPROPLONG
    Dim DevInfo As DirectInputDeviceInstance8
    Dim pBuffer(0 To BUFFER_SIZE) As DIDEVICEOBJECTDATA


    debugVal = "Error creating DI device"

    Set DI = DX.DirectInputCreate
    Set DIDevice = DI.CreateDevice("GUID_SysKeyboard")


    debugVal = "Error setting DI device"

    DIDevice.SetCommonDataFormat DIFORMAT_KEYBOARD
    DIDevice.SetCooperativeLevel Me.hWnd, DISCL_NONEXCLUSIVE Or DISCL_FOREGROUND


    debugVal = "Error setting DI properties"

    DevProp.lHow = DIPH_DEVICE
    DevProp.lData = BUFFER_SIZE
    DIDevice.SetProperty DIPROP_BUFFERSIZE, DevProp


    debugVal = "Error setting DI device notification"

    hEvent = DX.CreateEvent(Me)
    DIDevice.SetEventNotification hEvent


    debugVal = "Error getting device info"

    Set DevInfo = DIDevice.GetDeviceInfo()


    debugVal = "Error acquiring device"

    Me.SetFocus
    DIDevice.Acquire
    acquired = True

    Exit Sub

ErrorHandler:

    If debugVal <> "Error acquiring device" Then
        MsgBox "DirectInput initialization failed" & vbNewLine & D3DX.GetErrorString(err.Number) & vbNewLine & debugVal
    End If

End Sub

Public Sub resetDevice()

    On Error GoTo ErrorHandler
    Dim i As Integer

    noRedraw = True
    If selectionChanged Then
        SaveUndo
        selectionChanged = False
    End If
    SaveUndo
    mnuSelectAll_Click
    deletePolys

    Set mapTexture = Nothing
    Set particleTexture = Nothing
    Set patternTexture = Nothing
    Set sketchTexture = Nothing
    Set lineTexture = Nothing
    Set pathTexture = Nothing
    Set rCenterTexture = Nothing
    Set D3DDevice = Nothing
    Init
    For i = 1 To frmScenery.lstScenery.ListCount
        RefreshSceneryTextures i
    Next

    setMapTexture gTextureFile

    initGrid

    initialized = True

    loadUndo False
    loadUndo False

    noRedraw = False

    Render

    Exit Sub

ErrorHandler:

    MsgBox "Error resetting device" & vbNewLine & D3DX.GetErrorString(err.Number)

End Sub

Public Sub RegainFocus()

    On Error Resume Next

    Me.SetFocus
    DIDevice.Acquire
    acquired = True
    ctrlDown = False
    altDown = False
    shiftDown = False
    SetCursor currentFunction + 1

End Sub

Public Sub newMap()

    Dim i As Integer

    On Error GoTo ErrorHandler

    prompt = False

    Version = 11

    commonDialog.FileName = ""

    numVerts = 0
    toolAction = False

    mapTitle = "New Soldat Map"

    Options.BackgroundColor = ARGB(255, RGB(224, 224, 224))
    Options.BackgroundColor2 = ARGB(255, RGB(32, 32, 32))

    Options.textureName(0) = 0
    Options.MapRandomID = 0
    Options.GrenadePacks = 5
    Options.Medikits = 5
    Options.StartJet = 190
    Options.Steps = 0
    Options.Weather = 0

    numSelectedPolys = 0
    ReDim selectedPolys(0)
    ReDim vertexList(0)

    mPolyCount = 0
    ReDim Polys(0)
    ReDim vertexList(0)
    ReDim PolyCoords(0)

    sceneryCount = 0
    ReDim Scenery(0)
    sceneryElements = 0
    ReDim Preserve SceneryTextures(0)
    frmScenery.lstScenery.Clear
    setCurrentScenery 0
    tvwScenery.Nodes.Remove "In Use"
    tvwScenery.Nodes.Add "Master List", tvwFirst, "In Use", "In Use"

    spawnPoints = 0
    colliderCount = 0
    ReDim Spawns(0)
    ReDim Colliders(0)
    Colliders(0).radius = clrRadius

    waypointCount = 0
    ReDim Waypoints(0)
    conCount = 0
    ReDim Connections(0)

    lightCount = 0
    ReDim Lights(0)

    sketchLines = 0
    ReDim Preserve sketch(0)

    bgColors(1) = makeColor(224, 224, 224)
    bgColors(2) = makeColor(32, 32, 32)

    maxX = 0
    maxY = 0
    minX = 0
    minY = 0

    bgPolys(1) = CreateCustomVertex(-640, -640, 1, 1, RGB(bgColors(1).blue, bgColors(1).green, bgColors(1).red), 0, 0)
    bgPolys(2) = CreateCustomVertex(-640, 640, 1, 1, RGB(bgColors(2).blue, bgColors(2).green, bgColors(2).red), 0, 0)
    bgPolys(3) = CreateCustomVertex(640, -640, 1, 1, RGB(bgColors(1).blue, bgColors(1).green, bgColors(1).red), 0, 0)
    bgPolys(4) = CreateCustomVertex(640, 640, 1, 1, RGB(bgColors(2).blue, bgColors(2).green, bgColors(2).red), 0, 0)

    For i = 1 To 4
        bgPolyCoords(i).X = bgPolys(i).X
        bgPolyCoords(i).Y = bgPolys(i).Y
    Next

    scrollCoords(1).X = 0
    scrollCoords(1).Y = 0
    scrollCoords(2).X = -Me.ScaleWidth / 2 - 1
    scrollCoords(2).Y = -Me.ScaleHeight / 2 - 1
    zoomFactor = 1

    setMapData

    txtZoom.Text = Int(zoomFactor * 1000 + 0.5) / 10 & "%"

    If Len(Dir(soldatDir & "Textures\" & gTextureFile)) <> 0 Then
        setMapTexture gTextureFile
        frmTexture.setTexture gTextureFile
    Else
        Set mapTexture = Nothing
    End If

    currentFileName = "Untitled.pms"
    lblFileName.Caption = "Untitled.pms"

    centerView

    numUndo = 0
    numRedo = 0
    currentUndo = 0
    SaveUndo

    Render

    Exit Sub

ErrorHandler:

    MsgBox "error creating new file" & vbNewLine & Error$

End Sub

Public Sub LoadFile(theFileName As String)

    On Error GoTo ErrorHandler

    Dim errorVal As String
    Dim fileOpen As Boolean

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim temp As Long
    Dim tempString As String

    Dim polyIndex As Integer
    Dim polysInSector As Integer

    Const SECTOR_NUM As Long = 25

    Dim Scenery_New As TMapFile_Scenery
    Dim newWaypoint As TNewWaypoint
    Dim Prop As TProp
    Dim spawn As TSaveSpawnPoint

    Dim toTGARes As Long

    prompt = False

    scrollCoords(1).X = 0
    scrollCoords(1).Y = 0
    scrollCoords(2).X = -Me.ScaleWidth / 2
    scrollCoords(2).Y = -Me.ScaleHeight / 2
    zoomFactor = 1
    toolAction = False
    numVerts = 0

    sceneryCount = 0
    sceneryElements = 0
    frmScenery.lstScenery.Clear
    tvwScenery.Nodes.Remove "In Use"
    tvwScenery.Nodes.Add "Master List", tvwFirst, "In Use", "In Use"
    numSelectedPolys = 0
    ReDim selectedPolys(numSelectedPolys)

    currentFileName = ""
    For i = 0 To Len(theFileName) - 1
        If Mid(theFileName, Len(theFileName) - i, 1) <> "\" Then
            currentFileName = Mid(theFileName, Len(theFileName) - i, 1) + currentFileName
        Else
            Exit For
        End If
    Next

    lblFileName.Caption = currentFileName

    Open theFileName For Binary Access Read Lock Read As #1

        fileOpen = True
        errorVal = "Error loading polys"

        maxX = 0
        maxY = 0
        minX = 0
        minY = 0

        Get #1, , Version
        Get #1, , Options
        Get #1, , mPolyCount
        ReDim Polys(0 To mPolyCount)
        ReDim PolyCoords(0 To mPolyCount)
        ReDim vertexList(0 To mPolyCount)

        For i = 1 To mPolyCount
            Get #1, , Polys(i)
            Get #1, , vertexList(i).polyType

            For j = 1 To 3
                PolyCoords(i).vertex(j).X = Polys(i).vertex(j).X
                PolyCoords(i).vertex(j).Y = Polys(i).vertex(j).Y
                vertexList(i).color(j) = getRGB(Polys(i).vertex(j).color)
                If PolyCoords(i).vertex(j).X > maxX Then maxX = PolyCoords(i).vertex(j).X
                If PolyCoords(i).vertex(j).X < minX Then minX = PolyCoords(i).vertex(j).X
                If PolyCoords(i).vertex(j).Y > maxY Then maxY = PolyCoords(i).vertex(j).Y
                If PolyCoords(i).vertex(j).Y < minY Then minY = PolyCoords(i).vertex(j).Y
                Polys(i).Perp.vertex(j).Z = Sqr(Polys(i).Perp.vertex(j).X ^ 2 + Polys(i).Perp.vertex(j).Y ^ 2)
            Next
        Next

        Get #1, , temp  ' sectorsdivision
        Get #1, , temp  ' num sectors

        For i = -SECTOR_NUM To SECTOR_NUM
            For j = -SECTOR_NUM To SECTOR_NUM
                Get #1, , polysInSector     ' number of polys in sector
                For k = 1 To polysInSector  ' for each poly in sector
                    Get #1, , polyIndex     ' load and discard poly index
                Next
            Next
        Next

        errorVal = "Error loading scenery"

        Get #1, , sceneryCount

        ReDim Scenery(sceneryCount)

        If sceneryCount > 0 Then
            Dim offset As Integer
            offset = 0

            For i = 1 To sceneryCount
                Get #1, , Prop

                If Prop.X > 32766 Or Prop.X < -32766 Or Prop.Y > 32766 Or Prop.Y < -32766 Then
                    offset = offset + 1
                ElseIf Prop.Width < 0 Or Prop.Height < 0 Or Int(Prop.ScaleX * 1000) = 0 Or Int(Prop.ScaleY * 1000) = 0 Then
                    offset = offset + 1
                ElseIf Prop.ScaleX < -10000 Or Prop.ScaleX > 10000 Or Prop.ScaleY < -10000 Or Prop.ScaleY > 10000 Then
                    offset = offset + 1
                ElseIf Prop.Style < 1 Then
                    offset = offset + 1
                Else
                    Scenery(i - offset).Style = Prop.Style
                    Scenery(i - offset).Translation.X = Prop.X
                    Scenery(i - offset).Translation.Y = Prop.Y
                    Scenery(i - offset).screenTr.X = (Prop.X - scrollCoords(2).X) * zoomFactor
                    Scenery(i - offset).screenTr.Y = (Prop.Y - scrollCoords(2).Y) * zoomFactor
                    Scenery(i - offset).rotation = Prop.rotation
                    Scenery(i - offset).Scaling.X = Prop.ScaleX
                    Scenery(i - offset).Scaling.Y = Prop.ScaleY
                    If Prop.alpha < 1 Then
                        Scenery(i - offset).alpha = 255
                    ElseIf Prop.alpha <= 255 Then
                        Scenery(i - offset).alpha = Prop.alpha
                    Else
                        Scenery(i - offset).alpha = 255
                    End If
                    Scenery(i - offset).color = Prop.color
                    If Prop.level <= 255 And Prop.level >= 0 Then
                        Scenery(i - offset).level = Prop.level
                    Else
                        Scenery(i - offset).level = 0
                    End If
                    Scenery(i - offset).color = ARGB(Scenery(i - offset).alpha, Scenery(i - offset).color)
                End If
            Next

            sceneryCount = sceneryCount - offset
        End If

        ReDim Preserve Scenery(sceneryCount)

        errorVal = "Error loading scenery elements"

        offset = 0

        Get #1, , sceneryElements

        ReDim Preserve SceneryTextures(sceneryElements)

        Dim scenIndex As Integer
        Dim firstOccurence As Integer

        If sceneryElements > 0 And sceneryElements < 500 Then
            For i = 1 To sceneryElements
                tempString = ""

                Get #1, , Scenery_New

                For j = 1 To Scenery_New.sceneryName(0)
                    tempString = tempString & Chr$(Scenery_New.sceneryName(j))
                Next

                Dim loadName As String

                If tempString = "" Then
                    Set SceneryTextures(i).Texture = D3DX.CreateTextureFromFileEx(D3DDevice, appPath & "\skins\" & gfxDir & "\notfound.bmp", D3DX_DEFAULT, D3DX_DEFAULT, _
                            D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
                            D3DX_FILTER_POINT, COLOR_KEY, imageInfo, ByVal 0)
                    frmScenery.lstScenery.AddItem tempString
                    tvwScenery.Nodes.Add "In Use", tvwChild, tempString, tempString
                ElseIf checkLoaded(tempString) > -1 Then

                    loadName = soldatDir & "Scenery-gfx\" & tempString
                    toTGARes = GifToBmp(loadName, appPath & "\Temp\gif.tga")
                    If Right$(loadName, 4) = ".gif" Then
                        loadName = appPath & "\Temp\gif.tga"
                    End If

                    If toTGARes = -1 Then
                        Set SceneryTextures(i).Texture = D3DX.CreateTextureFromFileEx(D3DDevice, loadName, D3DX_DEFAULT, D3DX_DEFAULT, _
                                D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
                                D3DX_FILTER_POINT, COLOR_KEY, imageInfo, ByVal 0)
                    Else
                        Set SceneryTextures(i).Texture = D3DX.CreateTextureFromFileEx(D3DDevice, appPath & "\skins\" & gfxDir & "\notfound.bmp", D3DX_DEFAULT, D3DX_DEFAULT, _
                                D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
                                D3DX_FILTER_POINT, COLOR_KEY, imageInfo, ByVal 0)
                    End If

                    frmScenery.lstScenery.AddItem tempString
                    tvwScenery.Nodes.Add "In Use", tvwChild, , tempString
                ElseIf confirmExists(tempString) Then  ' if scenery texture is in master list

                    loadName = soldatDir & "Scenery-gfx\" & tempString
                    toTGARes = GifToBmp(loadName, appPath & "\Temp\gif.tga")
                    If Right$(loadName, 4) = ".gif" Then
                        loadName = appPath & "\Temp\gif.tga"
                    End If

                    If toTGARes = -1 Then
                        Set SceneryTextures(i).Texture = D3DX.CreateTextureFromFileEx(D3DDevice, loadName, D3DX_DEFAULT, D3DX_DEFAULT, _
                                D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
                                D3DX_FILTER_POINT, COLOR_KEY, imageInfo, ByVal 0)
                    Else
                        Set SceneryTextures(i).Texture = D3DX.CreateTextureFromFileEx(D3DDevice, appPath & "\skins\" & gfxDir & "\notfound.bmp", D3DX_DEFAULT, D3DX_DEFAULT, _
                                D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
                                D3DX_FILTER_POINT, COLOR_KEY, imageInfo, ByVal 0)
                    End If
                    frmScenery.lstScenery.AddItem tempString
                    tvwScenery.Nodes.Add "In Use", tvwChild, tempString, tempString
                Else
                    Set SceneryTextures(i).Texture = D3DX.CreateTextureFromFileEx(D3DDevice, appPath & "\skins\" & gfxDir & "\notfound.bmp", D3DX_DEFAULT, D3DX_DEFAULT, _
                            D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
                            D3DX_FILTER_POINT, COLOR_KEY, imageInfo, ByVal 0)
                    frmScenery.lstScenery.AddItem tempString
                    tvwScenery.Nodes.Add "In Use", tvwChild, tempString, tempString
                End If

                SceneryTextures(i).Texture.GetLevelDesc 0, textureDesc

                SceneryTextures(i).Width = imageInfo.Width
                SceneryTextures(i).Height = imageInfo.Height
                SceneryTextures(i).reScale.X = SceneryTextures(i).Width / textureDesc.Width
                SceneryTextures(i).reScale.Y = SceneryTextures(i).Height / textureDesc.Height

                If SceneryTextures(i).reScale.X = 0 Or SceneryTextures(i).reScale.Y = 0 Then
                    SceneryTextures(i).reScale.X = 1
                    SceneryTextures(i).reScale.Y = 1
                End If
            Next

            For i = 1 To sceneryCount
                If Scenery(i).Style > sceneryElements Then
                    Scenery(i).Style = sceneryElements
                ElseIf Scenery(i).Style < 1 Then
                    Scenery(i).Style = 1
                End If
            Next
        ElseIf sceneryElements <> 0 Then
            ' if we got to this point it means that scenery were loaded but scenery elements are borked
            ' or scenery are borked too
            sceneryElements = 0
            For i = 1 To sceneryCount
                Scenery(i).Style = 0
            Next
            GoTo ErrorHandler

        End If

        errorVal = "Error loading colliders"

        Get #1, , colliderCount

        ReDim Colliders(colliderCount)

        For i = 1 To colliderCount
            Get #1, , Colliders(i)
            Colliders(i).active = 0
        Next

        errorVal = "Error loading spawn points"

        Get #1, , spawnPoints
        ReDim Spawns(spawnPoints)

        For i = 1 To spawnPoints
            Get #1, , spawn
            Spawns(i).X = spawn.X
            Spawns(i).Y = spawn.Y
            Spawns(i).Team = spawn.Team
            If Spawns(i).Team > 31 Then Spawns(i).Team = 31
            Spawns(i).active = 0
        Next

        errorVal = "Error loading waypoints"

        Get #1, , waypointCount
        ReDim Waypoints(waypointCount)
        conCount = 0
        ReDim Connections(conCount)

        For i = 1 To waypointCount
            Get #1, , newWaypoint
            Waypoints(i).tempIndex = i
            Waypoints(i).pathNum = newWaypoint.pathNum
            If newWaypoint.connectionsNum >= 0 Then
                Waypoints(i).numConnections = newWaypoint.connectionsNum
            Else
                Waypoints(i).numConnections = 0
            End If
            Waypoints(i).special = newWaypoint.special
            Waypoints(i).X = newWaypoint.X
            Waypoints(i).Y = newWaypoint.Y
            Waypoints(i).wayType(0) = CBool(newWaypoint.Left)
            Waypoints(i).wayType(1) = CBool(newWaypoint.Right)
            Waypoints(i).wayType(2) = CBool(newWaypoint.up)
            Waypoints(i).wayType(3) = CBool(newWaypoint.down)
            Waypoints(i).wayType(4) = CBool(newWaypoint.m2)
            If newWaypoint.connectionsNum > 0 And newWaypoint.connectionsNum <= 20 Then
                conCount = conCount + newWaypoint.connectionsNum
                ReDim Preserve Connections(conCount)
                For j = 1 To newWaypoint.connectionsNum
                    Connections(conCount - newWaypoint.connectionsNum + j).point1 = i
                    Connections(conCount - newWaypoint.connectionsNum + j).point2 = newWaypoint.Connections(j)
                Next
            End If
        Next

        If Options.MapRandomID < 0 Then
            Get #1, , lightCount
            ReDim Lights(lightCount)

            For i = 1 To lightCount
                Get #1, , Lights(i)
            Next

            Get #1, , sketchLines
            ReDim Preserve sketch(sketchLines)

            For i = 1 To sketchLines
                Get #1, , sketch(i)
            Next
        Else
            lightCount = 0
            ReDim Lights(lightCount)
            sketchLines = 0
            ReDim Preserve sketch(sketchLines)
        End If

    Close #1

    errorVal = "Error reloading scenery"

    fileOpen = False

    errorVal = "Error setting map data"

    setCurrentScenery 0
    If sceneryElements > 0 Then
        frmScenery.lstScenery.ListIndex = 0
    End If

    ' get map title and texture
    mapTitle = ""
    For i = 1 To Options.mapName(0)
        mapTitle = mapTitle + Chr$(Options.mapName(i))
    Next
    gTextureFile = ""
    For i = 1 To Options.textureName(0)
        gTextureFile = gTextureFile + Chr$(Options.textureName(i))
    Next

    mapTitle = ""
    For i = 1 To Options.mapName(0)
        mapTitle = mapTitle + Chr$(Options.mapName(i))
    Next

    ' get background colors
    bgColors(1) = getRGB(Options.BackgroundColor)
    bgColors(2) = getRGB(Options.BackgroundColor2)

    ' set background poly colors
    bgPolys(1) = CreateCustomVertex(-maxX - 640, -maxX - 640, 1, 1, RGB(bgColors(1).blue, bgColors(1).green, bgColors(1).red), 0, 0)
    bgPolys(2) = CreateCustomVertex(-maxX, maxX, 1, 1, RGB(bgColors(2).blue, bgColors(2).green, bgColors(2).red), 0, 1)
    bgPolys(3) = CreateCustomVertex(maxX, -maxX, 1, 1, RGB(bgColors(1).blue, bgColors(1).green, bgColors(1).red), 1, 0)
    bgPolys(4) = CreateCustomVertex(maxX, maxX, 1, 1, RGB(bgColors(2).blue, bgColors(2).green, bgColors(2).red), 1, 1)

    If (maxX - minX) > (maxY - minY) Then
        bgPolys(1).X = minX - 640
        bgPolys(1).Y = Midpoint(maxY, minY) - ((maxX - minX) / 2) - 640
        bgPolys(2).X = minX - 640
        bgPolys(2).Y = Midpoint(maxY, minY) + ((maxX - minX) / 2) + 640
        bgPolys(3).X = maxX + 640
        bgPolys(3).Y = Midpoint(maxY, minY) - ((maxX - minX) / 2) - 640
        bgPolys(4).X = maxX + 640
        bgPolys(4).Y = Midpoint(maxY, minY) + ((maxX - minX) / 2) + 640
    Else
        bgPolys(1).X = Midpoint(maxX, minX) - ((maxY - minY) / 2) - 640
        bgPolys(1).Y = minY - 640
        bgPolys(2).X = Midpoint(maxX, minX) - ((maxY - minY) / 2) - 640
        bgPolys(2).Y = maxY + 640
        bgPolys(3).X = Midpoint(maxX, minX) + ((maxY - minY) / 2) + 640
        bgPolys(3).Y = minY - 640
        bgPolys(4).X = Midpoint(maxX, minX) + ((maxY - minY) / 2) + 640
        bgPolys(4).Y = maxY + 640
    End If

    For i = 1 To 4
        bgPolyCoords(i).X = bgPolys(i).X
        bgPolyCoords(i).Y = bgPolys(i).Y
    Next

    If Len(Dir$(soldatDir & "textures\" & gTextureFile)) <> 0 Then
        setMapTexture gTextureFile
        frmTexture.setTexture gTextureFile
    End If

    Colliders(0).radius = clrRadius

    setMapData
    txtZoom.Text = Int(zoomFactor * 1000 + 0.5) / 10 & "%"

    centerView

    numUndo = 0
    numRedo = 0
    currentUndo = 0

    If lightCount > 0 Then
        frmDisplay.setLayer 9, showLights
        applyLights
    End If

    SaveUndo

    Render

    Exit Sub

ErrorHandler:

    MsgBox "error loading map" & vbNewLine & Error$ & vbNewLine & errorVal
    If fileOpen Then Close #1
    noRedraw = False

End Sub

Private Function checkLoaded(sceneryName As String) As Integer

    Dim i As Integer

    On Error GoTo ErrorHandler

    checkLoaded = -1

    For i = 0 To frmScenery.lstScenery.ListCount - 1
        If frmScenery.lstScenery.List(i) = sceneryName Then checkLoaded = i
    Next

    Exit Function

ErrorHandler:

    MsgBox "error checking loaded scenery" & vbNewLine & Error$

End Function

Private Function getMapDimensions() As String

    getMapDimensions = Int(maxX - minX) & "x" & Int(maxY - minY)

End Function

Private Function getMapArea() As Long

    Dim i As Integer
    Dim area As Double
    Dim A As Single
    Dim B As Single
    Dim c As Single
    Dim x1 As Single
    Dim y1 As Single
    Dim x2 As Single
    Dim y2 As Single

    For i = 1 To mPolyCount
        If vertexList(i).polyType <> 3 Then
            x1 = (PolyCoords(i).vertex(3).X - PolyCoords(i).vertex(2).X)
            y1 = (PolyCoords(i).vertex(3).Y - PolyCoords(i).vertex(2).Y)
            x2 = (PolyCoords(i).vertex(1).X - PolyCoords(i).vertex(3).X)
            y2 = (PolyCoords(i).vertex(1).Y - PolyCoords(i).vertex(3).Y)
            A = Sqr(x1 ^ 2 + y1 ^ 2)
            B = Sqr(x2 ^ 2 + y2 ^ 2)
            c = GetAngle(x1, y1) - GetAngle(x2, y2)
            area = area + (A * B * Sin(c) / 2)
        End If
    Next

    MsgBox Int(area / ((maxX - minX) * (maxY - minY)) * 100 + 0.5) & "%"

End Function

Public Sub setMapData()

    frmInfo.lblCount(0).Caption = mPolyCount
    frmInfo.lblCount(1).Caption = sceneryCount & "/500 (" & sceneryElements & ")"
    frmInfo.lblCount(2).Caption = spawnPoints & "/128"
    frmInfo.lblCount(3).Caption = colliderCount & "/128"
    frmInfo.lblCount(4).Caption = waypointCount & "/500"
    frmInfo.lblCount(5).Caption = conCount
    frmInfo.lblCount(6).Caption = getMapDimensions

End Sub

Public Sub setCurrentScenery(Optional styleVal As Integer = -1, Optional sceneryName As String = "")

    On Error GoTo ErrorHandler

    If styleVal > -1 Then
        Scenery(0).Style = styleVal
    End If

    If sceneryName <> "" Then
        currentScenery = sceneryName
    End If

    Scenery(0).alpha = opacity * 255
    Scenery(0).color = ARGB(opacity * 255, RGB(gPolyClr.blue, gPolyClr.green, gPolyClr.red))
    Scenery(0).level = frmScenery.level
    Scenery(0).Scaling.X = 1
    Scenery(0).Scaling.Y = 1
    Scenery(0).screenTr.X = mouseCoords.X
    Scenery(0).screenTr.Y = mouseCoords.Y
    Scenery(0).rotation = 0

    Exit Sub

ErrorHandler:

    MsgBox "Error setting current scenery" & vbNewLine & Error$

End Sub

Public Sub setCurrentTexture(sceneryName As String)

    On Error GoTo ErrorHandler

    Dim loadName As String
    Dim toTGARes As Long

    loadName = soldatDir & "Scenery-gfx\" & sceneryName
    toTGARes = GifToBmp(loadName, appPath & "\Temp\gif.tga")
    If Right$(loadName, 4) = ".gif" Then
        loadName = appPath & "\Temp\gif.tga"
    End If

    If toTGARes = -1 Then
        Set SceneryTextures(0).Texture = D3DX.CreateTextureFromFileEx(D3DDevice, loadName, D3DX_DEFAULT, D3DX_DEFAULT, _
                D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
                D3DX_FILTER_POINT, COLOR_KEY, imageInfo, ByVal 0)
    Else
        Set SceneryTextures(0).Texture = D3DX.CreateTextureFromFileEx(D3DDevice, appPath & "\skins\" & gfxDir & "\notfound.bmp", D3DX_DEFAULT, D3DX_DEFAULT, _
                D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
                D3DX_FILTER_POINT, COLOR_KEY, imageInfo, ByVal 0)
    End If

    SceneryTextures(0).Texture.GetLevelDesc 0, textureDesc

    SceneryTextures(0).Width = imageInfo.Width
    SceneryTextures(0).Height = imageInfo.Height

    SceneryTextures(0).reScale.X = SceneryTextures(0).Width / textureDesc.Width
    SceneryTextures(0).reScale.Y = SceneryTextures(0).Height / textureDesc.Height

    If SceneryTextures(0).reScale.X = 0 Or SceneryTextures(0).reScale.Y = 0 Then
        SceneryTextures(0).reScale.X = 1
        SceneryTextures(0).reScale.Y = 1
    End If

    setCurrentScenery 0
    Scenery(0).Style = 0

    Exit Sub

ErrorHandler:

    MsgBox "Error creating current scenery texture" & vbNewLine & Error$

End Sub

Public Sub setSceneryLevel(ByVal level As Byte)

    Scenery(0).level = level

End Sub

Public Sub CreateSceneryTexture(sceneryName As String)

    On Error GoTo ErrorHandler

    sceneryElements = sceneryElements + 1
    ReDim Preserve SceneryTextures(sceneryElements)

    Dim loadName As String
    Dim toTGARes As Long

    loadName = soldatDir & "Scenery-gfx\" & sceneryName
    toTGARes = GifToBmp(loadName, appPath & "\Temp\gif.tga")
    If Right$(loadName, 4) = ".gif" Then
        loadName = appPath & "\Temp\gif.tga"
    End If

    If toTGARes = -1 Then
        Set SceneryTextures(sceneryElements).Texture = D3DX.CreateTextureFromFileEx(D3DDevice, loadName, D3DX_DEFAULT, D3DX_DEFAULT, _
                D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
                D3DX_FILTER_POINT, COLOR_KEY, imageInfo, ByVal 0)
    Else
        Set SceneryTextures(sceneryElements).Texture = D3DX.CreateTextureFromFileEx(D3DDevice, appPath & "\skins\" & gfxDir & "\notfound.bmp", D3DX_DEFAULT, D3DX_DEFAULT, _
                D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
                D3DX_FILTER_POINT, COLOR_KEY, imageInfo, ByVal 0)
    End If

    frmScenery.lstScenery.AddItem sceneryName
    tvwScenery.Nodes.Add "In Use", tvwChild, sceneryName, sceneryName

    SceneryTextures(sceneryElements).Texture.GetLevelDesc 0, textureDesc

    SceneryTextures(sceneryElements).Width = imageInfo.Width
    SceneryTextures(sceneryElements).Height = imageInfo.Height

    SceneryTextures(sceneryElements).reScale.X = SceneryTextures(sceneryElements).Width / textureDesc.Width
    SceneryTextures(sceneryElements).reScale.Y = SceneryTextures(sceneryElements).Height / textureDesc.Height

    If SceneryTextures(sceneryElements).reScale.X = 0 Or SceneryTextures(sceneryElements).reScale.Y = 0 Then
        SceneryTextures(sceneryElements).reScale.X = 1
        SceneryTextures(sceneryElements).reScale.Y = 1
    End If

    Exit Sub

ErrorHandler:

    MsgBox "Error creating scenery texture: " & sceneryName & vbNewLine & Error$
    SceneryTextures(sceneryElements) = SceneryTextures(0)

End Sub

Public Sub RefreshSceneryTextures(Index As Integer)

    If frmScenery.lstScenery.ListCount = 0 Then Exit Sub

    Dim sceneryName As String
    Dim scenNum As Integer

    sceneryName = frmScenery.lstScenery.List(Index - 1)

    Dim loadName As String
    Dim toTGARes As Long

    loadName = soldatDir & "Scenery-gfx\" & sceneryName
    toTGARes = GifToBmp(loadName, appPath & "\Temp\gif.tga")
    If Right$(loadName, 4) = ".gif" Then
        loadName = appPath & "\Temp\gif.tga"
    End If

    If toTGARes = -1 Then
        Set SceneryTextures(Index).Texture = D3DX.CreateTextureFromFileEx(D3DDevice, loadName, D3DX_DEFAULT, D3DX_DEFAULT, _
                D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
                D3DX_FILTER_POINT, COLOR_KEY, imageInfo, ByVal 0)
    Else
        Set SceneryTextures(Index).Texture = D3DX.CreateTextureFromFileEx(D3DDevice, appPath & "\skins\" & gfxDir & "\notfound.bmp", D3DX_DEFAULT, D3DX_DEFAULT, _
                D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
                D3DX_FILTER_POINT, COLOR_KEY, imageInfo, ByVal 0)
    End If

    SceneryTextures(Index).Texture.GetLevelDesc 0, textureDesc

    SceneryTextures(Index).Width = imageInfo.Width
    SceneryTextures(Index).Height = imageInfo.Height

    SceneryTextures(Index).reScale.X = SceneryTextures(Index).Width / textureDesc.Width
    SceneryTextures(Index).reScale.Y = SceneryTextures(Index).Height / textureDesc.Height

    If SceneryTextures(Index).reScale.X = 0 Or SceneryTextures(Index).reScale.Y = 0 Then
        SceneryTextures(Index).reScale.X = 1
        SceneryTextures(Index).reScale.Y = 1
    End If

End Sub

Private Sub SaveFile(theFileName As String)

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim X As Integer
    Dim Y As Integer

    Dim xOffset As Integer
    Dim yOffset As Integer

    Dim xDiff As Single
    Dim yDiff As Single
    Dim length As Single
    Dim VertNum As Byte
    Dim mapWidth As Long
    Dim mapHeight As Long

    Const SECTOR_NUM As Long = 25

    Dim Polygon As TMapFile_Polygon
    Dim sectorsDivision As Long

    Const ZERO As Integer = 0

    Dim Scenery_New As TMapFile_Scenery
    Dim newWaypoint As TNewWaypoint
    Dim sceneryName As String
    Dim Prop As TProp
    Dim spawn As TSaveSpawnPoint
    Dim tempClr As TColor
    Dim connectedNum As Integer

    Dim fileOpen As Boolean

    Me.MousePointer = vbHourglass

    ' refresh background
    mnuRefreshBG_Click

    mapWidth = maxX - minX
    mapHeight = maxY - minY

    Options.BackgroundColor = ARGB(255, RGB(bgColors(1).blue, bgColors(1).green, bgColors(1).red))
    Options.BackgroundColor2 = ARGB(255, RGB(bgColors(2).blue, bgColors(2).green, bgColors(2).red))
    ' set texture name
    Options.textureName(0) = Len(gTextureFile)
    For i = 1 To Len(gTextureFile)
        Options.textureName(i) = Asc(Mid(gTextureFile, i, 1))
    Next
    ' set map name
    Options.mapName(0) = Len(mapTitle)
    If Options.mapName(0) > 38 Then Options.mapName(0) = 38
    For i = 1 To Options.mapName(0)
        Options.mapName(i) = Asc(Mid(mapTitle, i, 1))
    Next

    Options.MapRandomID = -1

    If mapWidth > mapHeight Then
        sectorsDivision = Int((mapWidth + 100) / 25)
    Else
        sectorsDivision = Int((mapHeight + 100) / 25)
    End If

    Open theFileName For Binary Access Write Lock Write As #1

        fileOpen = True

        Put #1, , Version
        Put #1, , Options

        ' save polys
        Put #1, , mPolyCount
        For i = 1 To mPolyCount
            Polygon.Poly = Polys(i)

            For j = 1 To 3
                Polygon.Poly.vertex(j).X = PolyCoords(i).vertex(j).X
                Polygon.Poly.vertex(j).Y = PolyCoords(i).vertex(j).Y

                Polygon.Poly.vertex(j).color = ARGB(getAlpha(Polys(i).vertex(j).color), RGB(vertexList(i).color(j).blue, vertexList(i).color(j).green, vertexList(i).color(j).red))

                VertNum = j + 1
                If VertNum > 3 Then VertNum = 1

                xDiff = PolyCoords(i).vertex(VertNum).X - PolyCoords(i).vertex(j).X
                yDiff = PolyCoords(i).vertex(j).Y - PolyCoords(i).vertex(VertNum).Y
                If xDiff = 0 And yDiff = 0 Then
                    length = 1
                Else
                    length = Sqr(xDiff ^ 2 + yDiff ^ 2)
                End If
                Polygon.Poly.Perp.vertex(j).X = (yDiff / length) * Polygon.Poly.Perp.vertex(j).Z
                Polygon.Poly.Perp.vertex(j).Y = (xDiff / length) * Polygon.Poly.Perp.vertex(j).Z
                Polygon.Poly.Perp.vertex(j).Z = 1
            Next

            Polygon.polyType = vertexList(i).polyType

            Put #1, , Polygon
        Next

        Put #1, , sectorsDivision
        Put #1, , SECTOR_NUM

        For i = -25 To 25
            For j = -25 To 25
                Put #1, , ZERO
            Next
        Next

        Put #1, , sceneryCount

        For i = 1 To sceneryCount
            Prop.active = True
            Prop.alpha = Scenery(i).alpha
            tempClr = getRGB(Scenery(i).color)
            Prop.color = ARGB(255, RGB(tempClr.blue, tempClr.green, tempClr.red))
            Prop.Width = SceneryTextures(Scenery(i).Style).Width
            Prop.Height = SceneryTextures(Scenery(i).Style).Height
            Prop.level = Scenery(i).level
            Prop.rotation = Scenery(i).rotation
            Prop.ScaleX = Scenery(i).Scaling.X
            Prop.ScaleY = Scenery(i).Scaling.Y
            Prop.X = Scenery(i).Translation.X - xOffset
            Prop.Y = Scenery(i).Translation.Y - yOffset
            Prop.Style = Scenery(i).Style

            Put #1, , Prop
        Next

        Put #1, , sceneryElements

        For i = 1 To sceneryElements
            sceneryName = frmScenery.lstScenery.List(i - 1)
            Scenery_New.sceneryName(0) = Len(sceneryName)
            For j = 1 To Scenery_New.sceneryName(0)
                Scenery_New.sceneryName(j) = Asc(Mid(sceneryName, j, 1))
            Next
            Scenery_New.Date = getFileDate(sceneryName)
            Put #1, , Scenery_New
        Next

        Put #1, , colliderCount

        For i = 1 To colliderCount
            Colliders(i).active = 1
            Put #1, , Colliders(i)
            Colliders(i).active = 0
        Next

        Put #1, , spawnPoints

        For i = 1 To spawnPoints
            spawn.active = 1
            spawn.X = Spawns(i).X
            spawn.Y = Spawns(i).Y
            spawn.Team = Spawns(i).Team
            Put #1, , spawn
            Spawns(i).active = 0
        Next

        Put #1, , waypointCount

        For i = 1 To waypointCount
            newWaypoint.active = 1
            newWaypoint.X = Waypoints(i).X
            newWaypoint.Y = Waypoints(i).Y
            newWaypoint.connectionsNum = Waypoints(i).numConnections
            If Waypoints(i).wayType(0) Then newWaypoint.Left = 1 Else newWaypoint.Left = 0
            If Waypoints(i).wayType(1) Then newWaypoint.Right = 1 Else newWaypoint.Right = 0
            If Waypoints(i).wayType(2) Then newWaypoint.up = 1 Else newWaypoint.up = 0
            If Waypoints(i).wayType(3) Then newWaypoint.down = 1 Else newWaypoint.down = 0
            If Waypoints(i).wayType(4) Then newWaypoint.m2 = 1 Else newWaypoint.m2 = 0
            newWaypoint.id = i
            newWaypoint.pathNum = Waypoints(i).pathNum
            newWaypoint.special = Waypoints(i).special
            connectedNum = 0
            For j = 1 To conCount
                If Connections(j).point1 = i And connectedNum < 20 Then
                    connectedNum = connectedNum + 1
                    newWaypoint.Connections(connectedNum) = Connections(j).point2
                End If
            Next
            Waypoints(i).numConnections = connectedNum
            newWaypoint.connectionsNum = connectedNum
            Put #1, , newWaypoint
        Next

        Put #1, , lightCount

        For i = 1 To lightCount
            Put #1, , Lights(i)
        Next

        Put #1, , sketchLines

        For i = 1 To sketchLines
            Put #1, , sketch(i)
        Next

    Close #1

    fileOpen = False

    currentFileName = ""
    For i = 0 To Len(theFileName) - 1
        If Mid(theFileName, Len(theFileName) - i, 1) <> "\" Then
            currentFileName = Mid(theFileName, Len(theFileName) - i, 1) + currentFileName
        Else
            Exit For
        End If
    Next

    lblFileName.Caption = currentFileName

    Me.MousePointer = vbCustom

    Exit Sub

ErrorHandler:

    MsgBox "Error saving map" & vbNewLine & Error$
    If fileOpen Then
        Close #1
    End If

End Sub

Public Sub SaveAndCompile(theFileName As String)

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim X As Integer
    Dim Y As Integer

    Dim xOffset As Integer
    Dim yOffset As Integer

    Dim xDiff As Single
    Dim yDiff As Single
    Dim length As Single
    Dim VertNum As Byte
    Dim sector(1 To 256) As Integer
    Dim xSecNum As Integer
    Dim ySecNum As Integer
    Dim mapWidth As Integer
    Dim mapHeight As Integer

    Const SECTOR_NUM As Long = 25

    Dim Polygon As TMapFile_Polygon
    Dim sectorsDivision As Long
    Dim polysInSector As Integer

    Dim Scenery_New As TMapFile_Scenery
    Dim newWaypoint As TNewWaypoint
    Dim sceneryName As String
    Dim Prop As TProp
    Dim tempClr As TColor
    Dim connectedNum As Integer

    Dim newSpawnPoint As TSaveSpawnPoint
    Dim newCollider As TCollider

    Const ZERO As Integer = 0

    Dim fileOpen As Boolean

    On Error GoTo ErrorHandler

    Me.MousePointer = vbHourglass

    Randomize

    polysInSector = 0

    newSpawnPoint.active = 1
    newCollider.active = 1

    ' refresh background
    mnuRefreshBG_Click

    ' find offsets to center map
    xOffset = Int(Midpoint(maxX, minX))
    yOffset = Int(Midpoint(maxY, minY))

    mapWidth = maxX - xOffset
    mapHeight = maxY - yOffset

    Options.BackgroundColor = ARGB(255, RGB(bgColors(1).blue, bgColors(1).green, bgColors(1).red))
    Options.BackgroundColor2 = ARGB(255, RGB(bgColors(2).blue, bgColors(2).green, bgColors(2).red))
    ' set texture name
    Options.textureName(0) = Len(gTextureFile)
    If Options.textureName(0) > 24 Then Options.textureName(0) = 24
    For i = 1 To Options.textureName(0)
        Options.textureName(i) = Asc(Mid(gTextureFile, i, 1))
    Next
    ' set map name
    Options.mapName(0) = Len(mapTitle)
    If Options.mapName(0) > 38 Then Options.mapName(0) = 38
    For i = 1 To Options.mapName(0)
        Options.mapName(i) = Asc(Mid(mapTitle, i, 1))
    Next

    ' set map random ID
    Options.MapRandomID = (Rnd * 999999) + 10000

    xSecNum = SECTOR_NUM
    ySecNum = SECTOR_NUM

    If mapWidth > mapHeight Then
        sectorsDivision = Int((mapWidth + 100) / 25)
        ySecNum = (mapHeight + 100) / sectorsDivision
    Else
        sectorsDivision = Int((mapHeight + 100) / 25)
        xSecNum = (mapWidth + 100) / sectorsDivision
    End If

    Open theFileName For Binary Access Write Lock Write As #1

        fileOpen = True

        Put #1, , Version
        Put #1, , Options

        ' save polys
        Put #1, , mPolyCount
        For i = 1 To mPolyCount
            Polygon.Poly = Polys(i)
            Polygon.polyType = vertexList(i).polyType

            Polygon.Poly.vertex(1).X = PolyCoords(i).vertex(1).X - xOffset
            Polygon.Poly.vertex(1).Y = PolyCoords(i).vertex(1).Y - yOffset
            Polygon.Poly.vertex(2).X = PolyCoords(i).vertex(2).X - xOffset
            Polygon.Poly.vertex(2).Y = PolyCoords(i).vertex(2).Y - yOffset
            Polygon.Poly.vertex(3).X = PolyCoords(i).vertex(3).X - xOffset
            Polygon.Poly.vertex(3).Y = PolyCoords(i).vertex(3).Y - yOffset

            For j = 1 To 3
                VertNum = j + 1
                If VertNum > 3 Then VertNum = 1

                xDiff = Polygon.Poly.vertex(VertNum).X - Polygon.Poly.vertex(j).X
                yDiff = Polygon.Poly.vertex(j).Y - Polygon.Poly.vertex(VertNum).Y
                If xDiff = 0 And yDiff = 0 Then
                    length = 1
                Else
                    length = Sqr(xDiff ^ 2 + yDiff ^ 2)
                End If
                If Polygon.polyType = 18 Then
                    If Polygon.Poly.Perp.vertex(j).Z < 1 Then
                        Polygon.Poly.Perp.vertex(j).Z = 1
                    End If
                Else
                    Polygon.Poly.Perp.vertex(j).Z = 1
                End If
                Polygon.Poly.Perp.vertex(j).X = (yDiff / length) * Polygon.Poly.Perp.vertex(j).Z
                Polygon.Poly.Perp.vertex(j).Y = (xDiff / length) * Polygon.Poly.Perp.vertex(j).Z
                Polygon.Poly.Perp.vertex(j).Z = 1
                Polygon.Poly.vertex(j).Z = 1
            Next

            Put #1, , Polygon
        Next

        Put #1, , sectorsDivision
        Put #1, , SECTOR_NUM

        ' generate sectors
        For X = -SECTOR_NUM To SECTOR_NUM
            For Y = -SECTOR_NUM To SECTOR_NUM
                polysInSector = 0

                If X >= -xSecNum And X <= xSecNum And Y >= -ySecNum And Y <= ySecNum Then  ' if sectors within range
                    For i = 1 To mPolyCount
                        If vertexList(i).polyType <> 3 Then
                        If isInSector(i, sectorsDivision * (X - 0.5) + xOffset - 1, sectorsDivision * (Y - 0.5) + yOffset - 1, sectorsDivision + 2) Then
                            polysInSector = polysInSector + 1
                            If polysInSector > 256 Then
                                polysInSector = 256
                            Else
                                sector(polysInSector) = i
                            End If
                        End If
                        End If
                    Next

                    If polysInSector > 256 Then polysInSector = 256
                End If

                Put #1, , polysInSector

                If polysInSector > 0 Then
                    For k = 1 To polysInSector
                        Put #1, , sector(k)
                    Next
                End If
            Next
            picProgress.Line ((X + SECTOR_NUM) * 2, 0)-((X + SECTOR_NUM) * 2, 12), RGB(61, 75, 97)
            picProgress.Line ((X + SECTOR_NUM) * 2 + 1, 0)-((X + SECTOR_NUM) * 2 + 1, 12), RGB(61, 75, 97)
            picProgress.Refresh
        Next

        picProgress.Cls

        Put #1, , sceneryCount

        For i = 1 To sceneryCount
            Prop.active = True
            Prop.alpha = Scenery(i).alpha
            tempClr = getRGB(Scenery(i).color)
            Prop.color = ARGB(255, RGB(tempClr.blue, tempClr.green, tempClr.red))
            Prop.Width = SceneryTextures(Scenery(i).Style).Width
            Prop.Height = SceneryTextures(Scenery(i).Style).Height
            Prop.level = Scenery(i).level
            Prop.rotation = Scenery(i).rotation
            Prop.ScaleX = Scenery(i).Scaling.X
            Prop.ScaleY = Scenery(i).Scaling.Y
            Prop.X = Scenery(i).Translation.X - xOffset
            Prop.Y = Scenery(i).Translation.Y - yOffset
            Prop.Style = Scenery(i).Style

            Put #1, , Prop
        Next

        Put #1, , sceneryElements

        For i = 1 To sceneryElements
            sceneryName = frmScenery.lstScenery.List(i - 1)
            Scenery_New.sceneryName(0) = Len(sceneryName)
            For j = 1 To Scenery_New.sceneryName(0)
                Scenery_New.sceneryName(j) = Asc(Mid(sceneryName, j, 1))
            Next
            Scenery_New.Date = getFileDate(sceneryName)
            Put #1, , Scenery_New
        Next

        Put #1, , colliderCount

        For i = 1 To colliderCount
            newCollider.radius = Colliders(i).radius
            newCollider.X = Colliders(i).X - xOffset
            newCollider.Y = Colliders(i).Y - yOffset
            Put #1, , newCollider
        Next

        Put #1, , spawnPoints

        For i = 1 To spawnPoints
            newSpawnPoint.Team = Spawns(i).Team
            newSpawnPoint.X = Spawns(i).X - xOffset
            newSpawnPoint.Y = Spawns(i).Y - yOffset
            Put #1, , newSpawnPoint
        Next

        Put #1, , waypointCount

        For i = 1 To waypointCount
            newWaypoint.active = 1
            newWaypoint.X = Waypoints(i).X - xOffset
            newWaypoint.Y = Waypoints(i).Y - yOffset
            newWaypoint.connectionsNum = Waypoints(i).numConnections
            If Waypoints(i).wayType(0) Then newWaypoint.Left = 1 Else newWaypoint.Left = 0
            If Waypoints(i).wayType(1) Then newWaypoint.Right = 1 Else newWaypoint.Right = 0
            If Waypoints(i).wayType(2) Then newWaypoint.up = 1 Else newWaypoint.up = 0
            If Waypoints(i).wayType(3) Then newWaypoint.down = 1 Else newWaypoint.down = 0
            If Waypoints(i).wayType(4) Then newWaypoint.m2 = 1 Else newWaypoint.m2 = 0
            newWaypoint.id = i
            newWaypoint.pathNum = Waypoints(i).pathNum
            newWaypoint.special = Waypoints(i).special
            connectedNum = 0
            For j = 1 To conCount
                If Connections(j).point1 = i And connectedNum < 20 Then
                    connectedNum = connectedNum + 1
                    newWaypoint.Connections(connectedNum) = Connections(j).point2
                End If
            Next
            Waypoints(i).numConnections = connectedNum
            newWaypoint.connectionsNum = connectedNum
            Put #1, , newWaypoint
        Next

        Put #1, , ZERO
        Put #1, , ZERO
        Put #1, , ZERO
        Put #1, , ZERO

    Close #1

    fileOpen = False

    Me.MousePointer = vbCustom
    SetCursor currentFunction + 1

    Render

    Exit Sub

ErrorHandler:

    MsgBox "Error saving/compiling map: " & Error$
    If fileOpen Then
        Close #1
    End If

End Sub

Private Sub SaveUndo()

    On Error GoTo ErrorHandler

    Dim i As Integer
    Dim j As Integer
    Dim Polygon As TPolygon
    Dim theFileName As String

    selectionChanged = False

    numRedo = 0
    numUndo = numUndo + 1
    If numUndo > maxUndo Then
        numUndo = maxUndo
    End If
    currentUndo = currentUndo + 1
    If currentUndo > maxUndo Then
        currentUndo = 0
    End If

    theFileName = appPath & "\undo\undo" & currentUndo & ".pwn"

    If Dir(appPath & "\undo\", vbDirectory) = "" Then
         MkDir (appPath & "\undo\")
    End If

    Open theFileName For Binary Access Write Lock Write As #1

        ' save polys
        Put #1, , mPolyCount
        For i = 1 To mPolyCount
            Polygon = Polys(i)
            For j = 1 To 3
                Polygon.vertex(j).X = PolyCoords(i).vertex(j).X
                Polygon.vertex(j).Y = PolyCoords(i).vertex(j).Y
            Next
            Put #1, , Polygon
            Put #1, , vertexList(i)
        Next

        Put #1, , sceneryCount
        For i = 1 To sceneryCount
            Put #1, , Scenery(i)
        Next

        Put #1, , colliderCount
        For i = 1 To colliderCount
            Put #1, , Colliders(i)
        Next

        Put #1, , spawnPoints
        For i = 1 To spawnPoints
            Put #1, , Spawns(i)
        Next

        Put #1, , lightCount
        For i = 1 To lightCount
            Put #1, , Lights(i)
        Next

        Put #1, , waypointCount
        For i = 1 To waypointCount
            Put #1, , Waypoints(i)
        Next

        Put #1, , conCount
        For i = 1 To conCount
            Put #1, , Connections(i)
        Next

        Put #1, , numSelectedPolys
        For i = 1 To numSelectedPolys
            Put #1, , selectedPolys(i)
        Next

        Put #1, , numSelectedScenery
        Put #1, , numSelSpawns
        Put #1, , numSelColliders
        Put #1, , numSelWaypoints

        For i = 0 To 3
            Put #1, , selRect(i)
        Next

    Close #1

    Exit Sub

ErrorHandler:

    MsgBox "Error saving undo" & vbNewLine & Error$

End Sub

Private Sub loadUndo(redo As Boolean)

    Dim i As Integer
    Dim j As Integer
    Dim theFileName As String
    Dim errorVal As String

    On Error GoTo ErrorHandler

    If selectionChanged Then
        SaveUndo
        selectionChanged = False
    End If

    If toolAction = True And numVerts > 0 Then
        toolAction = False
        numVerts = 0
        Render
        Exit Sub
    End If

    currentWaypoint = 0

    If redo Then
        If numRedo < 1 Then Exit Sub
        currentUndo = currentUndo + 1
        numUndo = numUndo + 1
        numRedo = numRedo - 1
    Else  ' undo
        If numUndo <= 1 Then Exit Sub
        currentUndo = currentUndo - 1
        numUndo = numUndo - 1
        numRedo = numRedo + 1
    End If
    If currentUndo < 0 Then
        currentUndo = maxUndo
    ElseIf currentUndo > maxUndo Then
        currentUndo = 0
    End If

    numSelectedPolys = 0
    ReDim selectedPolys(0)

    theFileName = appPath & "\undo\undo" & currentUndo & ".pwn"

    errorVal = "Error opening file"

    Open theFileName For Binary Access Read Lock Read As #1

        errorVal = "Error loading polygons"

        Get #1, , mPolyCount
        ReDim Polys(0 To mPolyCount)
        ReDim PolyCoords(0 To mPolyCount)
        ReDim vertexList(0 To mPolyCount)

        For i = 1 To mPolyCount
            Get #1, , Polys(i)
            Get #1, , vertexList(i)
            For j = 1 To 3
                PolyCoords(i).vertex(j).X = Polys(i).vertex(j).X
                PolyCoords(i).vertex(j).Y = Polys(i).vertex(j).Y
                Polys(i).vertex(j).X = (PolyCoords(i).vertex(j).X - scrollCoords(2).X) * zoomFactor
                Polys(i).vertex(j).Y = (PolyCoords(i).vertex(j).Y - scrollCoords(2).Y) * zoomFactor
            Next
        Next

        errorVal = "Error loading scenery"

        Get #1, , sceneryCount
        ReDim Preserve Scenery(sceneryCount)
        If sceneryCount > 0 Then
            For i = 1 To sceneryCount
                Get #1, , Scenery(i)
                Scenery(i).screenTr.X = (Scenery(i).Translation.X - scrollCoords(2).X) * zoomFactor
                Scenery(i).screenTr.Y = (Scenery(i).Translation.Y - scrollCoords(2).Y) * zoomFactor
            Next
        End If

        errorVal = "Error loading colliders"

        Get #1, , colliderCount
        ReDim Preserve Colliders(colliderCount)
        For i = 1 To colliderCount
            Get #1, , Colliders(i)
        Next

        errorVal = "Error loading spawnpoints"

        Get #1, , spawnPoints
        ReDim Preserve Spawns(spawnPoints)
        For i = 1 To spawnPoints
            Get #1, , Spawns(i)
        Next

        errorVal = "Error loading lights"

        Get #1, , lightCount
        ReDim Preserve Lights(lightCount)
        For i = 1 To lightCount
            Get #1, , Lights(i)
        Next

        errorVal = "Error loading waypoints"

        Get #1, , waypointCount
        ReDim Waypoints(waypointCount)
        For i = 1 To waypointCount
            Get #1, , Waypoints(i)
        Next

        errorVal = "Error loading connections"

        Get #1, , conCount
        ReDim Connections(conCount)
        For i = 1 To conCount
            Get #1, , Connections(i)
        Next

        errorVal = "Error loading selected polys"

        Get #1, , numSelectedPolys
        ReDim selectedPolys(numSelectedPolys)
        For i = 1 To numSelectedPolys
            Get #1, , selectedPolys(i)
        Next

        errorVal = "Error loading selected scenery"

        Get #1, , numSelectedScenery
        Get #1, , numSelSpawns
        Get #1, , numSelColliders
        Get #1, , numSelWaypoints

        For i = 0 To 3
            Get #1, , selRect(i)
        Next

    Close #1

    errorVal = "Error loading undo state"

    setMapData

    getRCenter

    Render

    Exit Sub

ErrorHandler:

    MsgBox Error$ & vbNewLine & errorVal

End Sub

Private Function isInSector(Index As Integer, X As Integer, Y As Integer, ByVal div As Long) As Boolean

    On Error GoTo ErrorHandler

    isInSector = False

    ' is poly outside of sector for sure
    If (PolyCoords(Index).vertex(1).X < X) And (PolyCoords(Index).vertex(2).X < X) And (PolyCoords(Index).vertex(3).X < X) Then
        Exit Function
    ElseIf (PolyCoords(Index).vertex(1).X > X + div) And (PolyCoords(Index).vertex(2).X > X + div) And (PolyCoords(Index).vertex(3).X > X + div) Then
        Exit Function
    ElseIf (PolyCoords(Index).vertex(1).Y < Y) And (PolyCoords(Index).vertex(2).Y < Y) And (PolyCoords(Index).vertex(3).Y < Y) Then
        Exit Function
    ElseIf (PolyCoords(Index).vertex(1).Y > Y + div) And (PolyCoords(Index).vertex(2).Y > Y + div) And (PolyCoords(Index).vertex(3).Y > Y + div) Then
        Exit Function
    End If

    ' is vertex in sector
    If isBetween(X, PolyCoords(Index).vertex(1).X, X + div) And isBetween(Y, PolyCoords(Index).vertex(1).Y, Y + div) Then
        isInSector = True
        Exit Function
    ElseIf isBetween(X, PolyCoords(Index).vertex(2).X, X + div) And isBetween(Y, PolyCoords(Index).vertex(2).Y, Y + div) Then
        isInSector = True
        Exit Function
    ElseIf isBetween(X, PolyCoords(Index).vertex(3).X, X + div) And isBetween(Y, PolyCoords(Index).vertex(3).Y, Y + div) Then
        isInSector = True
        Exit Function
    End If

    ' check if sector corner is in poly
    If Not isInSector Then
        If pointInPoly(X, Y, Index) Then
            isInSector = True
            Exit Function
        ElseIf pointInPoly(X + div, Y, Index) Then
            isInSector = True
            Exit Function
        ElseIf pointInPoly(X, Y + div, Index) Then
            isInSector = True
            Exit Function
        ElseIf pointInPoly(X + div, Y + div, Index) Then
            isInSector = True
            Exit Function
        End If
    End If

    Dim A1 As D3DVECTOR2
    Dim B1 As D3DVECTOR2
    Dim A2 As D3DVECTOR2
    Dim B2 As D3DVECTOR2

    Dim indexA1 As Integer
    Dim indexB1 As Integer

    For indexA1 = 1 To 3
        indexB1 = indexA1 + 1
        If indexB1 > 3 Then indexB1 = 1
        A1.X = PolyCoords(Index).vertex(indexA1).X
        A1.Y = PolyCoords(Index).vertex(indexA1).Y
        B1.X = PolyCoords(Index).vertex(indexB1).X
        B1.Y = PolyCoords(Index).vertex(indexB1).Y

        A2.X = X
        A2.Y = Y
        B2.X = X + div
        B2.Y = Y
        If SegXSeg(A1, B1, A2, B2) Then  ' top
            isInSector = True
            Exit Function
        End If
        A2.X = X
        A2.Y = Y + div
        B2.X = X + div
        B2.Y = Y + div
        If SegXSeg(A1, B1, A2, B2) Then  ' bottom
            isInSector = True
            Exit Function
        End If
        A2.X = X
        A2.Y = Y
        B2.X = X
        B2.Y = Y + div
        If SegXSeg(A1, B1, A2, B2) Then  ' left
            isInSector = True
            Exit Function
        End If
        A2.X = X + div
        A2.Y = Y
        B2.X = X + div
        B2.Y = Y + div
        If SegXSeg(A1, B1, A2, B2) Then  ' right
            isInSector = True
            Exit Function
        End If
    Next

    Exit Function

ErrorHandler:

    MsgBox "Sector error, " & Error$

End Function

Private Function isInSector2(Index As Integer, X As Integer, Y As Integer, div As Long) As Integer

    Dim i As Integer
    Dim j As Integer
    Dim x1 As Integer
    Dim x2 As Integer
    Dim y1 As Integer
    Dim y2 As Integer

    Dim VertNum As Byte

    On Error GoTo ErrorHandler

    isInSector2 = False

    For j = 1 To 3
        VertNum = j + 1
        If VertNum > 3 Then VertNum = 1
        x1 = PolyCoords(Index).vertex(j).X
        x2 = PolyCoords(Index).vertex(VertNum).X
        y1 = PolyCoords(Index).vertex(j).Y
        y2 = PolyCoords(Index).vertex(VertNum).Y

        If segmentsIntersect(x1, y1, x2, y2, X, Y, X + div, Y) Then
            isInSector2 = True
        ElseIf segmentsIntersect(x1, y1, x2, y2, X, Y, X, Y + div) Then
            isInSector2 = True
        ElseIf segmentsIntersect(x1, y1, x2, y2, X + div, Y, X + div, Y + div) Then
            isInSector2 = True
        ElseIf segmentsIntersect(x1, y1, x2, y2, X, Y + div, X + div, Y + div) Then
            isInSector2 = True
        End If
    Next

    Exit Function

ErrorHandler:

    MsgBox Error$

End Function

Private Function SegXHorizSeg(ByRef A1 As D3DVECTOR2, ByRef B1 As D3DVECTOR2, _
        ByRef A2 As D3DVECTOR2, ByRef length As Long) As Boolean

    Dim U As D3DVECTOR2
    Dim VX As Integer
    Dim D As Single
    Dim epsilon As Single

    SegXHorizSeg = False

    U.X = B1.X - A1.X
    U.Y = B1.Y - A1.Y
    D = -U.Y * length

    If (D = 0) Then  ' the poly line seg is also horizontal
        Exit Function
    End If

    Dim W As D3DVECTOR2
    Dim s As Single
    Dim T As Single

    W.X = A1.X - A2.X
    W.Y = A1.Y - A2.Y

    s = (length * W.Y) / D
    If (s <= 0 Or s >= 1) Then
        Exit Function
    End If

    T = (U.X * W.Y - U.Y * W.X) / D
    If (T <= 0 Or T >= 1) Then
        Exit Function
    End If

    SegXHorizSeg = True

End Function

Private Function SegXVertSeg(ByRef A1 As D3DVECTOR2, ByRef B1 As D3DVECTOR2, _
        ByRef A2 As D3DVECTOR2, ByRef length As Long) As Boolean

    Dim U As D3DVECTOR2
    Dim D As Single

    SegXVertSeg = False

    U.X = B1.X - A1.X  ' length of poly seg x
    U.Y = B1.Y - A1.Y  ' y
    D = U.X * length

    If (D = 0) Then  ' the poly line seg is also vertical
        Exit Function
    End If

    Dim W As D3DVECTOR2
    Dim s As Single
    Dim T As Single

    W.X = A1.X - A2.X
    W.Y = A1.Y - A2.Y

    s = (-length * W.X) / D
    If (s <= 0 Or s >= 1) Then
        Exit Function
    End If

    T = (U.X * W.Y - U.Y * W.X) / D
    If (T <= 0 Or T >= 1) Then
        Exit Function
    End If

    SegXVertSeg = True

End Function

Private Function segmentsIntersect(ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer, _
        ByVal A1 As Integer, ByVal B1 As Integer, ByVal A2 As Integer, ByVal B2 As Integer) As Boolean

    On Error GoTo ErrorHandler

    Dim DX As Long
    Dim dy As Long
    Dim da As Long
    Dim db As Long
    Dim T As Single
    Dim s As Single

    DX = x2 - x1
    dy = y2 - y1
    da = A2 - A1
    db = B2 - B1

    If (da * dy - db * DX) = 0 Then
        ' the segments are parallel
        segmentsIntersect = False
        Exit Function
    End If

    s = (DX * (B1 - y1) + dy * (x1 - A1)) / (da * dy - db * DX)
    T = (da * (y1 - B1) + db * (A1 - x1)) / (db * DX - da * dy)
    segmentsIntersect = (s >= 0 And s <= 1 And T >= 0 And T <= 1)

    Exit Function

ErrorHandler:

    MsgBox Error$

End Function

Private Function SegXSeg(ByRef A1 As D3DVECTOR2, ByRef B1 As D3DVECTOR2, _
        ByRef A2 As D3DVECTOR2, ByRef B2 As D3DVECTOR2) As Boolean

    Dim U As D3DVECTOR2
    Dim V As D3DVECTOR2
    Dim D As Single

    SegXSeg = False

    U.X = B1.X - A1.X
    U.Y = B1.Y - A1.Y
    V.X = B2.X - A2.X
    V.Y = B2.Y - A2.Y
    D = U.X * V.Y - U.Y * V.X

    If (D = 0) Then  ' the poly line seg is also horizontal
        Exit Function
    End If

    Dim W As D3DVECTOR2
    Dim s As Single
    Dim T As Single

    W.X = A1.X - A2.X
    W.Y = A1.Y - A2.Y

    s = (V.X * W.Y - V.Y * W.X) / D
    If (s <= 0# Or s >= 1#) Then
        Exit Function
    End If

    T = (U.X * W.Y - U.Y * W.X) / D
    If (T <= 0# Or T >= 1#) Then
        Exit Function
    End If

    SegXSeg = True

End Function

Private Function isBetween(p1, p2, p3) As Boolean

    isBetween = False

    If (p1 >= p2 And p2 >= p3) Or (p3 >= p2 And p2 >= p1) Then
        isBetween = True
    End If

End Function

Private Sub initGrid()

    On Error GoTo ErrorHandler

    Dim i As Integer

    Dim clrString As String
    Dim clr1 As Long
    Dim clr2 As Long

    clr1 = ARGB(gridOp1, gridColor1)
    clr2 = ARGB(gridOp2, gridColor2)

    ReDim xGridLines(gridDivisions)
    ReDim yGridLines(gridDivisions)

    xGridLines(1).vertex(1) = CreateCustomVertex(0, 0, 1, 1, clr1, 0, 0)
    xGridLines(1).vertex(2) = CreateCustomVertex(Me.ScaleWidth, 0, 1, 1, clr1, 0, 0)

    yGridLines(1).vertex(1) = CreateCustomVertex(0, 0, 1, 1, clr1, 0, 0)
    yGridLines(1).vertex(2) = CreateCustomVertex(0, Me.ScaleHeight, 1, 1, clr1, 0, 0)

    For i = 2 To gridDivisions
        xGridLines(i).vertex(1) = CreateCustomVertex(0, 0, 1, 1, clr2, 0, 0)
        xGridLines(i).vertex(2) = CreateCustomVertex(Me.ScaleWidth, 0, 1, 1, clr2, 0, 0)
        yGridLines(i).vertex(1) = CreateCustomVertex(0, 0, 1, 1, clr2, 0, 0)
        yGridLines(i).vertex(2) = CreateCustomVertex(0, Me.ScaleHeight, 1, 1, clr2, 0, 0)
    Next

    inc = (gridSpacing / gridDivisions)

    Exit Sub

ErrorHandler:

    MsgBox "Error initializing grid"

End Sub

Private Sub setGrid()

    Dim xGridOffset As Single
    Dim yGridOffset As Single
    Dim i As Integer

    xGridOffset = (scrollCoords(2).X - (Int(scrollCoords(2).X / gridSpacing) * gridSpacing)) * zoomFactor
    yGridOffset = (scrollCoords(2).Y - (Int(scrollCoords(2).Y / gridSpacing) * gridSpacing)) * zoomFactor

    xGridLines(1).vertex(1).Y = 0 - yGridOffset
    xGridLines(1).vertex(2).Y = 0 - yGridOffset

    yGridLines(1).vertex(1).X = 0 - xGridOffset
    yGridLines(1).vertex(2).X = 0 - xGridOffset

    For i = 2 To gridDivisions
        xGridLines(i).vertex(1).Y = xGridLines(1).vertex(1).Y + (gridSpacing / gridDivisions) * zoomFactor * (i - 1)
        xGridLines(i).vertex(2).Y = xGridLines(i).vertex(1).Y
        yGridLines(i).vertex(1).X = yGridLines(1).vertex(1).X + (gridSpacing / gridDivisions) * zoomFactor * (i - 1)
        yGridLines(i).vertex(2).X = yGridLines(i).vertex(1).X
    Next

End Sub

Private Function CreateCustomVertex(ByVal X As Single, ByVal Y As Single, Z As Single, rhw As Single, color As Long, _
                                            tu As Single, tv As Single) As TCustomVertex

    CreateCustomVertex.X = X
    CreateCustomVertex.Y = Y
    CreateCustomVertex.Z = Z
    CreateCustomVertex.rhw = rhw
    CreateCustomVertex.color = color
    CreateCustomVertex.tu = tu
    CreateCustomVertex.tv = tv

End Function

Private Function ExModeActive() As Boolean

    Dim TestCoopRes As Long

    TestCoopRes = D3DDevice.TestCooperativeLevel

    If (TestCoopRes = D3D_OK) Then
        ExModeActive = True
    Else
        ExModeActive = False
    End If

End Function

Public Sub Render()

    If Not initialized Or noRedraw Then Exit Sub

    Dim i As Integer
    Dim j As Integer
    Dim lineCoords(1 To 4) As TCustomVertex
    Dim sceneryCoords(4) As TCustomVertex
    Dim circleCoords(0 To 32) As TCustomVertex
    Dim numPolys As Integer
    Dim scenR As Single
    Dim backtypePolys() As TPolygon

    Dim xVal As Single
    Dim yVal As Single
    Dim theta As Single
    Dim R As Single

    Dim srcRect As RECT
    Dim rc As D3DVECTOR2
    Dim sc As D3DVECTOR2
    Dim tr As D3DVECTOR2
    Dim sVal As Integer
    Dim objClr As Long


    Dim matView As D3DMATRIX
    Dim viewVector As D3DVECTOR
    Dim upVector As D3DVECTOR
    Dim atVector As D3DVECTOR
    Dim matProj As D3DMATRIX

    upVector.Y = -1
    atVector.Z = 1
    atVector.X = scrollCoords(2).X + Me.ScaleWidth / 2 / zoomFactor
    atVector.Y = (scrollCoords(2).Y + Me.ScaleHeight / 2 / zoomFactor)

    viewVector.X = scrollCoords(2).X + Me.ScaleWidth / 2 / zoomFactor
    viewVector.Y = (scrollCoords(2).Y + Me.ScaleHeight / 2 / zoomFactor)
    viewVector.Z = 0

    D3DXMatrixLookAtLH matView, viewVector, atVector, upVector
    D3DDevice.SetTransform D3DTS_VIEW, matView

    D3DXMatrixPerspectiveLH matProj, Me.ScaleWidth / zoomFactor, -Me.ScaleHeight / zoomFactor, -1, 0
    D3DDevice.SetTransform D3DTS_PROJECTION, matProj

    rc.X = 0
    rc.Y = 0

    srcRect.Left = 0
    srcRect.Top = 0

    For i = 1 To 4
        lineCoords(i).rhw = 1
        lineCoords(i).Z = 1
    Next

    initialized = False
    If ExModeActive Then  ' check if in focus
        initialized = True
    Else
        resetDevice
        initialized = True
    End If

    If numVerts > 0 And currentTool = TOOL_CREATE Then
        numPolys = mPolyCount + 1
    Else
        numPolys = mPolyCount
    End If

    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, backClr, 1#, 0

    D3DDevice.BeginScene
    ' ----

    D3DDevice.setTexture 0, Nothing

    ' draw background
    If showBG Then
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, bgPolys(1), Len(bgPolys(1))
    End If

    ' draw Polys
    If showPolys And numPolys > 0 Then
        If showTexture Then  ' set texture
            D3DDevice.setTexture 0, mapTexture
        End If

        D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        D3DDevice.SetRenderState D3DRS_COLORWRITEENABLE, D3DCOLORWRITEENABLE_BLUE Or D3DCOLORWRITEENABLE_GREEN Or D3DCOLORWRITEENABLE_RED
        D3DDevice.SetRenderState D3DRS_COLORWRITEENABLE, D3DCOLORWRITEENABLE_ALPHA Or D3DCOLORWRITEENABLE_BLUE Or D3DCOLORWRITEENABLE_GREEN Or D3DCOLORWRITEENABLE_RED

        D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1

        If clrPolys Then
            D3DDevice.SetRenderState D3DRS_SRCBLEND, polyBlendSrc
            D3DDevice.SetRenderState D3DRS_DESTBLEND, polyBlendDest
        End If

        For i = 1 To numPolys
            If vertexList(i).polyType = 24 Or vertexList(i).polyType = 25 Then
                D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 1, Polys(i).vertex(1), Len(Polys(1).vertex(1))
            End If
        Next

        D3DDevice.SetRenderState D3DRS_SRCBLEND, polyBlendSrc
        D3DDevice.SetRenderState D3DRS_DESTBLEND, polyBlendDest
    ElseIf showPolys = False And numPolys > 0 Then
        D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ZERO
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
        For i = 1 To numPolys
            If vertexList(i).polyType = 24 Or vertexList(i).polyType = 25 Then
                D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 1, Polys(i).vertex(1), Len(Polys(1).vertex(1))
            End If
        Next
        D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
    End If

    scenerySprite.Begin
    If sceneryCount > 0 And showScenery And sslBack Then
        For i = 1 To sceneryCount
            If Scenery(i).level = 0 Then
                sVal = Scenery(i).Style
                If Scenery(i).selected = 1 Then
                    If scaleDiff.X <> 1 Or scaleDiff.Y <> 1 Then
                        xVal = SceneryTextures(Scenery(i).Style).Width * Scenery(i).Scaling.X
                        yVal = SceneryTextures(Scenery(i).Style).Height * Scenery(i).Scaling.Y
                        theta = GetAngle(xVal, yVal) + Scenery(i).rotation
                        R = Sqr(xVal ^ 2 + yVal ^ 2)

                        xVal = Cos(theta) * R * scaleDiff.X
                        yVal = -Sin(theta) * R * scaleDiff.Y
                        theta = GetAngle(xVal, yVal) - Scenery(i).rotation
                        R = Sqr(xVal ^ 2 + yVal ^ 2)

                        sc.X = SceneryTextures(sVal).reScale.X * ((Cos(theta) * R) / (SceneryTextures(Scenery(i).Style).Width)) * zoomFactor
                        sc.Y = SceneryTextures(sVal).reScale.Y * (-(Sin(theta) * R) / (SceneryTextures(Scenery(i).Style).Height)) * zoomFactor
                        scenR = Scenery(i).rotation - rDiff
                    Else
                        sc.X = SceneryTextures(sVal).reScale.X * Scenery(i).Scaling.X * zoomFactor
                        sc.Y = SceneryTextures(sVal).reScale.Y * Scenery(i).Scaling.Y * zoomFactor
                        scenR = Scenery(i).rotation - rDiff
                    End If
                Else
                    sc.X = SceneryTextures(sVal).reScale.X * Scenery(i).Scaling.X * zoomFactor
                    sc.Y = SceneryTextures(sVal).reScale.Y * Scenery(i).Scaling.Y * zoomFactor
                    scenR = Scenery(i).rotation
                End If
                srcRect.Right = SceneryTextures(sVal).Width / SceneryTextures(sVal).reScale.X
                srcRect.bottom = SceneryTextures(sVal).Height / SceneryTextures(sVal).reScale.Y
                scenerySprite.Draw SceneryTextures(sVal).Texture, ByVal 0, sc, rc, scenR, Scenery(i).screenTr, Scenery(i).color
            End If
        Next
    End If

    If sceneryCount > 0 And showScenery And sslMid Then
        For i = 1 To sceneryCount
            If Scenery(i).level = 1 Then
                sVal = Scenery(i).Style
                If Scenery(i).selected = 1 Then
                    If scaleDiff.X <> 1 Or scaleDiff.Y <> 1 Then
                        xVal = SceneryTextures(Scenery(i).Style).Width * Scenery(i).Scaling.X
                        yVal = SceneryTextures(Scenery(i).Style).Height * Scenery(i).Scaling.Y
                        theta = GetAngle(xVal, yVal) + Scenery(i).rotation
                        R = Sqr(xVal ^ 2 + yVal ^ 2)

                        xVal = Cos(theta) * R * scaleDiff.X
                        yVal = -Sin(theta) * R * scaleDiff.Y
                        theta = GetAngle(xVal, yVal) - Scenery(i).rotation
                        R = Sqr(xVal ^ 2 + yVal ^ 2)

                        sc.X = SceneryTextures(sVal).reScale.X * ((Cos(theta) * R) / (SceneryTextures(Scenery(i).Style).Width)) * zoomFactor
                        sc.Y = SceneryTextures(sVal).reScale.Y * (-(Sin(theta) * R) / (SceneryTextures(Scenery(i).Style).Height)) * zoomFactor
                        scenR = Scenery(i).rotation - rDiff
                    Else
                        sc.X = SceneryTextures(sVal).reScale.X * Scenery(i).Scaling.X * zoomFactor
                        sc.Y = SceneryTextures(sVal).reScale.Y * Scenery(i).Scaling.Y * zoomFactor
                        scenR = Scenery(i).rotation - rDiff
                    End If
                Else
                    sc.X = SceneryTextures(sVal).reScale.X * Scenery(i).Scaling.X * zoomFactor
                    sc.Y = SceneryTextures(sVal).reScale.Y * Scenery(i).Scaling.Y * zoomFactor
                    scenR = Scenery(i).rotation
                End If
                srcRect.Right = SceneryTextures(sVal).Width / SceneryTextures(sVal).reScale.X
                srcRect.bottom = SceneryTextures(sVal).Height / SceneryTextures(sVal).reScale.Y
                scenerySprite.Draw SceneryTextures(sVal).Texture, ByVal 0, sc, rc, scenR, Scenery(i).screenTr, Scenery(i).color
            End If
        Next
    End If

    If currentFunction = TOOL_SCENERY And Not (ctrlDown Or altDown) Then
        If Scenery(0).level < 2 Then
            sVal = Scenery(0).Style
            sc.X = SceneryTextures(sVal).reScale.X * Scenery(0).Scaling.X * zoomFactor
            sc.Y = SceneryTextures(sVal).reScale.Y * Scenery(0).Scaling.Y * zoomFactor
            srcRect.Right = SceneryTextures(sVal).Width / SceneryTextures(sVal).reScale.X
            srcRect.bottom = SceneryTextures(sVal).Height / SceneryTextures(sVal).reScale.Y
            scenerySprite.Draw SceneryTextures(sVal).Texture, srcRect, sc, rc, Scenery(0).rotation, Scenery(0).screenTr, Scenery(0).color
        End If
    End If

    scenerySprite.End

    ' draw Polys
    If showPolys And numPolys > 0 Then
        If showTexture Then  ' set texture
            D3DDevice.setTexture 0, mapTexture
        End If

        D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        D3DDevice.SetRenderState D3DRS_COLORWRITEENABLE, D3DCOLORWRITEENABLE_BLUE Or D3DCOLORWRITEENABLE_GREEN Or D3DCOLORWRITEENABLE_RED
        D3DDevice.SetRenderState D3DRS_COLORWRITEENABLE, D3DCOLORWRITEENABLE_ALPHA Or D3DCOLORWRITEENABLE_BLUE Or D3DCOLORWRITEENABLE_GREEN Or D3DCOLORWRITEENABLE_RED

        D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1

        If clrPolys Then
            D3DDevice.SetRenderState D3DRS_SRCBLEND, polyBlendSrc
            D3DDevice.SetRenderState D3DRS_DESTBLEND, polyBlendDest
        End If

        For i = 1 To numPolys
            If Not (vertexList(i).polyType = 24 Or vertexList(i).polyType = 25) Then
                D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 1, Polys(i).vertex(1), Len(Polys(1).vertex(1))
            End If
        Next

        D3DDevice.SetRenderState D3DRS_SRCBLEND, polyBlendSrc
        D3DDevice.SetRenderState D3DRS_DESTBLEND, polyBlendDest

    ElseIf showPolys = False And numPolys > 0 Then
        D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ZERO
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
        For i = 1 To numPolys
            If Not (vertexList(i).polyType = 24 Or vertexList(i).polyType = 25) Then
                D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 1, Polys(i).vertex(1), Len(Polys(1).vertex(1))
            End If
        Next
        D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
    End If

    ' draw selected polys
    If numSelectedPolys > 0 And showPolys And Not (currentTool = TOOL_TEXTURE Or currentTool = TOOL_VCOLOR Or currentTool = TOOL_PCOLOR) Then
        D3DDevice.setTexture 0, patternTexture
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE

        D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
        For i = 1 To numSelectedPolys
            objClr = gPolyTypeClrs(vertexList(selectedPolys(i)).polyType)
            lineCoords(1) = Polys(selectedPolys(i)).vertex(1)
            lineCoords(2) = Polys(selectedPolys(i)).vertex(2)
            lineCoords(3) = Polys(selectedPolys(i)).vertex(3)

            lineCoords(1).tu = Polys(selectedPolys(i)).vertex(1).X / 128
            lineCoords(1).tv = Polys(selectedPolys(i)).vertex(1).Y / 128
            lineCoords(2).tu = Polys(selectedPolys(i)).vertex(2).X / 128
            lineCoords(2).tv = Polys(selectedPolys(i)).vertex(2).Y / 128
            lineCoords(3).tu = Polys(selectedPolys(i)).vertex(3).X / 128
            lineCoords(3).tv = Polys(selectedPolys(i)).vertex(3).Y / 128

            lineCoords(1).color = 0
            lineCoords(2).color = 0
            lineCoords(3).color = 0

            lineCoords(1).Z = 1
            lineCoords(2).Z = 1
            lineCoords(3).Z = 1
            lineCoords(1).rhw = 1
            lineCoords(2).rhw = 1
            lineCoords(3).rhw = 1
            If vertexList(selectedPolys(i)).vertex(1) = 1 Then lineCoords(1).color = objClr
            If vertexList(selectedPolys(i)).vertex(2) = 1 Then lineCoords(2).color = objClr
            If vertexList(selectedPolys(i)).vertex(3) = 1 Then lineCoords(3).color = objClr

            D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 1, lineCoords(1), Len(lineCoords(1))
        Next
        D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0
    End If

    ' draw depthmap
    If showPolys And currentTool = TOOL_DEPTHMAP Then
        lineCoords(1).tu = 0
        lineCoords(1).tv = 0
        lineCoords(2).tu = 0
        lineCoords(2).tv = 0
        lineCoords(3).tu = 0
        lineCoords(3).tv = 0
        lineCoords(1).Z = 1
        lineCoords(2).Z = 1
        lineCoords(3).Z = 1
        lineCoords(1).rhw = 1
        lineCoords(2).rhw = 1
        lineCoords(3).rhw = 1

        D3DDevice.setTexture 0, Nothing

        For i = 1 To mPolyCount
            lineCoords(1) = Polys(i).vertex(1)
            lineCoords(2) = Polys(i).vertex(2)
            lineCoords(3) = Polys(i).vertex(3)

            If Polys(i).vertex(1).Z >= 0 And Polys(i).vertex(2).Z >= 0 And Polys(i).vertex(3).Z >= 0 Then
                lineCoords(1).color = ARGB(255, RGB(Polys(i).vertex(1).Z, Polys(i).vertex(1).Z, Polys(i).vertex(1).Z))
                lineCoords(2).color = ARGB(255, RGB(Polys(i).vertex(2).Z, Polys(i).vertex(2).Z, Polys(i).vertex(2).Z))
                lineCoords(3).color = ARGB(255, RGB(Polys(i).vertex(3).Z, Polys(i).vertex(3).Z, Polys(i).vertex(3).Z))
            End If

            D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 1, lineCoords(1), Len(lineCoords(1))
        Next
    End If

    ' draw scenery
    scenerySprite.Begin
    If sceneryCount > 0 And showScenery And sslFront Then
        For i = 1 To sceneryCount
            If Scenery(i).level = 2 Then
                sVal = Scenery(i).Style
                If Scenery(i).selected = 1 Then
                    If scaleDiff.X <> 1 Or scaleDiff.Y <> 1 Then
                        xVal = SceneryTextures(Scenery(i).Style).Width * Scenery(i).Scaling.X
                        yVal = SceneryTextures(Scenery(i).Style).Height * Scenery(i).Scaling.Y
                        theta = GetAngle(xVal, yVal) + Scenery(i).rotation
                        R = Sqr(xVal ^ 2 + yVal ^ 2)

                        xVal = Cos(theta) * R * scaleDiff.X
                        yVal = -Sin(theta) * R * scaleDiff.Y
                        theta = GetAngle(xVal, yVal) - Scenery(i).rotation
                        R = Sqr(xVal ^ 2 + yVal ^ 2)

                        sc.X = SceneryTextures(sVal).reScale.X * ((Cos(theta) * R) / (SceneryTextures(Scenery(i).Style).Width)) * zoomFactor
                        sc.Y = SceneryTextures(sVal).reScale.Y * (-(Sin(theta) * R) / (SceneryTextures(Scenery(i).Style).Height)) * zoomFactor
                        scenR = Scenery(i).rotation - rDiff
                    Else
                        sc.X = SceneryTextures(sVal).reScale.X * Scenery(i).Scaling.X * zoomFactor
                        sc.Y = SceneryTextures(sVal).reScale.Y * Scenery(i).Scaling.Y * zoomFactor
                        scenR = Scenery(i).rotation - rDiff
                    End If
                Else
                    sc.X = SceneryTextures(sVal).reScale.X * Scenery(i).Scaling.X * zoomFactor
                    sc.Y = SceneryTextures(sVal).reScale.Y * Scenery(i).Scaling.Y * zoomFactor
                    scenR = Scenery(i).rotation
                End If
                srcRect.Right = SceneryTextures(sVal).Width / SceneryTextures(sVal).reScale.X
                srcRect.bottom = SceneryTextures(sVal).Height / SceneryTextures(sVal).reScale.Y
                scenerySprite.Draw SceneryTextures(sVal).Texture, ByVal 0, sc, rc, scenR, Scenery(i).screenTr, Scenery(i).color
            End If
        Next
    End If

    ' draw current scenery
    If currentFunction = TOOL_SCENERY And Not (ctrlDown Or altDown) Then
        If Scenery(0).level = 2 Then
            sVal = Scenery(0).Style
            sc.X = SceneryTextures(sVal).reScale.X * Scenery(0).Scaling.X * zoomFactor
            sc.Y = SceneryTextures(sVal).reScale.Y * Scenery(0).Scaling.Y * zoomFactor
            srcRect.Right = SceneryTextures(sVal).Width / SceneryTextures(sVal).reScale.X + 0
            srcRect.bottom = SceneryTextures(sVal).Height / SceneryTextures(sVal).reScale.Y + 0
            scenerySprite.Draw SceneryTextures(sVal).Texture, srcRect, sc, rc, Scenery(0).rotation, Scenery(0).screenTr, Scenery(0).color
        End If
    End If

    ' draw objects
    objClr = ARGB(255, RGB(255, 255, 255))
    sc.X = 32 / (objTexSize.X / 8)
    sc.Y = 32 / (objTexSize.Y / 4)
    rc.X = (objTexSize.X / 8) / 2
    rc.Y = (objTexSize.Y / 4) / 2
    If showObjects Then
        If spawnPoints > 0 Then
            For i = 1 To spawnPoints
                tr.X = Int((Spawns(i).X - scrollCoords(2).X) * zoomFactor - 15 + 0.5)
                tr.Y = Int((Spawns(i).Y - scrollCoords(2).Y) * zoomFactor - 15 + 0.5)
                srcRect.Top = Int(Spawns(i).Team / 8) * (objTexSize.Y / 4)
                srcRect.Left = (Spawns(i).Team - (Int(Spawns(i).Team / 8) * 8)) * (objTexSize.X / 8)
                srcRect.Right = srcRect.Left + (objTexSize.X / 8)
                srcRect.bottom = srcRect.Top + (objTexSize.Y / 4)
                If Spawns(i).active = 1 Then
                    scenerySprite.Draw objectsTexture, srcRect, sc, rc, 0, tr, ARGB(255, selectionColor)
                Else
                    scenerySprite.Draw objectsTexture, srcRect, sc, rc, 0, tr, objClr
                End If
            Next
        End If
        If colliderCount > 0 Then
            objClr = ARGB(128, RGB(255, 255, 255))
            For i = 1 To colliderCount
                sc.X = Colliders(i).radius / (objTexSize.X / 8) * zoomFactor
                sc.Y = Colliders(i).radius / (objTexSize.Y / 4) * zoomFactor
                tr.X = Int((Colliders(i).X - scrollCoords(2).X) * zoomFactor - (objTexSize.X / 8) / 2 * sc.X + 0.5)
                tr.Y = Int((Colliders(i).Y - scrollCoords(2).Y) * zoomFactor - (objTexSize.Y / 4) / 2 * sc.Y + 0.5)
                If Colliders(i).active = 1 Then
                    srcRect.Left = 0
                    srcRect.Top = (objTexSize.Y / 4) * 3
                    srcRect.Right = srcRect.Left + (objTexSize.X / 8)
                    srcRect.bottom = srcRect.Top + (objTexSize.Y / 4)
                    scenerySprite.Draw objectsTexture, srcRect, sc, rc, 0, tr, objClr
                Else
                    srcRect.Left = (objTexSize.X / 8)
                    srcRect.Top = (objTexSize.Y / 4) * 2
                    srcRect.Right = srcRect.Left + (objTexSize.X / 8)
                    srcRect.bottom = srcRect.Top + (objTexSize.Y / 4)
                    scenerySprite.Draw objectsTexture, srcRect, sc, rc, 0, tr, objClr
                End If
            Next
        End If
    End If

    If showLights Then
        objClr = ARGB(255, RGB(255, 255, 255))
        sc.X = 32 / (objTexSize.X / 8)
        sc.Y = 32 / (objTexSize.Y / 4)
        rc.X = (objTexSize.X / 8) / 2
        rc.Y = (objTexSize.Y / 4) / 2
        If lightCount > 0 Then
            srcRect.Left = (objTexSize.X / 8) * 7
            srcRect.Top = (objTexSize.Y / 4) * 2
            srcRect.Right = srcRect.Left + (objTexSize.X / 8)
            srcRect.bottom = srcRect.Top + (objTexSize.Y / 4)
            For i = 1 To lightCount
                objClr = ARGB(255, RGB(Lights(i).color.blue, Lights(i).color.green, Lights(i).color.red))
                sc.X = 32 / (objTexSize.X / 8)
                sc.Y = 32 / (objTexSize.Y / 4)
                tr.X = Int((Lights(i).X - scrollCoords(2).X) * zoomFactor - 16 * sc.X + 0.5)
                tr.Y = Int((Lights(i).Y - scrollCoords(2).Y) * zoomFactor - 16 * sc.Y + 0.5)
                If Lights(i).selected = 1 Then
                    scenerySprite.Draw objectsTexture, srcRect, sc, rc, 0, tr, ARGB(255, selectionColor)
                Else
                    scenerySprite.Draw objectsTexture, srcRect, sc, rc, 0, tr, objClr
                End If
            Next
        End If
    End If

    ' draw current object
    If currentTool = TOOL_OBJECTS And Not (ctrlDown Or altDown) Then
        objClr = ARGB(128, RGB(255, 255, 255))
        If mnuGostek.Checked Then  ' gostek
            sc.X = 32 / (objTexSize.X / 8) * zoomFactor
            sc.Y = 32 / (objTexSize.Y / 4) * zoomFactor
            srcRect.Left = (objTexSize.X / 8) * 2 + 1
            srcRect.Top = (objTexSize.Y / 4) * 2
            srcRect.Right = srcRect.Left + (objTexSize.X / 8) - 2
            srcRect.bottom = srcRect.Top + (objTexSize.Y / 4)
            tr.X = mouseCoords.X - 16 * zoomFactor
            tr.Y = mouseCoords.Y - 16 * zoomFactor
            scenerySprite.Draw objectsTexture, srcRect, sc, rc, 0, tr, objClr
        ElseIf mnuCollider.Checked = True Then  ' collider
            srcRect.Left = (objTexSize.X / 8)
            srcRect.Top = (objTexSize.Y / 4) * 2
            srcRect.Right = srcRect.Left + (objTexSize.X / 8)
            srcRect.bottom = srcRect.Top + (objTexSize.Y / 4)
            sc.X = Colliders(0).radius / (objTexSize.X / 8) * zoomFactor
            sc.Y = Colliders(0).radius / (objTexSize.Y / 4) * zoomFactor
            tr.X = Colliders(0).X - (objTexSize.X / 8) / 2 * sc.X
            tr.Y = Colliders(0).Y - (objTexSize.Y / 4) / 2 * sc.Y
            scenerySprite.Draw objectsTexture, srcRect, sc, rc, 0, tr, objClr
        Else  ' spawn
            sc.X = 32 / (objTexSize.X / 8)
            sc.Y = 32 / (objTexSize.Y / 4)
            tr.X = Spawns(0).X - 15
            tr.Y = Spawns(0).Y - 15
            srcRect.Top = Int(Spawns(0).Team / 8) * (objTexSize.Y / 4)
            srcRect.Left = (Spawns(0).Team - (Int(Spawns(0).Team / 8) * 8)) * (objTexSize.X / 8)
            srcRect.Right = srcRect.Left + (objTexSize.X / 8)
            srcRect.bottom = srcRect.Top + (objTexSize.Y / 4)
            scenerySprite.Draw objectsTexture, srcRect, sc, rc, 0, tr, objClr
        End If
    End If

    ' draw gostek
    If gostek.X <> 0 Or gostek.Y <> 0 Then
        sc.X = 32 / (objTexSize.X / 8) * zoomFactor
        sc.Y = 32 / (objTexSize.Y / 4) * zoomFactor
        srcRect.Left = ((objTexSize.X / 8) * 2) + 1
        srcRect.Top = (objTexSize.Y / 4) * 2
        srcRect.Right = srcRect.Left + (objTexSize.X / 8) - 2
        srcRect.bottom = srcRect.Top + (objTexSize.Y / 4)
        tr.X = (gostek.X - 16 - scrollCoords(2).X) * zoomFactor
        tr.Y = (gostek.Y - 16 - scrollCoords(2).Y) * zoomFactor
        scenerySprite.Draw objectsTexture, srcRect, sc, rc, 0, tr, ARGB(255, RGB(128, 128, 128))
    End If

    scenerySprite.End

    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.setTexture 0, Nothing

    ' draw grid
    If showGrid Then
        D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True

        setGrid

        For i = 0 To (Int((Me.ScaleWidth / gridSpacing) / zoomFactor) + 1)
            If inc * zoomFactor >= 8 Then
                D3DDevice.DrawPrimitiveUP D3DPT_LINELIST, gridDivisions, yGridLines(1).vertex(1), Len(yGridLines(1).vertex(1))
            ElseIf gridSpacing * zoomFactor >= 8 Then
                D3DDevice.DrawPrimitiveUP D3DPT_LINELIST, 1, yGridLines(1).vertex(1), Len(yGridLines(1).vertex(1))
            End If

            For j = 1 To gridDivisions
                yGridLines(j).vertex(1).X = yGridLines(j).vertex(1).X + gridSpacing * zoomFactor
                yGridLines(j).vertex(2).X = yGridLines(j).vertex(1).X
            Next
        Next
        For i = 0 To (Int((Me.ScaleHeight / gridSpacing) / zoomFactor) + 1)
            If inc * zoomFactor >= 8 Then
                D3DDevice.DrawPrimitiveUP D3DPT_LINELIST, gridDivisions, xGridLines(1).vertex(1), Len(xGridLines(1).vertex(1))
            ElseIf gridSpacing * zoomFactor >= 8 Then
                D3DDevice.DrawPrimitiveUP D3DPT_LINELIST, 1, xGridLines(1).vertex(1), Len(xGridLines(1).vertex(1))
            End If

            For j = 1 To gridDivisions
                xGridLines(j).vertex(1).Y = xGridLines(j).vertex(1).Y + gridSpacing * zoomFactor
                xGridLines(j).vertex(2).Y = xGridLines(j).vertex(1).Y
            Next
        Next
    End If

    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False

    If clrWireframe And (showWireframe Or showPoints) Then
        D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
        D3DDevice.SetRenderState D3DRS_SRCBLEND, wireBlendSrc
        D3DDevice.SetRenderState D3DRS_DESTBLEND, wireBlendDest
    End If

    ' draw wireframe
    If showWireframe And mPolyCount > 0 Then
        D3DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_WIREFRAME
        For i = 1 To mPolyCount
            D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 1, Polys(i).vertex(1), Len(Polys(1).vertex(1))
        Next
        D3DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
    End If

    ' draw scenery boxes
    sc.X = 1
    sc.Y = 1
    srcRect.Right = 8
    srcRect.bottom = 8
    If sceneryCount > 0 And showScenery Then
        For i = 1 To sceneryCount
            sVal = Scenery(i).Style

            sceneryCoords(0) = CreateCustomVertex(0, 0, 1, 1, ARGB(255, Scenery(i).color), 0, 0)
            If Scenery(i).selected = 1 Or Scenery(i).selected = 3 Then
                sceneryCoords(0).color = ARGB(255, pointColor)
            End If
            sceneryCoords(1) = sceneryCoords(0)
            sceneryCoords(2) = sceneryCoords(0)
            sceneryCoords(3) = sceneryCoords(0)
            sceneryCoords(0).X = Scenery(i).screenTr.X
            sceneryCoords(0).Y = Scenery(i).screenTr.Y

            If Scenery(i).selected = 1 And ctrlDown And (scaleDiff.X <> 1 Or scaleDiff.Y <> 1) Then
                xVal = SceneryTextures(Scenery(i).Style).Width * Scenery(i).Scaling.X
                yVal = SceneryTextures(Scenery(i).Style).Height * Scenery(i).Scaling.Y
                theta = GetAngle(xVal, yVal) + Scenery(i).rotation
                R = Sqr(xVal ^ 2 + yVal ^ 2)

                xVal = Cos(theta) * R * scaleDiff.X
                yVal = -Sin(theta) * R * scaleDiff.Y
                theta = GetAngle(xVal, yVal) - Scenery(i).rotation
                R = Sqr(xVal ^ 2 + yVal ^ 2)

                sc.X = (Cos(theta) * R)
                sc.Y = -(Sin(theta) * R)

                sceneryCoords(1).X = sceneryCoords(0).X + Cos(Scenery(i).rotation) * sc.X * zoomFactor
                sceneryCoords(1).Y = sceneryCoords(0).Y - Sin(Scenery(i).rotation) * sc.X * zoomFactor
                sceneryCoords(3).X = sceneryCoords(0).X + Sin(Scenery(i).rotation) * sc.Y * zoomFactor
                sceneryCoords(3).Y = sceneryCoords(0).Y + Cos(Scenery(i).rotation) * sc.Y * zoomFactor
            ElseIf Scenery(i).selected = 1 And (rDiff <> 0 Or (scaleDiff.X <> 0 Or scaleDiff.Y <> 0)) Then
                sceneryCoords(1).X = sceneryCoords(0).X + Cos(Scenery(i).rotation - rDiff) * (SceneryTextures(sVal).Width) * Scenery(i).Scaling.X * zoomFactor
                sceneryCoords(1).Y = sceneryCoords(0).Y - Sin(Scenery(i).rotation - rDiff) * (SceneryTextures(sVal).Width) * Scenery(i).Scaling.X * zoomFactor
                sceneryCoords(3).X = sceneryCoords(0).X + Sin(Scenery(i).rotation - rDiff) * (SceneryTextures(sVal).Height) * Scenery(i).Scaling.Y * zoomFactor
                sceneryCoords(3).Y = sceneryCoords(0).Y + Cos(Scenery(i).rotation - rDiff) * (SceneryTextures(sVal).Height) * Scenery(i).Scaling.Y * zoomFactor
            Else
                sceneryCoords(1).X = sceneryCoords(0).X + Cos(Scenery(i).rotation) * (SceneryTextures(sVal).Width) * Scenery(i).Scaling.X * zoomFactor
                sceneryCoords(1).Y = sceneryCoords(0).Y - Sin(Scenery(i).rotation) * (SceneryTextures(sVal).Width) * Scenery(i).Scaling.X * zoomFactor
                sceneryCoords(3).X = sceneryCoords(0).X + Sin(Scenery(i).rotation) * (SceneryTextures(sVal).Height) * Scenery(i).Scaling.Y * zoomFactor
                sceneryCoords(3).Y = sceneryCoords(0).Y + Cos(Scenery(i).rotation) * (SceneryTextures(sVal).Height) * Scenery(i).Scaling.Y * zoomFactor
            End If

            sceneryCoords(2).X = sceneryCoords(3).X + sceneryCoords(1).X - sceneryCoords(0).X
            sceneryCoords(2).Y = sceneryCoords(3).Y + sceneryCoords(1).Y - sceneryCoords(0).Y
            sceneryCoords(4) = sceneryCoords(0)

            If showWireframe Or Scenery(i).selected = 1 Or Scenery(i).selected = 3 Then
                D3DDevice.DrawPrimitiveUP D3DPT_LINESTRIP, 4, sceneryCoords(0), Len(sceneryCoords(0))
            End If

            If showPoints Or Scenery(i).selected = 1 Or Scenery(i).selected = 3 Then
                If sceneryVerts Then
                    D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, 4, sceneryCoords(0), Len(sceneryCoords(0))
                Else
                    D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, 1, sceneryCoords(0), Len(sceneryCoords(0))
                End If
            End If
        Next
        If currentTool = TOOL_SCENERY And Scenery(0).Style > 0 And Not (ctrlDown Or altDown) Then
            sVal = Scenery(0).Style
            sceneryCoords(0) = CreateCustomVertex(0, 0, 1, 1, Scenery(0).color, 0, 0)
            sceneryCoords(1) = CreateCustomVertex(0, 0, 1, 1, Scenery(0).color, 0, 0)
            sceneryCoords(2) = CreateCustomVertex(0, 0, 1, 1, Scenery(0).color, 0, 0)
            sceneryCoords(3) = CreateCustomVertex(0, 0, 1, 1, Scenery(0).color, 0, 0)
            sceneryCoords(0).X = Scenery(0).screenTr.X
            sceneryCoords(0).Y = Scenery(0).screenTr.Y
            sceneryCoords(1).X = sceneryCoords(0).X + Cos(Scenery(0).rotation) * (SceneryTextures(sVal).Width + 0) * Scenery(0).Scaling.X * zoomFactor
            sceneryCoords(1).Y = sceneryCoords(0).Y - Sin(Scenery(0).rotation) * (SceneryTextures(sVal).Width + 0) * Scenery(0).Scaling.X * zoomFactor
            sceneryCoords(3).X = sceneryCoords(0).X + Sin(Scenery(0).rotation) * (SceneryTextures(sVal).Height + 0) * Scenery(0).Scaling.Y * zoomFactor
            sceneryCoords(3).Y = sceneryCoords(0).Y + Cos(Scenery(0).rotation) * (SceneryTextures(sVal).Height + 0) * Scenery(0).Scaling.Y * zoomFactor
            sceneryCoords(2).X = sceneryCoords(3).X + sceneryCoords(1).X - sceneryCoords(0).X
            sceneryCoords(2).Y = sceneryCoords(3).Y + sceneryCoords(1).Y - sceneryCoords(0).Y
            sceneryCoords(4) = sceneryCoords(0)

            D3DDevice.DrawPrimitiveUP D3DPT_LINESTRIP, 4, sceneryCoords(0), Len(sceneryCoords(0))
        End If
    End If

    If numVerts > 0 And currentTool = TOOL_CREATE Then
        D3DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_WIREFRAME
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 1, Polys(mPolyCount + 1).vertex(1), Len(Polys(mPolyCount + 1).vertex(1))
        D3DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
    End If

    D3DDevice.setTexture 0, particleTexture

    ' draw points
    If showPoints And numPolys > 0 Then
        D3DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_POINT

        For i = 1 To numPolys
            lineCoords(1) = Polys(i).vertex(1)
            lineCoords(2) = Polys(i).vertex(2)
            lineCoords(3) = Polys(i).vertex(3)

            lineCoords(1).Z = 1
            lineCoords(2).Z = 1
            lineCoords(3).Z = 1
            lineCoords(1).rhw = 1
            lineCoords(2).rhw = 1
            lineCoords(3).rhw = 1

            D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, 3, lineCoords(1), Len(lineCoords(1))
        Next

        D3DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
    End If

    If showPoints And showObjects And colliderCount > 0 Then
        sceneryCoords(0) = CreateCustomVertex(0, 0, 1, 1, Scenery(0).color, 0, 0)
        For i = 1 To colliderCount
            sceneryCoords(0).X = (Colliders(i).X - scrollCoords(2).X) * zoomFactor
            sceneryCoords(0).Y = (Colliders(i).Y - scrollCoords(2).Y) * zoomFactor
            If Colliders(i).active = 1 Then
                sceneryCoords(0).color = selectionColor
            Else
                sceneryCoords(0).color = ARGB(255, RGB(255, 255, 255))
            End If
            D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, 1, sceneryCoords(0), Len(sceneryCoords(0))
        Next
    End If

    ' draw selected poly wireframes
    D3DDevice.setTexture 0, Nothing
    If numSelectedPolys > 0 Then
        D3DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_WIREFRAME
        For i = 1 To numSelectedPolys
            lineCoords(1) = Polys(selectedPolys(i)).vertex(1)
            lineCoords(2) = Polys(selectedPolys(i)).vertex(2)
            lineCoords(3) = Polys(selectedPolys(i)).vertex(3)

            lineCoords(1).Z = 1: lineCoords(1).rhw = 1
            lineCoords(2).Z = 1: lineCoords(2).rhw = 1
            lineCoords(3).Z = 1: lineCoords(3).rhw = 1

            If vertexList(selectedPolys(i)).vertex(1) = 1 Or vertexList(selectedPolys(i)).vertex(1) = 3 Then
                lineCoords(1).color = pointColor
            End If
            If vertexList(selectedPolys(i)).vertex(2) = 1 Or vertexList(selectedPolys(i)).vertex(2) = 3 Then
                lineCoords(2).color = pointColor
            End If
            If vertexList(selectedPolys(i)).vertex(3) = 1 Or vertexList(selectedPolys(i)).vertex(3) = 3 Then
                lineCoords(3).color = pointColor
            End If

            D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 1, lineCoords(1), Len(lineCoords(1))

            If showPoints Then
                If vertexList(selectedPolys(i)).vertex(1) = 1 Then lineCoords(1).color = pointColor
                If vertexList(selectedPolys(i)).vertex(2) = 1 Then lineCoords(2).color = pointColor
                If vertexList(selectedPolys(i)).vertex(3) = 1 Then lineCoords(3).color = pointColor
                D3DDevice.setTexture 0, particleTexture
                D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, 3, lineCoords(1), Len(lineCoords(1))
                D3DDevice.setTexture 0, Nothing
            End If
        Next
        D3DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
    End If
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 0

    ' draw selection rect
    If currentTool = TOOL_MOVE And (numSelectedPolys > 0 Or numSelectedScenery > 0) And noneSelected = False Then
        objClr = &H80FFFFFF

        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True

        D3DDevice.setTexture 0, lineTexture

        sceneryCoords(0) = CreateCustomVertex(0, 0, 1, 1, objClr, 0, 0)
        sceneryCoords(1) = CreateCustomVertex(0, 0, 1, 1, objClr, 0, 0)
        sceneryCoords(2) = CreateCustomVertex(0, 0, 1, 1, objClr, 0, 0)
        sceneryCoords(3) = CreateCustomVertex(0, 0, 1, 1, objClr, 0, 0)

        If rDiff <> 0 Then
            For i = 0 To 3
                xVal = (selRect(i).X - rCenter.X)
                yVal = (selRect(i).Y - rCenter.Y)
                R = Sqr(xVal ^ 2 + yVal ^ 2)
                theta = GetAngle(xVal, yVal) - rDiff
                sceneryCoords(i).X = (rCenter.X + R * Cos(theta) - scrollCoords(2).X) * zoomFactor
                sceneryCoords(i).Y = (rCenter.Y + R * -Sin(theta) - scrollCoords(2).Y) * zoomFactor
            Next
        ElseIf scaleDiff.X <> 1 Or scaleDiff.Y <> 1 Then
            For i = 0 To 3
                sceneryCoords(i).X = (rCenter.X + ((selRect(i).X - rCenter.X) * scaleDiff.X) - scrollCoords(2).X) * zoomFactor
                sceneryCoords(i).Y = (rCenter.Y + ((selRect(i).Y - rCenter.Y) * scaleDiff.Y) - scrollCoords(2).Y) * zoomFactor
            Next
        Else
            For i = 0 To 3
                sceneryCoords(i).X = (selRect(i).X - scrollCoords(2).X) * zoomFactor
                sceneryCoords(i).Y = (selRect(i).Y - scrollCoords(2).Y) * zoomFactor
            Next
        End If

        sceneryCoords(0).tu = 0
        sceneryCoords(0).tv = 0
        sceneryCoords(1).tu = Sqr((sceneryCoords(1).X - sceneryCoords(0).X) ^ 2 + (sceneryCoords(1).Y - sceneryCoords(0).Y) ^ 2) / 64
        sceneryCoords(1).tv = 0
        sceneryCoords(2).tu = sceneryCoords(1).tu
        sceneryCoords(2).tv = Sqr((sceneryCoords(2).X - sceneryCoords(1).X) ^ 2 + (sceneryCoords(2).Y - sceneryCoords(1).Y) ^ 2) / 64
        sceneryCoords(3).tu = 0
        sceneryCoords(3).tv = sceneryCoords(2).tv

        sceneryCoords(4) = sceneryCoords(0)

        D3DDevice.DrawPrimitiveUP D3DPT_LINESTRIP, 4, sceneryCoords(0), Len(sceneryCoords(0))
        D3DDevice.setTexture 0, Nothing
        D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, 4, sceneryCoords(0), Len(sceneryCoords(0))

        For i = 0 To 3
            sceneryCoords(i).X = Midpoint(sceneryCoords(i).X, sceneryCoords(i + 1).X)
            sceneryCoords(i).Y = Midpoint(sceneryCoords(i).Y, sceneryCoords(i + 1).Y)
        Next
        D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, 4, sceneryCoords(0), Len(sceneryCoords(0))

        If Not mnuFixedRCenter.Checked Then
            sceneryCoords(0).X = (rCenter.X - scrollCoords(2).X) * zoomFactor
            sceneryCoords(0).Y = (rCenter.Y - scrollCoords(2).Y) * zoomFactor
            D3DDevice.setTexture 0, rCenterTexture
            D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, 1, sceneryCoords(0), Len(sceneryCoords(0))
        End If

        D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
    End If

    If showWaypoints Then
        objClr = &HFFFFFFFF
        For i = 1 To waypointCount
            If (Waypoints(i).pathNum = 1 And frmWaypoints.showPaths <> 2) Or (Waypoints(i).pathNum = 2 And frmWaypoints.showPaths <> 1) Then
                If Waypoints(i).selected = True Then
                    If Waypoints(i).pathNum = 1 Then
                        srcRect.Left = (objTexSize.X / 8) * 5
                    Else
                        srcRect.Left = (objTexSize.X / 8) * 6
                    End If
                Else
                    If Waypoints(i).pathNum = 1 Then
                        srcRect.Left = (objTexSize.X / 8) * 3
                    Else
                        srcRect.Left = (objTexSize.X / 8) * 4
                    End If
                End If
                sc.X = 32 / (objTexSize.X / 8)
                sc.Y = 32 / (objTexSize.Y / 4)
                srcRect.Top = (objTexSize.Y / 4) * 2
                srcRect.Right = srcRect.Left + (objTexSize.X / 8)
                srcRect.bottom = srcRect.Top + (objTexSize.Y / 4)
                tr.X = Int((Waypoints(i).X - scrollCoords(2).X) * zoomFactor - 15 + 0.5)
                tr.Y = Int((Waypoints(i).Y - scrollCoords(2).Y) * zoomFactor - 15 + 0.5)
                scenerySprite.Draw objectsTexture, srcRect, sc, rc, 0, tr, objClr
            End If
        Next

        D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
        D3DDevice.setTexture 0, pathTexture
        For i = 1 To conCount
            If (Waypoints(Connections(i).point1).pathNum = 1 And frmWaypoints.showPaths <> 2) _
                    Or (Waypoints(Connections(i).point1).pathNum = 2 And frmWaypoints.showPaths <> 1) _
                    Or (Waypoints(Connections(i).point2).pathNum = 1 And frmWaypoints.showPaths <> 2) _
                    Or (Waypoints(Connections(i).point2).pathNum = 2 And frmWaypoints.showPaths <> 1) Then
                lineCoords(1).X = (Waypoints(Connections(i).point1).X - scrollCoords(2).X) * zoomFactor
                lineCoords(1).Y = (Waypoints(Connections(i).point1).Y - scrollCoords(2).Y) * zoomFactor
                lineCoords(2).X = (Waypoints(Connections(i).point2).X - scrollCoords(2).X) * zoomFactor
                lineCoords(2).Y = (Waypoints(Connections(i).point2).Y - scrollCoords(2).Y) * zoomFactor
                If Waypoints(Connections(i).point2).wayType(2) Then
                    lineCoords(1).color = &HFFFFFF22
                    lineCoords(2).color = &HFFFFFF22
                ElseIf Waypoints(Connections(i).point2).wayType(3) Then
                    lineCoords(1).color = &HFF22FFFF
                    lineCoords(2).color = &HFF22FFFF
                ElseIf Waypoints(Connections(i).point2).wayType(0) Then
                    lineCoords(1).color = &HFF22FF22
                    lineCoords(2).color = &HFF22FF22
                ElseIf Waypoints(Connections(i).point2).wayType(1) Then
                    lineCoords(1).color = &HFFFF2222
                    lineCoords(2).color = &HFFFF2222
                ElseIf Waypoints(Connections(i).point2).wayType(4) Then
                    lineCoords(1).color = &HFFFFFFFF
                    lineCoords(2).color = &HFFFFFFFF
                Else
                    lineCoords(1).color = &HFF000000
                    lineCoords(2).color = &HFF000000
                End If
                lineCoords(1).tu = 0
                lineCoords(1).tv = 0
                lineCoords(2).tu = 1
                lineCoords(2).tv = 0

                D3DDevice.DrawPrimitiveUP D3DPT_LINESTRIP, 1, lineCoords(1), Len(lineCoords(1))
            End If
        Next

        If currentWaypoint > 0 Then
            lineCoords(1).X = (Waypoints(currentWaypoint).X - scrollCoords(2).X) * zoomFactor
            lineCoords(1).Y = (Waypoints(currentWaypoint).Y - scrollCoords(2).Y) * zoomFactor
            lineCoords(2).X = mouseCoords.X
            lineCoords(2).Y = mouseCoords.Y
            If mnuWayType(2).Checked Then
                lineCoords(1).color = &HFFFFFF22
                lineCoords(2).color = &HFFFFFF22
            ElseIf mnuWayType(3).Checked Then
                lineCoords(1).color = &HFF22FFFF
                lineCoords(2).color = &HFF22FFFF
            ElseIf mnuWayType(0).Checked Then
                lineCoords(1).color = &HFF22FF22
                lineCoords(2).color = &HFF22FF22
            ElseIf mnuWayType(1).Checked Then
                lineCoords(1).color = &HFFFF2222
                lineCoords(2).color = &HFFFF2222
            ElseIf mnuWayType(4).Checked Then
                lineCoords(1).color = &HFFFFFFFF
                lineCoords(2).color = &HFFFFFFFF
            Else
                lineCoords(1).color = &HFF000000
                lineCoords(2).color = &HFF000000
            End If
            lineCoords(1).tu = 0
            lineCoords(1).tv = 0
            lineCoords(2).tu = 1
            lineCoords(2).tv = 0

            D3DDevice.DrawPrimitiveUP D3DPT_LINESTRIP, 1, lineCoords(1), Len(lineCoords(1))
        End If

        D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
    End If

    If showSketch Then
        D3DDevice.SetVertexShader FVF2
        D3DDevice.setTexture 0, sketchTexture
        If sketchLines > 0 Then
            D3DDevice.DrawPrimitiveUP D3DPT_LINELIST, sketchLines, sketch(1).vertex(1), Len(sketch(1).vertex(1))
        End If
        D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
        If currentFunction = TOOL_SKETCH And shiftDown Then
            D3DDevice.DrawPrimitiveUP D3DPT_LINELIST, 1, sketch(0).vertex(1), Len(sketch(0).vertex(1))
        End If
        D3DDevice.SetVertexShader FVF
    End If

    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_INVDESTCOLOR
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True

    ' draw circle
    If circleOn Then
        For i = 0 To 32
            circleCoords(i).color = ARGB(255, RGB(255, 255, 255))
            circleCoords(i).X = mouseCoords.X + zoomFactor * clrRadius * Math.Cos(PI * i / 16)
            circleCoords(i).Y = mouseCoords.Y + zoomFactor * clrRadius * Math.Sin(PI * i / 16)
        Next
        D3DDevice.DrawPrimitiveUP D3DPT_LINESTRIP, 32, circleCoords(0), Len(circleCoords(0))
    End If

    ' vertex selection
    If currentFunction = TOOL_VSELECT Or currentFunction = TOOL_VSELADD Or currentFunction = TOOL_VSELSUB Then
        If toolAction Then
            circleCoords(0).color = ARGB(255, RGB(255, 255, 255))
            circleCoords(1).color = ARGB(255, RGB(255, 255, 255))
            circleCoords(2).color = ARGB(255, RGB(255, 255, 255))
            circleCoords(3).color = ARGB(255, RGB(255, 255, 255))
            circleCoords(4).color = ARGB(255, RGB(255, 255, 255))
            circleCoords(0).X = selectedCoords(1).X
            circleCoords(1).X = mouseCoords.X
            circleCoords(2).X = mouseCoords.X
            circleCoords(3).X = selectedCoords(1).X
            circleCoords(4).X = selectedCoords(1).X
            circleCoords(0).Y = selectedCoords(1).Y
            circleCoords(1).Y = selectedCoords(1).Y
            circleCoords(2).Y = mouseCoords.Y
            circleCoords(3).Y = mouseCoords.Y
            circleCoords(4).Y = selectedCoords(1).Y
            D3DDevice.DrawPrimitiveUP D3DPT_LINESTRIP, 4, circleCoords(0), Len(circleCoords(0))
        End If
    End If

    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False

    ' ----
    D3DDevice.EndScene

    D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0

    eraseCircle = False

    Exit Sub

ErrorHandler:

    MsgBox "Error Rendering with Direct3D" & vbNewLine & D3DX.GetErrorString(err.Number)

End Sub

Function FtoDW(f As Single) As Long

    Dim buf As D3DXBuffer
    Dim l As Long

    Set buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData buf, 0, 4, 1, f
    D3DX.BufferGetData buf, 0, 4, 1, l
    FtoDW = l

End Function

Private Sub DirectXEvent8_DXCallback(ByVal eventid As Long)

    Dim i As Long
    Dim hotKeyPressed As Integer
    Dim wayptKeyPressed As Integer
    Dim layerKeyPressed As Integer
    Dim pBuffer(0 To BUFFER_SIZE) As DIDEVICEOBJECTDATA
    Static tempFunction As Byte

    On Error GoTo ErrorHandler

    If DIDevice Is Nothing Then Exit Sub

    If eventid = hEvent Then
        DIDevice.GetDeviceStateKeyboard DIState
        DIDevice.GetDeviceData pBuffer, DIGDD_DEFAULT

        If tvwScenery.Visible = True Then Exit Sub

        If Screen.ActiveForm.hWnd <> Me.hWnd Or Me.ActiveControl = txtZoom Then Exit Sub

        If DIState.Key(DIK_SPACE) = 128 And Not spaceDown Then
            circleOn = False
            spaceDown = True
            scrollCoords(1).X = mouseCoords.X
            scrollCoords(1).Y = mouseCoords.Y
            SetCursor TOOL_HAND + 1
            Exit Sub
        ElseIf (DIState.Key(DIK_LSHIFT) = 128 Or DIState.Key(DIK_RSHIFT) = 128) And Not shiftDown Then
            circleOn = False
            shiftDown = True
            Select Case currentTool
            Case Is = TOOL_VSELECT  ' add verts
                currentFunction = TOOL_VSELADD
            Case Is = TOOL_PSELECT  ' add polys
                currentFunction = TOOL_PSELADD
            Case Is = TOOL_WAYPOINT
                currentFunction = TOOL_CONNECT
            Case Is = TOOL_COLORPICKER
                currentFunction = TOOL_PIXPICKER
            Case Is = TOOL_SKETCH
                sketch(0).vertex(1).X = mouseCoords.X / zoomFactor + scrollCoords(2).X
                sketch(0).vertex(1).Y = mouseCoords.Y / zoomFactor + scrollCoords(2).Y
            End Select
            SetCursor currentFunction + 1
            lblCurrentTool.Caption = frmSoldatMapEditor.ImageList.ListImages(currentFunction + 1).Tag
            Exit Sub
        ElseIf (DIState.Key(DIK_LCONTROL) = 128 Or DIState.Key(DIK_RCONTROL) = 128) And Not ctrlDown Then
            circleOn = False
            ctrlDown = True
            Select Case currentTool
            Case Is = TOOL_MOVE
                currentFunction = TOOL_SCALE
                If altDown Then
                    ApplyTransform True
                End If
                toolAction = False
            Case Is = TOOL_SKETCH
                currentFunction = TOOL_SMUDGE
                circleOn = True
            Case Is > TOOL_MOVE
                currentFunction = TOOL_MOVE
                If currentTool <> TOOL_CREATE Then
                    toolAction = False
                End If
            End Select
            Render
            SetCursor currentFunction + 1
            lblCurrentTool.Caption = frmSoldatMapEditor.ImageList.ListImages(currentFunction + 1).Tag
            Exit Sub
        ElseIf (DIState.Key(DIK_LALT) = 128 Or DIState.Key(DIK_RALT) = 128) And Not altDown Then
            circleOn = False
            altDown = True
            Select Case currentTool
            Case Is = TOOL_MOVE
                currentFunction = TOOL_ROTATE
                If toolAction Then
                    If ctrlDown Then
                        ApplyTransform False
                    End If
                    toolAction = False
                End If
            Case Is = TOOL_VSELECT  ' subtract verts
                currentFunction = TOOL_VSELSUB
            Case Is = TOOL_PSELECT  ' subtract polys
                currentFunction = TOOL_PSELSUB
            Case Is = TOOL_VCOLOR  ' color picker
                currentFunction = TOOL_COLORPICKER
            Case Is = TOOL_PCOLOR  ' color picker
                currentFunction = TOOL_COLORPICKER
            Case Is = TOOL_DEPTHMAP
                currentFunction = TOOL_COLORPICKER
            Case Is = TOOL_COLORPICKER
                currentFunction = TOOL_LITPICKER
            Case Is = TOOL_SKETCH
                currentFunction = TOOL_ERASER
                circleOn = True
            Case Else
                currentFunction = TOOL_VSELECT
            End Select
            If currentFunction = TOOL_TEXTURE Then toolAction = False
            If currentFunction = TOOL_VCOLOR Or currentFunction = TOOL_DEPTHMAP Then circleOn = True
            Render
            SetCursor currentFunction + 1
            lblCurrentTool.Caption = frmSoldatMapEditor.ImageList.ListImages(currentFunction + 1).Tag
            Exit Sub
        End If

        hotKeyPressed = -1
        For i = 0 To 13
            If (DIState.Key(frmTools.getHotKey(i))) Then hotKeyPressed = i
        Next
        wayptKeyPressed = -1
        For i = 0 To 4
            If (DIState.Key(frmWaypoints.getWayptKey(i))) Then wayptKeyPressed = i
        Next
        layerKeyPressed = -1
        For i = 0 To 7
            If (DIState.Key(frmDisplay.getLayerKey(i))) Then layerKeyPressed = i
        Next

        ' key up
        If (pBuffer(0).lData = 0) Then
            If ((pBuffer(0).lOfs = DIK_RSHIFT Or pBuffer(0).lOfs = DIK_LSHIFT) And shiftDown) Then
                shiftDown = False
                currentFunction = currentTool
                If currentFunction = TOOL_SKETCH Then
                    toolAction = False
                    Render
                ElseIf currentFunction = TOOL_MOVE Then
                    If altDown Then
                        currentFunction = TOOL_ROTATE
                    ElseIf ctrlDown Then
                        currentFunction = TOOL_SCALE
                    End If
                End If
            ElseIf ((pBuffer(0).lOfs = DIK_RCONTROL Or pBuffer(0).lOfs = DIK_LCONTROL) And ctrlDown) Then
                ctrlDown = False
                If currentTool = TOOL_VSELECT Then
                    toolAction = False
                ElseIf currentTool = TOOL_MOVE Then
                    ApplyTransform False
                ElseIf currentTool = TOOL_SCENERY Then
                    Scenery(0).screenTr.X = mouseCoords.X
                    Scenery(0).screenTr.Y = mouseCoords.Y
                    Scenery(0).Translation.X = mouseCoords.X
                    Scenery(0).Translation.Y = mouseCoords.Y
                ElseIf currentTool = TOOL_OBJECTS Then
                    Spawns(0).X = mouseCoords.X
                    Spawns(0).Y = mouseCoords.Y
                ElseIf currentTool = TOOL_DEPTHMAP Then
                    circleOn = True
                ElseIf currentTool = TOOL_VCOLOR Then
                    circleOn = True
                ElseIf currentTool = TOOL_SKETCH Then
                    circleOn = False
                End If
                Render
                currentFunction = currentTool
            ElseIf ((pBuffer(0).lOfs = DIK_RALT Or pBuffer(0).lOfs = DIK_LALT) And altDown) Then
                altDown = False
                If currentTool = TOOL_MOVE Then
                    ApplyTransform True
                ElseIf currentTool = TOOL_SCENERY Then
                    Scenery(0).screenTr.X = mouseCoords.X
                    Scenery(0).screenTr.Y = mouseCoords.Y
                    Scenery(0).Translation.X = mouseCoords.X
                    Scenery(0).Translation.Y = mouseCoords.Y
                ElseIf currentTool = TOOL_OBJECTS Then
                    Spawns(0).X = mouseCoords.X
                    Spawns(0).Y = mouseCoords.Y
                ElseIf currentTool = TOOL_DEPTHMAP Then
                    circleOn = True
                ElseIf currentTool = TOOL_VCOLOR Then
                    circleOn = True
                ElseIf currentTool = TOOL_SKETCH Then
                    circleOn = False
                End If
                Render
                currentFunction = currentTool
            ElseIf (pBuffer(0).lOfs = DIK_SPACE And spaceDown) Then  ' scrolling
                spaceDown = False
            End If

            SetCursor currentFunction + 1
            lblCurrentTool.Caption = frmSoldatMapEditor.ImageList.ListImages(currentFunction + 1).Tag
        End If

        If ctrlDown Then  ' shortcuts
            If DIState.Key(DIK_EQUALS) = 128 Then  ' ctrl++
                Zoom getZoomDir(2)
            ElseIf DIState.Key(DIK_MINUS) = 128 Then  ' ctrl+-
                Zoom getZoomDir(0.5)
            ElseIf DIState.Key(DIK_0) = 128 Then  ' ctrl+0
                zoomFactor = 1
                scrollCoords(2).X = -ScaleWidth / 2
                scrollCoords(2).Y = -ScaleHeight / 2
                Zoom 1
            ElseIf DIState.Key(MapVirtualKey(78, 0)) = 128 Then  ' ctrl+n
                mnuNew_Click
            ElseIf DIState.Key(MapVirtualKey(79, 0)) = 128 And shiftDown Then  ' ctrl+shift+o
                mnuOpenCompiled_Click
            ElseIf DIState.Key(MapVirtualKey(79, 0)) = 128 Then  ' ctrl+o
                mnuOpen_Click
            ElseIf DIState.Key(MapVirtualKey(83, 0)) = 128 And shiftDown Then  ' ctrl+shift+s
                mnuSaveAs_Click
            ElseIf DIState.Key(MapVirtualKey(83, 0)) = 128 Then  ' ctrl+s
                mnuSave_Click
            ElseIf DIState.Key(MapVirtualKey(69, 0)) = 128 Then  ' ctrl+e
                mnuCreate_Click
            ElseIf DIState.Key(MapVirtualKey(86, 0)) = 128 Then  ' ctrl+v
                mnuPaste_Click
            ElseIf DIState.Key(MapVirtualKey(67, 0)) = 128 Then  ' ctrl+c
                mnuCopy_Click
            ElseIf DIState.Key(MapVirtualKey(90, 0)) = 128 Then  ' ctrl+z
                loadUndo False
            ElseIf DIState.Key(MapVirtualKey(89, 0)) = 128 Then  ' ctrl+y
                loadUndo True
            ElseIf DIState.Key(MapVirtualKey(65, 0)) = 128 Then  ' ctrl+a
                mnuSelectAll_Click
            ElseIf DIState.Key(MapVirtualKey(68, 0)) = 128 Then  ' ctrl+d
                mnuDuplicate_Click
            ElseIf DIState.Key(MapVirtualKey(73, 0)) = 128 Then  ' ctrl+i
                mnuInvertSel_Click
            ElseIf DIState.Key(MapVirtualKey(66, 0)) = 128 Then  ' ctrl+b
                mnuSelColor_Click
            ElseIf DIState.Key(MapVirtualKey(74, 0)) = 128 Then  ' ctrl+j
                mnuJoinVertices_Click
            ElseIf DIState.Key(MapVirtualKey(85, 0)) = 128 Then  ' ctrl+u
                mnuUntexture_Click
            ElseIf DIState.Key(MapVirtualKey(70, 0)) = 128 Then  ' ctrl+f
                mnuFixTexture_Click
            ElseIf DIState.Key(MapVirtualKey(76, 0)) = 128 Then  ' ctrl+l
                mnuSplit_Click
            ElseIf DIState.Key(MapVirtualKey(77, 0)) = 128 Then  ' ctrl+m
                mnuMap_Click
            ElseIf DIState.Key(MapVirtualKey(80, 0)) = 128 Then  ' ctrl+p
                mnuPreferences_Click
            ElseIf DIState.Key(MapVirtualKey(71, 0)) = 128 Then  ' ctrl+g
                AverageVertices
            ElseIf DIState.Key(DIK_APOSTROPHE) = 128 Then  ' ctrl+'
                mnuGrid_Click
            ElseIf DIState.Key(MapVirtualKey(84, 0)) = 128 Then  ' ctrl+t
                AutoTexture
            End If
        Else
            If hotKeyPressed > -1 And Not (shiftDown Or ctrlDown Or altDown) Then  ' hotkey
                setCurrentTool hotKeyPressed
                frmTools.picTools_MouseDown hotKeyPressed, 1, 0, 1, 1
            ElseIf wayptKeyPressed > -1 And Not (shiftDown Or ctrlDown Or altDown) Then  ' waypoint key
                frmWaypoints.picType_MouseUp wayptKeyPressed, 1, 0, 0, 0
            ElseIf layerKeyPressed > -1 And Not (shiftDown Or ctrlDown Or altDown) Then  ' layer key
                frmDisplay.picLayer_MouseUp layerKeyPressed, 1, 0, 0, 0
            ElseIf DIState.Key(DIK_NUMPADPLUS) = 128 Then  ' +
                Zoom getZoomDir(2)
            ElseIf DIState.Key(DIK_NUMPADMINUS) = 128 Then  ' -
                Zoom getZoomDir(0.5)
            ElseIf DIState.Key(DIK_NUMPADSTAR) = 128 Then  ' *
                Zoom 1 / zoomFactor
            ElseIf DIState.Key(DIK_DELETE) = 128 Then  ' delete
                deletePolys
            ElseIf DIState.Key(DIK_TAB) = 128 Then  ' tab
                TabPressed
            ElseIf (DIState.Key(DIK_ESCAPE) = 128) Then  ' esc
                If numVerts > 0 Or numCorners > 0 Or currentWaypoint > 0 Then
                    numVerts = 0
                    numCorners = 0
                    currentWaypoint = 0
                    toolAction = False
                    Render
                Else
                    mnuDeselect_Click
                End If
            ElseIf (DIState.Key(DIK_BACKSPACE) = 128) Then  ' backspace
                mnuSever_Click
            ElseIf (DIState.Key(DIK_INSERT) = 128 And shiftDown) Then  ' shift+insert
                mnuDuplicate_Click
            ElseIf (DIState.Key(DIK_HOME) = 128) Then  ' Home
                mnuBringToFront_Click
            ElseIf (DIState.Key(DIK_END) = 128) Then  ' End
                mnuSendToBack_Click
            ElseIf (DIState.Key(DIK_PGUP) = 128) Then  ' Page Up
                mnuBringForward_Click
            ElseIf (DIState.Key(DIK_PGDN) = 128) Then  ' Page Down
                mnuSendBackward_Click
            ElseIf (DIState.Key(DIK_F1) = 128) Then  ' F1
                RunHelp
            ElseIf (DIState.Key(DIK_F5) = 128) Then  ' F5
                mnuRefreshBG_Click
            ElseIf (DIState.Key(DIK_F8) = 128) Then  ' F8
                mnuRunSoldat_Click
            ElseIf (DIState.Key(DIK_F9) = 128) Then  ' F9
                mnuCompileAs_Click
            ElseIf (DIState.Key(DIK_F4) = 128 And altDown) Then  ' alt+F4
                mnuExit_Click
            ElseIf (DIState.Key(DIK_LBRACKET) = 128) Then  ' [
                If currentTool = 0 Then
                    setCurrentTool TOOL_DEPTHMAP
                Else
                    setCurrentTool currentTool - 1
                End If
                frmTools.picTools_MouseDown CInt(currentTool), 1, 0, 1, 1
            ElseIf (DIState.Key(DIK_RBRACKET) = 128) Then  ' ]
                If currentTool = TOOL_DEPTHMAP Then
                    setCurrentTool TOOL_MOVE
                Else
                    setCurrentTool currentTool + 1
                End If
                frmTools.picTools_MouseDown CInt(currentTool), 1, 0, 1, 1
            ElseIf (DIState.Key(DIK_LEFT) = 128 Or DIState.Key(DIK_UP) = 128 _
                    Or DIState.Key(DIK_RIGHT) = 128 _
                    Or DIState.Key(DIK_DOWN) = 128) Then  ' arrow keys
                Dim n As Single
                moveCoords(1).X = 0
                moveCoords(1).Y = 0
                If shiftDown Then
                    n = gridSpacing / gridDivisions * zoomFactor
                Else
                    n = zoomFactor
                End If
                If currentTool = TOOL_TEXTURE And numSelectedPolys > 0 Then
                    If selectionChanged Then
                        SaveUndo
                        selectionChanged = False
                    End If
                    If DIState.Key(DIK_LEFT) = 128 Then  ' left
                        StretchingTexture -n, 0
                    ElseIf DIState.Key(DIK_UP) = 128 Then  ' up
                        StretchingTexture 0, -n
                    ElseIf DIState.Key(DIK_RIGHT) = 128 Then  ' right
                        StretchingTexture n, 0
                    ElseIf DIState.Key(DIK_DOWN) = 128 Then  ' down
                        StretchingTexture 0, n
                    End If
                    SaveUndo
                Else
                    If selectionChanged Then
                        SaveUndo
                        selectionChanged = False
                    End If
                    If DIState.Key(DIK_LEFT) = 128 Then  ' left
                        Moving -n, 0
                    ElseIf DIState.Key(DIK_UP) = 128 Then  ' up
                        Moving 0, -n
                    ElseIf DIState.Key(DIK_RIGHT) = 128 Then  ' right
                        Moving n, 0
                    ElseIf DIState.Key(DIK_DOWN) = 128 Then  ' down
                        Moving 0, n
                    End If
                    SaveUndo
                End If
            End If
        End If

        lblCurrentTool.Caption = frmSoldatMapEditor.ImageList.ListImages(currentFunction + 1).Tag
    End If

    Exit Sub

ErrorHandler:

    If err.Number = DIERR_INPUTLOST Then
        acquired = False
    ElseIf err.Number = DIERR_NOTACQUIRED Then
        ' no-op
    Else
        MsgBox "DirectInput error" & vbNewLine & D3DX.GetErrorString(err.Number)
    End If

End Sub

Private Sub TabPressed()

    Dim scenNum As Integer
    Dim tempSel As Byte
    Dim i As Integer

    If numSelectedPolys = 1 And numSelectedScenery = 0 Then
        If vertexList(selectedPolys(1)).vertex(1) + vertexList(selectedPolys(1)).vertex(2) + vertexList(selectedPolys(1)).vertex(3) = 3 Then
            vertexList(selectedPolys(1)).vertex(1) = 0
            vertexList(selectedPolys(1)).vertex(2) = 0
            vertexList(selectedPolys(1)).vertex(3) = 0
            If Not shiftDown Then
                If selectedPolys(1) = mPolyCount Then
                    selectedPolys(1) = 1
                Else
                    selectedPolys(1) = selectedPolys(1) + 1
                End If
            Else
                Beep
                If selectedPolys(1) = 1 Then
                    selectedPolys(1) = mPolyCount
                Else
                    selectedPolys(1) = selectedPolys(1) - 1
                End If
            End If
            vertexList(selectedPolys(1)).vertex(1) = 1
            vertexList(selectedPolys(1)).vertex(2) = 1
            vertexList(selectedPolys(1)).vertex(3) = 1
        Else
            tempSel = vertexList(selectedPolys(1)).vertex(1)
            vertexList(selectedPolys(1)).vertex(1) = vertexList(selectedPolys(1)).vertex(2)
            vertexList(selectedPolys(1)).vertex(2) = vertexList(selectedPolys(1)).vertex(3)
            vertexList(selectedPolys(1)).vertex(3) = tempSel
        End If

        Render

    ElseIf numSelectedScenery = 1 And numSelectedPolys = 0 Then
        For i = 1 To sceneryCount
            If Scenery(i).selected = 1 Then
                    scenNum = i
                End If
            Next
            Scenery(scenNum).selected = 0
        If Not shiftDown Then
            If scenNum = sceneryCount Then
                Scenery(1).selected = 1
            Else
                Scenery(scenNum + 1).selected = 1
            End If
        Else
            If scenNum = 1 Then
                Scenery(sceneryCount).selected = 1
            Else
                Scenery(scenNum - 1).selected = 1
            End If
        End If

        Render
    End If

    getInfo

End Sub

Private Sub findDragPoint(X As Single, Y As Single)

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim midCoords As D3DVECTOR2

    ' check if user moused down on corner drag point of sel rect
    For i = 0 To 3
        j = i + 2
        If j > 3 Then
            j = i - 2
        End If
        If nearCoord((selRect(i).X - scrollCoords(2).X) * zoomFactor, moveCoords(1).X, 8) And _
                nearCoord((selRect(i).Y - scrollCoords(2).Y) * zoomFactor, moveCoords(1).Y, 8) Then
            If mnuFixedRCenter.Checked Then
                rCenter.X = selRect(j).X
                rCenter.Y = selRect(j).Y
            End If
            moveCoords(1).X = (selRect(i).X - scrollCoords(2).X) * zoomFactor
            moveCoords(1).Y = (selRect(i).Y - scrollCoords(2).Y) * zoomFactor
            X = moveCoords(1).X
            Y = moveCoords(1).Y
            toolAction = True
        End If
    Next

    If toolAction = False Then
        For i = 0 To 3
            j = i + 2
            If j > 3 Then
                j = i - 2
            End If
            k = i + 1
            If k > 3 Then
                k = 0
            End If
            midCoords.X = Midpoint(selRect(i).X, selRect(k).X)
            midCoords.Y = Midpoint(selRect(i).Y, selRect(k).Y)
            k = i - 1
            If k < 0 Then
                k = 3
            End If
            If nearCoord((midCoords.X - scrollCoords(2).X) * zoomFactor, moveCoords(1).X, 8) And _
                    nearCoord((midCoords.Y - scrollCoords(2).Y) * zoomFactor, moveCoords(1).Y, 8) Then
                If mnuFixedRCenter.Checked Then
                    rCenter.X = Midpoint(selRect(j).X, selRect(k).X)
                    rCenter.Y = Midpoint(selRect(j).Y, selRect(k).Y)
                End If
                moveCoords(1).X = (midCoords.X - scrollCoords(2).X) * zoomFactor
                moveCoords(1).Y = (midCoords.Y - scrollCoords(2).Y) * zoomFactor
                X = moveCoords(1).X
                Y = moveCoords(1).Y
                toolAction = True
            End If
        Next
        Render
    End If

End Sub

Private Sub findDragPoint2(X As Single, Y As Single)

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim midCoords As D3DVECTOR2

    toolAction = checkDragPoint(selRect(0).X, selRect(0).Y, selRect(2).X, selRect(2).Y)
    If Not toolAction Then toolAction = checkDragPoint(selRect(1).X, selRect(1).Y, selRect(3).X, selRect(3).Y)
    If Not toolAction Then toolAction = checkDragPoint(selRect(2).X, selRect(2).Y, selRect(0).X, selRect(0).Y)
    If Not toolAction Then toolAction = checkDragPoint(selRect(3).X, selRect(3).Y, selRect(1).X, selRect(1).Y)

    midCoords.X = Midpoint(selRect(0).X, selRect(1).X)
    midCoords.Y = Midpoint(selRect(0).Y, selRect(1).Y)
    If Not toolAction Then toolAction = checkDragPoint(midCoords.X, midCoords.Y, Midpoint(selRect(2).X, selRect(3).X), Midpoint(selRect(2).Y, selRect(3).Y))
    midCoords.X = Midpoint(selRect(1).X, selRect(2).X)
    midCoords.Y = Midpoint(selRect(1).Y, selRect(2).Y)
    If Not toolAction Then toolAction = checkDragPoint(midCoords.X, midCoords.Y, Midpoint(selRect(3).X, selRect(0).X), Midpoint(selRect(3).Y, selRect(0).Y))
    midCoords.X = Midpoint(selRect(2).X, selRect(3).X)
    midCoords.Y = Midpoint(selRect(2).Y, selRect(3).Y)
    If Not toolAction Then toolAction = checkDragPoint(midCoords.X, midCoords.Y, Midpoint(selRect(0).X, selRect(1).X), Midpoint(selRect(0).Y, selRect(1).Y))
    midCoords.X = Midpoint(selRect(3).X, selRect(0).X)
    midCoords.Y = Midpoint(selRect(3).Y, selRect(0).Y)
    If Not toolAction Then toolAction = checkDragPoint(midCoords.X, midCoords.Y, Midpoint(selRect(1).X, selRect(2).X), Midpoint(selRect(1).Y, selRect(2).Y))

End Sub

Private Function checkDragPoint(x1 As Single, y1 As Single, x2 As Single, y2 As Single) As Boolean

    If nearCoord((x1 - scrollCoords(2).X) * zoomFactor, moveCoords(1).X, 8) And nearCoord((y1 - scrollCoords(2).Y) * zoomFactor, moveCoords(1).Y, 8) Then
        If mnuFixedRCenter.Checked Then
            rCenter.X = x2
            rCenter.Y = y2
        End If
        moveCoords(1).X = (x1 - scrollCoords(2).X) * zoomFactor
        moveCoords(1).Y = (y1 - scrollCoords(2).Y) * zoomFactor
        checkDragPoint = True
    End If

End Function

Private Sub Form_DblClick()

    If currentTool = TOOL_CREATE Then  ' poly creation
        toolAction = True
    ElseIf currentTool = TOOL_VSELECT Then  ' vertex selection
        toolAction = True
        selectedCoords(1).X = MouseHelper.CursorX - (Me.Left / Screen.TwipsPerPixelX) - 1
        selectedCoords(1).Y = MouseHelper.CursorY - (Me.Top / Screen.TwipsPerPixelY) - 1
        selectedCoords(2).X = selectedCoords(1).X
        selectedCoords(2).Y = selectedCoords(1).Y

        Render
    End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    On Error GoTo ErrorHandler

    If acquired = False Then
        DIDevice.Acquire
        acquired = True
    End If

    If Button = 2 Then  ' popup menus
        If currentFunction = TOOL_CREATE Or currentFunction = TOOL_QUAD Then
            Me.PopupMenu mnuPolyTypes
        ElseIf currentTool = TOOL_MOVE Then
            mouseCoords.X = X
            mouseCoords.Y = Y
            Me.PopupMenu mnuMove
        ElseIf currentTool = TOOL_PSELECT Or currentTool = TOOL_VSELECT Then
            Me.PopupMenu mnuVertexSelect
        ElseIf currentFunction = TOOL_SCENERY Then
            If tvwScenery.Visible = False Then
                If Me.Tag = vbMaximized Then
                    tvwScenery.Left = mouseCoords.X
                    If mouseCoords.Y + tvwScenery.Height > Me.ScaleHeight - 17 Then
                        tvwScenery.Top = Me.ScaleHeight - tvwScenery.Height - 17
                    Else
                        tvwScenery.Top = mouseCoords.Y
                    End If
                Else
                    tvwScenery.Left = 0
                    tvwScenery.Top = 41
                End If
            End If
            tvwScenery.Visible = Not tvwScenery.Visible
        ElseIf currentFunction = TOOL_OBJECTS Then
            Me.PopupMenu mnuObjects, , X, Y
            Render
        ElseIf currentFunction = TOOL_WAYPOINT Then
            Me.PopupMenu mnuWaypoint, , X, Y
        End If
    ElseIf Button = 4 Then
        scrollCoords(1).X = X
        scrollCoords(1).Y = Y
        SetCursor TOOL_HAND + 1
    Else
        If tvwScenery.Visible Then tvwScenery.Visible = False
    End If

    If Button <> 1 Then Exit Sub

    If spaceDown Then
        scrollCoords(1).X = X
        scrollCoords(1).Y = Y
    ElseIf currentFunction = TOOL_MOVE Then  ' move
        If selectionChanged Then
            SaveUndo
            selectionChanged = False
        End If

        toolAction = True
        MouseDownMove X, Y
    ElseIf currentFunction = TOOL_ROTATE Or currentFunction = TOOL_SCALE Then  ' scaling/rotation
        If selectionChanged Then
            SaveUndo
            selectionChanged = False
        End If

        moveCoords(1).X = X
        moveCoords(1).Y = Y
        moveCoords(2).X = X
        moveCoords(2).Y = Y

        findDragPoint2 X, Y
    ElseIf (currentFunction = TOOL_CREATE Or currentFunction = TOOL_QUAD) Then  ' poly creation
        If selectionChanged Then
            SaveUndo
            selectionChanged = False
        End If

        If Shift = 0 Then
            If Not (showPolys Or showWireframe Or showPoints) Then
                showPolys = True
                frmDisplay.setLayer 1, True
            End If
            toolAction = True
        ElseIf Shift = KEY_SHIFT Then  ' constrained
            If Not (showPolys Or showWireframe Or showPoints) Then
                showPolys = True
            End If
            toolAction = True
        End If
    ElseIf currentFunction = TOOL_VSELECT Or currentFunction = TOOL_VSELADD Or currentFunction = TOOL_VSELSUB Then  ' vertex selection
        toolAction = True
        selectedCoords(1).X = X
        selectedCoords(1).Y = Y
        selectedCoords(2).X = X
        selectedCoords(2).Y = Y
    ElseIf currentFunction = TOOL_PSELECT Or currentFunction = TOOL_PSELADD Or currentFunction = TOOL_PSELSUB Then ' poly selection
        polySelection X, Y
    ElseIf currentFunction = TOOL_VCOLOR Then  ' vertex color
        If selectionChanged Then
            SaveUndo
            selectionChanged = False
        End If

        toolAction = True
        If colorMode > 0 Then
            VertexColoring X, Y
        Else
            PrecisionColoring X, Y
        End If
    ElseIf currentFunction = TOOL_PCOLOR Then  ' poly color
        If selectionChanged Then
            SaveUndo
            selectionChanged = False
        End If

        ColorFill X, Y
    ElseIf currentFunction = TOOL_TEXTURE Then  ' texture
        If selectionChanged Then
            SaveUndo
            selectionChanged = False
        End If

        If Shift = 0 Then
            toolAction = True
            MouseDownMove X, Y
        ElseIf Shift = KEY_SHIFT Then  ' constrained
            toolAction = True
            moveCoords(2).X = X
            moveCoords(2).Y = Y
            moveCoords(1).X = X
            moveCoords(1).Y = Y
        End If
    ElseIf currentFunction = TOOL_SCENERY Then
        If selectionChanged Then
            SaveUndo
            selectionChanged = False
        End If

        If Not showScenery Then
            showScenery = True
            frmDisplay.setLayer 5, showScenery
        End If
        toolAction = True
    ElseIf currentFunction = TOOL_COLORPICKER Then  ' color picker
        If currentTool = TOOL_DEPTHMAP Then
            depthPicker X, Y
        ElseIf currentTool = TOOL_SCENERY Then
        Else
            ColorPicker X, Y
        End If
    ElseIf currentFunction = TOOL_PIXPICKER Then
        Dim tempClr As TColor
        tempClr = getRGB(GetPixel(Me.hDC, X, Y))
        If frmPalette.Enabled = False Then
            frmColor.InitColor tempClr.blue, tempClr.green, tempClr.red
        Else
            gPolyClr.red = tempClr.blue
            gPolyClr.green = tempClr.green
            gPolyClr.blue = tempClr.red
            Scenery(0).color = ARGB(Scenery(0).alpha, RGB(gPolyClr.blue, gPolyClr.green, gPolyClr.red))
            frmPalette.setValues gPolyClr.red, gPolyClr.green, gPolyClr.blue
        End If
    ElseIf currentFunction = TOOL_LITPICKER Then
        lightPicker X, Y
    ElseIf currentFunction = TOOL_OBJECTS Then  ' objects
        If selectionChanged Then
            SaveUndo
            selectionChanged = False
        End If

        If Not showObjects And Not mnuGostek.Checked Then
            showObjects = True
            frmDisplay.setLayer 6, showObjects
        End If
        If mnuGostek.Checked Then
            gostek.X = X / zoomFactor + scrollCoords(2).X
            gostek.Y = Y / zoomFactor + scrollCoords(2).Y
        ElseIf Not mnuCollider.Checked Then
            spawnPoints = spawnPoints + 1
            ReDim Preserve Spawns(spawnPoints)

            Spawns(spawnPoints).Team = Spawns(0).Team
            Spawns(spawnPoints).X = X / zoomFactor + scrollCoords(2).X
            Spawns(spawnPoints).Y = Y / zoomFactor + scrollCoords(2).Y

            If showGrid And snapToGrid Then
                Spawns(spawnPoints).X = Int((Spawns(spawnPoints).X + inc / 2) / inc) * inc
                Spawns(spawnPoints).Y = Int((Spawns(spawnPoints).Y + inc / 2) / inc) * inc
            End If
        Else
            colliderCount = colliderCount + 1
            ReDim Preserve Colliders(colliderCount)

            Colliders(colliderCount).radius = Colliders(0).radius
            Colliders(colliderCount).X = X / zoomFactor + scrollCoords(2).X
            Colliders(colliderCount).Y = Y / zoomFactor + scrollCoords(2).Y

            If showGrid And snapToGrid Then
                Colliders(colliderCount).X = Int((Colliders(colliderCount).X + inc / 2) / inc) * inc
                Colliders(colliderCount).Y = Int((Colliders(colliderCount).Y + inc / 2) / inc) * inc
            End If
        End If
        Render
        toolAction = True
    ElseIf currentFunction = TOOL_WAYPOINT Then  ' waypoints
        If selectionChanged Then
            SaveUndo
            selectionChanged = False
        End If

        If Not showWaypoints Then
            showWaypoints = True
            frmDisplay.setLayer 7, showWaypoints
        End If

        If frmWaypoints.showPaths = 1 And frmWaypoints.wayptPath = 1 Or frmWaypoints.showPaths = 2 And frmWaypoints.wayptPath = 0 Then
            frmWaypoints.picShow_MouseUp 0, 1, 0, 0, 0
            mouseEvent2 frmWaypoints.picShow(0), 0, 0, BUTTON_SMALL, True, BUTTON_UP
        End If

        mnuDeselect_Click

        waypointCount = waypointCount + 1
        ReDim Preserve Waypoints(waypointCount)

        Waypoints(waypointCount).selected = True
        numSelWaypoints = numSelWaypoints + 1

        Waypoints(waypointCount).X = X / zoomFactor + scrollCoords(2).X
        Waypoints(waypointCount).Y = Y / zoomFactor + scrollCoords(2).Y

        Waypoints(waypointCount).pathNum = frmWaypoints.wayptPath + 1

        For i = 0 To 4
            Waypoints(waypointCount).wayType(i) = mnuWayType(i).Checked
        Next

        If currentWaypoint > 0 Then  ' connecting waypoints
            conCount = conCount + 1
            ReDim Preserve Connections(conCount)
            Connections(conCount).point1 = currentWaypoint
            Connections(conCount).point2 = waypointCount
            Waypoints(waypointCount).numConnections = Waypoints(waypointCount).numConnections + 1
            currentWaypoint = waypointCount
        End If
        getInfo
        Render
        toolAction = True
    ElseIf currentFunction = TOOL_CONNECT Then
        If selectionChanged Then
            SaveUndo
            selectionChanged = False
        End If

        toolAction = True
    ElseIf currentFunction = TOOL_DEPTHMAP Then
        If selectionChanged Then
            SaveUndo
            selectionChanged = False
        End If

        EditDepthMap X, Y

        toolAction = True
    ElseIf currentFunction = TOOL_LIGHTS Then
        CreateLight X, Y
    ElseIf currentFunction = TOOL_SKETCH Then
        If Shift = 0 Then  ' freeform
            startSketch X, Y
            toolAction = True
        ElseIf Shift = 1 Then
            showSketch = True
            frmDisplay.setLayer 10, showSketch
        End If
    ElseIf currentFunction = TOOL_ERASER Then
        If eraseSketch(X / zoomFactor + scrollCoords(2).X, Y / zoomFactor + scrollCoords(2).Y) = 1 Then
            Render
        End If
        toolAction = True
    ElseIf currentFunction = TOOL_SMUDGE Then
        moveCoords(2).X = X
        moveCoords(2).Y = Y
        moveCoords(1).X = X
        moveCoords(1).Y = Y
        toolAction = True
    End If

    Exit Sub

ErrorHandler:

    MsgBox Error$

End Sub

Private Sub CreateLight(X As Single, Y As Single)

    If selectionChanged Then
        SaveUndo
        selectionChanged = False
    End If

    showLights = True
    frmDisplay.setLayer 9, showLights

    lightCount = lightCount + 1
    ReDim Preserve Lights(lightCount)
    Lights(lightCount).X = X / zoomFactor + scrollCoords(2).X
    Lights(lightCount).Y = Y / zoomFactor + scrollCoords(2).Y
    Lights(lightCount).Z = 255
    Lights(lightCount).color = gPolyClr
    Lights(lightCount).intensity = opacity
    Lights(lightCount).range = 0

    If showGrid And snapToGrid Then
        Lights(lightCount).X = Int((Lights(lightCount).X + inc / 2) / inc) * inc
        Lights(lightCount).Y = Int((Lights(lightCount).Y + inc / 2) / inc) * inc
    End If

    applyLights
    SaveUndo
    Render

End Sub

Private Sub applyLights(Optional toSel As Boolean = False)

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    Dim lightDir As D3DVECTOR
    Dim polyNormal As D3DVECTOR
    Dim v1 As D3DVECTOR
    Dim v2 As D3DVECTOR
    Dim mag As Single
    Dim diffuseFactor As Single
    Dim totalDiffuse As Single

    Dim clr As TColor

    Dim rVal As Integer
    Dim gVal As Integer
    Dim bVal As Integer

    If lightCount = 0 Then Exit Sub

    For i = 1 To mPolyCount
        ' get poly vectors
        v1.X = PolyCoords(i).vertex(1).X - PolyCoords(i).vertex(2).X
        v1.Y = PolyCoords(i).vertex(1).Y - PolyCoords(i).vertex(2).Y
        v1.Z = Polys(i).vertex(1).Z - Polys(i).vertex(2).Z

        v2.X = PolyCoords(i).vertex(1).X - PolyCoords(i).vertex(3).X
        v2.Y = PolyCoords(i).vertex(1).Y - PolyCoords(i).vertex(3).Y
        v2.Z = Polys(i).vertex(1).Z - Polys(i).vertex(3).Z

        ' get poly normal
        polyNormal.X = (v1.Y * v2.Z) - (v1.Z * v2.Y)
        polyNormal.Y = (v1.Z * v2.X) - (v1.X * v2.Z)
        polyNormal.Z = (v1.X * v2.Y) - (v1.Y * v2.X)

        ' normalize poly normal
        mag = Sqr(polyNormal.X ^ 2 + polyNormal.Y ^ 2 + polyNormal.Z ^ 2)
        If mag > 0 Then
            polyNormal.X = polyNormal.X / mag
            polyNormal.Y = polyNormal.Y / mag
            polyNormal.Z = polyNormal.Z / mag
        End If

        For j = 1 To 3
            If (vertexList(i).vertex(j) = 1 And toSel) Or toSel = False Then

                For k = 1 To lightCount
                    ' get light dir vector
                    lightDir.X = Lights(k).X - PolyCoords(i).vertex(j).X
                    lightDir.Y = Lights(k).Y - PolyCoords(i).vertex(j).Y
                    lightDir.Z = Lights(k).Z - Polys(i).vertex(j).Z
                    ' normalize light dir
                    mag = Sqr(lightDir.X ^ 2 + lightDir.Y ^ 2 + lightDir.Z ^ 2)
                    If mag > 0 Then
                        lightDir.X = lightDir.X / mag
                        lightDir.Y = lightDir.Y / mag
                        lightDir.Z = lightDir.Z / mag
                    End If
                    ' get angle between light dir and poly normal (dot product)
                    diffuseFactor = (polyNormal.X * lightDir.X) + (polyNormal.Y * lightDir.Y) + (polyNormal.Z * lightDir.Z)
                    If diffuseFactor < 0 Then diffuseFactor = 0

                    If Lights(k).range = 0 Then  ' normal
                        mag = 1
                    Else  ' range > 0
                        If mag > 0 Then
                            If mag <= Lights(k).range Then
                                mag = 1 - mag / Lights(k).range
                            Else  ' vertex is out of range
                                mag = 0
                            End If
                        Else
                            mag = 0
                        End If
                    End If

                    ' calculate final color components
                    rVal = rVal + (Lights(k).color.red * diffuseFactor) * mag
                    gVal = gVal + (Lights(k).color.green * diffuseFactor) * mag
                    bVal = bVal + (Lights(k).color.blue * diffuseFactor) * mag

                    totalDiffuse = totalDiffuse + diffuseFactor
                Next

                totalDiffuse = totalDiffuse / lightCount

                clr = vertexList(i).color(j)
                rVal = rVal + clr.red
                gVal = gVal + clr.green
                bVal = bVal + clr.blue

                If rVal > 255 Then rVal = 255
                If gVal > 255 Then gVal = 255
                If bVal > 255 Then bVal = 255

                Polys(i).vertex(j).color = ARGB(getAlpha(Polys(i).vertex(j).color), RGB(Int(bVal), Int(gVal), Int(rVal)))

                rVal = 0
                gVal = 0
                bVal = 0
                totalDiffuse = 0

            End If
        Next
    Next

    Render

End Sub

Private Sub applyLightsToVert(pIndex As Integer, vIndex As Integer)

    On Error GoTo ErrorHandler

    If lightCount <= 0 Or Not showLights Then Exit Sub

    Dim k As Integer
    Dim lightDir As D3DVECTOR
    Dim polyNormal As D3DVECTOR
    Dim v1 As D3DVECTOR
    Dim v2 As D3DVECTOR
    Dim mag As Single
    Dim diffuseFactor As Single
    Dim totalDiffuse As Single
    Dim clr As TColor
    Dim rVal As Integer
    Dim gVal As Integer
    Dim bVal As Integer

    ' get poly vectors
    v1.X = PolyCoords(pIndex).vertex(1).X - PolyCoords(pIndex).vertex(2).X
    v1.Y = PolyCoords(pIndex).vertex(1).Y - PolyCoords(pIndex).vertex(2).Y
    v1.Z = Polys(pIndex).vertex(1).Z - Polys(pIndex).vertex(2).Z

    v2.X = PolyCoords(pIndex).vertex(1).X - PolyCoords(pIndex).vertex(3).X
    v2.Y = PolyCoords(pIndex).vertex(1).Y - PolyCoords(pIndex).vertex(3).Y
    v2.Z = Polys(pIndex).vertex(1).Z - Polys(pIndex).vertex(3).Z

    ' get poly normal
    polyNormal.X = (v1.Y * v2.Z) - (v1.Z * v2.Y)
    polyNormal.Y = (v1.Z * v2.X) - (v1.X * v2.Z)
    polyNormal.Z = (v1.X * v2.Y) - (v1.Y * v2.X)

    ' normalize poly normal
    mag = Sqr(polyNormal.X ^ 2 + polyNormal.Y ^ 2 + polyNormal.Z ^ 2)
    If mag > 0 Then
        polyNormal.X = polyNormal.X / mag
        polyNormal.Y = polyNormal.Y / mag
        polyNormal.Z = polyNormal.Z / mag
    End If

    For k = 1 To lightCount
        ' get light dir vector
        lightDir.X = Lights(k).X - PolyCoords(pIndex).vertex(vIndex).X
        lightDir.Y = Lights(k).Y - PolyCoords(pIndex).vertex(vIndex).Y
        lightDir.Z = Lights(k).Z - Polys(pIndex).vertex(vIndex).Z
        ' normalize light dir
        mag = Sqr(lightDir.X ^ 2 + lightDir.Y ^ 2 + lightDir.Z ^ 2)
        If mag > 0 Then
            lightDir.X = lightDir.X / mag
            lightDir.Y = lightDir.Y / mag
            lightDir.Z = lightDir.Z / mag
        End If
        ' get angle between light dir and poly normal (dot product)
        diffuseFactor = (polyNormal.X * lightDir.X) + (polyNormal.Y * lightDir.Y) + (polyNormal.Z * lightDir.Z)
        If diffuseFactor < 0 Then diffuseFactor = 0

        ' calculate final color components
        rVal = rVal + (Lights(k).color.red * diffuseFactor)
        gVal = gVal + (Lights(k).color.green * diffuseFactor)
        bVal = bVal + (Lights(k).color.blue * diffuseFactor)

        totalDiffuse = totalDiffuse + diffuseFactor
    Next

    totalDiffuse = totalDiffuse / lightCount

    clr = vertexList(pIndex).color(vIndex)
    rVal = rVal + clr.red
    gVal = gVal + clr.green
    bVal = bVal + clr.blue

    If rVal > 255 Then rVal = 255
    If gVal > 255 Then gVal = 255
    If bVal > 255 Then bVal = 255

    Polys(pIndex).vertex(vIndex).color = ARGB(getAlpha(Polys(pIndex).vertex(vIndex).color), RGB(Int(bVal), Int(gVal), Int(rVal)))

    rVal = 0
    gVal = 0
    bVal = 0
    totalDiffuse = 0

    Exit Sub

ErrorHandler:

    MsgBox Error$

End Sub

Private Sub SnapSelection()

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    Dim PolyNum As Integer

    For i = 1 To numSelectedPolys
        PolyNum = selectedPolys(i)
        For j = 1 To 3
            If vertexList(PolyNum).vertex(j) = 1 Then
                Polys(PolyNum).vertex(j).X = GetVertSnapCoord(PolyNum, j, 1)
                Polys(PolyNum).vertex(j).Y = GetVertSnapCoord(PolyNum, j, 0)
                If snapToGrid And showGrid Then
                    Polys(PolyNum).vertex(j).X = snapVertexToGrid(Polys(PolyNum).vertex(j).X, (scrollCoords(2).X - Int(scrollCoords(2).X / inc) * inc) * zoomFactor)
                    Polys(PolyNum).vertex(j).Y = snapVertexToGrid(Polys(PolyNum).vertex(j).Y, (scrollCoords(2).Y - Int(scrollCoords(2).Y / inc) * inc) * zoomFactor)
                End If
                PolyCoords(PolyNum).vertex(j).X = Polys(PolyNum).vertex(j).X / zoomFactor + scrollCoords(2).X
                PolyCoords(PolyNum).vertex(j).Y = Polys(PolyNum).vertex(j).Y / zoomFactor + scrollCoords(2).Y
            End If
        Next
    Next

End Sub

Private Function GetVertSnapCoord(PolyNum As Integer, VertNum As Integer, GetXVal As Boolean) As Integer

    Dim i As Integer
    Dim j As Integer
    Dim xVal As Integer
    Dim yVal As Integer
    Dim nearPoly As Integer
    Dim nearVert As Integer
    Dim minDiff As Long
    Dim thisDiff As Long
    Dim prevDiff As Long

    xVal = Polys(PolyNum).vertex(VertNum).X
    yVal = Polys(PolyNum).vertex(VertNum).Y
    If GetXVal Then
        GetVertSnapCoord = xVal
    Else
        GetVertSnapCoord = yVal
    End If

    If ohSnap Then
        nearPoly = -1
        minDiff = snapRadius ^ 2 + 1
        For i = 1 To mPolyCount
            For j = 1 To 3
                If nearPoly = -1 Then
                    prevDiff = (Polys(i).vertex(j).X - xVal) ^ 2 + (Polys(i).vertex(j).Y - yVal) ^ 2
                    If prevDiff < minDiff Then
                        nearPoly = i
                        nearVert = j
                    End If
                End If
            Next
        Next

        If Not nearPoly = -1 Then
            If GetXVal Then
                GetVertSnapCoord = Polys(nearPoly).vertex(nearVert).X
            Else
                GetVertSnapCoord = Polys(nearPoly).vertex(nearVert).Y
            End If
        End If
    End If

End Function

Private Sub AverageVerts()

    Dim i As Integer
    Dim j As Integer
    Dim finalR As Integer
    Dim finalG As Integer
    Dim finalB As Integer
    Dim tehClr As TColor

    For i = 1 To numSelectedPolys
        For j = 1 To 3
            If vertexList(selectedPolys(i)).vertex(j) = 1 Then
                tehClr = getRGB(Polys(selectedPolys(i)).vertex(j).color)
                finalR = finalR + tehClr.red
                finalG = finalG + tehClr.green
                finalB = finalB + tehClr.blue
            End If
        Next
    Next

    finalR = finalR / numSelectedPolys
    finalG = finalG / numSelectedPolys
    finalB = finalB / numSelectedPolys

    For i = 1 To numSelectedPolys
        For j = 1 To 3
            If vertexList(selectedPolys(i)).vertex(j) = 1 Then
                Polys(selectedPolys(i)).vertex(j).color = ARGB(getAlpha(Polys(selectedPolys(i)).vertex(j).color), RGB(finalR, finalG, finalB))
            End If
        Next
    Next

End Sub

Private Sub AverageVertices()

    Dim i As Integer
    Dim j As Integer
    Dim P As Integer
    Dim V As Integer
    Dim finalR As Integer
    Dim finalG As Integer
    Dim finalB As Integer
    Dim tehClr As TColor
    Dim vertexClr As TColor
    Dim numVertices As Integer
    Dim xVal As Single
    Dim yVal As Single
    Dim connectedPolys() As Integer
    Dim numConnectedPolys As Integer

    On Error GoTo ErrorHandler

    If selectionChanged Then
        SaveUndo
        selectionChanged = False
    End If

    Me.MousePointer = vbHourglass

    If numSelectedPolys = 0 Then
        For i = 1 To mPolyCount
            For j = 1 To 3
                If vertexList(i).vertex(j) = 0 Then
                    xVal = PolyCoords(i).vertex(j).X
                    yVal = PolyCoords(i).vertex(j).Y
                    finalR = 0
                    finalG = 0
                    finalB = 0
                    For P = 1 To mPolyCount
                        For V = 1 To 3
                            If nearCoord(xVal, PolyCoords(P).vertex(V).X, 2) And nearCoord(yVal, PolyCoords(P).vertex(V).Y, 2) Then
                                vertexList(P).vertex(V) = 1
                                tehClr.red = vertexList(P).color(V).red
                                tehClr.green = vertexList(P).color(V).green
                                tehClr.blue = vertexList(P).color(V).blue
                                finalR = finalR + tehClr.red
                                finalG = finalG + tehClr.green
                                finalB = finalB + tehClr.blue
                                numConnectedPolys = numConnectedPolys + 1
                                ReDim Preserve connectedPolys(numConnectedPolys)
                                connectedPolys(numConnectedPolys) = P
                            End If
                        Next
                    Next
                    finalR = finalR / numConnectedPolys
                    finalG = finalG / numConnectedPolys
                    finalB = finalB / numConnectedPolys

                    For P = 1 To numConnectedPolys
                        For V = 1 To 3
                            If vertexList(connectedPolys(P)).vertex(V) = 1 Then
                                vertexList(connectedPolys(P)).vertex(V) = 2
                                vertexList(connectedPolys(P)).color(V).red = finalR
                                vertexList(connectedPolys(P)).color(V).green = finalG
                                vertexList(connectedPolys(P)).color(V).blue = finalB
                                Polys(connectedPolys(P)).vertex(V).color = ARGB(getAlpha(Polys(connectedPolys(P)).vertex(V).color), RGB(finalB, finalG, finalR))
                            End If
                        Next
                    Next
                    numConnectedPolys = 0
                    ReDim connectedPolys(0)
                End If
            Next
        Next

        For i = 1 To mPolyCount
            vertexList(i).vertex(1) = 0
            vertexList(i).vertex(2) = 0
            vertexList(i).vertex(3) = 0
        Next

        applyLights
    Else
        For i = 1 To mPolyCount
            For j = 1 To 3
                If vertexList(i).vertex(j) = 1 Then
                    xVal = PolyCoords(i).vertex(j).X
                    yVal = PolyCoords(i).vertex(j).Y
                    finalR = 0
                    finalG = 0
                    finalB = 0
                    For P = 1 To mPolyCount
                        For V = 1 To 3
                            If nearCoord(xVal, PolyCoords(P).vertex(V).X, 2) And nearCoord(yVal, PolyCoords(P).vertex(V).Y, 2) Then
                                If vertexList(P).vertex(V) = 1 Then
                                    vertexList(P).vertex(V) = 2
                                    tehClr.red = vertexList(P).color(V).red
                                    tehClr.green = vertexList(P).color(V).green
                                    tehClr.blue = vertexList(P).color(V).blue
                                    finalR = finalR + tehClr.red
                                    finalG = finalG + tehClr.green
                                    finalB = finalB + tehClr.blue
                                    numConnectedPolys = numConnectedPolys + 1
                                    ReDim Preserve connectedPolys(numConnectedPolys)
                                    connectedPolys(numConnectedPolys) = P
                                End If
                            End If
                        Next
                    Next
                    finalR = finalR / numConnectedPolys
                    finalG = finalG / numConnectedPolys
                    finalB = finalB / numConnectedPolys
                    For P = 1 To numConnectedPolys
                        For V = 1 To 3
                            If vertexList(connectedPolys(P)).vertex(V) = 2 Then
                                vertexList(connectedPolys(P)).vertex(V) = 3
                                vertexList(connectedPolys(P)).color(V).red = finalR
                                vertexList(connectedPolys(P)).color(V).green = finalG
                                vertexList(connectedPolys(P)).color(V).blue = finalB
                                Polys(connectedPolys(P)).vertex(V).color = ARGB(getAlpha(Polys(connectedPolys(P)).vertex(V).color), RGB(finalB, finalG, finalR))
                            End If
                        Next
                    Next
                    numConnectedPolys = 0
                    ReDim connectedPolys(0)
                End If
            Next
        Next

        For i = 1 To mPolyCount
            For j = 1 To 3
                If vertexList(i).vertex(j) > 1 Then
                    vertexList(i).vertex(j) = 1
                End If
            Next
        Next

        applyLights True
    End If

    Me.MousePointer = vbCustom

    ctrlDown = False
    currentFunction = currentTool
    SetCursor currentFunction + 1
    lblCurrentTool.Caption = frmSoldatMapEditor.ImageList.ListImages(currentFunction + 1).Tag
    SaveUndo
    Render

    Exit Sub

ErrorHandler:

    MsgBox "Error averaging colors" & vbNewLine & Error$

End Sub

Private Sub MouseDownMove(X As Single, Y As Single)

    If numSelectedPolys + numSelectedScenery + numSelSpawns + numSelColliders + numSelWaypoints + numSelLights = 0 Then
        noneSelected = True
        SelNearest X, Y
    End If
    If snapToGrid And showGrid Then
        X = snapVertexToGrid(X, (scrollCoords(2).X - Int(scrollCoords(2).X / inc) * inc) * zoomFactor)
        Y = snapVertexToGrid(Y, (scrollCoords(2).Y - Int(scrollCoords(2).Y / inc) * inc) * zoomFactor)
    End If
    moveCoords(1).X = X
    moveCoords(1).Y = Y
    moveCoords(2).X = X
    moveCoords(2).Y = Y

End Sub

Private Sub SelNearest(X As Single, Y As Single)

    Dim i As Integer
    Dim j As Integer
    Dim addPoly As Integer
    Dim addVert As Integer
    Dim notSel As Integer
    Dim currentDist As Long
    Dim shortestDist As Long
    Dim xVal As Single
    Dim yVal As Single

    xVal = X / zoomFactor + scrollCoords(2).X
    yVal = Y / zoomFactor + scrollCoords(2).Y

    addPoly = 0
    shortestDist = 64 ^ 2 + 1
    If showPolys Then
    For i = 1 To mPolyCount
        For j = 1 To 3
            If nearCoord(X, Polys(i).vertex(j).X, 8) And nearCoord(Y, Polys(i).vertex(j).Y, 8) Then  ' move by vertex
                If addPoly <> i Then
                    numSelectedPolys = numSelectedPolys + 1
                    ReDim Preserve selectedPolys(numSelectedPolys)
                    selectedPolys(numSelectedPolys) = i
                End If
                vertexList(i).vertex(j) = 1
                addPoly = i
            End If
        Next
        If (pointInPoly(X, Y, i)) And addPoly = 0 Then
            For j = 1 To 3
                If nearCoord(X, Polys(i).vertex(j).X, 64) And nearCoord(Y, Polys(i).vertex(j).Y, 64) Then  ' move by region
                    currentDist = (Polys(i).vertex(j).X - X) ^ 2 + (Polys(i).vertex(j).Y - Y) ^ 2
                    If currentDist < shortestDist Then
                        shortestDist = currentDist
                        addPoly = i
                        addVert = j
                    End If
                End If
            Next
        End If
    Next
    End If

    If numSelectedPolys = 0 And addPoly > 0 Then
        numSelectedPolys = numSelectedPolys + 1
        ReDim Preserve selectedPolys(numSelectedPolys)
        selectedPolys(numSelectedPolys) = addPoly
        vertexList(addPoly).vertex(addVert) = 1
    End If

    If numSelectedPolys = 0 And addPoly = 0 And showScenery Then  ' select scenery
        For i = 1 To sceneryCount
            If PointInProp(X, Y, i) And addPoly = 0 Then
                Scenery(i).selected = 1
                numSelectedScenery = numSelectedScenery + 1
                addPoly = 1
            End If
        Next
    End If

    If addPoly = 0 And showObjects Then
        notSel = 0
        shortestDist = (8 ^ 2 + 1)
        For i = 1 To spawnPoints
            Spawns(i).active = 0
            If nearCoord(xVal, Spawns(i).X, 8 / zoomFactor) And nearCoord(yVal, Spawns(i).Y, 8 / zoomFactor) Then
                currentDist = (Spawns(i).X - xVal) ^ 2 + (Spawns(i).Y - yVal) ^ 2
                If currentDist < shortestDist Then
                    shortestDist = currentDist
                    notSel = i
                End If
            End If
        Next
        If notSel > 0 Then
            Spawns(notSel).active = 1
            numSelSpawns = numSelSpawns + 1
            addPoly = notSel
        End If
    End If

    If addPoly = 0 And showObjects Then
        notSel = 0
        shortestDist = 64 ^ 2 + 1
        For i = 1 To colliderCount
            Colliders(i).active = 0
            If nearCoord(xVal, Colliders(i).X, Colliders(i).radius / 2) And nearCoord(yVal, Colliders(i).Y, Colliders(i).radius / 2) Then
                currentDist = (Colliders(i).X - xVal) ^ 2 + (Colliders(i).Y - yVal) ^ 2
                If currentDist < shortestDist Then
                    shortestDist = currentDist
                    notSel = i
                End If
            End If
        Next
        If notSel > 0 Then
            Colliders(notSel).active = 1
            numSelColliders = numSelColliders + 1
            addPoly = notSel
        End If
    End If

    If addPoly = 0 And showWaypoints Then
        notSel = 0
        shortestDist = (8 ^ 2 + 1)
        For i = 1 To waypointCount
            Waypoints(i).selected = False
            If nearCoord(xVal, Waypoints(i).X, 8 / zoomFactor) And nearCoord(yVal, Waypoints(i).Y, 8 / zoomFactor) Then
                currentDist = (Waypoints(i).X - xVal) ^ 2 + (Waypoints(i).Y - yVal) ^ 2
                If currentDist < shortestDist Then
                    shortestDist = currentDist
                    notSel = i
                End If
            End If
        Next
        If notSel > 0 Then
            Waypoints(notSel).selected = True
            numSelWaypoints = numSelWaypoints + 1
        End If
    End If

    Render

End Sub

Private Sub CreatingPoly(Shift As Integer, X As Single, Y As Single)

    Dim xVal As Integer
    Dim yVal As Integer
    Dim rtheta As D3DVECTOR2

    xVal = X
    yVal = Y

    If Shift = KEY_SHIFT Then
        rtheta = ConstrainAngle(X - Polys(mPolyCount + 1).vertex(numVerts).X, Y - Polys(mPolyCount + 1).vertex(numVerts).Y)
        xVal = Polys(mPolyCount + 1).vertex(numVerts).X + rtheta.X * Cos(rtheta.Y)
        yVal = Polys(mPolyCount + 1).vertex(numVerts).Y + rtheta.X * Sin(rtheta.Y)
    End If

    If snapToGrid And showGrid Then
        xVal = snapVertexToGrid(xVal, (scrollCoords(2).X - Int(scrollCoords(2).X / inc) * inc) * zoomFactor)
        yVal = snapVertexToGrid(yVal, (scrollCoords(2).Y - Int(scrollCoords(2).Y / inc) * inc) * zoomFactor)
    End If

    Polys(mPolyCount + 1).vertex(numVerts + 1).X = xVal
    Polys(mPolyCount + 1).vertex(numVerts + 1).Y = yVal

    PolyCoords(mPolyCount + 1).vertex(numVerts + 1).X = xVal / zoomFactor + scrollCoords(2).X
    PolyCoords(mPolyCount + 1).vertex(numVerts + 1).Y = yVal / zoomFactor + scrollCoords(2).Y

    If mnuCustomX.Checked And mnuQuad.Checked Then
        If creatingQuad Then
            Polys(mPolyCount + 1).vertex(numVerts + 1).tu = (frmTexture.x1tex * 2 + 0.5) / xTexture
        Else
            If numVerts = 1 Or numVerts = 2 Then
                Polys(mPolyCount + 1).vertex(numVerts + 1).tu = (frmTexture.x2tex * 2 - 0.5) / xTexture
            Else
                Polys(mPolyCount + 1).vertex(numVerts + 1).tu = (frmTexture.x1tex * 2 + 0.5) / xTexture
            End If
        End If
    Else
        Polys(mPolyCount + 1).vertex(numVerts + 1).tu = (xVal / zoomFactor + scrollCoords(2).X) / xTexture
    End If

    If mnuCustomY.Checked And mnuQuad.Checked Then
        If creatingQuad Then
            Polys(mPolyCount + 1).vertex(numVerts + 1).tv = (frmTexture.y2tex * 2 - 0.5) / yTexture
        Else
            If numVerts > 1 Then
                Polys(mPolyCount + 1).vertex(numVerts + 1).tv = (frmTexture.y2tex * 2 - 0.5) / yTexture
            Else
                Polys(mPolyCount + 1).vertex(numVerts + 1).tv = (frmTexture.y1tex * 2 + 0.5) / yTexture
            End If
        End If
    Else
        Polys(mPolyCount + 1).vertex(numVerts + 1).tv = (yVal / zoomFactor + scrollCoords(2).Y) / yTexture
    End If

    Polys(mPolyCount + 1).vertex(numVerts + 1).color = ARGB(255 * opacity, RGB(gPolyClr.blue, gPolyClr.green, gPolyClr.red))

    Render

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error GoTo ErrorHandler

    ' if not in focus and not properties in focus and not text in focus
    If Screen.ActiveForm Is Nothing Then
        Exit Sub
    ElseIf Screen.ActiveForm.hWnd <> Me.hWnd And Screen.ActiveForm.hWnd <> frmInfo.hWnd Then
        If Not (Screen.ActiveForm.hWnd = frmPalette.hWnd And frmPalette.textControl) Then
            RegainFocus
        End If
    End If

    mouseCoords.X = X
    mouseCoords.Y = Y

    lblMousePosition.Caption = "Position: " & Round(X / zoomFactor + scrollCoords(2).X) & ", " & Round(Y / zoomFactor + scrollCoords(2).Y)

    ' draw circle
    If circleOn Then
        Render
    End If

    If Button = 4 Or Button = 5 Then  ' scrolling
        Scrolling X, Y
    End If

    If (currentFunction = TOOL_CREATE Or currentFunction = TOOL_QUAD) And toolAction Then
        If (Shift = 0 Or Shift = KEY_SHIFT) And numVerts > 0 Then  ' poly creation
            CreatingPoly Shift, X, Y
        End If
    ElseIf Button = 0 And currentFunction = TOOL_SCENERY Then
        CreatingScenery Shift, X, Y
    ElseIf Button = 0 And currentFunction = TOOL_OBJECTS And Shift < 2 Then
        Spawns(0).X = X
        Spawns(0).Y = Y
        Colliders(0).X = X
        Colliders(0).Y = Y
        Render
    ElseIf Button = 0 And (currentFunction = TOOL_WAYPOINT Or currentFunction = TOOL_CONNECT) And currentWaypoint > 0 Then
        Render
    ElseIf Button = 0 And currentFunction = TOOL_SKETCH And toolAction Then
        sketch(0).vertex(2).X = X / zoomFactor + scrollCoords(2).X
        sketch(0).vertex(2).Y = Y / zoomFactor + scrollCoords(2).Y
        Render
    End If

    If Button <> 1 Then Exit Sub

    If spaceDown Then  ' scrolling
        If currentFunction = TOOL_SCENERY And numCorners = 0 Then
            Scenery(0).screenTr.X = X
            Scenery(0).screenTr.Y = Y
        ElseIf currentFunction = TOOL_OBJECTS Then
            If Not mnuCollider.Checked Then
                Spawns(0).X = X
                Spawns(0).Y = Y
            ElseIf mnuCollider.Checked Then
                Colliders(0).X = X
                Colliders(0).Y = Y
            End If
        End If

        Scrolling X, Y

        If Button = 5 Then
            moveCoords(1).X = X
            moveCoords(1).Y = Y
        End If
    ElseIf currentFunction = TOOL_MOVE And toolAction Then  ' moving
        If Shift = KEY_SHIFT Then  ' constrained
            If Abs(X - moveCoords(2).X) > Abs(Y - moveCoords(2).Y) Then
                Y = moveCoords(2).Y
            Else
                X = moveCoords(2).X
            End If
        End If
        Moving X, Y
    ElseIf currentFunction = TOOL_SCALE And toolAction Then  ' scaling
        If Shift = KEY_CTRL Then
            Scaling X, Y, False
        ElseIf Shift = KEY_SHIFT + KEY_CTRL Then  ' constrained scaling
            Scaling X, Y, True
        End If
    ElseIf currentFunction = TOOL_ROTATE And toolAction Then  ' rotating
        If Shift = KEY_ALT Then
            Rotating X, Y, False
        ElseIf Shift = KEY_SHIFT + KEY_ALT Then  ' constrained rotating
            Rotating X, Y, True
        End If
    ElseIf (currentFunction = TOOL_CREATE Or currentFunction = TOOL_CREATE) And toolAction Then  ' poly creation
        ' no-op
    ElseIf currentFunction = TOOL_VSELECT Or currentFunction = TOOL_VSELADD Or currentFunction = TOOL_VSELSUB Then  ' vertex selection
        If toolAction Then
            Render
            selectedCoords(2).X = X
            selectedCoords(2).Y = Y
        End If
    ElseIf currentFunction = TOOL_PSELECT And toolAction Then  ' poly selection
        ' no-op
    ElseIf currentFunction = TOOL_VCOLOR And toolAction Then  ' vertex coloring
        If colorMode > 0 Then
            VertexColoring X, Y
        End If
    ElseIf currentFunction = TOOL_PCOLOR Then  ' poly coloring
        ' no-op
    ElseIf currentFunction = TOOL_TEXTURE And toolAction Then  ' texture
        If Shift = 0 Then
            StretchingTexture X, Y
        ElseIf Shift = KEY_SHIFT Then
            If Abs(X - moveCoords(2).X) > Abs(Y - moveCoords(2).Y) Then
                Y = moveCoords(2).Y
            Else
                X = moveCoords(2).X
            End If
            StretchingTexture X, Y
        End If
    ElseIf currentFunction = TOOL_SCENERY Then  ' scenery
        ' no-op
    ElseIf currentFunction = TOOL_COLORPICKER Then  ' color picker

        If currentTool = TOOL_DEPTHMAP Then
            depthPicker X, Y
        ElseIf currentTool = TOOL_SCENERY Then

        Else
            ColorPicker X, Y
        End If
    ElseIf currentFunction = TOOL_PIXPICKER Then  ' pixel picker
        Dim tempClr As TColor
        tempClr = getRGB(GetPixel(Me.hDC, X, Y))
        If frmPalette.Enabled = False Then
            frmColor.InitColor tempClr.blue, tempClr.green, tempClr.red
        Else
            gPolyClr.red = tempClr.blue
            gPolyClr.green = tempClr.green
            gPolyClr.blue = tempClr.red
            Scenery(0).color = ARGB(Scenery(0).alpha, RGB(gPolyClr.blue, gPolyClr.green, gPolyClr.red))
            frmPalette.setValues gPolyClr.red, gPolyClr.green, gPolyClr.blue
        End If

        Render
    ElseIf currentFunction = TOOL_LITPICKER Then  ' light picker
        ' no-op
    ElseIf currentFunction = TOOL_OBJECTS Then  ' objects
        Spawns(0).X = X
        Spawns(0).Y = Y
        If mnuGostek.Checked Then
            gostek.X = X / zoomFactor + scrollCoords(2).X
            gostek.Y = Y / zoomFactor + scrollCoords(2).Y
            Render
        End If
    ElseIf currentFunction = TOOL_WAYPOINT And toolAction Then  ' waypoints
        ' no-op
    ElseIf currentFunction = TOOL_DEPTHMAP And toolAction Then  ' depthmap
        EditDepthMap X, Y
    ElseIf currentFunction = TOOL_SKETCH And toolAction Then  ' sketch
        If Shift = 0 Then  ' freeform
            linkSketch X, Y
            sketch(sketchLines).vertex(2).X = X / zoomFactor + scrollCoords(2).X
            sketch(sketchLines).vertex(2).Y = Y / zoomFactor + scrollCoords(2).Y
            Render
        ElseIf Shift = KEY_SHIFT Then ' lines
            sketch(0).vertex(2).X = X / zoomFactor + scrollCoords(2).X
            sketch(0).vertex(2).Y = Y / zoomFactor + scrollCoords(2).Y
            Render
        End If
    ElseIf currentFunction = TOOL_ERASER And toolAction Then
        If eraseSketch(X / zoomFactor + scrollCoords(2).X, Y / zoomFactor + scrollCoords(2).Y) = 1 Then
            Render
        End If
    ElseIf currentFunction = TOOL_SMUDGE And toolAction Then
        If moveLines(X / zoomFactor + scrollCoords(2).X, Y / zoomFactor + scrollCoords(2).Y, X - moveCoords(2).X, Y - moveCoords(2).Y) = 1 Then
            Render
        End If
        moveCoords(2).X = X
        moveCoords(2).Y = Y
    End If

    Exit Sub

ErrorHandler:

    MsgBox "form_mousemove error" & vbNewLine & Error$

End Sub

Private Sub CreatingScenery(Shift As Integer, X As Single, Y As Single)

    Dim xVal As Single
    Dim yVal As Single
    Dim angle As Single

    xVal = X
    yVal = Y

    If snapToGrid And showGrid Then
        xVal = Int(snapVertexToGrid(X, (scrollCoords(2).X - Int(scrollCoords(2).X / inc) * inc) * zoomFactor) + 0.5)
        yVal = Int(snapVertexToGrid(Y, (scrollCoords(2).Y - Int(scrollCoords(2).Y / inc) * inc) * zoomFactor) + 0.5)
    End If

    If numCorners = 0 Then
        Scenery(0).screenTr.X = xVal
        Scenery(0).screenTr.Y = yVal
    End If

    xVal = xVal - Int(Scenery(0).screenTr.X + 0.5)
    yVal = yVal - Int(Scenery(0).screenTr.Y + 0.5)

    angle = GetAngle(xVal, yVal)

    If numCorners = 1 And toolAction Then
        If Shift = 1 Then
            angle = (Int((angle * 180 / PI + 0) / 15) * 15) / 180 * PI
        End If

        Scenery(0).rotation = angle
    ElseIf numCorners = 2 And toolAction Then
        angle = angle - Scenery(0).rotation

        Scenery(0).Scaling.X = (Cos(angle) * Sqr(xVal ^ 2 + yVal ^ 2)) / (SceneryTextures(Scenery(0).Style).Width) / zoomFactor
        Scenery(0).Scaling.Y = -(Sin(angle) * Sqr(xVal ^ 2 + yVal ^ 2)) / (SceneryTextures(Scenery(0).Style).Height) / zoomFactor

        If Shift = 1 Then
            If Scenery(0).Scaling.X < 0 Then
                Scenery(0).Scaling.X = -Sqr((xVal ^ 2 + yVal ^ 2) / (SceneryTextures(Scenery(0).Style).Width ^ 2 + SceneryTextures(Scenery(0).Style).Height ^ 2)) / zoomFactor
            Else
                Scenery(0).Scaling.X = Sqr((xVal ^ 2 + yVal ^ 2) / (SceneryTextures(Scenery(0).Style).Width ^ 2 + SceneryTextures(Scenery(0).Style).Height ^ 2)) / zoomFactor
            End If
            If Scenery(0).Scaling.Y * Scenery(0).Scaling.X < 0 Then
                Scenery(0).Scaling.Y = -Scenery(0).Scaling.X
            Else
                Scenery(0).Scaling.Y = Scenery(0).Scaling.X
            End If
        End If
    End If

    Render

End Sub

Private Function ConstrainAngle(xDiff As Integer, yDiff As Integer) As D3DVECTOR2

    Dim theta As Single
    Dim R As Single

    R = Sqr(xDiff ^ 2 + yDiff ^ 2)
    If xDiff = 0 Then
        If yDiff > 0 Then
            theta = PI / 2
        Else
            theta = 3 * PI / 2
        End If
    ElseIf xDiff > 0 Then
        theta = Atn(yDiff / xDiff)
    ElseIf xDiff < 0 Then
        theta = PI + Atn(yDiff / xDiff)
    End If

    theta = (Int((theta * 180 / PI + 7.5) / 15) * 15) / 180 * PI

    ConstrainAngle.X = R
    ConstrainAngle.Y = theta

End Function

Private Sub Scrolling(X As Single, Y As Single)

    Dim i As Integer

    scrollCoords(2).X = scrollCoords(2).X - (X - scrollCoords(1).X) / zoomFactor
    scrollCoords(2).Y = scrollCoords(2).Y - (Y - scrollCoords(1).Y) / zoomFactor

    For i = 1 To mPolyCount  ' move polys
        Polys(i).vertex(1).X = Polys(i).vertex(1).X + X - scrollCoords(1).X
        Polys(i).vertex(1).Y = Polys(i).vertex(1).Y + Y - scrollCoords(1).Y
        Polys(i).vertex(2).X = Polys(i).vertex(2).X + X - scrollCoords(1).X
        Polys(i).vertex(2).Y = Polys(i).vertex(2).Y + Y - scrollCoords(1).Y
        Polys(i).vertex(3).X = Polys(i).vertex(3).X + X - scrollCoords(1).X
        Polys(i).vertex(3).Y = Polys(i).vertex(3).Y + Y - scrollCoords(1).Y
    Next

    For i = 1 To 4  ' move background
        bgPolys(i).X = bgPolys(i).X + X - scrollCoords(1).X
        bgPolys(i).Y = bgPolys(i).Y + Y - scrollCoords(1).Y
    Next

    For i = 1 To sceneryCount
        Scenery(i).screenTr.X = Scenery(i).screenTr.X + X - scrollCoords(1).X
        Scenery(i).screenTr.Y = Scenery(i).screenTr.Y + Y - scrollCoords(1).Y
    Next

    If numVerts > 0 Then  ' move existing vertices of poly being created
        For i = 1 To 3
            Polys(mPolyCount + 1).vertex(i).X = Polys(mPolyCount + 1).vertex(i).X + X - scrollCoords(1).X
            Polys(mPolyCount + 1).vertex(i).Y = Polys(mPolyCount + 1).vertex(i).Y + Y - scrollCoords(1).Y
        Next
    End If

    If numCorners > 0 Then
        Scenery(0).screenTr.X = Scenery(0).screenTr.X + X - scrollCoords(1).X
        Scenery(0).screenTr.Y = Scenery(0).screenTr.Y + Y - scrollCoords(1).Y
    ElseIf currentFunction = TOOL_SCENERY And numCorners = 0 Then
        Scenery(0).screenTr.X = X
        Scenery(0).screenTr.Y = Y
    ElseIf currentFunction = TOOL_OBJECTS Then
        Spawns(0).X = X
        Spawns(0).Y = Y
        Colliders(0).X = X
        Colliders(0).Y = Y
    End If

    If (currentFunction = TOOL_VSELECT Or currentFunction = TOOL_VSELADD Or currentFunction = TOOL_VSELSUB) And toolAction Then
        selectedCoords(1).X = selectedCoords(1).X + X - scrollCoords(1).X
        selectedCoords(1).Y = selectedCoords(1).Y + Y - scrollCoords(1).Y
        selectedCoords(2).X = X
        selectedCoords(2).Y = Y
    End If

    scrollCoords(1).X = X
    scrollCoords(1).Y = Y

    Render

    If (currentFunction = TOOL_VSELECT Or currentFunction = TOOL_VSELADD Or currentFunction = TOOL_VSELSUB) And toolAction Then
        Render
    End If

End Sub

Private Sub Moving(ByVal X As Single, ByVal Y As Single)

    Dim i As Integer
    Dim j As Integer
    Dim PolyNum As Integer
    Dim xDiff As Single
    Dim yDiff As Single
    Dim xVal As Single
    Dim yVal As Single

    If snapToGrid And showGrid And toolAction Then
        X = snapVertexToGrid(X, (scrollCoords(2).X - Int(scrollCoords(2).X / inc) * inc) * zoomFactor)
        Y = snapVertexToGrid(Y, (scrollCoords(2).Y - Int(scrollCoords(2).Y / inc) * inc) * zoomFactor)
    End If

    xVal = X - moveCoords(1).X
    yVal = Y - moveCoords(1).Y

    For i = 1 To numSelectedPolys
        PolyNum = selectedPolys(i)
        For j = 1 To 3
            If vertexList(PolyNum).vertex(j) = 1 Then
                xDiff = Polys(PolyNum).vertex(j).tu - PolyCoords(PolyNum).vertex(j).X / xTexture
                yDiff = Polys(PolyNum).vertex(j).tv - PolyCoords(PolyNum).vertex(j).Y / yTexture
                PolyCoords(PolyNum).vertex(j).X = PolyCoords(PolyNum).vertex(j).X + xVal / zoomFactor
                PolyCoords(PolyNum).vertex(j).Y = PolyCoords(PolyNum).vertex(j).Y + yVal / zoomFactor
                ' switch
                Polys(PolyNum).vertex(j).X = (PolyCoords(PolyNum).vertex(j).X - scrollCoords(2).X) * zoomFactor
                Polys(PolyNum).vertex(j).Y = (PolyCoords(PolyNum).vertex(j).Y - scrollCoords(2).Y) * zoomFactor

                If fixedTexture Then
                    Polys(PolyNum).vertex(j).tu = (Polys(PolyNum).vertex(j).X / zoomFactor + scrollCoords(2).X) / xTexture + xDiff
                    Polys(PolyNum).vertex(j).tv = (Polys(PolyNum).vertex(j).Y / zoomFactor + scrollCoords(2).Y) / yTexture + yDiff
                End If
            End If
        Next
    Next

    For i = 1 To sceneryCount
        If Scenery(i).selected = 1 Then
            Scenery(i).Translation.X = Scenery(i).Translation.X + xVal / zoomFactor
            Scenery(i).Translation.Y = Scenery(i).Translation.Y + yVal / zoomFactor
            Scenery(i).screenTr.X = (Scenery(i).Translation.X - scrollCoords(2).X) * zoomFactor
            Scenery(i).screenTr.Y = (Scenery(i).Translation.Y - scrollCoords(2).Y) * zoomFactor
        End If
    Next

    For i = 1 To spawnPoints
        If Spawns(i).active = 1 Then
            Spawns(i).X = Spawns(i).X + xVal / zoomFactor
            Spawns(i).Y = Spawns(i).Y + yVal / zoomFactor
        End If
    Next
    For i = 1 To colliderCount
        If Colliders(i).active = 1 Then
            Colliders(i).X = Colliders(i).X + xVal / zoomFactor
            Colliders(i).Y = Colliders(i).Y + yVal / zoomFactor
        End If
    Next

    For i = 1 To lightCount
        If Lights(i).selected = 1 Then
            Lights(i).X = Lights(i).X + xVal / zoomFactor
            Lights(i).Y = Lights(i).Y + yVal / zoomFactor
        End If
    Next

    For i = 1 To waypointCount
        If Waypoints(i).selected = True Then
            Waypoints(i).X = Waypoints(i).X + xVal / zoomFactor
            Waypoints(i).Y = Waypoints(i).Y + yVal / zoomFactor
        End If
    Next

    rCenter.X = rCenter.X + xVal / zoomFactor
    rCenter.Y = rCenter.Y + yVal / zoomFactor

    For i = 0 To 3
        selRect(i).X = selRect(i).X + xVal / zoomFactor
        selRect(i).Y = selRect(i).Y + yVal / zoomFactor
    Next

    moveCoords(1).X = X
    moveCoords(1).Y = Y

    getInfo

    prompt = True

    Render

End Sub

Private Sub Scaling(ByVal X As Single, ByVal Y As Single, constrained As Boolean)

    Dim i As Integer
    Dim j As Integer
    Dim xVal As Single
    Dim yVal As Single
    Dim xCenter As Single
    Dim yCenter As Single
    Dim PolyNum As Integer
    Dim theta As Single

    xCenter = (rCenter.X - scrollCoords(2).X) * zoomFactor
    yCenter = (rCenter.Y - scrollCoords(2).Y) * zoomFactor

    If snapToGrid And showGrid Then
        X = snapVertexToGrid(X, (scrollCoords(2).X - Int(scrollCoords(2).X / inc) * inc) * zoomFactor)
        Y = snapVertexToGrid(Y, (scrollCoords(2).Y - Int(scrollCoords(2).Y / inc) * inc) * zoomFactor)
    End If

    If Not constrained Then
        If moveCoords(1).X = xCenter Then
            scaleDiff.X = 1
        Else
            scaleDiff.X = 1 + (X - moveCoords(1).X) / (moveCoords(1).X - xCenter)
        End If
        If moveCoords(1).Y = yCenter Then
            scaleDiff.Y = 1
        Else
            scaleDiff.Y = 1 + (Y - moveCoords(1).Y) / (moveCoords(1).Y - yCenter)
        End If
    Else
        If (moveCoords(1).X - xCenter) * (moveCoords(1).Y - yCenter) > 0 Then
            scaleDiff.X = (((X - xCenter) + (Y - yCenter)) / ((moveCoords(1).X - xCenter) + (moveCoords(1).Y - yCenter)))
            scaleDiff.Y = scaleDiff.X
        Else
            scaleDiff.X = (((X - xCenter) - (Y - yCenter)) / ((moveCoords(1).X - xCenter) - (moveCoords(1).Y - yCenter)))
            scaleDiff.Y = scaleDiff.X
        End If

    End If

    If numSelectedPolys > 0 Then
        For i = 1 To numSelectedPolys
            PolyNum = selectedPolys(i)
            For j = 1 To 3
                If vertexList(PolyNum).vertex(j) = 1 Then
                    Polys(PolyNum).vertex(j).X = ((rCenter.X + (PolyCoords(PolyNum).vertex(j).X - rCenter.X) * scaleDiff.X) - scrollCoords(2).X) * zoomFactor
                    Polys(PolyNum).vertex(j).Y = ((rCenter.Y + (PolyCoords(PolyNum).vertex(j).Y - rCenter.Y) * scaleDiff.Y) - scrollCoords(2).Y) * zoomFactor
                End If
            Next
        Next
    End If

    If numSelectedScenery > 0 Then
        For i = 1 To sceneryCount
            If Scenery(i).selected = 1 Then
                Scenery(i).screenTr.X = (rCenter.X + (Scenery(i).Translation.X - rCenter.X) * scaleDiff.X - scrollCoords(2).X) * zoomFactor
                Scenery(i).screenTr.Y = (rCenter.Y + (Scenery(i).Translation.Y - rCenter.Y) * scaleDiff.Y - scrollCoords(2).Y) * zoomFactor
            End If
        Next
    End If

    moveCoords(2).X = X
    moveCoords(2).Y = Y

    frmInfo.txtScale(0).Text = Int(scaleDiff.X * 1000) / 10
    frmInfo.txtScale(1).Text = Int(scaleDiff.Y * 1000) / 10

    prompt = True

    Render

End Sub

Private Sub ApplyTransform(Rotating As Boolean)

    Dim i As Integer
    Dim j As Integer
    Dim pNum As Integer
    Dim temp As D3DVECTOR2
    Dim tempVertex As TCustomVertex
    Dim vertSel As Byte
    Dim xVal As Single
    Dim yVal As Single
    Dim angle As Single
    Dim theta As Single
    Dim R As Single
    Dim tempClr As TColor

    If selectionChanged Then
        SaveUndo
        selectionChanged = False
    End If

    For i = 1 To numSelectedPolys
        pNum = selectedPolys(i)
        For j = 1 To 3
            PolyCoords(pNum).vertex(j).X = Polys(pNum).vertex(j).X / zoomFactor + scrollCoords(2).X
            PolyCoords(pNum).vertex(j).Y = Polys(pNum).vertex(j).Y / zoomFactor + scrollCoords(2).Y

            If (scaleDiff.X * scaleDiff.Y < 0) Then
                ' make sure polys are cw
                If Not isCW(pNum) Then  ' switch to make cw
                    temp = PolyCoords(pNum).vertex(3)
                    PolyCoords(pNum).vertex(3) = PolyCoords(pNum).vertex(2)
                    PolyCoords(pNum).vertex(2) = temp

                    tempVertex = Polys(pNum).vertex(3)
                    Polys(pNum).vertex(3) = Polys(pNum).vertex(2)
                    Polys(pNum).vertex(2) = tempVertex

                    vertSel = vertexList(pNum).vertex(3)
                    vertexList(pNum).vertex(3) = vertexList(pNum).vertex(2)
                    vertexList(pNum).vertex(2) = vertSel

                    tempClr = vertexList(pNum).color(3)
                    vertexList(pNum).color(3) = vertexList(pNum).color(2)
                    vertexList(pNum).color(2) = tempClr
                End If
            End If
        Next
    Next

    If numSelectedScenery > 0 Then
    For i = 1 To sceneryCount
        If Scenery(i).selected = 1 Then
            Scenery(i).Translation.X = Scenery(i).screenTr.X / zoomFactor + scrollCoords(2).X
            Scenery(i).Translation.Y = Scenery(i).screenTr.Y / zoomFactor + scrollCoords(2).Y

            If Not Rotating Then
                xVal = SceneryTextures(Scenery(i).Style).Width * Scenery(i).Scaling.X
                yVal = SceneryTextures(Scenery(i).Style).Height * Scenery(i).Scaling.Y
                angle = GetAngle(xVal, yVal) + Scenery(i).rotation
                R = Sqr(xVal ^ 2 + yVal ^ 2)

                xVal = Cos(angle) * R * scaleDiff.X
                yVal = -Sin(angle) * R * scaleDiff.Y
                angle = GetAngle(xVal, yVal) - Scenery(i).rotation
                R = Sqr(xVal ^ 2 + yVal ^ 2)

                Scenery(i).Scaling.X = (Cos(angle) * R) / (SceneryTextures(Scenery(i).Style).Width)
                Scenery(i).Scaling.Y = -(Sin(angle) * R) / (SceneryTextures(Scenery(i).Style).Height)
            End If

            If scaleDiff.X * scaleDiff.Y < 0 And Rotating Then
                Scenery(i).rotation = -(Scenery(i).rotation - rDiff)
            Else
                Scenery(i).rotation = (Scenery(i).rotation - rDiff)
            End If
        End If
    Next
    End If

    If Not Rotating Then
        For i = 0 To 3
            selRect(i).X = rCenter.X + (selRect(i).X - rCenter.X) * scaleDiff.X
            selRect(i).Y = rCenter.Y + (selRect(i).Y - rCenter.Y) * scaleDiff.Y
        Next
    Else
        For i = 0 To 3
            xVal = (selRect(i).X - rCenter.X)
            yVal = (selRect(i).Y - rCenter.Y)
            R = Sqr((xVal) ^ 2 + (yVal) ^ 2)  ' distance of point from rotation center
            angle = GetAngle(xVal, yVal) - rDiff
            selRect(i).X = rCenter.X + R * Cos(angle)
            selRect(i).Y = rCenter.Y + R * -Sin(angle)
        Next
    End If

    scaleDiff.X = 1
    scaleDiff.Y = 1

    rDiff = 0

    getRCenter

    SaveUndo

    getInfo

    Render

End Sub

Public Sub applyScale(tehXvalue As Single, tehYvalue As Single)

    Dim i As Integer
    Dim j As Integer
    Dim PolyNum As Integer
    Dim vertSel As Byte
    Dim temp As D3DVECTOR2
    Dim tempVertex As TCustomVertex
    Dim xVal As Single
    Dim yVal As Single
    Dim R As Single
    Dim angle As Single

    If selectionChanged Then
        SaveUndo
        selectionChanged = False
    End If

    scaleDiff.X = tehXvalue
    scaleDiff.Y = tehYvalue

    rCenter.X = Midpoint(selRect(0).X, selRect(2).X)
    rCenter.Y = Midpoint(selRect(0).Y, selRect(2).Y)

    If numSelectedPolys > 0 Then
        For i = 1 To numSelectedPolys
            PolyNum = selectedPolys(i)
            For j = 1 To 3
                If vertexList(PolyNum).vertex(j) = 1 Then
                    PolyCoords(PolyNum).vertex(j).X = (rCenter.X + (PolyCoords(PolyNum).vertex(j).X - rCenter.X) * scaleDiff.X)
                    PolyCoords(PolyNum).vertex(j).Y = (rCenter.Y + (PolyCoords(PolyNum).vertex(j).Y - rCenter.Y) * scaleDiff.Y)
                    Polys(PolyNum).vertex(j).X = (PolyCoords(PolyNum).vertex(j).X - scrollCoords(2).X) * zoomFactor
                    Polys(PolyNum).vertex(j).Y = (PolyCoords(PolyNum).vertex(j).Y - scrollCoords(2).Y) * zoomFactor
                End If
            Next

            ' make sure polys are cw
            If Not isCW(PolyNum) Then  ' switch to make cw
                temp = PolyCoords(PolyNum).vertex(3)
                PolyCoords(PolyNum).vertex(3) = PolyCoords(PolyNum).vertex(2)
                PolyCoords(PolyNum).vertex(2) = temp

                tempVertex = Polys(PolyNum).vertex(3)
                Polys(PolyNum).vertex(3) = Polys(PolyNum).vertex(2)
                Polys(PolyNum).vertex(2) = tempVertex

                vertSel = vertexList(PolyNum).vertex(3)
                vertexList(PolyNum).vertex(3) = vertexList(PolyNum).vertex(2)
                vertexList(PolyNum).vertex(2) = vertSel
            End If
        Next
    End If

    If numSelectedScenery > 0 Then
        For i = 1 To sceneryCount
            If Scenery(i).selected = 1 Then

                Scenery(i).Translation.X = rCenter.X + (Scenery(i).Translation.X - rCenter.X) * scaleDiff.X
                Scenery(i).Translation.Y = rCenter.Y + (Scenery(i).Translation.Y - rCenter.Y) * scaleDiff.Y

                Scenery(i).screenTr.X = (Scenery(i).Translation.X - scrollCoords(2).X) * zoomFactor
                Scenery(i).screenTr.Y = (Scenery(i).Translation.Y - scrollCoords(2).Y) * zoomFactor

                xVal = SceneryTextures(Scenery(i).Style).Width * Scenery(i).Scaling.X
                yVal = SceneryTextures(Scenery(i).Style).Height * Scenery(i).Scaling.Y
                angle = GetAngle(xVal, yVal) + Scenery(i).rotation
                R = Sqr(xVal ^ 2 + yVal ^ 2)

                xVal = Cos(angle) * R * scaleDiff.X
                yVal = -Sin(angle) * R * scaleDiff.Y
                angle = GetAngle(xVal, yVal) - Scenery(i).rotation
                R = Sqr(xVal ^ 2 + yVal ^ 2)

                Scenery(i).Scaling.X = (Cos(angle) * R) / (SceneryTextures(Scenery(i).Style).Width)
                Scenery(i).Scaling.Y = -(Sin(angle) * R) / (SceneryTextures(Scenery(i).Style).Height)
            End If
        Next
    End If

    ' MESS!
    If numSelSpawns > 0 Then
        For i = 1 To spawnPoints
            If Spawns(i).active = 1 Then
                Spawns(i).X = rCenter.X + (Spawns(i).X - rCenter.X) * scaleDiff.X
                Spawns(i).Y = rCenter.Y + (Spawns(i).Y - rCenter.Y) * scaleDiff.Y
            End If
        Next
    End If

    If numSelColliders > 0 Then
        For i = 1 To colliderCount
            If Colliders(i).active = 1 Then
                Colliders(i).X = rCenter.X + (Colliders(i).X - rCenter.X) * scaleDiff.X
                Colliders(i).Y = rCenter.Y + (Colliders(i).Y - rCenter.Y) * scaleDiff.Y
            End If
        Next
    End If

    If numSelLights > 0 Then
        For i = 1 To lightCount
            If Lights(i).selected = 1 Then
                Lights(i).X = rCenter.X + (Lights(i).X - rCenter.X) * scaleDiff.X
                Lights(i).Y = rCenter.Y + (Lights(i).Y - rCenter.Y) * scaleDiff.Y
            End If
        Next
    End If

    If numSelWaypoints > 0 Then
        For i = 1 To waypointCount
            If Waypoints(i).selected = True Then
                Waypoints(i).X = rCenter.X + (Waypoints(i).X - rCenter.X) * scaleDiff.X
                Waypoints(i).Y = rCenter.Y + (Waypoints(i).Y - rCenter.Y) * scaleDiff.Y
            End If
        Next
    End If

    scaleDiff.X = 1
    scaleDiff.Y = 1

    getRCenter
    getInfo
    SaveUndo
    Render

End Sub

Public Sub applyRotate(tehValue As Single)

    Dim R As Single
    Dim theta As Single
    Dim xDiff As Single
    Dim yDiff As Single
    Dim i As Integer
    Dim j As Integer
    Dim PolyNum As Integer

    If selectionChanged Then
        SaveUndo
        selectionChanged = False
    End If

    rDiff = tehValue

    rCenter.X = Midpoint(selRect(0).X, selRect(2).X)
    rCenter.Y = Midpoint(selRect(0).Y, selRect(2).Y)

    If numSelectedPolys > 0 Then
        For i = 1 To numSelectedPolys
            PolyNum = selectedPolys(i)
            For j = 1 To 3
                If vertexList(PolyNum).vertex(j) = 1 Then
                    xDiff = (PolyCoords(PolyNum).vertex(j).X - rCenter.X)
                    yDiff = (PolyCoords(PolyNum).vertex(j).Y - rCenter.Y)

                    theta = rDiff

                    PolyCoords(PolyNum).vertex(j).X = (Cos(theta) * xDiff - Sin(theta) * yDiff) + rCenter.X
                    PolyCoords(PolyNum).vertex(j).Y = (Sin(theta) * xDiff + Cos(theta) * yDiff) + rCenter.Y

                    Polys(PolyNum).vertex(j).X = (PolyCoords(PolyNum).vertex(j).X - scrollCoords(2).X) * zoomFactor
                    Polys(PolyNum).vertex(j).Y = (PolyCoords(PolyNum).vertex(j).Y - scrollCoords(2).Y) * zoomFactor
                End If
            Next
        Next
    End If

    If numSelectedScenery > 0 Then
        For i = 1 To sceneryCount
            If Scenery(i).selected = 1 Then
                xDiff = (Scenery(i).Translation.X - rCenter.X)
                yDiff = (Scenery(i).Translation.Y - rCenter.Y)

                R = Sqr((xDiff) ^ 2 + (yDiff) ^ 2)  ' distance of point from rotation center
                If xDiff = 0 Then
                    If yDiff > 0 Then
                        theta = PI / 2
                    Else
                        theta = 3 * PI / 2
                    End If
                ElseIf xDiff > 0 Then
                    theta = Atn(yDiff / xDiff)
                ElseIf xDiff < 0 Then
                    theta = PI + Atn(yDiff / xDiff)
                End If
                theta = theta + rDiff

                Scenery(i).Translation.X = rCenter.X + R * Cos(theta)
                Scenery(i).Translation.Y = rCenter.Y + R * Sin(theta)

                Scenery(i).screenTr.X = (Scenery(i).Translation.X - scrollCoords(2).X) * zoomFactor
                Scenery(i).screenTr.Y = (Scenery(i).Translation.Y - scrollCoords(2).Y) * zoomFactor

                If scaleDiff.X * scaleDiff.Y < 0 Then
                    Scenery(i).rotation = -(Scenery(i).rotation - rDiff)
                Else
                    Scenery(i).rotation = (Scenery(i).rotation - rDiff)
                End If
            End If
        Next
    End If

    ' MESS!
    If numSelSpawns > 0 Then
        For i = 1 To spawnPoints
            If Spawns(i).active = 1 Then
                xDiff = (Spawns(i).X - rCenter.X)
                yDiff = (Spawns(i).Y - rCenter.Y)
                theta = rDiff
                Spawns(i).X = (Cos(theta) * xDiff - Sin(theta) * yDiff) + rCenter.X
                Spawns(i).Y = (Sin(theta) * xDiff + Cos(theta) * yDiff) + rCenter.Y
            End If
        Next
    End If

    If numSelColliders > 0 Then
        For i = 1 To colliderCount
            If Colliders(i).active = 1 Then
                xDiff = (Colliders(i).X - rCenter.X)
                yDiff = (Colliders(i).Y - rCenter.Y)
                theta = rDiff
                Colliders(i).X = (Cos(theta) * xDiff - Sin(theta) * yDiff) + rCenter.X
                Colliders(i).Y = (Sin(theta) * xDiff + Cos(theta) * yDiff) + rCenter.Y
            End If
        Next
    End If

    If numSelLights > 0 Then
        For i = 1 To lightCount
            If Lights(i).selected = 1 Then
                xDiff = (Lights(i).X - rCenter.X)
                yDiff = (Lights(i).Y - rCenter.Y)
                theta = rDiff
                Lights(i).X = (Cos(theta) * xDiff - Sin(theta) * yDiff) + rCenter.X
                Lights(i).Y = (Sin(theta) * xDiff + Cos(theta) * yDiff) + rCenter.Y
            End If
        Next
    End If

    If numSelWaypoints > 0 Then
        For i = 1 To waypointCount
            If Waypoints(i).selected = True Then
                xDiff = (Waypoints(i).X - rCenter.X)
                yDiff = (Waypoints(i).Y - rCenter.Y)
                theta = rDiff
                Waypoints(i).X = (Cos(theta) * xDiff - Sin(theta) * yDiff) + rCenter.X
                Waypoints(i).Y = (Sin(theta) * xDiff + Cos(theta) * yDiff) + rCenter.Y
            End If
        Next
    End If

    rCenter.X = selRect(0).X
    rCenter.Y = selRect(0).Y
    rDiff = 0

    getRCenter
    getInfo
    SaveUndo
    Render

End Sub

Private Sub Rotating(X As Single, Y As Single, constrained As Boolean)

    Dim i As Integer
    Dim j As Integer
    Dim angle As Single
    Dim oldAngle As Single
    Dim xCenter As Single
    Dim yCenter As Single
    Dim xDiff As Integer
    Dim yDiff As Integer
    Dim PolyNum As Integer
    Dim R As Single
    Dim theta As Single

    If snapToGrid And showGrid Then
        X = snapVertexToGrid(X, (scrollCoords(2).X - Int(scrollCoords(2).X / inc) * inc) * zoomFactor)
        Y = snapVertexToGrid(Y, (scrollCoords(2).Y - Int(scrollCoords(2).Y / inc) * inc) * zoomFactor)
    End If

    xCenter = (rCenter.X - scrollCoords(2).X) * zoomFactor
    yCenter = (rCenter.Y - scrollCoords(2).Y) * zoomFactor
    If xCenter = moveCoords(1).X Then
        If moveCoords(1).Y - yCenter > 0 Then
            oldAngle = PI / 2
        Else
            oldAngle = 3 * PI / 2
        End If
    ElseIf moveCoords(1).X - xCenter > 0 Then
        oldAngle = Atn((moveCoords(1).Y - yCenter) / (moveCoords(1).X - xCenter))
    ElseIf moveCoords(1).X - xCenter < 0 Then
        oldAngle = PI + Atn((moveCoords(1).Y - yCenter) / (moveCoords(1).X - xCenter))
    End If

    If xCenter = X Then
        If Y - yCenter > 0 Then
            angle = PI / 2
        Else
            angle = 3 * PI / 2
        End If
    ElseIf X - xCenter > 0 Then
        angle = Atn((Y - yCenter) / (X - xCenter))
    ElseIf X - xCenter < 0 Then
        angle = PI + Atn((Y - yCenter) / (X - xCenter))
    End If

    rDiff = angle - oldAngle

    If constrained Then
        rDiff = (Int((rDiff * 180 / PI + 7.5) / 15) * 15) / 180 * PI
    End If

    If numSelectedPolys > 0 Then
        For i = 1 To numSelectedPolys
            PolyNum = selectedPolys(i)
            For j = 1 To 3
                If vertexList(PolyNum).vertex(j) = 1 Then
                    xDiff = (PolyCoords(PolyNum).vertex(j).X - rCenter.X) * zoomFactor
                    yDiff = (PolyCoords(PolyNum).vertex(j).Y - rCenter.Y) * zoomFactor

                    R = Sqr((xDiff) ^ 2 + (yDiff) ^ 2)  ' distance of point from rotation center
                    If xDiff = 0 Then
                        If yDiff > 0 Then
                            theta = PI / 2 + rDiff
                        Else
                            theta = 3 * PI / 2 + rDiff
                        End If
                    ElseIf xDiff > 0 Then
                        theta = Atn(yDiff / xDiff) + rDiff
                    ElseIf xDiff < 0 Then
                        theta = PI + Atn(yDiff / xDiff) + rDiff
                    End If

                    Polys(PolyNum).vertex(j).X = xCenter + R * Cos(theta)
                    Polys(PolyNum).vertex(j).Y = yCenter + R * Sin(theta)
                End If
            Next
        Next
    End If

    If numSelectedScenery > 0 Then
        For i = 1 To sceneryCount
            If Scenery(i).selected = 1 Then
                xDiff = (Scenery(i).Translation.X - rCenter.X) * zoomFactor
                yDiff = (Scenery(i).Translation.Y - rCenter.Y) * zoomFactor

                R = Sqr((xDiff) ^ 2 + (yDiff) ^ 2)  ' distance of point from rotation center
                If xDiff = 0 Then
                    If yDiff > 0 Then
                        theta = PI / 2 + rDiff
                    Else
                        theta = 3 * PI / 2 + rDiff
                    End If
                ElseIf xDiff > 0 Then
                    theta = Atn(yDiff / xDiff) + rDiff
                ElseIf xDiff < 0 Then
                    theta = PI + Atn(yDiff / xDiff) + rDiff
                End If

                Scenery(i).screenTr.X = xCenter + R * Cos(theta)
                Scenery(i).screenTr.Y = yCenter + R * Sin(theta)
            End If
        Next
    End If

    If numSelWaypoints Then
        For i = 1 To waypointCount
            If Waypoints(i).selected Then
                xDiff = (Scenery(i).Translation.X - rCenter.X) * zoomFactor
                yDiff = (Scenery(i).Translation.Y - rCenter.Y) * zoomFactor

                R = Sqr((xDiff) ^ 2 + (yDiff) ^ 2)  ' distance of point from rotation center
                If xDiff = 0 Then
                    If yDiff > 0 Then
                        theta = PI / 2 + rDiff
                    Else
                        theta = 3 * PI / 2 + rDiff
                    End If
                ElseIf xDiff > 0 Then
                    theta = Atn(yDiff / xDiff) + rDiff
                ElseIf xDiff < 0 Then
                    theta = PI + Atn(yDiff / xDiff) + rDiff
                End If

                Scenery(i).screenTr.X = xCenter + R * Cos(theta)
                Scenery(i).screenTr.Y = yCenter + R * Sin(theta)
            End If
        Next
    End If

    moveCoords(2).X = X
    moveCoords(2).Y = Y

    frmInfo.txtRotate.Text = Int(rDiff / PI * 180 * 100) / 100

    prompt = True

    Render

End Sub

Private Sub PrecisionColoring(X As Single, Y As Single)

    Dim i As Integer
    Dim j As Integer
    Dim closestPoly As Single
    Dim closestVert As Single
    Dim currentDist As Long
    Dim shortestDist As Long
    Dim PolyNum As Integer
    Dim destClr As TColor
    Dim R As Integer

    R = clrRadius * zoomFactor

    shortestDist = R ^ 2 + 1
    If numSelectedPolys > 0 Then

        For i = 1 To numSelectedPolys  ' find closest
            PolyNum = selectedPolys(i)
            If pointInPoly(X, Y, i) Then
                For j = 1 To 3
                    If nearCoord(X, Polys(PolyNum).vertex(j).X, R) And nearCoord(Y, Polys(PolyNum).vertex(j).Y, R) Then
                        currentDist = (Polys(PolyNum).vertex(j).X - X) ^ 2 + (Polys(PolyNum).vertex(j).Y - Y) ^ 2
                        If currentDist < shortestDist Then
                            shortestDist = currentDist
                            closestPoly = PolyNum
                            closestVert = j
                        End If
                    End If
                Next
            End If
        Next

        If closestPoly > 0 And closestVert > 0 Then
            destClr = getRGB(Polys(closestPoly).vertex(closestVert).color)
            destClr = applyBlend(destClr)
            Polys(closestPoly).vertex(closestVert).color = ARGB(getAlpha(Polys(closestPoly).vertex(closestVert).color), RGB(destClr.blue, destClr.green, destClr.red))
            vertexList(closestPoly).color(closestVert).red = destClr.red
            vertexList(closestPoly).color(closestVert).green = destClr.green
            vertexList(closestPoly).color(closestVert).blue = destClr.blue
        End If
    Else
        For i = 1 To mPolyCount  ' find closest
            If pointInPoly(X, Y, i) Then
                For j = 1 To 3
                    If nearCoord(X, Polys(i).vertex(j).X, R) And nearCoord(Y, Polys(i).vertex(j).Y, R) Then
                        currentDist = (Polys(i).vertex(j).X - X) ^ 2 + (Polys(i).vertex(j).Y - Y) ^ 2
                        If currentDist < shortestDist Then
                            shortestDist = currentDist
                            closestPoly = i
                            closestVert = j
                        End If
                    End If
                Next
            End If
        Next

        If closestPoly > 0 And closestVert > 0 Then
            destClr = getRGB(Polys(closestPoly).vertex(closestVert).color)
            destClr = applyBlend(destClr)
            Polys(closestPoly).vertex(closestVert).color = ARGB(getAlpha(Polys(closestPoly).vertex(closestVert).color), RGB(destClr.blue, destClr.green, destClr.red))
            vertexList(closestPoly).color(closestVert).red = destClr.red
            vertexList(closestPoly).color(closestVert).green = destClr.green
            vertexList(closestPoly).color(closestVert).blue = destClr.blue
        End If
    End If

    prompt = True

    Render

End Sub

Private Sub VertexColoring(X As Single, Y As Single)

    Dim i As Integer
    Dim j As Integer
    Dim pNum As Integer
    Dim destClr As TColor
    Dim R As Integer
    Dim colored As Boolean

    R = clrRadius * zoomFactor

    If numSelectedPolys > 0 And (showPolys Or showWireframe Or showPoints) Then
        For i = 1 To numSelectedPolys
            pNum = selectedPolys(i)
            For j = 1 To 3
                If vertexList(pNum).vertex(j) = 1 Then
                    If nearCoord(X, Polys(pNum).vertex(j).X, R) And nearCoord(Y, Polys(pNum).vertex(j).Y, R) Then
                        If (Polys(pNum).vertex(j).X - X) ^ 2 + (Polys(pNum).vertex(j).Y - Y) ^ 2 <= R ^ 2 Then
                            destClr = getRGB(Polys(pNum).vertex(j).color)
                            destClr = applyBlend(destClr)
                            Polys(pNum).vertex(j).color = ARGB(getAlpha(Polys(pNum).vertex(j).color), RGB(destClr.blue, destClr.green, destClr.red))
                            vertexList(pNum).color(j).red = destClr.red
                            vertexList(pNum).color(j).green = destClr.green
                            vertexList(pNum).color(j).blue = destClr.blue
                            If lightCount > 0 Then applyLightsToVert pNum, j
                            If colorMode = 1 Then vertexList(pNum).vertex(j) = 3
                            colored = True
                        End If
                    End If
                End If
            Next
        Next
    ElseIf (showPolys Or showWireframe Or showPoints) And numSelectedScenery = 0 Then
        For i = 1 To mPolyCount
            For j = 1 To 3
                If vertexList(i).vertex(j) = 0 Then
                    If nearCoord(X, Polys(i).vertex(j).X, R) And nearCoord(Y, Polys(i).vertex(j).Y, R) Then
                        If (Polys(i).vertex(j).X - X) ^ 2 + (Polys(i).vertex(j).Y - Y) ^ 2 <= R ^ 2 Then
                            destClr = getRGB(Polys(i).vertex(j).color)
                            destClr = applyBlend(destClr)
                            Polys(i).vertex(j).color = ARGB(getAlpha(Polys(i).vertex(j).color), RGB(destClr.blue, destClr.green, destClr.red))
                            vertexList(i).color(j).red = destClr.red
                            vertexList(i).color(j).green = destClr.green
                            vertexList(i).color(j).blue = destClr.blue
                            If lightCount > 0 Then applyLightsToVert i, j
                            If colorMode = 1 Then vertexList(i).vertex(j) = 2
                            colored = True
                        End If
                    End If
                End If
            Next
        Next
    End If

    If numSelectedScenery > 0 And showScenery Then
        For i = 1 To sceneryCount
            If Scenery(i).selected = 1 Then
                If nearCoord(X, Scenery(i).screenTr.X, R) And nearCoord(Y, Scenery(i).screenTr.Y, R) Then
                    If (Scenery(i).screenTr.X - X) ^ 2 + (Scenery(i).screenTr.Y - Y) ^ 2 <= R ^ 2 Then
                        destClr = getRGB(Scenery(i).color)
                        destClr = applyBlend(destClr)
                        Scenery(i).color = ARGB(Scenery(i).alpha, RGB(destClr.blue, destClr.green, destClr.red))
                        If colorMode = 1 Then Scenery(i).selected = 3
                        colored = True
                    End If
                End If
            End If
        Next
    ElseIf showScenery And numSelectedPolys = 0 Then
        For i = 1 To sceneryCount
            If Scenery(i).selected = 0 Then
                If nearCoord(X, Scenery(i).screenTr.X, R) And nearCoord(Y, Scenery(i).screenTr.Y, R) Then
                    If (Scenery(i).screenTr.X - X) ^ 2 + (Scenery(i).screenTr.Y - Y) ^ 2 <= R ^ 2 Then
                        destClr = getRGB(Scenery(i).color)
                        destClr = applyBlend(destClr)
                        Scenery(i).color = ARGB(Scenery(i).alpha, RGB(destClr.blue, destClr.green, destClr.red))
                        If colorMode = 1 Then Scenery(i).selected = 2
                        colored = True
                    End If
                End If
            End If
        Next
    End If

    If colored Then
        prompt = True
        Render
    End If

End Sub

Private Sub EditDepthMap(X As Single, Y As Single)

    Dim i As Integer
    Dim j As Integer
    Dim pNum As Integer
    Dim destClr As TColor
    Dim R As Integer
    Dim edited As Boolean

    R = clrRadius * zoomFactor

    If numSelectedPolys > 0 And (showPolys Or showWireframe Or showPoints) Then
        For i = 1 To numSelectedPolys
            pNum = selectedPolys(i)
            For j = 1 To 3
                If vertexList(pNum).vertex(j) = 1 Then
                    If nearCoord(X, Polys(pNum).vertex(j).X, R) And nearCoord(Y, Polys(pNum).vertex(j).Y, R) Then
                        If (Polys(pNum).vertex(j).X - X) ^ 2 + (Polys(pNum).vertex(j).Y - Y) ^ 2 <= R ^ 2 Then
                            Polys(pNum).vertex(j).Z = Polys(pNum).vertex(j).Z * (1 - opacity) + gPolyClr.red * opacity
                            If colorMode = 1 Then vertexList(pNum).vertex(j) = 3
                            edited = True
                        End If
                    End If
                End If
            Next
        Next
    ElseIf (showPolys Or showWireframe Or showPoints) And numSelectedScenery = 0 Then
        For i = 1 To mPolyCount
            For j = 1 To 3
                If vertexList(i).vertex(j) = 0 Then
                    If nearCoord(X, Polys(i).vertex(j).X, R) And nearCoord(Y, Polys(i).vertex(j).Y, R) Then
                        If (Polys(i).vertex(j).X - X) ^ 2 + (Polys(i).vertex(j).Y - Y) ^ 2 <= R ^ 2 Then
                            Polys(i).vertex(j).Z = Polys(i).vertex(j).Z * (1 - opacity) + gPolyClr.red * opacity
                            If colorMode = 1 Then vertexList(i).vertex(j) = 2
                            edited = True
                        End If
                    End If
                End If
            Next
        Next
    End If

    If edited Then
        prompt = True
        Render
    End If

End Sub

Private Sub ColorPicker(X As Single, Y As Single)

    Dim i As Integer
    Dim j As Integer
    Dim shortestDist As Integer
    Dim currentDist As Integer
    Dim pNum As Integer
    Dim vNum As Integer
    Dim tempClr As TColor

    If showPolys Or showWireframe Or showPoints Then
        shortestDist = 32 ^ 2 + 1
        For i = 1 To mPolyCount
            If pointInPoly(X, Y, i) Then
                For j = 1 To 3
                    If nearCoord(X, Polys(i).vertex(j).X, 32) And nearCoord(Y, Polys(i).vertex(j).Y, 32) Then
                        currentDist = (Polys(i).vertex(j).X - X) ^ 2 + (Polys(i).vertex(j).Y - Y) ^ 2
                        If currentDist < shortestDist Then
                            shortestDist = currentDist
                            pNum = i
                            vNum = j
                        End If
                    End If
                Next
            End If
        Next
    End If

    If vNum > 0 Then  ' poly color absorbed
        tempClr = vertexList(pNum).color(vNum)
        If tempClr.red = gPolyClr.red And tempClr.green = gPolyClr.green And tempClr.blue = gPolyClr.blue Then
        ElseIf frmPalette.Enabled = False Then  ' non modal
            frmColor.InitColor tempClr.red, tempClr.green, tempClr.blue
        Else
            gPolyClr = tempClr
            Scenery(0).color = ARGB(Scenery(0).alpha, Polys(pNum).vertex(vNum).color)
            frmPalette.setValues gPolyClr.red, gPolyClr.green, gPolyClr.blue
            frmPalette.checkPalette gPolyClr.red, gPolyClr.green, gPolyClr.blue
        End If
    ElseIf showScenery Then  ' no poly clrs absorbed, do scenery
        For i = 1 To sceneryCount
            If PointInProp(X, Y, i) And vNum = 0 Then
                vNum = i
            End If
        Next
        If vNum > 0 Then
            tempClr = getRGB(Scenery(vNum).color)
            If tempClr.red = gPolyClr.red And tempClr.green = gPolyClr.green And tempClr.blue = gPolyClr.blue Then

            ElseIf frmPalette.Enabled = False Then  ' non modal
                frmColor.InitColor tempClr.red, tempClr.green, tempClr.blue
            Else
                gPolyClr = tempClr
                Scenery(0).color = ARGB(Scenery(0).alpha, Scenery(vNum).color)
                frmPalette.setValues gPolyClr.red, gPolyClr.green, gPolyClr.blue
                frmPalette.checkPalette gPolyClr.red, gPolyClr.green, gPolyClr.blue
            End If
        End If
    End If

End Sub

Private Sub depthPicker(X As Single, Y As Single)

    Dim i As Integer
    Dim j As Integer
    Dim shortestDist As Integer
    Dim currentDist As Integer
    Dim pNum As Integer
    Dim vNum As Integer

    If showPolys Or showWireframe Or showPoints Then
        shortestDist = 32 ^ 2 + 1
        For i = 1 To mPolyCount
            If pointInPoly(X, Y, i) Then
                For j = 1 To 3
                    If nearCoord(X, Polys(i).vertex(j).X, 32) And nearCoord(Y, Polys(i).vertex(j).Y, 32) Then
                        currentDist = (Polys(i).vertex(j).X - X) ^ 2 + (Polys(i).vertex(j).Y - Y) ^ 2
                        If currentDist < shortestDist Then
                            shortestDist = currentDist
                            pNum = i
                            vNum = j
                        End If
                    End If
                Next
            End If
        Next
    End If

    If vNum > 0 Then  ' poly color absorbed
        If Polys(pNum).vertex(vNum).Z >= 0 And Polys(pNum).vertex(vNum).Z <= 255 Then
            gPolyClr.red = Polys(pNum).vertex(vNum).Z
        ElseIf Polys(pNum).vertex(vNum).Z < 0 Then
            gPolyClr.red = 0
        ElseIf Polys(pNum).vertex(vNum).Z > 255 Then
            gPolyClr.red = 255
        End If
        gPolyClr.green = gPolyClr.red
        gPolyClr.blue = gPolyClr.red
        Scenery(0).color = ARGB(Scenery(0).alpha, Polys(pNum).vertex(vNum).color)
        frmPalette.setValues gPolyClr.red, gPolyClr.green, gPolyClr.blue
        frmPalette.checkPalette gPolyClr.red, gPolyClr.green, gPolyClr.blue
    End If

End Sub

Private Sub lightPicker(X As Single, Y As Single)

    Dim i As Integer
    Dim j As Integer
    Dim shortestDist As Integer
    Dim currentDist As Integer
    Dim pNum As Integer
    Dim vNum As Integer
    Dim tempClr As TColor

    If showPolys Or showWireframe Or showPoints Then
        shortestDist = 32 ^ 2 + 1
        For i = 1 To mPolyCount
            If pointInPoly(X, Y, i) Then
                For j = 1 To 3
                    If nearCoord(X, Polys(i).vertex(j).X, 32) And nearCoord(Y, Polys(i).vertex(j).Y, 32) Then
                        currentDist = (Polys(i).vertex(j).X - X) ^ 2 + (Polys(i).vertex(j).Y - Y) ^ 2
                        If currentDist < shortestDist Then
                            shortestDist = currentDist
                            pNum = i
                            vNum = j
                        End If
                    End If
                Next
            End If
        Next
    End If

    If vNum > 0 Then  ' poly color absorbed
        tempClr = getRGB(Polys(pNum).vertex(vNum).color)
        If tempClr.red = gPolyClr.red And tempClr.green = gPolyClr.green And tempClr.blue = gPolyClr.blue Then
        ElseIf frmPalette.Enabled = False Then  ' non modal
            frmColor.InitColor tempClr.red, tempClr.green, tempClr.blue
        Else
            gPolyClr = tempClr
            Scenery(0).color = ARGB(Scenery(0).alpha, Polys(pNum).vertex(vNum).color)
            frmPalette.setValues gPolyClr.red, gPolyClr.green, gPolyClr.blue
            frmPalette.checkPalette gPolyClr.red, gPolyClr.green, gPolyClr.blue
        End If
    End If

End Sub

Private Sub StretchingTexture(X As Single, Y As Single)

    Dim i As Integer
    Dim j As Integer
    Dim PolyNum As Integer

    If numSelectedPolys > 0 Then
        For i = 1 To numSelectedPolys
            PolyNum = selectedPolys(i)
            For j = 1 To 3
                If vertexList(PolyNum).vertex(j) = 1 Then
                    Polys(PolyNum).vertex(j).tu = (Polys(PolyNum).vertex(j).tu - (X - moveCoords(1).X) / zoomFactor / xTexture)
                    Polys(PolyNum).vertex(j).tv = (Polys(PolyNum).vertex(j).tv - (Y - moveCoords(1).Y) / zoomFactor / yTexture)
                End If
            Next
        Next
        moveCoords(1).X = X
        moveCoords(1).Y = Y
        prompt = True
    End If

    getInfo

    Render

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim i As Integer
    Dim j As Integer

    If Button = 4 Then
        SetCursor currentFunction + 1
    End If

    If Button <> 1 Then
        Exit Sub
    End If

    If spaceDown Then
        Render
    ElseIf currentFunction = TOOL_MOVE And toolAction Then  ' snap selected vertex
        If Shift = KEY_SHIFT Then  ' constrained, don't snap
        Else
            snapSelected X, Y
            If noneSelected Then
                mnuDeselect_Click
                noneSelected = False
            End If
        End If

        If lightCount > 0 And showLights Then
            If numSelLights > 0 Then
                applyLights
                Render
            ElseIf numSelectedPolys > 0 Then
                applyLights True
                Render
            End If
        End If

        SaveUndo
    ElseIf currentFunction = TOOL_SCALE And toolAction Then  ' apply scaling
        If Shift = KEY_CTRL Then
            ApplyTransform False
        ElseIf Shift = KEY_SHIFT + KEY_CTRL Then  ' constrained scaling
            ApplyTransform False
        End If
    ElseIf currentFunction = TOOL_ROTATE And toolAction Then  ' apply rotation
        If Shift = KEY_ALT Then
            ApplyTransform True
        ElseIf Shift = KEY_SHIFT + KEY_ALT Then  ' constrained rotation
            ApplyTransform True
        End If
    ElseIf (currentFunction = TOOL_CREATE Or currentFunction = TOOL_QUAD) And toolAction Then  ' create polys
        If Shift = KEY_SHIFT And numVerts > 0 Then
            X = Polys(mPolyCount + 1).vertex(numVerts + 1).X
            Y = Polys(mPolyCount + 1).vertex(numVerts + 1).Y
        End If
        CreatePolys X, Y
    ElseIf currentFunction = TOOL_SCENERY And toolAction Then  ' create scenery
        CreateScenery X, Y
    ElseIf currentFunction = TOOL_VSELECT Or currentFunction = TOOL_VSELADD Or currentFunction = TOOL_VSELSUB Then  ' select vertices
        If toolAction Then
            selectedCoords(2).X = X
            selectedCoords(2).Y = Y
            eraseLines = False
            noRedraw = True
            If selectedCoords(2).X = selectedCoords(1).X And selectedCoords(2).Y = selectedCoords(1).Y Then
                regionSelection X, Y
            Else
                VertexSelection X, Y
            End If
            noRedraw = False
            selectedCoords(1).X = X
            selectedCoords(1).Y = Y
            Render
            If numSelectedPolys = 0 And numSelectedScenery = 0 And numSelLights = 0 And numSelSpawns = 0 And numSelWaypoints = 0 And numSelColliders = 1 Then
                For i = 1 To colliderCount
                    If Colliders(i).active = 1 Then
                        frmPalette.txtRadius.Text = LTrim$(Str$(Colliders(i).radius))
                        setRadius CInt(Colliders(i).radius)
                    End If
                Next
            End If
        End If
    ElseIf currentFunction = TOOL_PSELECT And toolAction Then  ' poly selection
    ElseIf currentFunction = TOOL_VCOLOR And toolAction Then  ' vertex coloring
        toolAction = False
        If colorMode = 1 Then
            For i = 1 To mPolyCount
                For j = 1 To 3
                    If vertexList(i).vertex(j) > 1 Then
                        vertexList(i).vertex(j) = vertexList(i).vertex(j) - 2
                    End If
                Next
            Next
            For i = 1 To sceneryCount
                If Scenery(i).selected > 1 Then
                    Scenery(i).selected = Scenery(i).selected - 2
                End If
            Next
        End If
        SaveUndo
    ElseIf currentFunction = TOOL_PCOLOR And toolAction Then  ' poly color
    ElseIf currentFunction = TOOL_TEXTURE And toolAction Then  ' texture
        SaveUndo
    ElseIf currentFunction = TOOL_OBJECTS And toolAction Then  ' objects
        SaveUndo
    ElseIf currentFunction = TOOL_WAYPOINT And toolAction Then  ' waypoints
        SaveUndo
    ElseIf currentFunction = TOOL_CONNECT And toolAction Then
        CreateConnection X, Y
    ElseIf currentFunction = TOOL_SKETCH Then
        If Shift = 0 And toolAction Then  ' freeform
            endSketch X, Y
            toolAction = False
        ElseIf Shift = 1 Then  ' lines
            If toolAction Then
                lineSketch X, Y
            Else
                toolAction = True
            End If
            sketch(0).vertex(1).X = X / zoomFactor + scrollCoords(2).X
            sketch(0).vertex(1).Y = Y / zoomFactor + scrollCoords(2).Y
            sketch(0).vertex(2).X = X / zoomFactor + scrollCoords(2).X
            sketch(0).vertex(2).Y = Y / zoomFactor + scrollCoords(2).Y
        End If

        deleteSmallLines
    ElseIf currentFunction = TOOL_ERASER Then
        toolAction = False
    ElseIf currentFunction = TOOL_DEPTHMAP Then
        toolAction = False
        If colorMode = 1 Then
            For i = 1 To mPolyCount
                For j = 1 To 3
                    If vertexList(i).vertex(j) > 1 Then
                        vertexList(i).vertex(j) = vertexList(i).vertex(j) - 2
                    End If
                Next
            Next
        End If
        SaveUndo
    End If

    If currentFunction <> TOOL_CREATE And currentFunction <> TOOL_QUAD And currentFunction <> TOOL_SKETCH And currentFunction <> TOOL_SCENERY Then
        If numVerts = 0 Then
            toolAction = False
        End If
    End If

    If noneSelected Then
        mnuDeselect_Click
        noneSelected = False
    End If

    If numSelWaypoints = 0 And frmWaypoints.Visible = True Then
        frmWaypoints.ClearWaypt
    End If

    selectedCoords(1).X = 0
    selectedCoords(1).Y = 0
    selectedCoords(2).X = 0
    selectedCoords(2).Y = 0

End Sub

Private Sub CreateConnection(X As Single, Y As Single)

    Dim i As Integer
    Dim notSel As Integer
    Dim currentDist As Long
    Dim shortestDist As Long
    Dim xVal As Single
    Dim yVal As Single

    xVal = X / zoomFactor + scrollCoords(2).X
    yVal = Y / zoomFactor + scrollCoords(2).Y

    notSel = 0
    shortestDist = (8 ^ 2 + 1)
    For i = 1 To waypointCount
        If (Waypoints(i).pathNum = 1 And frmWaypoints.showPaths <> 2) Or (Waypoints(i).pathNum = 2 And frmWaypoints.showPaths <> 1) Then
            If nearCoord(xVal, Waypoints(i).X, 8 / zoomFactor) And nearCoord(yVal, Waypoints(i).Y, 8 / zoomFactor) Then
                currentDist = (Waypoints(i).X - xVal) ^ 2 + (Waypoints(i).Y - yVal) ^ 2
                If currentDist < shortestDist Then
                    shortestDist = currentDist
                    notSel = i
                End If
            End If
        End If
    Next
    If notSel > 0 And currentWaypoint <> notSel Then
        If currentWaypoint > 0 Then  ' connecting waypoints
            If selectionChanged Then
                SaveUndo
                selectionChanged = False
            End If
            conCount = conCount + 1
            ReDim Preserve Connections(conCount)
            Connections(conCount).point1 = currentWaypoint
            Connections(conCount).point2 = notSel
            Waypoints(currentWaypoint).numConnections = Waypoints(currentWaypoint).numConnections + 1
            SaveUndo
        End If
        currentWaypoint = notSel
    ElseIf notSel > 0 Then
        currentWaypoint = notSel
    Else
        currentWaypoint = 0
        For i = 1 To waypointCount
            Waypoints(i).selected = False
        Next
        numSelWaypoints = 0
    End If

    getInfo
    Render

End Sub

Private Sub CreatePolys(X As Single, Y As Single)

    Dim i As Integer
    Dim j As Integer
    Dim xVal As Single
    Dim yVal As Single
    Dim shortestDist As Integer
    Dim currentDist As Long
    Dim temp As D3DVECTOR2
    Dim tempVertex As TCustomVertex

    If numVerts = 0 Then
        ReDim Preserve Polys(mPolyCount + 1)
        ReDim Preserve PolyCoords(mPolyCount + 1)
        ReDim Preserve vertexList(mPolyCount + 1)
        vertexList(mPolyCount + 1).polyType = polyType
    End If
    numVerts = numVerts + 1

    xVal = X
    yVal = Y

    If snapToGrid And showGrid Then
        xVal = snapVertexToGrid(xVal, (scrollCoords(2).X - Int(scrollCoords(2).X / inc) * inc) * zoomFactor)
        yVal = snapVertexToGrid(yVal, (scrollCoords(2).Y - Int(scrollCoords(2).Y / inc) * inc) * zoomFactor)
        PolyCoords(mPolyCount + 1).vertex(numVerts).X = Int(xVal / zoomFactor + scrollCoords(2).X + 0.5)
        PolyCoords(mPolyCount + 1).vertex(numVerts).Y = Int(yVal / zoomFactor + scrollCoords(2).Y + 0.5)
    ElseIf ohSnap Then  ' snap
        shortestDist = snapRadius ^ 2 + 1
        For i = 1 To mPolyCount
            For j = 1 To 3
                If nearCoord(xVal, Polys(i).vertex(j).X, shortestDist) And nearCoord(yVal, Polys(i).vertex(j).Y, shortestDist) Then
                    currentDist = ((Polys(i).vertex(j).X - xVal) ^ 2 + (Polys(i).vertex(j).Y - yVal) ^ 2)
                    If currentDist < shortestDist Then
                        shortestDist = currentDist
                        xVal = Polys(i).vertex(j).X
                        yVal = Polys(i).vertex(j).Y
                        PolyCoords(mPolyCount + 1).vertex(numVerts).X = PolyCoords(i).vertex(j).X
                        PolyCoords(mPolyCount + 1).vertex(numVerts).Y = PolyCoords(i).vertex(j).Y
                    End If
                End If
            Next
        Next
    End If

    If (xVal = X And yVal = Y) Or (Not ohSnap And Not snapToGrid) Then  ' no snapping occured
        PolyCoords(mPolyCount + 1).vertex(numVerts).X = Int(xVal / zoomFactor + scrollCoords(2).X + 0.5)
        PolyCoords(mPolyCount + 1).vertex(numVerts).Y = Int(yVal / zoomFactor + scrollCoords(2).Y + 0.5)
    End If

    Polys(mPolyCount + 1).vertex(numVerts) = CreateCustomVertex(xVal, yVal, _
            0, 1, ARGB(255 * opacity, RGB(gPolyClr.blue, gPolyClr.green, gPolyClr.red)), _
            (xVal / zoomFactor + scrollCoords(2).X) / xTexture, (yVal / zoomFactor + scrollCoords(2).Y) / yTexture)
    vertexList(mPolyCount + 1).color(numVerts).red = gPolyClr.red
    vertexList(mPolyCount + 1).color(numVerts).green = gPolyClr.green
    vertexList(mPolyCount + 1).color(numVerts).blue = gPolyClr.blue

    If mnuQuad.Checked And mnuCustomX.Checked Then
        If creatingQuad Then
            Polys(mPolyCount + 1).vertex(numVerts).tu = (frmTexture.x1tex * 2 + 0.5) / xTexture
        Else
            If numVerts = 2 Or numVerts = 3 Then
                Polys(mPolyCount + 1).vertex(numVerts).tu = (frmTexture.x2tex * 2 - 0.5) / xTexture
            Else
                Polys(mPolyCount + 1).vertex(numVerts).tu = (frmTexture.x1tex * 2 + 0.5) / xTexture
            End If
        End If
    End If

    If mnuQuad.Checked And mnuCustomY.Checked Then
        If creatingQuad Then
            Polys(mPolyCount + 1).vertex(numVerts).tv = (frmTexture.y2tex * 2 - 0.5) / yTexture
        Else
            If numVerts > 2 Then
                Polys(mPolyCount + 1).vertex(numVerts).tv = (frmTexture.y2tex * 2 - 0.5) / yTexture
            Else
                Polys(mPolyCount + 1).vertex(numVerts).tv = (frmTexture.y1tex * 2 + 0.5) / yTexture
            End If
        End If
    End If


    If numVerts = 1 Then
        Polys(mPolyCount + 1).vertex(2) = Polys(mPolyCount + 1).vertex(1)
        Polys(mPolyCount + 1).vertex(numVerts + 1).X = X
        Polys(mPolyCount + 1).vertex(numVerts + 1).Y = Y
        Polys(mPolyCount + 1).vertex(numVerts + 2) = Polys(mPolyCount + 1).vertex(1)
        PolyCoords(mPolyCount + 1).vertex(numVerts + 2) = PolyCoords(mPolyCount + 1).vertex(1)
    ElseIf numVerts = 3 Then
        numVerts = 0
        mPolyCount = mPolyCount + 1
        If Not isCW(mPolyCount) Then  ' switch to make cw
            temp = PolyCoords(mPolyCount).vertex(3)
            PolyCoords(mPolyCount).vertex(3) = PolyCoords(mPolyCount).vertex(2)
            PolyCoords(mPolyCount).vertex(2) = temp

            tempVertex = Polys(mPolyCount).vertex(3)
            Polys(mPolyCount).vertex(3) = Polys(mPolyCount).vertex(2)
            Polys(mPolyCount).vertex(2) = tempVertex
        End If
        toolAction = False
        frmInfo.lblCount(0).Caption = mPolyCount
        frmInfo.lblCount(6).Caption = getMapDimensions

        applyLightsToVert CInt(mPolyCount), 1
        applyLightsToVert CInt(mPolyCount), 2
        applyLightsToVert CInt(mPolyCount), 3

        Polys(mPolyCount).Perp.vertex(1).Z = 2
        Polys(mPolyCount).Perp.vertex(2).Z = 2
        Polys(mPolyCount).Perp.vertex(3).Z = 2

        SaveUndo
        If mnuQuad.Checked And Not creatingQuad Then
            ReDim Preserve Polys(mPolyCount + 1)
            ReDim Preserve PolyCoords(mPolyCount + 1)
            ReDim Preserve vertexList(mPolyCount + 1)
            vertexList(mPolyCount + 1).polyType = polyType
            Polys(mPolyCount + 1).vertex(1) = Polys(mPolyCount).vertex(1)
            Polys(mPolyCount + 1).vertex(2) = Polys(mPolyCount).vertex(3)
            PolyCoords(mPolyCount + 1).vertex(1) = PolyCoords(mPolyCount).vertex(1)
            PolyCoords(mPolyCount + 1).vertex(2) = PolyCoords(mPolyCount).vertex(3)
            vertexList(mPolyCount + 1).color(1) = vertexList(mPolyCount).color(1)
            vertexList(mPolyCount + 1).color(2) = vertexList(mPolyCount).color(3)
            numVerts = 2
            Polys(mPolyCount + 1).vertex(3) = Polys(mPolyCount).vertex(3)
            PolyCoords(mPolyCount + 1).vertex(3) = PolyCoords(mPolyCount).vertex(3)
            toolAction = True
            creatingQuad = True
        ElseIf creatingQuad Then
            creatingQuad = False
        End If
        prompt = True
    End If

    Render

End Sub

Private Sub startSketch(X As Single, Y As Single)

    On Error GoTo ErrorHandler

    showSketch = True
    frmDisplay.setLayer 10, showSketch

    sketchLines = sketchLines + 1
    ReDim Preserve sketch(sketchLines)

    sketch(sketchLines).vertex(1).X = X / zoomFactor + scrollCoords(2).X
    sketch(sketchLines).vertex(1).Y = Y / zoomFactor + scrollCoords(2).Y
    sketch(sketchLines).vertex(2).X = sketch(sketchLines).vertex(1).X
    sketch(sketchLines).vertex(2).Y = sketch(sketchLines).vertex(1).Y

    sketch(sketchLines).vertex(1).Z = 1
    sketch(sketchLines).vertex(2).Z = 1

    Render

    Exit Sub

ErrorHandler:

    MsgBox "Error starting sketch" & vbNewLine & Error$

End Sub

Private Sub lineSketch(X As Single, Y As Single)

    On Error GoTo ErrorHandler

    sketchLines = sketchLines + 1
    ReDim Preserve sketch(sketchLines)

    sketch(sketchLines).vertex(1).X = sketch(0).vertex(1).X
    sketch(sketchLines).vertex(1).Y = sketch(0).vertex(1).Y
    sketch(sketchLines).vertex(2).X = Int(X / zoomFactor + scrollCoords(2).X + 0.5)
    sketch(sketchLines).vertex(2).Y = Int(Y / zoomFactor + scrollCoords(2).Y + 0.5)

    sketch(sketchLines).vertex(1).Z = 1
    sketch(sketchLines).vertex(2).Z = 1

    Exit Sub

ErrorHandler:

    MsgBox "Error sketching line" & vbNewLine & Error$

End Sub

Private Sub linkSketch(X As Single, Y As Single)

    Dim xVal As Single
    Dim yVal As Single

    On Error GoTo ErrorHandler

    xVal = X / zoomFactor + scrollCoords(2).X
    yVal = Y / zoomFactor + scrollCoords(2).Y

    If (xVal - sketch(sketchLines).vertex(1).X) ^ 2 + (yVal - sketch(sketchLines).vertex(1).Y) ^ 2 > 16 ^ 2 Then
        sketch(sketchLines).vertex(2).X = X / zoomFactor + scrollCoords(2).X
        sketch(sketchLines).vertex(2).Y = Y / zoomFactor + scrollCoords(2).Y

        sketchLines = sketchLines + 1
        ReDim Preserve sketch(sketchLines)

        sketch(sketchLines).vertex(1).X = X / zoomFactor + scrollCoords(2).X
        sketch(sketchLines).vertex(1).Y = Y / zoomFactor + scrollCoords(2).Y
        sketch(sketchLines).vertex(2).X = X / zoomFactor + scrollCoords(2).X
        sketch(sketchLines).vertex(2).Y = Y / zoomFactor + scrollCoords(2).Y

        sketch(sketchLines).vertex(1).Z = 1
        sketch(sketchLines).vertex(2).Z = 1
    End If

    Exit Sub

ErrorHandler:

    MsgBox "Error linking sketch" & vbNewLine & Error$

End Sub

Private Sub endSketch(X As Single, Y As Single)

    sketch(sketchLines).vertex(2).X = X / zoomFactor + scrollCoords(2).X
    sketch(sketchLines).vertex(2).Y = Y / zoomFactor + scrollCoords(2).Y

    Render

    Exit Sub

ErrorHandler:
    MsgBox "Error ending sketch" & vbNewLine & Error$

End Sub

Private Sub CreateScenery(X As Single, Y As Single)

    Dim xVal As Integer
    Dim yVal As Integer
    Dim i As Integer

    On Error GoTo ErrorHandler

    If numCorners = 0 Then
        Scenery(0).screenTr.X = X
        Scenery(0).screenTr.Y = Y
    End If

    numCorners = numCorners + 1

    xVal = (Scenery(0).screenTr.X)
    yVal = (Scenery(0).screenTr.Y)

    If snapToGrid And showGrid Then
        xVal = snapVertexToGrid(xVal, (scrollCoords(2).X - Int(scrollCoords(2).X / inc) * inc) * zoomFactor)
        yVal = snapVertexToGrid(yVal, (scrollCoords(2).Y - Int(scrollCoords(2).Y / inc) * inc) * zoomFactor)

        If numCorners = 1 Then
            Scenery(0).screenTr.X = xVal
            Scenery(0).screenTr.Y = yVal
        ElseIf numCorners = 2 Then
        End If
    End If

    If numCorners = 1 And Not frmScenery.rotateScenery Then numCorners = numCorners + 1
    If numCorners = 2 And Not frmScenery.scaleScenery Then numCorners = numCorners + 1

    If numCorners = 3 Then
        sceneryCount = sceneryCount + 1
        ReDim Preserve Scenery(sceneryCount)

        Scenery(sceneryCount) = Scenery(0)
        Scenery(sceneryCount).Translation.X = Int(Scenery(0).screenTr.X / zoomFactor + scrollCoords(2).X + 0.5)
        Scenery(sceneryCount).Translation.Y = Int(Scenery(0).screenTr.Y / zoomFactor + scrollCoords(2).Y + 0.5)

        If Scenery(0).Style = 0 Then  ' create scenery texture
            CreateSceneryTexture currentScenery
            Scenery(0).Style = sceneryElements
            Scenery(sceneryCount).Style = sceneryElements
            frmScenery.notClicked = True
        End If

        setCurrentScenery
        frmInfo.lblCount(1).Caption = sceneryCount & "/500 (" & sceneryElements & ")"
        numCorners = 0

        prompt = True
        SaveUndo
    End If

    Exit Sub

ErrorHandler:

    MsgBox "Error creating scenery" & vbNewLine & Error$

End Sub

Private Sub snapSelected(X As Single, Y As Single)

    Dim i As Integer
    Dim j As Integer
    Dim PolyNum As Integer
    Dim xVal As Single
    Dim yVal As Single
    Dim temp As D3DVECTOR2
    Dim tempVertex As TCustomVertex
    Dim shortestDist As Single
    Dim currentDist As Single
    Dim xDiff As Single
    Dim yDiff As Single

    ' make sure polys are cw
    For i = 1 To numSelectedPolys
        If Not isCW(selectedPolys(i)) Then  ' switch to make cw
            temp = PolyCoords(selectedPolys(i)).vertex(3)
            PolyCoords(selectedPolys(i)).vertex(3) = PolyCoords(selectedPolys(i)).vertex(2)
            PolyCoords(selectedPolys(i)).vertex(2) = temp

            tempVertex = Polys(selectedPolys(i)).vertex(3)
            Polys(selectedPolys(i)).vertex(3) = Polys(selectedPolys(i)).vertex(2)
            Polys(selectedPolys(i)).vertex(2) = tempVertex

            PolyNum = vertexList(selectedPolys(i)).vertex(3)
            vertexList(selectedPolys(i)).vertex(3) = vertexList(selectedPolys(i)).vertex(2)
            vertexList(selectedPolys(i)).vertex(2) = PolyNum

            PolyNum = 0
        End If
    Next

    ' if grid is on, snap to grid
    ' else, if vert snapping is on then snap to verts

    ' find which vertex of poly is selected
    PolyNum = 0
    If numSelectedPolys > 0 Then
        For j = 1 To 3
            If vertexList(selectedPolys(1)).vertex(j) = 1 Then  ' which vertex in poly is selected
                If PolyNum > 0 And Not (snapToGrid And showGrid) Then  ' if more than one vertex in poly selected
                    Render
                    Exit Sub
                Else
                    PolyNum = j
                End If
            End If
        Next

        xVal = (Polys(selectedPolys(1)).vertex(PolyNum).X)
        yVal = (Polys(selectedPolys(1)).vertex(PolyNum).Y)
    Else
        For i = 1 To sceneryCount
            If Scenery(i).selected = 1 Then
                xVal = Scenery(i).screenTr.X
                yVal = Scenery(i).screenTr.Y
                Exit For
            End If
        Next
    End If

    If snapToGrid And showGrid Then
        xDiff = xVal - snapVertexToGrid(xVal, (scrollCoords(2).X - Int(scrollCoords(2).X / inc) * inc) * zoomFactor)
        yDiff = yVal - snapVertexToGrid(yVal, (scrollCoords(2).Y - Int(scrollCoords(2).Y / inc) * inc) * zoomFactor)

        If numSelectedPolys > 0 Then
            For i = 1 To numSelectedPolys
                PolyNum = selectedPolys(i)
                For j = 1 To 3
                    If vertexList(PolyNum).vertex(j) = 1 Then  ' if selected
                        Polys(PolyNum).vertex(j).X = Polys(PolyNum).vertex(j).X - xDiff
                        Polys(PolyNum).vertex(j).Y = Polys(PolyNum).vertex(j).Y - yDiff

                        PolyCoords(PolyNum).vertex(j).X = Int(Polys(PolyNum).vertex(j).X / zoomFactor + scrollCoords(2).X + 0.5)
                        PolyCoords(PolyNum).vertex(j).Y = Int(Polys(PolyNum).vertex(j).Y / zoomFactor + scrollCoords(2).Y + 0.5)

                        If fixedTexture Then
                            Polys(PolyNum).vertex(j).tu = PolyCoords(PolyNum).vertex(j).X / xTexture
                            Polys(PolyNum).vertex(j).tv = PolyCoords(PolyNum).vertex(j).Y / yTexture
                        End If
                    End If
                Next
            Next
        End If

        If numSelectedScenery > 0 Then
            For i = 1 To sceneryCount
                If Scenery(i).selected = 1 Then
                    Scenery(i).screenTr.X = Scenery(i).screenTr.X - xDiff
                    Scenery(i).screenTr.Y = Scenery(i).screenTr.Y - yDiff

                    Scenery(i).Translation.X = Int(Scenery(i).screenTr.X / zoomFactor + scrollCoords(2).X + 0.5)
                    Scenery(i).Translation.Y = Int(Scenery(i).screenTr.Y / zoomFactor + scrollCoords(2).Y + 0.5)
                End If
            Next
        End If

        If numSelSpawns > 0 Then
            For i = 1 To spawnPoints
                If Spawns(i).active = 1 Then
                    Spawns(i).X = Int((Spawns(i).X + inc / 2) / inc) * inc
                    Spawns(i).Y = Int((Spawns(i).Y + inc / 2) / inc) * inc
                End If
            Next
        End If

        If numSelColliders > 0 Then
            For i = 1 To colliderCount
                If Colliders(i).active = 1 Then
                    Colliders(i).X = Int((Colliders(i).X + inc / 2) / inc) * inc
                    Colliders(i).Y = Int((Colliders(i).Y + inc / 2) / inc) * inc
                End If
            Next
        End If

        If numSelLights > 0 Then
            For i = 1 To lightCount
                If Lights(i).selected Then
                    Lights(i).X = Int((Lights(i).X + inc / 2) / inc) * inc
                    Lights(i).Y = Int((Lights(i).Y + inc / 2) / inc) * inc
                End If
            Next
        End If

        rCenter.X = rCenter.X - xDiff / zoomFactor
        rCenter.Y = rCenter.Y - yDiff / zoomFactor
        For i = 0 To 3
            selRect(i).X = selRect(i).X - xDiff / zoomFactor
            selRect(i).Y = selRect(i).Y - yDiff / zoomFactor
        Next
    ElseIf ohSnap And numSelectedPolys > 0 Then
        ' if vertices with different coords are selected then exit sub
        If numSelectedPolys > 1 Then  ' check if any different coords
            For i = 2 To numSelectedPolys
                For j = 1 To 3
                    If vertexList(selectedPolys(i)).vertex(j) = 1 Then  ' if selected and has same coords
                        If Polys(selectedPolys(i)).vertex(j).X <> xVal Or Polys(selectedPolys(i)).vertex(j).Y <> yVal Then
                            Render
                            Exit Sub
                        End If
                    End If
                Next
            Next
        End If

        ' snap
        shortestDist = snapRadius ^ 2 + 1
        For i = 1 To mPolyCount
            For j = 1 To 3
                If nearCoord(xVal, Polys(i).vertex(j).X, shortestDist) And nearCoord(yVal, Polys(i).vertex(j).Y, shortestDist) Then
                    currentDist = (Polys(i).vertex(j).X - xVal) ^ 2 + (Polys(i).vertex(j).Y - yVal) ^ 2
                    If currentDist <= shortestDist And vertexList(i).vertex(j) = 0 Then
                        shortestDist = currentDist
                        xDiff = xVal - Polys(i).vertex(j).X
                        yDiff = yVal - Polys(i).vertex(j).Y
                        xVal = Polys(i).vertex(j).X
                        yVal = Polys(i).vertex(j).Y
                    End If
                End If
            Next
        Next

        ' if snapping occured
        If xVal <> (Polys(selectedPolys(1)).vertex(PolyNum).X) Or yVal <> (Polys(selectedPolys(1)).vertex(PolyNum).Y) Then
            For i = 1 To numSelectedPolys
                For j = 1 To 3
                    If vertexList(selectedPolys(i)).vertex(j) = 1 Then
                        Polys(selectedPolys(i)).vertex(j).X = xVal
                        Polys(selectedPolys(i)).vertex(j).Y = yVal

                        PolyCoords(selectedPolys(i)).vertex(j).X = (xVal / zoomFactor + scrollCoords(2).X)
                        PolyCoords(selectedPolys(i)).vertex(j).Y = (yVal / zoomFactor + scrollCoords(2).Y)

                    End If
                Next
            Next
            rCenter.X = rCenter.X - xDiff / zoomFactor
            rCenter.Y = rCenter.Y - yDiff / zoomFactor
            For i = 0 To 3
                selRect(i).X = selRect(i).X - xDiff / zoomFactor
                selRect(i).Y = selRect(i).Y - yDiff / zoomFactor
            Next
        End If

        PolyNum = 0
    End If

    getInfo

    Render

End Sub

Private Sub regionSelection(X As Single, Y As Single)

    Dim i As Integer
    Dim j As Integer
    Dim xVal As Single
    Dim yVal As Single
    Dim isSelected As Boolean

    xVal = X / zoomFactor + scrollCoords(2).X
    yVal = Y / zoomFactor + scrollCoords(2).Y

    If currentFunction = TOOL_VSELECT Then
        numSelectedPolys = 0
        ReDim selectedPolys(numSelectedPolys)
        numSelectedScenery = 0
        numSelSpawns = 0
        numSelColliders = 0
        numSelWaypoints = 0
        numSelLights = 0
        For i = 1 To sceneryCount
            Scenery(i).selected = 0
        Next
    End If

    If showPolys Or showWireframe Or showPoints Then
        isSelected = RegionSelPolys(X, Y)
    ElseIf currentFunction = TOOL_VSELECT Then
        For i = 1 To mPolyCount
            vertexList(i).vertex(1) = 0
            vertexList(i).vertex(2) = 0
            vertexList(i).vertex(3) = 0
        Next
    End If
    If showObjects Then
        isSelected = RegionSelObjects(xVal, yVal, isSelected)
    ElseIf currentFunction = TOOL_VSELECT Then
        For i = 1 To spawnPoints
            Spawns(i).active = 0
        Next
        For i = 1 To colliderCount
            Colliders(i).active = 0
        Next
    End If
    If showWaypoints Then
        isSelected = RegionSelWaypoints(xVal, yVal, isSelected)
    Else
        For i = 1 To waypointCount
            Waypoints(i).selected = False
        Next
    End If
    If showLights Then
        isSelected = regionSelLights(xVal, yVal, isSelected)
    ElseIf currentFunction = TOOL_VSELECT Then
        For i = 1 To lightCount
            Lights(i).selected = 0
        Next
    End If

    currentWaypoint = 0

    selectedCoords(1).X = 0
    selectedCoords(1).Y = 0
    selectedCoords(2).X = 0
    selectedCoords(2).Y = 0

    getRCenter
    getInfo
    selectionChanged = True
    Render

End Sub

Private Function RegionSelPolys(X As Single, Y As Single) As Boolean

    Dim i As Integer
    Dim j As Integer
    Dim currentDist As Long
    Dim shortestDist As Long
    Dim pIndex As Integer
    Dim vIndex As Integer
    Dim selVerts As Byte
    Dim selected As Byte
    Dim xVal As Single
    Dim yVal As Single

    xVal = X / zoomFactor + scrollCoords(2).X
    yVal = Y / zoomFactor + scrollCoords(2).Y

    For i = 1 To mPolyCount
        If currentFunction = TOOL_VSELECT Then
            vertexList(i).vertex(1) = 0
            vertexList(i).vertex(2) = 0
            vertexList(i).vertex(3) = 0
        End If

        If (pointInPoly(X, Y, i)) Then
            shortestDist = 64 ^ 2 + 1
            For j = 1 To 3
                currentDist = (PolyCoords(i).vertex(j).X - xVal) ^ 2 + (PolyCoords(i).vertex(j).Y - yVal) ^ 2
                If currentDist < shortestDist Then
                    shortestDist = currentDist
                    pIndex = i
                    vIndex = j
                End If
            Next
            If pIndex > 0 And vIndex > 0 Then
                If (currentFunction = TOOL_VSELADD And vertexList(pIndex).vertex(vIndex) = 1) Or (currentFunction = TOOL_VSELSUB And vertexList(pIndex).vertex(vIndex) = 0) Then
                    pIndex = 0
                    vIndex = 0
                ElseIf currentFunction <> TOOL_VSELECT Then
                    Exit For
                End If
            End If
        End If
    Next

    If pIndex > 0 And vIndex > 0 Then
        If currentFunction = TOOL_VSELECT Then
            numSelectedPolys = numSelectedPolys + 1
            ReDim Preserve selectedPolys(numSelectedPolys)
            selectedPolys(numSelectedPolys) = pIndex
            vertexList(pIndex).vertex(vIndex) = 1
            RegionSelPolys = True
        ElseIf currentFunction = TOOL_VSELADD Then
            For j = 1 To 3
                selVerts = selVerts + vertexList(pIndex).vertex(j)
            Next
            If selVerts > 0 Then  ' poly already selected
                vertexList(pIndex).vertex(vIndex) = 1
            Else
                numSelectedPolys = numSelectedPolys + 1
                ReDim Preserve selectedPolys(numSelectedPolys)
                selectedPolys(numSelectedPolys) = pIndex
                vertexList(pIndex).vertex(vIndex) = 1
            End If
            RegionSelPolys = True
        ElseIf currentFunction = TOOL_VSELSUB Then
            vertexList(pIndex).vertex(vIndex) = 0
            For i = 1 To numSelectedPolys
                For j = 1 To 3
                    selVerts = selVerts + vertexList(selectedPolys(i)).vertex(j)
                Next
                If selVerts = 0 Then  ' no longer selected, put last here and shorten array
                    selectedPolys(i) = selectedPolys(numSelectedPolys)
                    numSelectedPolys = numSelectedPolys - 1
                End If
                selVerts = 0
            Next
            ReDim Preserve selectedPolys(numSelectedPolys)
            RegionSelPolys = True
        End If
    End If

End Function

Private Function RegionSelObjects(xVal As Single, yVal As Single, skipSel As Boolean) As Boolean

    Dim i As Integer
    Dim j As Integer
    Dim currentDist As Long
    Dim shortestDist As Long
    Dim Index As Integer

    shortestDist = (8 ^ 2 + 1)
    For i = 1 To spawnPoints
        If currentFunction = TOOL_VSELECT Then Spawns(i).active = 0
        If nearCoord(xVal, Spawns(i).X, 8 / zoomFactor) And nearCoord(yVal, Spawns(i).Y, 8 / zoomFactor) Then
            currentDist = (Spawns(i).X - xVal) ^ 2 + (Spawns(i).Y - yVal) ^ 2
            If currentDist < shortestDist Then
                shortestDist = currentDist
                Index = i
            End If
        End If
    Next

    If Index > 0 Then
        If currentFunction = TOOL_VSELECT Then
            Spawns(Index).active = 1
            numSelSpawns = numSelSpawns + 1
            skipSel = True
        ElseIf currentFunction = TOOL_VSELADD Then
            Spawns(Index).active = 1
            numSelSpawns = numSelSpawns + 1
            skipSel = True
        ElseIf currentFunction = TOOL_VSELSUB Then
            Spawns(Index).active = 0
            numSelSpawns = numSelSpawns - 1
            skipSel = True
        End If
    End If

    shortestDist = 64 ^ 2 + 1
    For i = 1 To colliderCount
        If currentFunction = TOOL_VSELECT Then Colliders(i).active = 0
        If nearCoord(xVal, Colliders(i).X, Colliders(i).radius / 2) And nearCoord(yVal, Colliders(i).Y, Colliders(i).radius / 2) Then
            currentDist = (Colliders(i).X - xVal) ^ 2 + (Colliders(i).Y - yVal) ^ 2
            If currentDist < shortestDist Then
                shortestDist = currentDist
                Index = i
            End If
        End If
    Next

    If Index > 0 And Not skipSel Then
        If currentFunction = TOOL_VSELECT Then
            Colliders(Index).active = 1
            numSelColliders = numSelColliders + 1
            skipSel = True
        ElseIf currentFunction = TOOL_VSELADD Then
            Colliders(Index).active = 1
            numSelColliders = numSelColliders + 1
            skipSel = True
        ElseIf currentFunction = TOOL_VSELSUB Then
            Colliders(Index).active = 0
            numSelColliders = numSelColliders - 1
            skipSel = True
        End If
    End If

    RegionSelObjects = skipSel

End Function

Private Function regionSelLights(xVal As Single, yVal As Single, skipSel As Boolean) As Boolean

    Dim i As Integer
    Dim j As Integer
    Dim currentDist As Long
    Dim shortestDist As Long
    Dim Index As Integer

    shortestDist = (8 ^ 2 + 1)
    For i = 1 To lightCount
        If currentFunction = TOOL_VSELECT Then Lights(i).selected = 0
        If nearCoord(xVal, Lights(i).X, 8 / zoomFactor) And nearCoord(yVal, Lights(i).Y, 8 / zoomFactor) Then
            currentDist = (Lights(i).X - xVal) ^ 2 + (Lights(i).Y - yVal) ^ 2
            If currentDist < shortestDist Then
                shortestDist = currentDist
                Index = i
            End If
        End If
    Next

    If Index > 0 And Not skipSel Then
        If currentFunction = TOOL_VSELECT Then
            Lights(Index).selected = 1
            numSelLights = numSelLights + 1
            skipSel = True
        ElseIf currentFunction = TOOL_VSELADD Then
            Lights(Index).selected = 1
            numSelLights = numSelLights + 1
            skipSel = True
        ElseIf currentFunction = TOOL_VSELSUB Then
            Lights(Index).selected = 0
            numSelLights = numSelLights - 1
            skipSel = True
        End If
    End If

    regionSelLights = skipSel

End Function

Private Function RegionSelWaypoints(xVal As Single, yVal As Single, skipSel As Boolean) As Boolean

    Dim i As Integer
    Dim j As Integer
    Dim currentDist As Long
    Dim shortestDist As Long
    Dim Index As Integer

    shortestDist = (8 ^ 2 + 1)
    For i = 1 To waypointCount
        If currentFunction = TOOL_VSELECT Then Waypoints(i).selected = False
        If (frmWaypoints.showPaths = Waypoints(i).pathNum) Or frmWaypoints.showPaths = 0 Then
            If nearCoord(xVal, Waypoints(i).X, 8 / zoomFactor) And nearCoord(yVal, Waypoints(i).Y, 8 / zoomFactor) Then
                currentDist = (Waypoints(i).X - xVal) ^ 2 + (Waypoints(i).Y - yVal) ^ 2
                If currentDist < shortestDist Then
                    shortestDist = currentDist
                    Index = i
                End If
            End If
        End If
    Next

    If Index > 0 And Not skipSel Then
        If currentFunction = TOOL_VSELECT Then
            Waypoints(Index).selected = True
            numSelWaypoints = numSelWaypoints + 1
        ElseIf currentFunction = TOOL_VSELADD Then
            Waypoints(Index).selected = True
            numSelWaypoints = numSelWaypoints + 1
        ElseIf currentFunction = TOOL_VSELSUB Then
            Waypoints(Index).selected = False
            numSelWaypoints = numSelWaypoints - 1
        End If
    End If

End Function

Private Function eraseSketch(X As Single, Y As Single) As Byte

    Dim i As Integer
    Dim j As Integer
    Dim currentDist As Long
    Dim shortestDist As Long
    Dim lineIndex As Integer

    On Error GoTo ErrorHandler

    eraseSketch = 0

    shortestDist = clrRadius ^ 2 + 1
    For i = 1 To sketchLines
        For j = 1 To 2
            currentDist = (X - sketch(i).vertex(j).X) ^ 2 + (Y - sketch(i).vertex(j).Y) ^ 2
            If (currentDist < shortestDist) Then
                shortestDist = currentDist
                lineIndex = i
            End If
        Next
    Next

    If lineIndex > 0 Then
        sketch(lineIndex) = sketch(sketchLines)
        sketchLines = sketchLines - 1
        ReDim Preserve sketch(sketchLines)
        Render
        eraseSketch = 1
    End If

    Exit Function

ErrorHandler:

    MsgBox "Error erasing sketch" & vbNewLine & Error$

End Function

Private Function moveLines(X As Single, Y As Single, xDiff As Single, yDiff As Single) As Byte

    Dim i As Integer
    Dim j As Integer
    Dim dist As Single

    On Error GoTo ErrorHandler

    xDiff = xDiff / zoomFactor
    yDiff = yDiff / zoomFactor

    moveLines = 0

    For i = 1 To sketchLines
        For j = 1 To 2
            dist = (X - sketch(i).vertex(j).X) ^ 2 + (Y - sketch(i).vertex(j).Y) ^ 2
            If dist < clrRadius ^ 2 Then
                sketch(i).vertex(j).X = sketch(i).vertex(j).X + xDiff * Cos((dist / (clrRadius ^ 2)) * PI / 2)
                sketch(i).vertex(j).Y = sketch(i).vertex(j).Y + yDiff * Cos((dist / (clrRadius ^ 2)) * PI / 2)
                moveLines = 1
            End If
        Next
    Next

    Exit Function

ErrorHandler:

    MsgBox "Error moving sketch lines" & vbNewLine & Error$

End Function

Private Sub deleteSmallLines()

    Dim i As Integer

    On Error GoTo ErrorHandler

    For i = 1 To sketchLines
        If (Int(sketch(i).vertex(1).X + 0.5) = Int(sketch(i).vertex(2).X + 0.5)) And (Int(sketch(i).vertex(1).Y + 0.5) = Int(sketch(i).vertex(2).Y + 0.5)) Then
            sketch(i) = sketch(sketchLines)
            sketchLines = sketchLines - 1
        End If
    Next

    ReDim Preserve sketch(sketchLines)

    Render

    Exit Sub

ErrorHandler:

    MsgBox "Error deleting small sketch lines" & vbNewLine & Error$

End Sub

Private Sub VertexSelection(X As Single, Y As Single)

    Dim i As Integer
    Dim j As Integer

    On Error GoTo ErrorHandler

    If currentFunction = TOOL_VSELECT Then
        numSelectedPolys = 0
        ReDim selectedPolys(numSelectedPolys)
        numSelectedScenery = 0
        numSelSpawns = 0
        numSelColliders = 0
        numSelWaypoints = 0
    ElseIf currentFunction = TOOL_VSELSUB Then
        numSelectedPolys = 0
        ReDim selectedPolys(numSelectedPolys)
    End If

    If showPolys Or showWireframe Or showPoints Then
        VertexSelPolys
    ElseIf currentFunction = TOOL_VSELECT Then
        For i = 1 To mPolyCount
            For j = 1 To 3
                vertexList(i).vertex(j) = 0
            Next
        Next
    End If

    If showScenery Then
        VertexSelScenery
    ElseIf currentFunction = TOOL_VSELECT Then
        For i = 1 To sceneryCount
            Scenery(i).selected = 0
        Next
    End If

    If showObjects Then
        VertexSelObjects
    ElseIf currentFunction = TOOL_VSELECT Then
        For i = 1 To spawnPoints
            Spawns(i).active = 0
        Next
        For i = 1 To colliderCount
            Colliders(i).active = 0
        Next
    End If

    If showWaypoints Then
        VertexSelWaypoints
    ElseIf currentFunction = TOOL_VSELECT Then
        For i = 1 To waypointCount
            Waypoints(i).selected = False
        Next
    End If

    If showLights Then
        VertexSelLights
    ElseIf currentFunction = TOOL_VSELECT Then
        For i = 1 To lightCount
            Lights(i).selected = 0
        Next
    End If

    currentWaypoint = 0

    selectedCoords(1).X = X
    selectedCoords(1).Y = Y
    selectedCoords(2).X = X
    selectedCoords(2).Y = Y

    getRCenter
    getInfo
    selectionChanged = True
    Render

    Exit Sub

ErrorHandler:

    MsgBox "Error selecting vertices" & vbNewLine & Error$

End Sub

Private Sub VertexSelPolys()

    Dim i As Integer
    Dim j As Integer
    Dim addPoly As Integer
    Dim notSel As Integer

    If currentFunction = TOOL_VSELECT Then
        For i = 1 To mPolyCount
            For j = 1 To 3
                vertexList(i).vertex(j) = 0
                If inSelRect(Polys(i).vertex(j).X, Polys(i).vertex(j).Y) Then
                    addPoly = 1
                    vertexList(i).vertex(j) = 1
                End If
            Next
            If addPoly = 1 Then
                numSelectedPolys = numSelectedPolys + 1
                ReDim Preserve selectedPolys(numSelectedPolys)
                selectedPolys(numSelectedPolys) = i
            End If
            addPoly = 0
            notSel = 0
        Next
    ElseIf currentFunction = TOOL_VSELADD Then
        For i = 1 To mPolyCount
            For j = 1 To 3
                If vertexList(i).vertex(j) = 0 Then
                    notSel = notSel + 1
                    If inSelRect(Polys(i).vertex(j).X, Polys(i).vertex(j).Y) Then
                        addPoly = 1
                        vertexList(i).vertex(j) = 1
                    End If
                End If
            Next
            If addPoly = 1 And notSel = 3 Then
                numSelectedPolys = numSelectedPolys + 1
                ReDim Preserve selectedPolys(numSelectedPolys)
                selectedPolys(numSelectedPolys) = i
            End If
            addPoly = 0
            notSel = 0
        Next
    ElseIf currentFunction = TOOL_VSELSUB Then
        For i = 1 To mPolyCount
            For j = 1 To 3
                If vertexList(i).vertex(j) = 1 Then  ' if already selected and if in range
                    If inSelRect(Polys(i).vertex(j).X, Polys(i).vertex(j).Y) Then
                        notSel = notSel + 1
                        vertexList(i).vertex(j) = 0
                    Else  ' if already selected but not in range
                        addPoly = 1
                    End If
                End If
            Next
            If addPoly = 1 Then
                numSelectedPolys = numSelectedPolys + 1
                ReDim Preserve selectedPolys(numSelectedPolys)
                selectedPolys(numSelectedPolys) = i
            End If
            addPoly = 0
            notSel = 0
        Next
    End If

End Sub

Private Sub VertexSelScenery()

    Dim i As Integer
    Dim sVal As Integer
    Dim sceneryCoords(3) As TCustomVertex
    Dim selected(3) As Boolean

    For i = 1 To sceneryCount
        sVal = Scenery(i).Style

        sceneryCoords(0).X = (Scenery(i).Translation.X - scrollCoords(2).X) * zoomFactor
        sceneryCoords(0).Y = (Scenery(i).Translation.Y - scrollCoords(2).Y) * zoomFactor
        sceneryCoords(1).X = sceneryCoords(0).X + Cos(Scenery(i).rotation) * (SceneryTextures(sVal).Width) * Scenery(i).Scaling.X * zoomFactor
        sceneryCoords(1).Y = sceneryCoords(0).Y - Sin(Scenery(i).rotation) * (SceneryTextures(sVal).Width) * Scenery(i).Scaling.X * zoomFactor
        sceneryCoords(3).X = sceneryCoords(0).X + Sin(Scenery(i).rotation) * (SceneryTextures(sVal).Height) * Scenery(i).Scaling.Y * zoomFactor
        sceneryCoords(3).Y = sceneryCoords(0).Y + Cos(Scenery(i).rotation) * (SceneryTextures(sVal).Height) * Scenery(i).Scaling.Y * zoomFactor
        sceneryCoords(2).X = sceneryCoords(3).X + sceneryCoords(1).X - sceneryCoords(0).X
        sceneryCoords(2).Y = sceneryCoords(3).Y + sceneryCoords(1).Y - sceneryCoords(0).Y

        selected(0) = inSelRect(sceneryCoords(0).X, sceneryCoords(0).Y)
        If sceneryVerts Then
            selected(1) = inSelRect(sceneryCoords(1).X, sceneryCoords(1).Y)
            selected(2) = inSelRect(sceneryCoords(2).X, sceneryCoords(2).Y)
            selected(3) = inSelRect(sceneryCoords(3).X, sceneryCoords(3).Y)
        Else
            selected(1) = False
            selected(2) = False
            selected(3) = False
        End If

        If currentFunction = TOOL_VSELECT Then
            Scenery(i).selected = 0
        End If

        If showWireframe Or ((Scenery(i).level = 0 And sslBack) Or (Scenery(i).level = 1 And sslMid) Or (Scenery(i).level = 2 And sslFront)) Then
            If selected(0) Or selected(1) Or selected(2) Or selected(3) Then
                If currentFunction = TOOL_VSELECT Then
                    Scenery(i).selected = 1
                    numSelectedScenery = numSelectedScenery + 1
                ElseIf currentFunction = TOOL_VSELADD Then
                    If Scenery(i).selected = 0 Then
                        numSelectedScenery = numSelectedScenery + 1
                    End If
                    Scenery(i).selected = 1
                ElseIf currentFunction = TOOL_VSELSUB Then
                    If Scenery(i).selected = 1 Then
                        numSelectedScenery = numSelectedScenery - 1
                    End If
                    Scenery(i).selected = 0
                End If
            End If
        End If
    Next

End Sub

Private Sub VertexSelObjects()

    Dim i As Integer
    Dim j As Integer
    Dim xCoord As Single
    Dim yCoord As Single

    For i = 1 To spawnPoints
        xCoord = (Spawns(i).X - scrollCoords(2).X) * zoomFactor
        yCoord = (Spawns(i).Y - scrollCoords(2).Y) * zoomFactor
        If currentFunction = TOOL_VSELECT Then Spawns(i).active = 0
        If inSelRect(xCoord, yCoord) Then
            If currentFunction = TOOL_VSELECT Then
                Spawns(i).active = 1
                numSelSpawns = numSelSpawns + 1
            ElseIf currentFunction = TOOL_VSELADD And Spawns(i).active = 0 Then
                numSelSpawns = numSelSpawns + 1
                Spawns(i).active = 1
            ElseIf currentFunction = TOOL_VSELSUB And Spawns(i).active = 1 Then
                numSelSpawns = numSelSpawns - 1
                Spawns(i).active = 0
            End If
        End If
    Next

    For i = 1 To colliderCount
        xCoord = (Colliders(i).X - scrollCoords(2).X) * zoomFactor
        yCoord = (Colliders(i).Y - scrollCoords(2).Y) * zoomFactor
        If currentFunction = TOOL_VSELECT Then Colliders(i).active = 0
        If inSelRect(xCoord, yCoord) Then
            If currentFunction = TOOL_VSELECT Then
                numSelColliders = numSelColliders + 1
                Colliders(i).active = 1
            ElseIf currentFunction = TOOL_VSELADD And Colliders(i).active = 0 Then
                numSelColliders = numSelColliders + 1
                Colliders(i).active = 1
            ElseIf currentFunction = TOOL_VSELSUB And Colliders(i).active = 1 Then
                numSelColliders = numSelColliders - 1
                Colliders(i).active = 0
            End If
        End If
    Next

End Sub

Private Sub VertexSelLights()

    Dim i As Integer
    Dim j As Integer
    Dim xCoord As Long
    Dim yCoord As Long

    For i = 1 To lightCount
        xCoord = (Lights(i).X - scrollCoords(2).X) * zoomFactor
        yCoord = (Lights(i).Y - scrollCoords(2).Y) * zoomFactor
        If currentFunction = TOOL_VSELECT Then Lights(i).selected = 0
        If inSelRect(xCoord, yCoord) Then
            If currentFunction = TOOL_VSELECT Then
                Lights(i).selected = 1
                numSelLights = numSelLights + 1
            ElseIf currentFunction = TOOL_VSELADD And Lights(i).selected = 0 Then
                numSelLights = numSelLights + 1
                Lights(i).selected = 1
            ElseIf currentFunction = TOOL_VSELSUB And Lights(i).selected = 1 Then
                numSelLights = numSelLights - 1
                Lights(i).selected = 0
            End If
        End If
    Next

End Sub

Private Sub VertexSelWaypoints()

    Dim i As Integer
    Dim j As Integer
    Dim xCoord As Long
    Dim yCoord As Long

    For i = 1 To waypointCount
        If (frmWaypoints.showPaths = Waypoints(i).pathNum) Or frmWaypoints.showPaths = 0 Then
            xCoord = (Waypoints(i).X - scrollCoords(2).X) * zoomFactor
            yCoord = (Waypoints(i).Y - scrollCoords(2).Y) * zoomFactor
            If currentFunction = TOOL_VSELECT Then Waypoints(i).selected = False
            If inSelRect(xCoord, yCoord) Then
                If currentFunction = TOOL_VSELECT Then
                    Waypoints(i).selected = True
                    numSelWaypoints = numSelWaypoints + 1
                ElseIf currentFunction = TOOL_VSELADD And Not Waypoints(i).selected Then
                    numSelWaypoints = numSelWaypoints + 1
                    Waypoints(i).selected = True
                ElseIf currentFunction = TOOL_VSELSUB And Waypoints(i).selected Then
                    numSelWaypoints = numSelWaypoints - 1
                    Waypoints(i).selected = False
                End If
            End If
        End If
    Next

End Sub

Private Sub getRCenter()

    Dim i As Integer
    Dim j As Integer
    Dim setCoords As Boolean
    Dim xVal As Single
    Dim yVal As Single
    Dim Width As Single
    Dim Height As Single

    On Error GoTo ErrorHandler

    If numSelectedPolys > 0 Then
        For j = 1 To 3
            If vertexList(selectedPolys(1)).vertex(j) = 1 Then
                selRect(0).X = PolyCoords(selectedPolys(1)).vertex(j).X
                selRect(0).Y = PolyCoords(selectedPolys(1)).vertex(j).Y
                selRect(2).X = PolyCoords(selectedPolys(1)).vertex(j).X
                selRect(2).Y = PolyCoords(selectedPolys(1)).vertex(j).Y
            End If
        Next
        For i = 1 To numSelectedPolys
            For j = 1 To 3
                If vertexList(selectedPolys(i)).vertex(j) = 1 Then
                    compareRect PolyCoords(selectedPolys(i)).vertex(j).X, PolyCoords(selectedPolys(i)).vertex(j).Y
                End If
            Next
        Next
    End If
    If numSelectedScenery > 0 Then
        setCoords = False
        For i = 1 To sceneryCount
            If Scenery(i).selected = 1 Then

                If Not setCoords And numSelectedPolys = 0 Then
                    setCoords = True
                    selRect(0).X = Scenery(i).Translation.X
                    selRect(0).Y = Scenery(i).Translation.Y
                    selRect(2).X = Scenery(i).Translation.X
                    selRect(2).Y = Scenery(i).Translation.Y
                End If
                compareRect Scenery(i).Translation.X, Scenery(i).Translation.Y

                Width = SceneryTextures(Scenery(i).Style).Width * Scenery(i).Scaling.X
                Height = SceneryTextures(Scenery(i).Style).Height * Scenery(i).Scaling.Y

                xVal = Scenery(i).Translation.X + (Cos(Scenery(i).rotation) * Width) + (Sin(Scenery(i).rotation) * Height)
                yVal = Scenery(i).Translation.Y - (Sin(Scenery(i).rotation) * Width) + (Cos(Scenery(i).rotation) * Height)

                compareRect xVal, yVal

            End If
        Next
    End If

    If numSelWaypoints > 0 Then
        setCoords = False
        For i = 1 To waypointCount
            If Waypoints(i).selected Then
                If Not setCoords And numSelectedPolys = 0 And numSelectedScenery = 0 Then
                    setCoords = True
                    selRect(0).X = Waypoints(i).X
                    selRect(0).Y = Waypoints(i).Y
                    selRect(2).X = Waypoints(i).X
                    selRect(2).Y = Waypoints(i).Y
                End If
                compareRect Waypoints(i).X, Waypoints(i).Y
            End If
        Next
    End If

    If numSelColliders > 0 Then
        setCoords = False
        For i = 1 To colliderCount
            If Colliders(i).active Then
                If Not setCoords And numSelectedPolys = 0 And numSelectedScenery = 0 Then
                    setCoords = True
                    selRect(0).X = Colliders(i).X
                    selRect(0).Y = Colliders(i).Y
                    selRect(2).X = Colliders(i).X
                    selRect(2).Y = Colliders(i).Y
                End If
                compareRect Colliders(i).X, Colliders(i).Y
            End If
        Next
    End If

    If numSelSpawns > 0 Then
        setCoords = False
        For i = 1 To spawnPoints
            If Spawns(i).active Then
                If Not setCoords And numSelectedPolys = 0 And numSelectedScenery = 0 Then
                    setCoords = True
                    selRect(0).X = Spawns(i).X
                    selRect(0).Y = Spawns(i).Y
                    selRect(2).X = Spawns(i).X
                    selRect(2).Y = Spawns(i).Y
                End If
                compareRect Spawns(i).X, Spawns(i).Y
            End If
        Next
    End If

    If numSelLights > 0 Then
        setCoords = False
        For i = 1 To lightCount
            If Lights(i).selected Then
                If Not setCoords And numSelectedPolys = 0 And numSelectedScenery = 0 Then
                    setCoords = True
                    selRect(0).X = Lights(i).X
                    selRect(0).Y = Lights(i).Y
                    selRect(2).X = Lights(i).X
                    selRect(2).Y = Lights(i).Y
                End If
                compareRect Lights(i).X, Lights(i).Y
            End If
        Next
    End If

    selRect(1).X = selRect(2).X
    selRect(1).Y = selRect(0).Y
    selRect(3).X = selRect(0).X
    selRect(3).Y = selRect(2).Y

    If mnuFixedRCenter.Checked Then
        rCenter.X = Midpoint(selRect(0).X, selRect(2).X)
        rCenter.Y = Midpoint(selRect(0).Y, selRect(2).Y)
    End If

    Exit Sub

ErrorHandler:

    MsgBox Error$

End Sub

Private Sub compareRect(ByVal xVal As Single, ByVal yVal As Single)

    If xVal < selRect(0).X Then selRect(0).X = xVal
    If xVal > selRect(2).X Then selRect(2).X = xVal
    If yVal < selRect(0).Y Then selRect(0).Y = yVal
    If yVal > selRect(2).Y Then selRect(2).Y = yVal

End Sub

Private Sub vertexSelAlt(X As Single, Y As Single)

    Dim i As Integer
    Dim j As Integer
    Dim xDist As Integer
    Dim yDist As Integer
    Dim xCenter As Integer
    Dim yCenter As Integer
    Dim addPoly As Integer
    Dim notSel As Integer

    xDist = (X - selectedCoords(1).X) / 2  ' x distance from coord
    yDist = (Y - selectedCoords(1).Y) / 2  ' y distance from coord

    xCenter = X - xDist
    yCenter = Y - yDist

    numSelectedPolys = 0
    ReDim selectedPolys(numSelectedPolys)

    For i = 1 To mPolyCount
        For j = 1 To 3
            ' if in range
            If nearCoord(xCenter, Polys(i).vertex(j).X, Abs(xDist)) And nearCoord(yCenter, Polys(i).vertex(j).Y, Abs(yDist)) Then
                If vertexList(i).vertex(j) = 0 Then
                    vertexList(i).vertex(j) = 1
                    addPoly = 1
                Else
                    vertexList(i).vertex(j) = 0
                End If
            ElseIf vertexList(i).vertex(j) = 1 Then
                addPoly = 1
            End If
        Next
        If addPoly = 1 Then
            numSelectedPolys = numSelectedPolys + 1
            ReDim Preserve selectedPolys(numSelectedPolys)
            selectedPolys(numSelectedPolys) = i
        End If
        addPoly = 0
    Next

    selectedCoords(1).X = X
    selectedCoords(1).Y = Y
    selectedCoords(2).X = X
    selectedCoords(2).Y = Y

    Render

End Sub

Private Sub polySelection(X As Single, Y As Single)

    Dim i As Integer
    Dim j As Integer
    Dim addPoly As Integer
    Dim shortestDist As Integer
    Dim firstClicked As Integer
    Dim foundSelected As Integer

    addPoly = 0
    If currentFunction = TOOL_PSELECT Then  ' select polys (destroy and remake)
        ReDim selectedPolys(0)
        numSelectedPolys = 0
        numSelectedScenery = 0

        If showPolys Or showWireframe Or showPoints Then
            shortestDist = 16 ^ 2 + 1
            For i = 1 To mPolyCount
                If (pointInPoly(X, Y, i)) And addPoly = 0 Then  ' if in poly and no other poly selected
                    If firstClicked = 0 Then
                        firstClicked = i
                    End If
                    ' not selected and after selected
                    If foundSelected > 0 And (vertexList(i).vertex(1) + vertexList(i).vertex(1) + vertexList(i).vertex(1) < 3) Then
                        vertexList(i).vertex(1) = 1
                        vertexList(i).vertex(2) = 1
                        vertexList(i).vertex(3) = 1
                        numSelectedPolys = numSelectedPolys + 1
                        ReDim Preserve selectedPolys(numSelectedPolys)
                        selectedPolys(numSelectedPolys) = i
                        addPoly = 1
                    ' not selected, not found
                    ElseIf (vertexList(i).vertex(1) + vertexList(i).vertex(1) + vertexList(i).vertex(1) < 3) Then
                    Else  ' poly is selected
                        foundSelected = i
                        vertexList(i).vertex(1) = 0
                        vertexList(i).vertex(2) = 0
                        vertexList(i).vertex(3) = 0
                    End If
                Else
                    vertexList(i).vertex(1) = 0
                    vertexList(i).vertex(2) = 0
                    vertexList(i).vertex(3) = 0
                End If
            Next
        End If

        If addPoly = 0 And firstClicked > 0 Then
            vertexList(firstClicked).vertex(1) = 1
            vertexList(firstClicked).vertex(2) = 1
            vertexList(firstClicked).vertex(3) = 1
            numSelectedPolys = numSelectedPolys + 1
            ReDim Preserve selectedPolys(numSelectedPolys)
            selectedPolys(numSelectedPolys) = firstClicked
            addPoly = 1
        End If

        If showScenery And addPoly = 0 Then
            For i = 1 To sceneryCount
                Scenery(i).selected = 0
                If showWireframe Or ((Scenery(i).level = 0 And sslBack) Or (Scenery(i).level = 1 And sslMid) Or (Scenery(i).level = 2 And sslFront)) Then
                    If PointInProp(X, Y, i) And addPoly = 0 Then
                        Scenery(i).selected = 1
                        numSelectedScenery = numSelectedScenery + 1
                        addPoly = 1
                    End If
                End If
            Next
        Else
            For i = 1 To sceneryCount
                Scenery(i).selected = 0
            Next
        End If

        If showObjects Then
            For i = 1 To spawnPoints
                Spawns(i).active = 0
            Next
            numSelSpawns = 0
            For i = 1 To colliderCount
                Colliders(i).active = 0
            Next
            numSelColliders = 0
        End If
        If showLights Then
            For i = 1 To lightCount
                Lights(i).selected = 0
            Next
            numSelLights = 0
        End If
        If showWaypoints Then
            For i = 1 To waypointCount
                If (frmWaypoints.showPaths = Waypoints(i).pathNum) Or frmWaypoints.showPaths = 0 Then
                    Waypoints(i).selected = False
                End If
            Next
            numSelWaypoints = 0
        End If
    ElseIf currentFunction = TOOL_PSELADD Then  ' add polys
        addPoly = 0
        If showPolys Or showWireframe Or showPoints Then
            For i = 1 To mPolyCount
                If pointInPoly(X, Y, i) And vertexList(i).vertex(1) = 0 And addPoly = 0 Then  ' if in poly and not already selected
                    numSelectedPolys = numSelectedPolys + 1
                    ReDim Preserve selectedPolys(numSelectedPolys)
                    selectedPolys(numSelectedPolys) = i
                    vertexList(i).vertex(1) = 1
                    vertexList(i).vertex(2) = 1
                    vertexList(i).vertex(3) = 1
                    addPoly = 1
                End If
            Next
        End If

        If showScenery And addPoly = 0 Then
            For i = 1 To sceneryCount
                If Scenery(i).selected = 0 And addPoly = 0 Then
                    If PointInProp(X, Y, i) Then
                        Scenery(i).selected = 1
                        numSelectedScenery = numSelectedScenery + 1
                        addPoly = 1
                    End If
                End If
            Next
        End If
    ElseIf currentFunction = TOOL_PSELSUB Then  ' subtract polys
        ReDim selectedPolys(1)
        numSelectedPolys = 0

        If showPolys Or showWireframe Or showPoints Then
            For i = 1 To mPolyCount
                If vertexList(i).vertex(1) = 1 Then  ' if poly already selected
                    If (pointInPoly(X, Y, i)) And addPoly = 0 Then  ' if poly clicked
                        vertexList(i).vertex(1) = 0
                        vertexList(i).vertex(2) = 0
                        vertexList(i).vertex(3) = 0
                        addPoly = 1
                    Else
                        numSelectedPolys = numSelectedPolys + 1
                        ReDim Preserve selectedPolys(numSelectedPolys)
                        selectedPolys(numSelectedPolys) = i
                    End If
                End If
            Next
        End If

        If showScenery And addPoly = 0 Then
            For i = 1 To sceneryCount
                If Scenery(i).selected = 1 And addPoly = 0 Then
                    If PointInProp(X, Y, i) Then
                        Scenery(i).selected = 0
                        numSelectedScenery = numSelectedScenery - 1
                        addPoly = 1
                    End If
                End If
            Next
        End If
    End If

    getRCenter
    getInfo
    selectionChanged = True
    Render

End Sub

Private Function PointInProp(ByVal X As Single, ByVal Y As Single, Index As Integer) As Boolean

    Dim xDiff As Long
    Dim yDiff As Long
    Dim theta As Single
    Dim R As Single

    On Error GoTo ErrorHandler

    PointInProp = False

    xDiff = (X - Scenery(Index).screenTr.X)
    yDiff = (Y - Scenery(Index).screenTr.Y)

    R = Sqr((xDiff) ^ 2 + (yDiff) ^ 2)  ' distance of point from scenery rotation center
    If xDiff = 0 Then
        If yDiff > 0 Then
            theta = PI / 2 + Scenery(Index).rotation
        Else
            theta = 3 * PI / 2 + Scenery(Index).rotation
        End If
    ElseIf xDiff > 0 Then
        theta = Atn(yDiff / xDiff) + Scenery(Index).rotation
    ElseIf xDiff < 0 Then
        theta = PI + Atn(yDiff / xDiff) + Scenery(Index).rotation
    End If

    X = R * Cos(theta)
    Y = R * Sin(theta)

    If isBetween(0, X, SceneryTextures(Scenery(Index).Style).Width * Scenery(Index).Scaling.X * zoomFactor) Then
        If isBetween(0, Y, SceneryTextures(Scenery(Index).Style).Height * Scenery(Index).Scaling.Y * zoomFactor) Then
            PointInProp = True
        End If
    End If

    Exit Function

ErrorHandler:

    MsgBox "Error selecting scenery" & vbNewLine & Error$

End Function

Private Sub ColorFill(X As Single, Y As Single)

    Dim i As Integer
    Dim j As Integer
    Dim PolyNum As Integer
    Dim destClr As TColor
    Dim polyColored As Boolean

    If numSelectedPolys > 0 Or numSelectedScenery > 0 Then
        If showPolys Or showWireframe Or showPoints Then
            For i = 1 To numSelectedPolys
                PolyNum = selectedPolys(i)
                For j = 1 To 3
                    If vertexList(PolyNum).vertex(j) = 1 Then
                        If selectionChanged Then
                            SaveUndo
                            selectionChanged = False
                        End If
                        destClr = getRGB(Polys(PolyNum).vertex(j).color)
                        destClr = applyBlend(destClr)
                        Polys(PolyNum).vertex(j).color = ARGB(getAlpha(Polys(PolyNum).vertex(j).color), RGB(destClr.blue, destClr.green, destClr.red))
                        vertexList(PolyNum).color(j).red = destClr.red
                        vertexList(PolyNum).color(j).green = destClr.green
                        vertexList(PolyNum).color(j).blue = destClr.blue
                        applyLightsToVert PolyNum, j
                        polyColored = True
                    End If
                Next
            Next
        End If

        If showScenery Then
            For i = 1 To sceneryCount
                If Scenery(i).selected = 1 Then
                    If selectionChanged Then
                        SaveUndo
                        selectionChanged = False
                    End If
                    destClr = getRGB(Scenery(i).color)
                    destClr = applyBlend(destClr)
                    Scenery(i).color = ARGB(Scenery(i).alpha, RGB(destClr.blue, destClr.green, destClr.red))
                    polyColored = True
                End If
            Next
        End If

        If polyColored Then
            SaveUndo
        End If
    Else
        If showPolys Or showWireframe Or showPoints Then
            For i = 1 To mPolyCount
                If (pointInPoly(X, Y, i)) Then
                    For j = 1 To 3
                        If selectionChanged Then
                            SaveUndo
                            selectionChanged = False
                        End If
                        destClr = getRGB(Polys(i).vertex(j).color)  ' get clr of poly
                        destClr = applyBlend(destClr)
                        Polys(i).vertex(j).color = ARGB(getAlpha(Polys(i).vertex(j).color), RGB(destClr.blue, destClr.green, destClr.red))
                        vertexList(i).color(j).red = destClr.red
                        vertexList(i).color(j).green = destClr.green
                        vertexList(i).color(j).blue = destClr.blue
                        applyLightsToVert i, j
                        polyColored = True
                    Next
                End If
            Next
        End If

        If Not polyColored And showScenery Then
            For i = 1 To sceneryCount
                If PointInProp(X, Y, i) Then
                    If selectionChanged Then
                        SaveUndo
                        selectionChanged = False
                    End If
                    destClr = getRGB(Scenery(i).color)
                    destClr = applyBlend(destClr)
                    Scenery(i).color = ARGB(Scenery(i).alpha, RGB(destClr.blue, destClr.green, destClr.red))
                    polyColored = True
                End If
            Next
        End If

        If polyColored Then
            SaveUndo
        End If
    End If

    prompt = True

    Render

End Sub

Private Function applyBlend(dClr As TColor) As TColor

    If blendMode = 0 Then  ' normal
        applyBlend.red = gPolyClr.red * opacity + dClr.red * (1 - opacity)
        applyBlend.green = gPolyClr.green * opacity + dClr.green * (1 - opacity)
        applyBlend.blue = gPolyClr.blue * opacity + dClr.blue * (1 - opacity)
    ElseIf blendMode = 1 Then  ' multiply
        applyBlend.red = (dClr.red / 255 * gPolyClr.red) * opacity + dClr.red * (1 - opacity)
        applyBlend.green = (dClr.green / 255 * gPolyClr.green) * opacity + dClr.green * (1 - opacity)
        applyBlend.blue = (dClr.blue / 255 * gPolyClr.blue) * opacity + dClr.blue * (1 - opacity)
    ElseIf blendMode = 2 Then  ' screen
        applyBlend.red = (dClr.red - dClr.red / 255 * gPolyClr.red + gPolyClr.red) * opacity + dClr.red * (1 - opacity)
        applyBlend.green = (dClr.green - dClr.green / 255 * gPolyClr.green + gPolyClr.green) * opacity + dClr.green * (1 - opacity)
        applyBlend.blue = (dClr.blue - dClr.blue / 255 * gPolyClr.blue + gPolyClr.blue) * opacity + dClr.blue * (1 - opacity)
    ElseIf blendMode = 3 Then  ' AND ' darken
        applyBlend.red = lowerVal(dClr.red, gPolyClr.red) * opacity + dClr.red * (1 - opacity)
        applyBlend.green = lowerVal(dClr.green, gPolyClr.green) * opacity + dClr.green * (1 - opacity)
        applyBlend.blue = lowerVal(dClr.blue, gPolyClr.blue) * opacity + dClr.blue * (1 - opacity)
    ElseIf blendMode = 4 Then  ' OR ' lighten
        applyBlend.red = higherVal(dClr.red, gPolyClr.red) * opacity + dClr.red * (1 - opacity)
        applyBlend.green = higherVal(dClr.green, gPolyClr.green) * opacity + dClr.green * (1 - opacity)
        applyBlend.blue = higherVal(dClr.blue, gPolyClr.blue) * opacity + dClr.blue * (1 - opacity)
    ElseIf blendMode = 5 Then  ' XOR ' difference
        applyBlend.red = diffVal(dClr.red, gPolyClr.red) * opacity + dClr.red * (1 - opacity)
        applyBlend.green = diffVal(dClr.green, gPolyClr.green) * opacity + dClr.green * (1 - opacity)
        applyBlend.blue = diffVal(dClr.blue, gPolyClr.blue) * opacity + dClr.blue * (1 - opacity)
    Else
        applyBlend.red = 0
        applyBlend.green = 0
        applyBlend.blue = 0
    End If

End Function

Private Function snapVertexToGrid(ByVal coord As Single, offset As Single) As Single

    Dim target As Single

    offset = (inc * zoomFactor) - offset

    target = (Int(coord / (inc * zoomFactor)) * (inc * zoomFactor) + offset)
    If target > coord Then target = target - inc * zoomFactor

    If (coord - target) < ((inc * zoomFactor) / 2) Then
        snapVertexToGrid = target
    Else
        snapVertexToGrid = target + inc * zoomFactor
    End If

End Function

Private Sub deletePolys()

    Dim i As Integer
    Dim j As Integer
    Dim offset As Integer

    On Error GoTo ErrorHandler

    If selectionChanged Then
        SaveUndo
        selectionChanged = False
    End If

    prompt = True

    If numSelectedScenery > 0 Then
        offset = 1
        For i = 1 To sceneryCount
            Scenery(offset) = Scenery(i)
            If Scenery(i).selected = 1 Then  ' scenery selected
                sceneryCount = sceneryCount - 1
            Else  ' not selected
                offset = offset + 1
            End If
        Next
        ReDim Preserve Scenery(sceneryCount)
    End If

    If numSelSpawns > 0 Then
        offset = 1
        For i = 1 To spawnPoints
            Spawns(offset) = Spawns(i)
            If Spawns(i).active = 1 Then
                spawnPoints = spawnPoints - 1
            Else  ' not selected
                offset = offset + 1
            End If
        Next
        ReDim Preserve Spawns(spawnPoints)
    End If

    If numSelColliders > 0 Then
        offset = 1
        For i = 1 To colliderCount
            Colliders(offset) = Colliders(i)
            If Colliders(i).active = 1 Then  ' scenery selected
                colliderCount = colliderCount - 1
            Else  ' not selected
                offset = offset + 1
            End If
        Next
        ReDim Preserve Colliders(colliderCount)
    End If

    If numSelWaypoints > 0 Then
        currentWaypoint = 0
        offset = 1
        For i = 1 To waypointCount
            Waypoints(i).tempIndex = Waypoints(offset).tempIndex
            Waypoints(offset) = Waypoints(i)
            If Waypoints(i).selected Then
                waypointCount = waypointCount - 1
                Waypoints(i).tempIndex = -1
            Else  ' not selected
                Waypoints(i).tempIndex = offset
                offset = offset + 1
            End If
        Next

        offset = 1
        For i = 1 To conCount
            Connections(offset) = Connections(i)
            If Waypoints(Connections(i).point1).tempIndex < 0 Or Waypoints(Connections(i).point2).tempIndex < 0 Then
                conCount = conCount - 1
            Else  ' not selected
                Connections(offset).point1 = Waypoints(Connections(offset).point1).tempIndex
                Connections(offset).point2 = Waypoints(Connections(offset).point2).tempIndex
                offset = offset + 1
            End If
        Next
        For i = 1 To waypointCount
            Waypoints(i).tempIndex = i
            Waypoints(i).numConnections = 0
        Next
        ReDim Preserve Waypoints(waypointCount)
        ReDim Preserve Connections(conCount)
        For i = 1 To conCount
            Waypoints(Connections(i).point1).numConnections = Waypoints(Connections(i).point1).numConnections + 1
        Next
    End If

    If numSelLights > 0 Then
        offset = 1
        For i = 1 To lightCount
            Lights(offset) = Lights(i)
            If Lights(i).selected = 1 Then
                lightCount = lightCount - 1
            Else  ' not selected
                offset = offset + 1
            End If
        Next
        ReDim Preserve Lights(lightCount)
        If lightCount > 0 Then
            applyLights
        ElseIf lightCount = 0 Then
            For i = 1 To mPolyCount
                For j = 1 To 3
                    Polys(i).vertex(j).color = ARGB(getAlpha(Polys(i).vertex(j).color), RGB(vertexList(i).color(j).blue, vertexList(i).color(j).green, vertexList(i).color(j).red))
                Next
            Next
        End If
    End If

    numSelectedScenery = 0
    numSelSpawns = 0
    numSelColliders = 0
    numSelWaypoints = 0
    numSelLights = 0

    If numSelectedPolys > 0 Then  ' delete polys
        numSelectedPolys = 0
        ReDim selectedPolys(0)

        offset = 1

        For i = 1 To mPolyCount
            Polys(offset) = Polys(i)
            PolyCoords(offset) = PolyCoords(i)
            vertexList(offset) = vertexList(i)

            If (vertexList(i).vertex(1) + vertexList(i).vertex(2) + vertexList(i).vertex(3)) = 3 Then  ' poly selected
                vertexList(offset).vertex(1) = 0
                vertexList(offset).vertex(2) = 0
                vertexList(offset).vertex(3) = 0
                mPolyCount = mPolyCount - 1
            ElseIf (vertexList(i).vertex(1) + vertexList(i).vertex(2) + vertexList(i).vertex(3)) > 0 Then  ' vertices selected
                numSelectedPolys = numSelectedPolys + 1
                ReDim Preserve selectedPolys(numSelectedPolys)
                selectedPolys(numSelectedPolys) = offset
                offset = offset + 1
            Else  ' not selected
                offset = offset + 1
            End If
        Next

        ReDim Preserve Polys(mPolyCount)
        ReDim Preserve PolyCoords(mPolyCount)
        ReDim Preserve vertexList(mPolyCount)
    End If

    setMapData

    SaveUndo
    Render
    getInfo

    Exit Sub

ErrorHandler:

    MsgBox "Error deleting" & vbNewLine & Error$

End Sub

Private Function nearCoord(ByVal mouseCoord As Single, ByVal polyCoord As Single, ByVal range As Single) As Boolean

    If mouseCoord <= (polyCoord + range) And mouseCoord >= (polyCoord - range) Then
        nearCoord = True
    End If

End Function

Private Function inSelRect(ByVal X As Single, ByVal Y As Single) As Boolean

    If (X > selectedCoords(1).X And X < selectedCoords(2).X) Or (X > selectedCoords(2).X And X < selectedCoords(1).X) Then
        If (Y > selectedCoords(1).Y And Y < selectedCoords(2).Y) Or (Y > selectedCoords(2).Y And Y < selectedCoords(1).Y) Then
            inSelRect = True
        End If
    End If

End Function

Private Sub lblMousePosition_Click()
    ReleaseCapture
    SendMessage Me.hWnd, WM_NCLBUTTONDOWN, 2, 0&
    formLeft = Me.Left / Screen.TwipsPerPixelX
    formTop = Me.Top / Screen.TwipsPerPixelY
End Sub

Private Sub mnuClrSketch_Click()

    sketchLines = 0
    ReDim Preserve sketch(0)

End Sub

Private Sub mnuCopy_Click()

    savePrefab appPath & "\Temp\copy.PFB"

End Sub

Private Sub mnuVSelBringForward_Click()

    mnuBringForward_Click

End Sub

Private Sub mnuVSelBringToFront_Click()

    mnuBringToFront_Click

End Sub

Private Sub mnuVSelClear_Click()

    mnuClear_Click

End Sub

Private Sub mnuVSelCopy_Click()

    mnuCopy_Click

End Sub

Private Sub mnuFlip_Click(Index As Integer)

    Dim i As Integer
    Dim j As Integer
    Dim PolyNum As Integer
    Dim vertSel As Byte
    Dim temp As D3DVECTOR2
    Dim tempVertex As TCustomVertex
    Dim tempClr As TColor

    If selectionChanged Then
        SaveUndo
        selectionChanged = False
    End If

    If Index = 0 Then
        scaleDiff.X = -1
    ElseIf Index = 1 Then
        scaleDiff.Y = -1
    End If

    rCenter.X = selRect(0).X + (selRect(2).X - selRect(0).X) / 2
    rCenter.Y = selRect(0).Y + (selRect(2).Y - selRect(0).Y) / 2

    If numSelectedPolys > 0 Then
        For i = 1 To numSelectedPolys
            PolyNum = selectedPolys(i)
            For j = 1 To 3
                If vertexList(PolyNum).vertex(j) = 1 Then
                    PolyCoords(PolyNum).vertex(j).X = (rCenter.X + (PolyCoords(PolyNum).vertex(j).X - rCenter.X) * scaleDiff.X)
                    PolyCoords(PolyNum).vertex(j).Y = (rCenter.Y + (PolyCoords(PolyNum).vertex(j).Y - rCenter.Y) * scaleDiff.Y)
                    Polys(PolyNum).vertex(j).X = (PolyCoords(PolyNum).vertex(j).X - scrollCoords(2).X) * zoomFactor
                    Polys(PolyNum).vertex(j).Y = (PolyCoords(PolyNum).vertex(j).Y - scrollCoords(2).Y) * zoomFactor
                End If
            Next

            ' make sure polys are cw
            If Not isCW(PolyNum) Then  ' switch to make cw
                temp = PolyCoords(PolyNum).vertex(3)
                PolyCoords(PolyNum).vertex(3) = PolyCoords(PolyNum).vertex(2)
                PolyCoords(PolyNum).vertex(2) = temp

                tempVertex = Polys(PolyNum).vertex(3)
                Polys(PolyNum).vertex(3) = Polys(PolyNum).vertex(2)
                Polys(PolyNum).vertex(2) = tempVertex

                vertSel = vertexList(PolyNum).vertex(3)
                vertexList(PolyNum).vertex(3) = vertexList(PolyNum).vertex(2)
                vertexList(PolyNum).vertex(2) = vertSel

                tempClr = vertexList(PolyNum).color(3)
                vertexList(PolyNum).color(3) = vertexList(PolyNum).color(2)
                vertexList(PolyNum).color(2) = tempClr
            End If
        Next
    End If

    If numSelectedScenery > 0 Then
        For i = 1 To sceneryCount
            If Scenery(i).selected = 1 Then
                If scaleDiff.X * scaleDiff.Y < 0 Then
                    Scenery(i).rotation = -Scenery(i).rotation
                Else
                    Scenery(i).rotation = Scenery(i).rotation
                End If

                Scenery(i).Translation.X = rCenter.X + (Scenery(i).Translation.X - rCenter.X) * scaleDiff.X
                Scenery(i).Translation.Y = rCenter.Y + (Scenery(i).Translation.Y - rCenter.Y) * scaleDiff.Y

                Scenery(i).screenTr.X = (Scenery(i).Translation.X - scrollCoords(2).X) * zoomFactor
                Scenery(i).screenTr.Y = (Scenery(i).Translation.Y - scrollCoords(2).Y) * zoomFactor

                Scenery(i).Scaling.X = Scenery(i).Scaling.X * scaleDiff.X
                Scenery(i).Scaling.Y = Scenery(i).Scaling.Y * scaleDiff.Y
            End If
        Next
    End If

    If numSelWaypoints > 0 Then
        For i = 1 To waypointCount
            If Waypoints(i).selected Then
                Waypoints(i).X = (rCenter.X + (Waypoints(i).X - rCenter.X) * scaleDiff.X)
                Waypoints(i).Y = (rCenter.Y + (Waypoints(i).Y - rCenter.Y) * scaleDiff.Y)
                If Waypoints(i).wayType(0) Then
                    Waypoints(i).wayType(0) = False
                    Waypoints(i).wayType(1) = True
                ElseIf Waypoints(i).wayType(1) Then
                    Waypoints(i).wayType(0) = True
                    Waypoints(i).wayType(1) = False
                End If
            End If
        Next
    End If

    scaleDiff.X = 1
    scaleDiff.Y = 1

    SaveUndo
    Render
    getInfo

End Sub

Private Sub mnuFlipTexture_Click(Index As Integer)

    Dim i As Integer
    Dim j As Integer
    Dim PolyNum As Integer
    Dim avgMul As Single

    If selectionChanged Then
        SaveUndo
        selectionChanged = False
    End If

    If Index = 0 Then
        scaleDiff.X = -1
    ElseIf Index = 1 Then
        scaleDiff.Y = -1
    End If

    rCenter.X = 0
    rCenter.Y = 0

    avgMul = 1
    If numSelectedPolys > 0 Then
        For i = 1 To numSelectedPolys
            PolyNum = selectedPolys(i)
            For j = 1 To 3
                If vertexList(PolyNum).vertex(j) = 1 Then
                    rCenter.X = rCenter.X * (1 - 1 / avgMul) + Polys(PolyNum).vertex(j).tu / avgMul
                    rCenter.Y = rCenter.Y * (1 - 1 / avgMul) + Polys(PolyNum).vertex(j).tv / avgMul
                    avgMul = avgMul + 1
                End If
            Next
        Next
    End If

    If numSelectedPolys > 0 Then
        For i = 1 To numSelectedPolys
            PolyNum = selectedPolys(i)
            For j = 1 To 3
                If vertexList(PolyNum).vertex(j) = 1 Then
                    Polys(PolyNum).vertex(j).tu = (rCenter.X + (Polys(PolyNum).vertex(j).tu - rCenter.X) * scaleDiff.X)
                    Polys(PolyNum).vertex(j).tv = (rCenter.Y + (Polys(PolyNum).vertex(j).tv - rCenter.Y) * scaleDiff.Y)
                End If
            Next
        Next
    End If

    scaleDiff.X = 1
    scaleDiff.Y = 1

    SaveUndo
    Render
    getInfo

End Sub

Private Sub mnuInvertSel_Click()

    Dim i As Integer
    Dim j As Integer
    Dim addPoly As Boolean

    If showPolys Or showWireframe Or showPoints Then
        numSelectedPolys = 0
        ReDim selectedPolys(mPolyCount)
        For i = 1 To mPolyCount
            addPoly = False
            For j = 1 To 3
                If vertexList(i).vertex(j) = 0 Then
                    vertexList(i).vertex(j) = 1
                Else
                    vertexList(i).vertex(j) = 0
                End If
                If vertexList(i).vertex(j) = 1 Then
                    addPoly = True
                End If
            Next
            If addPoly Then
                numSelectedPolys = numSelectedPolys + 1
                selectedPolys(numSelectedPolys) = i
            End If
        Next
        ReDim Preserve selectedPolys(numSelectedPolys)
    End If

    If showScenery Or showWireframe Or showPoints Then
        numSelectedScenery = 0
        For i = 1 To sceneryCount
            If (Scenery(i).level = 0 And sslBack) Or (Scenery(i).level = 1 And sslMid) Or (Scenery(i).level = 2 And sslFront) Then
                If Scenery(i).selected = 0 Then
                    Scenery(i).selected = 1
                Else
                    Scenery(i).selected = 0
                End If
                If Scenery(i).selected = 1 Then
                    numSelectedScenery = numSelectedScenery + 1
                End If
            End If
        Next
    End If

    If showObjects Then
        numSelSpawns = 0
        For i = 1 To spawnPoints
            If Spawns(i).active = 0 Then
                Spawns(i).active = 1
            Else
                Spawns(i).active = 0
            End If
            If Spawns(i).active = 1 Then
                numSelSpawns = numSelSpawns + 1
            End If
        Next
        numSelColliders = 0
        For i = 1 To colliderCount
            If Colliders(i).active = 0 Then
                Colliders(i).active = 1
            Else
                Colliders(i).active = 0
            End If
            If Colliders(i).active Then
                numSelColliders = numSelColliders + 1
            End If
        Next
    End If

    If showLights Then
        numSelLights = 0
        For i = 1 To lightCount
            If Lights(i).selected = 0 Then
                Lights(i).selected = 1
            Else
                Lights(i).selected = 0
            End If
            If Lights(i).selected Then
                numSelLights = numSelLights + 1
            End If
        Next
    End If

    If showWaypoints Then
        numSelWaypoints = 0
        For i = 1 To waypointCount
            Waypoints(i).selected = Not Waypoints(i).selected
            If Waypoints(i).selected Then
                numSelWaypoints = numSelWaypoints + 1
            End If
        Next
    End If

    getRCenter
    getInfo

    Render

End Sub

Private Sub mnuPaste_Click()

    On Error GoTo ErrorHandler

    If (GetAttr(appPath & "\Temp\copy.PFB") And vbDirectory) = 0 Then
        loadPrefab appPath & "\Temp\copy.PFB"
    End If

ErrorHandler:

End Sub

Private Sub mnuRecent_Click(Index As Integer)

    Dim i As Integer
    Dim Result As VbMsgBoxResult
    Dim theFileName As String
    
    Dim prevMousePointer As Integer

    theFileName = mnuRecent(Index).Caption

    If Len(Dir$(theFileName)) <> 0 And theFileName <> "" Then
        If prompt Then
            Result = MsgBox("Save changes to " & currentFileName & "?", vbYesNoCancel)
            DoEvents
            If Result = vbCancel Then
                Exit Sub
            ElseIf Result = vbYes Then
                mnuSave_Click
                If prompt Then Exit Sub
            End If
        End If
        DoEvents

        prevMousePointer = Me.MousePointer
        Me.MousePointer = vbHourglass

        LoadFile theFileName

        Me.MousePointer = prevMousePointer

        For i = Index To 1 Step -1
            mnuRecent(i).Caption = mnuRecent(i - 1).Caption
        Next
        mnuRecent(0).Caption = theFileName
    ElseIf Len(Dir$(theFileName)) = 0 Then
        MsgBox "File not found: " & theFileName
    End If

End Sub

' put in recent files if it isn't already
Private Sub updateRecent(theFileName As String)

    Dim i As Integer

    mnuRecentFiles.Enabled = True

    For i = 9 To 1 Step -1
        mnuRecent(i).Caption = mnuRecent(i - 1).Caption
        If mnuRecent(i).Caption = "" Then
            mnuRecent(i).Visible = False
        Else
            mnuRecent(i).Visible = True
        End If
    Next
    mnuRecent(0).Caption = theFileName

End Sub

Private Sub mnuResetView_Click()

    zoomFactor = 1
    scrollCoords(2).X = -ScaleWidth / 2
    scrollCoords(2).Y = -ScaleHeight / 2
    Zoom 1

    Render

End Sub

Private Sub mnuRotate_Click(Index As Integer)

    Dim R As Single
    Dim theta As Single
    Dim xDiff As Single
    Dim yDiff As Single
    Dim i As Integer
    Dim j As Integer
    Dim PolyNum As Integer

    If selectionChanged Then
        SaveUndo
        selectionChanged = False
    End If

    If Index = 0 Then
        rDiff = PI
    ElseIf Index = 1 Then
        rDiff = PI / 2
    ElseIf Index = 2 Then
        rDiff = 3 * PI / 2
    End If

    rCenter.X = selRect(0).X + (selRect(2).X - selRect(0).X) / 2
    rCenter.Y = selRect(0).Y + (selRect(2).Y - selRect(0).Y) / 2

    If numSelectedPolys > 0 Then
        For i = 1 To numSelectedPolys
            PolyNum = selectedPolys(i)
            For j = 1 To 3
                If vertexList(PolyNum).vertex(j) = 1 Then
                    xDiff = (PolyCoords(PolyNum).vertex(j).X - rCenter.X)
                    yDiff = (PolyCoords(PolyNum).vertex(j).Y - rCenter.Y)

                    R = Sqr((xDiff) ^ 2 + (yDiff) ^ 2)  ' distance of point from rotation center
                    If xDiff = 0 Then
                        If yDiff > 0 Then
                            theta = PI / 2
                        Else
                            theta = 3 * PI / 2
                        End If
                    ElseIf xDiff > 0 Then
                        theta = Atn(yDiff / xDiff)
                    ElseIf xDiff < 0 Then
                        theta = PI + Atn(yDiff / xDiff)
                    End If
                    theta = theta + rDiff

                    PolyCoords(PolyNum).vertex(j).X = rCenter.X + R * Cos(theta)
                    PolyCoords(PolyNum).vertex(j).Y = rCenter.Y + R * Sin(theta)

                    Polys(PolyNum).vertex(j).X = (PolyCoords(PolyNum).vertex(j).X - scrollCoords(2).X) * zoomFactor
                    Polys(PolyNum).vertex(j).Y = (PolyCoords(PolyNum).vertex(j).Y - scrollCoords(2).Y) * zoomFactor
                End If
            Next
        Next
    End If

    If numSelectedScenery > 0 Then
        For i = 1 To sceneryCount
            If Scenery(i).selected = 1 Then
                xDiff = (Scenery(i).Translation.X - rCenter.X)
                yDiff = (Scenery(i).Translation.Y - rCenter.Y)

                R = Sqr((xDiff) ^ 2 + (yDiff) ^ 2)  ' distance of point from rotation center
                If xDiff = 0 Then
                    If yDiff > 0 Then
                        theta = PI / 2
                    Else
                        theta = 3 * PI / 2
                    End If
                ElseIf xDiff > 0 Then
                    theta = Atn(yDiff / xDiff)
                ElseIf xDiff < 0 Then
                    theta = PI + Atn(yDiff / xDiff)
                End If
                theta = theta + rDiff

                Scenery(i).Translation.X = rCenter.X + R * Cos(theta)
                Scenery(i).Translation.Y = rCenter.Y + R * Sin(theta)

                Scenery(i).screenTr.X = (Scenery(i).Translation.X - scrollCoords(2).X) * zoomFactor
                Scenery(i).screenTr.Y = (Scenery(i).Translation.Y - scrollCoords(2).Y) * zoomFactor

                If scaleDiff.X * scaleDiff.Y < 0 Then
                    Scenery(i).rotation = -(Scenery(i).rotation - rDiff)
                Else
                    Scenery(i).rotation = (Scenery(i).rotation - rDiff)
                End If
            End If
        Next
    End If

    rCenter.X = selRect(0).X
    rCenter.Y = selRect(0).Y
    rDiff = 0

    getRCenter
    getInfo

    SaveUndo
    Render

End Sub

Private Sub mnuRotateTexture_Click(Index As Integer)

    Dim R As Single
    Dim theta As Single
    Dim xDiff As Single
    Dim yDiff As Single
    Dim i As Integer
    Dim j As Integer
    Dim PolyNum As Integer
    Dim avgMul As Single
    Dim texRate As Single

    If selectionChanged Then
        SaveUndo
        selectionChanged = False
    End If

    If Index = 0 Then
        rDiff = PI
    ElseIf Index = 2 Then
        rDiff = PI / 2
    ElseIf Index = 1 Then
        rDiff = 3 * PI / 2
    End If

    texRate = CSng(xTexture) / CSng(yTexture)

    rCenter.X = 0
    rCenter.Y = 0

    avgMul = 1
    If numSelectedPolys > 0 Then
        For i = 1 To numSelectedPolys
            PolyNum = selectedPolys(i)
            For j = 1 To 3
                If vertexList(PolyNum).vertex(j) = 1 Then
                    rCenter.X = rCenter.X * (1 - 1 / avgMul) + Polys(PolyNum).vertex(j).tu * texRate / avgMul
                    rCenter.Y = rCenter.Y * (1 - 1 / avgMul) + Polys(PolyNum).vertex(j).tv / avgMul
                    avgMul = avgMul + 1
                End If
            Next
        Next
    End If

    If numSelectedPolys > 0 Then
        For i = 1 To numSelectedPolys
            PolyNum = selectedPolys(i)
            For j = 1 To 3
                If vertexList(PolyNum).vertex(j) = 1 Then
                    xDiff = (Polys(PolyNum).vertex(j).tu * texRate - rCenter.X)
                    yDiff = (Polys(PolyNum).vertex(j).tv - rCenter.Y)

                    R = Sqr((xDiff) ^ 2 + (yDiff) ^ 2)  ' distance of point from rotation center
                    If xDiff = 0 Then
                        If yDiff > 0 Then
                            theta = PI / 2
                        Else
                            theta = 3 * PI / 2
                        End If
                    ElseIf xDiff > 0 Then
                        theta = Atn(yDiff / xDiff)
                    ElseIf xDiff < 0 Then
                        theta = PI + Atn(yDiff / xDiff)
                    End If
                    theta = theta + rDiff

                    Polys(PolyNum).vertex(j).tu = (rCenter.X + R * Cos(theta)) / texRate
                    Polys(PolyNum).vertex(j).tv = rCenter.Y + R * Sin(theta)
                End If
            Next
        Next
    End If

    rCenter.X = selRect(0).X
    rCenter.Y = selRect(0).Y
    rDiff = 0

    getRCenter
    getInfo

    SaveUndo
    Render

End Sub

Private Sub mnuSetRCenter_Click()

    mnuFixedRCenter.Checked = False
    mnuSetRCenter.Checked = True
    mnuCenterRCenter.Checked = False
    rCenter.X = mouseCoords.X / zoomFactor + scrollCoords(2).X
    rCenter.Y = mouseCoords.Y / zoomFactor + scrollCoords(2).Y

End Sub

Private Sub mnuFixedRCenter_Click()

    mnuFixedRCenter.Checked = True
    mnuSetRCenter.Checked = False
    mnuCenterRCenter.Checked = False
    rCenter.X = Midpoint(selRect(0).X, selRect(2).X)
    rCenter.Y = Midpoint(selRect(0).Y, selRect(2).Y)

End Sub

Private Sub mnuCenterRCenter_Click()

    mnuFixedRCenter.Checked = False
    mnuSetRCenter.Checked = False
    mnuCenterRCenter.Checked = True
    rCenter.X = Midpoint(selRect(0).X, selRect(2).X)
    rCenter.Y = Midpoint(selRect(0).Y, selRect(2).Y)

End Sub

Private Sub mnuShowSceneryLayer_Click(Index As Integer)

    mnuShowSceneryLayer(Index).Checked = Not mnuShowSceneryLayer(Index).Checked

    If Index = 0 Then
        sslBack = mnuShowSceneryLayer(0).Checked
    ElseIf Index = 1 Then
        sslMid = mnuShowSceneryLayer(1).Checked
    ElseIf Index = 2 Then
        sslFront = mnuShowSceneryLayer(2).Checked
    End If

End Sub

Private Sub mnuSnapSelected_Click()

    SnapSelection

End Sub

Private Sub mnuVSelDuplicate_Click()

    mnuDuplicate_Click

End Sub

Private Sub mnuVSelFlip_Click(Index As Integer)

    mnuFlip_Click Index

End Sub

Private Sub mnuVSelPaste_Click()

    mnuPaste_Click

End Sub

Private Sub mnuVSelRotate_Click(Index As Integer)

    mnuRotate_Click Index

End Sub

Private Sub mnuVSelSendBackward_Click()

    mnuSendBackward_Click

End Sub

Private Sub mnuVSelSendToBack_Click()

    mnuSendToBack_Click

End Sub

Private Sub mnuWayType_Click(Index As Integer)

    Dim i As Integer

    mnuWayType(Index).Checked = Not mnuWayType(Index).Checked
    If Index = 0 Then
        mnuWayType(1).Checked = False
    ElseIf Index = 1 Then
        mnuWayType(0).Checked = False
    ElseIf Index = 2 Then
        mnuWayType(3).Checked = False
    ElseIf Index = 3 Then
        mnuWayType(2).Checked = False
    End If

    lblCurrentTool.Caption = frmSoldatMapEditor.ImageList.ListImages(currentFunction + 1).Tag

    For i = 0 To 4
        If mnuWayType(i).Checked Then
            lblCurrentTool.Caption = lblCurrentTool.Caption & " (" & mnuWayType(i).Caption & ")"
        End If
    Next

End Sub

Private Sub picResize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Me.Tag = vbNormal Then
        mIsResizingWindow = True
        picResize.Visible = False
        noRedraw = True

        mInitialWindowWidth = Me.Width
        mInitialWindowHeight = Me.Height

        mMouseStartPosX = X
        mMouseStartPosY = Y
    End If

End Sub

Private Sub picResize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Me.Tag = vbNormal And mIsResizingWindow = True Then
       Dim newWidth As Long
       Dim newHeight As Long

       newWidth = mInitialWindowWidth + (X - mMouseStartPosX) * Screen.TwipsPerPixelX
       newHeight = mInitialWindowHeight + (Y - mMouseStartPosY) * Screen.TwipsPerPixelY

       If newHeight > MIN_FORM_HEIGHT * Screen.TwipsPerPixelY Then
           Me.Height = newHeight
       End If

       If newWidth > MIN_FORM_WIDTH * Screen.TwipsPerPixelX Then
           Me.Width = newWidth
       End If
   End If

End Sub

Private Sub picResize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mIsResizingWindow = False

    If Me.Tag = vbNormal Then
        formHeight = Me.Height / Screen.TwipsPerPixelY
        formWidth = Me.Width / Screen.TwipsPerPixelX

        picResize.Top = formHeight - picResize.Height
        picResize.Left = formWidth - picResize.Width

        picResize.Visible = True
        noRedraw = False
        If mInitialWindowWidth <> Me.Width Or mInitialWindowHeight <> Me.Height Then
            resetDevice
        Else
            Render
        End If
    End If

End Sub

Private Sub picTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Me.Tag = vbMinimized Or Me.Tag = vbNormal And MouseHelper.ButtonIsPressed Then
        If Len(frmDisplay.Tag) <> 0 Then
            QuickHide frmDisplay
        End If
        If Len(frmInfo.Tag) <> 0 Then
            QuickHide frmInfo
        End If
        If Len(frmPalette.Tag) <> 0 Then
            QuickHide frmPalette
        End If
        If Len(frmScenery.Tag) <> 0 Then
            QuickHide frmScenery
        End If
        If Len(frmTexture.Tag) <> 0 Then
            QuickHide frmTexture
        End If
        If Len(frmTools.Tag) <> 0 Then
            QuickHide frmTools
        End If
        If Len(frmWaypoints.Tag) <> 0 Then
            QuickHide frmWaypoints
        End If

        ReleaseCapture
        SendMessage Me.hWnd, WM_NCLBUTTONDOWN, 2, 0&

        If Len(frmDisplay.Tag) <> 0 Then
            QuickMoveAndShow frmDisplay, (frmDisplay.Left + (Me.Left - (formLeft * Screen.TwipsPerPixelX))), (frmDisplay.Top + (Me.Top - (formTop * Screen.TwipsPerPixelY)))
        End If
        If Len(frmInfo.Tag) <> 0 Then
            QuickMoveAndShow frmInfo, (frmInfo.Left + (Me.Left - (formLeft * Screen.TwipsPerPixelX))), (frmInfo.Top + (Me.Top - (formTop * Screen.TwipsPerPixelY)))
        End If
        If Len(frmPalette.Tag) <> 0 Then
            QuickMoveAndShow frmPalette, (frmPalette.Left + (Me.Left - (formLeft * Screen.TwipsPerPixelX))), (frmPalette.Top + (Me.Top - (formTop * Screen.TwipsPerPixelY)))
        End If
        If Len(frmScenery.Tag) <> 0 Then
            QuickMoveAndShow frmScenery, (frmScenery.Left + (Me.Left - (formLeft * Screen.TwipsPerPixelX))), (frmScenery.Top + (Me.Top - (formTop * Screen.TwipsPerPixelY)))
        End If
        If Len(frmTexture.Tag) <> 0 Then
            QuickMoveAndShow frmTexture, (frmTexture.Left + (Me.Left - (formLeft * Screen.TwipsPerPixelX))), (frmTexture.Top + (Me.Top - (formTop * Screen.TwipsPerPixelY)))
        End If
        If Len(frmTools.Tag) <> 0 Then
            QuickMoveAndShow frmTools, (frmTools.Left + (Me.Left - (formLeft * Screen.TwipsPerPixelX))), (frmTools.Top + (Me.Top - (formTop * Screen.TwipsPerPixelY)))
        End If
        If Len(frmWaypoints.Tag) <> 0 Then
            QuickMoveAndShow frmWaypoints, (frmWaypoints.Left + (Me.Left - (formLeft * Screen.TwipsPerPixelX))), (frmWaypoints.Top + (Me.Top - (formTop * Screen.TwipsPerPixelY)))
        End If

        formLeft = Me.Left / Screen.TwipsPerPixelX
        formTop = Me.Top / Screen.TwipsPerPixelY
    End If

End Sub

Private Sub tvwScenery_Expand(ByVal Node As MSComctlLib.Node)

    If Node.Key <> "Master List" And Node.Key <> "In Use" And Node.Key <> "" Then
        mnuScenList.Tag = Node.Key
        mnuScenList.Caption = "Add to " & Node.Key & " list"
    End If

End Sub

Private Sub tvwScenery_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
        If tvwScenery.SelectedItem.FirstSibling <> "In Use" Then
            If tvwScenery.SelectedItem.Parent.Key = "Master List" Then
                If mnuScenList.Tag <> "" Then
                    mnuScenList.Caption = "Add " & tvwScenery.SelectedItem.Text & " to " & mnuScenList.Tag & " List"
                    mnuScenList.Enabled = True
                    mnuScenRemove.Enabled = False
                    PopupMenu mnuScenTree
                End If
            ElseIf tvwScenery.SelectedItem.Parent.Key <> "In Use" Then
                mnuScenRemove.Caption = "Remove " & tvwScenery.SelectedItem.Text & " from List"
                mnuScenList.Enabled = False
                mnuScenRemove.Enabled = True
                PopupMenu mnuScenTree
            End If
        End If
    End If

End Sub

Private Sub mnuScenList_Click()

    Dim i As Integer
    Dim tempNode As Node

    tvwScenery.Nodes.Add mnuScenList.Tag, tvwChild, , tvwScenery.SelectedItem.Text

    Open appPath & "\lists\" & mnuScenList.Tag & ".txt" For Output As #1

        Set tempNode = tvwScenery.Nodes.Item(mnuScenList.Tag).Child
        For i = 1 To tvwScenery.Nodes(mnuScenList.Tag).Children
            Print #1, tempNode.Text
            Set tempNode = tempNode.Next
        Next

    Close #1

End Sub

Private Sub mnuScenRemove_Click()

    Dim i As Integer
    Dim tempNode As Node

    tvwScenery.Nodes.Remove (tvwScenery.SelectedItem.Index)

    Open appPath & "\lists\" & mnuScenList.Tag & ".txt" For Output As #1

        Set tempNode = tvwScenery.Nodes.Item(mnuScenList.Tag).Child
        For i = 1 To tvwScenery.Nodes(mnuScenList.Tag).Children
            Print #1, tempNode.Text
            Set tempNode = tempNode.Next
        Next

    Close #1

End Sub

Public Sub tvwScenery_NodeClick(ByVal Node As MSComctlLib.Node)

    Dim i As Integer
    Dim isInList As Boolean
    Dim token As Long
    Dim tempNode As Node

    On Error GoTo ErrorHandler

    ' if there is no parent
    If tvwScenery.SelectedItem.FirstSibling = "In Use" Then Exit Sub

    If Len(Dir$(frmSoldatMapEditor.soldatDir & "Scenery-gfx\" & tvwScenery.SelectedItem.Text)) = 0 Then
        frmScenery.picScenery.Picture = LoadPicture(appPath & "\skins\" & gfxDir & "\notfound.bmp")
        Exit Sub
    End If

    If tvwScenery.SelectedItem.Parent.Key = "In Use" Then
        currentScenery = tvwScenery.SelectedItem.Text

        token = InitGDIPlus
        frmScenery.picScenery.Picture = LoadPictureGDIPlus(frmSoldatMapEditor.soldatDir & "Scenery-gfx\" & currentScenery, , , RGB(0, 255, 0))
        FreeGDIPlus token

        Set tempNode = tvwScenery.Nodes.Item("In Use").Child

        For i = 1 To (tvwScenery.Nodes.Item("In Use").Children)
            If currentScenery = tempNode.Text Then
                frmSoldatMapEditor.setCurrentScenery i
                frmScenery.lstScenery.ListIndex = i - 1
            End If
            Set tempNode = tempNode.Next
        Next
    Else
        If Len(Dir$(frmSoldatMapEditor.soldatDir & "Scenery-gfx\" & tvwScenery.SelectedItem.Text)) <> 0 Then

            currentScenery = tvwScenery.SelectedItem.Text

            token = InitGDIPlus
            frmScenery.picScenery.Picture = LoadPictureGDIPlus(frmSoldatMapEditor.soldatDir & "Scenery-gfx\" & currentScenery, , , RGB(0, 255, 0))
            FreeGDIPlus token

            ' check if already in list
            Set tempNode = tvwScenery.Nodes.Item("In Use").Child

            For i = 1 To (tvwScenery.Nodes.Item("In Use").Children)
                If currentScenery = tempNode.Text Then
                    isInList = True
                    frmSoldatMapEditor.setCurrentScenery i
                End If
                Set tempNode = tempNode.Next
            Next

            If Not isInList Then
                frmSoldatMapEditor.setCurrentTexture currentScenery
            End If
        End If

        frmScenery.lstScenery.ListIndex = -1
    End If

    Exit Sub

ErrorHandler:

    MsgBox "Error clicking scenery tree" & vbNewLine & Error$

End Sub

Private Function confirmExists(theFileName As String) As Boolean

    Dim tempNode As Node
    Dim i As Integer

    Set tempNode = tvwScenery.Nodes.Item("Master List").Child

    For i = 1 To (tvwScenery.Nodes.Item("Master List").Children)
        If LCase$(theFileName) = LCase$(tempNode.Text) Then
            confirmExists = True
        End If
        Set tempNode = tempNode.Next
    Next

End Function

Private Sub lblZoom_Click()
    txtZoom.Text = gResetZoom * 100 & "%"
    txtZoom_LostFocus
End Sub

Private Sub txtZoom_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        picTitle.SetFocus
    End If

End Sub

Private Sub txtZoom_LostFocus()

    Dim zoomInput As Single

    ' check if valid value was input
    If txtZoom.Text = "" Then
        txtZoom.Text = Int(zoomFactor * 1000 + 0.5) / 10 & "%"
    ElseIf IsNumeric(txtZoom.Text) Then
        zoomInput = txtZoom.Text
    ElseIf IsNumeric(Mid$(txtZoom.Text, 1, Len(txtZoom.Text) - 1)) And Right$(txtZoom.Text, 1) = "%" Then
        zoomInput = Mid$(txtZoom.Text, 1, Len(txtZoom.Text) - 1)
    Else
        txtZoom.Text = Int(zoomFactor * 1000 + 0.5) / 10 & "%"
    End If

    If (zoomInput / 100) >= gMinZoom Or (zoomInput / 100) <= gMaxZoom Then
        Zoom ((zoomInput / 100) / zoomFactor)
        txtZoom.Text = Int(zoomFactor * 1000 + 0.5) / 10 & "%"
    Else
        txtZoom.Text = Int(zoomFactor * 1000 + 0.5) / 10 & "%"
    End If

End Sub

Private Function getZoomDir(zoomDir As Single) As Single

    Dim zoomVal As Single
    Dim i As Integer

    getZoomDir = zoomDir

    zoomVal = gMinZoom
    For i = 1 To 8
        If zoomDir > 1 Then  ' zooming in
            If (zoomFactor) > zoomVal And (zoomFactor) < (zoomVal * 2) Then
                getZoomDir = (zoomVal * 2) / zoomFactor
                Exit For
            End If
        ElseIf zoomDir < 1 Then  ' zooming out
            If (zoomFactor) < zoomVal And (zoomFactor) > (zoomVal * 0.5) Then
                getZoomDir = (zoomVal * 0.5) / zoomFactor
                Exit For
            End If
        End If
        zoomVal = zoomVal * 2
    Next

End Function

Public Sub Zoom(zoomDir As Single)

    Dim i As Integer
    Dim j As Integer
    Dim zoomVal As Single

    If zoomFactor * zoomDir < gMinZoom Or zoomFactor * zoomDir > gMaxZoom Then Exit Sub

    Scenery(0).screenTr.X = Scenery(0).screenTr.X / zoomFactor + scrollCoords(2).X
    Scenery(0).screenTr.Y = Scenery(0).screenTr.Y / zoomFactor + scrollCoords(2).Y

    zoomFactor = zoomFactor * zoomDir

    If zoomDir > 1 Then
        ' zoom to middle
        scrollCoords(2).X = scrollCoords(2).X + Me.ScaleWidth / zoomFactor / (2 / (zoomDir - 1))
        scrollCoords(2).Y = scrollCoords(2).Y + Me.ScaleHeight / zoomFactor / (2 / (zoomDir - 1))
    ElseIf zoomDir < 1 Then
        scrollCoords(2).X = scrollCoords(2).X - Me.ScaleWidth / zoomFactor / (2 / (1 - zoomDir))
        scrollCoords(2).Y = scrollCoords(2).Y - Me.ScaleHeight / zoomFactor / (2 / (1 - zoomDir))
    End If

    For i = 1 To mPolyCount
        For j = 1 To 3
            Polys(i).vertex(j).X = (PolyCoords(i).vertex(j).X - scrollCoords(2).X) * zoomFactor
            Polys(i).vertex(j).Y = (PolyCoords(i).vertex(j).Y - scrollCoords(2).Y) * zoomFactor
        Next
    Next

    For i = 1 To sceneryCount
        Scenery(i).screenTr.X = (Scenery(i).Translation.X - scrollCoords(2).X) * zoomFactor
        Scenery(i).screenTr.Y = (Scenery(i).Translation.Y - scrollCoords(2).Y) * zoomFactor
    Next

    If numVerts > 0 Then
        For j = 1 To 3
            Polys(mPolyCount + 1).vertex(j).X = (PolyCoords(mPolyCount + 1).vertex(j).X - scrollCoords(2).X) * zoomFactor
            Polys(mPolyCount + 1).vertex(j).Y = (PolyCoords(mPolyCount + 1).vertex(j).Y - scrollCoords(2).Y) * zoomFactor
        Next
    End If

    For i = 1 To 4
        bgPolys(i).X = (bgPolyCoords(i).X - scrollCoords(2).X) * zoomFactor
        bgPolys(i).Y = (bgPolyCoords(i).Y - scrollCoords(2).Y) * zoomFactor
    Next

    Scenery(0).screenTr.X = (Scenery(0).screenTr.X - scrollCoords(2).X) * zoomFactor
    Scenery(0).screenTr.Y = (Scenery(0).screenTr.Y - scrollCoords(2).Y) * zoomFactor

    selectedCoords(1).X = 0
    selectedCoords(1).Y = 0
    selectedCoords(2).X = 0
    selectedCoords(2).Y = 0

    txtZoom.Text = Int(zoomFactor * 1000 + 0.5) / 10 & "%"

    Render

    If circleOn Then
        Render
    End If

End Sub

Public Sub zoomScroll(zoomDir As Single, ByVal X As Integer, ByVal Y As Integer)

    Dim i As Integer
    Dim j As Integer

    If (zoomFactor * zoomDir < gMinZoom) And zoomFactor > gMinZoom Then
        zoomDir = gMinZoom / zoomFactor
    ElseIf zoomFactor * zoomDir > gMaxZoom And zoomFactor < gMaxZoom Then
        zoomDir = gMaxZoom / zoomFactor
    End If

    If zoomFactor * zoomDir < gMinZoom Or zoomFactor * zoomDir > gMaxZoom Then Exit Sub

    Scenery(0).screenTr.X = Scenery(0).screenTr.X / zoomFactor + scrollCoords(2).X
    Scenery(0).screenTr.Y = Scenery(0).screenTr.Y / zoomFactor + scrollCoords(2).Y

    selectedCoords(1).X = selectedCoords(1).X / zoomFactor + scrollCoords(2).X
    selectedCoords(1).Y = selectedCoords(1).Y / zoomFactor + scrollCoords(2).Y

    zoomFactor = (zoomFactor * zoomDir)

    If zoomDir > 1 Then
        scrollCoords(2).X = scrollCoords(2).X + X / zoomFactor / ((2 / (zoomDir - 1)) / 2)
        scrollCoords(2).Y = scrollCoords(2).Y + Y / zoomFactor / ((2 / (zoomDir - 1)) / 2)
    ElseIf zoomDir < 1 Then
        scrollCoords(2).X = scrollCoords(2).X - Me.ScaleWidth / zoomFactor / (2 / (1 - zoomDir))
        scrollCoords(2).Y = scrollCoords(2).Y - Me.ScaleHeight / zoomFactor / (2 / (1 - zoomDir))
    End If

    For i = 1 To mPolyCount
        For j = 1 To 3
            Polys(i).vertex(j).X = (PolyCoords(i).vertex(j).X - scrollCoords(2).X) * zoomFactor
            Polys(i).vertex(j).Y = (PolyCoords(i).vertex(j).Y - scrollCoords(2).Y) * zoomFactor
        Next
    Next

    For i = 1 To sceneryCount
        Scenery(i).screenTr.X = (Scenery(i).Translation.X - scrollCoords(2).X) * zoomFactor
        Scenery(i).screenTr.Y = (Scenery(i).Translation.Y - scrollCoords(2).Y) * zoomFactor
    Next

    If numVerts > 0 Then
        For j = 1 To 3
            Polys(mPolyCount + 1).vertex(j).X = (PolyCoords(mPolyCount + 1).vertex(j).X - scrollCoords(2).X) * zoomFactor
            Polys(mPolyCount + 1).vertex(j).Y = (PolyCoords(mPolyCount + 1).vertex(j).Y - scrollCoords(2).Y) * zoomFactor
        Next
    End If

    For i = 1 To 4
        bgPolys(i).X = (bgPolyCoords(i).X - scrollCoords(2).X) * zoomFactor
        bgPolys(i).Y = (bgPolyCoords(i).Y - scrollCoords(2).Y) * zoomFactor
    Next

    Scenery(0).screenTr.X = (Scenery(0).screenTr.X - scrollCoords(2).X) * zoomFactor
    Scenery(0).screenTr.Y = (Scenery(0).screenTr.Y - scrollCoords(2).Y) * zoomFactor

    selectedCoords(1).X = (selectedCoords(1).X - scrollCoords(2).X) * zoomFactor
    selectedCoords(1).Y = (selectedCoords(1).Y - scrollCoords(2).Y) * zoomFactor

    txtZoom.Text = Int(zoomFactor * 1000 + 0.5) / 10 & "%"

    Render

End Sub

Private Function pointInPoly(ByVal X As Single, ByVal Y As Single, ByVal i As Integer) As Boolean

    Dim xDist As Single
    Dim yDist As Single
    Dim xDiff As Single
    Dim yDiff As Single
    Dim length As Single
    Dim D As Single

    pointInPoly = True

    xDist = X - Polys(i).vertex(1).X
    yDist = Y - Polys(i).vertex(1).Y
    xDiff = Polys(i).vertex(2).X - Polys(i).vertex(1).X
    yDiff = Polys(i).vertex(1).Y - Polys(i).vertex(2).Y
    If xDiff = 0 And yDiff = 0 Then
        length = 1
    Else
        length = Sqr(xDiff ^ 2 + yDiff ^ 2)
    End If
    D = (yDiff / length) * xDist + (xDiff / length) * yDist
    If D < 0 Then pointInPoly = False

    xDist = X - Polys(i).vertex(2).X
    yDist = Y - Polys(i).vertex(2).Y
    xDiff = Polys(i).vertex(3).X - Polys(i).vertex(2).X
    yDiff = Polys(i).vertex(2).Y - Polys(i).vertex(3).Y
    If xDiff = 0 And yDiff = 0 Then
        length = 1
    Else
        length = Sqr(xDiff ^ 2 + yDiff ^ 2)
    End If
    D = (yDiff / length) * xDist + (xDiff / length) * yDist
    If D < 0 Then pointInPoly = False

    xDist = X - Polys(i).vertex(3).X
    yDist = Y - Polys(i).vertex(3).Y
    xDiff = Polys(i).vertex(1).X - Polys(i).vertex(3).X
    yDiff = Polys(i).vertex(3).Y - Polys(i).vertex(1).Y
    If xDiff = 0 And yDiff = 0 Then
        length = 1
    Else
        length = Sqr(xDiff ^ 2 + yDiff ^ 2)
    End If
    D = (yDiff / length) * xDist + (xDiff / length) * yDist
    If D < 0 Then pointInPoly = False

End Function

Private Function isCW(ByVal i As Integer) As Boolean

    Dim xVal As Single
    Dim yVal As Single

    xVal = Midpoint(Polys(i).vertex(1).X, Midpoint(Polys(i).vertex(2).X, Polys(i).vertex(3).X))
    yVal = Midpoint(Polys(i).vertex(1).Y, Midpoint(Polys(i).vertex(2).Y, Polys(i).vertex(3).Y))

    isCW = pointInPoly(xVal, yVal, i)

End Function

Public Sub setDispOptions(layerNum As Integer, value As Boolean)

    If layerNum = 0 Then
        showBG = value
    ElseIf layerNum = 1 Then
        showPolys = value
    ElseIf layerNum = 2 Then
        showTexture = value
    ElseIf layerNum = 3 Then
        showWireframe = value
    ElseIf layerNum = 4 Then
        showPoints = value
    ElseIf layerNum = 5 Then
        showScenery = value
    ElseIf layerNum = 6 Then
        showObjects = value
    ElseIf layerNum = 7 Then
        showWaypoints = value
    ElseIf layerNum = 8 Then
        showGrid = value
        mnuGrid.Checked = value
    ElseIf layerNum = 9 Then
        showLights = value
        setLightsMode showLights
    ElseIf layerNum = 10 Then
        showSketch = value
    End If

    Render

End Sub

Private Sub setLightsMode(lightsOn As Boolean)

    Dim i As Integer
    Dim j As Integer

    If Not lightsOn Then
        For i = 1 To mPolyCount
            For j = 1 To 3
                Polys(i).vertex(j).color = ARGB(getAlpha(Polys(i).vertex(j).color), RGB(vertexList(i).color(j).blue, vertexList(i).color(j).green, vertexList(i).color(j).red))
            Next
        Next
    Else
        applyLights
    End If

End Sub

Public Sub setColorMode(ByVal clrVal As Byte)

    colorMode = clrVal

End Sub

Public Sub setCurrentTool(ByVal Index As Integer)

    Dim i As Integer

    currentTool = Index
    currentFunction = Index
    If currentTool = TOOL_CREATE And mnuQuad.Checked Then
        currentFunction = TOOL_QUAD
    ElseIf currentTool <> TOOL_SCENERY Then
        frmSoldatMapEditor.tvwScenery.Visible = False
    End If

    circleOn = False

    If numVerts > 0 And currentTool <> TOOL_CREATE Then  ' abort poly creation
        numVerts = 0
    ElseIf numCorners > 0 And currentTool <> TOOL_SCENERY Then
        numCorners = 0
    ElseIf currentWaypoint > 0 And currentTool <> TOOL_WAYPOINT Then
        currentWaypoint = 0
    End If
    toolAction = False

    If currentTool = TOOL_PSELECT And numSelectedPolys > 0 Then
        For i = 1 To numSelectedPolys
            vertexList(selectedPolys(i)).vertex(1) = 1
            vertexList(selectedPolys(i)).vertex(2) = 1
            vertexList(selectedPolys(i)).vertex(3) = 1
        Next
        getRCenter
    ElseIf currentTool = TOOL_MOVE Then
        If numSelectedPolys = 0 And numSelectedScenery = 1 Then
            frmInfo.mnuProp_Click 1
        Else
            frmInfo.mnuProp_Click 2
        End If
    ElseIf currentTool = TOOL_TEXTURE Then
        frmInfo.mnuProp_Click 3
    ElseIf currentTool = TOOL_VCOLOR Then
        circleOn = True
    ElseIf currentTool = TOOL_DEPTHMAP Then
        circleOn = True
    End If

    SetCursor currentFunction + 1
    lblCurrentTool.Caption = frmSoldatMapEditor.ImageList.ListImages(currentFunction + 1).Tag

    If currentTool = TOOL_CREATE Then
        lblCurrentTool.Caption = lblCurrentTool.Caption & " (" & mnuPolyType(polyType).Caption & ")"
    ElseIf currentTool = TOOL_WAYPOINT Then
        For i = 0 To 4
            If mnuWayType(i).Checked Then
                lblCurrentTool.Caption = lblCurrentTool.Caption & " (" & mnuWayType(i).Caption & ")"
            End If
        Next
    End If

    Render

End Sub

Public Function setTempTool(toolNum As Byte) As Byte

    setTempTool = currentTool
    currentTool = toolNum

End Function

Public Sub setMapTexture(texturePath As String)

    On Error GoTo ErrorHandler

    Set mapTexture = D3DX.CreateTextureFromFileEx(D3DDevice, frmSoldatMapEditor.soldatDir & "textures\" & texturePath, D3DX_DEFAULT, D3DX_DEFAULT, _
            D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_TRIANGLE, _
            D3DX_FILTER_TRIANGLE, COLOR_KEY, imageInfo, ByVal 0)

    gTextureFile = texturePath

    xTexture = imageInfo.Width
    yTexture = imageInfo.Height

    frmInfo.lblDimensions.Caption = "Dimensions: " & xTexture & " x " & yTexture
    frmInfo.txtQuadX(0).Text = 0
    frmInfo.txtQuadY(0).Text = 0
    frmInfo.txtQuadX(1).Text = xTexture
    frmInfo.txtQuadY(1).Text = yTexture

    Render

ErrorHandler:

End Sub

' set gpolyclr when rgb modified
Public Sub setPolyColor(Index As Integer, value As Byte)

    If Index = 0 Then
        gPolyClr.red = value
    ElseIf Index = 1 Then
        gPolyClr.green = value
    ElseIf Index = 2 Then
        gPolyClr.blue = value
    ElseIf Index = 3 Then
        opacity = value / 100
    End If
    If numVerts > 0 And (currentFunction = TOOL_CREATE Or currentFunction = TOOL_QUAD) Then
        Polys(mPolyCount + 1).vertex(numVerts + 1).color = ARGB(255 * opacity, RGB(gPolyClr.blue, gPolyClr.green, gPolyClr.red))
    End If
    Scenery(0).alpha = opacity * 255
    Scenery(0).color = ARGB(opacity * 255, RGB(gPolyClr.blue, gPolyClr.green, gPolyClr.red))

End Sub

' set gpolyclr when palette clicked
Public Sub setPaletteColor(red As Byte, green As Byte, blue As Byte)

    gPolyClr.red = red
    gPolyClr.green = green
    gPolyClr.blue = blue
    If numVerts > 0 And (currentFunction = TOOL_CREATE Or currentFunction = TOOL_QUAD) Then
        Polys(mPolyCount + 1).vertex(numVerts + 1).color = ARGB(255 * opacity, RGB(gPolyClr.blue, gPolyClr.green, gPolyClr.red))
    End If
    Scenery(0).alpha = opacity * 255
    Scenery(0).color = ARGB(Scenery(0).alpha, RGB(gPolyClr.blue, gPolyClr.green, gPolyClr.red))

End Sub

Public Sub setBlendMode(Index As Integer)

    blendMode = Index

End Sub

Public Sub getOptions()

    Dim i As Integer

    frmMap.txtDesc = mapTitle
    frmMap.txtJet = Options.StartJet
    frmMap.cboGrenades.ListIndex = Options.GrenadePacks
    frmMap.cboMedikits.ListIndex = Options.Medikits
    frmMap.cboSteps.ListIndex = Options.Steps
    frmMap.cboWeather.ListIndex = Options.Weather
    frmMap.picBackClr(0).BackColor = RGB(bgColors(1).red, bgColors(1).green, bgColors(1).blue)
    frmMap.picBackClr(1).BackColor = RGB(bgColors(2).red, bgColors(2).green, bgColors(2).blue)

    For i = 0 To frmMap.cboTexture.ListCount - 1
        If frmMap.cboTexture.List(i) = gTextureFile Then
            frmMap.cboTexture.ListIndex = i
        End If
    Next

End Sub

Public Sub setOptions()

    Options.GrenadePacks = frmMap.cboGrenades.ListIndex
    Options.Medikits = frmMap.cboMedikits.ListIndex
    Options.StartJet = frmMap.txtJet.Text
    Options.Steps = frmMap.cboSteps.ListIndex
    Options.Weather = frmMap.cboWeather.ListIndex
    Options.BackgroundColor = ARGB(255, RGB(bgColors(1).blue, bgColors(1).green, bgColors(1).red))
    Options.BackgroundColor = ARGB(255, RGB(bgColors(2).blue, bgColors(2).green, bgColors(2).red))

    mapTitle = frmMap.txtDesc.Text

End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim Result As VbMsgBoxResult
    Dim temp As String
    
    Dim prevMousePointer As Integer

    temp = Data.Files.Item(1)
    If Right(temp, 1) = """" Then
        temp = Left(temp, Len(temp) - 1)
        temp = Right(temp, Len(temp) - 1)
    End If

    If LCase$(Right(temp, 4)) = ".pms" Then
        If prompt Then
            Result = MsgBox("Save changes to " & currentFileName & "?", vbYesNoCancel)
            DoEvents
            If Result = vbCancel Then
                Exit Sub
            ElseIf Result = vbYes Then
                mnuSave_Click
                If prompt Then Exit Sub
            End If
        End If
        DoEvents

        recentFiles Data.Files.Item(1)

        prevMousePointer = Me.MousePointer
        Me.MousePointer = vbHourglass

        LoadFile Data.Files.Item(1)

        Me.MousePointer = prevMousePointer
    End If

End Sub

Public Sub Form_Paint()

    Render

End Sub

Public Sub Terminate()  ' You are on the way to destruction.

    Dim Result As VbMsgBoxResult

    On Error GoTo ErrorHandler

    If prompt Then
        Result = MsgBox("Save changes to " & currentFileName & "?", vbYesNoCancel)
        DoEvents
        If Result = vbCancel Then
            Exit Sub
        ElseIf Result = vbYes Then
            mnuSave_Click
            If prompt Then Exit Sub
        End If
    End If
    DoEvents

    saveSettings

    Set mapTexture = Nothing
    Set particleTexture = Nothing
    Set patternTexture = Nothing
    Set sketchTexture = Nothing
    Set objectsTexture = Nothing
    Set lineTexture = Nothing
    Set pathTexture = Nothing
    Set rCenterTexture = Nothing

    ReDim SceneryTextures(0)
    Set SceneryTextures(0).Texture = Nothing

    DIDevice.Unacquire

    If hEvent <> 0 Then DX.DestroyEvent hEvent

    Set D3DDevice = Nothing
    Set DIDevice = Nothing
    Set DI = Nothing
    Set D3D = Nothing
    Set DX = Nothing

    Unload Me
    End

    Exit Sub

ErrorHandler:

    MsgBox "Error terminating" & vbNewLine & Error$

End Sub

Private Sub Form_Resize()

    picHelp.Left = frmSoldatMapEditor.ScaleWidth - 80
    picMinimize.Left = frmSoldatMapEditor.ScaleWidth - 48
    picMaximize.Left = frmSoldatMapEditor.ScaleWidth - 32
    picExit.Left = frmSoldatMapEditor.ScaleWidth - 16

    picProgress.Left = frmSoldatMapEditor.ScaleWidth - 136

End Sub

Private Sub MouseHelper_MouseWheel(ctrl As Variant, Direction As MBMouseHelper.mbDirectionConstants, Button As Long, Shift As Long, Cancel As Boolean)

    Dim zoomVal As Single

    If Direction = mbBackward Then
        zoomScroll 0.8, mouseCoords.X, mouseCoords.Y
    ElseIf Direction = mbForward Then
        zoomScroll 1.25, mouseCoords.X, mouseCoords.Y
    End If

End Sub

Public Sub setPreferences()

    inc = (gridSpacing / gridDivisions)
    tvwScenery.Height = formHeight - 41 - 20
    resetDevice
    Render

End Sub

Public Function setBGColor(Index As Integer) As Long

    frmColor.InitColor bgColors(Index).red, bgColors(Index).green, bgColors(Index).blue
    frmColor.Show 1
    bgColors(Index).red = frmColor.red
    bgColors(Index).green = frmColor.green
    bgColors(Index).blue = frmColor.blue

    bgPolys(1).color = RGB(bgColors(1).blue, bgColors(1).green, bgColors(1).red)
    bgPolys(2).color = RGB(bgColors(2).blue, bgColors(2).green, bgColors(2).red)
    bgPolys(3).color = RGB(bgColors(1).blue, bgColors(1).green, bgColors(1).red)
    bgPolys(4).color = RGB(bgColors(2).blue, bgColors(2).green, bgColors(2).red)

    setBGColor = RGB(bgColors(Index).red, bgColors(Index).green, bgColors(Index).blue)

    Render

End Function

Public Sub setLightColor()

    Dim i As Integer
    Dim Index As Integer

    For i = 1 To lightCount
        If Lights(i).selected = 1 Then
            Index = i
            Exit For
        End If
    Next

    frmColor.InitColor Lights(Index).color.red, Lights(Index).color.green, Lights(Index).color.blue
    frmColor.Show 1

    For i = 1 To lightCount
        If Lights(i).selected = 1 Then
            Lights(i).color.red = frmColor.red
            Lights(i).color.green = frmColor.green
            Lights(i).color.blue = frmColor.blue
        End If
    Next

    frmInfo.picLight.BackColor = RGB(frmColor.red, frmColor.green, frmColor.blue)

    applyLights

End Sub

Public Sub setRadius(R As Integer)

    Dim i As Integer

    clrRadius = R
    Colliders(0).radius = R

    If numSelColliders > 0 Then
        For i = 1 To colliderCount
            If Colliders(i).active Then
                Colliders(i).radius = R
            End If
        Next
        Render
    End If

End Sub

Public Function setWayType(Index As Integer, tehValue As Boolean) As Boolean

    If numSelWaypoints = 0 Then
        setWayType = False
        Exit Function
    End If

    Dim i As Integer

    For i = 1 To waypointCount
        If Waypoints(i).selected Then
            Waypoints(i).wayType(Index) = tehValue
            If Index = 0 Then
                Waypoints(i).wayType(1) = False
            ElseIf Index = 1 Then
                Waypoints(i).wayType(0) = False
            ElseIf Index = 2 Then
                Waypoints(i).wayType(3) = False
            ElseIf Index = 3 Then
                Waypoints(i).wayType(2) = False
            End If
        End If
    Next

    setWayType = True

    Render

End Function

Public Sub setPathNum(tehValue As Byte)

    Dim i As Integer

    For i = 1 To waypointCount
        If Waypoints(i).selected Then
            Waypoints(i).pathNum = tehValue
        End If
    Next

    Render

End Sub

Public Function setSpecial(tehValue As Byte) As Boolean

    Dim i As Integer

    If numSelWaypoints = 0 Then
        setSpecial = False
        Exit Function
    End If

    For i = 1 To waypointCount
        If Waypoints(i).selected Then
            Waypoints(i).special = tehValue
        End If
    Next

    setSpecial = True

End Function

Public Sub setShowPaths()

    Render

End Sub

Public Sub ClearUnused()

    Dim i As Integer
    Dim j As Integer
    Dim doesExist As Boolean
    Dim offset As Integer
    Dim numDeleted As Integer

    On Error GoTo ErrorHandler

    offset = 1
    For i = 1 To sceneryElements
        For j = 1 To sceneryCount  ' check if exists
            If Scenery(j).Style = i Then
                doesExist = True
                Exit For
            End If
        Next
        ' check if duplicate
        For j = 0 To offset - 2
            If frmScenery.lstScenery.List(j) = frmScenery.lstScenery.List(offset - 1) Then
                doesExist = False
                Exit For
            End If
        Next
        SceneryTextures(offset) = SceneryTextures(i)
        If doesExist Then  ' if does not exist, will get overwritten next time
            offset = offset + 1
        Else
            numDeleted = numDeleted + 1
            frmScenery.lstScenery.RemoveItem offset - 1
        End If
        For j = 1 To sceneryCount
            If Scenery(j).Style = i Then
                Scenery(j).Style = Scenery(j).Style - numDeleted
            End If
        Next
        doesExist = False
    Next

    If numDeleted > 0 Then
        Scenery(0).Style = 0

        sceneryElements = sceneryElements - numDeleted
        ReDim Preserve SceneryTextures(sceneryElements)

        tvwScenery.Nodes.Remove "In Use"
        tvwScenery.Nodes.Add "Master List", tvwFirst, "In Use", "In Use"
        For i = 0 To frmScenery.lstScenery.ListCount - 1
            tvwScenery.Nodes.Add "In Use", tvwChild, frmScenery.lstScenery.List(i), frmScenery.lstScenery.List(i)
        Next
    End If

    numUndo = 0

    Exit Sub

ErrorHandler:

    MsgBox "Error clearing unused scenery" & vbNewLine & Error$

End Sub

Public Sub saveWindow(sectionName As String, window As Form, collapsed As Boolean, isNewFile As Boolean, Optional theFileName As String = "current.ini")

    Dim leftVal As Integer
    Dim topVal As Integer
    Dim iniString As String
    Dim sNull As String
    sNull = Chr$(0)

    leftVal = window.Left / Screen.TwipsPerPixelX
    topVal = window.Top / Screen.TwipsPerPixelY

    iniString = _
        "Visible=" & window.Visible & sNull & _
        "Left=" & leftVal & sNull & _
        "Top=" & topVal & sNull & _
        "Collapsed=" & collapsed & sNull & _
        "Snapped=" & IIf(Len(window.Tag) > 0, "True", "False") & _
        IIf(isNewFile, vbNewLine, "") & sNull & sNull

    saveSection sectionName, iniString, appPath & "\workspace\" & theFileName

End Sub

Private Function getNextValue(sectionString As String, ByRef eIndex As Integer) As String

    Dim nIndex As Integer

    eIndex = InStr(eIndex, sectionString, "=") + 1
    nIndex = InStr(eIndex, sectionString, vbNullChar)
    getNextValue = Mid$(sectionString, eIndex, nIndex)

End Function

Public Sub loadColors()

    On Error GoTo ErrorHandler

    bgColor = CLng("&H" + loadString("GUIColors", "Background", appPath & "\skins\" & gfxDir & "\colors.ini"))
    lblBackClr = CLng("&H" + loadString("GUIColors", "LabelBack", appPath & "\skins\" & gfxDir & "\colors.ini"))
    lblTextClr = CLng("&H" + loadString("GUIColors", "LabelText", appPath & "\skins\" & gfxDir & "\colors.ini"))
    txtBackClr = CLng("&H" + loadString("GUIColors", "TextBoxBack", appPath & "\skins\" & gfxDir & "\colors.ini"))
    txtTextClr = CLng("&H" + loadString("GUIColors", "TextBoxText", appPath & "\skins\" & gfxDir & "\colors.ini"))
    frameClr = CLng("&H" + loadString("GUIColors", "Frame", appPath & "\skins\" & gfxDir & "\colors.ini"))
    font1 = loadString("GUIColors", "font1", appPath & "\skins\" & gfxDir & "\colors.ini", 40)
    font2 = loadString("GUIColors", "font2", appPath & "\skins\" & gfxDir & "\colors.ini", 40)

    If font1 = "" Then font1 = "Arial"
    If font2 = "" Then font2 = "Arial"

    Exit Sub

ErrorHandler:

    MsgBox "Error loading colors" & vbNewLine & Error$

End Sub

Private Sub mnuExit_Click()

    Terminate

End Sub

Private Sub mnuNew_Click()

    Dim Result As VbMsgBoxResult

    If prompt Then
        Result = MsgBox("Save changes to " & currentFileName & "?", vbYesNoCancel)
        DoEvents
        If Result = vbCancel Then
            Exit Sub
        ElseIf Result = vbYes Then
            mnuSave_Click
            If prompt Then Exit Sub
        End If
    End If
    newMap

End Sub

Private Sub mnuOpen_Click()

    On Error GoTo ErrorHandler

    Dim Result As VbMsgBoxResult
    
    Dim prevMousePointer As Integer

    If prompt Then
        Result = MsgBox("Save changes to " & currentFileName & "?", vbYesNoCancel)
        DoEvents
        If Result = vbCancel Then
            Exit Sub
        ElseIf Result = vbYes Then
            mnuSave_Click
            If prompt Then Exit Sub
        End If
    End If
    DoEvents

    frmSoldatMapEditor.commonDialog.Filter = "Map File (*.pms)|*.pms"
    commonDialog.InitDir = uncompDir
    commonDialog.FileName = uncompDir & currentFileName
    frmSoldatMapEditor.commonDialog.DialogTitle = "Load Map"
    commonDialog.ShowOpen

    If commonDialog.FileName <> "" Then
        prompt = False
        recentFiles commonDialog.FileName
        mPolyCount = 0
        numSelectedPolys = 0
        ReDim selectedPolys(0)
        ReDim vertexList(0)
        ReDim Polys(0)
        selectedCoords(1).X = 0
        selectedCoords(1).Y = 0
        selectedCoords(2).X = 0
        selectedCoords(2).Y = 0

        prevMousePointer = Me.MousePointer
        Me.MousePointer = vbHourglass

        LoadFile commonDialog.FileName

        Me.MousePointer = prevMousePointer
    End If

    RegainFocus

    Exit Sub

ErrorHandler:

    If Error$ <> "Cancel was selected." Then
        MsgBox "Error opening file" & vbNewLine & Error$
    End If
    RegainFocus

End Sub

Private Sub mnuOpenCompiled_Click()

    On Error GoTo ErrorHandler

    Dim Result As VbMsgBoxResult

    Dim prevMousePointer As Integer
    

    If prompt Then
        Result = MsgBox("Save changes to " & currentFileName & "?", vbYesNoCancel)
        DoEvents
        If Result = vbCancel Then
            Exit Sub
        ElseIf Result = vbYes Then
            mnuSave_Click
            If prompt Then Exit Sub
        End If
    End If
    DoEvents

    frmSoldatMapEditor.commonDialog.Filter = "Map File (*.pms)|*.pms"
    commonDialog.InitDir = soldatDir & "Maps\"
    commonDialog.FileName = soldatDir & "Maps\" & currentFileName
    frmSoldatMapEditor.commonDialog.DialogTitle = "Load Compiled Map"
    commonDialog.ShowOpen

    If commonDialog.FileName <> "" Then
        prompt = False
        recentFiles commonDialog.FileName
        mPolyCount = 0
        numSelectedPolys = 0
        ReDim selectedPolys(0)
        ReDim vertexList(0)
        ReDim Polys(0)
        selectedCoords(1).X = 0
        selectedCoords(1).Y = 0
        selectedCoords(2).X = 0
        selectedCoords(2).Y = 0

        prevMousePointer = Me.MousePointer
        Me.MousePointer = vbHourglass

        LoadFile commonDialog.FileName

        Me.MousePointer = prevMousePointer
    End If

    RegainFocus

    Exit Sub

ErrorHandler:

    If Error$ <> "Cancel was selected." Then
        MsgBox "Error opening compiled map" & vbNewLine & Error$
    End If
    RegainFocus

End Sub

Private Sub mnuSave_Click()

    Dim i As Integer

    On Error GoTo ErrorHandler

    frmSoldatMapEditor.commonDialog.Filter = "Map File (*.pms)|*.pms"
    frmSoldatMapEditor.commonDialog.DialogTitle = "Save Map"
    commonDialog.FileName = uncompDir & currentFileName
    commonDialog.InitDir = uncompDir

    If lblFileName.Caption = "Untitled.pms" Then
        commonDialog.ShowSave

        If commonDialog.FileName <> "" Then
            recentFiles commonDialog.FileName

            DoEvents
            SaveFile commonDialog.FileName
            prompt = False
        End If
    Else
        SaveFile commonDialog.FileName
        prompt = False
    End If

    RegainFocus

    Exit Sub

ErrorHandler:

    If Error$ <> "Cancel was selected." Then
        MsgBox "Error saving file" & vbNewLine & Error$
    End If
    RegainFocus

End Sub

Private Sub mnuSaveAs_Click()

    Dim i As Integer

    On Error GoTo ErrorHandler

    frmSoldatMapEditor.commonDialog.Filter = "Map File (*.pms)|*.pms"
    commonDialog.InitDir = appPath & "\Maps\"
    commonDialog.FileName = appPath & "\Maps\" & currentFileName
    frmSoldatMapEditor.commonDialog.DialogTitle = "Save Map"
    commonDialog.ShowSave

    If commonDialog.FileName <> "" Then
        recentFiles commonDialog.FileName

        DoEvents
        SaveFile commonDialog.FileName
        prompt = False
    End If

    RegainFocus

    Exit Sub

ErrorHandler:

    If Error$ <> "Cancel was selected." Then
        MsgBox "Error saving as" & vbNewLine & Error$
    End If
    RegainFocus

End Sub

Private Sub mnuCompile_Click()

    Dim i As Integer
    Dim length As Integer

    On Error GoTo ErrorHandler

    frmSoldatMapEditor.commonDialog.Filter = "Map File (*.pms)|*.pms"
    commonDialog.InitDir = frmSoldatMapEditor.soldatDir & "Maps\"
    commonDialog.FileName = frmSoldatMapEditor.soldatDir & "Maps\" & currentFileName
    frmSoldatMapEditor.commonDialog.DialogTitle = "Compile to pms"

    If lblFileName.Caption = "Untitled.pms" Then
        commonDialog.ShowSave
        DoEvents

        If commonDialog.FileName <> "" Then
            SaveAndCompile commonDialog.FileName
            prompt = False

            For i = 1 To Len(commonDialog.FileName)
                If Mid(commonDialog.FileName, i, 1) = "\" Then
                    length = i + 1
                End If
            Next
            lastCompiled = Mid(commonDialog.FileName, length, Len(commonDialog.FileName) - length - 3)
        End If
    Else
        SaveAndCompile commonDialog.FileName
        prompt = False

        For i = 1 To Len(commonDialog.FileName)
            If Mid(commonDialog.FileName, i, 1) = "\" Then
                length = i + 1
            End If
        Next
        lastCompiled = Mid(commonDialog.FileName, length, Len(commonDialog.FileName) - length - 3)
    End If

    RegainFocus

    Exit Sub

ErrorHandler:

    If Error$ <> "Cancel was selected." Then
        MsgBox "Error compiling map" & vbNewLine & Error$
    End If
    RegainFocus

End Sub

Private Sub mnuCompileAs_Click()

    Dim i As Integer
    Dim length As Integer

    On Error GoTo ErrorHandler

    frmSoldatMapEditor.commonDialog.Filter = "Map File (*.pms)|*.pms"
    commonDialog.InitDir = frmSoldatMapEditor.soldatDir & "Maps\"
    commonDialog.FileName = frmSoldatMapEditor.soldatDir & "Maps\" & currentFileName
    frmSoldatMapEditor.commonDialog.DialogTitle = "Compile to pms"
    commonDialog.ShowSave

    If commonDialog.FileName <> "" Then
        DoEvents
        SaveAndCompile commonDialog.FileName
        prompt = False

        For i = 1 To Len(commonDialog.FileName)
            If Mid(commonDialog.FileName, i, 1) = "\" Then
                length = i + 1
            End If
        Next
        lastCompiled = Mid(commonDialog.FileName, length, Len(commonDialog.FileName) - length - 3)
    End If

    RegainFocus

    Exit Sub

ErrorHandler:

    If Error$ <> "Cancel was selected." Then
        MsgBox "Error compiling map" & vbNewLine & Error$
    End If
    RegainFocus

End Sub

Private Function recentFiles(theFileName As String) As Boolean

    Dim i As Integer
    Dim inRecent As Boolean
    Dim Index As Integer

    For i = 0 To 9
        If mnuRecent(i).Caption = theFileName Then
            inRecent = True
            Index = i
        End If
    Next
    If Not inRecent Then
        updateRecent theFileName
    Else
        For i = Index To 1 Step -1
            mnuRecent(i).Caption = mnuRecent(i - 1).Caption
        Next
        mnuRecent(0).Caption = theFileName
    End If

End Function

Private Sub mnuExport_Click()

    On Error GoTo ErrorHandler

    frmSoldatMapEditor.commonDialog.Filter = "Prefab (*.pfb)|*.pfb"
    commonDialog.InitDir = prefabDir
    commonDialog.FileName = "Untitled.pfb"
    frmSoldatMapEditor.commonDialog.DialogTitle = "Save Prefab"
    commonDialog.ShowSave

    If commonDialog.FileName <> "" Then
        savePrefab commonDialog.FileName
    End If

    RegainFocus

    Exit Sub

ErrorHandler:

    If Error$ <> "Cancel was selected." Then
        MsgBox "Error exporting" & vbNewLine & Error$
    End If
    RegainFocus

End Sub

Private Sub mnuImport_Click()

    On Error GoTo ErrorHandler

    commonDialog.Filter = "Prefab (*.pfb)|*.pfb"
    commonDialog.InitDir = prefabDir
    commonDialog.FileName = ""
    commonDialog.DialogTitle = "Import"
    commonDialog.ShowOpen

    If commonDialog.FileName <> "" Then
        loadPrefab commonDialog.FileName
    End If

    RegainFocus

    Exit Sub

ErrorHandler:

    If Error$ <> "Cancel was selected." Then
        MsgBox "Error importing" & vbNewLine & Error$
    End If
    RegainFocus

End Sub

Private Sub savePrefab(theFileName As String)

    On Error GoTo ErrorHandler

    Dim i As Integer
    Dim j As Integer
    Dim Polygon As TPolygon
    Dim elementName(50) As Byte
    Dim elementString As String
    Dim numSelCon As Integer
    Dim offset As Integer
    Dim tempConnection As TConnection
    Dim alpha As Byte

    Open theFileName For Binary Access Write Lock Write As #1

        Put #1, , numSelectedPolys
        For i = 1 To numSelectedPolys
            Polygon = Polys(selectedPolys(i))
            For j = 1 To 3
                Polygon.vertex(j).X = PolyCoords(selectedPolys(i)).vertex(j).X
                Polygon.vertex(j).Y = PolyCoords(selectedPolys(i)).vertex(j).Y
                vertexList(selectedPolys(i)).vertex(j) = 1
                alpha = getAlpha(Polys(selectedPolys(i)).vertex(j).color)
                Polygon.vertex(j).color = ARGB(alpha, RGB(vertexList(selectedPolys(i)).color(j).blue, vertexList(selectedPolys(i)).color(j).green, vertexList(selectedPolys(i)).color(j).red))
            Next
            Put #1, , Polygon
            Put #1, , vertexList(selectedPolys(i)).vertex(1)
            Put #1, , vertexList(selectedPolys(i)).vertex(2)
            Put #1, , vertexList(selectedPolys(i)).vertex(3)
            Put #1, , vertexList(selectedPolys(i)).polyType
        Next

        Put #1, , numSelectedScenery
        For i = 1 To sceneryCount
            If Scenery(i).selected = 1 Then
                Put #1, , Scenery(i)
                elementString = frmScenery.lstScenery.List(Scenery(i).Style - 1)
                elementName(0) = Len(elementString)
                For j = 1 To elementName(0)
                    elementName(j) = Asc(Mid(elementString, j, 1))
                Next
                Put #1, , elementName
            End If
        Next

        Put #1, , numSelColliders
        For i = 1 To colliderCount
            If Colliders(i).active = 1 Then
                Put #1, , Colliders(i)
            End If
        Next

        Put #1, , numSelSpawns
        For i = 1 To spawnPoints
            If Spawns(i).active = 1 Then
                Put #1, , Spawns(i)
            End If
        Next

        offset = 0
        Put #1, , numSelWaypoints
        For i = 1 To waypointCount
            If Waypoints(i).selected Then
                offset = offset + 1
                Waypoints(i).tempIndex = offset
                Put #1, , Waypoints(i)
            End If
        Next

        numSelCon = 0
        For i = 1 To conCount
            If Waypoints(Connections(i).point1).selected And Waypoints(Connections(i).point2).selected Then
                numSelCon = numSelCon + 1
            End If
        Next

        Put #1, , numSelCon
        For i = 1 To conCount
            If Waypoints(Connections(i).point1).selected And Waypoints(Connections(i).point2).selected Then
                tempConnection.point1 = Waypoints(Connections(i).point1).tempIndex
                tempConnection.point2 = Waypoints(Connections(i).point2).tempIndex
                Put #1, , tempConnection
            End If
        Next

        For i = 1 To waypointCount
            Waypoints(i).tempIndex = i
        Next

    Close #1

    Exit Sub

ErrorHandler:

    MsgBox Error$

End Sub

Private Sub loadPrefab(theFileName As String)

    On Error GoTo ErrorHandler

    Dim newPolys As Integer
    Dim newScenery As Integer
    Dim newElements As Integer
    Dim elementName(50) As Byte
    Dim elementString As String
    Dim newColliders As Integer
    Dim newSpawnPoints As Integer
    Dim newWaypoints As Integer
    Dim newConnections As Integer
    Dim i As Integer
    Dim j As Integer
    Dim tehValue As Integer
    Dim tempClr As TColor

    mnuDeselect_Click

    Open theFileName For Binary Access Read Lock Read As #1

        Get #1, , newPolys
        If newPolys > 0 Then
            ReDim Preserve Polys(mPolyCount + newPolys)
            ReDim Preserve PolyCoords(mPolyCount + newPolys)
            ReDim Preserve vertexList(mPolyCount + newPolys)
            numSelectedPolys = newPolys
            ReDim selectedPolys(newPolys)

            For i = 1 To newPolys
                tehValue = mPolyCount + i
                Get #1, , Polys(tehValue)
                Get #1, , vertexList(tehValue).vertex(1)
                Get #1, , vertexList(tehValue).vertex(2)
                Get #1, , vertexList(tehValue).vertex(3)
                Get #1, , vertexList(tehValue).polyType
                For j = 1 To 3
                    PolyCoords(tehValue).vertex(j).X = Polys(tehValue).vertex(j).X
                    PolyCoords(tehValue).vertex(j).Y = Polys(tehValue).vertex(j).Y
                    Polys(tehValue).vertex(j).X = (PolyCoords(tehValue).vertex(j).X - scrollCoords(2).X) * zoomFactor
                    Polys(tehValue).vertex(j).Y = (PolyCoords(tehValue).vertex(j).Y - scrollCoords(2).Y) * zoomFactor
                    tempClr = getRGB(Polys(tehValue).vertex(j).color)
                    vertexList(tehValue).color(j).red = tempClr.red
                    vertexList(tehValue).color(j).green = tempClr.green
                    vertexList(tehValue).color(j).blue = tempClr.blue
                Next
                selectedPolys(i) = tehValue
            Next
            mPolyCount = mPolyCount + newPolys
        End If

        Get #1, , newScenery
        If newScenery > 0 Then
            If Not showScenery Then
                showScenery = True
                frmDisplay.setLayer 5, showScenery
            End If
            numSelectedScenery = newScenery
            ReDim Preserve Scenery(sceneryCount + newScenery)
            If newScenery > 0 Then
                For i = 1 To newScenery
                    tehValue = sceneryCount + i
                    Get #1, , Scenery(tehValue)
                    Scenery(tehValue).screenTr.X = (Scenery(tehValue).Translation.X - scrollCoords(2).X) * zoomFactor
                    Scenery(tehValue).screenTr.Y = (Scenery(tehValue).Translation.Y - scrollCoords(2).Y) * zoomFactor
                    Scenery(tehValue).Style = 0

                    Get #1, , elementName
                    ' get scenery name
                    elementString = ""
                    For j = 1 To elementName(0)
                        elementString = elementString + Chr$(elementName(j))
                    Next
                    ' find scenery in list
                    For j = 1 To sceneryElements
                        If frmScenery.lstScenery.List(j - 1) = elementString Then
                            Scenery(tehValue).Style = j
                        End If
                    Next
                    ' scenery not in list, so load it
                    If Scenery(tehValue).Style = 0 Then
                        CreateSceneryTexture elementString
                        Scenery(tehValue).Style = sceneryElements
                    End If
                Next
            End If
        End If
        sceneryCount = sceneryCount + newScenery

        Get #1, , newColliders
        If newColliders > 0 Then
            showObjects = True
            numSelColliders = newColliders
            ReDim Preserve Colliders(colliderCount + newColliders)
            For i = 1 To newColliders
                Get #1, , Colliders(colliderCount + i)
            Next
            colliderCount = colliderCount + newColliders
        End If

        Get #1, , newSpawnPoints
        If newSpawnPoints > 0 Then
            showObjects = True
            numSelSpawns = newSpawnPoints
            ReDim Preserve Spawns(spawnPoints + newSpawnPoints)
            For i = 1 To newSpawnPoints
                Get #1, , Spawns(spawnPoints + i)
            Next
            spawnPoints = spawnPoints + newSpawnPoints
        End If

        Get #1, , newWaypoints
        If newWaypoints > 0 Then
            showWaypoints = True
            numSelWaypoints = newWaypoints
            ReDim Preserve Waypoints(waypointCount + newWaypoints)
            For i = 1 To newWaypoints
                Get #1, , Waypoints(waypointCount + i)
            Next
            Get #1, , newConnections
            If newConnections > 0 Then
                ReDim Preserve Connections(conCount + newConnections)
                For i = 1 To newConnections
                    Get #1, , Connections(conCount + i)
                    Connections(conCount + i).point1 = Connections(conCount + i).point1 + waypointCount
                    Connections(conCount + i).point2 = Connections(conCount + i).point2 + waypointCount
                Next
                conCount = conCount + newConnections
            End If
            waypointCount = waypointCount + newWaypoints
            For i = 1 To waypointCount
                Waypoints(i).tempIndex = i
            Next
        End If

        frmDisplay.setLayer 6, showObjects

    Close #1

    setMapData

    getInfo
    getRCenter

    Exit Sub

ErrorHandler:

    MsgBox Error$

End Sub

Private Sub mnuRunSoldat_Click()

    SetGameMode lastCompiled
    SetMapList lastCompiled
    RunSoldat

End Sub

Private Sub SetMapList(theFileName As String)

    Open soldatDir & "mapslist.txt" For Output As #1
        Print #1, theFileName
    Close #1

End Sub

Private Sub mnuUndo_Click()

    loadUndo False

End Sub

Private Sub mnuRedo_Click()

    loadUndo True

End Sub

Private Sub mnuDuplicate_Click()

    Dim i As Integer
    Dim j As Integer
    Dim offset As Integer

    On Error GoTo ErrorHandler

    If selectionChanged Then
        SaveUndo
        selectionChanged = False
    End If

    If numSelectedPolys > 0 Then
        mPolyCount = mPolyCount + numSelectedPolys
        ReDim Preserve Polys(mPolyCount)
        ReDim Preserve PolyCoords(mPolyCount)
        ReDim Preserve vertexList(mPolyCount)
        For i = 1 To numSelectedPolys
            PolyCoords(mPolyCount - numSelectedPolys + i) = PolyCoords(selectedPolys(i))
            PolyCoords(mPolyCount - numSelectedPolys + i).vertex(1).X = PolyCoords(selectedPolys(i)).vertex(1).X + 32
            PolyCoords(mPolyCount - numSelectedPolys + i).vertex(2).X = PolyCoords(selectedPolys(i)).vertex(2).X + 32
            PolyCoords(mPolyCount - numSelectedPolys + i).vertex(3).X = PolyCoords(selectedPolys(i)).vertex(3).X + 32
            Polys(mPolyCount - numSelectedPolys + i) = Polys(selectedPolys(i))
            Polys(mPolyCount - numSelectedPolys + i).vertex(1).X = (PolyCoords(mPolyCount - numSelectedPolys + i).vertex(1).X - scrollCoords(2).X) * zoomFactor
            Polys(mPolyCount - numSelectedPolys + i).vertex(2).X = (PolyCoords(mPolyCount - numSelectedPolys + i).vertex(2).X - scrollCoords(2).X) * zoomFactor
            Polys(mPolyCount - numSelectedPolys + i).vertex(3).X = (PolyCoords(mPolyCount - numSelectedPolys + i).vertex(3).X - scrollCoords(2).X) * zoomFactor
            vertexList(mPolyCount - numSelectedPolys + i).polyType = vertexList(selectedPolys(i)).polyType
            vertexList(mPolyCount - numSelectedPolys + i).color(1) = vertexList(selectedPolys(i)).color(1)
            vertexList(mPolyCount - numSelectedPolys + i).color(2) = vertexList(selectedPolys(i)).color(2)
            vertexList(mPolyCount - numSelectedPolys + i).color(3) = vertexList(selectedPolys(i)).color(3)
            For j = 1 To 3
                vertexList(selectedPolys(i)).vertex(j) = 0
                vertexList(mPolyCount - numSelectedPolys + i).vertex(j) = 1
            Next
            selectedPolys(i) = mPolyCount - numSelectedPolys + i
        Next
    End If
    offset = 0
    If numSelectedScenery > 0 Then
        ReDim Preserve Scenery(sceneryCount + numSelectedScenery)
        For i = 1 To sceneryCount
            If Scenery(i).selected = 1 Then
                offset = offset + 1
                Scenery(sceneryCount + offset) = Scenery(i)
                Scenery(sceneryCount + offset).Translation.X = Scenery(sceneryCount + offset).Translation.X + 32
                Scenery(sceneryCount + offset).screenTr.X = Scenery(sceneryCount + offset).screenTr.X + 32 * zoomFactor
                Scenery(i).selected = 0
            End If
        Next
        sceneryCount = sceneryCount + numSelectedScenery
    End If

    If numSelectedScenery > 0 Or numSelectedPolys > 0 Then
        rCenter.X = rCenter.X + 32
        selRect(0).X = selRect(0).X + 32
        selRect(1).X = selRect(1).X + 32
        selRect(2).X = selRect(2).X + 32
        selRect(3).X = selRect(3).X + 32
    End If

    offset = 0
    For i = 1 To spawnPoints
        If Spawns(i).active = 1 Then
            offset = offset + 1
            ReDim Preserve Spawns(spawnPoints + offset)
            Spawns(spawnPoints + offset) = Spawns(i)
            Spawns(spawnPoints + offset).X = Spawns(spawnPoints + offset).X + 32
            Spawns(i).active = 0
        End If
    Next
    spawnPoints = spawnPoints + offset

    offset = 0
    For i = 1 To colliderCount
        If Colliders(i).active = 1 Then
            offset = offset + 1
            ReDim Preserve Colliders(colliderCount + offset)
            Colliders(colliderCount + offset) = Colliders(i)
            Colliders(colliderCount + offset).X = Colliders(colliderCount + offset).X + 32
            Colliders(i).active = 0
        End If
    Next
    colliderCount = colliderCount + offset

    If numSelWaypoints > 0 Then
        offset = 0
        For i = 1 To waypointCount
            If Waypoints(i).selected Then
                offset = offset + 1
                ReDim Preserve Waypoints(waypointCount + offset)
                Waypoints(waypointCount + offset) = Waypoints(i)
                Waypoints(waypointCount + offset).X = Waypoints(waypointCount + offset).X + 32
                Waypoints(waypointCount + offset).tempIndex = 0
                Waypoints(i).tempIndex = waypointCount + offset
            End If
        Next

        waypointCount = waypointCount + offset

        offset = 0
        For i = 1 To conCount
            If Waypoints(Connections(i).point1).selected And Waypoints(Connections(i).point2).selected Then
                offset = offset + 1
                ReDim Preserve Connections(conCount + offset)
                Connections(conCount + offset).point1 = Waypoints(Connections(i).point1).tempIndex
                Connections(conCount + offset).point2 = Waypoints(Connections(i).point2).tempIndex
            End If
        Next

        conCount = conCount + offset

        For i = 1 To waypointCount
            If Waypoints(i).tempIndex > 0 Then
                Waypoints(i).selected = False
            End If
            Waypoints(i).tempIndex = i
        Next
    End If

    setMapData

    getRCenter

    SaveUndo
    Render
    getInfo

    prompt = True

    Exit Sub

ErrorHandler:

    MsgBox "Duplicate error" & vbNewLine & Error$

End Sub

Private Sub mnuClear_Click()

    deletePolys

End Sub

Private Sub mnuSelectAll_Click()

    Dim i As Integer
    Dim j As Integer

    If showPolys Or showWireframe Or showPoints Then
        ReDim selectedPolys(mPolyCount)
        For i = 1 To mPolyCount
            selectedPolys(i) = i
            For j = 1 To 3
                vertexList(i).vertex(j) = 1
            Next
        Next
        numSelectedPolys = mPolyCount
    End If

    If showScenery Or showWireframe Or showPoints Then
        numSelectedScenery = 0
        For i = 1 To sceneryCount
            If (Scenery(i).level = 0 And sslBack) Or (Scenery(i).level = 1 And sslMid) Or (Scenery(i).level = 2 And sslFront) Then
                Scenery(i).selected = 1
                numSelectedScenery = numSelectedScenery + 1
            End If
        Next
    End If

    If showObjects Then
        For i = 1 To spawnPoints
            Spawns(i).active = 1
        Next
        numSelSpawns = spawnPoints
        For i = 1 To colliderCount
            Colliders(i).active = 1
        Next
        numSelColliders = colliderCount
    End If

    If showLights Then
        For i = 1 To lightCount
            Lights(i).selected = 1
        Next
        numSelLights = lightCount
    End If

    If showWaypoints Then
        For i = 1 To waypointCount
            Waypoints(i).selected = True
        Next
        numSelWaypoints = waypointCount
    End If

    getRCenter
    getInfo

    Render

End Sub

Private Sub mnuDeselect_Click()

    Dim i As Integer
    Dim j As Integer

    numSelectedPolys = 0
    ReDim selectedPolys(0)
    numSelectedScenery = 0
    numSelSpawns = 0
    numSelColliders = 0
    numSelWaypoints = 0

    For i = 1 To mPolyCount
        For j = 1 To 3
            vertexList(i).vertex(j) = 0
        Next
    Next
    For i = 1 To sceneryCount
        Scenery(i).selected = 0
    Next
    For i = 1 To colliderCount
        Colliders(i).active = 0
    Next
    For i = 1 To spawnPoints
        Spawns(i).active = 0
    Next
    For i = 1 To waypointCount
        Waypoints(i).selected = False
    Next

    Render
    getInfo

End Sub

Private Sub mnuSelColor_Click()

    Dim i As Integer
    Dim j As Integer
    Dim addPoly As Byte
    Dim clrVal As TColor

    numSelectedPolys = 0
    ReDim selectedPolys(0)

    For i = 1 To mPolyCount
        For j = 1 To 3
            vertexList(i).vertex(j) = 0
            clrVal = getRGB(Polys(i).vertex(j).color)
            If clrVal.red = gPolyClr.red And clrVal.green = gPolyClr.green And clrVal.blue = gPolyClr.blue Then
                addPoly = 1
                vertexList(i).vertex(j) = 1
            End If
        Next
        If addPoly = 1 Then
            numSelectedPolys = numSelectedPolys + 1
            ReDim Preserve selectedPolys(numSelectedPolys)
            selectedPolys(numSelectedPolys) = i
        End If
        addPoly = 0
    Next

    Render

End Sub

Private Sub mnuBringToFront_Click()

    Dim i As Integer
    Dim j As Integer
    Dim tempTri As TTriangle
    Dim tempPoly As TPolygon
    Dim tempScenery As TScenery
    Dim tempVertex As TVertexData
    Dim offset As Integer

    If selectionChanged Then
        SaveUndo
        selectionChanged = False
    End If

    If numSelectedPolys > 0 Then
        offset = mPolyCount
        For i = mPolyCount To 1 Step -1
            If (vertexList(i).vertex(1) + vertexList(i).vertex(2) + vertexList(i).vertex(3)) > 0 Then  ' if selected
                tempPoly = Polys(i)
                tempTri = PolyCoords(i)
                tempVertex = vertexList(i)
                For j = i To (offset - 1)
                    Polys(j) = Polys(j + 1)
                    PolyCoords(j) = PolyCoords(j + 1)
                    vertexList(j) = vertexList(j + 1)
                Next
                Polys(offset) = tempPoly
                PolyCoords(offset) = tempTri
                vertexList(offset) = tempVertex

                selectedPolys(mPolyCount - offset + 1) = offset
                offset = offset - 1
            End If
        Next
    End If

    If numSelectedScenery > 0 Then
        offset = sceneryCount
        For i = sceneryCount To 1 Step -1
            If Scenery(i).selected Then  ' if selected
                tempScenery = Scenery(i)
                For j = i To (offset - 1)
                    Scenery(j) = Scenery(j + 1)
                Next
                Scenery(offset) = tempScenery
                offset = offset - 1
            End If
        Next
    End If

    prompt = True
    SaveUndo
    Render
    getInfo

End Sub

Private Sub mnuSendToBack_Click()

    Dim i As Integer
    Dim j As Integer
    Dim tempTri As TTriangle
    Dim tempPoly As TPolygon
    Dim tempScenery As TScenery
    Dim tempVertex As TVertexData
    Dim offset As Integer

    If selectionChanged Then
        SaveUndo
        selectionChanged = False
    End If

    If numSelectedPolys > 0 Then
        offset = 1
        For i = 1 To mPolyCount
            If (vertexList(i).vertex(1) + vertexList(i).vertex(2) + vertexList(i).vertex(3)) > 0 Then  ' if selected
                tempPoly = Polys(i)
                tempTri = PolyCoords(i)
                tempVertex = vertexList(i)
                For j = i To offset + 1 Step -1
                    Polys(j) = Polys(j - 1)
                    PolyCoords(j) = PolyCoords(j - 1)
                    vertexList(j) = vertexList(j - 1)
                Next
                Polys(offset) = tempPoly
                PolyCoords(offset) = tempTri
                vertexList(offset) = tempVertex

                selectedPolys(offset) = offset
                offset = offset + 1
            End If
        Next
    End If

    If numSelectedScenery > 0 Then
        offset = 1
        For i = 1 To sceneryCount
            If Scenery(i).selected = 1 Then  ' if selected
                tempScenery = Scenery(i)
                For j = i To offset + 1 Step -1
                    Scenery(j) = Scenery(j - 1)
                Next
                Scenery(offset) = tempScenery
                offset = offset + 1
            End If
        Next
    End If

    prompt = True
    SaveUndo
    Render
    getInfo

End Sub

Private Sub mnuBringForward_Click()

    Dim i As Integer
    Dim tempTri As TTriangle
    Dim tempPoly As TPolygon
    Dim tempScenery As TScenery
    Dim tempVertex As TVertexData
    Dim offset As Integer

    If selectionChanged Then
        SaveUndo
        selectionChanged = False
    End If

    If numSelectedPolys > 0 Then
        offset = mPolyCount
        For i = (mPolyCount - 1) To 1 Step -1
            If (vertexList(i).vertex(1) + vertexList(i).vertex(2) + vertexList(i).vertex(3)) > 0 Then  ' if selected
                If (vertexList(i + 1).vertex(1) + vertexList(i + 1).vertex(2) + vertexList(i + 1).vertex(3)) > 0 Then
                    selectedPolys(mPolyCount - offset + 1) = i + 1
                    offset = offset - 1
                Else
                    tempPoly = Polys(i)
                    tempTri = PolyCoords(i)
                    tempVertex = vertexList(i)

                    Polys(i) = Polys(i + 1)
                    PolyCoords(i) = PolyCoords(i + 1)
                    vertexList(i) = vertexList(i + 1)

                    Polys(i + 1) = tempPoly
                    PolyCoords(i + 1) = tempTri
                    vertexList(i + 1) = tempVertex

                    selectedPolys(mPolyCount - offset + 1) = i + 1
                    offset = offset - 1
                End If
            End If
        Next
    End If

    If numSelectedScenery > 0 Then
        offset = sceneryCount
        For i = (sceneryCount - 1) To 1 Step -1
            If Scenery(i).selected = 1 Then  ' if selected
                If Scenery(i + 1).selected = 1 Then
                    offset = offset - 1
                Else
                    tempScenery = Scenery(i)
                    Scenery(i) = Scenery(i + 1)
                    Scenery(i + 1) = tempScenery
                    offset = offset - 1
                End If
            End If
        Next
    End If

    prompt = True
    SaveUndo
    Render
    getInfo

End Sub

Private Sub mnuSendBackward_Click()

    Dim i As Integer
    Dim tempTri As TTriangle
    Dim tempPoly As TPolygon
    Dim tempVertex As TVertexData
    Dim offset As Integer
    Dim tempScenery As TScenery

    If selectionChanged Then
        SaveUndo
        selectionChanged = False
    End If

    If numSelectedPolys > 0 Then
        offset = 1
        For i = 2 To mPolyCount
            If (vertexList(i).vertex(1) + vertexList(i).vertex(2) + vertexList(i).vertex(3)) > 0 Then  ' if selected
                If (vertexList(i - 1).vertex(1) + vertexList(i - 1).vertex(2) + vertexList(i - 1).vertex(3)) > 0 Then
                    selectedPolys(offset) = i - 1
                    offset = offset + 1
                Else
                    tempPoly = Polys(i)
                    tempTri = PolyCoords(i)
                    tempVertex = vertexList(i)

                    Polys(i) = Polys(i - 1)
                    PolyCoords(i) = PolyCoords(i - 1)
                    vertexList(i) = vertexList(i - 1)

                    Polys(i - 1) = tempPoly
                    PolyCoords(i - 1) = tempTri
                    vertexList(i - 1) = tempVertex

                    selectedPolys(offset) = i - 1
                    offset = offset + 1
                End If
            End If
        Next
    End If

    If numSelectedScenery > 0 Then
        offset = 1
        For i = 2 To sceneryCount
            If Scenery(i).selected = 1 Then  ' if selected
                If Scenery(i - 1).selected = 1 Then
                    offset = offset + 1
                Else
                    tempScenery = Scenery(i)
                    Scenery(i) = Scenery(i - 1)
                    Scenery(i - 1) = tempScenery
                    offset = offset + 1
                End If
            End If
        Next
    End If

    prompt = True
    SaveUndo
    Render
    getInfo

End Sub

Private Sub mnuFixTexture_Click()

    Dim PolyNum As Integer
    Dim i As Integer
    Dim j As Integer

    If selectionChanged Then
        SaveUndo
        selectionChanged = False
    End If

    If numSelectedPolys > 0 Then
        For i = 1 To numSelectedPolys
            PolyNum = selectedPolys(i)
            For j = 1 To 3
                If vertexList(PolyNum).vertex(j) = 1 Then
                    Polys(PolyNum).vertex(j).tu = (PolyCoords(PolyNum).vertex(j).X / xTexture)
                    Polys(PolyNum).vertex(j).tv = (PolyCoords(PolyNum).vertex(j).Y / yTexture)
                End If
            Next
        Next
        prompt = True
    End If

    SaveUndo
    Render
    getInfo

End Sub

Private Sub mnuUntexture_Click()

    Dim i As Integer
    Dim j As Integer

    If selectionChanged Then
        SaveUndo
        selectionChanged = False
    End If

    If numSelectedPolys > 0 Then
        For i = 1 To numSelectedPolys
            For j = 1 To 3
                If vertexList(selectedPolys(i)).vertex(j) = 1 Then
                    Polys(selectedPolys(i)).vertex(j).tu = 1
                    Polys(selectedPolys(i)).vertex(j).tv = 1
                End If
            Next
        Next
        prompt = True
    End If

    SaveUndo
    Render
    getInfo

End Sub

Private Sub mnuVisible_Click()

    Dim i As Integer
    Dim j As Integer

    On Error GoTo ErrorHandler

    If selectionChanged Then
        SaveUndo
        selectionChanged = False
    End If

    For i = 1 To numSelectedPolys
        For j = 1 To 3
            If Polys(selectedPolys(i)).vertex(j).Z < 0 Then
                Polys(selectedPolys(i)).vertex(j).rhw = 1
                Polys(selectedPolys(i)).vertex(j).Z = 1
            Else
                Polys(selectedPolys(i)).vertex(j).rhw = -10
                Polys(selectedPolys(i)).vertex(j).Z = -1
            End If
        Next
    Next

    prompt = True
    SaveUndo
    Render

    Exit Sub

ErrorHandler:

    MsgBox Error$

End Sub

Private Sub mnuAverage_Click()

    AverageVertices

End Sub

Private Sub mnuApplyLight_Click()

    Dim i As Integer
    Dim j As Integer
    Dim tehClr As TColor

    If lightCount = 0 Then Exit Sub

    If numSelectedPolys > 0 Then
        For i = 1 To numSelectedPolys
            For j = 1 To 3
                ' apply poly color to base color
                tehClr = getRGB(Polys(selectedPolys(i)).vertex(j).color)
                vertexList(selectedPolys(i)).color(j).red = tehClr.red
                vertexList(selectedPolys(i)).color(j).green = tehClr.green
                vertexList(selectedPolys(i)).color(j).blue = tehClr.blue
            Next
        Next
    Else
        For i = 1 To mPolyCount
            For j = 1 To 3
                ' apply poly color to base color
                tehClr = getRGB(Polys(i).vertex(j).color)
                vertexList(i).color(j).red = tehClr.red
                vertexList(i).color(j).green = tehClr.green
                vertexList(i).color(j).blue = tehClr.blue
            Next
        Next
    End If

    ReDim Lights(0)
    lightCount = 0

    Render

End Sub

Private Sub mnuSplit_Click()

    If numSelectedPolys < 1 Then Exit Sub

    Dim i As Integer
    Dim j As Integer
    Dim Left As Byte
    Dim Right As Byte
    Dim clr1 As TColor
    Dim clr2 As TColor
    Dim alpha1 As Byte
    Dim alpha2 As Byte
    Dim newPolys As Integer

    If selectionChanged Then
        SaveUndo
        selectionChanged = False
    End If

    For i = 1 To numSelectedPolys
        For j = 1 To 3
            If vertexList(selectedPolys(i)).vertex(j) = 1 Then
                If j = 1 Then
                    Left = 2
                    Right = 3
                ElseIf j = 2 Then
                    Left = 3
                    Right = 1
                ElseIf j = 3 Then
                    Left = 1
                    Right = 2
                End If
                mPolyCount = mPolyCount + 1
                newPolys = newPolys + 1

                ReDim Preserve Polys(mPolyCount)
                ReDim Preserve PolyCoords(mPolyCount)
                ReDim Preserve vertexList(mPolyCount)

                ReDim Preserve selectedPolys(numSelectedPolys + newPolys)
                selectedPolys(numSelectedPolys + newPolys) = mPolyCount
                vertexList(mPolyCount).vertex(j) = 1

                PolyCoords(mPolyCount).vertex(j) = PolyCoords(selectedPolys(i)).vertex(j)
                PolyCoords(mPolyCount).vertex(Left) = PolyCoords(selectedPolys(i)).vertex(Left)

                PolyCoords(mPolyCount).vertex(Right).X = Midpoint(PolyCoords(selectedPolys(i)).vertex(Left).X, PolyCoords(selectedPolys(i)).vertex(Right).X)
                PolyCoords(mPolyCount).vertex(Right).Y = Midpoint(PolyCoords(selectedPolys(i)).vertex(Left).Y, PolyCoords(selectedPolys(i)).vertex(Right).Y)

                PolyCoords(selectedPolys(i)).vertex(Left) = PolyCoords(mPolyCount).vertex(Right)

                Polys(mPolyCount).vertex(j) = Polys(selectedPolys(i)).vertex(j)
                Polys(mPolyCount).vertex(Left) = Polys(selectedPolys(i)).vertex(Left)
                Polys(mPolyCount).Perp.vertex(1).Z = 2
                Polys(mPolyCount).Perp.vertex(2).Z = 2
                Polys(mPolyCount).Perp.vertex(3).Z = 2

                ' coords
                Polys(mPolyCount).vertex(Right) = Polys(selectedPolys(i)).vertex(Right)
                Polys(mPolyCount).vertex(Right).X = (PolyCoords(mPolyCount).vertex(Right).X - scrollCoords(2).X) * zoomFactor
                Polys(mPolyCount).vertex(Right).Y = (PolyCoords(mPolyCount).vertex(Right).Y - scrollCoords(2).Y) * zoomFactor

                ' texture coords
                Polys(mPolyCount).vertex(Right).tu = Midpoint(Polys(selectedPolys(i)).vertex(Right).tu, Polys(mPolyCount).vertex(Left).tu)
                Polys(mPolyCount).vertex(Right).tv = Midpoint(Polys(selectedPolys(i)).vertex(Right).tv, Polys(mPolyCount).vertex(Left).tv)

                vertexList(mPolyCount).color(j) = vertexList(selectedPolys(i)).color(j)
                vertexList(mPolyCount).color(Left) = vertexList(selectedPolys(i)).color(Left)

                ' colors
                clr1 = vertexList(selectedPolys(i)).color(Right)
                clr2 = vertexList(mPolyCount).color(Left)
                vertexList(mPolyCount).color(Right).red = clr1.red * 0.5 + clr2.red * 0.5
                vertexList(mPolyCount).color(Right).green = clr1.green * 0.5 + clr2.green * 0.5
                vertexList(mPolyCount).color(Right).blue = clr1.blue * 0.5 + clr2.blue * 0.5

                vertexList(selectedPolys(i)).color(Left) = vertexList(mPolyCount).color(Right)

                clr1 = getRGB(Polys(selectedPolys(i)).vertex(Right).color)
                clr2 = getRGB(Polys(mPolyCount).vertex(Left).color)
                alpha1 = getAlpha(Polys(selectedPolys(i)).vertex(Right).color)
                alpha2 = getAlpha(Polys(mPolyCount).vertex(Left).color)
                Polys(mPolyCount).vertex(Right).color = ARGB((alpha1 * 0.5 + alpha2 * 0.5), RGB((clr1.blue * 0.5 + clr2.blue * 0.5), (clr1.green * 0.5 + clr2.green * 0.5), (clr1.red * 0.5 + clr2.red * 0.5)))

                Polys(selectedPolys(i)).vertex(Left) = Polys(mPolyCount).vertex(Right)

                vertexList(mPolyCount).polyType = vertexList(selectedPolys(i)).polyType
            End If
        Next
    Next

    numSelectedPolys = numSelectedPolys + newPolys
    SaveUndo
    Render
    getInfo

    frmInfo.lblCount(0).Caption = mPolyCount
    frmInfo.lblCount(6).Caption = getMapDimensions

End Sub

Private Sub mnuJoinVertices_Click()

    Dim firstVertex As Integer
    Dim i As Integer
    Dim j As Integer

    If selectionChanged Then
        SaveUndo
        selectionChanged = False
    End If

    If numSelectedPolys > 0 Then
        For j = 1 To 3
            If vertexList(selectedPolys(1)).vertex(4 - j) = 1 Then
                firstVertex = 4 - j
            End If
        Next

        For i = 2 To numSelectedPolys
            For j = 1 To 3
                If vertexList(selectedPolys(i)).vertex(j) = 1 Then
                    PolyCoords(selectedPolys(i)).vertex(j).X = PolyCoords(selectedPolys(1)).vertex(firstVertex).X
                    PolyCoords(selectedPolys(i)).vertex(j).Y = PolyCoords(selectedPolys(1)).vertex(firstVertex).Y
                    Polys(selectedPolys(i)).vertex(j).X = Polys(selectedPolys(1)).vertex(firstVertex).X
                    Polys(selectedPolys(i)).vertex(j).Y = Polys(selectedPolys(1)).vertex(firstVertex).Y
                End If
            Next
        Next

        prompt = True
    End If

    SaveUndo
    Render
    getInfo

End Sub

Private Sub mnuCreate_Click()

    If numSelectedPolys < 1 Or numSelectedPolys > 3 Then Exit Sub

    Dim i As Integer
    Dim j As Integer
    Dim numSelVerts As Integer
    Dim temp As D3DVECTOR2
    Dim tempVertex As TCustomVertex
    Dim tempClr As TColor

    If selectionChanged Then
        SaveUndo
        selectionChanged = False
    End If

    ReDim Preserve Polys(mPolyCount + 1)
    ReDim Preserve PolyCoords(mPolyCount + 1)
    ReDim Preserve vertexList(mPolyCount + 1)

    For i = 1 To numSelectedPolys
        For j = 1 To 3
            If vertexList(selectedPolys(i)).vertex(j) = 1 Then
                numSelVerts = numSelVerts + 1
                Polys(mPolyCount + 1).vertex(numSelVerts) = Polys(selectedPolys(i)).vertex(j)
                PolyCoords(mPolyCount + 1).vertex(numSelVerts) = PolyCoords(selectedPolys(i)).vertex(j)
                vertexList(mPolyCount + 1).color(numSelVerts) = vertexList(selectedPolys(i)).color(j)
                vertexList(mPolyCount + 1).polyType = vertexList(selectedPolys(i)).polyType
            End If
            If numSelVerts = 3 Then Exit For
        Next
        If numSelVerts = 3 Then Exit For
    Next

    If numSelVerts > 2 Then
        mPolyCount = mPolyCount + 1
    End If

    If Not isCW(mPolyCount) Then  ' switch to make cw
        temp = PolyCoords(mPolyCount).vertex(3)
        PolyCoords(mPolyCount).vertex(3) = PolyCoords(mPolyCount).vertex(2)
        PolyCoords(mPolyCount).vertex(2) = temp

        tempVertex = Polys(mPolyCount).vertex(3)
        Polys(mPolyCount).vertex(3) = Polys(mPolyCount).vertex(2)
        Polys(mPolyCount).vertex(2) = tempVertex

        tempClr = vertexList(mPolyCount).color(3)
        vertexList(mPolyCount).color(3) = vertexList(mPolyCount).color(2)
        vertexList(mPolyCount).color(2) = tempClr
    End If

    Polys(mPolyCount).Perp.vertex(1).Z = 2
    Polys(mPolyCount).Perp.vertex(2).Z = 2
    Polys(mPolyCount).Perp.vertex(3).Z = 2

    frmInfo.lblCount(0).Caption = mPolyCount
    frmInfo.lblCount(6).Caption = getMapDimensions

    SaveUndo
    Render

End Sub

Private Sub mnuSever_Click()

    Dim i As Integer
    Dim offset As Integer
    Dim numConnections As Integer

    If selectionChanged Then
        SaveUndo
        selectionChanged = False
    End If

    numConnections = conCount

    If numSelWaypoints > 1 Then
        offset = 1
        For i = 1 To conCount
            Connections(offset) = Connections(i)
            If Waypoints(Connections(i).point1).selected And Waypoints(Connections(i).point2).selected Then
                numConnections = numConnections - 1
                Waypoints(Connections(i).point1).numConnections = Waypoints(Connections(i).point1).numConnections - 1
            Else  ' not selected
                offset = offset + 1
            End If
        Next
    ElseIf numSelWaypoints = 1 Then
        offset = 1
        For i = 1 To conCount
            Connections(offset) = Connections(i)
            If Waypoints(Connections(i).point1).selected Or Waypoints(Connections(i).point2).selected Then
                numConnections = numConnections - 1
                Waypoints(Connections(i).point1).numConnections = Waypoints(Connections(i).point1).numConnections - 1
            Else  ' not selected
                offset = offset + 1
            End If
        Next
    End If

    conCount = numConnections
    ReDim Preserve Connections(conCount)

    SaveUndo
    Render

End Sub

Private Sub mnuRefreshBG_Click()

    Dim i As Integer
    Dim j As Integer
    Dim bgSize As Integer
    Dim xOffset As Integer
    Dim yOffset As Integer

    maxX = 0
    maxY = 0
    minX = 0
    minY = 0

    If mPolyCount > 0 Then
        For i = 1 To mPolyCount
            For j = 1 To 3
                If PolyCoords(i).vertex(j).X > maxX Then maxX = PolyCoords(i).vertex(j).X
                If PolyCoords(i).vertex(j).X < minX Then minX = PolyCoords(i).vertex(j).X
                If PolyCoords(i).vertex(j).Y > maxY Then maxY = PolyCoords(i).vertex(j).Y
                If PolyCoords(i).vertex(j).Y < minY Then minY = PolyCoords(i).vertex(j).Y
            Next
        Next
    End If

    xOffset = Int(Midpoint(maxX, minX))
    yOffset = Int(Midpoint(maxY, minY))

    If (maxX - minX) > (maxY - minY) Then
        bgSize = maxX - xOffset
    Else
        bgSize = maxY - xOffset
    End If

    bgPolyCoords(1).X = xOffset - (bgSize + 640)
    bgPolyCoords(1).Y = yOffset - (bgSize + 640)

    bgPolyCoords(2).X = xOffset - (bgSize + 640)
    bgPolyCoords(2).Y = yOffset + (bgSize + 640)

    bgPolyCoords(3).X = xOffset + (bgSize + 640)
    bgPolyCoords(3).Y = yOffset - (bgSize + 640)

    bgPolyCoords(4).X = xOffset + (bgSize + 640)
    bgPolyCoords(4).Y = yOffset + (bgSize + 640)

    For i = 1 To 4
        bgPolys(i).X = (bgPolyCoords(i).X - scrollCoords(2).X) * zoomFactor
        bgPolys(i).Y = (bgPolyCoords(i).Y - scrollCoords(2).Y) * zoomFactor
    Next

    frmInfo.lblCount(6).Caption = getMapDimensions

    Render

End Sub

Private Sub mnuPreferences_Click()

    frmPreferences.Show 1
    gPolyTypeClrs(0) = frmSoldatMapEditor.selectionColor

End Sub

Private Sub mnuMap_Click()

    frmMap.Show 1
    ctrlDown = False
    setCurrentTool currentTool

End Sub

Private Sub mnuZoomIn_Click()

    Zoom 2

End Sub

Private Sub mnuZoomOut_Click()

    Zoom 0.5

End Sub

Private Sub mnuGrid_Click()

    mnuGrid.Checked = Not mnuGrid.Checked
    showGrid = mnuGrid.Checked
    frmDisplay.setLayer 8, mnuGrid.Checked
    Render

End Sub

Private Sub mnuSnapToGrid_Click()

    mnuSnapToGrid.Checked = Not mnuSnapToGrid.Checked
    snapToGrid = mnuSnapToGrid.Checked

End Sub

Private Sub mnuRefresh_Click()

    resetDevice

End Sub

Private Sub mnuTools_Click()

    mnuTools.Checked = Not mnuTools.Checked
    frmTools.Visible = mnuTools.Checked

End Sub

Private Sub mnuDisplay_Click()

    mnuDisplay.Checked = Not mnuDisplay.Checked
    frmDisplay.Visible = mnuDisplay.Checked

End Sub

Private Sub mnuPalette_Click()

    mnuPalette.Checked = Not mnuPalette.Checked
    frmPalette.Visible = mnuPalette.Checked

End Sub

Private Sub mnuWaypoints_Click()

    mnuWaypoints.Checked = Not mnuWaypoints.Checked
    frmWaypoints.Visible = mnuWaypoints.Checked

End Sub

Private Sub mnuScenery_Click()

    mnuScenery.Checked = Not mnuScenery.Checked
    frmScenery.Visible = mnuScenery.Checked

End Sub

Private Sub mnuinfo_Click()

    mnuInfo.Checked = Not mnuInfo.Checked
    frmInfo.Visible = mnuInfo.Checked

End Sub

Private Sub mnuTexture_Click()

    mnuTexture.Checked = Not mnuTexture.Checked
    frmTexture.Visible = mnuTexture.Checked

End Sub

Private Sub mnuBlendWireframe_Click()

    mnuBlendWireframe.Checked = Not mnuBlendWireframe.Checked
    clrWireframe = mnuBlendWireframe.Checked

End Sub

Private Sub mnuBlendPolys_Click()

    mnuBlendPolys.Checked = Not mnuBlendPolys.Checked
    clrPolys = mnuBlendPolys.Checked

End Sub

Private Sub mnuFixedTexture_Click()

    mnuFixedTexture.Checked = Not mnuFixedTexture.Checked
    fixedTexture = mnuFixedTexture.Checked

End Sub

Private Sub mnuSnapToVerts_Click()

    mnuSnapToVerts.Checked = Not mnuSnapToVerts.Checked
    ohSnap = mnuSnapToVerts.Checked

End Sub

Private Sub mnuLoadSpace_Click()

    On Error GoTo ErrorHandler

    frmSoldatMapEditor.commonDialog.Filter = "Ini File (*.ini)|*.ini"
    commonDialog.InitDir = appPath & "\Workspace\"
    commonDialog.FileName = ""
    frmSoldatMapEditor.commonDialog.DialogTitle = "Load Workspace"
    commonDialog.ShowOpen

    If commonDialog.FileName <> "" Then
        If Len(Dir$(appPath & "\Workspace\" & commonDialog.FileTitle)) <> 0 Then
            loadWorkspace commonDialog.FileTitle
            frmTools.setForm
            frmDisplay.setForm
            frmInfo.setForm
            frmPalette.setForm
            frmScenery.setForm
            frmTexture.setForm
            frmWaypoints.setForm
            resetDevice
        End If
    End If

    RegainFocus

    Exit Sub

ErrorHandler:

    RegainFocus

End Sub

Private Sub mnuSaveSpace_Click()

    On Error GoTo ErrorHandler

    Dim iniString As String
    Dim sNull As String
    Dim isNewFile As Boolean

    sNull = Chr$(0)
    isNewFile = False

    frmSoldatMapEditor.commonDialog.Filter = "Ini File (*.ini)|*.ini"
    commonDialog.InitDir = appPath & "\Workspace\"
    commonDialog.FileName = ""
    frmSoldatMapEditor.commonDialog.DialogTitle = "Save Workspace"
    commonDialog.ShowSave

    If commonDialog.FileName <> "" Then
        isNewFile = Not FileExists(appPath & "\workspace\" & commonDialog.FileTitle)

        iniString = _
            "WindowState=" & Me.Tag & sNull & _
            "Width=" & formWidth & sNull & _
            "Height=" & formHeight & sNull & _
            "Left=" & formLeft & sNull & _
            "Top=" & formTop & _
            IIf(isNewFile, vbNewLine, "") & sNull & sNull
        saveSection "Main", iniString, appPath & "\workspace\" & commonDialog.FileTitle

        saveWindow "Tools", frmTools, False, isNewFile, commonDialog.FileTitle
        saveWindow "Display", frmDisplay, frmDisplay.collapsed, isNewFile, commonDialog.FileTitle
        saveWindow "Properties", frmInfo, frmInfo.collapsed, isNewFile, commonDialog.FileTitle
        saveWindow "Palette", frmPalette, frmPalette.collapsed, isNewFile, commonDialog.FileTitle
        saveWindow "Scenery", frmScenery, frmScenery.collapsed, isNewFile, commonDialog.FileTitle
        saveWindow "Waypoints", frmWaypoints, frmWaypoints.collapsed, isNewFile, commonDialog.FileTitle
        saveWindow "Texture", frmTexture, frmTexture.collapsed, isNewFile, commonDialog.FileTitle
    End If

    RegainFocus

    Exit Sub

ErrorHandler:

    RegainFocus

End Sub

Private Sub mnuResetWindows_Click()

    If Me.Tag = vbNormal Then
        formWidth = Screen.Width / Screen.TwipsPerPixelX - (64 + 208 + 208)
        formHeight = formWidth * 3 / 4
        formLeft = Screen.Width / Screen.TwipsPerPixelX / 2 - formWidth / 2 - 1
        formTop = Screen.Height / Screen.TwipsPerPixelY / 2 - formHeight / 2 - 1

        tvwScenery.Height = formHeight - 41 - 20

        Me.Width = formWidth * Screen.TwipsPerPixelX
        Me.Height = formHeight * Screen.TwipsPerPixelY
        Me.Left = Screen.Width / 2 - Me.Width / 2 - Screen.TwipsPerPixelX
        Me.Top = Screen.Height / 2 - Me.Height / 2 - Screen.TwipsPerPixelY

        picResize.Top = formHeight - picResize.Height
        picResize.Left = formWidth - picResize.Width

        frmTools.Left = Me.Left - frmTools.Width + Screen.TwipsPerPixelX
        frmTools.Top = Me.Top + 41 * Screen.TwipsPerPixelY
        frmPalette.Left = Me.Left + Me.Width - Screen.TwipsPerPixelX
        frmPalette.Top = Me.Top + 41 * Screen.TwipsPerPixelY
        frmDisplay.Left = frmPalette.Left
        frmDisplay.Top = frmPalette.Top + frmPalette.Height - Screen.TwipsPerPixelY
        frmScenery.Left = Me.Left + Me.Width - Screen.TwipsPerPixelX
        frmScenery.Top = frmDisplay.Top + frmDisplay.Height - Screen.TwipsPerPixelY
        frmInfo.Left = Me.Left - frmInfo.Width + Screen.TwipsPerPixelX
        frmInfo.Top = frmTools.Top + frmTools.Height - Screen.TwipsPerPixelY
        frmWaypoints.Left = Me.Left - frmWaypoints.Width + Screen.TwipsPerPixelX
        frmWaypoints.Top = frmInfo.Top + frmInfo.Height - Screen.TwipsPerPixelY
        frmTexture.Top = frmPalette.Top
        frmTexture.Left = frmPalette.Left - frmTexture.Width + Screen.TwipsPerPixelX

        resetDevice
    Else
        frmTools.Left = Me.Left
        frmTools.Top = Me.Top + 41 * Screen.TwipsPerPixelY
        frmPalette.Left = Me.Left + Me.Width - frmPalette.Width
        frmPalette.Top = Me.Top + 41 * Screen.TwipsPerPixelY
        frmDisplay.Left = frmPalette.Left
        frmDisplay.Top = frmPalette.Top + frmPalette.Height - Screen.TwipsPerPixelY
        frmWaypoints.Left = Me.Left
        frmWaypoints.Top = Me.Top + Me.Height - frmWaypoints.Height - 19 * Screen.TwipsPerPixelY
        frmScenery.Left = Me.Left + Me.Width - frmScenery.Width
        frmScenery.Top = frmDisplay.Top + frmDisplay.Height - Screen.TwipsPerPixelY
        frmInfo.Left = Me.Left
        frmInfo.Top = frmWaypoints.Top - frmInfo.Height + Screen.TwipsPerPixelY
        frmTexture.Top = frmPalette.Top
        frmTexture.Left = frmPalette.Left - frmTexture.Width + Screen.TwipsPerPixelX

    End If

End Sub

Private Sub mnuShowAll_Click()

    mnuTools.Checked = True
    frmTools.Visible = True

    mnuPalette.Checked = True
    frmPalette.Visible = True

    mnuDisplay.Checked = True
    frmDisplay.Visible = True

    mnuScenery.Checked = True
    frmScenery.Visible = True

    mnuInfo.Checked = True
    frmInfo.Visible = True

    mnuTexture.Checked = True
    frmTexture.Visible = True

    mnuWaypoints.Checked = True
    frmWaypoints.Visible = True

End Sub

Private Sub mnuHideAll_Click()

    mnuTools.Checked = False
    frmTools.Visible = False

    mnuPalette.Checked = False
    frmPalette.Visible = False

    mnuDisplay.Checked = False
    frmDisplay.Visible = False

    mnuScenery.Checked = False
    frmScenery.Visible = False

    mnuInfo.Checked = False
    frmInfo.Visible = False

    mnuTexture.Checked = False
    frmTexture.Visible = False

    mnuWaypoints.Checked = False
    frmWaypoints.Visible = False

End Sub

Private Sub mnuGostek_Click()

    If mnuGostek.Checked Then
        gostek.X = 0
        gostek.Y = 0
    Else
        mnuGostek.Checked = True
        mnuSpawn(Spawns(0).Team).Checked = False
        mnuCollider.Checked = False
    End If

End Sub

Private Sub mnuCollider_Click()

    mnuCollider.Checked = True
    mnuSpawn(Spawns(0).Team).Checked = False
    mnuGostek.Checked = False
    Colliders(0).radius = clrRadius
End Sub

Private Sub mnuSpawn_Click(Index As Integer)

    mnuCollider.Checked = False
    mnuSpawn(Spawns(0).Team).Checked = False
    mnuSpawn(Index).Checked = True
    mnuGostek.Checked = False
    Spawns(0).Team = Index

End Sub

Private Sub mnuPolyType_Click(Index As Integer)

    mnuPolyType(polyType).Checked = False
    mnuPolyType(Index).Checked = True
    polyType = Index
    lblCurrentTool.Caption = frmSoldatMapEditor.ImageList.ListImages(currentFunction + 1).Tag & " (" & mnuPolyType(polyType).Caption & ")"

End Sub

Private Sub mnuQuad_Click()

    mnuQuad.Checked = Not mnuQuad.Checked

    If mnuQuad.Checked Then
        currentFunction = TOOL_QUAD
        SetCursor currentFunction + 1
        lblCurrentTool.Caption = frmSoldatMapEditor.ImageList.ListImages(currentFunction + 1).Tag
    Else
        currentFunction = TOOL_CREATE
        SetCursor currentFunction + 1
        lblCurrentTool.Caption = frmSoldatMapEditor.ImageList.ListImages(currentFunction + 1).Tag
    End If

    lblCurrentTool.Caption = frmSoldatMapEditor.ImageList.ListImages(currentFunction + 1).Tag & " (" & mnuPolyType(polyType).Caption & ")"

End Sub

Private Sub mnuCustomX_Click()

    mnuCustomX.Checked = Not mnuCustomX.Checked

End Sub

Private Sub mnuCustomY_Click()

    mnuCustomY.Checked = Not mnuCustomY.Checked

End Sub

Private Sub mnuFitOnScreen_Click()

    If mPolyCount < 1 Then Exit Sub

    Dim Width As Integer
    Dim Height As Integer

    mnuRefreshBG_Click

    scrollCoords(2).X = -Me.ScaleWidth / 2 - 1 + Midpoint(minX, maxX)
    scrollCoords(2).Y = -Me.ScaleHeight / 2 - 25 + Midpoint(minY, maxY)
    zoomFactor = 1

    Width = maxX - minX
    Height = maxY - minY

    If Height / Width < (Me.ScaleHeight - 88) / (Me.ScaleWidth - 32) Then
        Zoom ((Me.ScaleWidth - 32) / Width)
    Else
        Zoom ((Me.ScaleHeight - 88) / Height)
    End If

End Sub

Private Sub mnuActualPixels_Click()

    zoomFactor = (Me.ScaleWidth + 2) / 640
    Zoom 1

End Sub

Private Sub mnuScenTrans_Click(Index As Integer)

    mnuScenTrans(Index).Checked = Not mnuScenTrans(Index).Checked

    If Index = 0 Then  ' rotate
        frmScenery.rotateScenery = mnuScenTrans(Index).Checked
        mouseEvent2 frmScenery.picRotate, 0, 0, BUTTON_SMALL, frmScenery.rotateScenery, BUTTON_UP
    ElseIf Index = 1 Then
        frmScenery.scaleScenery = mnuScenTrans(Index).Checked
        mouseEvent2 frmScenery.picScale, 0, 0, BUTTON_SMALL, frmScenery.scaleScenery, BUTTON_UP
    End If

End Sub

Public Sub getInfo()

    Dim i As Integer
    Dim j As Integer
    Dim scenNum As Integer

    On Error GoTo ErrorHandler

    frmInfo.noChange = True
    frmWaypoints.noChange = True

    For i = 1 To waypointCount
        If Waypoints(i).selected Then
            frmWaypoints.getPathNum Waypoints(i).pathNum
            For j = 0 To 4
                frmWaypoints.getWayType j, Waypoints(i).wayType(j)
            Next
            frmWaypoints.cboSpecial.ListIndex = Waypoints(i).special
            frmWaypoints.lblNumCon = Waypoints(i).numConnections
            Exit For
        End If
    Next

    If numSelectedPolys = 0 And numSelectedScenery = 0 Then
        If numSelLights > 0 Then
            For i = 1 To lightCount
                If Lights(i).selected = 1 Then
                    frmInfo.txtLightProp(0).Text = Lights(i).Z
                    frmInfo.txtLightProp(1).Text = Lights(i).range
                    frmInfo.picLight.BackColor = RGB(Lights(i).color.red, Lights(i).color.green, Lights(i).color.blue)
                    Exit For
                End If
            Next
            frmInfo.mnuProp_Click 4
        Else
            frmInfo.mnuProp_Click 5
        End If
        frmInfo.lblCoords = ""
        frmInfo.lblIndex = ""
        frmInfo.lblSelPolys = ""
        frmInfo.lblSelScenery = ""
        frmInfo.noChange = False
        frmWaypoints.noChange = False
        Exit Sub
    End If

    If numSelectedPolys > 0 Then
        frmInfo.cboPolyType.ListIndex = vertexList(selectedPolys(1)).polyType
        frmInfo.txtBounciness.Enabled = False
        For j = 1 To 3
            If vertexList(selectedPolys(1)).vertex(j) = 1 Then
                frmInfo.txtBounciness.Text = Int((Polys(selectedPolys(1)).Perp.vertex(j).Z - 1) * 100)
                If frmInfo.txtBounciness.Text < 0 Then
                    frmInfo.txtBounciness.Text = 0
                End If
                If frmInfo.cboPolyType.ListIndex = 18 Then
                    frmInfo.txtBounciness.Enabled = True
                End If
                frmInfo.txtTexture(0).Text = Int(Polys(selectedPolys(1)).vertex(j).tu * 10000 + 0.5) / 10000
                frmInfo.txtTexture(1).Text = Int(Polys(selectedPolys(1)).vertex(j).tv * 10000 + 0.5) / 10000
                frmInfo.txtVertexAlpha.Text = Int((getAlpha(Polys(selectedPolys(1)).vertex(j).color) / 255 * 100) * 100 + 0.5) / 100
                frmInfo.lblCoords.Caption = Int(PolyCoords(selectedPolys(1)).vertex(j).X * 100 + 0.5) / 100 & ", " & Int(PolyCoords(selectedPolys(1)).vertex(j).Y * 100) / 100
                Exit For
            End If
        Next
    End If

    If numSelectedScenery > 0 Then

        For i = 1 To sceneryCount
            If Scenery(i).selected = 1 Then
                scenNum = i
                frmInfo.txtScenProp(0).Text = Int(Scenery(i).Scaling.X * 100 * 100 + 0.5) / 100
                frmInfo.txtScenProp(1).Text = Int(Scenery(i).Scaling.Y * 100 * 100 + 0.5) / 100
                frmInfo.txtScenProp(2).Text = Int(Scenery(i).alpha / 255 * 100 * 10 + 0.5) / 10
                frmInfo.txtScenProp(3).Text = Int(Scenery(i).rotation * 180 / PI * 10 + 0.5) / 10
                frmInfo.cboLevel.ListIndex = Scenery(i).level
                If numSelectedPolys = 0 Then
                    frmInfo.lblCoords.Caption = Int(Scenery(i).Translation.X * 100 + 0.5) / 100 & ", " & Int(Scenery(i).Translation.Y * 100) / 100
                End If
                Exit For
            End If
        Next
    End If

    If numSelectedPolys = 1 And numSelectedScenery = 0 Then
        frmInfo.lblIndex.Caption = selectedPolys(1)
    ElseIf numSelectedPolys = 0 And numSelectedScenery = 1 Then
        frmInfo.lblIndex.Caption = scenNum
    Else
        frmInfo.lblIndex.Caption = ""
    End If

    If currentTool = TOOL_MOVE Then
        If numSelectedPolys = 0 And numSelectedScenery = 1 Then
            frmInfo.mnuProp_Click 1
        Else
            frmInfo.mnuProp_Click 2
        End If
    ElseIf numSelectedPolys > 0 And numSelectedScenery = 0 Then
        frmInfo.mnuProp_Click 0
    ElseIf numSelectedPolys = 0 And numSelectedScenery > 0 Then
        frmInfo.mnuProp_Click 1
    End If

    frmInfo.txtScale(0).Text = Int(scaleDiff.X * 1000 + 0.5) / 10
    frmInfo.txtScale(1).Text = Int(scaleDiff.Y * 1000 + 0.5) / 10
    frmInfo.txtRotate.Text = rDiff

    If numSelectedScenery = 1 And numSelectedPolys = 0 Then
        frmInfo.lblSelPolys = ""
        frmInfo.lblSelScenery = frmScenery.lstScenery.List(Scenery(scenNum).Style - 1)
    Else
        If numSelectedPolys = 0 Then
            frmInfo.lblSelPolys = ""
        Else
            frmInfo.lblSelPolys = "Polys: " & numSelectedPolys
        End If
        If numSelectedScenery = 0 Then
            frmInfo.lblSelScenery = ""
        Else
            frmInfo.lblSelScenery = "Scenery: " & numSelectedScenery
        End If
    End If

    If numSelWaypoints = 0 Then
        frmWaypoints.ClearWaypt
    End If

    frmInfo.noChange = False
    frmWaypoints.noChange = False

    Exit Sub

ErrorHandler:

    MsgBox "getInfo() error" & vbNewLine & Error$

End Sub

' apply scale/rotate

Public Sub applyPolyType(ByVal Index As Integer)

    Dim i As Integer

    If selectionChanged Then
        SaveUndo
        selectionChanged = False
    End If

    If numSelectedPolys > 0 Then
        For i = 1 To numSelectedPolys
            vertexList(selectedPolys(i)).polyType = Index
        Next
    End If
    SaveUndo
    Render

End Sub

Public Sub applyTextureCoords(ByVal tehValue As Single, Index As Integer)

    Dim i As Integer
    Dim j As Integer

    If selectionChanged Then
        SaveUndo
        selectionChanged = False
    End If

    If numSelectedPolys > 0 Then
        For i = 1 To numSelectedPolys
            For j = 1 To 3
                If vertexList(selectedPolys(i)).vertex(j) = 1 Then
                    If Index = 0 Then
                        Polys(selectedPolys(i)).vertex(j).tu = tehValue
                    Else
                        Polys(selectedPolys(i)).vertex(j).tv = tehValue
                    End If
                End If
            Next
        Next
    End If
    SaveUndo
    Render

End Sub

Public Sub applyVertexAlpha(tehValue As Single)

    Dim i As Integer
    Dim j As Integer

    If selectionChanged Then
        SaveUndo
        selectionChanged = False
    End If

    If numSelectedPolys > 0 Then
        For i = 1 To numSelectedPolys
            For j = 1 To 3
                If vertexList(selectedPolys(i)).vertex(j) = 1 Then
                    Polys(selectedPolys(i)).vertex(j).color = ARGB(tehValue * 255, Polys(selectedPolys(i)).vertex(j).color)
                End If
            Next
        Next
    End If
    SaveUndo
    Render

End Sub

Public Sub applyBounciness(tehValue As Single)

    Dim i As Integer
    Dim j As Integer

    If selectionChanged Then
        SaveUndo
        selectionChanged = False
    End If

    If numSelectedPolys > 0 Then
        For i = 1 To numSelectedPolys
            For j = 1 To 3
                Polys(selectedPolys(i)).Perp.vertex(j).Z = tehValue
            Next
        Next
    End If
    SaveUndo

End Sub

Public Sub applySceneryProp(ByVal tehValue As Single, Index As Integer)

    Dim i As Integer
    Dim tempClr As TColor

    If selectionChanged Then
        SaveUndo
        selectionChanged = False
    End If

    For i = 1 To sceneryCount
        If Scenery(i).selected = 1 Then
            If Index = 0 Then  ' x scale
                Scenery(i).Scaling.X = tehValue
            ElseIf Index = 1 Then  ' y scale
                Scenery(i).Scaling.Y = tehValue
            ElseIf Index = 2 Then  ' alpha
                tempClr = getRGB(Scenery(i).color)
                Scenery(i).alpha = tehValue
                Scenery(i).color = ARGB(tehValue, RGB(tempClr.blue, tempClr.green, tempClr.red))
            ElseIf Index = 3 Then  ' rotation
                Scenery(i).rotation = tehValue
            ElseIf Index = 4 Then  ' level
                Scenery(i).level = tehValue
            End If
        End If
    Next
    If Index = 0 Or Index = 1 Or Index = 3 Then
        getRCenter
    End If
    SaveUndo
    Render

End Sub

Public Sub applyLightProp(ByVal tehValue As Single, Index As Integer)

    Dim i As Integer

    If selectionChanged Then
        SaveUndo
        selectionChanged = False
    End If

    For i = 1 To lightCount
        If Lights(i).selected = 1 Then
            If Index = 0 Then  ' z-coord
                Lights(i).Z = tehValue
            ElseIf Index = 1 Then
                Lights(i).range = tehValue
            End If
        End If
    Next
    SaveUndo
    applyLights
    Render

End Sub

Private Sub picMenu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picMenu(Index), X, Y, BUTTON_MENU, 0, BUTTON_DOWN
    PopupMenu mnuMenu(Index), , Index * MENU_WIDTH, 41
    mouseEvent2 picMenu(Index), X, Y, BUTTON_MENU, 0, BUTTON_UP

End Sub

Private Sub picMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picMenu(Index), X, Y, BUTTON_MENU, 0, BUTTON_MOVE

End Sub

Private Sub picHelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picHelp, X, Y, BUTTON_SMALL, 0, BUTTON_DOWN

End Sub

Private Sub picHelp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picHelp, X, Y, BUTTON_SMALL, 0, BUTTON_MOVE

End Sub

Private Sub picHelp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RunHelp

    mouseEvent2 picHelp, X, Y, BUTTON_SMALL, 0, BUTTON_UP

End Sub

Public Sub SetColors()

    On Error Resume Next

    Dim c As Control

    frmSoldatMapEditor.picMenuBar.BackColor = bgColor
    frmSoldatMapEditor.picStatus.BackColor = bgColor
    frmSoldatMapEditor.picResize.BackColor = bgColor
    txtZoom.BackColor = bgColor
    txtZoom.ForeColor = lblTextClr
    picProgress.BackColor = bgColor
    lblFileName.BackColor = lblBackClr
    lblFileName.ForeColor = lblTextClr
    lblZoom.BackColor = lblBackClr
    lblZoom.ForeColor = lblTextClr
    lblCurrentTool.BackColor = lblBackClr
    lblCurrentTool.ForeColor = lblTextClr
    lblMousePosition.BackColor = lblBackClr
    lblMousePosition.ForeColor = lblTextClr

    SetFormFonts Me

End Sub

Private Sub picMaximize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picMaximize, X, Y, BUTTON_SMALL, (Me.Tag = vbNormal), BUTTON_DOWN

End Sub

Private Sub picMaximize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picMaximize, X, Y, BUTTON_SMALL, (Me.Tag = vbNormal), BUTTON_MOVE

End Sub

Private Sub picMaximize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Me.Tag = vbMaximized Then
        RestoreBorderLessForm
        picResize.Top = (Me.Height / Screen.TwipsPerPixelY) - picResize.Height
        picResize.Left = (Me.Width / Screen.TwipsPerPixelX) - picResize.Width
    Else
        MaximizeBorderLessForm
    End If

    picResize.Visible = Me.Tag = vbNormal

    mouseEvent2 picMaximize, X, Y, BUTTON_SMALL, (Me.Tag = vbNormal), BUTTON_UP

    resetDevice

End Sub

Private Sub picMinimize_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picMinimize, X, Y, BUTTON_SMALL, 0, BUTTON_DOWN

End Sub

Private Sub picMinimize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picMinimize, X, Y, BUTTON_SMALL, 0, BUTTON_MOVE

End Sub

Public Sub picMinimize_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picMinimize, X, Y, BUTTON_SMALL, 0, BUTTON_UP
    If mnuDisplay.Checked Then frmDisplay.Hide
    If mnuWaypoints.Checked Then frmWaypoints.Hide
    If mnuTools.Checked Then frmTools.Hide
    If mnuPalette.Checked Then frmPalette.Hide
    If mnuScenery.Checked Then frmScenery.Hide
    If mnuInfo.Checked Then frmInfo.Hide
    If mnuTexture.Checked Then frmTexture.Hide
    Me.Hide
    frmTaskBar.WindowState = vbMinimized

End Sub

Private Sub picExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picExit, X, Y, BUTTON_SMALL, 0, BUTTON_DOWN

End Sub

Private Sub picExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picExit, X, Y, BUTTON_SMALL, 0, BUTTON_MOVE

End Sub

Private Sub picExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picExit, X, Y, BUTTON_SMALL, 0, BUTTON_UP
    Terminate

End Sub

Private Sub picStatus_Click()

    If Me.Tag = vbMaximized Then
        Dim hwnd1 As Long
        hwnd1 = FindWindow("Shell_traywnd", "")
        Call SetWindowPos(hwnd1, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
    End If

End Sub

Private Sub picTitle_DblClick()

    If Me.Tag = vbMaximized Then
        RestoreBorderLessForm

        picResize.Top = (Me.Height / Screen.TwipsPerPixelY) - picResize.Height
        picResize.Left = (Me.Width / Screen.TwipsPerPixelX) - picResize.Width
    Else
        MaximizeBorderLessForm
    End If

    mouseEvent2 picMaximize, 0, 0, BUTTON_SMALL, (Me.Tag = vbNormal), BUTTON_UP

    picResize.Visible = Me.Tag = vbNormal

    resetDevice

End Sub

Private Sub AutoTexture()

    If (numSelectedPolys <= 0) Then
        Exit Sub
    End If

    Dim X As Single
    Dim Y As Single
    Dim Z As Single
    Dim vertIndex As Integer
    Dim i As Integer

    For i = 1 To 3
        If vertexList(selectedPolys(1)).vertex(i) > 0 Then
            vertIndex = i
        End If
    Next

    X = PolyCoords(selectedPolys(1)).vertex(vertIndex).X
    Y = PolyCoords(selectedPolys(1)).vertex(vertIndex).Y
    Z = Polys(selectedPolys(1)).vertex(vertIndex).Z

    numSelectedPolys = 0
    ReDim selectedPolys(0)

    SetTextureCoords X, Y, Z, 0, 0

    Render

End Sub

Private Sub SetTextureCoords(X As Single, Y As Single, Z As Single, tu As Single, tv As Single)

    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    For i = 0 To mPolyCount
        For j = 1 To 3
            ' if vertex is at these coords and not marked
            If Int(PolyCoords(i).vertex(j).X) = Int(X) _
                    And Int(PolyCoords(i).vertex(j).Y) = Int(Y) _
                    And Int(Polys(i).vertex(j).Z) = Int(Z) _
                    And vertexList(i).vertex(j) < 10 Then
                ' set its tex coords to these tex coords
                Polys(i).vertex(j).tu = tu
                Polys(i).vertex(j).tv = tv
                ' mark in vertex list
                vertexList(i).vertex(j) = 10
                ' find next vertex index
                k = j + 1
                If k > 3 Then k = 1
                ' check next vertex
                If vertexList(i).vertex(k) < 10 Then
                    ' calculate new tex coords

                    ' call this routine again with new coords & tex coords
                    SetTextureCoords PolyCoords(i).vertex(k).X, PolyCoords(i).vertex(k).Y, Polys(i).vertex(k).Z, 0, 0
                End If
            End If
        Next
    Next

    ' loop through all vertices to find vertices at this point, put into array
    ' set their coords
    ' set vertex list value to mark

    ' for each vertex at this point, find adjacent verts
    ' calc new coords, call this and set new coords?
    ' send new coords to this routine?
    ' call this routine on them

End Sub
