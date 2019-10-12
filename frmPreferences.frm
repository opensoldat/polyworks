VERSION 5.00
Begin VB.Form frmPreferences 
   BackColor       =   &H004A3C31&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6600
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6840
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   440
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   456
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picTopmost 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   4800
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   78
      TabStop         =   0   'False
      Tag             =   "4"
      Top             =   7440
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picScenery 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   4800
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   77
      TabStop         =   0   'False
      Tag             =   "4"
      Top             =   6960
      Width           =   240
   End
   Begin VB.ComboBox cboSkin 
      Appearance      =   0  'Flat
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
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Tag             =   "font1"
      Top             =   3600
      Width           =   2535
   End
   Begin VB.TextBox txtHotkey 
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
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
      Height          =   285
      Index           =   13
      Left            =   6240
      TabIndex        =   23
      Tag             =   "font1"
      Top             =   3720
      Width           =   255
   End
   Begin VB.TextBox txtHotkey 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
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
      Height          =   285
      Index           =   12
      Left            =   4800
      TabIndex        =   22
      Tag             =   "font1"
      Top             =   3720
      Width           =   255
   End
   Begin VB.PictureBox picSekrit 
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
      Left            =   120
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   74
      TabStop         =   0   'False
      Tag             =   "3"
      Top             =   6120
      Width           =   960
   End
   Begin VB.TextBox txtWayptKey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
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
      Height          =   285
      Index           =   4
      Left            =   6120
      TabIndex        =   28
      Tag             =   "font1"
      Text            =   "X"
      ToolTipText     =   "Fly"
      Top             =   5520
      Width           =   255
   End
   Begin VB.TextBox txtWayptKey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
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
      Height          =   285
      Index           =   3
      Left            =   5400
      TabIndex        =   27
      Tag             =   "font1"
      Text            =   "Z"
      ToolTipText     =   "Down"
      Top             =   5520
      Width           =   255
   End
   Begin VB.TextBox txtWayptKey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
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
      Height          =   285
      Index           =   2
      Left            =   5400
      TabIndex        =   24
      Tag             =   "font1"
      Text            =   "W"
      ToolTipText     =   "Up"
      Top             =   4800
      Width           =   255
   End
   Begin VB.TextBox txtWayptKey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
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
      Height          =   285
      Index           =   1
      Left            =   5760
      TabIndex        =   26
      Tag             =   "font1"
      Text            =   "S"
      ToolTipText     =   "Right"
      Top             =   5160
      Width           =   255
   End
   Begin VB.TextBox txtWayptKey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
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
      Height          =   285
      Index           =   0
      Left            =   5040
      TabIndex        =   25
      Tag             =   "font1"
      Text            =   "A"
      ToolTipText     =   "Left"
      Top             =   5160
      Width           =   255
   End
   Begin VB.PictureBox picPrefabs 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3960
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   71
      TabStop         =   0   'False
      Tag             =   "10"
      Top             =   5520
      Width           =   240
   End
   Begin VB.TextBox txtPrefabs 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   360
      TabIndex        =   9
      Tag             =   "font1"
      Top             =   5520
      Width           =   3495
   End
   Begin VB.TextBox txtUncomp 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   360
      TabIndex        =   8
      Tag             =   "font1"
      Top             =   4920
      Width           =   3495
   End
   Begin VB.PictureBox picUncomp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3960
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   68
      TabStop         =   0   'False
      Tag             =   "10"
      Top             =   4920
      Width           =   240
   End
   Begin VB.TextBox txtHotkey 
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
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
      Height          =   285
      Index           =   11
      Left            =   6240
      TabIndex        =   21
      Tag             =   "font1"
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox txtHotkey 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
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
      Height          =   285
      Index           =   10
      Left            =   4800
      TabIndex        =   20
      Tag             =   "font1"
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox txtHotkey 
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
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
      Height          =   285
      Index           =   9
      Left            =   6240
      TabIndex        =   19
      Tag             =   "font1"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox txtHotkey 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
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
      Height          =   285
      Index           =   8
      Left            =   4800
      TabIndex        =   18
      Tag             =   "font1"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox txtHotkey 
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
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
      Height          =   285
      Index           =   7
      Left            =   6240
      TabIndex        =   17
      Tag             =   "font1"
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox txtHotkey 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
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
      Height          =   285
      Index           =   6
      Left            =   4800
      TabIndex        =   16
      Tag             =   "font1"
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox txtHotkey 
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
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
      Height          =   285
      Index           =   5
      Left            =   6240
      TabIndex        =   15
      Tag             =   "font1"
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox txtHotkey 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
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
      Height          =   285
      Index           =   4
      Left            =   4800
      TabIndex        =   14
      Tag             =   "font1"
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox txtHotkey 
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
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
      Height          =   285
      Index           =   3
      Left            =   6240
      TabIndex        =   13
      Tag             =   "font1"
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox txtHotkey 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
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
      Height          =   285
      Index           =   2
      Left            =   4800
      TabIndex        =   12
      Tag             =   "font1"
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox txtHotkey 
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
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
      Height          =   285
      Index           =   1
      Left            =   6240
      TabIndex        =   11
      Tag             =   "font1"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox txtHotkey 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
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
      Height          =   285
      Index           =   0
      Left            =   4800
      TabIndex        =   10
      Tag             =   "font1"
      Top             =   840
      Width           =   255
   End
   Begin VB.PictureBox picHotkeys 
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3360
      Left            =   5160
      ScaleHeight     =   224
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   720
      Width           =   960
   End
   Begin VB.TextBox txtOpacity2 
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
      Left            =   2880
      TabIndex        =   5
      Tag             =   "font1"
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox txtOpacity1 
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
      Left            =   2880
      TabIndex        =   4
      Tag             =   "font1"
      Top             =   2280
      Width           =   495
   End
   Begin VB.PictureBox picGridClr2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3960
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   2640
      Width           =   255
   End
   Begin VB.PictureBox picFolder 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3960
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   33
      TabStop         =   0   'False
      Tag             =   "10"
      Top             =   4320
      Width           =   240
   End
   Begin VB.TextBox txtHeight 
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
      Left            =   1560
      TabIndex        =   1
      Tag             =   "font1"
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox txtWidth 
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
      Left            =   1560
      TabIndex        =   0
      Tag             =   "font1"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txtDir 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   360
      TabIndex        =   7
      Tag             =   "font1"
      Top             =   4320
      Width           =   3495
   End
   Begin VB.TextBox txtDivisions 
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
      Left            =   1560
      TabIndex        =   3
      Tag             =   "font1"
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox txtSpacing 
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
      Left            =   1560
      TabIndex        =   2
      Tag             =   "font1"
      Top             =   2280
      Width           =   495
   End
   Begin VB.PictureBox picGridClr 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3960
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   2280
      Width           =   255
   End
   Begin VB.PictureBox picApply 
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
      Left            =   5760
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   49
      TabStop         =   0   'False
      Tag             =   "2"
      Top             =   6120
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
      Left            =   3600
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   48
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   6120
      Width           =   960
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
      Left            =   4680
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   47
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   6120
      Width           =   960
   End
   Begin VB.PictureBox picBackClr 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3960
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   1680
      Width           =   255
   End
   Begin VB.PictureBox picPointClr 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3960
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   960
      Width           =   255
   End
   Begin VB.PictureBox picSelectionClr 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3960
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   1320
      Width           =   255
   End
   Begin VB.ComboBox cboWireSrc 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmPreferences.frx":0000
      Left            =   1200
      List            =   "frmPreferences.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   7080
      Width           =   1455
   End
   Begin VB.ComboBox cboWireDest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmPreferences.frx":0072
      Left            =   1200
      List            =   "frmPreferences.frx":008E
      Style           =   2  'Dropdown List
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   7440
      Width           =   1455
   End
   Begin VB.ComboBox cboPolyDest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmPreferences.frx":00E4
      Left            =   2760
      List            =   "frmPreferences.frx":0100
      Style           =   2  'Dropdown List
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   7440
      Width           =   1455
   End
   Begin VB.ComboBox cboPolySrc 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmPreferences.frx":0156
      Left            =   2760
      List            =   "frmPreferences.frx":0172
      Style           =   2  'Dropdown List
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   7080
      Width           =   1455
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
      ScaleWidth      =   456
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   0
      Width           =   6840
      Begin VB.PictureBox picHide 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H004A3C31&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   6600
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   35
         TabStop         =   0   'False
         Tag             =   "3"
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.Label lblPref 
      BackStyle       =   0  'Transparent
      Caption         =   "Fullscreen always on top"
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
      Index           =   24
      Left            =   5040
      TabIndex        =   80
      Tag             =   "font2"
      Top             =   7440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblPref 
      BackStyle       =   0  'Transparent
      Caption         =   "Use 4 verts for scenery"
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
      Index           =   23
      Left            =   5040
      TabIndex        =   79
      Tag             =   "font2"
      Top             =   6960
      Width           =   1575
   End
   Begin VB.Label lblOther 
      BackColor       =   &H004A3C31&
      Caption         =   "Other"
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
      Left            =   4800
      TabIndex        =   76
      Top             =   6600
      Width           =   735
   End
   Begin VB.Shape fraPref 
      BorderColor     =   &H000B3C0D&
      Height          =   1215
      Index           =   5
      Left            =   4560
      Top             =   6720
      Width           =   2175
   End
   Begin VB.Label lblPref 
      BackColor       =   &H004A3C31&
      BackStyle       =   0  'Transparent
      Caption         =   "Skin"
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
      Index           =   22
      Left            =   360
      TabIndex        =   75
      Tag             =   "font2"
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label lblWayKeys 
      Alignment       =   2  'Center
      BackColor       =   &H004A3C31&
      Caption         =   "Waypoint Keys"
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
      Left            =   4800
      TabIndex        =   73
      Tag             =   "font2"
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Shape fraPref 
      BorderColor     =   &H000B3C0D&
      Height          =   1455
      Index           =   3
      Left            =   4560
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label lblPref 
      BackColor       =   &H004A3C31&
      BackStyle       =   0  'Transparent
      Caption         =   "Prefabs"
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
      Index           =   17
      Left            =   360
      TabIndex        =   72
      Tag             =   "font2"
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label lblPref 
      BackColor       =   &H004A3C31&
      BackStyle       =   0  'Transparent
      Caption         =   "Uncompiled"
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
      Index           =   16
      Left            =   360
      TabIndex        =   70
      Tag             =   "font2"
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label lblPref 
      BackColor       =   &H004A3C31&
      BackStyle       =   0  'Transparent
      Caption         =   "Soldat"
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
      Left            =   360
      TabIndex        =   69
      Tag             =   "font2"
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label lblHotkeys 
      Alignment       =   2  'Center
      BackColor       =   &H004F3D31&
      Caption         =   "HotKeys"
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
      Left            =   4800
      TabIndex        =   66
      Tag             =   "font2"
      Top             =   360
      Width           =   975
   End
   Begin VB.Shape fraPref 
      BorderColor     =   &H000B3C0D&
      Height          =   3855
      Index           =   1
      Left            =   4560
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label lblPref 
      BackStyle       =   0  'Transparent
      Caption         =   "px"
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
      Index           =   12
      Left            =   2280
      TabIndex        =   65
      Tag             =   "font2"
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label lblPref 
      BackStyle       =   0  'Transparent
      Caption         =   "px"
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
      Index           =   11
      Left            =   2280
      TabIndex        =   64
      Tag             =   "font2"
      Top             =   960
      Width           =   375
   End
   Begin VB.Label lblPref 
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
      Index           =   15
      Left            =   3360
      TabIndex        =   63
      Tag             =   "font2"
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label lblPref 
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
      Index           =   14
      Left            =   3360
      TabIndex        =   62
      Tag             =   "font2"
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label lblPref 
      Alignment       =   2  'Center
      BackColor       =   &H004A3C31&
      BackStyle       =   0  'Transparent
      Caption         =   "Colours"
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
      Left            =   3000
      TabIndex        =   61
      Tag             =   "font2"
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblPref 
      BackColor       =   &H004A3C31&
      BackStyle       =   0  'Transparent
      Caption         =   "Window"
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
      TabIndex        =   59
      Tag             =   "font2"
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblPref 
      BackColor       =   &H00614B3D&
      Caption         =   "Height:"
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
      TabIndex        =   58
      Tag             =   "font2"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblPref 
      BackColor       =   &H00614B3D&
      Caption         =   "Width:"
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
      TabIndex        =   57
      Tag             =   "font2"
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblPref 
      BackStyle       =   0  'Transparent
      Caption         =   "px"
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
      Index           =   13
      Left            =   2160
      TabIndex        =   56
      Tag             =   "font2"
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label lblPref 
      BackColor       =   &H00614B3D&
      Caption         =   "Divisions:"
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
      TabIndex        =   55
      Tag             =   "font2"
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblPref 
      BackColor       =   &H00614B3D&
      Caption         =   "Spacing:"
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
      TabIndex        =   54
      Tag             =   "font2"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label lblPref 
      BackColor       =   &H004F3D31&
      BackStyle       =   0  'Transparent
      Caption         =   "Grid"
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
      Left            =   360
      TabIndex        =   52
      Tag             =   "font2"
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label lblPref 
      BackColor       =   &H00614B3D&
      BackStyle       =   0  'Transparent
      Caption         =   "Wireframe"
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
      Index           =   18
      Left            =   1200
      TabIndex        =   51
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Label lblPref 
      BackColor       =   &H00614B3D&
      BackStyle       =   0  'Transparent
      Caption         =   "Polygon"
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
      Index           =   19
      Left            =   2880
      TabIndex        =   50
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Label lblDirs 
      Alignment       =   2  'Center
      BackColor       =   &H004A3C31&
      Caption         =   "Directories"
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
      Left            =   360
      TabIndex        =   46
      Tag             =   "font2"
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Shape fraPref 
      BorderColor     =   &H000B3C0D&
      Height          =   2655
      Index           =   2
      Left            =   120
      Top             =   3360
      Width           =   4335
   End
   Begin VB.Label lblDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H004F3D31&
      Caption         =   "Display"
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
      Left            =   360
      TabIndex        =   45
      Tag             =   "font2"
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lblPref 
      BackColor       =   &H00614B3D&
      Caption         =   "Pattern:"
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
      Left            =   2880
      TabIndex        =   44
      Tag             =   "font2"
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblPref 
      BackColor       =   &H00614B3D&
      Caption         =   "SRC:"
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
      Index           =   20
      Left            =   360
      TabIndex        =   43
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label lblPref 
      BackColor       =   &H00614B3D&
      Caption         =   "DEST:"
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
      Index           =   21
      Left            =   360
      TabIndex        =   42
      Top             =   7440
      Width           =   735
   End
   Begin VB.Label lblPref 
      BackColor       =   &H00614B3D&
      Caption         =   "Back:"
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
      Left            =   2880
      TabIndex        =   41
      Tag             =   "font2"
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblPref 
      BackColor       =   &H00614B3D&
      Caption         =   "Point:"
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
      Left            =   2880
      TabIndex        =   40
      Tag             =   "font2"
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblBlending 
      BackColor       =   &H004A3C31&
      Caption         =   "Blending"
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
      Left            =   360
      TabIndex        =   39
      Top             =   6600
      Width           =   975
   End
   Begin VB.Shape fraPref 
      BorderColor     =   &H000B3C0D&
      Height          =   2655
      Index           =   0
      Left            =   120
      Top             =   480
      Width           =   4335
   End
   Begin VB.Shape fraPref 
      BorderColor     =   &H000B3C0D&
      Height          =   1215
      Index           =   4
      Left            =   120
      Top             =   6720
      Width           =   4335
   End
End
Attribute VB_Name = "frmPreferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TColour
    red As Byte
    green As Byte
    blue As Byte
End Type

Dim blendModes(0 To 7) As Integer

Dim backClr As TColour
Dim pointClr As TColour
Dim selClr As TColour
Dim gridClr As TColour
Dim gridClr2 As TColour

Dim spacing As Integer, divisions As Integer
Dim formWidth As Integer, formHeight As Integer
Dim opacity1 As Integer, opacity2 As Integer
Dim sceneryVerts As Boolean, topmost As Boolean

Private Sub picHide_Click()

    Me.ScaleHeight = 408
    Unload Me
    frmSoldatMapEditor.RegainFocus

End Sub

Private Sub picSekrit_Click()

    If Me.ScaleHeight < 460 Then
        Me.Height = 544 * Screen.TwipsPerPixelY
    Else
        Me.Height = 440 * Screen.TwipsPerPixelY
    End If

End Sub

Private Sub picTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ReleaseCapture
    SendMessage Me.hWnd, WM_NCLBUTTONDOWN, 2, 0&

End Sub

Private Sub picApply_Click()

    applyPreferences

End Sub

Private Sub picCancel_Click()

    Me.ScaleHeight = 408
    Unload Me
    frmSoldatMapEditor.RegainFocus

End Sub

Private Function applyPreferences() As Boolean

    Dim i As Integer

    On Error GoTo ErrorHandler

    mouseEvent2 picHide, 0, 0, BUTTON_SMALL, 0, BUTTON_UP
    mouseEvent2 picOK, 0, 0, BUTTON_LARGE, 0, BUTTON_UP
    mouseEvent2 picCancel, 0, 0, BUTTON_LARGE, 0, BUTTON_UP
    mouseEvent2 picSekrit, 0, 0, BUTTON_LARGE, 0, BUTTON_UP
    mouseEvent2 picApply, 0, 0, BUTTON_LARGE, 0, BUTTON_UP
    mouseEvent2 picFolder, 0, 0, BUTTON_SMALL, 0, BUTTON_UP

    If right(txtDir.Text, 1) <> "\" Then txtDir.Text = txtDir.Text + "\"

    If Len(Dir(txtDir.Text, vbDirectory)) <> 0 And frmSoldatMapEditor.soldatDir <> txtDir.Text Then
        If Not Len(Dir(txtDir.Text & "Maps\", vbDirectory)) <> 0 Then
            MsgBox "'Maps' folder does not exist in Soldat directory."
            Exit Function
        End If
        If Not Len(Dir(txtDir.Text & "Textures\", vbDirectory)) <> 0 Then
            MsgBox "'Textures' folder does not exist in Soldat directory."
            Exit Function
        End If
        If Not Len(Dir(txtDir.Text & "Scenery-gfx\", vbDirectory)) <> 0 Then
            MsgBox "'Scenery-gfx' folder does not exist in Soldat directory."
            Exit Function
        End If

        frmSoldatMapEditor.soldatDir = txtDir.Text
    ElseIf Len(Dir(txtDir.Text, vbDirectory)) = 0 Then
        MsgBox "Soldat directory does not exist."
        Exit Function
    End If

    If right(txtUncomp.Text, 1) <> "\" Then txtUncomp.Text = txtUncomp.Text + "\"

    If Len(Dir(txtUncomp.Text, vbDirectory)) <> 0 Then
        frmSoldatMapEditor.uncompDir = txtUncomp.Text
    Else
        MsgBox "Uncompiled Maps directory does not exist."
        Exit Function
    End If

    If right(txtPrefabs.Text, 1) <> "\" Then txtPrefabs.Text = txtPrefabs.Text + "\"

    If Len(Dir(txtPrefabs.Text, vbDirectory)) <> 0 Then
        frmSoldatMapEditor.prefabDir = txtPrefabs.Text
    Else
        MsgBox "Prefabs Maps directory does not exist."
        Exit Function
    End If

    frmSoldatMapEditor.wireBlendSrc = blendModes(cboWireSrc.ListIndex)
    frmSoldatMapEditor.wireBlendDest = blendModes(cboWireDest.ListIndex)
    frmSoldatMapEditor.polyBlendSrc = blendModes(cboPolySrc.ListIndex)
    frmSoldatMapEditor.polyBlendDest = blendModes(cboPolyDest.ListIndex)

    frmSoldatMapEditor.backClr = RGB(backClr.blue, backClr.green, backClr.red)
    frmSoldatMapEditor.pointClr = RGB(pointClr.blue, pointClr.green, pointClr.red)
    frmSoldatMapEditor.selectionClr = RGB(selClr.blue, selClr.green, selClr.red)
    frmSoldatMapEditor.gridClr = RGB(gridClr.blue, gridClr.green, gridClr.red)
    frmSoldatMapEditor.gridClr2 = RGB(gridClr2.blue, gridClr2.green, gridClr2.red)

    frmSoldatMapEditor.formWidth = formWidth
    frmSoldatMapEditor.formHeight = formHeight
    If frmSoldatMapEditor.WindowState = vbNormal Then
        frmSoldatMapEditor.Width = formWidth * Screen.TwipsPerPixelX
        frmSoldatMapEditor.Height = formHeight * Screen.TwipsPerPixelY
    ElseIf frmSoldatMapEditor.WindowState = vbMaximized Then
        frmSoldatMapEditor.WindowState = vbNormal
        frmSoldatMapEditor.Width = formWidth * Screen.TwipsPerPixelX
        frmSoldatMapEditor.Height = formHeight * Screen.TwipsPerPixelY
        frmSoldatMapEditor.WindowState = vbMaximized
    End If

    frmSoldatMapEditor.gridSpacing = spacing
    frmSoldatMapEditor.gridDivisions = divisions
    frmSoldatMapEditor.gridOp1 = opacity1 / 100 * 255
    frmSoldatMapEditor.gridOp2 = opacity2 / 100 * 255

    frmSoldatMapEditor.sceneryVerts = sceneryVerts
    frmSoldatMapEditor.topmost = topmost

    For i = 0 To 13
        frmTools.setHotKey i, MapVirtualKey(CInt(txtHotkey(i).Tag), 0)
        frmTools.picTools(i).ToolTipText = frmTools.picTools(i).Tag & " (" & (txtHotkey(i).Text) & ")"
    Next

    For i = 0 To 4
        frmWaypoints.setWayptKey i, MapVirtualKey(CInt(txtWayptKey(i).Tag), 0)
        frmWaypoints.picType(i).ToolTipText = " (" & (txtWayptKey(i).Text) & ")"
        frmWaypoints.lblType(i).ToolTipText = " (" & (txtWayptKey(i).Text) & ")"
    Next

    If cboSkin.List(cboSkin.ListIndex) <> gfxDir Then
        gfxDir = cboSkin.List(cboSkin.ListIndex)
        frmSoldatMapEditor.loadColours
        frmSoldatMapEditor.SetColours
        frmSoldatMapEditor.initGfx
        frmColour.SetColours
        frmDisplay.SetColours
        frmInfo.SetColours
        frmMap.SetColours
        frmPalette.SetColours
        frmPreferences.SetColours
        frmScenery.SetColours
        frmSoldatMapEditor.SetColours
        frmTexture.SetColours
        frmTools.SetColours
        frmWaypoints.SetColours
        frmDisplay.refreshButtons
    End If

    frmSoldatMapEditor.setPreferences

    applyPreferences = True

    Exit Function

ErrorHandler:

    MsgBox "Error applying preferences" & vbNewLine & Error$

End Function

Private Sub picOK_Click()

    Me.ScaleHeight = 408
    Me.Hide
    If applyPreferences Then
        Unload Me
        frmSoldatMapEditor.RegainFocus
    Else
        Me.Show
    End If

End Sub

Private Sub Form_Load()

    Dim i As Integer

    On Error GoTo ErrorHandler

    sceneryVerts = frmSoldatMapEditor.sceneryVerts
    topmost = frmSoldatMapEditor.topmost

    Me.SetColours

    blendModes(0) = 1
    blendModes(1) = 2
    blendModes(2) = 3
    blendModes(3) = 4
    blendModes(4) = 9
    blendModes(5) = 10
    blendModes(6) = 5
    blendModes(7) = 6

    backClr = getRGB(frmSoldatMapEditor.backClr)
    pointClr = getRGB(frmSoldatMapEditor.pointClr)
    selClr = getRGB(frmSoldatMapEditor.selectionClr)
    gridClr = getRGB(frmSoldatMapEditor.gridClr)
    gridClr2 = getRGB(frmSoldatMapEditor.gridClr2)

    For i = 0 To 7
        If frmSoldatMapEditor.wireBlendSrc = blendModes(i) Then cboWireSrc.ListIndex = i
        If frmSoldatMapEditor.wireBlendDest = blendModes(i) Then cboWireDest.ListIndex = i
        If frmSoldatMapEditor.polyBlendSrc = blendModes(i) Then cboPolySrc.ListIndex = i
        If frmSoldatMapEditor.polyBlendDest = blendModes(i) Then cboPolyDest.ListIndex = i
    Next

    Me.picBackClr.BackColor = RGB(backClr.red, backClr.green, backClr.blue)
    Me.picPointClr.BackColor = RGB(pointClr.red, pointClr.green, pointClr.blue)
    Me.picSelectionClr.BackColor = RGB(selClr.red, selClr.green, selClr.blue)
    Me.picGridClr.BackColor = RGB(gridClr.red, gridClr.green, gridClr.blue)
    Me.picGridClr2.BackColor = RGB(gridClr2.red, gridClr2.green, gridClr2.blue)

    txtWidth.Text = frmSoldatMapEditor.formWidth
    txtHeight.Text = frmSoldatMapEditor.formHeight
    formWidth = txtWidth.Text
    formHeight = txtHeight.Text

    txtSpacing.Text = frmSoldatMapEditor.gridSpacing
    txtDivisions.Text = frmSoldatMapEditor.gridDivisions
    spacing = txtSpacing.Text
    divisions = txtDivisions.Text
    opacity1 = frmSoldatMapEditor.gridOp1 / 255 * 100
    txtOpacity1.Text = opacity1
    opacity2 = frmSoldatMapEditor.gridOp2 / 255 * 100
    txtOpacity2.Text = opacity2

    For i = 0 To 13
        txtHotkey(i).Text = Chr$(MapVirtualKey(frmTools.getHotKey(i), 1))
        txtHotkey(i).Tag = Asc(txtHotkey(i).Text)
    Next

    For i = 0 To 4
        txtWayptKey(i).Text = Chr$(MapVirtualKey(frmWaypoints.getWayptKey(i), 1))
        txtWayptKey(i).Tag = Asc(txtWayptKey(i).Text)
    Next

    Dim file As Variant

    file = Dir$(appPath & "\*.*", vbDirectory)
    Do While Len(file)
        If FileExists(appPath & "\" & file & "\colours.ini") Then
            cboSkin.AddItem file
            If file = gfxDir Then cboSkin.ListIndex = cboSkin.ListCount - 1
        End If
        file = Dir$
    Loop

    txtDir.Text = frmSoldatMapEditor.soldatDir
    txtUncomp.Text = frmSoldatMapEditor.uncompDir
    txtPrefabs.Text = frmSoldatMapEditor.prefabDir

    Exit Sub
ErrorHandler:
    MsgBox Error$ & vbNewLine & "Error loading Preferences form"

End Sub

Private Function FileExists(fileName As String) As Boolean

    On Error GoTo ErrorHandler

    FileExists = FileLen(fileName) > 0

ErrorHandler:

End Function


Private Sub picPointClr_Click()

    frmColour.InitClr pointClr.red, pointClr.green, pointClr.blue
    frmColour.Show 1
    picPointClr.BackColor = RGB(frmColour.red, frmColour.green, frmColour.blue)
    pointClr.red = frmColour.red
    pointClr.green = frmColour.green
    pointClr.blue = frmColour.blue

End Sub

Private Sub picSelectionClr_Click()

    frmColour.InitClr selClr.red, selClr.green, selClr.blue
    frmColour.Show 1
    picSelectionClr.BackColor = RGB(frmColour.red, frmColour.green, frmColour.blue)
    selClr.red = frmColour.red
    selClr.green = frmColour.green
    selClr.blue = frmColour.blue

End Sub

Private Sub picBackClr_Click()

    frmColour.InitClr backClr.red, backClr.green, backClr.blue
    frmColour.Show 1
    picBackClr.BackColor = RGB(frmColour.red, frmColour.green, frmColour.blue)
    backClr.red = frmColour.red
    backClr.green = frmColour.green
    backClr.blue = frmColour.blue

End Sub

Private Sub picGridClr_Click()

    frmColour.InitClr gridClr.red, gridClr.green, gridClr.blue
    frmColour.Show 1
    picGridClr.BackColor = RGB(frmColour.red, frmColour.green, frmColour.blue)
    gridClr.red = frmColour.red
    gridClr.green = frmColour.green
    gridClr.blue = frmColour.blue

End Sub

Private Sub picGridClr2_Click()

    frmColour.InitClr gridClr2.red, gridClr2.green, gridClr2.blue
    frmColour.Show 1
    picGridClr2.BackColor = RGB(frmColour.red, frmColour.green, frmColour.blue)
    gridClr2.red = frmColour.red
    gridClr2.green = frmColour.green
    gridClr2.blue = frmColour.blue

End Sub

Private Sub picFolder_Click()

    Dim folder As String

    folder = SelectFolder(Me)

    If right(folder, 1) <> "\" Then folder = folder & "\"

    If Len(folder) > 1 Then
        txtDir.Text = folder
    End If

    mouseEvent2 picFolder, 0, 0, BUTTON_SMALL, 0, BUTTON_UP

End Sub

Private Sub picUncomp_Click()

    Dim folder As String

    folder = SelectFolder(Me)

    If right(folder, 1) <> "\" Then folder = folder & "\"

    If Len(folder) > 1 Then
        txtUncomp.Text = folder
    End If

    mouseEvent2 picUncomp, 0, 0, BUTTON_SMALL, 0, BUTTON_UP

End Sub

Private Sub picPrefabs_Click()

    Dim folder As String

    folder = SelectFolder(Me)

    If right(folder, 1) <> "\" Then folder = folder & "\"

    If Len(folder) > 1 Then
        txtPrefabs.Text = folder
    End If

    mouseEvent2 picPrefabs, 0, 0, BUTTON_SMALL, 0, BUTTON_UP

End Sub

Private Sub txtHotkey_GotFocus(Index As Integer)

    txtHotkey(Index).SelStart = 0
    txtHotkey(Index).SelLength = Len(txtHotkey(Index).Text)

End Sub

Private Sub txtHotkey_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    txtHotkey(Index).Tag = KeyCode

End Sub

Private Sub txtHotkey_KeyPress(Index As Integer, KeyAscii As Integer)

    txtHotkey(Index).Text = UCase$(Chr$(KeyAscii))
    KeyAscii = 0

End Sub

Private Sub txtWayptKey_GotFocus(Index As Integer)

    txtWayptKey(Index).SelStart = 0
    txtWayptKey(Index).SelLength = Len(txtWayptKey(Index).Text)

End Sub

Private Sub txtWayptKey_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    txtWayptKey(Index).Tag = KeyCode

End Sub

Private Sub txtWayptKey_KeyPress(Index As Integer, KeyAscii As Integer)

    txtWayptKey(Index).Text = UCase$(Chr$(KeyAscii))
    KeyAscii = 0

End Sub

Private Sub txtSpacing_Change()

    If IsNumeric(txtSpacing.Text) = False And txtSpacing.Text <> "" Then
        txtSpacing.Text = spacing
    End If

End Sub

Private Sub txtSpacing_GotFocus()

    txtSpacing.SelStart = 0
    txtSpacing.SelLength = Len(txtSpacing.Text)

End Sub

Private Sub txtSpacing_LostFocus()

    If IsNumeric(txtSpacing.Text) = False And txtSpacing.Text <> "" Then
        txtSpacing.Text = spacing
    ElseIf txtSpacing.Text = "" Then
        txtSpacing.Text = spacing
    ElseIf txtSpacing.Text >= 10 And txtSpacing.Text <= 100 Then
        spacing = Int(txtSpacing.Text)
        txtSpacing.Text = spacing
    Else
        If txtSpacing.Text < 10 Then spacing = 10
        If txtSpacing.Text > 100 Then spacing = 100
        txtSpacing.Text = spacing
    End If

End Sub

Private Sub txtDivisions_Change()

    If IsNumeric(txtDivisions.Text) = False And txtDivisions.Text <> "" Then
        txtDivisions.Text = divisions
    End If

End Sub

Private Sub txtDivisions_GotFocus()

    txtDivisions.SelStart = 0
    txtDivisions.SelLength = Len(txtDivisions.Text)

End Sub

Private Sub txtDivisions_LostFocus()

    If IsNumeric(txtDivisions.Text) = False And txtDivisions.Text <> "" Then
        txtDivisions.Text = divisions
    ElseIf txtDivisions.Text = "" Then
        txtDivisions.Text = divisions
    ElseIf txtDivisions.Text >= 1 And txtDivisions.Text <= 10 Then
        divisions = Int(txtDivisions.Text)
        txtDivisions = divisions
    Else
        If txtDivisions.Text < 1 Then divisions = 1
        If txtDivisions.Text > Int(spacing / 4) Then divisions = Int(spacing / 4)
        txtDivisions.Text = divisions
    End If

End Sub

Private Sub txtOpacity1_Change()

    If IsNumeric(txtOpacity1.Text) = False And txtOpacity1.Text <> "" Then
        txtOpacity1.Text = opacity1
    End If

End Sub

Private Sub txtOpacity1_GotFocus()

    txtOpacity1.SelStart = 0
    txtOpacity1.SelLength = Len(txtOpacity1.Text)

End Sub

Private Sub txtOpacity1_LostFocus()

    If IsNumeric(txtOpacity1.Text) = False And txtOpacity1.Text <> "" Then
        txtOpacity1.Text = opacity1
    ElseIf txtOpacity1.Text = "" Then
        txtOpacity1.Text = opacity1
    ElseIf txtOpacity1.Text >= 10 And txtOpacity1.Text <= 100 Then
        opacity1 = Int(txtOpacity1.Text)
        txtOpacity1.Text = opacity1
    Else
        txtOpacity1.Text = opacity1
    End If

End Sub

Private Sub txtOpacity2_Change()

    If IsNumeric(txtOpacity2.Text) = False And txtOpacity2.Text <> "" Then
        txtOpacity2.Text = opacity2
    End If

End Sub

Private Sub txtOpacity2_GotFocus()

    txtOpacity2.SelStart = 0
    txtOpacity2.SelLength = Len(txtOpacity2.Text)

End Sub

Private Sub txtOpacity2_LostFocus()

    If IsNumeric(txtOpacity2.Text) = False And txtOpacity2.Text <> "" Then
        txtOpacity2.Text = opacity2
    ElseIf txtOpacity2.Text = "" Then
        txtOpacity2.Text = opacity2
    ElseIf txtOpacity2.Text >= 10 And txtOpacity2.Text <= 100 Then
        opacity2 = Int(txtOpacity2.Text)
        txtOpacity2.Text = opacity2
    Else
        txtOpacity2.Text = opacity2
    End If

End Sub

Private Sub txtWidth_Change()

    If IsNumeric(txtWidth.Text) = False And txtWidth.Text <> "" Then
        txtWidth.Text = formWidth
    End If

End Sub

Private Sub txtWidth_GotFocus()

    txtWidth.SelStart = 0
    txtWidth.SelLength = Len(txtWidth.Text)

End Sub

Private Sub txtWidth_LostFocus()

    If IsNumeric(txtWidth.Text) = False And txtWidth.Text <> "" Then
        txtWidth.Text = formWidth
    ElseIf txtWidth.Text = "" Then
        txtWidth.Text = formWidth
    ElseIf txtWidth.Text >= 320 And txtWidth.Text <= Screen.Width / Screen.TwipsPerPixelX Then
        formWidth = Int(txtWidth.Text)
        txtWidth.Text = formWidth
    Else
        If txtWidth.Text < 320 Then formWidth = 320
        If txtWidth.Text > (Screen.Width / Screen.TwipsPerPixelX) Then formWidth = (Screen.Width / Screen.TwipsPerPixelX)
        txtWidth.Text = formWidth
    End If

End Sub

Private Sub txtHeight_Change()

    If IsNumeric(txtHeight.Text) = False And txtHeight.Text <> "" Then
        txtHeight.Text = formHeight
    End If

End Sub

Private Sub txtHeight_GotFocus()

    txtHeight.SelStart = 0
    txtHeight.SelLength = Len(txtHeight.Text)

End Sub

Private Sub txtHeight_LostFocus()

    If IsNumeric(txtHeight.Text) = False And txtHeight.Text <> "" Then
        txtHeight.Text = formHeight
    ElseIf txtHeight.Text = "" Then
        txtHeight.Text = formHeight
    ElseIf txtHeight.Text >= 320 And txtHeight.Text <= Screen.Height / Screen.TwipsPerPixelY Then
        formHeight = Int(txtHeight.Text)
        txtHeight.Text = formHeight
    Else
        If txtHeight.Text < 320 Then formHeight = 320
        If txtHeight.Text > (Screen.Height / Screen.TwipsPerPixelY) Then formHeight = (Screen.Height / Screen.TwipsPerPixelY)
        txtHeight.Text = formHeight
    End If

End Sub

Private Function getRGB(DecValue As Long) As TColour

    Dim hexValue As String

    hexValue = Hex(Val(DecValue))

    If Len(hexValue) < 6 Then
        hexValue = String(6 - Len(hexValue), "0") + hexValue
    End If

    getRGB.red = CLng("&H" + mid(hexValue, 1, 2))
    getRGB.green = CLng("&H" + mid(hexValue, 3, 2))
    getRGB.blue = CLng("&H" + mid(hexValue, 5, 2))

End Function

Private Sub picSekrit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picSekrit, X, Y, BUTTON_LARGE, 0, BUTTON_DOWN

End Sub

Private Sub picSekrit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picSekrit, X, Y, BUTTON_LARGE, 0, BUTTON_MOVE

End Sub

Private Sub picCancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picCancel, X, Y, BUTTON_LARGE, 0, BUTTON_DOWN

End Sub

Private Sub picCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picCancel, X, Y, BUTTON_LARGE, 0, BUTTON_MOVE

End Sub

Private Sub picOK_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picOK, X, Y, BUTTON_LARGE, 0, BUTTON_DOWN

End Sub

Private Sub picOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picOK, X, Y, BUTTON_LARGE, 0, BUTTON_MOVE

End Sub

Private Sub picApply_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picApply, X, Y, BUTTON_LARGE, 0, BUTTON_DOWN

End Sub

Private Sub picApply_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picApply, X, Y, BUTTON_LARGE, 0, BUTTON_MOVE

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

Private Sub picfolder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picFolder, X, Y, BUTTON_SMALL, 0, BUTTON_DOWN

End Sub

Private Sub picfolder_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picFolder, X, Y, BUTTON_SMALL, 0, BUTTON_MOVE

End Sub

Private Sub picUncomp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picUncomp, X, Y, BUTTON_SMALL, 0, BUTTON_DOWN

End Sub

Private Sub picUncomp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picUncomp, X, Y, BUTTON_SMALL, 0, BUTTON_MOVE

End Sub

Private Sub picPrefabs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picPrefabs, X, Y, BUTTON_SMALL, 0, BUTTON_DOWN

End Sub

Private Sub picPrefabs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picPrefabs, X, Y, BUTTON_SMALL, 0, BUTTON_MOVE

End Sub

Private Sub picScenery_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picScenery, X, Y, BUTTON_SMALL, sceneryVerts, BUTTON_DOWN

End Sub

Private Sub picScenery_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picScenery, X, Y, BUTTON_SMALL, sceneryVerts, BUTTON_MOVE

End Sub

Private Sub picScenery_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    sceneryVerts = Not sceneryVerts
    mouseEvent2 picScenery, X, Y, BUTTON_SMALL, sceneryVerts, BUTTON_UP

End Sub

Private Sub picTopmost_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picTopmost, X, Y, BUTTON_SMALL, topmost, BUTTON_DOWN

End Sub

Private Sub picTopmost_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    mouseEvent2 picTopmost, X, Y, BUTTON_SMALL, topmost, BUTTON_MOVE

End Sub

Private Sub picTopmost_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    topmost = Not topmost
    mouseEvent2 picTopmost, X, Y, BUTTON_SMALL, topmost, BUTTON_UP

End Sub

Public Sub SetColours()

    On Error Resume Next

    Dim i As Integer
    Dim c As Control

    '--------

    picTitle.Picture = LoadPicture(appPath & "\" & gfxDir & "\titlebar_preferences.bmp")
    picHotkeys.Picture = LoadPicture(appPath & "\" & gfxDir & "\tools.bmp")

    mouseEvent2 picHide, 0, 0, BUTTON_SMALL, 0, BUTTON_UP
    mouseEvent2 picOK, 0, 0, BUTTON_LARGE, 0, BUTTON_UP
    mouseEvent2 picCancel, 0, 0, BUTTON_LARGE, 0, BUTTON_UP
    mouseEvent2 picSekrit, 0, 0, BUTTON_LARGE, 0, BUTTON_UP
    mouseEvent2 picApply, 0, 0, BUTTON_LARGE, 0, BUTTON_UP
    mouseEvent2 picFolder, 0, 0, BUTTON_SMALL, 0, BUTTON_UP
    mouseEvent2 picUncomp, 0, 0, BUTTON_SMALL, 0, BUTTON_UP
    mouseEvent2 picPrefabs, 0, 0, BUTTON_SMALL, 0, BUTTON_UP
    mouseEvent2 picScenery, 0, 0, BUTTON_SMALL, sceneryVerts, BUTTON_UP
    mouseEvent2 picTopmost, 0, 0, BUTTON_SMALL, topmost, BUTTON_UP

    '--------

    Me.BackColor = bgClr

    For i = 0 To 22
        lblPref(i).BackColor = lblBackClr
        lblPref(i).ForeColor = lblTextClr
    Next

    lblDisplay.BackColor = bgClr
    lblDisplay.ForeColor = lblTextClr
    lblHotkeys.BackColor = bgClr
    lblHotkeys.ForeColor = lblTextClr
    lblDirs.BackColor = bgClr
    lblDirs.ForeColor = lblTextClr
    lblWayKeys.BackColor = bgClr
    lblWayKeys.ForeColor = lblTextClr
    lblBlending.BackColor = bgClr
    lblBlending.ForeColor = lblTextClr

    For i = 0 To 13
        txtHotkey(i).BackColor = bgClr
        txtHotkey(i).ForeColor = lblTextClr
    Next

    For i = 0 To 4
        txtWayptKey(i).BackColor = bgClr
        txtWayptKey(i).ForeColor = lblTextClr
        fraPref(i).BorderColor = frameClr
    Next

    txtWidth.BackColor = txtBackClr
    txtWidth.ForeColor = txtTextClr
    txtHeight.BackColor = txtBackClr
    txtHeight.ForeColor = txtTextClr
    txtSpacing.BackColor = txtBackClr
    txtSpacing.ForeColor = txtTextClr
    txtDivisions.BackColor = txtBackClr
    txtDivisions.ForeColor = txtTextClr
    txtOpacity1.BackColor = txtBackClr
    txtOpacity1.ForeColor = txtTextClr
    txtOpacity2.BackColor = txtBackClr
    txtOpacity2.ForeColor = txtTextClr
    txtDir.BackColor = txtBackClr
    txtDir.ForeColor = txtTextClr
    txtUncomp.BackColor = txtBackClr
    txtUncomp.ForeColor = txtTextClr
    txtPrefabs.BackColor = txtBackClr
    txtPrefabs.ForeColor = txtTextClr

    cboWireSrc.BackColor = txtBackClr
    cboWireSrc.ForeColor = txtTextClr
    cboWireDest.BackColor = txtBackClr
    cboWireDest.ForeColor = txtTextClr
    cboPolySrc.BackColor = txtBackClr
    cboPolySrc.ForeColor = txtTextClr
    cboPolyDest.BackColor = txtBackClr
    cboPolyDest.ForeColor = txtTextClr

    cboSkin.BackColor = txtBackClr
    cboSkin.ForeColor = txtTextClr

    For Each c In Me.Controls
        If c.Tag = "font1" Then
            c.Font.Name = font1
        ElseIf c.Tag = "font2" Then
            c.Font.Name = font2
        End If
    Next

End Sub
