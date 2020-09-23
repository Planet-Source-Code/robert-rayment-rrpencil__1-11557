VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "RRPencil"
   ClientHeight    =   6765
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   Icon            =   "Pencil.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   451
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame frmResize 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   4530
      TabIndex        =   175
      Top             =   870
      Width           =   435
      Begin VB.OptionButton optResize 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":0E42
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   60
         Picture         =   "Pencil.frx":0FC0
         Style           =   1  'Graphical
         TabIndex        =   177
         Top             =   420
         Width           =   315
      End
      Begin VB.OptionButton optResize 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":113E
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   60
         Picture         =   "Pencil.frx":12BC
         Style           =   1  'Graphical
         TabIndex        =   176
         Top             =   120
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdCanHt 
      BackColor       =   &H00C0C0C0&
      Caption         =   "720"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   1770
      Style           =   1  'Graphical
      TabIndex        =   167
      Top             =   5730
      Width           =   375
   End
   Begin VB.CommandButton cmdCanHt 
      BackColor       =   &H00C0C0C0&
      Caption         =   "520"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   1755
      Style           =   1  'Graphical
      TabIndex        =   166
      Top             =   5490
      Width           =   375
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   870
      LargeChange     =   10
      Left            =   1770
      Max             =   -200
      Min             =   4
      TabIndex        =   164
      Top             =   5955
      Width           =   375
   End
   Begin VB.Frame frmRotate 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   4515
      TabIndex        =   146
      Top             =   0
      Width           =   435
      Begin VB.OptionButton optRotate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":143A
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   60
         Picture         =   "Pencil.frx":15B8
         Style           =   1  'Graphical
         TabIndex        =   148
         Top             =   120
         Width           =   315
      End
      Begin VB.OptionButton optRotate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":1736
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   60
         Picture         =   "Pencil.frx":18B4
         Style           =   1  'Graphical
         TabIndex        =   147
         Top             =   420
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Reset"
      Height          =   375
      Left            =   11325
      Style           =   1  'Graphical
      TabIndex        =   142
      Top             =   7530
      Width           =   630
   End
   Begin VB.Frame frmBricksTiles 
      BackColor       =   &H00C0C0C0&
      Caption         =   "    AirBrush  "
      ForeColor       =   &H00FF0000&
      Height          =   1515
      Left            =   120
      TabIndex        =   133
      Top             =   1620
      Width           =   1215
      Begin ComCtl2.UpDown UDmx 
         Height          =   315
         Left            =   840
         TabIndex        =   136
         Top             =   660
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   556
         _Version        =   327681
         Value           =   12
         BuddyControl    =   "txtmx"
         BuddyDispid     =   196618
         OrigLeft        =   750
         OrigTop         =   240
         OrigRight       =   945
         OrigBottom      =   570
         Max             =   20
         Min             =   -20
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtmy 
         Height          =   315
         Left            =   360
         TabIndex        =   135
         Text            =   "txtmy"
         Top             =   1020
         Width           =   375
      End
      Begin VB.TextBox txtmx 
         Height          =   315
         Left            =   360
         TabIndex        =   134
         Text            =   "txtmx"
         Top             =   660
         Width           =   375
      End
      Begin ComCtl2.UpDown UDmy 
         Height          =   315
         Left            =   840
         TabIndex        =   137
         Top             =   1020
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   556
         _Version        =   327681
         Value           =   8
         BuddyControl    =   "txtmy"
         BuddyDispid     =   196617
         OrigLeft        =   750
         OrigTop         =   690
         OrigRight       =   945
         OrigBottom      =   1020
         Max             =   20
         Min             =   -20
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label LabABmy 
         BackColor       =   &H00C0C0C0&
         Caption         =   "my"
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   60
         TabIndex        =   140
         Top             =   1050
         Width           =   270
      End
      Begin VB.Label LabABmx 
         BackColor       =   &H00C0C0C0&
         Caption         =   "mx"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   60
         TabIndex        =   139
         Top             =   720
         Width           =   315
      End
      Begin VB.Label LabABBT 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Red tool spacing +/- 20"
         ForeColor       =   &H000000FF&
         Height          =   450
         Left            =   60
         TabIndex        =   138
         Top             =   210
         Width           =   1110
      End
   End
   Begin VB.Frame frmHairs 
      BackColor       =   &H00C0C0C0&
      Caption         =   "   Hair Lines  "
      ForeColor       =   &H00FF0000&
      Height          =   795
      Left            =   120
      TabIndex        =   124
      Top             =   840
      Width           =   1215
      Begin VB.CheckBox chkPerspec 
         DownPicture     =   "Pencil.frx":1A32
         Height          =   250
         Left            =   645
         Picture         =   "Pencil.frx":1B34
         Style           =   1  'Graphical
         TabIndex        =   168
         ToolTipText     =   "Perspective points"
         Top             =   435
         Width           =   250
      End
      Begin VB.CheckBox chkHairs 
         DownPicture     =   "Pencil.frx":1C36
         Height          =   250
         Index           =   1
         Left            =   180
         Picture         =   "Pencil.frx":1D38
         Style           =   1  'Graphical
         TabIndex        =   126
         ToolTipText     =   "X  hairs"
         Top             =   435
         Width           =   250
      End
      Begin VB.CheckBox chkHairs 
         DownPicture     =   "Pencil.frx":1E3A
         Height          =   250
         Index           =   0
         Left            =   180
         Picture         =   "Pencil.frx":1F3C
         Style           =   1  'Graphical
         TabIndex        =   125
         ToolTipText     =   "+  hairs"
         Top             =   180
         Width           =   250
      End
      Begin VB.Label LabPerspec 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Perspec"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   585
         TabIndex        =   169
         Top             =   210
         Width           =   495
      End
   End
   Begin VB.PictureBox picINFO 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   180
      ScaleHeight     =   195
      ScaleWidth      =   11400
      TabIndex        =   122
      Top             =   7965
      Width           =   11460
   End
   Begin VB.PictureBox picXY 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000013&
      Height          =   495
      Left            =   11265
      ScaleHeight     =   435
      ScaleWidth      =   615
      TabIndex        =   121
      Top             =   6960
      Width           =   675
   End
   Begin VB.CommandButton cmdUndoLast 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Undo Last"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   210
      Style           =   1  'Graphical
      TabIndex        =   119
      Top             =   60
      Width           =   1035
   End
   Begin VB.Frame frmZoom 
      BackColor       =   &H00C0C0C0&
      Caption         =   "     ZOOM    "
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   120
      TabIndex        =   116
      Top             =   3120
      Width           =   1215
      Begin VB.CheckBox chkZOOMAR 
         DownPicture     =   "Pencil.frx":203E
         Height          =   315
         Left            =   270
         Picture         =   "Pencil.frx":24C4
         Style           =   1  'Graphical
         TabIndex        =   123
         ToolTipText     =   "Zoom Active Rect"
         Top             =   630
         Width           =   315
      End
      Begin ComCtl2.UpDown UDZOOM 
         Height          =   660
         Left            =   825
         TabIndex        =   118
         ToolTipText     =   "Zoom Mag"
         Top             =   270
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   1164
         _Version        =   327681
         Value           =   1
         Min             =   1
         Enabled         =   -1  'True
      End
      Begin VB.CheckBox chkZOOM 
         BackColor       =   &H00C0C0C0&
         DownPicture     =   "Pencil.frx":2642
         Height          =   315
         Left            =   135
         Picture         =   "Pencil.frx":28C4
         Style           =   1  'Graphical
         TabIndex        =   117
         ToolTipText     =   "Toggle Zoom"
         Top             =   270
         Width           =   555
      End
   End
   Begin VB.Frame frmRSR 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Resize or Rotate amount"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   120
      TabIndex        =   113
      Top             =   6900
      Width           =   2055
      Begin ComCtl2.UpDown UDRSR 
         Height          =   330
         Left            =   1770
         TabIndex        =   151
         Top             =   270
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   582
         _Version        =   327681
         Value           =   5
         BuddyControl    =   "txtRSR"
         BuddyDispid     =   196633
         OrigLeft        =   1710
         OrigTop         =   210
         OrigRight       =   1905
         OrigBottom      =   765
         Max             =   180
         Min             =   -180
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtRSR 
         Height          =   330
         Left            =   1245
         TabIndex        =   115
         Text            =   "Text1"
         Top             =   270
         Width           =   495
      End
      Begin VB.PictureBox picRSR 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   585
         Left            =   90
         ScaleHeight     =   525
         ScaleWidth      =   1050
         TabIndex        =   114
         Top             =   195
         Width           =   1110
      End
      Begin VB.Label LabN 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "N"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1395
         TabIndex        =   120
         Top             =   615
         Width           =   165
      End
   End
   Begin VB.Frame frmScroller 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Scroll picture"
      ForeColor       =   &H00FF0000&
      Height          =   1035
      Left            =   120
      TabIndex        =   108
      Top             =   5880
      Width           =   1215
      Begin ComCtl2.UpDown UpDown3 
         Height          =   375
         Left            =   780
         TabIndex        =   112
         ToolTipText     =   "Scroll vert"
         Top             =   525
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   661
         _Version        =   327681
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown2 
         Height          =   195
         Left            =   690
         TabIndex        =   111
         ToolTipText     =   "Scroll horz"
         Top             =   270
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   344
         _Version        =   327681
         Orientation     =   1
         Enabled         =   -1  'True
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   390
         Left            =   390
         TabIndex        =   110
         ToolTipText     =   "Incr scroll"
         Top             =   390
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   688
         _Version        =   327681
         Value           =   1
         BuddyControl    =   "txtScroller"
         BuddyDispid     =   196637
         OrigLeft        =   465
         OrigTop         =   285
         OrigRight       =   660
         OrigBottom      =   630
         Max             =   50
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtScroller 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   75
         TabIndex        =   109
         Text            =   "Text1"
         Top             =   435
         Width           =   300
      End
      Begin VB.Label LaPixScroll 
         BackColor       =   &H00C0C0C0&
         Caption         =   "pix"
         Height          =   255
         Left            =   105
         TabIndex        =   145
         Top             =   165
         Width           =   285
      End
   End
   Begin VB.Frame frmActRect 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Active Rect "
      ForeColor       =   &H00FF0000&
      Height          =   1695
      Left            =   120
      TabIndex        =   105
      Top             =   4200
      Width           =   1215
      Begin VB.CommandButton cmdClearActRect 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Clear Rectangle"
         Height          =   555
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   107
         Top             =   1065
         Width           =   915
      End
      Begin VB.PictureBox picActRect 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   60
         ScaleHeight     =   735
         ScaleWidth      =   1035
         TabIndex        =   106
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame frmMCR 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1395
      Left            =   4050
      TabIndex        =   99
      Top             =   2445
      Width           =   435
      Begin VB.OptionButton optMCR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":2B46
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   3
         Left            =   60
         Picture         =   "Pencil.frx":2CC4
         Style           =   1  'Graphical
         TabIndex        =   103
         Top             =   1020
         Width           =   315
      End
      Begin VB.OptionButton optMCR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":2E42
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   2
         Left            =   60
         Picture         =   "Pencil.frx":2FC0
         Style           =   1  'Graphical
         TabIndex        =   102
         Top             =   720
         Width           =   315
      End
      Begin VB.OptionButton optMCR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":313E
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   60
         Picture         =   "Pencil.frx":32BC
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   420
         Width           =   315
      End
      Begin VB.OptionButton optMCR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":343A
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   60
         Picture         =   "Pencil.frx":35B8
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   105
         Width           =   315
      End
   End
   Begin VB.Frame frmFill 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   4995
      Left            =   3615
      TabIndex        =   91
      Top             =   780
      Width           =   435
      Begin VB.OptionButton optFill 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DisabledPicture =   "Pencil.frx":3736
         DownPicture     =   "Pencil.frx":38B4
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   15
         Left            =   60
         Picture         =   "Pencil.frx":3A32
         Style           =   1  'Graphical
         TabIndex        =   160
         Top             =   4620
         Width           =   315
      End
      Begin VB.OptionButton optFill 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DisabledPicture =   "Pencil.frx":3BB0
         DownPicture     =   "Pencil.frx":3D2E
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   14
         Left            =   60
         Picture         =   "Pencil.frx":3EAC
         Style           =   1  'Graphical
         TabIndex        =   159
         Top             =   4320
         Width           =   315
      End
      Begin VB.OptionButton optFill 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DisabledPicture =   "Pencil.frx":402A
         DownPicture     =   "Pencil.frx":41A8
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   13
         Left            =   60
         Picture         =   "Pencil.frx":4326
         Style           =   1  'Graphical
         TabIndex        =   158
         Top             =   4005
         Width           =   315
      End
      Begin VB.OptionButton optFill 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DisabledPicture =   "Pencil.frx":44A4
         DownPicture     =   "Pencil.frx":4622
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   12
         Left            =   60
         Picture         =   "Pencil.frx":47A0
         Style           =   1  'Graphical
         TabIndex        =   157
         Top             =   3720
         Width           =   315
      End
      Begin VB.OptionButton optFill 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":491E
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   11
         Left            =   60
         Picture         =   "Pencil.frx":4A9C
         Style           =   1  'Graphical
         TabIndex        =   156
         Top             =   3420
         Width           =   315
      End
      Begin VB.OptionButton optFill 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":4C1A
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   10
         Left            =   60
         Picture         =   "Pencil.frx":4D98
         Style           =   1  'Graphical
         TabIndex        =   155
         Top             =   3105
         Width           =   315
      End
      Begin VB.OptionButton optFill 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":4F16
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   9
         Left            =   60
         Picture         =   "Pencil.frx":5094
         Style           =   1  'Graphical
         TabIndex        =   154
         Top             =   2820
         Width           =   315
      End
      Begin VB.OptionButton optFill 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":5212
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   8
         Left            =   60
         Picture         =   "Pencil.frx":5390
         Style           =   1  'Graphical
         TabIndex        =   153
         Top             =   2520
         Width           =   315
      End
      Begin VB.OptionButton optFill 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":550E
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   7
         Left            =   60
         Picture         =   "Pencil.frx":568C
         Style           =   1  'Graphical
         TabIndex        =   152
         Top             =   2220
         Width           =   315
      End
      Begin VB.OptionButton optFill 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":580A
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   6
         Left            =   60
         Picture         =   "Pencil.frx":5988
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   1920
         Width           =   315
      End
      Begin VB.OptionButton optFill 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":5B06
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   5
         Left            =   60
         Picture         =   "Pencil.frx":5C84
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   1620
         Width           =   315
      End
      Begin VB.OptionButton optFill 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":5E02
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   4
         Left            =   60
         Picture         =   "Pencil.frx":5F80
         Style           =   1  'Graphical
         TabIndex        =   96
         Top             =   1320
         Width           =   315
      End
      Begin VB.OptionButton optFill 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":60FE
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   3
         Left            =   60
         Picture         =   "Pencil.frx":627C
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   1020
         Width           =   315
      End
      Begin VB.OptionButton optFill 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":63FA
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   2
         Left            =   60
         Picture         =   "Pencil.frx":6578
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   720
         Width           =   315
      End
      Begin VB.OptionButton optFill 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":66F6
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   60
         Picture         =   "Pencil.frx":6874
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   420
         Width           =   315
      End
      Begin VB.OptionButton optFill 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":69F2
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   60
         Picture         =   "Pencil.frx":6B70
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   120
         Width           =   315
      End
   End
   Begin VB.Frame frmRubber 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1440
      Left            =   4050
      TabIndex        =   88
      Top             =   15
      Width           =   435
      Begin VB.OptionButton optRubber 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":6CEE
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   3
         Left            =   60
         Picture         =   "Pencil.frx":6E6C
         Style           =   1  'Graphical
         TabIndex        =   171
         Top             =   1020
         Width           =   315
      End
      Begin VB.OptionButton optRubber 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":6FEA
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   2
         Left            =   60
         Picture         =   "Pencil.frx":7168
         Style           =   1  'Graphical
         TabIndex        =   104
         Top             =   720
         Width           =   315
      End
      Begin VB.OptionButton optRubber 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":72E6
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   60
         Picture         =   "Pencil.frx":7464
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   405
         Width           =   315
      End
      Begin VB.OptionButton optRubber 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":75E2
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   60
         Picture         =   "Pencil.frx":7760
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   120
         Width           =   315
      End
   End
   Begin VB.Frame frmSmudge 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   4050
      TabIndex        =   85
      Top             =   1620
      Width           =   435
      Begin VB.OptionButton optSmudge 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":78DE
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   75
         Picture         =   "Pencil.frx":7A5C
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   420
         Width           =   315
      End
      Begin VB.OptionButton optSmudge 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":7BDA
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   60
         Picture         =   "Pencil.frx":7D58
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   120
         Width           =   315
      End
   End
   Begin VB.Frame frmTPiece 
      BackColor       =   &H00C0C0C0&
      Height          =   795
      Left            =   3600
      TabIndex        =   73
      Top             =   0
      Width           =   435
      Begin VB.OptionButton optTPiece 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":7ED6
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   60
         Picture         =   "Pencil.frx":8054
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   420
         Width           =   315
      End
      Begin VB.OptionButton optTPiece 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":81D2
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   60
         Picture         =   "Pencil.frx":8350
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   120
         Width           =   315
      End
   End
   Begin VB.CheckBox chkTOOLS 
      DownPicture     =   "Pencil.frx":84CE
      Height          =   315
      Index           =   20
      Left            =   1380
      Picture         =   "Pencil.frx":8954
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   6105
      Width           =   315
   End
   Begin VB.CheckBox chkTOOLS 
      DownPicture     =   "Pencil.frx":8DDA
      Height          =   315
      Index           =   19
      Left            =   1380
      Picture         =   "Pencil.frx":9260
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   5790
      Width           =   315
   End
   Begin VB.CheckBox chkTOOLS 
      DownPicture     =   "Pencil.frx":96E6
      Height          =   315
      Index           =   18
      Left            =   1395
      Picture         =   "Pencil.frx":9B6C
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   5505
      Width           =   315
   End
   Begin VB.CheckBox chkTOOLS 
      DownPicture     =   "Pencil.frx":9FF2
      Height          =   315
      Index           =   17
      Left            =   1380
      Picture         =   "Pencil.frx":A170
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   5205
      Width           =   315
   End
   Begin VB.CheckBox chkTOOLS 
      DownPicture     =   "Pencil.frx":A2EE
      Height          =   315
      Index           =   16
      Left            =   1380
      Picture         =   "Pencil.frx":A774
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   4920
      Width           =   315
   End
   Begin VB.CheckBox chkTOOLS 
      DownPicture     =   "Pencil.frx":ABFA
      Height          =   315
      Index           =   15
      Left            =   1365
      Picture         =   "Pencil.frx":AD78
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   4620
      Width           =   315
   End
   Begin VB.CheckBox chkTOOLS 
      DownPicture     =   "Pencil.frx":B35E
      Height          =   315
      Index           =   14
      Left            =   1380
      Picture         =   "Pencil.frx":B4DC
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4305
      Width           =   315
   End
   Begin VB.CheckBox chkTOOLS 
      DownPicture     =   "Pencil.frx":B65A
      Height          =   315
      Index           =   13
      Left            =   1380
      Picture         =   "Pencil.frx":B7D8
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4020
      Width           =   315
   End
   Begin VB.CheckBox chkTOOLS 
      DownPicture     =   "Pencil.frx":B956
      Height          =   315
      Index           =   12
      Left            =   1380
      Picture         =   "Pencil.frx":BAD4
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3720
      Width           =   315
   End
   Begin VB.CheckBox chkTOOLS 
      DownPicture     =   "Pencil.frx":BC52
      Height          =   315
      Index           =   11
      Left            =   1380
      Picture         =   "Pencil.frx":C0D8
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3420
      Width           =   315
   End
   Begin VB.CheckBox chkTOOLS 
      DownPicture     =   "Pencil.frx":C55E
      Height          =   315
      Index           =   10
      Left            =   1380
      Picture         =   "Pencil.frx":C6DC
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3120
      Width           =   315
   End
   Begin VB.CheckBox chkTOOLS 
      DownPicture     =   "Pencil.frx":C85A
      Height          =   315
      Index           =   9
      Left            =   1380
      Picture         =   "Pencil.frx":C9D8
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2820
      Width           =   315
   End
   Begin VB.CheckBox chkTOOLS 
      DownPicture     =   "Pencil.frx":CB56
      Height          =   315
      Index           =   8
      Left            =   1395
      Picture         =   "Pencil.frx":CCD4
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2520
      Width           =   315
   End
   Begin VB.CheckBox chkTOOLS 
      DownPicture     =   "Pencil.frx":CE52
      Height          =   315
      Index           =   7
      Left            =   1380
      Picture         =   "Pencil.frx":CFD0
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2220
      Width           =   315
   End
   Begin VB.CheckBox chkTOOLS 
      DownPicture     =   "Pencil.frx":D14E
      Height          =   315
      Index           =   6
      Left            =   1380
      Picture         =   "Pencil.frx":D2CC
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1920
      Width           =   315
   End
   Begin VB.CheckBox chkTOOLS 
      DownPicture     =   "Pencil.frx":D44A
      Height          =   315
      Index           =   5
      Left            =   1380
      Picture         =   "Pencil.frx":D5C8
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1620
      Width           =   315
   End
   Begin VB.CheckBox chkTOOLS 
      DownPicture     =   "Pencil.frx":D746
      Height          =   315
      Index           =   4
      Left            =   1380
      Picture         =   "Pencil.frx":D8C4
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1320
      Width           =   315
   End
   Begin VB.CheckBox chkTOOLS 
      DownPicture     =   "Pencil.frx":DA42
      Height          =   315
      Index           =   3
      Left            =   1380
      Picture         =   "Pencil.frx":DBC0
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1035
      Width           =   315
   End
   Begin VB.CheckBox chkTOOLS 
      DownPicture     =   "Pencil.frx":DD3E
      Height          =   315
      Index           =   2
      Left            =   1380
      Picture         =   "Pencil.frx":DEBC
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   720
      Width           =   315
   End
   Begin VB.CheckBox chkTOOLS 
      DownPicture     =   "Pencil.frx":E03A
      Height          =   315
      Index           =   1
      Left            =   1380
      Picture         =   "Pencil.frx":E1B8
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   405
      Width           =   315
   End
   Begin VB.CheckBox chkTOOLS 
      DownPicture     =   "Pencil.frx":E336
      Height          =   315
      Index           =   0
      Left            =   1380
      Picture         =   "Pencil.frx":E4B4
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   120
      Width           =   315
   End
   Begin VB.PictureBox picPAL 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   11370
      ScaleHeight     =   233
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   1
      Top             =   435
      Width           =   495
   End
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   8895
      Left            =   5190
      ScaleHeight     =   589
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   405
      TabIndex        =   0
      Top             =   -1035
      Width           =   6135
      Begin VB.CommandButton cmdHome 
         Caption         =   "Close"
         Height          =   300
         Index           =   1
         Left            =   2970
         TabIndex        =   174
         Top             =   4950
         Width           =   600
      End
      Begin VB.CommandButton cmdHome 
         Caption         =   "Home"
         Height          =   300
         Index           =   0
         Left            =   1470
         TabIndex        =   173
         Top             =   4935
         Width           =   600
      End
      Begin VB.ListBox List1 
         Height          =   1035
         ItemData        =   "Pencil.frx":E632
         Left            =   1080
         List            =   "Pencil.frx":E634
         TabIndex        =   172
         Top             =   4800
         Width           =   3000
      End
      Begin VB.Frame frmText 
         Caption         =   "   TEXT   "
         Height          =   4185
         Left            =   2760
         TabIndex        =   127
         Top             =   1605
         Width           =   3315
         Begin ComCtl2.UpDown UpDown4 
            Height          =   405
            Left            =   2400
            TabIndex        =   162
            Top             =   2235
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   714
            _Version        =   327681
            BuddyControl    =   "txtAngle"
            BuddyDispid     =   196658
            OrigLeft        =   1665
            OrigTop         =   2145
            OrigRight       =   1860
            OrigBottom      =   2565
            Increment       =   5
            Max             =   180
            Min             =   -180
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtAngle 
            Height          =   285
            Left            =   1800
            TabIndex        =   161
            Text            =   "Text1"
            Top             =   2265
            Width           =   450
         End
         Begin VB.CommandButton cmdTextCancel 
            Caption         =   "Cancel"
            Height          =   315
            Left            =   2040
            TabIndex        =   130
            Top             =   300
            Width           =   675
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "Font"
            Height          =   315
            Left            =   480
            TabIndex        =   129
            Top             =   300
            Width           =   735
         End
         Begin VB.TextBox Text2 
            Height          =   1275
            Left            =   240
            TabIndex        =   128
            Text            =   "Text2"
            Top             =   765
            Width           =   2835
         End
         Begin VB.Image Image2 
            Height          =   315
            Left            =   810
            Picture         =   "Pencil.frx":E636
            Top             =   2520
            Width           =   255
         End
         Begin VB.Label Label2 
            Caption         =   "Text angle (deg)"
            Height          =   315
            Left            =   405
            TabIndex        =   163
            Top             =   2280
            Width           =   1290
         End
         Begin VB.Label labText 
            Caption         =   $"Pencil.frx":EABC
            Height          =   1155
            Left            =   285
            TabIndex        =   131
            Top             =   2835
            Width           =   2790
         End
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   8310
         Top             =   7320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.PictureBox picZOOMBox 
         Height          =   840
         Left            =   1695
         ScaleHeight     =   780
         ScaleWidth      =   570
         TabIndex        =   4
         Top             =   3150
         Width           =   630
      End
      Begin VB.PictureBox picMCRBox 
         Height          =   825
         Left            =   1695
         ScaleHeight     =   765
         ScaleWidth      =   570
         TabIndex        =   3
         Top             =   2250
         Width           =   630
      End
      Begin VB.PictureBox picCanvasStore 
         AutoRedraw      =   -1  'True
         Height          =   825
         Left            =   1680
         ScaleHeight     =   51
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   40
         TabIndex        =   2
         Top             =   1305
         Width           =   660
      End
      Begin VB.Shape shpARect 
         BorderColor     =   &H00FF0000&
         BorderStyle     =   3  'Dot
         Height          =   315
         Left            =   525
         Top             =   4920
         Width           =   375
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FF8080&
         BorderStyle     =   2  'Dash
         X1              =   47
         X2              =   86
         Y1              =   280
         Y2              =   265
      End
      Begin VB.Shape Shape1 
         Height          =   90
         Index           =   2
         Left            =   960
         Shape           =   3  'Circle
         Top             =   4470
         Width           =   90
      End
      Begin VB.Shape Shape1 
         Height          =   90
         Index           =   1
         Left            =   705
         Shape           =   3  'Circle
         Top             =   4485
         Width           =   90
      End
      Begin VB.Shape Shape1 
         Height          =   90
         Index           =   0
         Left            =   420
         Shape           =   3  'Circle
         Top             =   4485
         Width           =   90
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FF8080&
         BorderStyle     =   4  'Dash-Dot
         X1              =   29
         X2              =   77
         Y1              =   275
         Y2              =   248
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FF8080&
         BorderStyle     =   3  'Dot
         X1              =   31
         X2              =   79
         Y1              =   254
         Y2              =   268
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00C0C0FF&
         BorderStyle     =   3  'Dot
         X1              =   36
         X2              =   76
         Y1              =   235
         Y2              =   235
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FF0000&
         BorderStyle     =   3  'Dot
         X1              =   74
         X2              =   35
         Y1              =   154
         Y2              =   193
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF0000&
         BorderStyle     =   3  'Dot
         X1              =   31
         X2              =   80
         Y1              =   157
         Y2              =   190
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FF0000&
         BorderStyle     =   3  'Dot
         X1              =   58
         X2              =   58
         Y1              =   85
         Y2              =   140
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderStyle     =   3  'Dot
         X1              =   31
         X2              =   89
         Y1              =   115
         Y2              =   115
      End
      Begin VB.Label LabWait 
         Caption         =   "WAIT"
         Height          =   240
         Left            =   1845
         TabIndex        =   13
         Top             =   4245
         Width           =   450
      End
   End
   Begin VB.Frame frmLine 
      BackColor       =   &H00C0C0C0&
      Height          =   4695
      Left            =   2655
      TabIndex        =   60
      Top             =   0
      Width           =   435
      Begin VB.OptionButton optLine 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":EBB5
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   12
         Left            =   60
         Picture         =   "Pencil.frx":ED33
         Style           =   1  'Graphical
         TabIndex        =   180
         Top             =   3720
         Width           =   315
      End
      Begin VB.OptionButton optLine 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":EEB1
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   11
         Left            =   60
         Picture         =   "Pencil.frx":F02F
         Style           =   1  'Graphical
         TabIndex        =   179
         Top             =   3420
         Width           =   315
      End
      Begin VB.OptionButton optLine 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":F1AD
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   10
         Left            =   60
         Picture         =   "Pencil.frx":F32B
         Style           =   1  'Graphical
         TabIndex        =   178
         Top             =   3120
         Width           =   315
      End
      Begin VB.OptionButton optLine 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":F4A9
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   9
         Left            =   60
         Picture         =   "Pencil.frx":F627
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   2820
         Width           =   315
      End
      Begin VB.OptionButton optLine 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":F7A5
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   8
         Left            =   60
         Picture         =   "Pencil.frx":F923
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   2520
         Width           =   315
      End
      Begin VB.OptionButton optLine 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":FAA1
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   7
         Left            =   60
         Picture         =   "Pencil.frx":FC1F
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   2220
         Width           =   315
      End
      Begin VB.OptionButton optLine 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":FD9D
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   6
         Left            =   60
         Picture         =   "Pencil.frx":FF1B
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   1920
         Width           =   315
      End
      Begin VB.OptionButton optLine 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":10099
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   5
         Left            =   60
         Picture         =   "Pencil.frx":10217
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   1620
         Width           =   315
      End
      Begin VB.OptionButton optLine 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":10395
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   4
         Left            =   60
         Picture         =   "Pencil.frx":10513
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   1320
         Width           =   315
      End
      Begin VB.OptionButton optLine 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":10691
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   3
         Left            =   60
         Picture         =   "Pencil.frx":1080F
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   1020
         Width           =   315
      End
      Begin VB.OptionButton optLine 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":1098D
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   2
         Left            =   60
         Picture         =   "Pencil.frx":10B0B
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   720
         Width           =   315
      End
      Begin VB.OptionButton optLine 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":10C89
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   60
         Picture         =   "Pencil.frx":10E07
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   420
         Width           =   315
      End
      Begin VB.OptionButton optLine 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":10F85
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   60
         Picture         =   "Pencil.frx":11103
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   120
         Width           =   315
      End
   End
   Begin VB.Frame frmAirBrush 
      BackColor       =   &H00C0C0C0&
      Height          =   4695
      Left            =   2190
      TabIndex        =   51
      Top             =   0
      Width           =   435
      Begin VB.OptionButton optAirBrush 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":11281
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   10
         Left            =   60
         Picture         =   "Pencil.frx":113FF
         Style           =   1  'Graphical
         TabIndex        =   144
         Top             =   3120
         Width           =   315
      End
      Begin VB.OptionButton optAirBrush 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":1157D
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   9
         Left            =   60
         Picture         =   "Pencil.frx":116FB
         Style           =   1  'Graphical
         TabIndex        =   143
         Top             =   2820
         Width           =   315
      End
      Begin VB.OptionButton optAirBrush 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":11879
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   8
         Left            =   60
         Picture         =   "Pencil.frx":119F7
         Style           =   1  'Graphical
         TabIndex        =   132
         Top             =   2520
         Width           =   315
      End
      Begin VB.OptionButton optAirBrush 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":11B75
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   7
         Left            =   60
         Picture         =   "Pencil.frx":11CF3
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   2220
         Width           =   315
      End
      Begin VB.OptionButton optAirBrush 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":11E71
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   6
         Left            =   60
         Picture         =   "Pencil.frx":11FEF
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   1920
         Width           =   315
      End
      Begin VB.OptionButton optAirBrush 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":1216D
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   5
         Left            =   60
         Picture         =   "Pencil.frx":122EB
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   1620
         Width           =   315
      End
      Begin VB.OptionButton optAirBrush 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":12469
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   4
         Left            =   60
         Picture         =   "Pencil.frx":125E7
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   1320
         Width           =   315
      End
      Begin VB.OptionButton optAirBrush 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":12765
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   3
         Left            =   60
         Picture         =   "Pencil.frx":128E3
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   1020
         Width           =   315
      End
      Begin VB.OptionButton optAirBrush 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":12A61
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   2
         Left            =   60
         Picture         =   "Pencil.frx":12BDF
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   720
         Width           =   315
      End
      Begin VB.OptionButton optAirBrush 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":12D5D
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   60
         Picture         =   "Pencil.frx":12EDB
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   420
         Width           =   315
      End
      Begin VB.OptionButton optAirBrush 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":13059
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   60
         Picture         =   "Pencil.frx":131D7
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   120
         Width           =   315
      End
   End
   Begin VB.Frame frmBrush 
      BackColor       =   &H00C0C0C0&
      Height          =   4695
      Left            =   1770
      TabIndex        =   35
      Top             =   0
      Width           =   435
      Begin VB.OptionButton optBrush 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":13355
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   14
         Left            =   60
         Picture         =   "Pencil.frx":134D3
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   4320
         Width           =   315
      End
      Begin VB.OptionButton optBrush 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":13651
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   13
         Left            =   60
         Picture         =   "Pencil.frx":137CF
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   4020
         Width           =   315
      End
      Begin VB.OptionButton optBrush 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":1394D
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   12
         Left            =   60
         Picture         =   "Pencil.frx":13ACB
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   3720
         Width           =   315
      End
      Begin VB.OptionButton optBrush 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":13C49
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   11
         Left            =   60
         Picture         =   "Pencil.frx":13DC7
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   3405
         Width           =   315
      End
      Begin VB.OptionButton optBrush 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":13F45
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   10
         Left            =   60
         Picture         =   "Pencil.frx":140C3
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   3120
         Width           =   315
      End
      Begin VB.OptionButton optBrush 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":14241
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   9
         Left            =   60
         Picture         =   "Pencil.frx":143BF
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   2820
         Width           =   315
      End
      Begin VB.OptionButton optBrush 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":1453D
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   8
         Left            =   60
         Picture         =   "Pencil.frx":146BB
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   2520
         Width           =   315
      End
      Begin VB.OptionButton optBrush 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":14839
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   7
         Left            =   60
         Picture         =   "Pencil.frx":149B7
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   2220
         Width           =   315
      End
      Begin VB.OptionButton optBrush 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":14B35
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   6
         Left            =   60
         Picture         =   "Pencil.frx":14CB3
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   1905
         Width           =   315
      End
      Begin VB.OptionButton optBrush 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":14E31
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   5
         Left            =   60
         Picture         =   "Pencil.frx":14FAF
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   1620
         Width           =   315
      End
      Begin VB.OptionButton optBrush 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":1512D
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   4
         Left            =   60
         Picture         =   "Pencil.frx":152AB
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   1320
         Width           =   315
      End
      Begin VB.OptionButton optBrush 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":15429
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   3
         Left            =   60
         Picture         =   "Pencil.frx":155A7
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   1020
         Width           =   315
      End
      Begin VB.OptionButton optBrush 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":15725
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   2
         Left            =   60
         Picture         =   "Pencil.frx":158A3
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   720
         Width           =   315
      End
      Begin VB.OptionButton optBrush 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":15A21
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   60
         Picture         =   "Pencil.frx":15B9F
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   420
         Width           =   315
      End
      Begin VB.OptionButton optBrush 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":15D1D
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   60
         Picture         =   "Pencil.frx":15E9B
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   120
         Width           =   315
      End
   End
   Begin VB.Frame frmRect 
      BackColor       =   &H00C0C0C0&
      Height          =   4695
      Left            =   3120
      TabIndex        =   71
      Top             =   0
      Width           =   435
      Begin VB.OptionButton optRect 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":16019
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   9
         Left            =   60
         Picture         =   "Pencil.frx":16197
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   2820
         Width           =   315
      End
      Begin VB.OptionButton optRect 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":16315
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   8
         Left            =   60
         Picture         =   "Pencil.frx":16493
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   2520
         Width           =   315
      End
      Begin VB.OptionButton optRect 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":16611
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   7
         Left            =   60
         Picture         =   "Pencil.frx":1678F
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   2220
         Width           =   315
      End
      Begin VB.OptionButton optRect 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":1690D
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   6
         Left            =   60
         Picture         =   "Pencil.frx":16A8B
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   1920
         Width           =   315
      End
      Begin VB.OptionButton optRect 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":16C09
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   5
         Left            =   60
         Picture         =   "Pencil.frx":16D87
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   1620
         Width           =   315
      End
      Begin VB.OptionButton optRect 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":16F05
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   4
         Left            =   60
         Picture         =   "Pencil.frx":17083
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   1320
         Width           =   315
      End
      Begin VB.OptionButton optRect 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":17201
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   3
         Left            =   60
         Picture         =   "Pencil.frx":1737F
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   1020
         Width           =   315
      End
      Begin VB.OptionButton optRect 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":174FD
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   2
         Left            =   60
         Picture         =   "Pencil.frx":1767B
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   720
         Width           =   315
      End
      Begin VB.OptionButton optRect 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":177F9
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   60
         Picture         =   "Pencil.frx":17977
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   420
         Width           =   315
      End
      Begin VB.OptionButton optRect 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         DownPicture     =   "Pencil.frx":17AF5
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   60
         Picture         =   "Pencil.frx":17C73
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   120
         Width           =   315
      End
   End
   Begin VB.Label LabColors 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Colors"
      Height          =   255
      Left            =   11340
      TabIndex        =   170
      Top             =   0
      Width           =   555
   End
   Begin VB.Label LabCanvasHt 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Canvas   height"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1725
      TabIndex        =   165
      Top             =   5070
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   315
      Left            =   1410
      Picture         =   "Pencil.frx":17DF1
      Top             =   6525
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Coords"
      Height          =   210
      Left            =   11370
      TabIndex        =   150
      Top             =   6720
      Width           =   525
   End
   Begin VB.Label LabDraw 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Draw"
      Height          =   240
      Left            =   11430
      TabIndex        =   149
      Top             =   4140
      Width           =   465
   End
   Begin VB.Label LabMsg 
      Caption         =   "LabMsg"
      Height          =   195
      Left            =   11385
      TabIndex        =   141
      Top             =   7935
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label LabRIGHTCN 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LabRIGHTCN"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11490
      TabIndex        =   12
      Top             =   6060
      Width           =   375
   End
   Begin VB.Label LabLEFTCN 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LabLEFTCN"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11475
      TabIndex        =   11
      Top             =   5205
      Width           =   375
   End
   Begin VB.Label LabRIGHTCUL 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LabRIGHTCUL"
      Height          =   315
      Left            =   11490
      TabIndex        =   10
      ToolTipText     =   "Rubber color"
      Top             =   6300
      Width           =   495
   End
   Begin VB.Label LabRIGHT 
      Alignment       =   2  'Center
      BackColor       =   &H00C8C8C8&
      Caption         =   "Right"
      Height          =   195
      Left            =   11430
      TabIndex        =   9
      Top             =   5820
      Width           =   435
   End
   Begin VB.Label LabLEFTCUL 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LabLEFTCUL"
      Height          =   315
      Left            =   11460
      TabIndex        =   8
      ToolTipText     =   "Draw color"
      Top             =   5445
      Width           =   495
   End
   Begin VB.Label LabLeft 
      BackColor       =   &H00C8C8C8&
      Caption         =   "Left"
      Height          =   195
      Left            =   11475
      TabIndex        =   7
      Top             =   4995
      Width           =   375
   End
   Begin VB.Label LabCUL 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LabCUL"
      Height          =   315
      Left            =   11415
      TabIndex        =   6
      Top             =   4590
      Width           =   495
   End
   Begin VB.Label LabCN 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LabCN"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   11460
      TabIndex        =   5
      Top             =   4365
      Width           =   375
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu LoadFile 
         Caption         =   "&Load picture  File"
      End
      Begin VB.Menu AddFile 
         Caption         =   "&Add picture file"
      End
      Begin VB.Menu brk1 
         Caption         =   "-"
      End
      Begin VB.Menu SaveBMP 
         Caption         =   "Save As &BMP"
      End
      Begin VB.Menu SaveJPG 
         Caption         =   "Save As &JPG"
      End
      Begin VB.Menu SaveRectangle8bitBMP 
         Caption         =   "Save &Rectangle as grey BMP"
      End
      Begin VB.Menu brk2 
         Caption         =   "-"
      End
      Begin VB.Menu PrintBMP 
         Caption         =   "&Print"
      End
      Begin VB.Menu brk3 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu PencilHelp 
      Caption         =   "&Help"
   End
   Begin VB.Menu InstrucCaptions 
      Caption         =   "Instruc"
   End
   Begin VB.Menu ClearPicture 
      Caption         =   " [ &CLEAR PICTURE ]"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'RRPencil  by Robert Rayment

'Corrected  13/9/00, 14/9/00 getting there!

Option Base 1
DefInt A-Q  'a-q integers
DefSng R-Z  'rst, uvw, xyz real
Private Sub MouseTArr()
'Change MouseCursor
On Error GoTo curerror
tcur$ = App.Path & "\t.cur"
Open tcur$ For Input As #1
Close #1
   MouseIcon = LoadPicture(tcur$)
   MousePointer = 99
Exit Sub
'========
curerror:
Close
MousePointer = vbDefault
End Sub


Private Sub Form_Load()
MouseTArr

picINFO.Visible = False 'Make True for debug info

CurrentDirec$ = App.Path
If Right$(CurrentDirec$, 1) <> "\" Then CurrentDirec$ = CurrentDirec$ + "\"

With Form1
   .ScaleMode = vbPixels
   .BackColor = &HC0C0C0   'RGB(150, 150, 150)
   .Caption = "RRPencil by Robert Rayment"
End With

'Set up RRPencil Help
SETUPHELP
'----------
SETUPSCREEN
'----------
SetInstructions
'----------
SetZOOMParams
'----------
SetBitArray
'----------

TOOL = 0
InstrucCaptions.Caption = Instruction$(TOOL)
'------------------
ClearAllSubToolbars
frmBrush.Visible = True
'------------------
arleft = 0: artop = 0: arright = 0: arbottom = 0 'Active rect coords
ShowActRectCoords
'------------------

ReDim xsto(1000), ysto(1000)  'For Double & Poly Lines & Splines
PCount = 0  'Point count where needed
zpspac = 0  'Parallel spacing
LCount = 0  'LC COUNT

'CLEAR SWITCHES
DrawingMode = False
chkZOOM.Value = vbUnchecked
ZoomMode = False
TextSW = False
MCRSW = False
ResizeSW = False
PerspectiveSW = False
RotateSW = False
AddBMPSW = False
ActiveRectExists = False
zoomdrawing = False

LabWait.Visible = False
picMCRBox.DragMode = vbManual

picCanvas.PSet (0, 0), 0

'Clear hair & perspective shapes
ShowPlusLines = False
Line1.Visible = False
Line2.Visible = False
ShowXLines = False
Line3.Visible = False
Line4.Visible = False

ShowPerspecLines = False
Line5.Visible = False
Line6.Visible = False
Line7.Visible = False
Line8.Visible = False
Shape1(0).Visible = False
Shape1(1).Visible = False
Shape1(2).Visible = False
SetPerspecPts = False
NumPPts = 0
shpARect.Visible = False

'Set AirBrush mx,my
txtmx.Text = 12
txtmy.Text = 8

'Set start picture scroll amount
txtScroller.Text = "1"

'Set start text angle
frmText.Visible = False
txtAngle.Text = "0"

'720 Canvas scroller
VScroll1.Visible = False

RedrawPolCoTu = 0

DrawCanvasBorder

DoEvents
End Sub

Public Sub DrawCanvasBorder()
'In: picCanvas.Heights 520 or 720
Form1.DrawWidth = 1
Q& = Form1.BackColor
PL = picCanvas.Left - 1: PT = picCanvas.Top - 1
PW = picCanvas.Width + 2: PH = picCanvas.Height + 2
'Line (PL, PT)-Step(PW, 0), 0

Line (PL, PT)-Step(0, PT + 720), Q&
Line (PL + PW, PT)-Step(0, PT + 720), Q&

Line (PL, PT)-Step(0, PH), 0
Line (PL + PW, PT)-Step(0, PT + PH), 0
'Line (PL + PW, PT + PH)-Step(PL + PH, 0), QBColor(8)
Form1.DrawWidth = 1
Refresh
End Sub


Public Sub ShowActRectCoords()
'Global arleft, artop, arright, arbottom, arwidth, arheight 'Active rect coords
'arleft must be < arright.  artop must be < arbottom

FixRectCoords arleft, arright, artop, arbottom

picActRect.Cls
picActRect.Print "Top"; Tab(7); " ="; artop
picActRect.Print "Left"; Tab(7); " ="; arleft
arwidth = arright - arleft - 1
If arwidth < 0 Then arwidth = 0
picActRect.Print "Width"; Tab(7); " ="; arwidth
arheight = arbottom - artop - 1
If arheight < 0 Then arheight = 0
picActRect.Print "Height"; Tab(7); " ="; arheight
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Not DrawingMode Then
'   cul& = Form1.Point(X, Y)
'   Label1.Caption = Hex$(cul&)
'   GetGreyCN cul&, culnum
'   LabCN.Caption = Str$(255 - culnum)
'   LabCUL.BackColor = cul&
'End If
End Sub

Private Sub chkTOOLS_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index <> TOOL And Index <> prevIndex Then
   InstrucCaptions.Caption = Instruction$(Index)
End If
prevIndex = Index
End Sub

'SELECT TOOL
Private Sub chkTOOLS_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If ActionInProgress And Index <> TOOL Then
   chkTOOLS(Index).Value = vbUnchecked
   Exit Sub
End If

If Index = TOOL Then
   chkTOOLS(Index).Value = vbChecked
   InstrucCaptions.Caption = Instruction$(TOOL)
   Exit Sub
End If

prevTOOL = TOOL
chkTOOLS(TOOL).Value = vbUnchecked '0
TOOL = Index
chkTOOLS(TOOL).Value = vbChecked '1
picCanvas.SetFocus

ShowInstructionsAndSubToolBars

End Sub

Private Sub ShowInstructionsAndSubToolBars()
InstrucCaptions.Caption = Instruction$(TOOL)

ClearAllSubToolbars

'Show any sub-toolbars
Select Case TOOL
Case 0
   frmBrush.Visible = True
Case 1
   frmAirBrush.Visible = True
Case 2 To 4
   frmLine.Visible = True
   If TOOL = 3 Then  'PolyLine
      optLine(10).Visible = True 'Parallelogram
      optLine(11).Visible = True 'Frustrum
      optLine(12).Visible = True 'Auto-join
   Else
      optLine(10).Visible = False
      optLine(11).Visible = False
      optLine(12).Visible = False
      optLine(0).Value = True
      LineType = 0
   End If
   
Case 5 To 9
   frmRect.Visible = True
Case 10
   frmTPiece.Visible = True
Case 11
   TextSW = True
   ActionInProgress
   frmText.Visible = True

Case 12
   frmFill.Visible = True
Case 13
   frmRubber.Visible = True
Case 15
   frmSmudge.Visible = True
Case 16
   frmMCR.Visible = True
Case 17
   frmResize.Visible = True
Case 19
   frmRotate.Visible = True
End Select
End Sub

'SELECT SUB-TOOLS

Private Sub optBrush_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If ActionInProgress Then optBrush(BrushType).Value = True: Exit Sub
BrushType = Index
picCanvas.SetFocus
End Sub
Private Sub optAirBrush_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If ActionInProgress Then optAirBrush(AirBrushType).Value = True: Exit Sub
AirBrushType = Index
picCanvas.SetFocus
End Sub
Private Sub optLine_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If ActionInProgress Then optLine(LineType).Value = True: Exit Sub
LineType = Index
picCanvas.SetFocus
End Sub
Private Sub optRect_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If ActionInProgress Then optRect(RectangleType).Value = True: Exit Sub
RectangleType = Index
picCanvas.SetFocus
End Sub
Private Sub optTPiece_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If ActionInProgress Then optTPiece(TPieceType).Value = True: Exit Sub
TPieceType = Index
picCanvas.SetFocus
End Sub
Private Sub optFill_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If ActionInProgress Then optFill(FillType).Value = True: Exit Sub
FillType = Index
picCanvas.SetFocus
End Sub
Private Sub optRubber_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If ActionInProgress Then optLine(RubberType).Value = True: Exit Sub
RubberType = Index
picCanvas.SetFocus
End Sub
Private Sub optSmudge_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If ActionInProgress Then optRubber(SmudgeType).Value = True: Exit Sub
SmudgeType = Index
picCanvas.SetFocus
End Sub
Private Sub optMCR_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If ActionInProgress Then optMCR(MCRType).Value = True: Exit Sub
MCRType = Index
picCanvas.SetFocus
End Sub
Private Sub optResize_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If ActionInProgress Then optResize(ResizeType).Value = True: Exit Sub
ResizeType = Index
picCanvas.SetFocus
End Sub
Private Sub optRotate_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If ActionInProgress Then optRotate(RotateType).Value = True: Exit Sub
RotateType = Index
picCanvas.SetFocus
End Sub

Private Sub cmdClearPic_Click()
If ActionInProgress Then Exit Sub
resp = RRMsgBox$("RRPencil", "CLEAR PICTURE ?", 1, 1, 0, 1) 'YES,NO,OK, NumLines
If resp = vbYes Then
   picCanvas.DrawMode = 13
   picCanvas.Line (0, 0)-Step(picCanvas.Width, picCanvas.Height), QBColor(15), BF
   arleft = 0: artop = 0: arright = 0: arbottom = 0 'Active rect coords
   ShowActRectCoords
   ActiveRectExists = False
   '------------------
   LabWait.Visible = False
End If
picCanvas.SetFocus
End Sub

Private Sub cmdUndoLast_Click()
'UNDO LAST SHAPE ?
If ActionInProgress Then Exit Sub

CANVASRESTORE

End Sub
Private Sub SetVanishingPoints(Button, X, Y)
'From picCanvas_MouseDown
'Global ShowPerspecLines, SetPerspecPts,
'       YPr, XPr1, XPr2, XPr3, NumPPts
   NumPPts = NumPPts + 1
   If NumPPts = 1 Then
      YPr = Y: XPr1 = X
      Shape1(0).Left = XPr1 - Shape1(0).Width / 2
      Shape1(0).Top = YPr - Shape1(0).Height / 2
      Shape1(0).Visible = True
      Line5.Visible = True 'Horz perline
      Line6.Visible = True 'perline 1
   ElseIf NumPPts = 2 Then
      If Button = 2 Then 'RC
         NumPPts = 1
         SetPerspecPts = False
         ShowPerspecLines = True
         InstrucCaptions.Caption = Instruction$(TOOL)
         Exit Sub
      End If
      XPr2 = X
      Shape1(1).Left = XPr2 - Shape1(1).Width / 2
      Shape1(1).Top = YPr - Shape1(1).Height / 2
      Shape1(1).Visible = True
      Line7.Visible = True 'perline 2
   ElseIf NumPPts = 3 Then
      If Button = 2 Then 'RC
         NumPPts = 2
         SetPerspecPts = False
         ShowPerspecLines = True
         InstrucCaptions.Caption = Instruction$(TOOL)
         Exit Sub
      End If
      XPr3 = X
      Shape1(2).Left = XPr3 - Shape1(2).Width / 2
      Shape1(2).Top = YPr - Shape1(2).Height / 2
      Shape1(2).Visible = True
      Line8.Visible = True 'perline 3
      SetPerspecPts = False
      ShowPerspecLines = True
      InstrucCaptions.Caption = Instruction$(TOOL)
   End If
End Sub

Private Sub chkPerspec_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If DrawingMode = True Then chkPerspec.Value = vbUnchecked: Exit Sub

If chkPerspec.Value = vbChecked Then
   If ActionInProgress Then chkPerspec.Value = vbUnchecked: Exit Sub

   InstrucCaptions.Caption = Space$(8) & "SET HORIZON && PERSPECTIVE POINTS:  LC on Picture for each perspective point (max 3) - RC to Stop"
   SetPerspecPts = True
   
   If ActiveRectExists Then   'To stop Zoom box showing it
      cmdClearActRect_Click
   End If

Else
   InstrucCaptions.Caption = Instruction$(TOOL)
   SetPerspecPts = False
   ShowPerspecLines = False
   Line5.Visible = False
   Line6.Visible = False
   Line7.Visible = False
   Line8.Visible = False
   Shape1(0).Visible = False
   Shape1(1).Visible = False
   Shape1(2).Visible = False
   NumPPts = 0
End If

End Sub
Private Sub StartZoom(X, Y)
'From picCanvas_MouseDown
   DrawingMode = False
   If ActiveRectExists Then   'To stop Zoom box showing it
      cmdClearActRect_Click
   End If
   
   CANVASSAVE
   
   picZOOMBox.Cls
   picZOOMBox.Visible = False
   picCanvas.DrawMode = 13
   xzoom = X: yzoom = Y
   'Shift Mouspoint to bottom-right of zoom box
   'so that picture at x=0 y=0 positions correctly
   'during picCanvas to picZoomBox API transfer
   If X < ZoomSize(ZoomSpan) Then xzoom = ZoomSize(ZoomSpan)
   If Y < ZoomSize(ZoomSpan) * 2 Then yzoom = ZoomSize(ZoomSpan) * 2
   'Position Zoom grid. NB Global GridStep starts with default
   PositionZoomBox xzoom
   picZOOMBox.Visible = True
   picZOOMBox.Refresh
   'Transfer picCanvas points to  Zoom grid
   MousePointer = vbHourglass
   DoEvents
   TransferCanvasToZoomBox 'Global xzoom, yzoom
   MouseTArr
   'MousePointer = vbDefault
   picZOOMBox.SetFocus
End Sub

Private Sub chkZOOM_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If DrawingMode = True Then chkZOOM.Value = vbUnchecked: Exit Sub

If chkZOOM.Value = vbChecked Then
   If ActionInProgress Then chkZOOM.Value = vbUnchecked: Exit Sub

   InstrucCaptions.Caption = Space$(8) & "ZOOM:  LC ON PICTURE, ADJUST MAG WITH UPDOWN BUTTONS, THEN LC IN ZOOM BOX"
   ZoomMode = True
   
   If ActiveRectExists Then   'To stop Zoom box showing it
      cmdClearActRect_Click
   End If

Else
   InstrucCaptions.Caption = Instruction$(TOOL)
   ZoomMode = False

   picZOOMBox.Cls
   picZOOMBox.Visible = False

End If
End Sub
Private Sub TransferCanvasToZoomBox()
'Global xzoom & yzoom are the MousePos on picCanvas
MousePointer = vbHourglass
picZOOMBox.Cls
DestDC& = picZOOMBox.hdc
DestX = 0
DestY = 0
DestWidth = 240
DestHeight = 480
SorcDC& = picCanvas.hdc
SorcX = xzoom - ZoomSize(ZoomSpan)
SorcY = yzoom - 2 * ZoomSize(ZoomSpan)
SorcWidth = 2 * ZoomSize(ZoomSpan) '+ 1
SorcHeight = 4 * ZoomSize(ZoomSpan) '+ 1
dwRop& = &HCC0020  'SRCCOPY Src to Dest

'If StretchBlt SorcX/Y+SorcWidth/Height goes beyond
'picCanvas then nothing is tranferred, therefore
'need to test this and adjust Sorc & Dest Width/Height
overlap = SorcX + SorcWidth - (picCanvas.Width)
If overlap > 0 Then
   xzoom = xzoom - overlap
   SorcX = xzoom - ZoomSize(ZoomSpan)
End If
overlap = SorcY + SorcHeight - (picCanvas.Height)
If overlap > 0 Then
   yzoom = yzoom - overlap
   SorcY = yzoom - 2 * ZoomSize(ZoomSpan)
End If

DoEvents

Success& = StretchBlt(DestDC&, DestX, DestY, DestWidth, DestHeight, _
SorcDC&, SorcX, SorcY, SorcWidth, SorcHeight, dwRop&)
'Put grid on top of zoomed picture
DrawZoomGrid
picZOOMBox.Refresh
MouseTArr
'MousePointer = vbDefault
End Sub
Private Sub TransferActRectTopicMCRBox()
DestDC& = picMCRBox.hdc
SorcDC& = picCanvas.hdc
dwRop& = &HCC0020  'SRCCOPY Src to Dest
SorcX = arleft + 1: SorcY = artop + 1
DestX = 0
DestY = 0
DestWidth = arright - arleft - 1
DestHeight = arbottom - artop - 1
Success& = BitBlt(DestDC&, DestX, DestY, DestWidth, DestHeight, SorcDC&, SorcX, SorcY, dwRop&)
End Sub
Private Sub LRReflectActRectTopicMCRBox()
'Left-Right reflection inside Active rect to picMCRBox
DestDC& = picMCRBox.hdc
SorcDC& = picCanvas.hdc
dwRop& = &HCC0020  'SRCCOPY Src to Dest
For SorcX = arleft + 1 To arright - 1
   SorcY = artop + 1
   DestX = picMCRBox.Width - (SorcX - (arleft + 1)) - 3
   DestY = 0
   DestWidth = 1
   DestHeight = arbottom - 1 - (artop + 1) + 1
   Success& = BitBlt(DestDC&, DestX, DestY, DestWidth, DestHeight, SorcDC&, SorcX, SorcY, dwRop&)
Next SorcX
End Sub
Private Sub TBReflectActRectTopicMCRBox()
'Top-Bottom reflection inside Active rect to picMCRBox
DestDC& = picMCRBox.hdc
SorcDC& = picCanvas.hdc
dwRop& = &HCC0020  'SRCCOPY Src to Dest
For SorcY = artop + 1 To arbottom - 1
   SorcX = arleft + 1
   DestY = picMCRBox.Height - (SorcY - (artop + 1)) - 3
   DestX = 0
   DestHeight = 1
   DestWidth = arright - 1 - (arleft + 1) + 1
   Success& = BitBlt(DestDC&, DestX, DestY, DestWidth, DestHeight, SorcDC&, SorcX, SorcY, dwRop&)
Next SorcY
End Sub

Private Sub ResizeInActiveRect()

zsqueeze = Val(txtRSR.Text)
If zsqueeze > 50 Then
   txtRSR.Text = "50"
ElseIf zsqueeze < -50 Then
   txtRSR.Text = "-50"
End If
zsqueeze = Val(txtRSR.Text)

zsq = 1 + zsqueeze / 100

picMCRBox.Left = karleft
picMCRBox.Top = kartop
'these two above not really necessary
picMCRBox.Width = arright - arleft - 1
picMCRBox.Height = arbottom - artop - 1
'Transfer inside picCanvas Active rect
'to picMCRBox multiplying by zsq (ie +/- zsqueeze %)
picMCRBox.Cls
picMCRBox.BackColor = RubberCul&
picMCRBox.Line (0, 0)-Step(picMCRBox.Width, picMCRBox.Height), RubberCul&, BF
DestDC& = picMCRBox.hdc
DestX = 0
DestY = 0
DestWidth = zsq * picMCRBox.Width
DestHeight = zsq * picMCRBox.Height
SorcDC& = picCanvas.hdc
SorcX = arleft + 1
SorcY = artop + 1
SorcWidth = picMCRBox.Width
SorcHeight = picMCRBox.Height
dwRop& = &HCC0020  'SRCCOPY Src to Dest
DoEvents

Success& = StretchBlt(DestDC&, DestX, DestY, DestWidth, DestHeight, _
SorcDC&, SorcX, SorcY, SorcWidth, SorcHeight, dwRop&)

picMCRBox.Refresh

picMCRBox.Visible = True

End Sub

Private Sub StartPerspectiveShear(Button, X, Y)
If Button = 1 Then
   CANVASRESTORE
   'Test if in Active rectangle If So exit
   If X >= arleft And X <= arright And Y >= artop And Y <= arbottom Then
      Exit Sub
   Else  'X,Y outside AR
      MousePointer = vbHourglass
      PerspectiveSW = True
      '--------------
      PerspectiveShear X, Y
      '--------------
      MouseTArr
      'MousePointer = vbDefault
   End If
   'DrawingMode left = True Prevents repeated CANVASSAVE
Else
   PerspectiveSW = False
   DrawingMode = False
   cmdClearActRect_Click
   a = ActionInProgress
End If
End Sub

Private Sub PerspectiveShear(X, Y)
'X,Y lies outside active rect

svpicCanvasDrawMode = picCanvas.DrawMode
picCanvas.DrawMode = 13

picMCRBox.BackColor = picCanvas.BackColor
picMCRBox.Cls


zdx = karright - karleft
zdy = karbottom - kartop
x1 = karleft: y1 = kartop
x2 = karright: y2 = karbottom

zdx = x2 - x1 - 1
zdy = y2 - y1 - 1
xc = (x1 + x2) / 2

'--------------------
If X < x1 Or X > x2 Then
   If X - x1 = 0 Then X = X + 1

   If X < x1 Then
      'To left of ActRect
      zm1 = (Y - y1) / (X - x2)
      zm2 = (Y - y2) / (X - x2)
   ElseIf X > x2 Then
      'To right of ActRect
      zm1 = (Y - y1) / (X - x1)
      zm2 = (Y - y2) / (X - x1)
   End If
   'Set size of picMCRBox
   If Y > y2 Then
      xL1 = 0
      xL2 = (x2 - x1) * (Y - y2) / (X - x1)
   ElseIf Y < y1 Then
      xL1 = (x2 - x1) * (y1 - Y) / (X - x1)
      xL2 = 0
   End If
   picMCRBox.Left = karleft
   picMCRBox.Top = kartop - Abs(xL1)
   picMCRBox.Width = karright - karleft + 4
   picMCRBox.Height = karbottom - kartop + 4 + Abs(xL1) + Abs(xL2)
   picCanvas.DrawMode = 13
   
   ypmax = Y - zm1 * (X - x2)
   yp2max = Y - zm2 * (X - x2)
   If yp2max > ymax Then ymax = yp2max
   ylen = Sqr((x2 - x1) ^ 2 + (ypmax - yp1) ^ 2)
   zstep = (x2 - x1) / ylen
   For zi = x1 To x2 Step 0.1 'zstep '0.1   'Better but slower
      xx = zi - picMCRBox.Left
      yp1 = Y - zm1 * (X - zi)
      yp2 = Y - zm2 * (X - zi)
      zddy = (yp2 - yp1) / (zdy)
      yy = yp1
      For zj = y1 To y2
         c0& = GetPixel(picCanvas.hdc, zi, zj)
         zres = SetPixelV(picMCRBox.hdc, xx, yy - picMCRBox.Top, c0&)
         yy = yy + zddy
      Next zj
   Next zi
'--------------------
ElseIf X > x1 And X < x2 And Y < y1 Then
'Above ActRect
   zm1 = (y2 - Y) / (X - x1)
   zm2 = (y2 - Y) / (x2 - X)
   picMCRBox.Left = karleft
   picMCRBox.Top = kartop
   picMCRBox.Width = karright - karleft + 4
   picMCRBox.Height = karbottom - kartop + 4
   For zj = y1 To y2 Step 0.1
      xp1 = X - (zj - Y) / zm1
      xp2 = X + (zj - Y) / zm2
      zddx = (xp2 - xp1) / zdx
      xx = xp1
      For zi = x1 To x2
         c0& = GetPixel(picCanvas.hdc, zi, zj)
         zres = SetPixelV(picMCRBox.hdc, xx - picMCRBox.Left, zj - picMCRBox.Top, c0&)
         xx = xx + zddx
      Next zi
   Next zj

'--------------------
ElseIf X > x1 And X < x2 And Y > y2 Then
'Below ActRect
   zm1 = (Y - y1) / (X - x1)
   zm2 = (Y - y1) / (x2 - X)
   picMCRBox.Left = karleft
   picMCRBox.Top = kartop
   picMCRBox.Width = karright - karleft + 4
   picMCRBox.Height = karbottom - kartop + 4
   For zj = y1 To y2 Step 0.1
      xp1 = x1 + (zj - y1) / zm1
      xp2 = x2 - (zj - y1) / zm2
      zddx = (xp2 - xp1) / zdx
      xx = xp1
      For zi = x1 To x2
         c0& = GetPixel(picCanvas.hdc, zi, zj)
         zres = SetPixelV(picMCRBox.hdc, xx - picMCRBox.Left, zj - picMCRBox.Top, c0&)
         xx = xx + zddx
      Next zi
   Next zj
End If
'--------------------

t% = DoEvents()
picMCRBox.Refresh

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'TRANSFER image in picMCRBox to picCanvas
DestDC& = picCanvas.hdc
DestX = picMCRBox.Left
DestY = picMCRBox.Top
DestWidth = picMCRBox.Width
DestHeight = picMCRBox.Height

SorcDC& = picMCRBox.hdc
SorcX = 0
SorcY = 0
dwRop& = &HCC0020  'SRCCOPY Src to Dest

Success& = BitBlt(DestDC&, DestX, DestY, DestWidth, DestHeight, _
SorcDC&, SorcX, SorcY, dwRop&)

picMCRBox.DrawWidth = 1
picMCRBox.AutoSize = False
picMCRBox.Visible = False
picCanvas.Refresh


picCanvas.DrawMode = svpicCanvasDrawMode

End Sub
Private Sub Rotator()

zangdeg = Val(txtRSR.Text)
If zangdeg > 180 Then
   txtRSR.Text = "180"
ElseIf zangdeg < -180 Then
   txtRSR.Text = "-180"
End If
zangdeg = Val(txtRSR.Text)
zang = zangdeg / r2d#  'degrees to radians
'Active rect size
vw = karright - karleft - 1
vh = karbottom - kartop - 1
'Set picMCRBox size to receive whole rotated image
zW = Abs(vw * Cos(zang)) + Abs(vh * Sin(zang))
zH = Abs(vh * Cos(zang)) + Abs(vw * Sin(zang))

vwdiff = zW - vw
vhdiff = zH - vh
picMCRBox.Width = zW + 2
picMCRBox.Height = zH + 2
picMCRBox.Left = karleft
picMCRBox.Top = kartop
picMCRBox.BackColor = picCanvas.BackColor
picMCRBox.Cls

' Compute the centers.
c1x = (karleft + karright) / 2
c1y = (kartop + karbottom) / 2
c2x = picMCRBox.Width / 2
c2y = picMCRBox.Height / 2

' Compute half-image size.
nh = vh / 2
nw = vw / 2
For p1y = 0 To nh '+ 1
   For p1x = 0 To nw '+ 1
      ' Compute polar coordinate of p1.
      
      zp1y = 1& * p1y
      zp1x = 1& * p1x
      za = zAtn2(zp1y, zp1x)
      zr = Sqr(zp1x * zp1x + zp1y * zp1y)
      ' Compute rotated position of p1 -> p2.
      p2x = zr * Cos(za + zang)
      p2y = zr * Sin(za + zang)
      p3x = zr * Cos(za - zang)
      p3y = zr * Sin(za - zang)
      
      ' Get 4 p1 pixel colours
      ' Set these 4 pixel colours @ p2 rotated positions
         c0& = GetPixel(picCanvas.hdc, c1x + p1x, c1y + p1y)
            xx = SetPixelV(picMCRBox.hdc, c2x + p2x, c2y + p2y, c0&)
         
            xx = SetPixelV(picMCRBox.hdc, c2x + p2x - 1, c2y + p2y - 1, c0&)
            xx = SetPixelV(picMCRBox.hdc, c2x + p2x - 1, c2y + p2y + 1, c0&)
            xx = SetPixelV(picMCRBox.hdc, c2x + p2x + 1, c2y + p2y + 1, c0&)
            xx = SetPixelV(picMCRBox.hdc, c2x + p2x + 1, c2y + p2y - 1, c0&)
         
         c1& = GetPixel(picCanvas.hdc, c1x - p1x, c1y - p1y)
            xx = SetPixelV(picMCRBox.hdc, c2x - p2x, c2y - p2y, c1&)
         
            xx = SetPixelV(picMCRBox.hdc, c2x - p2x - 1, c2y - p2y - 1, c1&)
            xx = SetPixelV(picMCRBox.hdc, c2x - p2x - 1, c2y - p2y + 1, c1&)
            xx = SetPixelV(picMCRBox.hdc, c2x - p2x + 1, c2y - p2y + 1, c1&)
            xx = SetPixelV(picMCRBox.hdc, c2x - p2x + 1, c2y - p2y - 1, c1&)
         
         c2& = GetPixel(picCanvas.hdc, c1x - p1x, c1y + p1y)
            xx = SetPixelV(picMCRBox.hdc, c2x - p3x, c2y + p3y, c2&)
         
            xx = SetPixelV(picMCRBox.hdc, c2x - p3x - 1, c2y + p3y - 1, c2&)
            xx = SetPixelV(picMCRBox.hdc, c2x - p3x - 1, c2y + p3y + 1, c2&)
            xx = SetPixelV(picMCRBox.hdc, c2x - p3x + 1, c2y + p3y + 1, c2&)
            xx = SetPixelV(picMCRBox.hdc, c2x - p3x + 1, c2y + p3y - 1, c2&)
         
         c3& = GetPixel(picCanvas.hdc, c1x + p1x, c1y - p1y)
            xx = SetPixelV(picMCRBox.hdc, c2x + p3x, c2y - p3y, c3&)
   
            xx = SetPixelV(picMCRBox.hdc, c2x + p3x - 1, c2y - p3y - 1, c3&)
            xx = SetPixelV(picMCRBox.hdc, c2x + p3x - 1, c2y - p3y + 1, c3&)
            xx = SetPixelV(picMCRBox.hdc, c2x + p3x + 1, c2y - p3y + 1, c3&)
            xx = SetPixelV(picMCRBox.hdc, c2x + p3x + 1, c2y - p3y - 1, c3&)
   
   Next
   ' Allow pending Windows messages to be processed.
   t% = DoEvents()
   picMCRBox.Refresh
Next
picMCRBox.Visible = True

End Sub

Private Sub TileActiveRectangle(X, Y)
svdrawMode = picCanvas.DrawMode
picCanvas.DrawMode = 13

picMCRBox.Left = X + 4 'picCanvas.Left + arleft
picMCRBox.Top = Y + 4 'picCanvas.Top + artop
'these two above not really necessary
picMCRBox.BorderStyle = 0
picMCRBox.Width = arright - arleft - 1 ' + 1
picMCRBox.Height = arbottom - artop - 1 ' + 1
picMCRBox.Picture = LoadPicture
'--------------------
TransferActRectTopicMCRBox
'--------------------

DrawingMode = False
cmdClearActRect_Click
ShowActRectCoords

picMCRBox.Refresh

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'TRANSFER multiply, picMCRBox to picCanvas

Heights = 2 * picCanvas.Height / picMCRBox.Height
Widths = 2 * picCanvas.Width / picMCRBox.Width

DestDC& = picCanvas.hdc

DestWidth = picMCRBox.Width
DestHeight = picMCRBox.Height
SorcDC& = picMCRBox.hdc
SorcX = 0
SorcY = 0
dwRop& = &HCC0020  'SRCCOPY Src to Dest

DestY = 0
For yy = 1 To Heights
   DestX = 0
   For xx = 1 To Widths

      Success& = BitBlt(DestDC&, DestX, DestY, DestWidth, DestHeight, _
      SorcDC&, SorcX, SorcY, dwRop&)

      DestX = DestX + picMCRBox.Width - 1
   Next xx
   DestY = DestY + picMCRBox.Height - 1
Next yy

picMCRBox.BorderStyle = 1
picMCRBox.DrawWidth = 1
picMCRBox.AutoSize = False
picMCRBox.Visible = False
picCanvas.Refresh
picCanvas.DrawMode = svdrawMode

End Sub

Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim culb As Byte
ShowXY X, Y

'ZOOM BOX or PERSPECTIVE POINTS or DRAWING PICTURE START

If ZoomMode = True Then
   StartZoom X, Y
   Exit Sub
End If
   
If SetPerspecPts Then
   SetVanishingPoints Button, X, Y
   Exit Sub
End If

If AddBMPSW = True Then Exit Sub
If picMCRBox.Visible = True Then Exit Sub

If ActiveRectExists And Not InRectangle(X, Y, arleft, artop, arright, arbottom) Then
   Select Case TOOL
   Case MCR, Resize, Rotate, Tile: Exit Sub
   End Select
End If

If Not ActiveRectExists Then
   MSGneed = False
   Select Case TOOL
   Case MCR, Resize, PerspecShear, Rotate, Tile
      MSGneed = True
   Case Rubber
      If RubberType = 3 Then MSGneed = True
   Case Smudge
      If SmudgeType = 1 Then MSGneed = True
   End Select
   If MSGneed = True Then
      resp = RRMsgBox$("RRPencil", "Need to draw an active rectangle", 0, 0, 1, 1) 'YES,NO,OK, NumLines
      Exit Sub
   End If
End If

If DrawingMode = False Then  'Start new shape
   
   CANVASSAVE
   
   PCount = 0  'start new point count
   RedrawPolCoTu = 0
   zpspac = 0  'parallel line spacing off
   LCount = 0  'start new LC count
   picCanvas.DrawWidth = 1
   picCanvas.DrawStyle = vbSolid
   
   DrawingMode = True
   
   'Set DrawCul&
   If Button = 1 Then
      DrawCul& = LeftCul&
   ElseIf Button = 2 Then
      DrawCul& = RightCul&
   End If
   'Show DrawCul&
   GetGreyCN DrawCul&, culb
   LabCN.Caption = Str$(255 - culb)
   LabCUL.BackColor = DrawCul&

End If

picCanvas.FontTransparent = False

Select Case TOOL
Case Brush: StartBrush X, Y
Case AirBrush: StartAirBrush X, Y
Case ALine: StartLine X, Y 'ie 2 point shape
Case APolyline: StartPolyLine Button, X, Y 'ie multi-point shape
Case Spline: StartSpline Button, X, Y 'Similar logic to PolyLine ie multi-point shape
Case Rectangle: StartRectangle X, Y 'Similar logic to ALine ie 2 point shape
Case Cirlipse: StartCirlipse X, Y 'Similar logic to ALine ie 2 point shape
Case Cone: StartCone X, Y '3 left clicks
Case Tube: StartTube X, Y
Case Arch: StartArch X, Y 'Similar logic to ALine ie 2 point shape
Case TPiece: StartTPiece X, Y 'Similar logic to ALine ie 2 point shape
Case AText: StartText X, Y
Case Fill: StartFill X, Y 'Single click does whole action
Case Rubber: StartRubber X, Y 'Similar logic to ALine in that 2nd click clears rubber shape
Case ActiveRect: StartActiveRectangle X, Y 'Similar logic to ALine ie 2 point shape
Case Smudge: StartSmudge X, Y  'Single click does whole action
Case MCR: StartMCR X, Y 'MoveCopyReflect
Case Resize: StartResize X, Y
Case PerspecShear: StartPerspectiveShear Button, X, Y 'PerspecShear +/- degrees
Case Rotate: StartRotate X, Y 'Rotate +/- degrees
Case Tile: StartTile X, Y

End Select
End Sub

Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   

Dim culb As Byte
ShowXY X, Y
a = ActionInProgress

If SetPerspecPts = True Then
   If InstrucCaptions.Caption <> PerSpecInstr$ Then
      InstrucCaptions.Caption = PerSpecInstr$
   End If
   Exit Sub
End If

If InstrucCaptions.Caption <> Instruction$(TOOL) Then
   If Not ZoomMode Then
      InstrucCaptions.Caption = Instruction$(TOOL)
   Else
      If InstrucCaptions.Caption <> ZoomInstr$ Then
         InstrucCaptions.Caption = ZoomInstr$
      End If
   End If
End If
   

'HAIR-LINES
'Global ShowPerspecLines, SetPerspecPts,
'       YPr, XPr1, XPr2, XPr3, NumPPts
W = picCanvas.Width
H = picCanvas.Height
'--
Line1.x1 = X: Line1.y1 = 0
Line1.x2 = X: Line1.y2 = H
'|
Line2.x1 = 0: Line2.y1 = Y
Line2.x2 = W: Line2.y2 = Y
'\
Line3.x1 = X - Y:     Line3.y1 = 0
Line3.x2 = X - Y + H: Line3.y2 = H

'/
Line4.x1 = X + Y:     Line4.y1 = 0
Line4.x2 = X + Y - H: Line4.y2 = H
'-
Line5.x1 = 0: Line5.y1 = YPr
Line5.x2 = W: Line5.y2 = YPr
'\
Line6.x1 = XPr1: Line6.y1 = YPr
Line6.x2 = X:    Line6.y2 = Y
'/
If NumPPts > 1 Then
  Line7.x1 = X:    Line7.y1 = Y
  Line7.x2 = XPr2: Line7.y2 = YPr
End If
'/
If NumPPts > 2 Then
  Line8.x1 = X:    Line8.y1 = Y
  Line8.x2 = XPr3: Line8.y2 = YPr
End If

If Not ActionInProgress Then  'Show picture colors
   cul& = picCanvas.Point(X, Y)
   If cul& <> -1 Then
      GetGreyCN cul&, culb
      LabCN.Caption = Str$(255 - culb)
      LabCUL.BackColor = cul&
   End If
   Exit Sub
End If

If AddBMPSW = True Then Exit Sub
If ZoomMode = True Then Exit Sub 'ie Mousemove plays no part when zooming
'picCanvas.SetFocus


If DrawingMode = False Then Exit Sub

Select Case TOOL
Case Brush: DrawBrush X, Y
Case AirBrush: DrawAirBrush X, Y
Case ALine: DrawLine X, Y
Case APolyline: DrawPolyLine Button, X, Y
Case Spline: ADrawSpline Button, X, Y
Case Rectangle: DrawRectangle X, Y
Case Cirlipse: DrawCirlipse X, Y
Case Cone: DrawCone X, Y  'LC count used to decide Ellipse or cone point
Case Tube: DrawTube X, Y
Case Arch: DrawArch X, Y
Case TPiece: ADrawTPiece X, Y
Case Rubber: DrawRubber X, Y
Case ActiveRect: DrawActiveRect X, Y 'Similar logic to ALine ie 2 point shape

End Select
End Sub
Private Sub picCanvas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ShowXY X, Y
Select Case TOOL
Case Brush: DrawingMode = False
Case AirBrush: DrawingMode = False
Case ALine
Case APolyline
Case Spline
Case Rectangle
Case Cirlipse
Case Cone
Case Tube
Case Arch
Case TPiece

Case AText
Case Fill
Case Rubber
Case ActiveRect
Case Smudge: DrawingMode = False
Case MCR
Case Resize
Case PerspecShear
Case Rotate
Case Tile
End Select
End Sub

Private Sub cmdFont_Click()
CommonDialog1.CancelError = True
CommonDialog1.DialogTitle = "Select font"
CommonDialog1.Flags = &H103
If CommonDialog1.FontName = "" Then
   CommonDialog1.FontName = "Arial"
End If
On Error GoTo FontError
CommonDialog1.ShowFont
With Text2
.FontName = CommonDialog1.FontName
.FontSize = CommonDialog1.FontSize
.ForeColor = CommonDialog1.Color
.FontBold = CommonDialog1.FontBold
.FontItalic = CommonDialog1.FontItalic
.FontUnderline = CommonDialog1.FontUnderline
End With
With picCanvas
.FontName = CommonDialog1.FontName
.FontSize = CommonDialog1.FontSize
.ForeColor = CommonDialog1.Color
.FontBold = CommonDialog1.FontBold
.FontItalic = CommonDialog1.FontItalic
.FontUnderline = CommonDialog1.FontUnderline
End With
With picMCRBox
.FontName = CommonDialog1.FontName
.FontSize = CommonDialog1.FontSize
.ForeColor = CommonDialog1.Color
.FontBold = CommonDialog1.FontBold
.FontItalic = CommonDialog1.FontItalic
.FontUnderline = CommonDialog1.FontUnderline
End With
Exit Sub
'============
FontError:
Exit Sub
End Sub

Private Sub cmdTextCancel_Click()
frmText.Visible = False
TextSW = False
chkTOOLS(TOOL).Value = vbUnchecked '0
chkTOOLS(prevTOOL).Value = vbChecked '1
TOOL = prevTOOL
ShowInstructionsAndSubToolBars
MousePointer = 0
End Sub

Private Sub picMCRBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ShowXY X, Y
res = ActionInProgress
zDragX = X: zDragY = Y  'Needed for LC for picMCRBox_MouseMove
If Button = 2 Then 'RC complete MCR or Added BMP
   MousePointer = vbHourglass
   'If ActiveRectExists Then
   '   'Clear active rectangle before copying
   '   DrawingMode = False
   '   cmdClearActRect_Click
   'End If
   
   If TextSW = True Then
      svFT = picCanvas.FontTransparent
      picCanvas.FontTransparent = True
      picCanvas.ForeColor = DrawCul&
      picCanvas.CurrentX = picMCRBox.Left + 2
      picCanvas.CurrentY = picMCRBox.Top + 2
      'ie +2 if picMCRBox has a border
      TextLine$ = Text2.Text
      TextAngle = Val(txtAngle.Text)
      RotateText TextAngle, TextLine$
      
      'picCanvas.Print Text2.Text;
      'THIS PREVENTS DrawWidth>1 ?????
      picCanvas.FontTransparent = False
      TextSW = False
      
      chkTOOLS(TOOL).Value = vbUnchecked '0
      chkTOOLS(prevTOOL).Value = vbChecked '1
      TOOL = prevTOOL
      ShowInstructionsAndSubToolBars
      
      picCanvas.FontTransparent = svFT
      
   ElseIf AddBMPSW = True Then
        picMCRBoxToPicCanvas
        AddBMPSW = False
        
   ElseIf ResizeSW = True Then
      Select Case ResizeType
      Case 0   'MOVE
        BlankActiveRect
      End Select  'Else COPY
      picMCRBoxToPicCanvas
        
      picMCRBox.DrawWidth = 1
      ResizeSW = False
   
   ElseIf RotateSW = True Then
      Select Case RotateType
      Case 0   'MOVE
        BlankActiveRect
      End Select  'Else COPY
      picMCRBoxToPicCanvas
        
      picMCRBox.DrawWidth = 1
      RotateSW = False
   Else  '?? Resize gets here with MCRType=0
      Select Case MCRType
      Case 0   'MOVE
        BlankActiveRect
        picMCRBoxToPicCanvas
      Case 1   'COPY
        picMCRBoxToPicCanvas
      Case 2   'REFLECT LEFT->RIGHT
        picMCRBoxToPicCanvas
      Case 3   'REFLECT TOP->BELOW
        picMCRBoxToPicCanvas
      End Select
      'RedrawActiveRect   'For further copying
      MCRSW = False
   End If
   
   picMCRBox.Cls
   picMCRBox.Visible = False
   
   LCount = 0
   DrawingMode = False
   MouseTArr
   'MousePointer = vbDefault

End If
End Sub
Private Sub picMCRBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ShowXY X, Y
'Move & position copied rectangle in picMCRBox
If Button = 1 Then
   OffsetX = (X - zDragX): OffsetY = (Y - zDragY)
   xi = picMCRBox.Left + OffsetX
   yi = picMCRBox.Top + OffsetY
   picMCRBox.Move xi, yi
End If
End Sub
Private Sub BlankActiveRect()
svdrawMode = picCanvas.DrawMode
picCanvas.DrawMode = 13
'NB Needs to use sv because Active rect is temporarily cleared
cul& = picCanvas.BackColor
picCanvas.Line (karleft + 1, kartop + 1)-(karright - 1, karbottom - 1), cul&, BF
picCanvas.DrawMode = svdrawMode
End Sub
Private Sub picMCRBoxToPicCanvas()
svdrawMode = picCanvas.DrawMode
picCanvas.DrawMode = 13
SorcDC& = picMCRBox.hdc
DestDC& = picCanvas.hdc
dwRop& = &HCC0020  'SRCCOPY Src to Dest
SorcX = 0: SorcY = 0
DestX = picMCRBox.Left + 1 '- picCanvas.Left + 1
DestY = picMCRBox.Top + 1 '- picCanvas.Top + 1
DestWidth = picMCRBox.Width - 2
DestHeight = picMCRBox.Height - 2
Success& = BitBlt(DestDC&, DestX, DestY, DestWidth, DestHeight, SorcDC&, SorcX, SorcY, dwRop&)
picCanvas.DrawMode = svdrawMode
End Sub

Private Sub picPAL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If ActionInProgress Then Exit Sub
If Button = 1 Then   'Set left draw color
   DrawCul& = RGB(255 - Y, 255 - Y, 255 - Y)
   LeftCul& = RGB(255 - Y, 255 - Y, 255 - Y)
   LabLEFTCUL.BackColor = DrawCul&
   LabLEFTCN.Caption = Str$(Y)
   cn = Y
ElseIf Button = 2 Then   'Set right rubber color
   RubberCul& = RGB(255 - Y, 255 - Y, 255 - Y)
   RightCul& = RGB(255 - Y, 255 - Y, 255 - Y)
   LabRIGHTCUL.BackColor = RubberCul&
   LabRIGHTCN.Caption = Str$(Y)
   rubcn = Y
End If
End Sub
Private Sub picPAL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Show color
LabCN.Caption = Str$(Y)
LabCUL.BackColor = RGB(255 - Y, 255 - Y, 255 - Y)
End Sub


Private Sub PrintBMP_Click()
If ActionInProgress Then Exit Sub
LabWait.Visible = True

PrintFromIView    'In IView.bas

LabWait.Visible = False
End Sub

Private Sub UDmx_Change()
If ActionInProgress Then Exit Sub
End Sub
Private Sub UDmy_Change()
If ActionInProgress Then Exit Sub
End Sub

Private Sub UDZOOM_DownClick()
If ZoomMode = False Then Exit Sub

If ActiveRectExists Then
   cmdClearActRect_Click   'Clears picCanvas & picZOOMBox Act Rects
End If

ZoomSpan = ZoomSpan + 1
If ZoomSpan > 12 Then ZoomSpan = 12: Exit Sub
   MousePointer = vbHourglass
   GridStep = 240 / (2 * ZoomSize(ZoomSpan))
   picZOOMBox.Cls
   picZOOMBox.Visible = False
   'Position Zoom grid
   PositionZoomBox xzoom
   'Transfer picCanvas points to  Zoom grid
   TransferCanvasToZoomBox 'Global xzoom, yzoom
   picZOOMBox.Visible = True
   picZOOMBox.SetFocus
   MouseTArr
   'MousePointer = vbDefault
End Sub
Private Sub UDZOOM_UpClick()
If ZoomMode = False Then Exit Sub

If ActiveRectExists Then
   cmdClearActRect_Click   'Clears picCanvas & picZOOMBox Act Rects
End If

ZoomSpan = ZoomSpan - 1
If ZoomSpan < 0 Then ZoomSpan = 0: Exit Sub
   MousePointer = vbHourglass
   GridStep = 240 / (2 * ZoomSize(ZoomSpan))
   picZOOMBox.Cls
   picZOOMBox.Visible = False
   'Position Zoom grid
   PositionZoomBox xzoom
   'Transfer picCanvas points to  Zoom grid
   TransferCanvasToZoomBox 'Global xzoom, yzoom
   picZOOMBox.Visible = True
   picZOOMBox.SetFocus
   MouseTArr
   'MousePointer = vbDefault
End Sub

Private Sub picZoomBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'ZOOM BOX pixel setting & Act Rect
If picZOOMBox.Point(X, Y) = -1 Then Exit Sub

'Find top left coord of zoom grid point
x1 = X Mod GridStep
y1 = Y Mod GridStep
x1 = GridStep * (X \ GridStep) + 1
y1 = GridStep * (Y \ GridStep) + 1
'Find & show equivalent picCanvas coords
xo = xzoom + (x1 \ GridStep - ZoomSize(ZoomSpan))
yo = yzoom + (y1 \ GridStep - ZoomSize(ZoomSpan) * 2)
ShowXY xo, yo

If chkZOOMAR.Value = vbChecked Then
     
   If ActiveRectExists And zoomdrawing = False Then
      cmdClearActRect_Click   'Clears picCanvas & picZOOMBox Act Rects
   End If

   zpspac = 0
   picZOOMBox.DrawStyle = vbDot
   picZOOMBox.DrawWidth = 2
   zoomdrawing = Not zoomdrawing
   If zoomdrawing = True Then
      zoomRectCul& = picZOOMBox.BackColor Xor QBColor(9)
      picZOOMBox.DrawMode = 7
      picCanvas.DrawMode = 7
      xs = X: ys = Y
      xp = X: yp = Y
      xs = x1 + GridStep \ 2: ys = y1 + GridStep \ 2
      xp = x1 + GridStep \ 2: yp = y1 + GridStep \ 2
      picZOOMBox.Line (xs, ys)-(xp, yp), zoomRectCul&, B
      
      arleft = xo: artop = yo
      arright = xo: arbottom = yo
      ShowActRectCoords
         
      shpARect.Visible = True
      shpARect.Left = arleft
      shpARect.Top = artop
      shpARect.Width = (arright - arleft)
      shpARect.Height = (arbottom - artop)
      
      ActiveRectExists = True
   Else  'end zoombox rect
      picZOOMBox.DrawMode = 13
      picZOOMBox.DrawStyle = vbSolid
      picZOOMBox.DrawWidth = 1
      picCanvas.DrawStyle = vbSolid
      picCanvas.DrawMode = 13
      zoomend = True 'to prevent MouseMove putting a dot on the zoomgrid
      Exit Sub
   End If
Else
   zoomend = False
   GridAction Button, x1, y1, xo, yo
End If

picZOOMBox.SetFocus
End Sub
Private Sub picZoomBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'ZOOM BOX pixel setting & Act Rect
If picZOOMBox.Point(X, Y) = -1 Then Exit Sub
x1 = X Mod GridStep
y1 = Y Mod GridStep
x1 = GridStep * (X \ GridStep) + 1
y1 = GridStep * (Y \ GridStep) + 1
'Find & show picCanvas coords
xo = xzoom + (x1 \ GridStep - ZoomSize(ZoomSpan))
yo = yzoom + (y1 \ GridStep - ZoomSize(ZoomSpan) * 2)
ShowXY xo, yo

If chkZOOMAR.Value = vbChecked And zoomdrawing = True Then
   picZOOMBox.Line (xs, ys)-(xp, yp), zoomRectCul&, B
   xp = X: yp = Y
   picZOOMBox.Line (xs, ys)-(xp, yp), zoomRectCul&, B
   
   arright = xo: arbottom = yo
   ShowActRectCoords
   shpARect.Left = arleft
   shpARect.Top = artop
   shpARect.Width = (arright - arleft)
   shpARect.Height = (arbottom - artop)

Else
   If Button <> 1 And Button <> 2 Or zoomend = True Then Exit Sub
   GridAction Button, x1, y1, xo, yo
End If

picZOOMBox.SetFocus
End Sub
Private Sub GridAction(Button, x1, y1, xo, yo)
'ZOOM BOX pixel setting
'From picZoomBox MouseDown & MouseMove
If Button = 1 Then
   cul& = LeftCul&
ElseIf Button = 2 Then
   cul& = RightCul&
End If
picZOOMBox.Line (x1, y1)-(x1 + (GridStep - 2), y1 + (GridStep - 2)), cul&, BF
'Show drawing on picCanvas
svpicCanvasDrawWidth = picCanvas.DrawWidth
picCanvas.DrawWidth = 1
picCanvas.PSet (xo, yo), cul&
picCanvas.DrawWidth = svpicCanvasDrawWidth
End Sub

Private Sub LoadFile_Click()
If ActionInProgress Then Exit Sub

If ActiveRectExists Then cmdClearActRect_Click

CANVASSAVE

Title$ = "Load picture file"
Choice$ = "Pics(*.bmp *.jpg *.gif *.wmf *.emf *.ico *.cur)|*.bmp;*.jpg;*.gif;*.wmf;*.emf;*.ico;*.cur"
InitDir$ = CurrentDirec$
OpenLoadDialog Title$, Choice$, LoadFileSpec$, InitDir$
If LoadFileSpec$ <> "" Then
   CurrentDirec$ = ExtractPath(LoadFileSpec$)
   
   picCanvas.Picture = LoadPicture(LoadFileSpec$) ', , vbLPColor)
   
   Dim bmp As BITMAP
   GetObjectAPI picCanvas.Picture, Len(bmp), bmp
   
   bmw$ = ""
   bmw$ = bmw$ & "KEEP FILE ?" & vbCrLf & vbCrLf
   bmw$ = bmw$ & "Width =" & Str$(bmp.bmWidth) & vbCrLf
   bmw$ = bmw$ & "Height =" & Str$(bmp.bmHeight) & vbCrLf
   
res = RRMsgBox$(ExtractFileName$(LoadFileSpec$), bmw$, 1, 1, 0, 4) 'YES,NO,OK, NumLines
   
   If res = vbNo Then
      CANVASRESTORE
      DrawingMode = False
   End If
End If
End Sub

Private Sub AddFile_Click()
If ActionInProgress Then Exit Sub
If ActiveRectExists Then cmdClearActRect_Click

'Load picture into picMCRBox & transfer to picCanvas
Dim bmp As BITMAP
Title$ = "Add picture file"
Choice$ = "Pics(*.bmp *.jpg *.gif *.wmf *.emf *.ico *.cur)|*.bmp;*.jpg;*.gif;*.wmf;*.emf;*.ico;*.cur"
InitDir$ = CurrentDirec$
OpenLoadDialog Title$, Choice$, LoadFileSpec$, InitDir$

If LoadFileSpec$ <> "" Then
   
   CurrentDirec$ = ExtractPath(LoadFileSpec$)

   LoadFileSpec$ = CommonDialog1.FileName
   FName$ = ExtractFileName(LoadFileSpec$)

   CANVASSAVE
   AddBMPSW = True
   DrawingMode = True

   picMCRBox.Cls
   picMCRBox.Top = 50
   picMCRBox.Left = 50
   picMCRBox.Picture = LoadPicture(LoadFileSpec$, , vbLPColor)
   picMCRBox.Refresh
   picMCRBox.Visible = True

   GetObjectAPI picMCRBox.Picture, Len(bmp), bmp
 
   'Restrict metafiles whose width & height = 0
   '& would fill whole picture box
   If bmp.bmWidth = 0 Then bmp.bmWidth = 100
   picMCRBox.Width = bmp.bmWidth
   If bmp.bmHeight = 0 Then bmp.bmHeight = 100
   picMCRBox.Height = bmp.bmHeight
   
   bmw$ = ""
   bmw$ = bmw$ & "KEEP FILE ?" & vbCrLf & vbCrLf
   bmw$ = bmw$ & "Width =" & Str$(bmp.bmWidth) & vbCrLf
   bmw$ = bmw$ & "Height =" & Str$(bmp.bmHeight) & vbCrLf & vbCrLf
   bmw$ = bmw$ & "If KEPT LC-HOLD-MOVE on" & vbCrLf & " added picture" & vbCrLf
   bmw$ = bmw$ & " && RC to fix"
   FName$ = ExtractFileName$(LoadFileSpec$)
   res = RRMsgBox$(FName$, bmw$, 1, 1, 0, 6) 'YES,NO,OK, NumLines
   
   If res = vbNo Then
      picMCRBox.Cls
      picMCRBox.Visible = False
      CANVASRESTORE
      AddBMPSW = False
      DrawingMode = False
   Else
      MousePointer = 10
   End If
End If
picCanvas.Refresh
End Sub

Private Sub SaveBMP_Click()
If ActionInProgress Then Exit Sub

Title$ = "Save BMP file"
Choice$ = "BMP files(*.bmp)|*.bmp"
InitDir$ = CurrentDirec$
SFile$ = ""
OpenSaveDialog Title$, Choice$, SaveFileSpec$, InitDir$, SFile$
If SaveFileSpec$ <> "" Then
   FixFileExtension SaveFileSpec$, "bmp"
   CurrentDirec$ = ExtractPath(SaveFileSpec$)
   
   SavePicture picCanvas.Image, SaveFileSpec$
   'NB .Picture saves as original size, .Image as whole picture box
End If
End Sub

Private Sub SaveJPG_Click()
If ActionInProgress Then Exit Sub

'SAVE PICTURE AS JPG @ WIDTH=600
'& HEIGHT OF 520 OR 720

SaveAsJPG   'In IView.bas

End Sub

Private Sub SaveRectangle8BitBMP_Click()

Dim bwidth As Long
Dim bheight As Long
Dim icb As Byte

If ActionInProgress Then Exit Sub

'After ensuring bwidth divisible by 4
'save rectangle 1 pixel INSIDE Active rectangle
'NB if bwidth has to be increased then right hand
'part of rectangle could include colors on and
'outside the right hand part of the Active rectangle.

If Not ActiveRectExists Then
   resp = RRMsgBox$("RRPencil", "Need to draw an active rectangle", 0, 0, 1, 1) 'YES,NO,OK, NumLines
   Exit Sub
End If

'Ensure 1->2 positive direction
FixRectCoords arleft, arright, artop, arbottom

arwidth = arright - arleft - 1
arheight = arbottom - artop - 1
If arwidth <= 0 Or arheight <= 0 Then
   resp = RRMsgBox$("RRPencil", "Active rectangle width < 1", 0, 0, 1, 1) 'YES,NO,OK, NumLines
   cmdClearActRect_Click
   Exit Sub
End If

'Save these so that kartop etc not chnged
axleft = arleft
axright = arright
aytop = artop
aybottom = arbottom

bwidth = CLng(arwidth)
bheight = CLng(arheight)

cmdClearActRect_Click

'Must make bwidth divisible by 4
'by extending axright
iremainder = bwidth Mod 4
If iremainder <> 0 Then
   axright = axright + 4 - iremainder
   bwidth = CLng(axright - axleft - 1&)
End If

Title$ = "Save BMP Rectangle"
Choice$ = "BMP files(*.bmp)|*.bmp"
InitDir$ = CurrentDirec$
SFile$ = ""
OpenSaveDialog Title$, Choice$, SaveFileSpec$, InitDir$, SFile$
If SaveFileSpec$ <> "" Then
   
   FixFileExtension SaveFileSpec$, "bmp"
   CurrentDirec$ = ExtractPath(SaveFileSpec$)
   
   MousePointer = vbHourglass
   LabWait.Visible = True
   'Shorten file else binary output only overwrites
   Open SaveFileSpec$ For Output As #1
   Print #1, " "
   Close #1

   'SAVE as 8-bit BITMAP
   Open SaveFileSpec$ For Binary As #1
   BM = 19778:     Put #1, , BM
   fsize& = bwidth * bheight + 54 + 1024
                   Put #1, , fsize&
   dblint& = 0:    Put #1, , dblint&
   offset& = 1078: Put #1, , offset&
   bmh& = 40:      Put #1, , bmh&
                   Put #1, , bwidth
                   Put #1, , bheight
   plane = 1:      Put #1, , plane
   bpp = 8:        Put #1, , bpp
   comp& = 0:      Put #1, , comp&
   imsiz& = 0:     Put #1, , imsiz&
   hppm& = 0:      Put #1, , hpp&
   vppm& = 0:      Put #1, , vpp&
   cindex& = 0:    Put #1, , cindex&
   cimpor& = 0:    Put #1, , cimpor&
   'Put on grey palette
   For icb = 0 To 254
      Put #1, , icb
      Put #1, , icb
      Put #1, , icb
      Put #1, , icb
   Next icb
      Put #1, , icb
      Put #1, , icb
      Put #1, , icb
      Put #1, , icb

   'Find pixels' grey color number
   For IY = aybottom - 1 To aytop + 1 Step -1
   For IX = axleft + 1 To axright - 1
      'cul& = picCanvas.Point(IX, IY)
      cul& = GetPixel(picCanvas.hdc, IX, IY)
      If cul& <> -1 Then
         GetGreyCN cul&, icb
      Else
         icb = 255
      End If
      Put #1, , icb
   Next IX
   Next IY
   Close #1
   LabWait.Visible = False
   MouseTArr
   'MousePointer = vbDefault
End If
End Sub

Private Sub CANVASSAVE()
'BitBlt method
hSorcDC& = picCanvas.hdc
hDestDC& = picCanvasStore.hdc
dwRop& = &HCC0020         'Src to Dest
SorcX = 0: SorcY = 0
DestX = 0: DestY = 0
nWidth = picCanvasStore.Width
nHight = picCanvasStore.Height
Success& = BitBlt(hDestDC&, DestX, DestY, nWidth, nHight, hSorcDC&, SorcX, SorcY, dwRop&)
End Sub
Private Sub CANVASRESTORE()
'BitBlt method
hSorcDC& = picCanvasStore.hdc
hDestDC& = picCanvas.hdc
dwRop& = &HCC0020         'Src to Dest
SorcX = 0: SorcY = 0
DestX = 0: DestY = 0
nWidth = picCanvas.Width
nHight = picCanvas.Height
Success& = BitBlt(hDestDC&, DestX, DestY, nWidth, nHight, hSorcDC&, SorcX, SorcY, dwRop&)
picCanvas.Refresh
End Sub

Private Sub cmdClearActRect_Click()
'Global arleft, artop, arright, arbottom, arwidth, arheight 'Active rect coords
'If DrawingMode = True Or pzoomtool = ActiveRect Then Exit Sub
If DrawingMode = True Then Exit Sub
If ActiveRectExists = True Then
   shpARect.Visible = False
   'Keep Active rectangle coords in case it needs to be redrawn
   karleft = arleft: kartop = artop
   karright = arright: karbottom = arbottom
   arleft = 0: artop = 0
   arright = 0: arbottom = 0
   ShowActRectCoords
   If ZoomMode = True Then 'Clear picZoomBox active rectangle
      picZOOMBox.DrawStyle = vbDot
      picZOOMBox.DrawWidth = 2
      picZOOMBox.DrawMode = 7
      'xs,ys,xp,yp Global from picZoomBox
      picZOOMBox.Line (xs, ys)-(xp, yp), zoomRectCul&, B
      picZOOMBox.DrawMode = 13
      picZOOMBox.DrawStyle = vbSolid
      picZOOMBox.DrawWidth = 1
      ActiveRectExists = False
   End If
   ActiveRectExists = False
End If
End Sub
Private Sub RedrawActiveRect()
'NB assumes coords saved in k..
arleft = karleft: artop = kartop
arright = karright: arbottom = karbottom
ShowActRectCoords
shpARect.Visible = True
shpARect.Left = arleft
shpARect.Top = artop
shpARect.Width = (arright - arleft)
shpARect.Height = (arbottom - artop)
ActiveRectExists = True
End Sub

Private Sub chkHairs_Click(Index As Integer)
'Global ShowPerspecLines, SetPerspecPts,
'       YPr, XPr1, XPr2, XPr3, NumPPts

If ActionInProgress Then
   If ShowPlusLines Then
      chkHairs(0).Value = vbChecked
   Else
      chkHairs(0).Value = vbUnchecked
   End If
   If ShowXLines Then
      chkHairs(1).Value = vbChecked
   Else
      chkHairs(1).Value = vbUnchecked
   End If
   Exit Sub
End If

Select Case Index
Case 0
   Line1.Visible = Not Line1.Visible
   Line2.Visible = Not Line2.Visible
   ShowPlusLines = Not ShowPlusLines
Case 1
   Line3.Visible = Not Line3.Visible
   Line4.Visible = Not Line4.Visible
   ShowXLines = Not ShowXLines
End Select

End Sub


Private Sub Exit_Click()
Form_QueryUnload Cancel, IUnloadMode
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then    'Close on Form1 pressed

resp = RRMsgBox$("RRPencil", "Quit Application ?", 1, 1, 0, 1) 'YES,NO,OK, NumLines
   
   If resp = vbNo Then
      Cancel = True
   Else  'vbYes
      Cancel = False
      Unload Me
      End
   End If
End If
End Sub

'=====================================================================
Public Sub DrawShadedLine(StartCount)
'LineType, cn, rubcn, zpspac
If StartCount < 1 Then Exit Sub
If PCount - StartCount < 1 Then Exit Sub

zsn = 255 - rubcn
If zsn < 0 Then zsn = 0
If zsn > 255 Then zsn = 255

If zpspac = 0 Then
   zculsteps = CSng(cn - rubcn) / 2
Else
   zculsteps = CSng(cn - rubcn) / zpspac
End If
'Find 1st inner parallel line pts
For zsp = 1 To zpspac

   cul& = RGB(Int(zsn), Int(zsn), Int(zsn)) 'Colors(Int(zsn))

   'Set 1st outer parallel line pts
   x1 = xsto(StartCount): y1 = ysto(StartCount)
   x2 = xsto(StartCount + 1): y2 = ysto(StartCount + 1)
   Findxeye zsp, x1, y1, x2, y2, xdd, ydd, xa, ya, xb, yb
   xa1 = xa: ya1 = ya    'Save start of paraline to close end
   
   If PCount - StartCount = 1 Then   'Single segment
      picCanvas.Line (xa, ya)-(xb, yb), cul&
   Else
      hpen& = CreatePen(picCanvas.DrawStyle, picCanvas.DrawWidth, cul&)
      hpenold& = SelectObject(picCanvas.hdc, hpen&)
      
      ReDim zp(StartCount To PCount - 1) As POINTAPI
      ReDim zpp(StartCount To PCount - 1) As POINTAPI
   
      xprv = xa: yprv = ya
      For m = StartCount + 1 To PCount - 1
         x1 = xsto(m - 1): y1 = ysto(m - 1)
         x2 = xsto(m): y2 = ysto(m)
         x3 = xsto(m + 1): y3 = ysto(m + 1)
         Findxeye zsp, x1, y1, x2, y2, xdd, ydd, xa, ya, xx, yx
         Findxeye zsp, x2, y2, x3, y3, xdd, ydd, xx, yx, xb, yb
         Findxiyi x1, y1, x2, y2, x3, y3, xdd, ydd, xa, ya, xb, yb, xi, yi
      
         If x1 < -30000 Then x1 = -30000
         If x1 > 30000 Then x1 = 30000
         If y1 < -30000 Then y1 = -30000
         If y1 > 30000 Then y1 = 30000
      
         zp(m - 1).kX = x1: zp(m - 1).kY = y1
      
         If x2 < -30000 Then x2 = -30000
         If x2 > 30000 Then x2 = 30000
         If y2 < -30000 Then y2 = -30000
         If y2 > 30000 Then y2 = 30000
      
         zp(m).kX = x2: zp(m).kY = y2
         zpp(m - 1).kX = xprv: zpp(m - 1).kY = yprv
      
         If xi < -30000 Then xi = -30000
         If xi > 30000 Then xi = 30000
         If yi < -30000 Then yi = -30000
         If yi > 30000 Then yi = 30000
      
         zpp(m).kX = xi: zpp(m).kY = yi
         xprv = xi: yprv = yi                         'x,yprv -> x,yi
      Next m
      xy = Polyline(picCanvas.hdc, zpp(1), PCount - StartCount)
      'Last parallel line pt
      picCanvas.Line (xi, yi)-(xb, yb), cul&
      zres = SelectObject(picCanvas.hdc, hpenold&)
      zres = DeleteObject(hpen&)
   End If
   
   zsn = zsn - zculsteps
   If zsn > 255 Then zsn = 255: zculsteps = -zculsteps
   If zsn < 1 Then zsn = 1: zculsteps = Abs(zculsteps)
   
Next zsp

Erase zp, zpp
End Sub
Public Sub DrawDoubleLine(cul&, StartCount)
'FINDS AND DRAWS MAIN & PARALLEL LINES zpspac AWAY FROM MAIN LINE
'Needs (Cul&,StartCount)
'& Global PCount, xsto(PCount),ysto(PCount) & zpspac para-spacing
'Cul& = PenCul& (mode 7) or DrawCul& (mode 13)
'picCanvas stores its points in xsto(), ysto()
If StartCount < 1 Then Exit Sub
If PCount - StartCount < 1 Then Exit Sub
'Set 1st outer parallel line pts
x1 = xsto(StartCount): y1 = ysto(StartCount)
x2 = xsto(StartCount + 1): y2 = ysto(StartCount + 1)
'Find 1st inner parallel line pts
Findxeye zpspac, x1, y1, x2, y2, xdd, ydd, xa, ya, xb, yb
xa1 = xa: ya1 = ya    'Save start of paraline to close end
   
If PCount - StartCount = 1 Then   'Single segment
   picCanvas.Line (x1, y1)-(x2, y2), cul&
   picCanvas.Line (xa, ya)-(xb, yb), cul&
Else
   
   hpen& = CreatePen(picCanvas.DrawStyle, picCanvas.DrawWidth, cul&)
   hpenold& = SelectObject(picCanvas.hdc, hpen&)
   ReDim zp(StartCount To PCount - 1) As POINTAPI
   ReDim zpp(StartCount To PCount - 1) As POINTAPI
   
   xprv = xa: yprv = ya
   For m = StartCount + 1 To PCount - 1
      x1 = xsto(m - 1): y1 = ysto(m - 1)
      x2 = xsto(m): y2 = ysto(m)
      x3 = xsto(m + 1): y3 = ysto(m + 1)
      Findxeye zpspac, x1, y1, x2, y2, xdd, ydd, xa, ya, xx, yx
      Findxeye zpspac, x2, y2, x3, y3, xdd, ydd, xx, yx, xb, yb
      Findxiyi x1, y1, x2, y2, x3, y3, xdd, ydd, xa, ya, xb, yb, xi, yi
      
      If x1 < -30000 Then x1 = -30000
      If x1 > 30000 Then x1 = 30000
      If y1 < -30000 Then y1 = -30000
      If y1 > 30000 Then y1 = 30000
      
      zp(m - 1).kX = x1: zp(m - 1).kY = y1
      
      If x2 < -30000 Then x2 = -30000
      If x2 > 30000 Then x2 = 30000
      If y2 < -30000 Then y2 = -30000
      If y2 > 30000 Then y2 = 30000
      
      zp(m).kX = x2: zp(m).kY = y2
      
      zpp(m - 1).kX = xprv: zpp(m - 1).kY = yprv
      
      If xi < -30000 Then xi = -30000
      If xi > 30000 Then xi = 30000
      If yi < -30000 Then yi = -30000
      If yi > 30000 Then yi = 30000
      
      zpp(m).kX = xi: zpp(m).kY = yi
      
      xprv = xi: yprv = yi                         'x,yprv -> x,yi
   Next m
   xy = Polyline(picCanvas.hdc, zp(1), PCount - StartCount)
   xy = Polyline(picCanvas.hdc, zpp(1), PCount - StartCount)
   zres = SelectObject(picCanvas.hdc, hpenold&)
   zres = DeleteObject(hpen&)
   
   'Last parallel line pt
   picCanvas.Line (x2, y2)-(x3, y3), cul&
   picCanvas.Line (xi, yi)-(xb, yb), cul&
   Erase zp, zpp
End If
End Sub
Public Sub DrawSpline(cul&, Button)
'Global LineType,, zpspac
'Global PCount
'Global xsto(), ysto()

If PCount <= 1 Then Exit Sub
If PCount = 2 Then
   If zpspac = 0 Then   'Single line
      picCanvas.Line (xsto(1), ysto(1))-(xsto(2), ysto(2)), cul&
   Else
      Select Case LineType
      Case 7, 8, 9
         If Button <= 1 Then
            DrawDoubleLine cul&, 1
         Else
            DrawShadedLine 1
         End If
      Case Else
         DrawDoubleLine cul&, 1
      End Select
   End If
   Exit Sub
Else  'PCount > 2
   'Save xsto() & ysto()
   ReDim xstosav(PCount), ystosav(PCount)
   For N = 1 To PCount
       xstosav(N) = xsto(N): ystosav(N) = ysto(N)
   Next N

   xfrac = 0.25
   SUP = 3
   oldpts = PCount
   For S = 1 To SUP
       ReDim xaa(oldpts), yaa(oldpts)
       For I = 1 To oldpts
           xaa(I) = xsto(I): yaa(I) = ysto(I)
       Next I
       newpts = 2 * oldpts - 2
       ReDim xsto(newpts), ysto(newpts)
       xsto(1) = xaa(1): ysto(1) = yaa(1)
       For I = 2 To oldpts - 1
           xdx = xaa(I) - xaa(I - 1)
           xsto(2 * I - 2) = xaa(I) - xfrac * xdx
           ydy = yaa(I) - yaa(I - 1)
           ysto(2 * I - 2) = yaa(I) - xfrac * ydy
           xdx = xaa(I + 1) - xaa(I)
           xsto(2 * I - 1) = xaa(I) + xfrac * xdx
           ydy = yaa(I + 1) - yaa(I)
           ysto(2 * I - 1) = yaa(I) + xfrac * ydy
       Next I
       xsto(newpts) = xaa(oldpts): ysto(newpts) = yaa(oldpts)
       oldpts = newpts
   Next S
      
   svPCount = PCount
   PCount = newpts
      
   If zpspac = 0 Then
      hpen& = CreatePen(picCanvas.DrawStyle, picCanvas.DrawWidth, cul&)
      hpenold& = SelectObject(picCanvas.hdc, hpen&)
      ReDim zp(1 To PCount) As POINTAPI
      For I = 1 To PCount
         zp(I).kX = xsto(I): zp(I).kY = ysto(I)
      Next I
      xy = Polyline(picCanvas.hdc, zp(1), PCount)
      zres = SelectObject(picCanvas.hdc, hpenold&)
      zres = DeleteObject(hpen&)
      picCanvas.PSet (zp(1).kX, zp(1).kY), cul&   'to refresh picCanvas
   Else
      'NB DrawDoubleLine uses xsto(), ysto() as its primary line
      Select Case LineType
      Case 7, 8, 9
         If Button <= 1 Then
            DrawDoubleLine cul&, 1
         Else
            DrawShadedLine 1
         End If
      Case Else
         DrawDoubleLine cul&, 1
      End Select
   End If
   
   Erase zp, xaa, yaa
   
   PCount = svPCount
   ReDim xsto(10000), ysto(10000)
   For N = 1 To PCount
      xsto(N) = xstosav(N): ysto(N) = ystosav(N)
   Next N
   Erase xstosav, ystosav
End If

End Sub
Private Sub CirlipseConcentricShaded()
'++,--,+-,-+
If zrad = 0 Then zrad = 0.01 ': Exit Sub
If zrad > 5 Then picCanvas.DrawWidth = 2
zsn = 255 - rubcn
If zsn < 0 Then zsn = 0
If zsn > 255 Then zsn = 255
zculsteps = (cn - rubcn) / Abs(zrad)
For zr = 1 To zrad
   zsn = zsn + zculsteps
   If zsn > 255 Then zsn = 255: zculsteps = -zculsteps
   If zsn < 2 Then zsn = 2: zculsteps = Abs(zculsteps)
   jsn = Int(zsn)
   picCanvas.Circle (xs, ys), zr, RGB(jsn, jsn, jsn), , , zratio
Next zr
picCanvas.DrawWidth = 1
End Sub
Private Sub ConcentricShadedCone()
'xs, ys, zrad, xp, yp
zsn = 255 - rubcn
If zsn < 0 Then zsn = 0
If zsn > 255 Then zsn = 255
zculsteps = (cn - rubcn) / 720
For zang = 0 To 2 * pi# Step pi# / 360
   zsn = zsn + zculsteps
   If zsn > 255 Then zsn = 255: zculsteps = -zculsteps
   If zsn < 2 Then zsn = 2: zculsteps = Abs(zculsteps)
   jsn = Int(zsn)
   xi = xs + zrad * Cos(zang)
   yi = ys + zrad * Sin(zang)
   picCanvas.Line (xi, yi)-(xp, yp), RGB(jsn, jsn, jsn)
Next zang
End Sub
Public Sub DrawTPiece(cul&, ByVal zxs, ByVal zys, ByVal zxe, ByVal zye)
zTPLen = 10  'Length of side piece(s)
If TPieceType = 0 Then
   'Clear/draw main line
   picCanvas.Line (zxs, zys)-(zxe, zye), cul&
Else
   'Clear/draw TPiece to left  PLUS PIECE
   Findxeye zpspac, zxs, zys, zxe, zye, xdd, ydd, xa, ya, xb, yb
   x1 = (zxs + zxe) / 2
   y1 = (zys + zye) / 2
   x2 = (xa + xb) / 2
   y2 = (ya + yb) / 2
   Findxeye zpspac / 2, x1, y1, x2, y2, xdd, ydd, xs1, ys1, xx, yx
   Findxeye zpspac / 2, x2, y2, x1, y1, xdd, ydd, xx, yy, xe1, ye1
   Findxeye zTPLen, xe1, ye1, xs1, ys1, xdd, ydd, xe2, ye2, xs2, ys2
   picCanvas.Line (zxs, zys)-(xs1, ys1), cul&
   picCanvas.Line (xs1, ys1)-(xs2, ys2), cul&
   picCanvas.Line (zxe, zye)-(xe1, ye1), cul&
   picCanvas.Line (xe1, ye1)-(xe2, ye2), cul&
End If
        
'Clear/draw TPiece to right ALWAYS
Findxeye zpspac, zxs, zys, zxe, zye, xdd, ydd, xa, ya, xb, yb
x1 = (zxs + zxe) / 2
y1 = (zys + zye) / 2
x2 = (xa + xb) / 2
y2 = (ya + yb) / 2
Findxeye zpspac / 2, x1, y1, x2, y2, xdd, ydd, xx, yy, xa1, ya1
Findxeye zpspac / 2, x2, y2, x1, y1, xdd, ydd, xb1, yb1, xx, yy
Findxeye zTPLen, xa1, ya1, xb1, yb1, xdd, ydd, xa2, ya2, xb2, yb2
picCanvas.Line (xa, ya)-(xa1, ya1), cul&
picCanvas.Line (xa1, ya1)-(xa2, ya2), cul&
picCanvas.Line (xb1, yb1)-(xb, yb), cul&
picCanvas.Line (xb1, yb1)-(xb2, yb2), cul&
End Sub

'=====================================================================

Private Sub StartBrush(X, Y)
   picCanvas.DrawMode = 13
   picCanvas.DrawStyle = vbSolid
   Select Case BrushType
   Case 0: picCanvas.DrawWidth = 1
   Case 1: picCanvas.DrawWidth = 2
   Case 2: picCanvas.DrawWidth = 4
   Case 3: picCanvas.DrawWidth = 8
   Case Else: picCanvas.DrawWidth = 1
   End Select
   
   Select Case BrushType
   Case 0, 1, 2, 3   '* Brush
      xp = X: yp = Y
      picCanvas.PSet (X, Y), DrawCul&
   Case 4, 5, 6   '\ Brush (2,4,6 pixels)
     'Global xp,yp
      xp = X: yp = Y
      x1 = X - (BrushType - 3): y1 = Y - (BrushType - 3)
      x2 = X + (BrushType - 3): y2 = Y + (BrushType - 3)
      
      picCanvas.Line (x1, y1)-(x2, y2), DrawCul&
   Case 7, 8, 9  '/   Brush
      'Global xp,yp
      xp = X: yp = Y
      'x1 = X - (BrushType - 7): y1 = Y + (BrushType - 7)
      'x2 = X + (BrushType - 7): y2 = Y - (BrushType - 7)
      x1 = X - (BrushType - 6): y1 = Y + (BrushType - 6)
      x2 = X + (BrushType - 6): y2 = Y - (BrushType - 6)
      picCanvas.Line (x1, y1)-(x2, y2), DrawCul&
   Case 10, 11, 12 'Stalk & leaves, drooping twigs & side upshoots   Brush
      xp = X: yp = Y
      picCanvas.PSet (X, Y), DrawCul&
   Case 13  'Grass  Brush
      xp = X: yp = Y
      picCanvas.PSet (X, Y), DrawCul&
   Case 14  'Blocks  Brush
      xp = X: yp = Y
      picCanvas.PSet (X, Y), DrawCul&
   End Select
End Sub

Private Sub DrawBrush(X, Y)
   Select Case BrushType
   Case 0, 1, 2, 3   '*
      picCanvas.Line (xp, yp)-(X, Y), DrawCul&
      xp = X: yp = Y
   Case 4, 5, 6 '\
      'Global xp,yp
      zdx = X - xp: zdy = Y - yp
      If Sgn(zdx) <> Sgn(zdy) Then picCanvas.DrawWidth = 2
      zlines = Abs(zdx)
      If Abs(zdx) < Abs(zdy) Then zlines = Abs(zdy)
      If zlines >= 0 Then
         zstepx = zdx / (zlines + 1)
         zstepy = zdy / (zlines + 1)
         xi = xp: yi = yp
         For zn = 0 To zlines + 1
            x1 = xi - (BrushType - 3): y1 = yi - (BrushType - 3)
            x2 = xi + (BrushType - 3): y2 = yi + (BrushType - 3)
            picCanvas.Line (x1, y1)-(x2, y2), DrawCul&
            xi = xi + zstepx
            yi = yi + zstepy
         Next zn
      End If
      picCanvas.DrawWidth = 1
      xp = X: yp = Y
   Case 7, 8, 9   'Brush /
      'Global xp,yp
      zdx = X - xp: zdy = Y - yp
      If Sgn(zdx) = Sgn(zdy) Then picCanvas.DrawWidth = 2
      zlines = Abs(zdx)
      If Abs(zdx) < Abs(zdy) Then zlines = Abs(zdy)
      If zlines >= 0 Then
         zstepx = zdx / (zlines + 1)
         zstepy = zdy / (zlines + 1)
         xi = xp: yi = yp
         For zn = 0 To zlines + 1
            'x1 = xi - (BrushType - 7): y1 = yi + (BrushType - 7)
            'x2 = xi + (BrushType - 7): y2 = yi - (BrushType - 7)
            x1 = xi - (BrushType - 7): y1 = yi + (BrushType - 6)
            x2 = xi + (BrushType - 7): y2 = yi - (BrushType - 6)
            
            picCanvas.Line (x1, y1)-(x2, y2), DrawCul&
            
            xi = xi + zstepx
            yi = yi + zstepy
         Next zn
      End If
      picCanvas.DrawWidth = 1
      xp = X: yp = Y
   
   Case 10  'Brush
      'Stalk & drooping side twigs
      picCanvas.Line -(X, Y), DrawCul&
         zyscale = Rnd * 4
         zxscale = Rnd * 4 - 2
         picCanvas.Line (X, Y)-(X + 1 * zxscale, Y - 1 * zyscale), DrawCul&
         picCanvas.Line -Step(2 * zxscale, -1 * zyscale), DrawCul&
         picCanvas.Line -Step(2 * zxscale, 1 * zyscale), DrawCul&
         picCanvas.Line -Step(2 * zxscale, 2 * zyscale), DrawCul&
         picCanvas.PSet (X, Y), DrawCul&
   Case 11  'Brush
     'Stalk & side upward shoots
      picCanvas.Line -(X, Y), DrawCul&
         zyscale = Rnd * 6
         zxscale = Rnd * 6 - 3
         picCanvas.Line (X, Y)-(X + 1 * zxscale, Y - 1 * zyscale), DrawCul&
         picCanvas.Line -Step(1 * zxscale, -2 * zyscale), DrawCul&
         picCanvas.Line -Step(0 * zxscale, -1 * zyscale), DrawCul&
         picCanvas.PSet (X, Y), DrawCul&
   Case 12  'Brush
      'Stalk & leaves
      picCanvas.Line -(X, Y), DrawCul&
      If Rnd < 0.6 Then   'reduces frequency of leaves
         zscale = Rnd * 4 - 2
         picCanvas.Line (X, Y)-(X + 3 * zscale, Y - 1), DrawCul&
         picCanvas.Line -Step(1 * zscale, 2 * zscale), DrawCul&
         picCanvas.Line -Step(1 * zscale, 2 * zscale), DrawCul&
         picCanvas.Line -Step(-2 * zscale, -1 * zscale), DrawCul&
         picCanvas.Line -Step(-1 * zscale, -1 * zscale), DrawCul&
         picCanvas.Line -Step(-1 * zscale, -1 * zscale), DrawCul&
         picCanvas.Line -Step(-1 * zscale, -2 * zscale), DrawCul&
         picCanvas.PSet (X, Y), DrawCul&
      End If
   
   Case 13  'Brush Grass
      posdir = -1  'Up
      If X < xp Then posdir = 1  'Down
      If Rnd < 0.6 Then
         ystalklen = Rnd * 10 + 1
         xstalklen = Rnd * 10 + 1 - 5
         xs = X + xstalklen
         ys = Y + posdir * ystalklen
         
         picCanvas.Line (X, Y)-(xs, ys), DrawCul&
         
      End If
      xp = X: yp = Y
   
   Case 14  'Brush Blocks
      
      xp = X: yp = Y
      If Rnd < 0.6 Then
         yblocklen = Rnd * 20 + 1
         xblocklen = Rnd * 20 + 1 - 10
         xs = X + xblocklen: ys = Y - yblocklen
         
         picCanvas.Line (X, Y)-(xs, ys), DrawCul&, B
         
      End If
      xp = X: yp = Y
   End Select
End Sub

Private Sub StartAirBrush(X, Y)
   picCanvas.DrawMode = 13
   Randomize
   Select Case AirBrushType
   Case 0: zradmax = 2  'AirBrush Small spray
      For I = 1 To 4
         xi = X + 1 - zradmax * Rnd + 1: yi = Y + 1 - zradmax * Rnd + 1
         picCanvas.PSet (xi, yi), DrawCul&
      Next I
   Case 1: zradmax = 8  'AirBrush Medium spray
      For I = 1 To 8
         xi = X + 2 - zradmax * Rnd + 1: yi = Y + 2 - zradmax * Rnd + 1
         picCanvas.PSet (xi, yi), DrawCul&
      Next I
   Case 2: zradmax = 16 'AirBrush Large spray
      For I = 1 To 16
         xi = X + 4 - zradmax * Rnd + 1: yi = Y + 4 - zradmax * Rnd + 1
         picCanvas.PSet (xi, yi), DrawCul&
      Next I
   Case 3, 4, 5   'AirBrush Grass
      DrawAirBrushGrass X, Y
   Case 6   'AirBrush Stalk & flowers
      DrawAirBrushStalks X, Y
   Case 7, 8, 9, 10 'AirBrush Bricks, Tiles & Fences
      mx = Val(txtmx.Text)
      my = Val(txtmy.Text)
      If mx = 0 Then mx = 1
      If my = 0 Then my = 1
      DrawAirBrushBricks X, Y
      picCanvas.DrawWidth = 1
   End Select
End Sub

Private Sub DrawAirBrush(X, Y)
Select Case AirBrushType
Case 0: zradmax = 2 'AirBrush dotted
   For I = 1 To 4
      xi = X + 1 - zradmax * Rnd + 1: yi = Y + 1 - zradmax * Rnd + 1
      picCanvas.PSet (xi, yi), DrawCul&
   Next I
Case 1: zradmax = 8  'AirBrush dotted
   For I = 1 To 8
      xi = X + 2 - zradmax * Rnd + 1: yi = Y + 2 - zradmax * Rnd + 1
      picCanvas.PSet (xi, yi), DrawCul&
   Next I
Case 2: zradmax = 16 'AirBrush dotted
   For I = 1 To 16
      xi = X + 4 - zradmax * Rnd + 1: yi = Y + 4 - zradmax * Rnd + 1
      picCanvas.PSet (xi, yi), DrawCul&
   Next I
Case 3, 4, 5 'AirBrush Grass
   DrawAirBrushGrass X, Y
Case 6   'AirBrush Stalk & Leaves
    DrawAirBrushStalks X, Y
Case 7, 8, 9, 10 'AirBrush Bricks, Tiles & Fences
      DrawAirBrushBricks X, Y
End Select
End Sub
Private Sub DrawAirBrushStalks(X, Y)
zlen = 20
zwidth = zlen
xs = X: ys = Y
For I = 1 To 16
   xp = X + zwidth * (Rnd - 0.5): yp = Y - zlen * Rnd + 1
   picCanvas.Line (xs, ys)-(xp, yp), DrawCul&
   picCanvas.Circle (xp, yp), 1, DrawCul&
Next I
End Sub
Private Sub DrawAirBrushGrass(X, Y)
zlen = 4 * (AirBrushType - 1) '4*(1,2,4)
zwidth = zlen
xs = X: ys = Y
For I = 1 To 16
   xp = X + zwidth * (Rnd - 0.5)
   yp = Y - zlen * Rnd + 1
   picCanvas.Line (xs, ys)-(xp, yp), DrawCul&
Next I
End Sub
Private Sub DrawAirBrushBricks(X, Y)
Select Case AirBrushType
Case 7   'Bricks
xs = X: ys = Y
If 2 * my + 4 = 0 Then my = -3
xs = xs - (xs Mod mx): ys = ys - (ys Mod (2 * my + 4))
         
'=|
picCanvas.Line (xs, ys)-(xs + mx / 2, ys), DrawCul&
picCanvas.Line (xs + mx / 2, ys)-(xs + mx / 2, ys + my), DrawCul&
picCanvas.Line (xs + mx / 2, ys + my)-(xs, ys + my), DrawCul&
'|=
xp = xs + mx / 2 + 2
picCanvas.Line (xp, ys)-(xp + mx / 2, ys), DrawCul&
picCanvas.Line (xp, ys)-(xp, ys + my), DrawCul&
picCanvas.Line (xp, ys + my)-(xp + mx / 2, ys + my), DrawCul&
      
yp = ys + my + 2
picCanvas.Line (xs + 1, yp)-(xs + mx - 1, yp + my), DrawCul&, B

Case 8   'Slanted tiles
picCanvas.DrawWidth = 2 'To allow Fill to clear them
xs = X: ys = Y
xs = xs - (xs Mod (2 * mx)): ys = ys - (ys Mod (2 * my))
picCanvas.Line (xs, ys)-(xs + 2 * mx, ys - 2 * my), DrawCul&
picCanvas.Line (xs + 2 * mx, ys - 2 * my)-(xs + 3 * mx, ys - 1 * my), DrawCul&
picCanvas.Line (xs + 3 * mx, ys - 1 * my)-(xs + 1 * mx, ys + 1 * my), DrawCul&
picCanvas.Line (xs + 1 * mx, ys + 1 * my)-(xs, ys), DrawCul&
picCanvas.DrawWidth = 1

Case 9   'Horizontal tiles
xs = X: ys = Y
xs = xs - (xs Mod (2 * mx)): ys = ys - (ys Mod (2 * my))
picCanvas.Line (xs, ys)-(xs + 2 * mx, ys + 1 * my), DrawCul&, B
picCanvas.Line (xs + mx, ys + my)-(xs + 2 * mx + mx, ys + 1 * my + my), DrawCul&, B

Case 10   'Fence
xs = X: ys = Y
If mx = -2 Then mx = -3
xs = xs - (xs Mod (2 * mx + 4)): ys = ys - (ys Mod (2 * my))
picCanvas.Line (xs, ys)-(xs + 2 * mx, ys + 2 * my), DrawCul&, B
picCanvas.Line (xs + 2 * mx + 4, ys)-(xs + 2 * mx + 4 + 2 * mx, ys + 2 * my), DrawCul&, B
End Select

End Sub

Private Sub SetCanvasLine()
picCanvas.DrawStyle = vbSolid
Select Case LineType
Case 0: picCanvas.DrawWidth = 1: zpspac = 0
Case 1: picCanvas.DrawWidth = 2: zpspac = 0
Case 2: picCanvas.DrawWidth = 4: zpspac = 0
Case 3: picCanvas.DrawWidth = 1: zpspac = 2
Case 4: picCanvas.DrawWidth = 1: zpspac = 4
Case 5: picCanvas.DrawWidth = 1: zpspac = 6
Case 6: picCanvas.DrawWidth = 1: zspac = 0:
    picCanvas.DrawStyle = vbDot: zpspac = 0
    picCanvas.FontTransparent = True
Case 7: picCanvas.DrawWidth = 2: zpspac = 6
Case 8: picCanvas.DrawWidth = 2: zpspac = 12
Case 9: picCanvas.DrawWidth = 2: zpspac = 18
Case Else: picCanvas.DrawWidth = 1: zpspac = 0
End Select
End Sub

Private Sub SetCanvasRectangle()
Select Case RectangleType
Case 0: picCanvas.DrawWidth = 1
Case 1: picCanvas.DrawWidth = 2
Case 2: picCanvas.DrawWidth = 4
Case 3: picCanvas.DrawWidth = 1: zpspac = 2
Case 4: picCanvas.DrawWidth = 1: zpspac = 4
Case 5: picCanvas.DrawWidth = 1: zpspac = 6
Case 6: picCanvas.DrawWidth = 1: zpspac = 0
   picCanvas.DrawStyle = vbDot
   picCanvas.FontTransparent = True
Case Else: picCanvas.DrawWidth = 1
End Select
End Sub

Private Sub StartLine(X, Y)
SetCanvasLine
LCount = LCount + 1
If LCount = 1 Then 'Starts line picCanvas.Mousemove will position
   picCanvas.DrawMode = 7    'To allow rubber-banding in Mousemove
   Select Case LineType
   Case 0, 1, 2, 6   'ALine
      xs = X: ys = Y
      xp = X: yp = Y
      picCanvas.Line (xs, ys)-(X, Y), PenCul&
   Case 3, 4, 5, 7, 8, 9 'ALine
      PCount = 2 'For DrawDoubleLine
      xsto(1) = X: ysto(1) = Y
      xsto(2) = X: ysto(2) = Y
      DrawDoubleLine PenCul&, 1
      'FINDS AND DRAWS MAIN & PARALLEL LINES zpspac AWAY FROM MAIN LINE
      'Needs (Cul&,StartCount)
      '& Global PCount, xsto(PCount),ysto(PCount) & zpspac para-spacing
      'Cul& = PenCul& (mode 7) or Colors(cn) (mode 13)
   End Select
ElseIf LCount = 2 Then
      'Move whole line
Else  'LCount=2  Redraw Line, end of line
   CANVASRESTORE
   picCanvas.DrawMode = 13
   Select Case LineType
   Case 0, 1, 2, 6   'ALine
      picCanvas.Line (xs, ys)-(xp, yp), DrawCul&
   Case 3, 4, 5   'ALine
      DrawDoubleLine DrawCul&, 1
   Case 7, 8, 9 'ALine
      DrawShadedLine 1
   End Select
   picCanvas.DrawWidth = 1
   
   If LineType <> 6 Then
      picCanvas.FontTransparent = True
      picCanvas.DrawStyle = vbSolid
   End If
   
   DrawingMode = False  'Allows new start
End If
End Sub

Private Sub DrawLine(X, Y)
Select Case LineType
Case 0, 1, 2, 6   'Line widths 1,2,4 & dotted
            
   If LCount = 1 Then
      picCanvas.Line (xs, ys)-(xp, yp), PenCul&
      xp = X: yp = Y
      picCanvas.Line (xs, ys)-(xp, yp), PenCul&
      xprev = X: yprev = Y
   ElseIf LCount = 2 Then  'Move whole Line
      picCanvas.Line (xs, ys)-(xp, yp), PenCul&
      xs = xs + (X - xprev)
      ys = ys + (Y - yprev)
      xp = xp + (X - xprev)
      yp = yp + (Y - yprev)
      xprev = X: yprev = Y
      picCanvas.Line (xs, ys)-(xp, yp), PenCul&
   End If
   
Case 3, 4, 5, 7, 8, 9 'Double Lines spacings 2,4,6
      
   If LCount = 1 Then
      DrawDoubleLine PenCul&, 1
      xsto(PCount) = X: ysto(PCount) = Y
      DrawDoubleLine PenCul&, 1
      xprev = X: yprev = Y
   ElseIf LCount = 2 Then  'Move whole Double line
      DrawDoubleLine PenCul&, 1
      For I = 1 To PCount
         xsto(I) = xsto(I) + (X - xprev)
         ysto(I) = ysto(I) + (Y - yprev)
      Next I
      xprev = X: yprev = Y
      DrawDoubleLine PenCul&, 1
   End If
   
End Select
End Sub

Private Sub StartPolyLine(Button, X, Y)
SetCanvasLine
If (LineType = 10 Or LineType = 11) And PCount = 2 Then Button = 2

LCount = LCount + 1
If (Button = 1 Or LCount = 1) And RedrawPolCoTu = 0 Then 'LCul or  RCul to start fresh line
   picCanvas.DrawMode = 7
   Select Case LineType
   Case 0, 1, 2, 6, 10, 11, 12 'APolyline
      xs = X: ys = Y
      xp = X: yp = Y
      PCount = PCount + 1
      xsto(PCount) = X: ysto(PCount) = Y
      picCanvas.Line (xs, ys)-(xp, yp), PenCul&
   Case 3, 4, 5, 7, 8, 9 'Double APolyline
      PCount = PCount + 1
      xsto(PCount) = X: ysto(PCount) = Y
      If PCount = 1 Then
         PCount = PCount + 1
         xsto(PCount) = X: ysto(PCount) = Y
      End If
      DrawDoubleLine PenCul&, PCount - 1
   End Select

ElseIf Button = 2 And RedrawPolCoTu = 0 Then  'Move whole PolyLine
   RedrawPolCoTu = 1
   CANVASRESTORE

   Select Case LineType
   Case 0, 1, 2, 6  'APolyline
      PCount = PCount + 1
      xsto(PCount) = xp: ysto(PCount) = yp
   Case 3, 4, 5, 7, 8, 9 'Double APolyline
   
   Case 10, 11 'Parallelogram & Frustrum calculate 4th point
      PCount = PCount + 1
      xsto(PCount) = xp: ysto(PCount) = yp
      If LineType = 10 Then   'Parallelogram
         PCount = 5
         xsto(4) = xsto(1) - (xsto(2) - xsto(3))
         ysto(4) = ysto(1) + (ysto(3) - ysto(2))
      Else  'Frustrum
         PCount = 5
         x1 = xsto(1): y1 = ysto(1)
         x2 = xsto(2): y2 = ysto(2)
         x3 = xsto(3): y3 = ysto(3)
         'Find angle of 1st line to horizontal
         za = zAtn2(y2 - y1, x2 - x1)
         'Translate to make x1,y1 the origin
         x2 = x2 - x1: y2 = y2 - y1
         x3 = x3 - x1: y3 = y3 - y1
         'Rotate za about 1st pt
         xr2 = x2 * Cos(za) + y2 * Sin(za)
         'yr2 = -x2 * Sin(za) + y2 * Cos(za)
         xr3 = x3 * Cos(za) + y3 * Sin(za)
         yr3 = -x3 * Sin(za) + y3 * Cos(za)
         'Find 4th point
         xr4 = (xr2 - xr3)
         yr4 = yr3
         'Rotate back
         x2 = xr2 * Cos(-za) + yr2 * Sin(-za)
         'y2 = -xr2 * Sin(-za) + yr2 * Cos(-za)
         x3 = xr3 * Cos(-za) + yr3 * Sin(-za)
         y3 = -xr3 * Sin(-za) + yr3 * Cos(-za)
         x4 = xr4 * Cos(-za) + yr4 * Sin(-za)
         y4 = -xr4 * Sin(-za) + yr4 * Cos(-za)
         'Translate back
         xsto(2) = x2 + x1: ysto(2) = y2 + y1
         xsto(3) = x3 + x1: ysto(3) = y3 + y1
         xsto(4) = x4 + x1
         ysto(4) = y4 + y1
      End If
      'Join back to 1st point
      xsto(5) = xsto(1)
      ysto(5) = ysto(1)
   Case 12  'Auto-join last point
      PCount = PCount + 1
      xsto(PCount) = xp: ysto(PCount) = yp
      PCount = PCount + 1
      xsto(PCount) = xsto(1)
      ysto(PCount) = ysto(1)
   End Select

Else 'Fix PolyLine
   RedrawPolCoTu = 0
   CANVASRESTORE
   picCanvas.DrawMode = 13
   Select Case LineType
   Case 0, 1, 2, 6  'APolyline
      For I = 1 To PCount - 1
         picCanvas.Line (xsto(I), ysto(I))-(xsto(I + 1), ysto(I + 1)), DrawCul&
      Next I
   Case 3, 4, 5   'Double APolyline
      DrawDoubleLine DrawCul&, 1
   Case 7, 8, 9   'Shaded APolyline
      DrawShadedLine 1
   Case 10, 11, 12 'Parallelogram & Frustrum & Auto-join
      For I = 1 To PCount - 1
         picCanvas.Line (xsto(I), ysto(I))-(xsto(I + 1), ysto(I + 1)), DrawCul&
      Next I
   End Select
   DrawingMode = False  'Allows new start
   
End If
End Sub

Private Sub DrawPolyLine(Button, X, Y)

If RedrawPolCoTu > 0 Then  'Move whole PolyLine
   If RedrawPolCoTu > 1 Then   'Skip once after CANVASRESTORE just done
      Select Case LineType
      Case 0, 1, 2, 6, 10, 11, 12 'APolyline
         For I = 1 To PCount - 1
            picCanvas.Line (xsto(I), ysto(I))-(xsto(I + 1), ysto(I + 1)), PenCul&
         Next I
      Case 3, 4, 5, 7, 8, 9 'Double APolyline
         DrawDoubleLine PenCul&, 1
      End Select
   
      For I = 1 To PCount '+ 1
         xsto(I) = xsto(I) + (X - xprev)
         ysto(I) = ysto(I) + (Y - yprev)
      Next I
   End If
   
   Select Case LineType
   Case 0, 1, 2, 6, 10, 11, 12 'APolyline
      For I = 1 To PCount - 1
         picCanvas.Line (xsto(I), ysto(I))-(xsto(I + 1), ysto(I + 1)), PenCul&
      Next I
   Case 3, 4, 5, 7, 8, 9 'Double APolyline
      DrawDoubleLine PenCul&, 1
   End Select
   
   RedrawPolCoTu = 2 'RedrawPolCoTu + 1  'To allow full redraw
   xprev = X: yprev = Y

Else
   
   Select Case LineType
   Case 0, 1, 2, 6, 10, 11, 12
      picCanvas.Line (xs, ys)-(xp, yp), PenCul&
      xp = X: yp = Y
      picCanvas.Line (xs, ys)-(xp, yp), PenCul&
   Case 3, 4, 5, 7, 8, 9
      DrawDoubleLine PenCul&, PCount - 1
      xsto(PCount) = X: ysto(PCount) = Y
      DrawDoubleLine PenCul&, PCount - 1
   End Select
   xprev = X: yprev = Y

End If
End Sub

Private Sub StartSpline(Button, X, Y)
SetCanvasLine
LCount = LCount + 1
If (Button = 1 Or LCount = 1) And RedrawPolCoTu = 0 Then 'LCul or  RCul to start fresh line
   picCanvas.DrawMode = 7
   PCount = PCount + 1  'Start fresh Spline until RC
   xsto(PCount) = X: ysto(PCount) = Y
   If PCount = 1 Then
      PCount = PCount + 1
      xsto(PCount) = X: ysto(PCount) = Y
   End If
   PCount = PCount - 1
   DrawSpline PenCul&, Button
   PCount = PCount + 1
   'Refill since DrawSpline loses last point
   xsto(PCount) = X: ysto(PCount) = Y
   DrawSpline PenCul&, Button

'------
ElseIf Button = 2 And RedrawPolCoTu = 0 Then  'Move whole Spline outline
   RedrawPolCoTu = 1
   CANVASRESTORE

Else  'RC  Redraw Spline
   RedrawPolCoTu = 0
   CANVASRESTORE
   picCanvas.DrawMode = 13
   Button = 2  'Forces shading
   DrawSpline DrawCul&, Button
   DrawingMode = False  'Allows new start
   Button = 0
End If
End Sub

Private Sub ADrawSpline(Button, X, Y)
If RedrawPolCoTu > 0 Then  'Move whole Spline outline
   If RedrawPolCoTu > 1 Then   'Skip once after CANVASRESTORE just done
      DrawSpline PenCul&, Button
      For I = 1 To PCount '+ 1
         xsto(I) = xsto(I) + (X - xprev)
         ysto(I) = ysto(I) + (Y - yprev)
      Next I
   End If
   
   DrawSpline PenCul&, Button
   
   RedrawPolCoTu = 2 'RedrawPolCoTu + 1  'To allow full redraw
   xprev = X: yprev = Y
Else
   DrawSpline PenCul&, Button
   xsto(PCount) = X: ysto(PCount) = Y
   DrawSpline PenCul&, Button
   xprev = X: yprev = Y
End If
End Sub

Private Sub StartRectangle(X, Y)
SetCanvasRectangle
LCount = LCount + 1
If LCount = 1 Then
   picCanvas.DrawMode = 7
   xs = X: ys = Y
   xp = X: yp = Y
   Select Case RectangleType
   Case 0, 1, 2, 6, 7, 8, 9   'Rectangle
      picCanvas.Line (xs, ys)-(X, Y), PenCul&, B
   Case 3, 4, 5   ''Rectangle
      picCanvas.Line (xs, ys)-(X, Y), PenCul&, B
      picCanvas.Line (xs + zpspac, ys + zpspac)-(X - zpspac, Y - zpspac), PenCul&, B
   End Select
ElseIf LCount = 2 Then
   'Now move whole piece
   
Else  'LCount=2 Redraw rectangle
   CANVASRESTORE
   picCanvas.DrawMode = 13
   'Select Case RectangleType
   Select Case RectangleType
   Case 0, 1, 2, 6   'Rectangle
      picCanvas.Line (xs, ys)-(xp, yp), DrawCul&, B
   Case 3, 4, 5   'Double Rectangle
      picCanvas.Line (xs, ys)-(xp, yp), DrawCul&, B
      picCanvas.Line (xs + zpspac, ys + zpspac)-(xp - zpspac, yp - zpspac), DrawCul&, B
   Case 7   'Horz shading Colors cn to rubcn 'Rectangle
      RectangleHorzShading
   Case 8   'Vert shading DrawCul& to Colours rubcn 'Rectangle
      RectangleVertShading
   Case 9   'Concentric shading  'Rectangle
      RectangleConcentricShading
   End Select
   DrawingMode = False  'Allows new start
End If
End Sub

Private Sub DrawRectangle(X, Y)
Select Case RectangleType
Case 0, 1, 2, 6, 7, 8, 9   'Rectangle
   If LCount = 1 Then
      picCanvas.Line (xs, ys)-(xp, yp), PenCul&, B
      xp = X: yp = Y
      picCanvas.Line (xs, ys)-(xp, yp), PenCul&, B
      xprev = X: yprev = Y

   ElseIf LCount = 2 Then  'Move whole Rectangle
      picCanvas.Line (xs, ys)-(xp, yp), PenCul&, B
      xs = xs + (X - xprev)
      ys = ys + (Y - yprev)
      xp = xp + (X - xprev)
      yp = yp + (Y - yprev)
      xprev = X: yprev = Y
      picCanvas.Line (xs, ys)-(xp, yp), PenCul&, B
   End If

Case 3, 4, 5   'Double Rectangle
   If LCount = 1 Then
      picCanvas.Line (xs, ys)-(xp, yp), PenCul&, B
      picCanvas.Line (xs + zpspac, ys + zpspac)-(xp - zpspac, yp - zpspac), PenCul&, B
      xp = X: yp = Y
      picCanvas.Line (xs, ys)-(xp, yp), PenCul&, B
      picCanvas.Line (xs + zpspac, ys + zpspac)-(xp - zpspac, yp - zpspac), PenCul&, B
      xprev = X: yprev = Y
   ElseIf LCount = 2 Then  'Move whole Rectangle
      picCanvas.Line (xs, ys)-(xp, yp), PenCul&, B
      picCanvas.Line (xs + zpspac, ys + zpspac)-(xp - zpspac, yp - zpspac), PenCul&, B
      xs = xs + (X - xprev)
      ys = ys + (Y - yprev)
      xp = xp + (X - xprev)
      yp = yp + (Y - yprev)
      xprev = X: yprev = Y
      picCanvas.Line (xs, ys)-(xp, yp), PenCul&, B
      picCanvas.Line (xs + zpspac, ys + zpspac)-(xp - zpspac, yp - zpspac), PenCul&, B
   End If

End Select
End Sub

Private Sub RectangleHorzShading()
icn = 255 - cn
If icn < 0 Then icn = 0
If icn > 255 Then icn = 255
picCanvas.Line (xs, ys)-(xp, yp), RGB(icn, icn, icn), B
If yp <> ys Then
   zsn = 255 - rubcn
   If zsn < 0 Then zsn = 0
   If zsn > 255 Then zsn = 255
   zdy = Abs((yp - (ys + 1)))
   If zdy = 0 Then zdy = 0.02
   zculsteps = (cn - rubcn) / zdy
   For IY = ys + 1 To yp - 1 Step Sgn(yp - ys)
       jsn = Int(zsn)
       picCanvas.Line (xs + 1, IY)-(xp, IY), RGB(jsn, jsn, jsn)
       zsn = zsn + zculsteps
       If zsn > 255 Then zsn = 255: zculsteps = -zculsteps
       If zsn < 1 Then zsn = 1: zculsteps = Abs(zculsteps)
   Next IY
End If
End Sub
Private Sub RectangleVertShading()
icn = 255 - cn
If icn < 0 Then icn = 0
If icn > 255 Then icn = 255
picCanvas.Line (xs, ys)-(xp, yp), RGB(icn, icn, icn), B
If xp <> xs Then
   zsn = 255 - rubcn
   If zsn < 0 Then zsn = 0
   If zsn > 255 Then zsn = 255
   zdx = Abs((xp - (xs + 1)))
   If zdx = 0 Then zdx = 0.02
   zculsteps = (cn - rubcn) / Abs((xp - (xs + 1)))
   For IX = xs + 1 To xp - 1 Step Sgn(xp - xs)
      jsn = Int(zsn)
      picCanvas.Line (IX, ys + 1)-(IX, yp), RGB(jsn, jsn, jsn)
      zsn = zsn + zculsteps
      If zsn > 255 Then zsn = 255: zculsteps = -zculsteps
      If zsn < 1 Then zsn = 1: zculsteps = Abs(zculsteps)
   Next IX
End If
End Sub
Private Sub RectangleConcentricShading()
icn = 255 - cn
If icn < 0 Then icn = 0
If icn > 255 Then icn = 255
picCanvas.Line (xs, ys)-(xp, yp), RGB(icn, icn, icn), B
If xp <> xs Then
   iy1 = ys + 1
   iy2 = yp - 1
   ix1 = xs + 1
   ix2 = xp - 1
   nsq = Abs(iy2 - iy1) / 2
   If nsq > Abs(ix2 - ix1) / 2 Then nsq = Abs(ix2 - ix1) / 2
      zsn = 255 - rubcn
      If zsn < 0 Then zsn = 0
      If zsn > 255 Then zsn = 255
      If nsq <> 0 Then
         zculsteps = (cn - rubcn) / nsq
      Else
         nsq = 1
         zculsteps = (cn - rubcn) / nsq
      End If
      For N = 1 To nsq
         jsn = Int(zsn)
         picCanvas.Line (ix1, iy1)-(ix2, iy2), RGB(jsn, jsn, jsn), B
         zsn = zsn + zculsteps
         If zsn > 255 Then zsn = 255: zculsteps = -zculsteps
         If zsn < 1 Then zsn = 1: zculsteps = Abs(zculsteps)
         iy1 = iy1 + 1: iy2 = iy2 - 1
         ix1 = ix1 + 1: ix2 = ix2 - 1
      Next N
End If
End Sub


Private Sub StartCirlipse(X, Y)
SetCanvasRectangle
LCount = LCount + 1
If LCount = 1 Then
   picCanvas.DrawMode = 7
   xs = X: ys = Y
   EvalZradZratio X, Y
   Select Case RectangleType
   Case 0, 1, 2, 6, 7, 8, 9   'Cirlipse
      picCanvas.Circle (xs, ys), zrad, PenCul&, , , zratio
   Case 3, 4, 5   '2 Cirlipses
      picCanvas.Circle (xs, ys), zrad, PenCul&, , , zratio
      EvalZrad2Zratio2
      picCanvas.Circle (xs, ys), zrad2, PenCul&, , , zratio2
   End Select
ElseIf LCount = 2 Then
   'Now move whole piece
   
Else  'Redraw Cirlipse
   CANVASRESTORE
   picCanvas.DrawMode = 13
   picCanvas.PSet (xs, ys), DrawCul&   'Colors(cn)
   Select Case RectangleType
   Case 0, 1, 2, 6   '3 line thickness & dotted 'Cirlipse
      picCanvas.Circle (xs, ys), zrad, DrawCul&, , , zratio
   Case 3, 4, 5      '3 double-line spacings 'Cirlipse
      picCanvas.Circle (xs, ys), zrad, DrawCul&, , , zratio
      picCanvas.Circle (xs, ys), zrad2, DrawCul&, , , zratio2
      
   Case 7   'Cirlipse horz shading
      CirlipseHorzShading
   Case 8   'Cirlipse vert shading
      CirlipseVertShading
   Case 9   'Concentric shaded Cirlipse
      CirlipseConcentricShaded
   End Select
   DrawingMode = False  'Allows new start
End If
End Sub


Private Sub DrawCirlipse(X, Y)
Select Case RectangleType
Case 0, 1, 2, 6, 7, 8, 9   'Cirlipse
   If LCount = 1 Then
      picCanvas.Circle (xs, ys), zrad, PenCul&, , , zratio
      EvalZradZratio X, Y
      picCanvas.Circle (xs, ys), zrad, PenCul&, , , zratio
      xprev = X: yprev = Y
      DoEvents

   ElseIf LCount = 2 Then  'Move Cirlipse
      picCanvas.Circle (xs, ys), zrad, PenCul&, , , zratio
      'EvalZradZratio X, Y
      xs = xs + (X - xprev)
      ys = ys + (Y - yprev)
      xp = xp + (X - xprev)
      yp = yp + (Y - yprev)
      xprev = X: yprev = Y
      picCanvas.Circle (xs, ys), zrad, PenCul&, , , zratio
   End If

Case 3, 4, 5   '2 Cirlipses
   If LCount = 1 Then
      picCanvas.Circle (xs, ys), zrad, PenCul&, , , zratio
      picCanvas.Circle (xs, ys), zrad2, PenCul&, , , zratio2
      EvalZradZratio X, Y
      EvalZrad2Zratio2
      picCanvas.Circle (xs, ys), zrad, PenCul&, , , zratio
      picCanvas.Circle (xs, ys), zrad2, PenCul&, , , zratio2
      xprev = X: yprev = Y

   ElseIf LCount = 2 Then  'Move Cirlipses
      picCanvas.Circle (xs, ys), zrad, PenCul&, , , zratio
      picCanvas.Circle (xs, ys), zrad2, PenCul&, , , zratio2
      'EvalZradZratio X, Y
      'EvalZrad2Zratio2
      xs = xs + (X - xprev)
      ys = ys + (Y - yprev)
      xp = xp + (X - xprev)
      yp = yp + (Y - yprev)
      xprev = X: yprev = Y
      picCanvas.Circle (xs, ys), zrad, PenCul&, , , zratio
      picCanvas.Circle (xs, ys), zrad2, PenCul&, , , zratio2
   End If
   
End Select
End Sub

Private Sub CirlipseHorzShading()
If zrad = 0 Then zrad = 1
If zratio = 0 Then zratio = 0.001 ': Exit Sub
zyb = zrad * zratio   'horz ellipse zyb up - zxa right
zxa = zrad            'zrad to right  x=xs to xs+zrad
            
If zyb > zxa Then     'vert ellipse zyb up - zxa right
   zyb = zrad         'zrad up  y=ys to ys+zrad
   zxa = zyb / zratio
End If
            
y1 = ys - zyb: y2 = ys + zyb
zsn = 255 - rubcn
If zsn < 0 Then zsn = 0
If zsn > 255 Then zsn = 255
zculsteps = (cn - rubcn) / Abs(2 * zyb)
            
For IY = y1 To y2 Step Sgn(y2 - y1)
    zsn = zsn + zculsteps
    If zsn > 255 Then zsn = 255: zculsteps = -zculsteps
    If zsn < 1 Then zsn = 1: zculsteps = Abs(zculsteps)
    jsn = Int(zsn)
    xdx = (zxa / zyb) * Sqr(Abs(zyb ^ 2 - (IY - ys) ^ 2))
    picCanvas.Line (xs - xdx, IY)-(xs + xdx, IY), RGB(jsn, jsn, jsn)
Next IY
End Sub
Private Sub CirlipseVertShading()
If zrad = 0 Then zrad = 1 ': Exit Sub
If zratio = 0 Then zratio = 0.001 ': Exit Sub
zyb = zrad * zratio
zxa = zrad
            
If zyb > zxa Then
   zyb = zrad
   zxa = zyb / zratio
End If
            
x1 = xs - zxa: x2 = xs + zxa
zsn = rubcn
If zsn < 0 Then zsn = 0
If zsn > 255 Then zsn = 255
zculsteps = (cn - rubcn) / Abs(2 * zxa)
            
For IX = x1 To x2 Step Sgn(x2 - x1)
   zsn = zsn + zculsteps
   If zsn > 255 Then zsn = 255: zculsteps = -zculsteps
   If zsn < 1 Then zsn = 1: zculsteps = Abs(zculsteps)
   jsn = Int(zsn)
   ydy = (zyb / zxa) * Sqr(Abs(zxa ^ 2 - (IX - xs) ^ 2))
   picCanvas.Line (IX, ys - ydy)-(IX, ys + ydy), RGB(jsn, jsn, jsn)
Next IX

End Sub

Private Sub StartCone(X, Y)
SetCanvasRectangle
LCount = LCount + 1
If (Button = 1 Or LCount = 1) And RedrawPolCoTu = 0 Then 'LCul or  RCul to start fresh cone
   picCanvas.DrawMode = 7
   xs = X: ys = Y
   zrad = Abs(X - xs)
   zratio = 1
   Select Case RectangleType
   Case 0, 1, 2, 6, 7, 8, 9   'Cone
      picCanvas.Circle (xs, ys), zrad, PenCul&, , , zratio
   Case 3, 4, 5   'Cone
      picCanvas.Circle (xs, ys), zrad, PenCul&, , , zratio
      zrad2 = zrad - zpspac
      If zrad2 < 0 Then zrad2 = 0
      picCanvas.Circle (xs, ys), zrad2, PenCul&, , , zratio
   End Select
      
ElseIf LCount = 2 Then
   xp = X: yp = Y
   picCanvas.Line (xs, ys)-(X, Y), PenCul&

ElseIf LCount = 3 And RedrawPolCoTu = 0 Then  'Move whole Cone outline
   RedrawPolCoTu = 1
   CANVASRESTORE
   LCount = 3

Else
   'Redraw Cirlipse of cone
   CANVASRESTORE
   picCanvas.DrawMode = 13
   Select Case RectangleType
   Case 0, 1, 2, 6   '3 line thickness & dotted 'Cone
      picCanvas.Circle (xs, ys), zrad, DrawCul&, , , zratio
      picCanvas.PSet (xs, ys), DrawCul&
      EvalTangents xs, ys, zrad, xp, yp, x1, y1, x2, y2
      picCanvas.Line (x1, y1)-(xp, yp), DrawCul&
      picCanvas.Line (x2, y2)-(xp, yp), DrawCul&
   Case 3, 4, 5      '3 double-line spacings 'Cone
      picCanvas.Circle (xs, ys), zrad, DrawCul&, , , zratio
      picCanvas.PSet (xs, ys), DrawCul&
      EvalTangents xs, ys, zrad, xp, yp, x1, y1, x2, y2
      picCanvas.Line (x1, y1)-(xp, yp), DrawCul&
      picCanvas.Line (x2, y2)-(xp, yp), DrawCul&
         
      zrad2 = zrad - zpspac
      If zrad2 < 0 Then zrad2 = 0
      picCanvas.Circle (xs, ys), zrad2, DrawCul&, , , zratio
      EvalTangents xs, ys, zrad2, xp, yp, x1, y1, x2, y2
      picCanvas.Line (x1, y1)-(xp, yp), DrawCul&
      picCanvas.Line (x2, y2)-(xp, yp), DrawCul&
      
   Case 7   'Cone: along axis shading
      ConeAlongAxisShading
      picCanvas.DrawWidth = 1
      
   Case 8   'Cone: vert shading
      ConeVertShading
      picCanvas.DrawWidth = 1
      
   Case 9   'Cone: center shading
      ConeCenterShading
      picCanvas.DrawWidth = 1
   End Select
   DrawingMode = False
End If
End Sub

Private Sub DrawCone(X, Y)
If RedrawPolCoTu > 0 Then  'Move whole Cone outline
   If RedrawPolCoTu > 1 Then   'Skip once after CANVASRESTORE just done
      FixedConeTubeBaseOutLine
      picCanvas.Line (xs, ys)-(xp, yp), PenCul&
      xs = xs + (X - xprev)
      ys = ys + (Y - yprev)
      xp = xp + (X - xprev)
      yp = yp + (Y - yprev)
      xprev = X: yprev = Y
   End If
   
   FixedConeTubeBaseOutLine
   picCanvas.Line (xs, ys)-(xp, yp), PenCul&
   
   RedrawPolCoTu = 2 'RedrawPolCoTu + 1  'To allow full redraw
   xprev = X: yprev = Y
Else
   If LCount = 1 Then
      ConeTubeBaseOutLine X, Y
   ElseIf LCount = 2 Then  'Line from circle center along cone axis
      picCanvas.Line (xs, ys)-(xp, yp), PenCul&
      xp = X: yp = Y
      picCanvas.Line (xs, ys)-(xp, yp), PenCul&
   End If
   xprev = X: yprev = Y
End If

End Sub
Private Sub ConeTubeBaseOutLine(X, Y)
   Select Case RectangleType
   Case 0, 1, 2, 6, 7, 8, 9   'Cone
      picCanvas.Circle (xs, ys), zrad, PenCul&, , , zratio
      zrad = Abs(X - xs)
      zratio = 1
      picCanvas.Circle (xs, ys), zrad, PenCul&, , , zratio
      
   Case 3, 4, 5   'Cone
      picCanvas.Circle (xs, ys), zrad, PenCul&, , , zratio
      zrad2 = zrad - zpspac
      If zrad2 < 0 Then zrad2 = 0
      picCanvas.Circle (xs, ys), zrad2, PenCul&, , , zratio
      
      zrad = Abs(X - xs)
      zratio = 1
      picCanvas.Circle (xs, ys), zrad, PenCul&, , , zratio
      zrad2 = zrad - zpspac
      If zrad2 < 0 Then zrad2 = 0
      picCanvas.Circle (xs, ys), zrad2, PenCul&, , , zratio
      
   End Select
End Sub
   
Private Sub FixedConeTubeBaseOutLine()
   Select Case RectangleType
   Case 0, 1, 2, 6, 7, 8, 9   'Cone
      picCanvas.Circle (xs, ys), zrad, PenCul&, , , zratio
   Case 3, 4, 5   'Cone
      picCanvas.Circle (xs, ys), zrad, PenCul&, , , zratio
      picCanvas.Circle (xs, ys), zrad2, PenCul&, , , zratio
   End Select
End Sub


Private Sub ConeAlongAxisShading()
icn = 255 - cn
If icn < 0 Then icn = 0
If icn > 255 Then icn = 255
picCanvas.Circle (xs, ys), zrad, RGB(icn, icn, icn), , , zratio
picCanvas.PSet (xs, ys), RGB(icn, icn, icn)
EvalTangents xs, ys, zrad, xp, yp, x1, y1, x2, y2
If x1 = xp And y1 = yp Then
   picCanvas.DrawWidth = 2
   ConcentricShadedCone
   picCanvas.DrawWidth = 1
   DrawingMode = False
   Exit Sub
End If
         
picCanvas.Line (x1, y1)-(xp, yp), RGB(icn, icn, icn)
picCanvas.Line (x2, y2)-(xp, yp), RGB(icn, icn, icn)
         
picCanvas.DrawWidth = 2
'Use tangent chord theorem
zL = Sqr((yp - ys) ^ 2 + (xp - xs) ^ 2)
zPT2 = zL ^ 2 - zrad ^ 2
ztheta = zAtn2(yp - ys, xp - xs)
zthetadeg = r2d# * ztheta
If ztheta < -pi# / 2 And ztheta > -pi# Then
   ztheta = ztheta + 2 * pi#
End If
zstartangle = zAtn2(ys - y1, x1 - xs)
zendangle = zstartangle + 2 * ztheta
zsn = 255 - rubcn
If zsn < 0 Then zsn = 0
If zsn > 255 Then zsn = 255
zculsteps = (cn - rubcn) / 720
N = 0
For zi = -zstartangle To zendangle Step pi# / 720
   zsn = zsn + zculsteps
   If zsn > 255 Then zsn = 255: zculsteps = -zculsteps
   If zsn < 1 Then zsn = 1: zculsteps = Abs(zculsteps)
   jsn = Int(zsn)
   yQ = ys + zrad * Sin(zi)
   xQ = xs + zrad * Cos(zi)
   zPM = xp - xQ
   zQM = yp - yQ: If zQM = 0 Then zQM = 0.02
   zPQ2 = zPM ^ 2 + zQM ^ 2: If zPQ2 = 0 Then zPQ = 0.02
   zdy = (zPT2 / zPQ2) * zQM
   zdx = zPM * zdy / zQM
   y3 = yp - zdy
   x3 = xp - zdx
   picCanvas.Line (xQ, yQ)-(xp, yp), RGB(jsn, jsn, jsn)
   cnn = 255 - CInt(cn - (zsn - rubcn))
   If cnn < 0 Then cnn = 0
   If cnn > 255 Then cnn = 255
   picCanvas.Line (xQ, yQ)-(x3, y3), RGB(cnn, cnn, cnn)
Next zi

End Sub
Private Sub ConeVertShading()
icn = 255 - cn
If icn < 0 Then icn = 0
If icn > 255 Then icn = 255
picCanvas.Circle (xs, ys), zrad, RGB(icn, icn, icn), , , zratio
picCanvas.PSet (xs, ys), RGB(icn, icn, icn)
EvalTangents xs, ys, zrad, xp, yp, x1, y1, x2, y2
If x1 = xp And y1 = yp Then
   picCanvas.DrawWidth = 2
   ConcentricShadedCone
   picCanvas.DrawWidth = 1
   DrawingMode = False
   Exit Sub
End If
         
picCanvas.Line (x1, y1)-(xp, yp), RGB(icn, icn, icn)
picCanvas.Line (x2, y2)-(xp, yp), RGB(icn, icn, icn)
         
zL = Sqr((xp - xs) ^ 2 + (yp - ys) ^ 2)
ztheta = zAtn2(yp - ys, xp - xs)
zstep = 0.5
If zrad > 100 Then zstep = 0.2
         
zsn = 255 - rubcn
If zsn < 0 Then zsn = 0
If zsn > 255 Then zsn = 255
If zL = 0 Then zL = 1
zculsteps = (cn - rubcn) / (zL / zstep)
         
picCanvas.DrawWidth = 2
For zn = 0 To zL Step zstep
   jsn = Int(zsn)
   xc = xs + zn * Cos(ztheta)
   yc = ys + zn * Sin(ztheta)
   zr = zrad * (1 - zn / zL)  'zr = zrad gives a shaded tube
   picCanvas.Circle (xc, yc), zr, RGB(jsn, jsn, jsn), , , zratio
   zsn = zsn + zculsteps
   If zsn > 255 Then zsn = 255: zculsteps = -zculsteps
   If zsn < 2 Then zsn = 2: zculsteps = Abs(zculsteps)
Next zn

End Sub
Private Sub ConeCenterShading()
icn = 255 - cn
If icn < 0 Then icn = 0
If icn > 255 Then icn = 255
picCanvas.Circle (xs, ys), zrad, RGB(icn, icn, icn), , , zratio
picCanvas.PSet (xs, ys), RGB(icn, icn, icn)
EvalTangents xs, ys, zrad, xp, yp, x1, y1, x2, y2
         
If x1 = xp And y1 = yp Then
   picCanvas.DrawWidth = 2
   ConcentricShadedCone
   picCanvas.DrawWidth = 1
   DrawingMode = False
   Exit Sub
End If
         
picCanvas.Line (x1, y1)-(xp, yp), RGB(icn, icn, icn)
picCanvas.Line (x2, y2)-(xp, yp), RGB(icn, icn, icn)
         
picCanvas.DrawWidth = 2
'Use tangent chord theorem 'Cone
zL = Sqr((yp - ys) ^ 2 + (xp - xs) ^ 2)
zPT2 = zL ^ 2 - zrad ^ 2
ztheta = zAtn2(yp - ys, xp - xs)
zthetadeg = r2d# * ztheta
If ztheta < -pi# / 2 And ztheta > -pi# Then
   ztheta = ztheta + 2 * pi#
End If
zstartangle = zAtn2(ys - y1, x1 - xs)
zendangle = zstartangle + 2 * ztheta
         
zsn = 255 - rubcn
If zsn < 0 Then zsn = 0
If zsn > 255 Then zsn = 255
zculsteps = (cn - rubcn) / 90
N = 0
zs = (-zstartangle + zendangle) / 2
zii = zs
For zi = zs To zendangle Step pi# / 180
   zsn = zsn + zculsteps
   If zsn > 255 Then zsn = 255: zculsteps = -zculsteps
   If zsn < 1 Then zsn = 1: zculsteps = Abs(zculsteps)
   jsn = Int(zsn)
   yQ = ys + zrad * Sin(zi)
   xQ = xs + zrad * Cos(zi)
   zPM = xp - xQ
   zQM = yp - yQ: If zQM = 0 Then zQM = 0.02
   zPQ2 = zPM ^ 2 + zQM ^ 2: If zPQ2 = 0 Then zPQ2 = 0.02
   zdy = (zPT2 / zPQ2) * zQM
            
   If zQM <> 0 Then
      zdx = zPM * zdy / zQM
      y3 = yp - zdy
      x3 = xp - zdx
      picCanvas.Line (xQ, yQ)-(xp, yp), RGB(jsn, jsn, jsn)
      cnn = 255 - CInt(cn - (zsn - rubcn))
      If cnn < 1 Then cnn = 1
      If cnn > 255 Then cnn = 255
      picCanvas.Line (xQ, yQ)-(x3, y3), RGB(cnn, cnn, cnn)
   End If
            
   yQ = ys + zrad * Sin(zii)
   xQ = xs + zrad * Cos(zii)
   zPM = xp - xQ
   zQM = yp - yQ: If zQM = 0 Then zQM = 0.02
   zPQ2 = zPM ^ 2 + zQM ^ 2: If zPQ2 = 0 Then zPQ2 = 0.02
   zdy = (zPT2 / zPQ2) * zQM
            
   If zQM <> 0 Then
      zdx = zPM * zdy / zQM
      y3 = yp - zdy
      x3 = xp - zdx
      picCanvas.Line (xQ, yQ)-(xp, yp), RGB(jsn, jsn, jsn)
      picCanvas.Line (xQ, yQ)-(x3, y3), RGB(cnn, cnn, cnn)
   End If
            
   zii = zii - pi# / 180
Next zi

End Sub


Private Sub StartTube(X, Y)
SetCanvasRectangle
LCount = LCount + 1
If (Button = 1 Or LCount = 1) And RedrawPolCoTu = 0 Then 'LCul or  RCul to start fresh cone
'If LCount = 1 Then
   picCanvas.DrawMode = 7
   xs = X: ys = Y
   zrad = Abs(X - xs)
   zratio = 1
   Select Case RectangleType
   Case 0, 1, 2, 6, 7, 8, 9   'Tube
      picCanvas.Circle (xs, ys), zrad, PenCul&, , , zratio
   Case 3, 4, 5   'Tube
      picCanvas.Circle (xs, ys), zrad, PenCul&, , , zratio
      zrad2 = zrad - zpspac
      If zrad2 < 0 Then zrad2 = 0
      picCanvas.Circle (xs, ys), zrad2, PenCul&, , , zratio
   End Select
      
ElseIf LCount = 2 Then
   xp = X: yp = Y
   picCanvas.Line (xs, ys)-(X, Y), PenCul&

ElseIf LCount = 3 And RedrawPolCoTu = 0 Then  'Move whole Cone outline
   RedrawPolCoTu = 1
   CANVASRESTORE
   LCount = 3

Else
   'Redraw Tube
   CANVASRESTORE
   picCanvas.DrawMode = 13
   
   Select Case RectangleType
   Case 0, 1, 2, 6   '3 line thickness & dotted 'Tube
      TubeLineThickness_Dotted
      
   Case 3, 4, 5      '3 double-line spacings 'Tube
      TubeDoubleLine
      
   Case 7   'Tube: along axis shading
      TubeAlongAxisShading
      picCanvas.DrawWidth = 1
      
   Case 8   'Tube: across axis vert shading
      TubeAcrossAxisVertShading
      picCanvas.DrawWidth = 1
      
   Case 9   'Tube: concentric shading actually vert & horiz. need sorting
      TubeConcentricShading
      picCanvas.DrawWidth = 1
   
   End Select
   
   DrawingMode = False
End If
End Sub

Private Sub DrawTube(X, Y)
If RedrawPolCoTu > 0 Then  'Move whole Cone outline
   If RedrawPolCoTu > 1 Then   'Skip once after CANVASRESTORE just done
      FixedConeTubeBaseOutLine
      picCanvas.Line (xs, ys)-(xp, yp), PenCul&
      xs = xs + (X - xprev)
      ys = ys + (Y - yprev)
      xp = xp + (X - xprev)
      yp = yp + (Y - yprev)
      xprev = X: yprev = Y
   End If
   
   FixedConeTubeBaseOutLine
   picCanvas.Line (xs, ys)-(xp, yp), PenCul&
   
   RedrawPolCoTu = 2 'RedrawPolCoTu + 1  'To allow full redraw
   xprev = X: yprev = Y
Else
   If LCount = 1 Then
      ConeTubeBaseOutLine X, Y
   ElseIf LCount = 2 Then  'Line from circle center along cone axis
      picCanvas.Line (xs, ys)-(xp, yp), PenCul&
      xp = X: yp = Y
      picCanvas.Line (xs, ys)-(xp, yp), PenCul&
   End If
   xprev = X: yprev = Y
End If
End Sub

Private Sub TubeLineThickness_Dotted()
picCanvas.Circle (xs, ys), zrad, DrawCul&, , , zratio
picCanvas.PSet (xs, ys), DrawCul&
EvalDiameters xs, ys, zrad, xp, yp, x1, y1, x2, y2, x3, y3, x4, y4
         
'Make end of tube a semi-circle
zang2 = zAtn2(y1 - ys, x1 - xs)
If zang2 < 0 Then
   zang2 = Abs(zang2)
Else
   zang2 = 2 * pi# - zang2
End If
zang1 = zAtn2(y2 - ys, x2 - xs)
If zang1 < 0 Then
   zang1 = Abs(zang1)
Else
   zang1 = 2 * pi# - zang1
End If
picCanvas.Circle (xp, yp), zrad, DrawCul&, zang1, zang2, zratio
picCanvas.PSet (xp, yp), DrawCul&
picCanvas.Line (x1, y1)-(x3, y3), DrawCul&
picCanvas.Line (x2, y2)-(x4, y4), DrawCul&

End Sub
Private Sub TubeDoubleLine()
picCanvas.Circle (xs, ys), zrad, DrawCul&, , , zratio
picCanvas.PSet (xs, ys), DrawCul&
         
EvalDiameters xs, ys, zrad, xp, yp, x1, y1, x2, y2, x3, y3, x4, y4
         
'Make end of tube a semi-circle
zang2 = zAtn2(y1 - ys, x1 - xs)
If zang2 < 0 Then
   zang2 = Abs(zang2)
Else
   zang2 = 2 * pi# - zang2
End If
zang1 = zAtn2(y2 - ys, x2 - xs)
If zang1 < 0 Then
   zang1 = Abs(zang1)
Else
   zang1 = 2 * pi# - zang1
End If
picCanvas.Circle (xp, yp), zrad, DrawCul&, zang1, zang2, zratio
         
picCanvas.PSet (xp, yp), DrawCul&
picCanvas.Line (x1, y1)-(x3, y3), DrawCul&
picCanvas.Line (x2, y2)-(x4, y4), DrawCul&
         
zrad2 = zrad - zpspac
If zrad2 < 0 Then zrad2 = 0
picCanvas.Circle (xs, ys), zrad2, DrawCul&, , , zratio
EvalDiameters xs, ys, zrad2, xp, yp, x1, y1, x2, y2, x3, y3, x4, y4
'Make end of tube a semi-circle
zang2 = zAtn2(y1 - ys, x1 - xs)
If zang2 < 0 Then
   zang2 = Abs(zang2)
Else
   zang2 = 2 * pi# - zang2
End If
zang1 = zAtn2(y2 - ys, x2 - xs)
If zang1 < 0 Then
   zang1 = Abs(zang1)
Else
   zang1 = 2 * pi# - zang1
End If
         
picCanvas.Circle (xp, yp), zrad2, DrawCul&, zang1, zang2, zratio
         
picCanvas.PSet (xp, yp), DrawCul&
picCanvas.Line (x1, y1)-(x3, y3), DrawCul&
picCanvas.Line (x2, y2)-(x4, y4), DrawCul&

End Sub
Private Sub TubeAlongAxisShading()
icn = 255 - cn
If icn < 0 Then icn = 0
If icn > 255 Then icn = 255
picCanvas.Circle (xs, ys), zrad, RGB(icn, icn, icn), , , zratio
picCanvas.PSet (xs, ys), RGB(icn, icn, icn)
EvalDiameters xs, ys, zrad, xp, yp, x1, y1, x2, y2, x3, y3, x4, y4
picCanvas.Circle (xp, yp), zrad, RGB(icn, icn, icn), , , zratio
picCanvas.PSet (xp, yp), RGB(icn, icn, icn)
picCanvas.Line (x1, y1)-(x3, y3), RGB(icn, icn, icn)
picCanvas.Line (x2, y2)-(x4, y4), RGB(icn, icn, icn)
         
picCanvas.DrawWidth = 2
ztheta = zAtn2(yp - ys, xp - xs)
zthetadeg = r2d# * ztheta
zstartangle = -(pi# / 2 - ztheta)
zendangle = pi# / 2 + ztheta
         
zsn = 255 - rubcn
If zsn < 0 Then zsn = 0
If zsn > 255 Then zsn = 255
zculsteps = (cn - rubcn) / 180
For zi = zstartangle To zendangle Step pi# / 180
   jsn = Int(zsn)
   y1 = ys + zrad * Sin(zi)
   x1 = xs + zrad * Cos(zi)
   y2 = yp + zrad * Sin(zi)
   x2 = xp + zrad * Cos(zi)
   y3 = y1 - 2 * zrad * Cos(zi - ztheta) * Sin(ztheta)
   x3 = x1 - 2 * zrad * Cos(zi - ztheta) * Cos(ztheta)
   picCanvas.Line (x1, y1)-(x2, y2), RGB(jsn, jsn, jsn)
            
   cnn = CInt(cn - (zsn - rubcn))
   If cnn < 0 Then cnn = 0
   If cnn > 255 Then cnn = 255
   picCanvas.Line (x1, y1)-(x3, y3), RGB(cnn, cnn, cnn)
   zsn = zsn + zculsteps
   If zsn > 255 Then zsn = 255: zculsteps = -zculsteps
   If zsn < 1 Then zsn = 1: zculsteps = Abs(zculsteps)
Next zi

End Sub
Private Sub TubeAcrossAxisVertShading()
icn = 255 - cn
If icn < 0 Then icn = 0
If icn > 255 Then icn = 255
picCanvas.Circle (xs, ys), zrad, RGB(icn, icn, icn), , , zratio
picCanvas.PSet (xs, ys), RGB(icn, icn, icn)
EvalTangents xs, ys, zrad, xp, yp, x1, y1, x2, y2
picCanvas.Circle (xp, yp), zrad, RGB(icn, icn, icn), , , zratio
picCanvas.PSet (xp, yp), RGB(icn, icn, icn)
picCanvas.Line (x1, y1)-(xp, yp), RGB(icn, icn, icn)
picCanvas.Line (x2, y2)-(xp, yp), RGB(icn, icn, icn)
         
zL = Sqr((xp - xs) ^ 2 + (yp - ys) ^ 2)
ztheta = zAtn2(yp - ys, xp - xs)
zstep = 0.5
If zrad > 100 Then zstep = 0.2
         
zsn = 255 - rubcn
If zsn < 0 Then zsn = 0
If zsn > 255 Then zsn = 255
If zL = 0 Then zL = 1
zculsteps = (cn - rubcn) / (zL / zstep)
         
picCanvas.DrawWidth = 2
For zn = 0 To zL Step zstep
   jsn = Int(zsn)
   xc = xs + zn * Cos(ztheta)
   yc = ys + zn * Sin(ztheta)
   zr = zrad '* (1 - zn / zL)  'zr = zrad gives a shaded tube
   picCanvas.Circle (xc, yc), zr, RGB(jsn, jsn, jsn), , , zratio
   zsn = zsn + zculsteps
   If zsn > 255 Then zsn = 255: zculsteps = -zculsteps
   If zsn < 1 Then zsn = 1: zculsteps = Abs(zculsteps)
Next zn

End Sub
Private Sub TubeConcentricShading()
icn = 255 - cn
If icn < 0 Then icn = 0
If icn > 255 Then icn = 255
picCanvas.Circle (xs, ys), zrad, RGB(icn, icn, icn), , , zratio
picCanvas.PSet (xs, ys), RGB(icn, icn, icn)
EvalDiameters xs, ys, zrad, xp, yp, x1, y1, x2, y2, x3, y3, x4, y4
picCanvas.Circle (xp, yp), zrad, RGB(icn, icn, icn), , , zratio
picCanvas.PSet (xp, yp), RGB(icn, icn, icn)
picCanvas.Line (x1, y1)-(x3, y3), RGB(icn, icn, icn)
picCanvas.Line (x2, y2)-(x4, y4), RGB(icn, icn, icn)
         
picCanvas.DrawWidth = 2
ztheta = zAtn2(yp - ys, xp - xs)
zthetadeg = r2d# * ztheta
zstartangle = -(pi# / 2 - ztheta)
zendangle = pi# / 2 + ztheta
         
zsn = 255 - rubcn
If zsn < 0 Then zsn = 0
If zsn > 255 Then zsn = 255
zculsteps = (cn - rubcn) / 180
For zi = ztheta To zendangle Step pi# / 180
   jsn = Int(zsn)
   y1 = ys + zrad * Sin(zi)
   x1 = xs + zrad * Cos(zi)
   y2 = yp + zrad * Sin(zi)
   x2 = xp + zrad * Cos(zi)
   y3 = y1 - 2 * zrad * Cos(zi - ztheta) * Sin(ztheta)
   x3 = x1 - 2 * zrad * Cos(zi - ztheta) * Cos(ztheta)
   picCanvas.Line (x1, y1)-(x2, y2), RGB(jsn, jsn, jsn)
            
   cnn = CInt(cn - (zsn - rubcn))
   If cnn < 0 Then cnn = 0
   If cnn > 255 Then cnn = 255
   picCanvas.Line (x1, y1)-(x3, y3), RGB(cnn, cnn, cnn)
            
   zi2 = 2 * ztheta - zi '@ end zi2=(pi#/2+ztheta)-pi# = -(pi#/2-ztheta)
   y1 = ys + zrad * Sin(zi2)
   x1 = xs + zrad * Cos(zi2)
   y2 = yp + zrad * Sin(zi2)
   x2 = xp + zrad * Cos(zi2)
   y3 = y1 - 2 * zrad * Cos(zi2 - ztheta) * Sin(ztheta)
   x3 = x1 - 2 * zrad * Cos(zi2 - ztheta) * Cos(ztheta)
   picCanvas.Line (x1, y1)-(x2, y2), RGB(jsn, jsn, jsn)
   picCanvas.Line (x1, y1)-(x3, y3), RGB(cnn, cnn, cnn)
   zsn = zsn + zculsteps
   If zsn > 255 Then zsn = 255: zculsteps = -zculsteps
   If zsn < 1 Then zsn = 1: zculsteps = Abs(zculsteps)
Next zi

End Sub
      

Private Sub StartArch(X, Y)
SetCanvasRectangle
LCount = LCount + 1
If LCount = 1 Then
   picCanvas.DrawMode = 7
   xs = X: ys = Y
   xp = X: yp = Y
   Select Case RectangleType
   Case 0, 1, 2, 6, 7, 8, 9   'Arch
      DrawSingleArch PenCul&
   Case 3, 4, 5   'Double line Arch
      DrawDoubleArch PenCul&
   End Select
ElseIf LCount = 2 Then
   'Now move whole piece
   
Else  'Redraw Arch
   CANVASRESTORE
   picCanvas.DrawMode = 13
   Select Case RectangleType
   Case 0, 1, 2, 6   'Arch
      DrawSingleArch DrawCul&
   Case 3, 4, 5   'Double line Arch
      DrawDoubleArch DrawCul&
   Case 7   'Horz arch shading
      ArchHorzShading
      
   Case 8   'Vert Arch shading
      ArchVertShading
      
   Case 9   'Concentric Arch shading
      ArchConcentricShading
   End Select
      
   DrawingMode = False
End If
End Sub
Private Sub ArchHorzShading()
'NB Up mouse movement produces shaded top shell which is useful
'but if too far up nothing is produced!
If xs > xp Then
   xt = xs: xs = xp: xp = xt
End If
xc = (xs + xp) / 2
zrad = Abs(xp - xs) / 2
zH = yp - ys + zrad
If Abs(zH) = 0 Then zH = 1
         
zsn = 255 - rubcn
If zsn < 0 Then zsn = 0
If zsn > 255 Then zsn = 255
zculsteps = (cn - rubcn) / zH
xc = (xs + xp) / 2
For IY = ys - zrad To yp
   jsn = Int(zsn)
   If IY < ys Then  'top semi-circle
      zdx = Sqr(Abs(zrad ^ 2 - (IY - ys) ^ 2))
      x1 = xc - zdx: x2 = xc + zdx
      picCanvas.Line (x1, IY)-(x2, IY), RGB(jsn, jsn, jsn) 'BH
   Else             'bottom rectangle
      picCanvas.Line (xs, IY)-(xp, IY), RGB(jsn, jsn, jsn)  'BH
   End If
            
   zsn = zsn + zculsteps
   If zsn > 255 Then zsn = 255: zculsteps = -zculsteps
   If zsn < 1 Then zsn = 1: zculsteps = Abs(zculsteps)
Next IY

End Sub
Private Sub ArchVertShading()
If xs > xp Then
   xt = xs: xs = xp: xp = xt
End If
xc = (xs + xp) / 2
zrad = Abs(xp - xs) / 2
zW = Abs(xp - xs)
If Abs(zW) = 0 Then zW = 1
         
zsn = 255 - rubcn
If zsn < 0 Then zsn = 0
If zsn > 255 Then zsn = 255
zculsteps = (cn - rubcn) / zW
xc = (xs + xp) / 2
For IX = xs To xp
   jsn = Int(zsn)
   zytop = ys - Sqr(Abs(zrad ^ 2 - (IX - xc) ^ 2))
   picCanvas.Line (IX, zytop)-(IX, yp), RGB(jsn, jsn, jsn) 'BH
   zsn = zsn + zculsteps
   If zsn > 255 Then zsn = 255: zculsteps = -zculsteps
   If zsn < 1 Then zsn = 1: zculsteps = Abs(zculsteps)
Next IX

End Sub
      
Private Sub ArchConcentricShading()
If xs > xp Then
   xt = xs: xs = xp: xp = xt
End If
xc = (xs + xp) / 2
zrad = Abs(xp - xs) / 2
zW = Abs(xp - xs) / 2
If Abs(zW) = 0 Then zW = 1
         
zsn = 255 - rubcn
If zsn < 0 Then zsn = 0
If zsn > 255 Then zsn = 255
zculsteps = (cn - rubcn) / zW
For IX = xs To xc
   jsn = Int(zsn)
   zytop = ys - Sqr(Abs(zrad ^ 2 - (IX - xc) ^ 2))
   picCanvas.Line (IX, zytop)-(IX, yp), RGB(jsn, jsn, jsn) 'BH
   zsn = zsn + zculsteps
   If zsn > 255 Then zsn = 255: zculsteps = -zculsteps
   If zsn < 1 Then zsn = 1: zculsteps = Abs(zculsteps)
Next IX
For IX = xc To xp
   jsn = Int(zsn)
   zytop = ys - Sqr(Abs(zrad ^ 2 - (IX - xc) ^ 2))
   picCanvas.Line (IX, zytop)-(IX, yp), RGB(jsn, jsn, jsn)  'BH
   zsn = zsn + zculsteps
   If zsn > 255 Then zsn = 255: zculsteps = -zculsteps
   If zsn < 1 Then zsn = 1: zculsteps = Abs(zculsteps)
Next IX

End Sub

Private Sub DrawArch(X, Y)

Select Case RectangleType
Case 0, 1, 2, 6, 7, 8, 9   'Arch
   If LCount = 1 Then
      DrawSingleArch PenCul&
      xp = X: yp = Y
      DrawSingleArch PenCul&
      xprev = X: yprev = Y
   ElseIf LCount = 2 Then  'Move whole Arch
      DrawSingleArch PenCul&
      xs = xs + (X - xprev)
      ys = ys + (Y - yprev)
      xp = xp + (X - xprev)
      yp = yp + (Y - yprev)
      xprev = X: yprev = Y
      DrawSingleArch PenCul&
   End If
Case 3, 4, 5   'Double line Arch
   If LCount = 1 Then
      DrawDoubleArch PenCul&
      xp = X: yp = Y
      DrawDoubleArch PenCul&
      xprev = X: yprev = Y
   ElseIf LCount = 2 Then  'Move whole Arch
      DrawDoubleArch PenCul&
      xs = xs + (X - xprev)
      ys = ys + (Y - yprev)
      xp = xp + (X - xprev)
      yp = yp + (Y - yprev)
      xprev = X: yprev = Y
      DrawDoubleArch PenCul&
   End If
End Select
End Sub

Private Sub DrawSingleArch(cul&)
picCanvas.Line (xs, ys)-(xs, yp), cul&  'LV
picCanvas.Line (xs, yp)-(xp, yp), cul&  'BH
picCanvas.Line (xp, yp)-(xp, ys), cul&  'LV
xc = (xs + xp) / 2
zrad = Abs(xp - xs) / 2
picCanvas.Circle (xc, ys), zrad, cul&, 0, pi#
End Sub
Private Sub DrawDoubleArch(cul&)
picCanvas.Line (xs, ys)-(xs, yp), cul&  'LV
picCanvas.Line (xs, yp)-(xp, yp), cul&  'BH
picCanvas.Line (xp, yp)-(xp, ys), cul&  'LV
picCanvas.Line (xs + zpspac, ys)-(xs + zpspac, yp - zpspac), cul& 'LV
picCanvas.Line (xs + zpspac, yp - zpspac)-(xp - zpspac, yp - zpspac), cul& 'BH
picCanvas.Line (xp - zpspac, yp - zpspac)-(xp - zpspac, ys), cul& 'LV
xc = (xs + xp) / 2
zrad = Abs(xp - xs) / 2
picCanvas.Circle (xc, ys), zrad, cul&, 0, pi#
zrad2 = zrad - zpspac
If zrad2 < 0 Then zrad2 = 0
picCanvas.Circle (xc, ys), zrad2, cul&, 0, pi#
End Sub
      
Private Sub StartTPiece(X, Y)
SetCanvasLine
LCount = LCount + 1
If zpspac = 0 Then zpspac = 4
If LCount = 1 Then
   picCanvas.DrawMode = 7
   xs = X: ys = Y
   xp = X: yp = Y
   DrawTPiece PenCul&, xs, ys, xp, yp
        
ElseIf LCount = 2 Then
   'Now move whole piece
Else  'Redraw TPiece
   CANVASRESTORE
   picCanvas.DrawMode = 13
   DrawTPiece DrawCul&, xs, ys, xp, yp
      
   DrawingMode = False
End If
End Sub

Private Sub ADrawTPiece(X, Y)
If LCount = 1 Then
   DrawTPiece PenCul&, xs, ys, xp, yp
   xp = X: yp = Y
   DrawTPiece PenCul&, xs, ys, xp, yp
   xprev = X: yprev = Y
ElseIf LCount = 2 Then  'Move whole T-piece
   DrawTPiece PenCul&, xs, ys, xp, yp
   xs = xs + (X - xprev)
   ys = ys + (Y - yprev)
   xp = xp + (X - xprev)
   yp = yp + (Y - yprev)
   xprev = X: yprev = Y
   DrawTPiece PenCul&, xs, ys, xp, yp
End If
End Sub

Private Sub StartText(X, Y)
frmText.Visible = False
If TextSW = True Then
   a$ = Text2.Text
   picMCRBox.Width = picMCRBox.TextWidth(a$) + 4
   picMCRBox.Height = picMCRBox.TextHeight("Eg") + 4
   picMCRBox.Left = X '+ picCanvas.Left
   picMCRBox.Top = Y '+ picCanvas.Top
   picMCRBox.Visible = True
   picMCRBox.Cls
   picMCRBox.Print a$;
   picMCRBox.Refresh
   MousePointer = 10
End If
End Sub

Private Sub StartFill(X, Y)
picCanvas.DrawStyle = vbSolid
picCanvas.DrawMode = 13
picCanvas.DrawWidth = 1
FillPtcul& = picCanvas.Point(X, Y)

Select Case FillType
Case 0: picCanvas.FillStyle = vbFSSolid
Case 1: picCanvas.FillStyle = vbHorizontalLine
Case 2: picCanvas.FillStyle = vbVerticalLine
Case 3: picCanvas.FillStyle = vbDownwardDiagonal   '/ ?
Case 4: picCanvas.FillStyle = vbUpwardDiagonal     '\ ?
Case 5: picCanvas.FillStyle = vbCross
Case 6: picCanvas.FillStyle = vbDiagonalCross
Case Else
             
   picCanvas.ForeColor = DrawCul&
   bytindex = (FillType - 7) * 8 + 1
   hbmp = CreateBitmap(8, 8, 1, 1, BitArray(bytindex))
   hbr = CreatePatternBrush(hbmp)
   hprev = SelectObject(picCanvas.hdc, hbr)
   
End Select

picCanvas.FillColor = DrawCul&
FLOODFILLSURFACE = 1
'Fills with FillColor so long as point surrounded by FillPtcul&
rs = ExtFloodFill(picCanvas.hdc, X, Y, FillPtcul&, FLOODFILLSURFACE)
   
If FillType >= 7 Then
   rs = SelectObject(picCanvas.hdc, hprev)
   rs = DeleteObject(hbr)
   rs = DeleteObject(hbmp)
End If
   
picCanvas.FillStyle = 1  'Default (Transparent)
DrawingMode = False
picCanvas.Refresh

End Sub

Private Sub StartRubber(X, Y)
LCount = LCount + 1
RubRectCul& = picCanvas.BackColor Xor RGB(0, 0, 0)
picCanvas.DrawWidth = 1
   
Select Case RubberType
Case 0 To 2:
   Select Case RubberType
   Case 0: zpspac = 2
   Case 1: zpspac = 4
   Case 2: zpspac = 8
   End Select
   
   If LCount = 1 Then
      picCanvas.DrawMode = 7
      xs = X - zpspac: ys = Y - zpspac - 1
      xp = X + 1: yp = Y - 1
      picCanvas.Line (xs, ys)-(xp, yp), RubRectCul&, B
   Else  'Clear Rubber
      picCanvas.Line (xs, ys)-(xp, yp), RubRectCul&, B
      DrawingMode = False
   End If

Case 3:  'RGB Smudge in Active rectangle
   If InRectangle(X, Y, arleft, artop, arright, arbottom) Then
      svdrawMode = picCanvas.DrawMode
      picCanvas.DrawMode = 13
      y1 = artop + 1: y2 = arbottom - 1
      x1 = arleft + 1: x2 = arright - 1
      picCanvas.Line (x1, y1)-(x2, y2), RubberCul&, BF
      picCanvas.DrawMode = svdrawMode
      DrawingMode = False
   End If
End Select

End Sub

Private Sub DrawRubber(X, Y)

If RubberType = 3 Then Exit Sub

picCanvas.Line (xs, ys)-(xp, yp), RubRectCul&, B
picCanvas.DrawMode = 13
      
For R = ys To yp
   picCanvas.Line (xs, R)-(xp, R), RubberCul&
Next R
picCanvas.DrawMode = 7
xs = X - zpspac: ys = Y - zpspac - 1
xp = X + 1: yp = Y - 1
picCanvas.Line (xs, ys)-(xp, yp), RubRectCul&, B
End Sub

Private Sub StartActiveRectangle(X, Y)
LCount = LCount + 1
If LCount = 1 Then
   If ActiveRectExists Then
      LCount = 0
      DrawingMode = False
      cmdClearActRect_Click
   Else
      xs = X: ys = Y
      xp = X: yp = Y
      arleft = xs: artop = ys
      arright = xp: arbottom = yp
      ShowActRectCoords
      shpARect.Visible = True
      shpARect.Left = arleft
      shpARect.Top = artop
      shpARect.Width = (arright - arleft)
      shpARect.Height = (arbottom - artop)
      ActiveRectExists = True
   End If

ElseIf LCount = 2 Then
   'Now move whole piece
   

Else  'end rect
   'Make sure Act Rect big enough to get mouse in easily
   ShowActRectCoords
   DrawingMode = False
   If arwidth < 4 Or arheight < 4 Then
      cmdClearActRect_Click
      ShowActRectCoords
      ActiveRectExists = False
   Else
      'Make copy of AR points
      karleft = arleft: kartop = artop
      karright = arright: karbottom = arbottom
   End If
   LCount = 0
End If
End Sub
Private Sub DrawActiveRect(X, Y)
If LCount = 1 Then
xp = X: yp = Y
Else
      xs = xs + (X - xprev)
      ys = ys + (Y - yprev)
      xp = xp + (X - xprev)
      yp = yp + (Y - yprev)
End If
arleft = xs: artop = ys
arright = xp: arbottom = yp
ShowActRectCoords
shpARect.Left = arleft
shpARect.Top = artop
shpARect.Width = (arright - arleft)
shpARect.Height = (arbottom - artop)
xprev = X: yprev = Y
End Sub

Private Sub StartSmudge(X, Y)
'SmudgeType=NUM OK for smooth palettes
picCanvas.DrawStyle = vbSolid
picCanvas.DrawMode = 13
picCanvas.DrawWidth = 1
   
Select Case SmudgeType
Case 0:  'RGB Smudge brush in +- 4 rectangle
   For Yr = Y - 4 To Y + 4
   For Xr = X - 4 To X + 4
       Num = 0
       grey = 0
       For IY = Yr - 1 To Yr + 1
       For IX = Xr - 1 To Xr + 1
          cul& = GetPixel(picCanvas.hdc, IX, IY)
          If cul& <> -1 Then
             'All RGB components the same find average red-grey
             grey = grey + (cul& And &HFF&)
             Num = Num + 1
          End If
       Next IX
       Next IY
       If Num <> 0 Then
          grey = grey / Num
          zres = SetPixelV(picCanvas.hdc, Xr, Yr, RGB(grey, grey, grey))
       End If
    Next Xr
    Next Yr
    picCanvas.Refresh
Case 1:  'RGB Smudge in Active rectangle
   If InRectangle(X, Y, arleft, artop, arright, arbottom) Then
      MousePointer = vbHourglass
      
      For Yr = artop + 2 To arbottom - 2
      For Xr = arleft + 2 To arright - 2
          Num = 0
          grey = 0
          For IY = Yr - 1 To Yr + 1
          For IX = Xr - 1 To Xr + 1
             cul& = GetPixel(picCanvas.hdc, IX, IY)
             If cul& <> -1 Then
               'All RGB components the same find average red-grey
                grey = grey + (cul& And &HFF&)
                Num = Num + 1
             End If
          Next IX
          Next IY
          If Num <> 0 Then
             grey = grey / Num
             zres = SetPixelV(picCanvas.hdc, Xr, Yr, RGB(grey, grey, grey))
          End If
     Next Xr
     Next Yr
     MouseTArr
     'MousePointer = vbDefault
     DrawingMode = False
   End If

End Select
End Sub

Private Sub StartMCR(X, Y)
LCount = LCount + 1
If LCount = 1 Then
   'Test if in Active rectangle
   If X > arleft And X < arright And Y > artop And Y < arbottom Then
      
      MCRSW = True
      MousePointer = vbHourglass
      
      picMCRBox.Left = X - 2 '4
      picMCRBox.Top = Y - 2 '4
      picMCRBox.Width = arright - arleft + 2 '4
      picMCRBox.Height = arbottom - artop + 2 '4
      
      Select Case MCRType
      Case 0   'MOVE
         TransferActRectTopicMCRBox
      Case 1   'COPY
         TransferActRectTopicMCRBox
      Case 2   'REFLECT LEFT->RIGHT
         LRReflectActRectTopicMCRBox
      Case 3   'REFLECT TOP->BELOW
         TBReflectActRectTopicMCRBox
      End Select
      picMCRBox.Visible = True
      MousePointer = 10
   Else  'Not in Active rect, reset LCount
      LCount = 0
   End If
End If

End Sub
Private Sub StartResize(X, Y)
LCount = LCount + 1
If LCount = 1 Then
   'Test if in Active rectangle
   If InRectangle(X, Y, arleft, artop, arright, arbottom) Then
      ResizeSW = True
      MousePointer = vbHourglass
      '-----------------
       ResizeInActiveRect
      '-----------------
      MouseTArr
      'MousePointer = vbDefault
   End If
   LCount = 0
   DrawingMode = False
      MousePointer = 10
End If
End Sub

Private Sub StartRotate(X, Y)
LCount = LCount + 1
If LCount = 1 Then
   'Test if in Active rectangle 'Global arleft, artop, arright, arbottom
   If InRectangle(X, Y, arleft, artop, arright, arbottom) Then
      RotateSW = True
      MousePointer = vbHourglass
      '--------------
      Rotator
      '--------------
      LCount = 0
      DrawingMode = False
      MousePointer = 10
   End If
End If
End Sub

Private Sub StartTile(X, Y)
LCount = LCount + 1
If LCount = 1 Then
   'Test if in Active rectangle 'Global arleft, artop, arright, arbottom
   If InRectangle(X, Y, arleft, artop, arright, arbottom) Then
      MousePointer = vbHourglass
      '--------------
      TileActiveRectangle X, Y
      '--------------
      LCount = 0
      DrawingMode = False
      MouseTArr
      'MousePointer = vbDefault
   End If
End If
End Sub

Private Sub UpDown2_DownClick()  'SCROLL LEFT
If ActionInProgress Then Exit Sub
ScrollStep = Val(txtScroller.Text)
If ActiveRectExists Then
   arleft = arleft - ScrollStep
   arright = arright - ScrollStep
   If arleft < 0 Or arright < 0 Then
      arleft = arleft + ScrollStep
      arright = arright + ScrollStep
      cmdClearActRect_Click
   End If
   ShowActRectCoords
End If
ScrollParam = 4
ScrollPicture ScrollParam
picCanvas.Refresh
End Sub
Private Sub UpDown2_UpClick() 'Scroll RIGHT
If ActionInProgress Then Exit Sub
ScrollStep = Val(txtScroller.Text)
If ActiveRectExists Then
   arleft = arleft + ScrollStep
   arright = arright + ScrollStep
   If arleft >= picCanvas.Width Or arright >= picCanvas.Width Then
      arleft = arleft - ScrollStep
      arright = arright - ScrollStep
      cmdClearActRect_Click
   End If
   ShowActRectCoords
End If
ScrollParam = 3
ScrollPicture ScrollParam
picCanvas.Refresh
End Sub
Private Sub UpDown3_DownClick()  'SCROLL DOWN
If ActionInProgress Then Exit Sub
ScrollStep = Val(txtScroller.Text)
If ActiveRectExists Then
   artop = artop + ScrollStep
   arbottom = arbottom + ScrollStep
   If artop >= picCanvas.Height Or arbottom >= picCanvas.Height Then
      artop = artop - ScrollStep
      arbottom = arbottom - ScrollStep
      cmdClearActRect_Click
   End If
   ShowActRectCoords
End If
ScrollParam = 1
ScrollPicture ScrollParam
picCanvas.Refresh
End Sub
Private Sub UpDown3_UpClick() 'SCROLL UP
If ActionInProgress Then Exit Sub
ScrollStep = Val(txtScroller.Text)
If ActiveRectExists Then
   artop = artop - ScrollStep
   arbottom = arbottom - ScrollStep
   If artop < 0 Or arbottom < 0 Then
      artop = artop + ScrollStep
      arbottom = arbottom + ScrollStep
      cmdClearActRect_Click
   End If
   ShowActRectCoords
End If
ScrollParam = 2
ScrollPicture ScrollParam
picCanvas.Refresh
End Sub

Private Sub ScrollPicture(ScrollParam)
CANVASSAVE
      
picMCRBox.BorderStyle = 0
picMCRBox.Cls
picMCRBox.Width = picCanvas.Width
picMCRBox.Height = picCanvas.Height

picCanvas.DrawMode = 13
CanvasDC& = picCanvas.hdc
StoreDC& = picMCRBox.hdc
'ASSIGN SOURCE & DEST COORDS
dwRop& = &HCC0020  'SRCCOPY Src to Dest
Select Case ScrollParam
Case 1   'PIC DOWN
   'Save rectangle (height ScrollStep) (width picCanvas.Width)
   'from bottom of Canvas to top of Store
   SorcX = 0: SorcY = picCanvas.Height - ScrollStep
   DestX = 0: DestY = 0
   DestHeight = ScrollStep: DestWidth = picCanvas.Width
   Success& = BitBlt(StoreDC&, DestX, DestY, DestWidth, DestHeight, CanvasDC&, SorcX, SorcY, dwRop&)
   'Move Canvas rectangle (height picCanvas.Height - ScrollStep) (width picCanvas.Width)
   'down ScrollStep lines: bottom button v
   SorcX = 0: SorcY = 0
   DestX = 0: DestY = ScrollStep
   DestHeight = picCanvas.Height - ScrollStep
   DestWidth = picCanvas.Width
   Success& = BitBlt(CanvasDC&, DestX, DestY, DestWidth, DestHeight, CanvasDC&, SorcX, SorcY, dwRop&)
   'Move saved rectangle (height ScrollStep) (width picCanvas.Width)
   'from top of Store to top of Canvas
   SorcX = 0: SorcY = 0
   DestX = 0: DestY = 0
   DestHeight = ScrollStep
   DestWidth = picCanvas.Width
   Success& = BitBlt(CanvasDC&, DestX, DestY, DestWidth, DestHeight, StoreDC&, SorcX, SorcY, dwRop&)
Case 2   'PIC UP
   'Save rectangle (height ScrollStep) (width picCanvas.Width)
   'from top of Canvas to top of Store
   SorcX = 0: SorcY = 0
   DestX = 0: DestY = 0
   DestHeight = ScrollStep: DestWidth = picCanvas.Width
   Success& = BitBlt(StoreDC&, DestX, DestY, DestWidth, DestHeight, CanvasDC&, SorcX, SorcY, dwRop&)
   'Move Canvas rectangle (height picCanvas.Height - ScrollStep) (width picCanvas.Width)
   'up ScrollStep lines: top button ^
   SorcX = 0: SorcY = ScrollStep
   DestX = 0: DestY = 0
   DestHeight = picCanvas.Height - ScrollStep
   DestWidth = picCanvas.Width
   Success& = BitBlt(CanvasDC&, DestX, DestY, DestWidth, DestHeight, CanvasDC&, SorcX, SorcY, dwRop&)
   'Move saved rectangle (height ScrollStep) (width picCanvas.Width)
   'from top of Store to bottom of Canvas
   SorcX = 0: SorcY = 0
   DestX = 0: DestY = picCanvas.Height - ScrollStep
   DestHeight = ScrollStep
   DestWidth = picCanvas.Width
   Success& = BitBlt(CanvasDC&, DestX, DestY, DestWidth, DestHeight, StoreDC&, SorcX, SorcY, dwRop&)
Case 3   'PIC RIGHT
   'Save right rectangle (height picCanvas.Height) (width ScrollStep)
   'from Right of Canvas to Left of Store
   SorcX = picCanvas.Width - ScrollStep: SorcY = 0
   DestX = 0: DestY = 0
   DestHeight = picCanvas.Height: DestWidth = ScrollStep
   Success& = BitBlt(StoreDC&, DestX, DestY, DestWidth, DestHeight, CanvasDC&, SorcX, SorcY, dwRop&)
   'Move Canvas rectangle (height picCanvas.Height)  (width picCanvas.Width - ScrollStep)
   'ScrollStep pixel lines right: right button >
   SorcX = 0: SorcY = 0
   DestX = ScrollStep: DestY = 0
   DestHeight = picCanvas.Height
   DestWidth = picCanvas.Width - ScrollStep
   Success& = BitBlt(CanvasDC&, DestX, DestY, DestWidth, DestHeight, CanvasDC&, SorcX, SorcY, dwRop&)
   'Move saved rectangle (height picCanvas.Height) (width ScrollStep)
   'from Left of Store to Left of Canvas
   SorcX = 0: SorcY = 0
   DestX = 0: DestY = 0
   DestHeight = picCanvas.Height
   DestWidth = ScrollStep
   Success& = BitBlt(CanvasDC&, DestX, DestY, DestWidth, DestHeight, StoreDC&, SorcX, SorcY, dwRop&)
Case 4   'PIC LEFT
   'Save left rectangle (height picCanvas.Height) (width ScrollStep)
   'from Left of Canvas to Left of Store
   SorcX = 0: SorcY = 0
   DestX = 0: DestY = 0
   DestHeight = picCanvas.Height
   DestWidth = ScrollStep
   Success& = BitBlt(StoreDC&, DestX, DestY, DestWidth, DestHeight, CanvasDC&, SorcX, SorcY, dwRop&)
   'Move Canvas rectangle (height picCanvas.Height)  (width picCanvas.Width - ScrollStep)
   'ScrollStep pixel lines left: left button <
   SorcX = ScrollStep: SorcY = 0
   DestX = 0: DestY = 0
   DestHeight = picCanvas.Height
   DestWidth = picCanvas.Width - ScrollStep
   Success& = BitBlt(CanvasDC&, DestX, DestY, DestWidth, DestHeight, CanvasDC&, SorcX, SorcY, dwRop&)
   'Move saved rectangle (height picCanvas.Height) (width ScrollStep)
   'from Left of Store to Right of Canvas
   SorcX = 0: SorcY = 0
   DestX = picCanvas.Width - ScrollStep: DestY = 0
   DestHeight = picCanvas.Height
   DestWidth = ScrollStep
   Success& = BitBlt(CanvasDC&, DestX, DestY, DestWidth, DestHeight, StoreDC&, SorcX, SorcY, dwRop&)
End Select
picMCRBox.BorderStyle = 1
picMCRBox.Cls
End Sub

Private Sub cmdReset_Click()

frmText.Visible = False

chkHairs(0).Value = vbUnchecked
chkHairs(1).Value = vbUnchecked


chkZOOM.Value = vbUnchecked
ZoomMode = False
picZOOMBox.Cls
picZOOMBox.Visible = False


picMCRBox.Cls
picMCRBox.Visible = False

DrawingMode = False
cmdClearActRect_Click
ShowActRectCoords
ActiveRectExists = False
'------------------

ReDim xsto(1000), ysto(1000)
PCount = 0  'Point count where needed
zpspac = 0  'Parallel spacing ie none
LCount = 0  'LC COUNT

chkTOOLS(TOOL).Value = vbUnchecked '0
TOOL = 0
chkTOOLS(TOOL).Value = vbChecked '1
ShowInstructionsAndSubToolBars
MousePointer = 0
'CLEAR SWITCHES
DrawingMode = False
chkZOOM.Value = vbUnchecked
ZoomMode = False
TextSW = False
MCRSW = False
ResizeSW = False
PerspectiveSW = False
RotateSW = False
AddBMPSW = False
ActiveRectExists = False
zoomdrawing = False

LabWait.Visible = False

picCanvas.PSet (0, 0), 0

ShowPlusLines = False
Line1.Visible = False
Line2.Visible = False
ShowXLines = False
Line3.Visible = False
Line4.Visible = False
ShowPerspecLines = False
Line5.Visible = False
Line6.Visible = False
Line7.Visible = False
Line8.Visible = False
Shape1(0).Visible = False
Shape1(1).Visible = False
Shape1(2).Visible = False
SetPerspecPts = False
NumPPts = 0
shpARect.Visible = False

ClearHelpStuff
a = ActionInProgress
End Sub

Private Function RRMsgBox$(Cap$, Prom$, Com1, Com2, Com3, N)  'YES,NO,OK,NumLines
frmMsg.Height = 2000 + N * 200
frmMsg.Caption = Cap$
frmMsg.Label1.Caption = Prom$
If Com1 = 0 Then
frmMsg.Command1.Visible = False
Else
frmMsg.Command1.Visible = True
End If
If Com2 = 0 Then
frmMsg.Command2.Visible = False
Else
frmMsg.Command2.Visible = True
End If
If Com3 = 0 Then
frmMsg.Command3.Visible = False
Else
frmMsg.Command3.Visible = True
End If
frmMsg.Show vbModal
RRMsgBox$ = LabMsg.Caption
End Function

Private Sub ClearPicture_Click()
If ActionInProgress Then Exit Sub
resp = RRMsgBox$("RRPencil", "CLEAR PICTURE ?", 1, 1, 0, 1) 'YES,NO,OK, NumLines
If resp = vbYes Then
   picCanvas.DrawMode = 13
   picCanvas.Line (0, 0)-Step(picCanvas.Width, picCanvas.Height), QBColor(15), BF
   arleft = 0: artop = 0: arright = 0: arbottom = 0 'Active rect coords
   ShowActRectCoords
   ActiveRectExists = False
   '------------------
   LabWait.Visible = False
End If
picCanvas.SetFocus
End Sub

Public Sub RotateText(zAngle, TextLine$)
'NB NOT ALL FONTS CAN BE ROTATED
'eg MS Sans Serif !!
'picCanvas CurrentX & Y set
'Set rotation in tenths of a degree, i.e., 1800 = 180 degrees
'TextLine$ = Text2.Text
'Make +ve angles rotate clockwise

RotateFont.lfEscapement = -zAngle * 10

'RotateFont.lfWidth = 10

If picCanvas.FontItalic Then
   RotateFont.lfItalic = 1
Else
   RotateFont.lfItalic = 0
End If

If picCanvas.FontStrikethru Then
   RotateFont.lfStrikeOut = 1
Else
   RotateFont.lfStrikeOut = 0
End If

If picCanvas.FontUnderline Then
   RotateFont.lfUnderline = 1
Else
   RotateFont.lfUnderline = 0
End If

RotateFont.lfCharSet = 0   '0,1

RotateFont.lfFaceName = picCanvas.FontName & Chr$(0)
   
If picCanvas.FontBold Then
   RotateFont.lfWeight = 1200 / 2
Else
   RotateFont.lfWeight = 400 / 2
End If
'stppy = Screen.TwipsPerPixelY
'rlf = (FontSize * 8)   '1-20
stppy = 5
rlf = picCanvas.FontSize * 8
RotateFont.lfHeight = rlf / stppy

'------------------------------------
rFont = CreateFontIndirect(RotateFont)
CurFont = SelectObject(picCanvas.hdc, rFont)
   
'X,Y = picMCRBox Left & Right
'picCanvas.CurrentX = X
'picCanvas.CurrentY = Y
'picCanvas.FontTransparent = True

picCanvas.Print TextLine$

'picCanvas.FontTransparent = False
   
'Restore CurFont
res = SelectObject(picCanvas.hdc, CurFont)
res = DeleteObject(rFont)
'------------------------------------
 
End Sub

Private Sub VScroll1_Change()
picCanvas.Top = VScroll1.Value
End Sub
Private Sub cmdCanHt_Click(Index As Integer)
'Set picCanvas Height
If Index = 0 Then
   picCanvas.Top = 4
   picCanvas.Height = 520
   picCanvasStore.Height = picCanvas.Height
   VScroll1.Value = 4
   VScroll1.Visible = False
Else
   picCanvas.Top = 4
   picCanvas.Height = 720
   picCanvasStore.Height = picCanvas.Height
   VScroll1.Value = 4
   VScroll1.Visible = True
End If

DrawCanvasBorder

End Sub

Private Sub SETUPHELP()
'Set up for RRPencil Help
With List1
   .Top = 4
   .Left = 190
   .Width = 400
   .Height = 500
End With
c$ = App.Path + "/RRPencil.txt"
On Error GoTo helperror
HelpExist = True
Open c$ For Input As #1
Do Until EOF(1)
   Line Input #1, a$
   List1.AddItem a$
Loop
Close
cmdHome(0).Top = List1.Top + 2
cmdHome(1).Top = List1.Top + 2
cmdHome(0).Left = List1.Left + List1.Width - cmdHome(0).Width - 20
cmdHome(1).Left = cmdHome(0).Left - cmdHome(0).Width - 20
GoTo afterhelp
helperror:
Close
HelpExist = False
GoTo afterhelp
afterhelp:
cmdHome(0).ZOrder 0
cmdHome(1).ZOrder 0
ClearHelpStuff
'-------------------------
End Sub

Private Sub cmdHome_Click(Index As Integer)
'Help buttons
If Index = 0 Then 'Home button
   List1.Visible = False
   'List1.ListIndex = 0
   List1.TopIndex = 0
   List1.Visible = True
Else  'Close button
   ClearHelpStuff
End If
End Sub

Private Sub List1_Click()
'Select item
I = List1.ListIndex
Text$ = List1.List(I)
If Left(Text$, 1) = "#" Then
   Text$ = Mid$(Text$, 2)

   'Search List1 for Text$ & place at top
   List1.ListIndex = SendMessage(List1.hwnd, LB_FINDSTRING, -1, ByVal Text$)
   If List1.ListIndex > 0 Then
      List1.TopIndex = List1.ListIndex - 1
   End If
   'For j = i + 1 To List1.ListCount - 1
   '   If List1.List(j) = Text$ Then Exit For
   'Next j
   'If j <= List1.ListCount - 1 Then
   '  List1.TopIndex = j - 1
   'End If
End If
End Sub
Private Sub List1_DblClick()
ClearHelpStuff
End Sub

Private Sub PencilHelp_Click()
If ActionInProgress Then Exit Sub

If HelpExist Then
   List1.ListIndex = 0
   List1.Visible = True
   cmdHome(0).Visible = True
   cmdHome(1).Visible = True
Else
   resp = RRMsgBox$("RRPencil", "See RRPencil.txt", 0, 0, 1, 1) 'YES,NO,OK, NumLines
End If
'c$ = "Notepad.exe " + App.Path + "/BRRICTEXT.txt"
'res& = Shell(c$, 1)
End Sub

Private Sub ClearHelpStuff()
   List1.Visible = False
   cmdHome(0).Visible = False
   cmdHome(1).Visible = False
End Sub

