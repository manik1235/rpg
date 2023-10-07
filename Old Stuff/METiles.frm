VERSION 5.00
Begin VB.Form Tiles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tools"
   ClientHeight    =   9645
   ClientLeft      =   8880
   ClientTop       =   330
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   10665
   Begin VB.PictureBox ToolToolbar1 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   2655
      TabIndex        =   161
      Top             =   0
      Width           =   2655
      Begin VB.Image Button 
         Height          =   360
         Index           =   1
         Left            =   120
         Picture         =   "METiles.frx":0000
         Stretch         =   -1  'True
         ToolTipText     =   "Players"
         Top             =   120
         Width           =   360
      End
      Begin VB.Image Button 
         Height          =   360
         Index           =   2
         Left            =   120
         Picture         =   "METiles.frx":030A
         Stretch         =   -1  'True
         ToolTipText     =   "Monsters"
         Top             =   480
         Width           =   360
      End
      Begin VB.Image Button 
         Height          =   360
         Index           =   3
         Left            =   480
         Picture         =   "METiles.frx":0614
         Stretch         =   -1  'True
         ToolTipText     =   "Desert Tiles"
         Top             =   120
         Width           =   360
      End
      Begin VB.Image Button 
         Height          =   360
         Index           =   4
         Left            =   840
         Picture         =   "METiles.frx":0896
         Stretch         =   -1  'True
         ToolTipText     =   "Grass Tiles"
         Top             =   120
         Width           =   360
      End
      Begin VB.Image Button 
         Height          =   360
         Index           =   5
         Left            =   1200
         Picture         =   "METiles.frx":0BA0
         Stretch         =   -1  'True
         ToolTipText     =   "Mountain Tiles"
         Top             =   120
         Width           =   360
      End
      Begin VB.Image Button 
         Height          =   360
         Index           =   6
         Left            =   1560
         Picture         =   "METiles.frx":0EAA
         Stretch         =   -1  'True
         ToolTipText     =   "Passable Forest Tiles"
         Top             =   120
         Width           =   360
      End
      Begin VB.Image Button 
         Height          =   360
         Index           =   7
         Left            =   480
         Picture         =   "METiles.frx":11B4
         Stretch         =   -1  'True
         ToolTipText     =   "Impassible Forest Tiles"
         Top             =   480
         Width           =   360
      End
      Begin VB.Image Button 
         Height          =   360
         Index           =   8
         Left            =   840
         Picture         =   "METiles.frx":1DF6
         Stretch         =   -1  'True
         ToolTipText     =   "Water Tiles"
         Top             =   480
         Width           =   360
      End
      Begin VB.Image Button 
         Height          =   360
         Index           =   9
         Left            =   1200
         Picture         =   "METiles.frx":2078
         Stretch         =   -1  'True
         ToolTipText     =   "Snow Tiles"
         Top             =   480
         Width           =   360
      End
      Begin VB.Image Button 
         Height          =   360
         Index           =   10
         Left            =   2040
         Picture         =   "METiles.frx":2CBA
         Stretch         =   -1  'True
         ToolTipText     =   "Weapon Tiles"
         Top             =   120
         Width           =   360
      End
      Begin VB.Image Button 
         Height          =   360
         Index           =   11
         Left            =   2040
         Picture         =   "METiles.frx":2FC4
         Stretch         =   -1  'True
         ToolTipText     =   "Treasure Tiles"
         Top             =   480
         Width           =   360
      End
      Begin VB.Image ToolbarImage 
         Height          =   1005
         Left            =   0
         Picture         =   "METiles.frx":32CE
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2535
      End
   End
   Begin VB.VScrollBar ToolVScroll1 
      Height          =   3015
      LargeChange     =   1000
      Left            =   2040
      SmallChange     =   100
      TabIndex        =   80
      Top             =   960
      Width           =   255
   End
   Begin VB.PictureBox TileWindow 
      Height          =   1815
      Index           =   7
      Left            =   6960
      ScaleHeight     =   1755
      ScaleWidth      =   1755
      TabIndex        =   79
      Top             =   1200
      Width           =   1815
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   95
         Left            =   600
         Picture         =   "METiles.frx":AFF0
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   160
         Tag             =   "F72"
         Top             =   0
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   96
         Left            =   600
         Picture         =   "METiles.frx":BC32
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   159
         Tag             =   "F82"
         Top             =   600
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   113
         Left            =   600
         Picture         =   "METiles.frx":C874
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   158
         Tag             =   "F12"
         Top             =   1200
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   114
         Left            =   1200
         Picture         =   "METiles.frx":D4B6
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   157
         Tag             =   "F32"
         Top             =   0
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   115
         Left            =   1200
         Picture         =   "METiles.frx":E0F8
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   156
         Tag             =   "F42"
         Top             =   600
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   116
         Left            =   0
         Picture         =   "METiles.frx":ED3A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   155
         Tag             =   "F52"
         Top             =   0
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   117
         Left            =   0
         Picture         =   "METiles.frx":F97C
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   154
         Tag             =   "F62"
         Top             =   600
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   118
         Left            =   0
         Picture         =   "METiles.frx":105BE
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   153
         Tag             =   "F22"
         Top             =   1200
         Width           =   480
      End
   End
   Begin VB.PictureBox TileWindow 
      Height          =   615
      Index           =   2
      Left            =   120
      ScaleHeight     =   555
      ScaleWidth      =   1755
      TabIndex        =   74
      Top             =   1200
      Width           =   1815
      Begin VB.PictureBox Characters 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   1
         Left            =   0
         Picture         =   "METiles.frx":11200
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   78
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.PictureBox TileWindow 
      Height          =   615
      Index           =   1
      Left            =   120
      ScaleHeight     =   555
      ScaleWidth      =   1755
      TabIndex        =   73
      Top             =   1920
      Width           =   1815
      Begin VB.PictureBox Characters 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   3
         Left            =   1200
         Picture         =   "METiles.frx":1150A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   77
         Top             =   0
         Width           =   480
      End
      Begin VB.PictureBox Characters 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   2
         Left            =   600
         Picture         =   "METiles.frx":1178C
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   76
         Top             =   0
         Width           =   480
      End
      Begin VB.PictureBox Characters 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   0
         Left            =   0
         Picture         =   "METiles.frx":11A0E
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   75
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.PictureBox TileWindow 
      Height          =   615
      Index           =   11
      Left            =   120
      ScaleHeight     =   555
      ScaleWidth      =   1755
      TabIndex        =   25
      Top             =   2640
      Width           =   1815
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   35
         Left            =   0
         LinkItem        =   "Chalice"
         Picture         =   "METiles.frx":11D18
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   26
         Tag             =   "G0t"
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.PictureBox TileWindow 
      Height          =   1815
      Index           =   10
      Left            =   5040
      ScaleHeight     =   1755
      ScaleWidth      =   1755
      TabIndex        =   22
      Top             =   1200
      Width           =   1815
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   145
         Left            =   600
         Picture         =   "METiles.frx":12022
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   152
         Tag             =   "M0w"
         Top             =   1200
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   120
         Left            =   0
         Picture         =   "METiles.frx":12C64
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   151
         Tag             =   "M1w"
         Top             =   600
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   112
         Left            =   0
         Picture         =   "METiles.frx":138A6
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   125
         Tag             =   "F0w"
         Top             =   1200
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   93
         Left            =   600
         Picture         =   "METiles.frx":144E8
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   124
         Tag             =   "F1w"
         Top             =   600
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   86
         Left            =   1200
         Picture         =   "METiles.frx":1512A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   83
         Tag             =   "D0w"
         Top             =   600
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   67
         Left            =   0
         Picture         =   "METiles.frx":15D6C
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   81
         Tag             =   "G1w"
         Top             =   0
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   33
         Left            =   1200
         Picture         =   "METiles.frx":169AE
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   24
         Tag             =   "D1w"
         Top             =   0
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   34
         Left            =   600
         Picture         =   "METiles.frx":16C30
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   23
         Tag             =   "G0w"
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.PictureBox TileWindow 
      Height          =   7215
      Index           =   5
      Left            =   2640
      ScaleHeight     =   7155
      ScaleWidth      =   1755
      TabIndex        =   20
      Top             =   240
      Width           =   1815
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   144
         Left            =   0
         Picture         =   "METiles.frx":16F3A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   150
         Tag             =   "M22"
         Top             =   1800
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   143
         Left            =   0
         Picture         =   "METiles.frx":17B7C
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   149
         Tag             =   "M62"
         Top             =   1200
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   142
         Left            =   0
         Picture         =   "METiles.frx":187BE
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   148
         Tag             =   "M52"
         Top             =   600
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   141
         Left            =   1200
         Picture         =   "METiles.frx":19400
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   147
         Tag             =   "M42"
         Top             =   1200
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   140
         Left            =   1200
         Picture         =   "METiles.frx":1A042
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   146
         Tag             =   "M32"
         Top             =   600
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   139
         Left            =   600
         Picture         =   "METiles.frx":1AC84
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   145
         Tag             =   "M12"
         Top             =   1800
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   138
         Left            =   600
         Picture         =   "METiles.frx":1B8C6
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   144
         Tag             =   "Me1"
         Top             =   3600
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   137
         Left            =   0
         Picture         =   "METiles.frx":1C508
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   143
         Tag             =   "Md1"
         Top             =   3600
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   136
         Left            =   1200
         Picture         =   "METiles.frx":1D14A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   142
         Tag             =   "Mc1"
         Top             =   3600
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   135
         Left            =   600
         Picture         =   "METiles.frx":1DD8C
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   141
         Tag             =   "Mb1"
         Top             =   3000
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   134
         Left            =   1200
         Picture         =   "METiles.frx":1E9CE
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   140
         Tag             =   "Ma1"
         Top             =   4200
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   133
         Left            =   600
         Picture         =   "METiles.frx":1F610
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   139
         Tag             =   "M91"
         Top             =   4800
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   132
         Left            =   600
         Picture         =   "METiles.frx":20252
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   138
         Tag             =   "M81"
         Top             =   4200
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   131
         Left            =   0
         Picture         =   "METiles.frx":20E94
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   137
         Tag             =   "M71"
         Top             =   4800
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   130
         Left            =   0
         Picture         =   "METiles.frx":21AD6
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   136
         Tag             =   "M61"
         Top             =   3000
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   129
         Left            =   0
         Picture         =   "METiles.frx":22718
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   135
         Tag             =   "M51"
         Top             =   2400
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   128
         Left            =   1200
         Picture         =   "METiles.frx":2335A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   134
         Tag             =   "M41"
         Top             =   3000
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   127
         Left            =   1200
         Picture         =   "METiles.frx":23F9C
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   133
         Tag             =   "M31"
         Top             =   2400
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   126
         Left            =   0
         Picture         =   "METiles.frx":24BDE
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   132
         Tag             =   "M21"
         Top             =   4200
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   125
         Left            =   1200
         Picture         =   "METiles.frx":25820
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   131
         Top             =   1800
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   124
         Left            =   600
         Picture         =   "METiles.frx":26462
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   130
         Tag             =   "Mf1"
         Top             =   2400
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   123
         Left            =   1200
         Picture         =   "METiles.frx":270A4
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   129
         Tag             =   "M20"
         Top             =   0
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   122
         Left            =   600
         Picture         =   "METiles.frx":27CE6
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   128
         Tag             =   "M82"
         Top             =   1200
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   121
         Left            =   600
         Picture         =   "METiles.frx":28928
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   127
         Tag             =   "M72"
         Top             =   600
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   2
         Left            =   0
         Picture         =   "METiles.frx":2956A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   43
         Tag             =   "M00"
         Top             =   0
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   17
         Left            =   1200
         Picture         =   "METiles.frx":29874
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   42
         Tag             =   "Mm1"
         Top             =   6000
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   16
         Left            =   1200
         Picture         =   "METiles.frx":29B7E
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   41
         Tag             =   "Ml1"
         Top             =   5400
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   15
         Left            =   0
         Picture         =   "METiles.frx":29E88
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   40
         Tag             =   "Mk1"
         Top             =   5400
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   14
         Left            =   0
         Picture         =   "METiles.frx":2A192
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   39
         Tag             =   "Mj1"
         Top             =   6000
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   12
         Left            =   600
         Picture         =   "METiles.frx":2A49C
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   38
         Tag             =   "Mh1"
         Top             =   5400
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   13
         Left            =   600
         Picture         =   "METiles.frx":2A7A6
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   37
         Tag             =   "Mi1"
         Top             =   6000
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   11
         Left            =   1200
         Picture         =   "METiles.frx":2AAB0
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   36
         Tag             =   "Mg1"
         Top             =   6600
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   10
         Left            =   1200
         Picture         =   "METiles.frx":2ADBA
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   35
         Tag             =   "Mf1"
         Top             =   4800
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   27
         Left            =   600
         Picture         =   "METiles.frx":2B0C4
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   21
         Tag             =   "M10"
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.PictureBox TileWindow 
      Height          =   4215
      Index           =   6
      Left            =   6960
      Picture         =   "METiles.frx":2BD06
      ScaleHeight     =   4155
      ScaleWidth      =   1755
      TabIndex        =   19
      Top             =   3240
      Width           =   1815
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   119
         Left            =   0
         Picture         =   "METiles.frx":2C948
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   126
         Tag             =   "F61"
         Top             =   1200
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   111
         Left            =   1200
         Picture         =   "METiles.frx":2D58A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   123
         Tag             =   "f20"
         Top             =   3000
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   110
         Left            =   0
         Picture         =   "METiles.frx":2E1CC
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   122
         Tag             =   "F21"
         Top             =   2400
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   109
         Left            =   600
         Picture         =   "METiles.frx":2EE0E
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   121
         Tag             =   "Fe1"
         Top             =   1800
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   108
         Left            =   0
         Picture         =   "METiles.frx":2FA50
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   120
         Tag             =   "Fd1"
         Top             =   1800
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   107
         Left            =   1200
         Picture         =   "METiles.frx":30692
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   119
         Tag             =   "Fc1"
         Top             =   1800
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   106
         Left            =   600
         Picture         =   "METiles.frx":312D4
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   118
         Tag             =   "Fb1"
         Top             =   1200
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   105
         Left            =   0
         Picture         =   "METiles.frx":31F16
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   117
         Tag             =   "F51"
         Top             =   600
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   104
         Left            =   1200
         Picture         =   "METiles.frx":32B58
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   116
         Tag             =   "F41"
         Top             =   1200
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   103
         Left            =   1200
         Picture         =   "METiles.frx":3379A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   115
         Tag             =   "F31"
         Top             =   600
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   102
         Left            =   0
         Picture         =   "METiles.frx":343DC
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   114
         Tag             =   "F11"
         Top             =   3600
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   101
         Left            =   1200
         Picture         =   "METiles.frx":3501E
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   113
         Tag             =   "Fa1"
         Top             =   2400
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   100
         Left            =   600
         Picture         =   "METiles.frx":35C60
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   112
         Tag             =   "F91"
         Top             =   3000
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   99
         Left            =   600
         Picture         =   "METiles.frx":368A2
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   111
         Tag             =   "F81"
         Top             =   2400
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   98
         Left            =   0
         Picture         =   "METiles.frx":374E4
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   110
         Tag             =   "F71"
         Top             =   3000
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   97
         Left            =   600
         Picture         =   "METiles.frx":38126
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   109
         Tag             =   "F01"
         Top             =   600
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   94
         Left            =   1200
         Picture         =   "METiles.frx":38D68
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   108
         Tag             =   "F30"
         Top             =   0
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   7
         Left            =   600
         Picture         =   "METiles.frx":399AA
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   70
         Tag             =   "F10"
         Top             =   0
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   1
         Left            =   0
         Picture         =   "METiles.frx":3A5EC
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   69
         Tag             =   "F00"
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.PictureBox TileWindow 
      Height          =   1815
      Index           =   8
      Left            =   5040
      ScaleHeight     =   1755
      ScaleWidth      =   1755
      TabIndex        =   16
      Top             =   3120
      Width           =   1815
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   21
         Left            =   0
         Picture         =   "METiles.frx":3A8F6
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   68
         Tag             =   "W61"
         Top             =   1200
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   20
         Left            =   0
         Picture         =   "METiles.frx":3B538
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   67
         Tag             =   "W51"
         Top             =   600
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   19
         Left            =   1200
         Picture         =   "METiles.frx":3C17A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   66
         Tag             =   "W41"
         Top             =   1200
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   18
         Left            =   1200
         Picture         =   "METiles.frx":3CDBC
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   65
         Tag             =   "W31"
         Top             =   600
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   9
         Left            =   600
         Picture         =   "METiles.frx":3D9FE
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   64
         Tag             =   "W01"
         Top             =   600
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   6
         Left            =   600
         Picture         =   "METiles.frx":3E640
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   63
         Tag             =   "W11"
         Top             =   1200
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   5
         Left            =   600
         Picture         =   "METiles.frx":3E94A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   62
         Tag             =   "W21"
         Top             =   0
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   4
         Left            =   0
         Picture         =   "METiles.frx":3EBCC
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   17
         Tag             =   "W00"
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.PictureBox TileWindow 
      Height          =   7215
      Index           =   4
      Left            =   6120
      ScaleHeight     =   7155
      ScaleWidth      =   1755
      TabIndex        =   2
      Top             =   1320
      Width           =   1815
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   37
         Left            =   1200
         Picture         =   "METiles.frx":3F80E
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   84
         Tag             =   "G11"
         Top             =   1800
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   3
         Left            =   0
         Picture         =   "METiles.frx":40450
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   72
         Tag             =   "G00"
         Top             =   0
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   8
         Left            =   1200
         Picture         =   "METiles.frx":4075A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   71
         Tag             =   "G20"
         Top             =   0
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   26
         Left            =   600
         Picture         =   "METiles.frx":40A64
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   61
         Tag             =   "G10"
         Top             =   0
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   32
         Left            =   0
         Picture         =   "METiles.frx":416A6
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   60
         Tag             =   "G22"
         Top             =   1800
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   31
         Left            =   0
         Picture         =   "METiles.frx":422E8
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   59
         Tag             =   "G62"
         Top             =   1200
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   30
         Left            =   0
         Picture         =   "METiles.frx":42F2A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   58
         Tag             =   "G52"
         Top             =   600
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   29
         Left            =   1200
         Picture         =   "METiles.frx":43B6C
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   57
         Tag             =   "G42"
         Top             =   1200
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   28
         Left            =   1200
         Picture         =   "METiles.frx":447AE
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   56
         Tag             =   "G20"
         Top             =   600
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   24
         Left            =   600
         Picture         =   "METiles.frx":453F0
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   55
         Tag             =   "G82"
         Top             =   1200
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   23
         Left            =   600
         Picture         =   "METiles.frx":46032
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   54
         Tag             =   "G72"
         Top             =   600
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   22
         Left            =   600
         Picture         =   "METiles.frx":46C74
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   53
         Tag             =   "G12"
         Top             =   1800
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   36
         Left            =   600
         Picture         =   "METiles.frx":478B6
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   52
         Tag             =   "G01"
         Top             =   2400
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   58
         Left            =   600
         Picture         =   "METiles.frx":484F8
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   51
         Tag             =   "Gm1"
         Top             =   6000
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   57
         Left            =   1200
         Picture         =   "METiles.frx":4913A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   50
         Tag             =   "Gl1"
         Top             =   6600
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   56
         Left            =   0
         Picture         =   "METiles.frx":49D7C
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   49
         Tag             =   "Gk1"
         Top             =   6000
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   55
         Left            =   1200
         Picture         =   "METiles.frx":4A9BE
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   48
         Tag             =   "Gj1"
         Top             =   4800
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   54
         Left            =   1200
         Picture         =   "METiles.frx":4B600
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   47
         Tag             =   "Gi1"
         Top             =   6000
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   53
         Left            =   600
         Picture         =   "METiles.frx":4C242
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   46
         Tag             =   "Gh1"
         Top             =   5400
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   52
         Left            =   1200
         Picture         =   "METiles.frx":4CE84
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   45
         Tag             =   "Gg1"
         Top             =   5400
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   51
         Left            =   0
         Picture         =   "METiles.frx":4DAC6
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   44
         Tag             =   "Gf1"
         Top             =   5400
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   38
         Left            =   0
         Picture         =   "METiles.frx":4E708
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   15
         Tag             =   "G21"
         Top             =   4200
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   39
         Left            =   1200
         Picture         =   "METiles.frx":4F34A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   14
         Tag             =   "G31"
         Top             =   2400
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   40
         Left            =   1200
         Picture         =   "METiles.frx":4FF8C
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   13
         Tag             =   "G41"
         Top             =   3000
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   41
         Left            =   0
         Picture         =   "METiles.frx":50BCE
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   12
         Tag             =   "G51"
         Top             =   2400
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   42
         Left            =   0
         Picture         =   "METiles.frx":51810
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   11
         Tag             =   "G61"
         Top             =   3000
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   43
         Left            =   0
         Picture         =   "METiles.frx":52452
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   10
         Tag             =   "G71"
         Top             =   4800
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   44
         Left            =   600
         Picture         =   "METiles.frx":53094
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   9
         Tag             =   "G81"
         Top             =   4200
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   45
         Left            =   600
         Picture         =   "METiles.frx":53CD6
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   8
         Tag             =   "G91"
         Top             =   4800
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   46
         Left            =   1200
         Picture         =   "METiles.frx":54918
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   7
         Tag             =   "Ga1"
         Top             =   4200
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   47
         Left            =   600
         Picture         =   "METiles.frx":5555A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   6
         Tag             =   "Gb1"
         Top             =   3000
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   48
         Left            =   1200
         Picture         =   "METiles.frx":5619C
         ScaleHeight     =   480
         ScaleWidth      =   495
         TabIndex        =   5
         Tag             =   "Gc1"
         Top             =   3600
         Width           =   495
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   49
         Left            =   0
         Picture         =   "METiles.frx":56E5E
         ScaleHeight     =   480
         ScaleWidth      =   495
         TabIndex        =   4
         Tag             =   "Gd1"
         Top             =   3600
         Width           =   495
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   50
         Left            =   600
         Picture         =   "METiles.frx":57B20
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   3
         Tag             =   "Ge1"
         Top             =   3600
         Width           =   480
      End
   End
   Begin VB.PictureBox TileWindow 
      Height          =   7215
      Index           =   3
      Left            =   3240
      ScaleHeight     =   7155
      ScaleWidth      =   1755
      TabIndex        =   0
      Top             =   1320
      Width           =   1815
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   92
         Left            =   0
         Picture         =   "METiles.frx":58762
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   107
         Tag             =   "D22"
         Top             =   1800
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   91
         Left            =   0
         Picture         =   "METiles.frx":593A4
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   106
         Tag             =   "D62"
         Top             =   1200
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   90
         Left            =   0
         Picture         =   "METiles.frx":59FE6
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   105
         Tag             =   "D52"
         Top             =   600
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   89
         Left            =   1200
         Picture         =   "METiles.frx":5AC28
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   104
         Tag             =   "D42"
         Top             =   1200
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   88
         Left            =   1200
         Picture         =   "METiles.frx":5B86A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   103
         Tag             =   "D32"
         Top             =   600
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   87
         Left            =   600
         Picture         =   "METiles.frx":5C4AC
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   102
         Tag             =   "d12"
         Top             =   1800
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   85
         Left            =   1200
         Picture         =   "METiles.frx":5D0EE
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   101
         Tag             =   "D20"
         Top             =   0
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   84
         Left            =   0
         Picture         =   "METiles.frx":5DD30
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   100
         Tag             =   "D21"
         Top             =   4200
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   83
         Left            =   600
         Picture         =   "METiles.frx":5E972
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   99
         Tag             =   "De1"
         Top             =   3600
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   82
         Left            =   0
         Picture         =   "METiles.frx":5F5B4
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   98
         Tag             =   "Dd1"
         Top             =   3600
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   81
         Left            =   1200
         Picture         =   "METiles.frx":601F6
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   97
         Tag             =   "Dc1"
         Top             =   3600
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   80
         Left            =   600
         Picture         =   "METiles.frx":60E38
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   96
         Tag             =   "Db1"
         Top             =   3000
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   79
         Left            =   0
         Picture         =   "METiles.frx":61A7A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   95
         Tag             =   "D61"
         Top             =   3000
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   78
         Left            =   0
         Picture         =   "METiles.frx":626BC
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   94
         Tag             =   "D61"
         Top             =   2400
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   77
         Left            =   1200
         Picture         =   "METiles.frx":632FE
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   93
         Tag             =   "D41"
         Top             =   3000
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   76
         Left            =   1200
         Picture         =   "METiles.frx":63F40
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   92
         Tag             =   "D31"
         Top             =   2400
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   75
         Left            =   1200
         Picture         =   "METiles.frx":64B82
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   91
         Tag             =   "D11"
         Top             =   1800
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   74
         Left            =   1200
         Picture         =   "METiles.frx":657C4
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   90
         Tag             =   "Da1"
         Top             =   4200
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   73
         Left            =   600
         Picture         =   "METiles.frx":66406
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   89
         Tag             =   "D91"
         Top             =   4800
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   72
         Left            =   600
         Picture         =   "METiles.frx":67048
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   88
         Tag             =   "D81"
         Top             =   4200
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   71
         Left            =   0
         Picture         =   "METiles.frx":67C8A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   87
         Tag             =   "D71"
         Top             =   4800
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   70
         Left            =   600
         Picture         =   "METiles.frx":688CC
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   86
         Tag             =   "D01"
         Top             =   2400
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   69
         Left            =   600
         Picture         =   "METiles.frx":6950E
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   85
         Tag             =   "D82"
         Top             =   1200
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   68
         Left            =   600
         Picture         =   "METiles.frx":6A150
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   82
         Tag             =   "D72"
         Top             =   600
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   66
         Left            =   600
         Picture         =   "METiles.frx":6AD92
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   34
         Tag             =   "Dm1"
         Top             =   6000
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   65
         Left            =   1200
         Picture         =   "METiles.frx":6B9D4
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   33
         Tag             =   "Dl1"
         Top             =   6600
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   64
         Left            =   0
         Picture         =   "METiles.frx":6C616
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   32
         Tag             =   "Dk1"
         Top             =   6000
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   63
         Left            =   1200
         Picture         =   "METiles.frx":6D258
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   31
         Tag             =   "Dj1"
         Top             =   4800
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   62
         Left            =   1200
         Picture         =   "METiles.frx":6DE9A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   30
         Tag             =   "Di1"
         Top             =   6000
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   61
         Left            =   600
         Picture         =   "METiles.frx":6EADC
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   29
         Tag             =   "Dh1"
         Top             =   5400
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   60
         Left            =   1200
         Picture         =   "METiles.frx":6F71E
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   28
         Tag             =   "Dg1"
         Top             =   5400
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   59
         Left            =   0
         Picture         =   "METiles.frx":70360
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   27
         Tag             =   "Df1"
         Top             =   5400
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   0
         Left            =   0
         Picture         =   "METiles.frx":70FA2
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   1
         Tag             =   "D00"
         Top             =   0
         Width           =   480
      End
      Begin VB.PictureBox Terrain 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   25
         Left            =   600
         Picture         =   "METiles.frx":71224
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   18
         Tag             =   "D10"
         Top             =   0
         Width           =   480
      End
   End
End
Attribute VB_Name = "Tiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim w As Integer
Dim curtool As Integer
Public MaxTerrain As Integer

Private Sub Button_Click(Index As Integer)
On Error Resume Next
Toolbar1.ZOrder
    curtool = Index
    For w = 1 To 11
      TileWindow(w).Visible = False
    Next w
    TileWindow(Index).Visible = True
    SetTiles curtool
End Sub

Private Sub Form_Load()
On Error GoTo errh
ToolToolbar1.Height = ToolbarImage.Height
ToolToolbar1.Width = ToolbarImage.Width
MaxTerrain = 0
1 Me.Caption = Terrain(MaxTerrain).Tag
MaxTerrain = MaxTerrain + 1
GoTo 1
errh:
MaxTerrain = MaxTerrain - 1
curtool = 1
SetTiles curtool
Me.Caption = "Tools"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
Me.Hide
End Sub

Private Sub Terrain_Click(Index As Integer)
On Error GoTo errh
    MapEdit.CurTile.Picture = Terrain(Index).Picture
    MapEdit.CurTile.Tag = Terrain(Index).Tag
    MapEdit.Icon = Terrain(Index).Picture
Exit Sub
errh:
Resume Next
End Sub


Private Sub VScroll1_Change()
    TileWindow(curtool).Top = ((VScroll1.Value) + Toolbar1.Height)
End Sub

Public Sub SetTiles(curtool)
    For w = 1 To 11
      If w <> 9 Then TileWindow(w).Visible = False
    Next w
    TileWindow(curtool).Visible = True
    TileWindow(curtool).Left = 0
    TileWindow(curtool).Top = 0 + ToolToolbar1.Height
    ToolVScroll1.Max = 0 - (TileWindow(curtool).Height - ToolVScroll1.Height)
    ToolVScroll1.Left = 1920
    'VScroll1.Max = TileWindow(curtool).Height
End Sub

Private Sub VScroll1_Scroll()
VScroll1_Change
End Sub
