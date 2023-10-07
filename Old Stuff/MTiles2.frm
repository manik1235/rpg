VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form MTiles 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   9525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   741
      Appearance      =   1
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons = 10
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7}
            Caption = ""
            Key = "Players"
            Description = ""
            Object.ToolTipText = "Player Tiles"
            Object.Tag = ""
            ImageIndex = 3
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7}
            Caption = ""
            Key = "Monsters"
            Description = ""
            Object.ToolTipText = "Monster Tiles"
            Object.Tag = ""
            ImageIndex = 2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7}
            Caption = ""
            Key = "Desert"
            Description = ""
            Object.ToolTipText = "Desert Tiles"
            Object.Tag = ""
            ImageIndex = 7
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7}
            Caption = ""
            Key = "Grass"
            Description = ""
            Object.ToolTipText = "Grass Tiles"
            Object.Tag = ""
            ImageIndex = 1
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7}
            Caption = ""
            Key = "Mountain"
            Description = ""
            Object.ToolTipText = "Mountain Tiles"
            Object.Tag = ""
            ImageIndex = 9
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7}
            Caption = ""
            Key = "PForest"
            Description = ""
            Object.ToolTipText = "Passable Forest Tiles"
            Object.Tag = ""
            ImageIndex = 10
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7}
            Caption = ""
            Key = "IForest"
            Description = ""
            Object.ToolTipText = "Impassible Forest Tiles"
            Object.Tag = ""
            ImageIndex = 5
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7}
            Caption = ""
            Key = ""
            Description = ""
            Object.ToolTipText = ""
            Object.Tag = ""
            Style = 3
            MixedState = -1         'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7}
            Caption = ""
            Key = "Weapon"
            Description = ""
            Object.ToolTipText = "Weapon Tiles"
            Object.Tag = ""
            ImageIndex = 4
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7}
            Caption = ""
            Key = "Treasures"
            Description = ""
            Object.ToolTipText = "Treasure Tiles"
            Object.Tag = ""
            ImageIndex = 6
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox TreasureTiles
      Height = 1815
      Left = 0
      ScaleHeight = 1755
      ScaleWidth = 1755
      TabIndex = 57
      Top = 960
      Width = 1815
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 35
         Left = 0
         Picture         =   "METiles.frx":0000
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 58
         Top = 0
         Width = 480
      End
   End
   Begin VB.PictureBox WeaponTiles
      Height = 1695
      Left = 3840
      ScaleHeight = 1635
      ScaleWidth = 1755
      TabIndex = 54
      Top = 960
      Width = 1815
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 33
         Left = 600
         Picture         =   "METiles.frx":030A
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 56
         Top = 600
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 34
         Left = 1200
         Picture         =   "METiles.frx":058C
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 55
         Top = 0
         Width = 480
      End
   End
   Begin VB.PictureBox MountainTiles
      Height = 2415
      Left = 0
      ScaleHeight = 2355
      ScaleWidth = 1755
      TabIndex = 43
      Top = 2880
      Width = 1815
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 27
         Left = 600
         Picture         =   "METiles.frx":0896
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 53
         Top = 0
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 17
         Left = 1200
         Picture         =   "METiles.frx":14D8
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 52
         Top = 1200
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 16
         Left = 1200
         Picture         =   "METiles.frx":17E2
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 51
         Top = 600
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 15
         Left = 0
         Picture         =   "METiles.frx":1AEC
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 50
         Top = 600
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 14
         Left = 0
         Picture         =   "METiles.frx":1DF6
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 49
         Top = 1200
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 12
         Left = 600
         Picture         =   "METiles.frx":2100
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 48
         Top = 600
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 13
         Left = 600
         Picture         =   "METiles.frx":240A
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 47
         Top = 1200
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 11
         Left = 1200
         Picture         =   "METiles.frx":2714
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 46
         Top = 1800
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 10
         Left = 1200
         Picture         =   "METiles.frx":2A1E
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 45
         Top = 0
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 2
         Left = 0
         Picture         =   "METiles.frx":3660
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 44
         Top = 0
         Width = 480
      End
   End
   Begin VB.PictureBox ForestTiles
      Height = 1815
      Left = 3840
      ScaleHeight = 1755
      ScaleWidth = 1755
      TabIndex = 40
      Top = 2880
      Width = 1815
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 7
         Left = 600
         Picture         =   "METiles.frx":396A
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 42
         Top = 0
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 1
         Left = 0
         Picture         =   "METiles.frx":45AC
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 41
         Top = 0
         Width = 480
      End
   End
   Begin VB.PictureBox WaterTiles
      Height = 1815
      Left = 1920
      ScaleHeight = 1755
      ScaleWidth = 1755
      TabIndex = 30
      Top = 960
      Width = 1815
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 4
         Left = 0
         Picture         =   "METiles.frx":48B6
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 38
         Top = 0
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 5
         Left = 600
         Picture         =   "METiles.frx":4B38
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 37
         Top = 0
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 6
         Left = 600
         Picture         =   "METiles.frx":4DBA
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 36
         Top = 1200
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 9
         Left = 600
         Picture         =   "METiles.frx":50C4
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 35
         Top = 600
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 18
         Left = 1200
         Picture         =   "METiles.frx":5D06
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 34
         Top = 0
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 19
         Left = 1200
         Picture         =   "METiles.frx":6948
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 33
         Top = 600
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 20
         Left = 0
         Picture         =   "METiles.frx":758A
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 32
         Top = 600
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 21
         Left = 0
         Picture         =   "METiles.frx":81CC
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 31
         Top = 1200
         Width = 480
      End
   End
   Begin VB.PictureBox GrassTiles
      Height = 7215
      Left = 1920
      ScaleHeight = 7155
      ScaleWidth = 1755
      TabIndex = 3
      Top = 2880
      Width = 1815
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 58
         Left = 600
         Picture         =   "METiles.frx":8E0E
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 66
         Top = 6000
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 57
         Left = 1200
         Picture         =   "METiles.frx":9A50
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 65
         Top = 6600
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 56
         Left = 0
         Picture = "METiles.frx": A692
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 64
         Top = 6000
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 55
         Left = 1200
         Picture = "METiles.frx": B2D4
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 63
         Top = 4800
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 54
         Left = 1200
         Picture = "METiles.frx": BF16
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 62
         Top = 6000
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 53
         Left = 600
         Picture = "METiles.frx": CB58
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 61
         Top = 5400
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 52
         Left = 1200
         Picture = "METiles.frx": D79A
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 60
         Top = 5400
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 51
         Left = 0
         Picture = "METiles.frx": E3DC
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 59
         Top = 5400
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 3
         Left = 0
         Picture = "METiles.frx": F01E
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 29
         Top = 0
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 22
         Left = 600
         Picture = "METiles.frx": F328
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 28
         Top = 1800
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 23
         Left = 600
         Picture = "METiles.frx": FF6A
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 27
         Top = 600
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 24
         Left = 600
         Picture         =   "METiles.frx":10BAC
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 26
         Top = 1200
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 26
         Left = 600
         Picture         =   "METiles.frx":117EE
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 25
         Top = 0
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 28
         Left = 1200
         Picture         =   "METiles.frx":12430
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 24
         Top = 600
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 29
         Left = 1200
         Picture         =   "METiles.frx":13072
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 23
         Top = 1200
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 30
         Left = 0
         Picture         =   "METiles.frx":13CB4
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 22
         Top = 600
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 31
         Left = 0
         Picture         =   "METiles.frx":148F6
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 21
         Top = 1200
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 32
         Left = 0
         Picture         =   "METiles.frx":15538
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 20
         Top = 1800
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 36
         Left = 1200
         Picture         =   "METiles.frx":1617A
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 19
         Top = 1800
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 8
         Left = 1200
         Picture         =   "METiles.frx":16DBC
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 18
         Top = 0
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 37
         Left = 600
         Picture         =   "METiles.frx":170C6
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 17
         Top = 2400
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 38
         Left = 0
         Picture         =   "METiles.frx":17D08
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 16
         Top = 4200
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 39
         Left = 1200
         Picture         =   "METiles.frx":1894A
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 15
         Top = 2400
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 40
         Left = 1200
         Picture         =   "METiles.frx":1958C
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 14
         Top = 3000
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 41
         Left = 0
         Picture         =   "METiles.frx":1A1CE
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 13
         Top = 2400
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 42
         Left = 0
         Picture         =   "METiles.frx":1AE10
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 12
         Top = 3000
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 43
         Left = 0
         Picture         =   "METiles.frx":1BA52
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 11
         Top = 4800
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 44
         Left = 600
         Picture         =   "METiles.frx":1C694
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 10
         Top = 4200
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 45
         Left = 600
         Picture         =   "METiles.frx":1D2D6
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 9
         Top = 4800
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 46
         Left = 1200
         Picture         =   "METiles.frx":1DF18
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 8
         Top = 4200
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 47
         Left = 600
         Picture         =   "METiles.frx":1EB5A
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 7
         Top = 3000
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 48
         Left = 1200
         Picture         =   "METiles.frx":1F79C
         ScaleHeight = 480
         ScaleWidth = 495
         TabIndex = 6
         Top = 3600
         Width = 495
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 49
         Left = 0
         Picture         =   "METiles.frx":2045E
         ScaleHeight = 480
         ScaleWidth = 495
         TabIndex = 5
         Top = 3600
         Width = 495
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 50
         Left = 600
         Picture         =   "METiles.frx":21120
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 4
         Top = 3600
         Width = 480
      End
   End
   Begin VB.PictureBox DesertTiles
      Height = 3015
      Left = 5760
      ScaleHeight = 2955
      ScaleWidth = 1755
      TabIndex = 0
      Top = 2880
      Width = 1815
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 66
         Left = 600
         Picture         =   "METiles.frx":21D62
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 74
         Top = 1200
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 65
         Left = 1200
         Picture         =   "METiles.frx":229A4
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 73
         Top = 1800
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 64
         Left = 0
         Picture         =   "METiles.frx":235E6
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 72
         Top = 1200
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 63
         Left = 1200
         Picture         =   "METiles.frx":24228
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 71
         Top = 0
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 62
         Left = 1200
         Picture         =   "METiles.frx":24E6A
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 70
         Top = 1200
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 61
         Left = 600
         Picture         =   "METiles.frx":25AAC
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 69
         Top = 600
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 60
         Left = 1200
         Picture         =   "METiles.frx":266EE
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 68
         Top = 600
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 59
         Left = 0
         Picture         =   "METiles.frx":27330
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 67
         Top = 600
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 0
         Left = 0
         Picture         =   "METiles.frx":27F72
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 2
         Top = 0
         Width = 480
      End
      Begin VB.PictureBox Terrain
         Appearance = 0         'Flat
         AutoRedraw = -1         'True
         AutoSize = -1           'True
         BackColor = &H80000005
         BorderStyle = 0        'None
         ForeColor = &H80000008
         Height = 480
         Index = 25
         Left = 600
         Picture         =   "METiles.frx":281F4
         ScaleHeight = 480
         ScaleWidth = 480
         TabIndex = 1
         Top = 0
         Width = 480
      End
   End
   Begin ComctlLib.ImageList ImageList1
      Left = 8760
      Top = 840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor = -2147483643
      ImageWidth = 32
      ImageHeight = 32
      MaskColor = 12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7}
         NumListImages = 10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7}
            Picture         =   "METiles.frx":28E36
            Key = "Grass"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7}
            Picture         =   "METiles.frx":29150
            Key = ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7}
            Picture         =   "METiles.frx":2946A
            Key = ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7}
            Picture         =   "METiles.frx":29784
            Key = ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7}
            Picture         =   "METiles.frx":29A9E
            Key = ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7}
            Picture         =   "METiles.frx":2A6F0
            Key = ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7}
            Picture         =   "METiles.frx":2AA0A
            Key = "Desert"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7}
            Picture         =   "METiles.frx":2AC9C
            Key = "Water"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7}
            Picture         =   "METiles.frx":2AF2E
            Key = "Mountains"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7}
            Picture         =   "METiles.frx":2B248
            Key = "Forest"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "MTiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
