VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.ocx"
Begin VB.Form Properties 
   Caption         =   "Properties"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   5640
      TabIndex        =   66
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3240
      TabIndex        =   65
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   64
      Top             =   5880
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   10186
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Map"
      TabPicture(0)   =   "meProperties.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(5)=   "lblMapSize"
      Tab(0).Control(6)=   "lblNumPlayers"
      Tab(0).Control(7)=   "lblNumMonsters"
      Tab(0).Control(8)=   "MapName"
      Tab(0).Control(9)=   "MapDesc"
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Player"
      TabPicture(1)   =   "meProperties.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Monster"
      TabPicture(2)   =   "meProperties.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(1)=   "Frame3"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Events"
      TabPicture(3)   =   "meProperties.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      Begin VB.Frame Frame4 
         Caption         =   "Monster Properties"
         Height          =   4455
         Left            =   -74880
         TabIndex        =   33
         Top             =   1140
         Width           =   6615
         Begin VB.CommandButton btnMonsterRight 
            Caption         =   ">"
            Height          =   255
            Left            =   2760
            TabIndex        =   58
            Top             =   3120
            Width           =   255
         End
         Begin VB.CommandButton btnMonsterLeft 
            Caption         =   "<"
            Height          =   255
            Left            =   2400
            TabIndex        =   57
            Top             =   3120
            Width           =   255
         End
         Begin VB.CommandButton btnSave 
            Caption         =   "&Save"
            Height          =   615
            Index           =   1
            Left            =   2760
            TabIndex        =   56
            Top             =   3600
            Width           =   1215
         End
         Begin VB.CommandButton btnRedefine 
            Caption         =   "&Redefine"
            Height          =   615
            Left            =   5160
            TabIndex        =   54
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox YMax 
            Height          =   285
            Left            =   4440
            TabIndex        =   52
            Text            =   "YMax"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox XMax 
            Height          =   285
            Left            =   4440
            TabIndex        =   51
            Text            =   "XMax"
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox YMin 
            Height          =   285
            Left            =   3600
            TabIndex        =   50
            Text            =   "YMin"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox XMin 
            Height          =   285
            Left            =   3600
            TabIndex        =   49
            Text            =   "XMin"
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox MonsterX 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1080
            TabIndex        =   38
            Text            =   "MonsterX"
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox MonsterY 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1080
            TabIndex        =   37
            Text            =   "MonsterY"
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox MonsterWeapon 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1080
            TabIndex        =   36
            Text            =   "MonsterWeapon"
            Top             =   1440
            Width           =   1455
         End
         Begin VB.TextBox MonsterText 
            Height          =   285
            Left            =   600
            TabIndex        =   35
            Text            =   "MonsterText"
            Top             =   2160
            Width           =   4455
         End
         Begin VB.TextBox MonsterHealth 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1080
            TabIndex        =   34
            Text            =   "MonsterHealth"
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label23 
            Caption         =   "Change Image"
            Height          =   255
            Index           =   0
            Left            =   2280
            TabIndex        =   59
            Top             =   2880
            Width           =   1095
         End
         Begin VB.Label Label22 
            Caption         =   "to      to"
            Height          =   615
            Left            =   4200
            TabIndex        =   53
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label21 
            Caption         =   "YRange"
            Height          =   255
            Left            =   2640
            TabIndex        =   48
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label20 
            Caption         =   "XRange"
            Height          =   255
            Left            =   2640
            TabIndex        =   47
            Top             =   360
            Width           =   735
         End
         Begin VB.Image MonsterImage 
            Height          =   375
            Left            =   1440
            Top             =   2880
            Width           =   495
         End
         Begin VB.Label Label19 
            Caption         =   "Monster's Image"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   2880
            Width           =   1215
         End
         Begin VB.Label Label17 
            Caption         =   "X Position"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label16 
            Caption         =   "Y Position"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label15 
            Caption         =   "Weapon"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label14 
            Caption         =   "Text to display when encountered"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   1800
            Width           =   2655
         End
         Begin VB.Label Label13 
            Caption         =   "Hint: Use | to separate text into different windows (SHIFT + \)"
            Height          =   255
            Left            =   720
            TabIndex        =   40
            Top             =   2520
            Width           =   4455
         End
         Begin VB.Label Label12 
            Caption         =   "Health"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   1080
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Monster"
         Height          =   615
         Left            =   -74880
         TabIndex        =   32
         Top             =   480
         Width           =   2175
         Begin VB.ComboBox cmboMonster 
            Height          =   315
            Left            =   120
            TabIndex        =   63
            Text            =   "0"
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Player Properties"
         Height          =   4455
         Left            =   120
         TabIndex        =   20
         Top             =   1140
         Width           =   6615
         Begin VB.CommandButton btnPlayerRight 
            Caption         =   ">"
            Height          =   255
            Left            =   2760
            TabIndex        =   61
            Top             =   3120
            Width           =   255
         End
         Begin VB.CommandButton btnPlayerLeft 
            Caption         =   "<"
            Height          =   255
            Left            =   2400
            TabIndex        =   60
            Top             =   3120
            Width           =   255
         End
         Begin VB.CommandButton btnSave 
            Caption         =   "&Save"
            Height          =   615
            Index           =   0
            Left            =   2760
            TabIndex        =   55
            Top             =   3600
            Width           =   1215
         End
         Begin VB.TextBox PlayerHealth 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1080
            TabIndex        =   31
            Text            =   "PlayerHealth"
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox PlayerText 
            Height          =   285
            Left            =   600
            TabIndex        =   28
            Text            =   "PlayerText"
            Top             =   2160
            Width           =   4455
         End
         Begin VB.TextBox PlayerWeapon 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1080
            TabIndex        =   26
            Text            =   "PlayerWeapon"
            Top             =   1440
            Width           =   1455
         End
         Begin VB.TextBox PlayerY 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1080
            TabIndex        =   24
            Text            =   "PlayerY"
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox PlayerX 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1080
            TabIndex        =   23
            Text            =   "PlayerX"
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label23 
            Caption         =   "Change Image"
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   62
            Top             =   2880
            Width           =   1095
         End
         Begin VB.Image PlayerImage 
            Height          =   375
            Left            =   1440
            Top             =   2880
            Width           =   495
         End
         Begin VB.Label Label18 
            Caption         =   "Player's Image"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   2880
            Width           =   1095
         End
         Begin VB.Label Label11 
            Caption         =   "Health"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label10 
            Caption         =   "Hint: Use | to separate text into different windows (SHIFT + \)"
            Height          =   255
            Left            =   720
            TabIndex        =   29
            Top             =   2520
            Width           =   4455
         End
         Begin VB.Label Label9 
            Caption         =   "Text to display at start (for this player)"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   1800
            Width           =   2655
         End
         Begin VB.Label Label8 
            Caption         =   "Weapon"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Y Position"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "X Position"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Player"
         Height          =   615
         Left            =   120
         TabIndex        =   11
         Top             =   420
         Width           =   4215
         Begin VB.OptionButton POption 
            Caption         =   "&7"
            Height          =   255
            Index           =   7
            Left            =   3480
            TabIndex        =   19
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton POption 
            Caption         =   "&6"
            Height          =   255
            Index           =   6
            Left            =   3000
            TabIndex        =   18
            Top             =   240
            Width           =   495
         End
         Begin VB.OptionButton POption 
            Caption         =   "&5"
            Height          =   255
            Index           =   5
            Left            =   2520
            TabIndex        =   17
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton POption 
            Caption         =   "&4"
            Height          =   255
            Index           =   4
            Left            =   2040
            TabIndex        =   16
            Top             =   240
            Width           =   495
         End
         Begin VB.OptionButton POption 
            Caption         =   "&3"
            Height          =   255
            Index           =   3
            Left            =   1560
            TabIndex        =   15
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton POption 
            Caption         =   "&2"
            Height          =   255
            Index           =   2
            Left            =   1080
            TabIndex        =   14
            Top             =   240
            Width           =   495
         End
         Begin VB.OptionButton POption 
            Caption         =   "&1"
            Height          =   255
            Index           =   1
            Left            =   600
            TabIndex        =   13
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton POption 
            Caption         =   "&0"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Value           =   -1  'True
            Width           =   495
         End
      End
      Begin VB.TextBox MapDesc 
         Height          =   285
         Left            =   -73560
         TabIndex        =   7
         Text            =   "MapDesc"
         Top             =   900
         Width           =   3735
      End
      Begin VB.TextBox MapName 
         Height          =   285
         Left            =   -73560
         TabIndex        =   1
         Text            =   "MapName"
         Top             =   540
         Width           =   3735
      End
      Begin VB.Label lblNumMonsters 
         Alignment       =   1  'Right Justify
         Caption         =   "lblNumMonsters"
         Height          =   255
         Left            =   -72480
         TabIndex        =   10
         Top             =   1980
         Width           =   2535
      End
      Begin VB.Label lblNumPlayers 
         Alignment       =   1  'Right Justify
         Caption         =   "lblNumPlayers"
         Height          =   255
         Left            =   -72600
         TabIndex        =   9
         Top             =   1620
         Width           =   2655
      End
      Begin VB.Label lblMapSize 
         Alignment       =   1  'Right Justify
         Caption         =   "lblMapSize"
         Height          =   255
         Left            =   -72840
         TabIndex        =   8
         Top             =   1260
         Width           =   2895
      End
      Begin VB.Label Label5 
         Caption         =   "Number of Monsters"
         Height          =   255
         Left            =   -74880
         TabIndex        =   6
         Top             =   1980
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Number of Players"
         Height          =   255
         Left            =   -74880
         TabIndex        =   5
         Top             =   1620
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Description"
         Height          =   255
         Left            =   -74880
         TabIndex        =   4
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Size"
         Height          =   255
         Left            =   -74880
         TabIndex        =   3
         Top             =   1260
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Map Name"
         Height          =   255
         Left            =   -74880
         TabIndex        =   2
         Top             =   540
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Properties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btnApply_Click()
On Error GoTo errh
' Apply will only save what is on the current tab
If SSTab1.Tab = 0 Then
  'Map Properties
  MapEdit.MapName = Properties.MapName
  MapEdit.MapDesc = Properties.MapDesc
  MapEdit.Form_Resize
ElseIf SSTab1.Tab = 1 Then
  'Player Properties
End If

Exit Sub
errh:
MsgBox Err & Err.Description
Stop
Resume
End Sub

Private Sub btnMonsterLeft_Click()
MapEdit.btnMonsterLeft_Click_
End Sub

Private Sub btnMonsterRight_Click()
MapEdit.btnMonsterRight_Click_
End Sub

Private Sub btnPlayerLeft_Click()
MapEdit.btnPlayerLeft_Click_
End Sub

Private Sub btnPlayerRight_Click()
MapEdit.btnPlayerRight_Click_
End Sub

Private Sub btnRedefine_Click()
MapEdit.btnRedefine_Click_
End Sub

Private Sub btnSave_Click(Index As Integer)
MapEdit.btnSave_Click_ Index
End Sub


Private Sub Form_Load()
Me.Left = (Screen.Width - Me.ScaleWidth) / 2
Me.Top = (Screen.Height - Me.ScaleHeight) / 2
MapEdit.Form_Load_
End Sub

Private Sub cmboMonster_Click()
MapEdit.cmboMonster_Click cmboMonster.Text
End Sub

Private Sub POption_Click(Index As Integer)
    MapEdit.POption_Click_ Index
End Sub

