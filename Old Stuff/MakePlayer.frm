VERSION 5.00
Begin VB.Form MakePlayer 
   Caption         =   "Charater Properties"
   ClientHeight    =   2670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7605
   Icon            =   "MakePlayer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2670
   ScaleWidth      =   7605
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set Range First"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   6120
      TabIndex        =   6
      Top             =   480
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Creature Stats"
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.CheckBox Check1 
         Caption         =   "Monster doesn't move"
         Height          =   255
         Left            =   3360
         TabIndex        =   13
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox CharMsgText 
         Height          =   285
         Left            =   1800
         TabIndex        =   11
         Top             =   1440
         Width           =   4095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Monster"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Set Monster's Range"
         Height          =   375
         Left            =   1440
         TabIndex        =   8
         Top             =   2160
         Width           =   1815
      End
      Begin VB.ComboBox CharWeapon 
         Height          =   315
         ItemData        =   "MakePlayer.frx":030A
         Left            =   2280
         List            =   "MakePlayer.frx":0317
         TabIndex        =   5
         Text            =   "Club"
         Top             =   960
         Width           =   3615
      End
      Begin VB.ComboBox CharHitPts 
         Height          =   315
         ItemData        =   "MakePlayer.frx":0330
         Left            =   2280
         List            =   "MakePlayer.frx":033D
         TabIndex        =   4
         Text            =   "2"
         Top             =   600
         Width           =   3615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Player"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Hint: Use | to separate text into different windows (SHIFT + \)"
         Height          =   255
         Left            =   1560
         TabIndex        =   12
         Top             =   1800
         Width           =   4335
      End
      Begin VB.Label Label2 
         Caption         =   "Set Text to display when encountered"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Weapon 
         Caption         =   "Weapon"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Health"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   615
      End
   End
End
Attribute VB_Name = "MakePlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim doneset As Boolean

Private Sub Check1_Click()
If Check1.Value = 1 Then
  Command3.Enabled = False
  MapEdit.MonMinX1 = MapEdit.MonCurX
  MapEdit.MonMinY1 = MapEdit.MonCurY
  MapEdit.MonMaxX2 = MapEdit.MonCurX
  MapEdit.MonMaxY2 = MapEdit.MonCurY
  Option1_Click 1
Else
  Command3.Enabled = True
  MapEdit.MonMinX1 = -1
  MapEdit.MonMinY1 = -1
  MapEdit.MonMaxX2 = -1
  MapEdit.MonMaxY2 = -1
  MapEdit.Outline2(0).Visible = False
  MapEdit.Outline2(1).Visible = False
  Command3.Caption = "&Set Monster's Range"
End If
End Sub

Private Sub Command1_Click()
On Error GoTo errh
Dim R As Integer
Dim Temp1
  If MapEdit.MonMinX1 > MapEdit.MonMaxX2 Then
    Temp1 = MapEdit.MonMinX1
    MapEdit.MonMinX1 = MapEdit.MonMaxX2
    MapEdit.MonMaxX2 = Temp1
  ElseIf MapEdit.MonMinY1 > MapEdit.MonMaxY2 Then
    Temp1 = MapEdit.MonMinY1
    MapEdit.MonMinY1 = MapEdit.MonMaxY2
    MapEdit.MonMaxY2 = Temp1
  End If

If LCase(CharWeapon.Text) <> "club" And LCase(CharWeapon.Text) <> "dagger" And LCase(CharWeapon.Text) <> "sword" Then
  R = MsgBox("Invalid weapon choice", vbCritical, "Invalid Choice")
  CharWeapon.SetFocus
  CharWeapon.SelStart = 1
  CharWeapon.SelLength = Len(CharWeapon.Text)
  Exit Sub
ElseIf Val(CharHitPts.Text) < 1 Or Val(CharHitPts.Text) > 3 Then
  R = MsgBox("Hitpoints must be between 1 and 3", vbCritical, "Invalid Choice")
  CharHitPts.SetFocus
  CharHitPts.SelStart = 1
  CharHitPts.SelLength = Len(CharWeapon.Text)
  Exit Sub
ElseIf Option1(1).Value And (MapEdit.MonMinX1 < 0 Or (MapEdit.MonMinY1 = Empty And Not MapEdit.MonMinY1 = 0)) Then
  R = MsgBox("You must first set the monster's range!", vbCritical, "Set range")
  Command3.SetFocus
  Exit Sub
ElseIf Option1(1).Value And (MapEdit.MonMinX1 > MapEdit.MonCurX Or MapEdit.MonMinY1 > MapEdit.MonCurY Or MapEdit.MonMaxX2 < MapEdit.MonCurX Or MapEdit.MonMaxY2 < MapEdit.MonCurY) Then
    R = MsgBox("The monster's range must include the monster!! Reset it's range.", vbCritical, "Invalid range")
    Command3.SetFocus
    Exit Sub
End If
Command1.Enabled = True
MapEdit.SetBounds = False
MapEdit.SetFirst = False
MapEdit.SetLast = False
MapEdit.mnuSave.Enabled = True
MapEdit.mnuSaveAs.Enabled = True
MapEdit.Outline2(0).Visible = False
MapEdit.Outline2(1).Visible = False
'MapEdit.Tile(MapEdit.MonCurIndex).Picture = MapEdit.Terrain(((InStr(MapEdit.TagString, c)) - 1) / 4).Picture

MapEdit.StatusBar1.Panels(3).Text = ""

MapEdit.SaveMakePlayerData

'Make the temp circle disappear
MapEdit.shapetmpMonster.Visible = False

Exit Sub
errh:
MsgBox Err & Err.Description
Stop
Resume
End Sub

Private Sub Command2_Click()
MapEdit.SetBounds = False
MapEdit.SetFirst = False
MapEdit.SetLast = False
MapEdit.mnuSave.Enabled = True
MapEdit.mnuSaveAs.Enabled = True
MapEdit.Outline2(0).Visible = False
MapEdit.Outline2(1).Visible = False
MapEdit.StatusBar1.Panels(3).Text = ""
Command1.Enabled = True
Unload Me
End Sub

Private Sub Command3_Click()
MapEdit.SetBounds = True
MapEdit.SetFirst = True
MapEdit.SetLast = True
doneset = False
Command3.Caption = "Re&set Monster's Range"
MapEdit.Outline2(0).Visible = False
MapEdit.Outline2(1).Visible = False
MapEdit.shapetmpMonster.Visible = True
MapEdit.shapetmpMonster.Left = MapEdit.MonCurX * 32
MapEdit.shapetmpMonster.Top = MapEdit.MonCurY * 32
Command1.Enabled = True
Command1.Caption = "&Create Monster"
Me.Visible = False
End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
If Left(MapEdit.GetTileXY("SpecialTile", (MapEdit.MonCurX), (MapEdit.MonCurY)), 6) = "Player" Then
  Me.Caption = "Player " & Mid(MapEdit.GetTileXY("SpecialTile", (MapEdit.MonCurX), (MapEdit.MonCurY)), 9) & "  Properties"
  Option1(0).Value = True
  Option1_Click 0
ElseIf Left(MapEdit.GetTileXY("SpecialTile", (MapEdit.MonCurX), (MapEdit.MonCurY)), 7) = "Monster" Then
  Me.Caption = "Monster " & Mid(MapEdit.GetTileXY("SpecialTile", (MapEdit.MonCurX), (MapEdit.MonCurY)), 10) & " Properties"
  Option1(1).Value = True
  Option1_Click 1
Else
  Me.Caption = "Monster " & MapEdit.TotalMonsters + 1 & " Properties"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
MapEdit.SetBounds = False
MapEdit.SetFirst = False
MapEdit.SetLast = False
End Sub

Private Sub Option1_Click(Index As Integer)
If Option1(1).Value = True Then
  Check1.Visible = True
  Command3.Visible = True
  Label2 = "Set Text to display when encountered"
  If MapEdit.MonMinY1 = Empty Then
    ' Haven't set range yet
    Command1.Caption = "Set Range First"
    Command1.Enabled = False
  Else
    Command1.Caption = "&Create Monster"
    Command1.Enabled = True
  End If
  Me.Caption = "Monster " & MapEdit.TotalMonsters + 1 & " Properties"
Else
  Check1.Visible = False
  Command3.Visible = False
  Label2 = "Set Text to display when game starts"
  Me.Caption = "Player " & MapEdit.TotalPlayers + 1 & " Properties"
  Command1.Caption = "&Create Player"
  Command1.Enabled = True
End If
End Sub
