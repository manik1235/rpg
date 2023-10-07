VERSION 5.00
Begin VB.Form MapEdit 
   Caption         =   "RPG Map Editor"
   ClientHeight    =   9270
   ClientLeft      =   540
   ClientTop       =   630
   ClientWidth     =   11385
   Icon            =   "meMapEditor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9270
   ScaleWidth      =   11385
   Visible         =   0   'False
   Begin VB.PictureBox Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      ScaleHeight     =   360
      ScaleWidth      =   11325
      TabIndex        =   22
      Top             =   0
      Width           =   11385
   End
   Begin VB.PictureBox StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   11325
      TabIndex        =   21
      Top             =   8895
      Width           =   11385
   End
   Begin VB.PictureBox ImageList 
      BackColor       =   &H80000005&
      Height          =   480
      Left            =   4560
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   25
      Top             =   6240
      Width           =   1200
   End
   Begin VB.PictureBox ImageList3 
      BackColor       =   &H80000005&
      Height          =   480
      Left            =   3360
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   26
      Top             =   6240
      Width           =   1200
   End
   Begin VB.PictureBox CurTile 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   480
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   20
      Top             =   5520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame ToolWindow 
      Caption         =   "Tool Window"
      Height          =   8295
      Left            =   7800
      TabIndex        =   6
      Top             =   480
      Width           =   3495
      Begin VB.Frame frameTool 
         Caption         =   "Characters"
         Height          =   6855
         Index           =   3
         Left            =   0
         TabIndex        =   19
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
         Begin VB.ListBox listTilesC 
            Height          =   2010
            Left            =   120
            TabIndex        =   24
            Top             =   1200
            Width           =   3375
         End
         Begin VB.HScrollBar HScroll3 
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   840
            Width           =   3375
         End
         Begin VB.Shape shpTileOutline 
            BorderColor     =   &H0000C000&
            BorderWidth     =   3
            Height          =   495
            Index           =   2
            Left            =   1560
            Tag             =   "0"
            Top             =   240
            Width           =   495
         End
         Begin VB.Image imgCurTileC 
            Height          =   495
            Index           =   0
            Left            =   120
            Top             =   240
            Width           =   495
         End
         Begin VB.Image imgCurTileC 
            Height          =   495
            Index           =   1
            Left            =   600
            Top             =   240
            Width           =   495
         End
         Begin VB.Image imgCurTileC 
            Height          =   495
            Index           =   2
            Left            =   1080
            Top             =   240
            Width           =   495
         End
         Begin VB.Image imgCurTileC 
            Height          =   495
            Index           =   3
            Left            =   1560
            Top             =   240
            Width           =   495
         End
         Begin VB.Image imgCurTileC 
            Height          =   495
            Index           =   4
            Left            =   2040
            Top             =   240
            Width           =   495
         End
         Begin VB.Image imgCurTileC 
            Height          =   495
            Index           =   5
            Left            =   2520
            Top             =   240
            Width           =   495
         End
         Begin VB.Image imgCurTileC 
            Height          =   495
            Index           =   6
            Left            =   3000
            Top             =   240
            Width           =   495
         End
         Begin VB.Shape Shape3 
            Height          =   495
            Left            =   120
            Top             =   240
            Width           =   3375
         End
      End
      Begin VB.Frame frameTool 
         Caption         =   "Tiles"
         Height          =   4935
         Index           =   2
         Left            =   0
         TabIndex        =   10
         Top             =   1680
         Visible         =   0   'False
         Width           =   3615
         Begin VB.VScrollBar VScroll2 
            Height          =   3375
            Left            =   3240
            TabIndex        =   18
            Top             =   240
            Width           =   255
         End
         Begin VB.ListBox listTiles 
            Height          =   3375
            ItemData        =   "meMapEditor.frx":030A
            Left            =   720
            List            =   "meMapEditor.frx":030C
            TabIndex        =   17
            Top             =   240
            Width           =   2775
         End
         Begin VB.Frame Frame1 
            Caption         =   "Tiles to list"
            Height          =   1095
            Left            =   120
            TabIndex        =   11
            Top             =   3720
            Width           =   3375
            Begin VB.CheckBox chkShowTiles 
               Caption         =   "All"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   16
               Top             =   240
               Value           =   1  'Checked
               Width           =   615
            End
            Begin VB.CheckBox chkShowTiles 
               Caption         =   "Paths"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   15
               Top             =   480
               Width           =   735
            End
            Begin VB.CheckBox chkShowTiles 
               Caption         =   "Walls"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   14
               Top             =   720
               Width           =   735
            End
            Begin VB.CheckBox chkShowTiles 
               Caption         =   "Weapons"
               Height          =   255
               Index           =   3
               Left            =   1920
               TabIndex        =   13
               Top             =   240
               Width           =   1095
            End
            Begin VB.CheckBox chkShowTiles 
               Caption         =   "Objects"
               Height          =   255
               Index           =   4
               Left            =   1920
               TabIndex        =   12
               Top             =   480
               Width           =   855
            End
         End
         Begin VB.Shape shpTileOutline 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            Height          =   495
            Index           =   0
            Left            =   120
            Tag             =   "0"
            Top             =   1680
            Width           =   495
         End
         Begin VB.Image imgCurTile 
            Height          =   495
            Index           =   3
            Left            =   120
            Top             =   1680
            Width           =   495
         End
         Begin VB.Shape Shape2 
            Height          =   3375
            Left            =   120
            Top             =   240
            Width           =   495
         End
         Begin VB.Image imgCurTile 
            Height          =   495
            Index           =   0
            Left            =   120
            Top             =   240
            Width           =   495
         End
         Begin VB.Image imgCurTile 
            Height          =   495
            Index           =   1
            Left            =   120
            Top             =   720
            Width           =   495
         End
         Begin VB.Image imgCurTile 
            Height          =   495
            Index           =   2
            Left            =   120
            Top             =   1200
            Width           =   495
         End
         Begin VB.Image imgCurTile 
            Height          =   495
            Index           =   4
            Left            =   120
            Top             =   2160
            Width           =   495
         End
         Begin VB.Image imgCurTile 
            Height          =   495
            Index           =   5
            Left            =   120
            Top             =   2640
            Width           =   495
         End
         Begin VB.Image imgCurTile 
            Height          =   495
            Index           =   6
            Left            =   120
            Top             =   3120
            Width           =   495
         End
      End
      Begin VB.Frame frameTool 
         Caption         =   "Terrain"
         Height          =   3495
         Index           =   1
         Left            =   0
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   3615
         Begin VB.HScrollBar HScroll2 
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   840
            Width           =   3375
         End
         Begin VB.ListBox listTilesB 
            Height          =   2205
            Left            =   120
            TabIndex        =   8
            Top             =   1200
            Width           =   3375
         End
         Begin VB.Shape shpTileOutline 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   3
            Height          =   495
            Index           =   1
            Left            =   1560
            Tag             =   "0"
            Top             =   240
            Width           =   495
         End
         Begin VB.Shape Shape1 
            Height          =   495
            Left            =   120
            Top             =   240
            Width           =   3375
         End
         Begin VB.Image imgCurTileb 
            Height          =   495
            Index           =   6
            Left            =   3000
            Top             =   240
            Width           =   495
         End
         Begin VB.Image imgCurTileb 
            Height          =   495
            Index           =   5
            Left            =   2520
            Top             =   240
            Width           =   495
         End
         Begin VB.Image imgCurTileb 
            Height          =   495
            Index           =   4
            Left            =   2040
            Top             =   240
            Width           =   495
         End
         Begin VB.Image imgCurTileb 
            Height          =   495
            Index           =   3
            Left            =   1560
            Top             =   240
            Width           =   495
         End
         Begin VB.Image imgCurTileb 
            Height          =   495
            Index           =   2
            Left            =   1080
            Top             =   240
            Width           =   495
         End
         Begin VB.Image imgCurTileb 
            Height          =   495
            Index           =   1
            Left            =   600
            Top             =   240
            Width           =   495
         End
         Begin VB.Image imgCurTileb 
            Height          =   495
            Index           =   0
            Left            =   120
            Top             =   240
            Width           =   495
         End
      End
   End
   Begin VB.PictureBox ComDialog 
      Height          =   480
      Left            =   2040
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   27
      Top             =   5520
      Width           =   1200
   End
   Begin VB.PictureBox TempPic 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1440
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   5
      Top             =   5520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox CharTile 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   960
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   4
      Top             =   5520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox EraseTile 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   5520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   4815
      Left            =   6120
      SmallChange     =   10
      TabIndex        =   2
      Top             =   480
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      SmallChange     =   10
      TabIndex        =   1
      Top             =   5280
      Width           =   6135
   End
   Begin VB.PictureBox ContainPic 
      AutoRedraw      =   -1  'True
      Height          =   4695
      Left            =   0
      ScaleHeight     =   309
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   397
      TabIndex        =   0
      Top             =   480
      Width           =   6015
      Begin VB.Shape shapetmpMonster 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         Height          =   495
         Left            =   3840
         Shape           =   3  'Circle
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Shape shapeFill 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00FF0000&
         Height          =   495
         Left            =   1920
         Top             =   2400
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Shape sSel 
         Height          =   495
         Left            =   2640
         Top             =   1320
         Width           =   495
      End
      Begin VB.Shape shapeMonster 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         Height          =   495
         Index           =   0
         Left            =   3240
         Shape           =   3  'Circle
         Top             =   2400
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Shape shapePlayer 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         Height          =   495
         Index           =   0
         Left            =   2400
         Shape           =   4  'Rounded Rectangle
         Top             =   2400
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Shape Outline2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Height          =   495
         Index           =   1
         Left            =   3360
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Shape Outline2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Height          =   495
         Index           =   0
         Left            =   2400
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "S&ave As"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
   End
   Begin VB.Menu mnuToolsRoot 
      Caption         =   "&Tools"
      Begin VB.Menu mnuMapTiles 
         Caption         =   "&Map Tiles"
         Begin VB.Menu mnuCharacterTiles 
            Caption         =   "&Characters"
            Begin VB.Menu mnuCharacter 
               Caption         =   "Character"
               Index           =   0
            End
         End
         Begin VB.Menu sep8 
            Caption         =   "-"
         End
         Begin VB.Menu mnuBaseTerrainRoot 
            Caption         =   "&Terrain"
            Begin VB.Menu mnuBaseTerrain 
               Caption         =   "Terrain"
               Index           =   0
            End
         End
         Begin VB.Menu sep10 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAllTilesRoot 
            Caption         =   "&All Tiles"
            Begin VB.Menu mnuAllTiles 
               Caption         =   "All Tiles"
               Index           =   0
            End
         End
         Begin VB.Menu mnuPathsRoot 
            Caption         =   "&Paths"
            Begin VB.Menu mnuPaths 
               Caption         =   "Path"
               Index           =   0
            End
         End
         Begin VB.Menu mnuWallsRoot 
            Caption         =   "&Walls"
            Begin VB.Menu mnuWall 
               Caption         =   "Wall"
               Index           =   0
            End
         End
         Begin VB.Menu mnuWeaponRoot 
            Caption         =   "&Weapons"
            Begin VB.Menu mnuWeapon 
               Caption         =   "Weapon"
               Index           =   0
            End
         End
         Begin VB.Menu mnuObjectsRoot 
            Caption         =   "&Objects"
            Begin VB.Menu mnuObject 
               Caption         =   "Object"
               Index           =   0
            End
         End
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "Map &Properties"
         Index           =   0
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "&Player Properties"
         Index           =   1
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "M&onsterProperties"
         Index           =   2
      End
   End
   Begin VB.Menu mnuHelpRoot 
      Caption         =   "&Help"
      NegotiatePosition=   3  'Right
      Begin VB.Menu mnuHelp 
         Caption         =   "Rpg Map Editor &Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About Rpg Map Editor"
      End
   End
End
Attribute VB_Name = "MapEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'BitBlt const
Const SRCCOPY = &HCC0020
Const SRCAND = &H8800C6
Const SRCPAINT = &HEE0086
Const SRCOR = &HEE0086 'same as paint, but i like this name more

'--------------------------------------------------------------------------
'-----------------------------These are for ToolWIndow---------------------
Dim w As Integer
Dim curtool As Integer
'-----------------------------These are for ToolWIndow---------------------
'--------------------------------------------------------------------------

'This is the TagString
Public TagString As String
Public TagStringC As String 'Character
Public TagStringB As String 'Base Terrain

'These denote special tiles
Public IsCharTile As Boolean
Public CurMon As String ' The current player/monster the user clicked on

'These are how big the map is (in tiles)
Public MapX As Integer
Public MapY As Integer

'Map's Name
Public MapName As String

'used  to loop while tiling
Dim X As Integer
Dim Y As Integer
Dim c As Integer

'This is true if user is right clicking to erase
Dim RightClick As Boolean

'Neg. or 0=Not Erasing 1=Make Terrain Default 2=Keep Terrain, Erase Tile 3=Erase Character
Dim chkErase As Integer

'This hold the # of tiles there are
Public MaxTerrain

'Msgbox response
Dim R As Integer

'Used to remember stuff for Char stats
Public SetBounds As Boolean
Public SetLast As Boolean
Public SetFirst As Boolean

'Used to remember if Me is drawing with the selection tool (shift)
Dim AmSelecting As Boolean
'Holds X and Y of initial selection
Dim iX As Single
Dim iY As Single

'File Vars
Public FileName As String 'This holds filename
Public filepath As String 'This holds path
Dim Filter As String
Dim FileHandle As Integer
Dim SavedMap As String 'This holds the string of Chars for the saved map
Public FirstSave As Boolean
Public JustSaved As Boolean 'This is true if you just saved, false if you didn't and made changes

'These are for players and monsters
Const MAXPLAYERS = 7
Const MAXMONSTERS = 999

' This converts twips to pixels
Const TwipTOPixel As Long = 15



'These won't be written to file
Public Data
Public DefaultTile
Public MapSize
Public MapDesc
Public MonCurX 'Holds the X value to send to MakePlayer
Public MonCurY
Public MonCurIndex
Public MonCurImage
Public MonMinX1 'Holds the Min X range to send to MakePlayer
Public MonMinY1
Public MonMaxX2
Public MonMaxY2
'These hold the TagString values of each individual tile
Dim Tile As TileType
'These will be written to file
Public TotalMonsters
Public TotalPlayers
Public TotalSigns
Public TotalHouses
Public MapTopX    'This is to hold where on the map the Screen should start at
Public MapTopY
Dim Player(MAXPLAYERS) As PlayerType
Dim Monster(MAXMONSTERS) As MonsterType
Dim Sign(MAXMONSTERS) As SignType 'Sign's pos. (Sign#, (Text, Sound, Choice #1, Choice #2, Num. Times to visit (0 for infinite)))
Dim House(MAXMONSTERS) As HouseType  '(House#, (Text, Sound, Choice #1, Choice #2, Num. Times to visit (0 for infinite)))


Private Sub Characters_Click(Index As Integer)
CurTile.Picture = Characters(Index).Picture
MonCurImage = Index
IsCharTile = True
End Sub

Private Sub CharTile_Click()
MsgBox CharTile.Tag
End Sub

Private Sub chkShowTiles_Click(Index As Integer)
CreateList
End Sub



Private Sub ContainPic_KeyDown(KeyCode As Integer, Shift As Integer)
' Depress the erase button if ctrl is pressed
If Shift = 2 Or Shift = 3 Then
  'Toolbar1.Buttons(5).Value = tbrPressed
Else
  'Toolbar1.Buttons(5).Value = tbrUnpressed
End If
End Sub

Private Sub ContainPic_KeyUp(KeyCode As Integer, Shift As Integer)
' Depress the erase button if ctrl is pressed
If Shift = 2 Or Shift = 3 Then
  Toolbar1.Buttons(5).Value = tbrPressed
Else
  Toolbar1.Buttons(5).Value = tbrUnpressed
End If
End Sub

Private Sub ContainPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Error GoTo errh
On Error Resume Next

  Dim rc
  Dim tmpTexture As String
  Dim NewTag As String 'temp storage of the tag to change to
    
    If (Shift = 1 Or Shift = 3) And Shift <> 2 And Shift < 1 And Shift > 3 Then
      'if shift is let go of, you are no longer selecting
      shapeFill.Visible = False
      AmSelecting = False
    End If
    
    If IsCharTile And Button = vbLeftButton Then
      If Left(Tile.XY(pX(X), pY(Y)).SpecialTile, 6) = "Player" Or Left(Tile.XY(pX(X), pY(Y)).SpecialTile, 7) = "Monster" And Not SetBounds Then
        R = MsgBox("There is already a character here. Would you like to edit it's properties instead?", vbQuestion + vbYesNo)
        If R = vbNo Then
          Exit Sub
        Else
          CurMon = Mid(Tile.XY(pX(X), pY(Y)).SpecialTile, 8)
          Properties.Show
          Exit Sub
        End If
      End If
      ClickTile X, Y
      Exit Sub
    End If
    
    If Button And Shift = 0 Then
      If Button = vbRightButton Then
        'Right mouse button (erase)
        'If Left(Tile.XY(pX(X), pY(Y)).SpecialTile, 6) <> "" Then
          'KillCharacter X, Y
        'Else
          ' Cut out the "guts" of the Texture code Gxx_ and replace them with zeros, to make it the blank tile.
          tmpTexture = Tile.XY(pX(X), pY(Y)).Tag
          
          If Mid(tmpTexture, 2, 2) <> "00" Then
            'If it is plain terrain, do nothing:: other wise continue
            If Right(tmpTexture, 1) = "&" Then
              'This means the top tile can be over water, so find the tile it goes to
              Tile.XY(pX(X), pY(Y)).Tag = DrawMask(X, Y, Left(tmpTexture, 1), "", "")
            Else
              'otherwise continue as normal
              Tile.XY(pX(X), pY(Y)).Tag = DrawMask(X, Y, Tilefrm.BaseTerrain(TerrainVal(Left(tmpTexture, 1), TagStringB)).Tag)
            End If
          End If
        'End If
      Else
        'Left Mouse Button (Place Tile/Character)
        If chkErase Then
          ' The eraser is checked, place the Default tile.
          Tile.XY(pX(X), pY(Y)).Tag = DrawMask(X, Y, EraseTile.Tag, "", "")
        Else
          If Len(CurTile.Tag) = 4 And Left(CurTile.Tag, 1) <> " " Then
            ' If its Terrain (the begining char isnt a space)
            If Mid(Tile.XY(pX(X), pY(Y)).Tag, 2, 2) = "00" Then
              ' If it is being laid on a blank terrain, replace the symbol at the end
              ' Place it over w/o a mask
              Tile.XY(pX(X), pY(Y)).Tag = DrawMask(X, Y, CurTile.Tag, "", "")
            ElseIf Right(Tile.XY(pX(X), pY(Y)).Tag, 1) = "~" Then
              ' If its water,
            Else
              ' Draw it with a mask (terrain)
              If Right(Tile.XY(pX(X), pY(Y)).Tag, 1) = Right(CurTile.Tag, 1) Or (Right(Tile.XY(pX(X), pY(Y)).Tag, 1) = "&" And Right(CurTile.Tag, 1) = "_") Then 'if the tag is & then the tile can go under
                ' If it is being put on a path, make sure the symbols match, and then place it underneath
                Tile.XY(pX(X), pY(Y)).Tag = DrawMask(X, Y, CurTile.Tag, Tile.XY(pX(X), pY(Y)).Tag, TagString)
              Else
                ' If it is being put on something that doesnt have the same symbol, erase and overwrite
                Tile.XY(pX(X), pY(Y)).Tag = DrawMask(X, Y, CurTile.Tag, "", "")
              End If
            End If
          ElseIf Len(CurTile.Tag) = 4 And Left(CurTile.Tag, 1) = " " Then
            ' If its a Terrain modifier (eg. Path)
            If Left(Tile.XY(pX(X), pY(Y)).Tag, 1) = "F" Then
              ' If it is being placed on a forest, change it to grass (Cut down those trees!)
              Tile.XY(pX(X), pY(Y)).Tag = DrawMask(X, Y, "G00_", CurTile.Tag, TagString)
            ElseIf Right(Tile.XY(pX(X), pY(Y)).Tag, 1) = "~" Then
              'if it is being placed on water, it must be able to be placed there (has a ~ or &)
              If Right(CurTile.Tag, 1) = "~" Or Right(CurTile.Tag, 1) = "&" Then
                Tile.XY(pX(X), pY(Y)).Tag = DrawMask(X, Y, Left(Tile.XY(pX(X), pY(Y)).Tag, 1) & "00" & Right(Tile.XY(pX(X), pY(Y)).Tag, 1), CurTile.Tag, TagString)
              Else
                ' You are trying to place a non water tile on water
                Exit Sub
              End If
            ElseIf Left(Tile.XY(pX(X), pY(Y)).Tag, 1) <> "0" Then
              ' Make sure it isnt a Blank tile (Black) nothing can go on these, they are borders (only terrain tiles can overwrite these)
              Tile.XY(pX(X), pY(Y)).Tag = DrawMask(X, Y, Left(Tile.XY(pX(X), pY(Y)).Tag, 1) & "00" & Right(Tile.XY(pX(X), pY(Y)).Tag, 1), CurTile.Tag, TagString)
            Else
              ' if it can't print the tile (for example a path on a Blank (Black) square, exit
              Exit Sub
            End If
          Else
            ' shouldnt get here... for detecting errors
            MsgBox "There was an error in 'MapEdit.frm' in the sub 'ContainPic_MouseDown'. Most likely a terrain error"
            Stop
          End If
        End If
          
        'This will correct the tiles for crossover pictures (grass to dirt) type things
        TempPic.Picture = Tilefrm.Terrain(TerrainVal(Tile.XY(pX(X), pY(Y)).Tag)).Picture
        TileChange X, Y
      End If
    ElseIf Button = 1 And Shift = 4 Then
      ' Manually edit the current tag (Alt+Left Click)
      NewTag = InputBox("Enter a four character tile tag.", "Manually set tile", Tile.XY(pX(X), pY(Y)).Tag)
      If Mid(NewTag, 2, 2) = "00" Then
        ' It is only terrain, don't place a toptile
        Tile.XY(pX(X), pY(Y)).Tag = DrawMask(X, Y, NewTag)
      ElseIf NewTag = "" Then
        'Don't do anything if nothing is entered (or canceled)
      Else
        ' Place the terrain and the mask
        Tile.XY(pX(X), pY(Y)).Tag = DrawMask(X, Y, NewTag, " " & Mid(NewTag, 2, 3), TagString)
      End If
    ElseIf Button = 1 And Shift = 2 Then
      ' Erase, place the default tile
      Tile.XY(pX(X), pY(Y)).Tag = DrawMask(X, Y, EraseTile.Tag, "", "")
      ' Depress the erase button
      'Toolbar1.Buttons(5).Value = tbrPressed
    ElseIf Button = 1 And (Shift = 1 Or Shift = 3) Then
      If Shift = 3 Then
        ' Depress the erase button
        'Toolbar1.Buttons(5).Value = tbrPressed
      End If
      ' Fills in the rectangle with that texture
      If Not AmSelecting Then
        ' First, record where clicked
        iX = X
        iY = Y
        AmSelecting = True
        shapeFill.Visible = True
        shapeFill.Left = X
        shapeFill.Top = Y
      End If
      ' Place the X coordinate
      If X - iX > 0 Then
        shapeFill.Width = X - iX
        shapeFill.Left = iX
      ElseIf X - iX < 0 Then
        shapeFill.Left = X
        shapeFill.Width = Abs(iX - X)
      ElseIf X - iX = 0 Then
        shapeFill.Width = 1
        shapeFill.Left = X
      End If
      'Place the Y coordinate
      If Y - iY > 0 Then
        shapeFill.Height = Y - iY
        shapeFill.Top = iY
      ElseIf Y - iY < 0 Then
        shapeFill.Top = Y
        shapeFill.Height = Abs(Y - iY)
      ElseIf Y - iY = 0 Then
        shapeFill.Height = 1
        shapeFill.Top = Y
      End If
      
      
      
      
      
    End If
    ContainPic.Refresh
'  If Button Then
'    If X / 15 >= ContainPic.ScaleWidth - 30 And X / 15 <= ContainPic.ScaleWidth Then
'      If HScroll1.Value < HScroll1.Max - 10 Then HScroll1.Value = HScroll1.Value + 10
'    End If
'    If X / 15 <= 10 And X / 15 >= 0 Then
'      If HScroll1.Value > 9 Then HScroll1.Value = HScroll1.Value - 10
'    End If
'    If Y / 15 >= ContainPic.ScaleHeight - 10 And Y / 15 <= ContainPic.ScaleHeight Then
'      If VScroll1.Value < VScroll1.Max - 10 Then VScroll1.Value = VScroll1.Value + 10
'    End If
'    If Y / 15 <= 10 And Y / 15 >= 0 Then
'      If VScroll1.Value > 9 Then VScroll1.Value = VScroll1.Value - 10
'    End If '
'  End If

  Exit Sub
errh:
If Err = 9 Then
  ' subscript out of range (off the page)
  Exit Sub
End If
MsgBox Err & Err.Description
Stop
Resume
End Sub

Private Sub ContainPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errh

  If X < 0 Or Y < 0 Or X >= ContainPic.ScaleWidth Or Y >= ContainPic.ScaleHeight Then Exit Sub
  
  ' make the square cursor follow your mouse
  sSel.Left = Int(X / Tile.PixelWidth) * Tile.PixelWidth
  sSel.Top = Int(Y / Tile.PixelHeight) * Tile.PixelHeight
  
  'The X,Y position (blocks)
  'StatusBar1.Panels(1).Text = Int(X / Tile.PixelWidth) & "," & Int(Y / Tile.PixelHeight)
  'Shows if there is a special tile
  'StatusBar1.Panels(2).Text = Tile.XY(pX(X), pY(Y)).SpecialTile
  'shows the tile code
  'StatusBar1.Panels(4).Text = Tile.XY(pX(X), pY(Y)).Tag
  If RightClick = False And Button = 0 And Shift = 0 And Left(Tile.XY(pX(X), pY(Y)).SpecialTile, 7) = "Monster" And Not SetBounds Then
      'StatusBar1.Panels(3).Text = Monster(Val(Mid(Tile.XY(pX(X), pY(Y)).SpecialTile, 10))).MinX & "," & Monster(Val(Mid(Tile.XY(pX(X), pY(Y)).SpecialTile, 10))).MinY & "-" & Monster(Val(Mid(Tile.XY(pX(X), pY(Y)).SpecialTile, 10))).MaxX & "," & Monster(Val(Mid(Tile.XY(pX(X), pY(Y)).SpecialTile, 10))).MaxY
      Outline2(0).FillStyle = 5 'Downward Diagonal lines
      Outline2(0).Visible = True
      Outline2(0).Left = Monster(Val(Mid(Tile.XY(pX(X), pY(Y)).SpecialTile, 10))).MinX * Tile.PixelWidth
      Outline2(0).Top = Monster(Val(Mid(Tile.XY(pX(X), pY(Y)).SpecialTile, 10))).MinY * Tile.PixelHeight
      Outline2(0).Width = (Monster(Val(Mid(Tile.XY(pX(X), pY(Y)).SpecialTile, 10))).MaxX * Tile.PixelWidth) - Outline2(0).Left + Outline2(1).Width 'I added the height and width of Outline2(1) instead
      Outline2(0).Height = (Monster(Val(Mid(Tile.XY(pX(X), pY(Y)).SpecialTile, 10))).MaxY * Tile.PixelHeight) - Outline2(0).Top + Outline2(1).Height '     of Outline2(0) because #1 will stay constant
      'OutlinePic = (Monster(Val(Mid(Tile.XY(pX(X), pY(Y)).SpecialTile, 10))).MinY) * MapX
      'OutlinePic = OutlinePic + Monster(Val(Mid(Tile.XY(pX(X), pY(Y)).SpecialTile, 10))).MinX
      'Outline2(1).Visible = True
      'Outline2(1).Left = Monster(Val(Mid(Tile.XY(pX(X), pY(Y)).SpecialTile, 10))).MaxX * Tile.PixelWidth
      'Outline2(1).Top = (Monster(Val(Mid(Tile.XY(pX(X), pY(Y)).SpecialTile, 10))).MaxY) * Tile.PixelHeight
  Else
    If Not SetBounds Then
      Outline2(0).Visible = False
      Outline2(0).FillStyle = 1 'transparent
      Outline2(0).Width = Outline2(1).Width
      Outline2(0).Height = Outline2(1).Height
      Outline2(1).Visible = False
      'StatusBar1.Panels(3).Text = ""
    End If
  End If
  ContainPic_MouseDown Button, Shift, X, Y
  

  
  Exit Sub
errh:
If Err = 9 Then
  ' subscript out of range
  Exit Sub
End If
MsgBox Err & Err.Description
Stop
Resume
End Sub

Private Sub ContainPic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errh

Dim dx As Single
Dim dy As Single
Dim dXStep As Integer
Dim dYStep As Integer
Dim ErrWhere As String
Dim ThisTile As String

If Shift = 2 Or Shift = 3 Then
  ' unpress the erase button
  'Toolbar1.Buttons(5).Value = tbrUnpressed
End If

If Button = 1 And (Shift = 1 Or Shift = 3) Then
  ' This stops it from group selecting when the mouse is let go of
  
  If Shift = 1 Then
    ' Fill with curtile
    ThisTile = CurTile.Tag
  Else
    ' Fill with erase
    ThisTile = EraseTile.Tag
  End If
  
  ErrWhere = "x"
  dXStep = Abs(X - iX) / (X - iX)
  ErrWhere = "y"
  dYStep = Abs(Y - iY) / (Y - iY)
  ErrWhere = "else"
  
  ' Paint each tile
  For dx = iX To X Step dXStep * Tile.PixelWidth
    For dy = iY To Y Step dYStep * Tile.PixelHeight
      If Mid(ThisTile, 2, 2) = "00" Then
        Tile.XY(pX(dx), pY(dy)).Tag = DrawMask(dx, dy, Left(ThisTile, 1))
      Else
        Tile.XY(pX(dx), pY(dy)).Tag = DrawMask(dx, dy, Left(ThisTile, 1), Mid(ThisTile, 2), TagString)
      End If
    Next dy
  Next dx
  AmSelecting = False
  shapeFill.Visible = False
End If


Exit Sub
errh:
If Err = 9 Then
  'Subscript out of range, tile isn't on map
  Resume Next
End If
If Err = 11 Then
  'division by zero
  If ErrWhere = "x" Then
    dXStep = 0
  ElseIf ErrWhere = "y" Then
    dYStep = 0
  Else
    MsgBox "Divide by zero"
    Stop
    Resume
  End If
  Resume Next
End If
MsgBox Err & Err.Description
Stop
Resume
End Sub

Private Sub CurTile_Click()
MsgBox CurTile.Tag
End Sub

Private Sub EraseTile_Click()
MsgBox EraseTile.Tag
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errh

If KeyCode = 39 Then
  ' Right Arrow
  If frameTool(1).Visible Then
    'Terrain Toolwindow
    imgCurTileb_Click 4 'click the tile one to the right
  ElseIf frameTool(3).Visible Then
    'characters
    imgCurTileC_Click 4
  End If
ElseIf KeyCode = 37 Then
  ' Left arrow
  If frameTool(1).Visible Then
    'Terrain Toolwindow
    imgCurTileb_Click 2 'click the tile one to the left
  ElseIf frameTool(3).Visible Then
    'characters
    imgCurTileC_Click 2
  End If
ElseIf KeyCode = 40 Then
  ' Down arrow
  If frameTool(2).Visible Then
    'Tiles
    imgCurTile_Click 4 'click the tile one to the down
  End If
ElseIf KeyCode = 38 Then
  ' Down arrow
  If frameTool(2).Visible Then
    'Tiles
    imgCurTile_Click 2 'click the tile one to the down
  End If
End If




Exit Sub
errh:
MsgBox Err & Err.Description
Stop
Resume
End Sub


Private Sub Form_Load()

On Error GoTo errh

Dim TagSpot As Variant

'Me.Show
'center
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2

'Sets the total... vars
TotalMonsters = -1
TotalPlayers = -1
TotalSigns = -1
TotalHouses = -1


' Load all the tiles and assign tags to them, and create the TagString
LoadTiles Me



'This sets file basics
MapX = 0
MapY = 0
FileName = "Untitled"
filepath = "C:\"


' Create a Grass map
Load NewMapfrm
NewMapfrm.Option1(4).Value = True
NewMapfrm.Option(0).Value = True
Command2N_Click 'This goes to the sub that the Create New button goes to when clicked


While MapX = 0
  DoEvents
Wend

'This creates the Tile menu items
CreateMenu

' Create the list.
CreateList

' Create the Image List
CreateImageList

' Create the toolbar
CreateToolbar

' In the beginning, IsCharTile is set to True because it is the last toolwindow that is made. set it back
IsCharTile = False

' Set the caption and run the Resize sub
Me.Caption = "RPG Map Editor - " & FileName & ".rpg"
Form_Resize

On Error GoTo 0
Exit Sub
errh:
If Err = 360 Then Resume Next
MsgBox Err & Err.Description
Stop
Resume

End Sub



Public Sub Form_Resize()
On Error Resume Next
If Me.Width < 3180 Then Me.Width = 3180
If Me.Height < 3045 Then Me.Height = 3045
If Me.WindowState = 1 Then Exit Sub
VScroll1.Top = Toolbar1.Height
If Me.WindowState = 1 Then Exit Sub
VScroll1.Height = (Me.ScaleHeight - StatusBar1.Height - HScroll1.Height - Toolbar1.Height)
VScroll1.Left = (Me.ScaleWidth - VScroll1.Width - (ToolWindow.Width * ToolWindow.Visible * -1))
HScroll1.Width = (Me.ScaleWidth - VScroll1.Width - (ToolWindow.Width * ToolWindow.Visible * -1))
HScroll1.Top = VScroll1.Height + VScroll1.Top
If (Me.ScaleWidth - VScroll1.Width) > ContainPic.Width Then Me.ScaleWidth = ContainPic.Width + VScroll1.Width + 100
If (Me.ScaleHeight - HScroll1.Height - HScroll1.Height) > ContainPic.Height Then Me.ScaleHeight = ContainPic.Height + (HScroll1.Height * 2) + 140
Me.ScaleMode = 1  'Resets the scale so it works
HScroll1.LargeChange = ContainPic.Width / 10
VScroll1.LargeChange = ContainPic.Height / 10
HScroll1.Max = ContainPic.Width - (HScroll1.Width)   '+ pX((ToolWindow.Width * ToolWindow.Visible * -1)) 'Sets the Hor. Scroll Bar
VScroll1.Max = ContainPic.Height - VScroll1.Height '- ContainPic.Top - HScroll1.Height '(ContainPic.Height - VScroll1.Height) / 4
If Me.ScaleWidth + (VScroll1.Width * VScroll1.Visible) - (ToolWindow.Width * ToolWindow.Visible * -1) >= ContainPic.Width Then
  HScroll1.Visible = False
Else
  HScroll1.Visible = True
End If
If Me.ScaleHeight + (HScroll1.Height * HScroll1.Visible) - Toolbar1.Height - StatusBar1.Height >= ContainPic.Height Then
  VScroll1.Visible = False
Else
  VScroll1.Visible = True
End If

ToolWindow.Top = StatusBar1.Height
ToolWindow.Left = Me.ScaleWidth - ToolWindow.Width
ToolWindow.Height = Me.ScaleHeight - StatusBar1.Height - Toolbar1.Height + 40

Me.Caption = "RPG Map Editor - " & MapName & ".rpg"
'Me.SetFocus


End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not JustSaved Then
  R = MsgBox("Save changes to " & FileName & "?", vbQuestion + vbYesNoCancel, Me.Caption)
  If R = vbYes Then
    mnuSave_Click
    If Not JustSaved Then Cancel = 1: Exit Sub
  ElseIf R = vbCancel Then
    Cancel = 1
    Exit Sub
  End If
End If
End
  
End Sub

Private Sub HScroll1_Change()
HScroll1_Scroll
End Sub

Private Sub HScroll1_Scroll()
ContainPic.Left = -HScroll1.Value
End Sub

Private Sub HScroll2_Change()
HScroll2_Scroll
End Sub

Private Sub HScroll2_Scroll()
On Error Resume Next
listTilesB.ListIndex = HScroll2.Value
End Sub

Private Sub imgCurTile_Click(Index As Integer)
If imgCurTile(Index).Picture = LoadPicture() Then Exit Sub ' if it doesnt contain a picture

If listTiles.ListCount > 0 Then
  listTiles.ListIndex = Val(Left(imgCurTile(Index).Tag, InStr(1, imgCurTile(Index).Tag, " ")))
    IsCharTile = False
  ' The sub isnt called if its the middle one (you arent changing anything)
  If Index = 3 Then listTiles_Click
End If
End Sub

Private Sub imgCurTile_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Button = 1 Then
'  imgCurTile_Click Index
'End If
End Sub

Private Sub imgCurTileb_Click(Index As Integer)
If imgCurTileb(Index).Picture = LoadPicture() Then Exit Sub ' if it doesnt contain a picture

If listTilesB.ListCount > 0 Then
  listTilesB.ListIndex = Val(Left(imgCurTileb(Index).Tag, InStr(1, imgCurTileb(Index).Tag, " ")))
  IsCharTile = False
  ' The sub isnt called if its the middle one (you arent changing anything)
  If Index = 3 Then listTilesB_Click
End If
End Sub


Private Sub imgCurTileC_Click(Index As Integer)
If imgCurTileC(Index).Picture = LoadPicture() Then Exit Sub ' if it doesnt contain a picture

If listTilesC.ListCount > 0 Then
  listTilesC.ListIndex = Val(Left(imgCurTileC(Index).Tag, InStr(1, imgCurTileC(Index).Tag, " ")))
  IsCharTile = True
  ' The sub isnt called if its the middle one (you arent changing anything)
  If Index = 3 Then
    listTilesB_Click
  End If
End If
End Sub

Private Sub listTiles_Click()
On Error GoTo errh

Dim ChangeBy As Integer
Dim c1 As Integer
Dim d As Integer

' set the scroll bar to the correct position
VScroll2.Value = listTiles.ListIndex

imgCurTile(3).Picture = Tilefrm.Terrain(TerrainVal(Right(listTiles.List(listTiles.ListIndex), 4), TagString)).Picture

'make sure the list is set to the center picture
'listTiles.Selected(Val(Left(imgCurTile(Index).Tag, 1))) = True

d = (imgCurTile.Count - 1) / 2
If d / 2 = Int(d / 2) Then
  d = d - 1
End If
  

For c = -d To d
  If listTiles.ListIndex + c >= 0 And listTiles.ListIndex + c < listTiles.ListCount Then 'if a list item exsists...
    imgCurTile(d + c).Tag = ((listTiles.ListIndex) + c) & " " & Right(listTiles.List(listTiles.ListIndex + c), 4)
    c1 = (listTiles.ListIndex) + c
    Select Case Mid(imgCurTile(d + c).Tag, Len(Str(c1)) + 1, 1)
      Case "C"
        If c = 0 Then
          ' sets the current tile if it is the center image
          IsCharTile = True
          CurTile.Picture = Tilefrm.Characters(TerrainVal(Right(listTiles.List(listTiles.ListIndex), 4), TagString) + c).Picture
          CurTile.Tag = Tilefrm.Terrain(TerrainVal(Right(listTiles.List(listTiles.ListIndex), 4), TagString) + c).Tag
        End If
        imgCurTile(d + c).Picture = Tilefrm.Characters(TerrainVal(Right(listTiles.List(listTiles.ListIndex), 4), TagString) + c).Picture
      Case " "
        If c = 0 Then
          ' sets the current tile if it is the center image
          IsCharTile = False
          CurTile.Picture = Tilefrm.Terrain(TerrainVal(Right(listTiles.List(listTiles.ListIndex), 4), TagString) + c).Picture
          CurTile.Tag = Tilefrm.Terrain(TerrainVal(Right(listTiles.List(listTiles.ListIndex), 4), TagString) + c).Tag
        End If
        'imgCurTile(d + c).Picture = Tilefrm.Terrain(TerrainVal(Right(listTiles.List(listTiles.ListIndex), 4), TagString) + c).Picture
      Case Else
        If c = 0 Then
          ' sets the current tile if it is the center image
          IsCharTile = False
          CurTile.Picture = Tilefrm.BaseTerrain(TerrainVal(Right(listTiles.List(listTiles.ListIndex), 4), TagString) + c).Picture
          CurTile.Tag = Tilefrm.Terrain(TerrainVal(Right(listTiles.List(listTiles.ListIndex), 4), TagString) + c).Tag
        End If
        imgCurTile(d + c).Picture = Tilefrm.BaseTerrain(TerrainVal(Right(listTiles.List(listTiles.ListIndex), 4), TagString) + c).Picture
    End Select
  Else
    imgCurTile(d + c).Picture = LoadPicture()
    imgCurTile(d + c).Tag = ""
  End If
Next c

' Set the Frame title to the Icon name
frameTool(2).Caption = listTiles.List(listTiles.ListIndex)

Exit Sub
errh:
If Err = 340 Then
  ' Control array element doesnt exsist
  imgCurTile(d + c).Picture = LoadPicture()
  Resume Next
End If


MsgBox Err & Err.Description
Stop
Resume
End Sub

Private Sub listTiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
  listTiles_Click
End If
End Sub

Private Sub listTiles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
listTiles_MouseDown Button, Shift, X, Y
End Sub

Private Sub listTilesB_Click()
On Error GoTo errh

Dim ChangeBy As Integer
Dim c1 As Integer
Dim d As Integer

' set the scroll bar to the correct position
HScroll2.Value = listTilesB.ListIndex

imgCurTileb(3).Picture = Tilefrm.BaseTerrain(TerrainVal(Right(listTilesB.List(listTilesB.ListIndex), 4), TagStringB)).Picture

'make sure the list is set to the center picture
'listTiles.Selected(Val(Left(imgCurTile(Index).Tag, 1))) = True

d = (imgCurTileb.Count - 1) / 2
If d / 2 = Int(d / 2) Then
  d = d - 1
End If
  

For c = -d To d
  If listTilesB.ListIndex + c >= 0 And listTilesB.ListIndex + c < listTilesB.ListCount Then 'if a list item exsists...
    imgCurTileb(d + c).Tag = ((listTilesB.ListIndex) + c) & " " & Right(listTilesB.List(listTilesB.ListIndex + c), 4)
    c1 = (listTilesB.ListIndex) + c
    'Select Case Mid(imgCurTileb(d + c).Tag, Len(Str(c1)) + 1, 1)
    '  Case "C"
    '    If c = 0 Then
    '      ' sets the current tile if it is the center image
    '      IsCharTile = True
    '      BaseTile.Picture = Tilefrm.Characters(TerrainVal(Right(listTiles.List(listTiles.ListIndex), 4), TagString) + c).Picture
    '      CurTile.Tag = Tilefrm.Terrain(TerrainVal(Right(listTiles.List(listTiles.ListIndex), 4), TagString) + c).Tag
    '    End If
    '    imgCurTileb(d + c).Picture = Tilefrm.Characters(TerrainVal(Right(listTiles.List(listTiles.ListIndex), 4), TagString) + c).Picture
    '  Case " "
    '    If c = 0 Then
    '      ' sets the current tile if it is the center image
    '      IsCharTile = False
    '      CurTile.Picture = Tilefrm.Terrain(TerrainVal(Right(listTiles.List(listTiles.ListIndex), 4), TagString) + c).Picture
    '      CurTile.Tag = Tilefrm.Terrain(TerrainVal(Right(listTiles.List(listTiles.ListIndex), 4), TagString) + c).Tag
    '    End If
    '    imgCurTileb(d + c).Picture = Tilefrm.Terrain(TerrainVal(Right(listTiles.List(listTiles.ListIndex), 4), TagString) + c).Picture
    '  Case Else
        If c = 0 Then
          ' sets the current tile if it is the center image
          CurTile.Picture = Tilefrm.BaseTerrain(TerrainVal(Right(listTilesB.List(listTilesB.ListIndex), 4), TagStringB) + c).Picture
          CurTile.Tag = Tilefrm.BaseTerrain(TerrainVal(Right(listTilesB.List(listTilesB.ListIndex), 4), TagStringB) + c).Tag
        End If
        imgCurTileb(d + c).Picture = Tilefrm.BaseTerrain(TerrainVal(Right(listTilesB.List(listTilesB.ListIndex), 4), TagStringB) + c).Picture
    'End Select
  Else
    imgCurTileb(d + c).Picture = LoadPicture()
    imgCurTileb(d + c).Tag = ""
  End If
Next c

' Set the Frame title to the Icon name
frameTool(1).Caption = listTilesB.List(listTilesB.ListIndex)

Exit Sub
errh:
If Err = 340 Then
  ' Control array element doesnt exsist
  imgCurTileb(d + c).Picture = LoadPicture()
  Resume Next
End If


MsgBox Err & Err.Description
Stop
Resume
End Sub

Private Sub listTilesB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
  listTilesB_Click
End If
End Sub

Private Sub listTilesB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
listTilesB_MouseDown Button, Shift, X, Y
End Sub

Private Sub listTilesC_Click()
On Error GoTo errh

Dim ChangeBy As Integer
Dim c1 As Integer
Dim d As Integer

' set the scroll bar to the correct position
HScroll3.Value = listTilesC.ListIndex

imgCurTileC(3).Picture = Tilefrm.Characters(TerrainVal(Right(listTilesC.List(listTilesC.ListIndex), 4), TagStringC)).Picture

'make sure the list is set to the center picture
'listTiles.Selected(Val(Left(imgCurTile(Index).Tag, 1))) = True
d = (imgCurTileC.Count - 1) / 2
If d / 2 = Int(d / 2) Then
  d = d - 1
End If

For c = -d To d
  If listTilesC.ListIndex + c >= 0 And listTilesC.ListIndex + c < listTilesC.ListCount Then 'if a list item exsists...
    imgCurTileC(d + c).Tag = ((listTilesC.ListIndex) + c) & " " & Right(listTilesC.List(listTilesC.ListIndex + c), 4)
    c1 = (listTilesC.ListIndex) + c
    'Select Case Mid(imgcurtilec(d + c).Tag, Len(Str(c1)) + 1, 1)
    '  Case "C"
    '    If c = 0 Then
    '      ' sets the current tile if it is the center image
    '      IsCharTile = True
    '      BaseTile.Picture = Tilefrm.Characters(TerrainVal(Right(listTiles.List(listTiles.ListIndex), 4), TagString) + c).Picture
    '      CurTile.Tag = Tilefrm.Terrain(TerrainVal(Right(listTiles.List(listTiles.ListIndex), 4), TagString) + c).Tag
    '    End If
    '    imgcurtilec(d + c).Picture = Tilefrm.Characters(TerrainVal(Right(listTiles.List(listTiles.ListIndex), 4), TagString) + c).Picture
    '  Case " "
    '    If c = 0 Then
    '      ' sets the current tile if it is the center image
    '      IsCharTile = False
    '      CurTile.Picture = Tilefrm.Terrain(TerrainVal(Right(listTiles.List(listTiles.ListIndex), 4), TagString) + c).Picture
    '      CurTile.Tag = Tilefrm.Terrain(TerrainVal(Right(listTiles.List(listTiles.ListIndex), 4), TagString) + c).Tag
    '    End If
    '    imgcurtilec(d + c).Picture = Tilefrm.Terrain(TerrainVal(Right(listTiles.List(listTiles.ListIndex), 4), TagString) + c).Picture
    '  Case Else
        If c = 0 Then
          ' sets the current tile if it is the center image
          IsCharTile = True
          CurTile.Picture = Tilefrm.Characters(TerrainVal(Right(listTilesC.List(listTilesC.ListIndex), 4), TagStringC) + c).Picture
          CurTile.Tag = Tilefrm.Characters(TerrainVal(Right(listTilesC.List(listTilesC.ListIndex), 4), TagStringC) + c).Tag
        End If
        imgCurTileC(d + c).Picture = Tilefrm.Characters(TerrainVal(Right(listTilesC.List(listTilesC.ListIndex), 4), TagStringC) + c).Picture
    'End Select
  Else
    imgCurTileC(d + c).Picture = LoadPicture()
    imgCurTileC(d + c).Tag = ""
  End If
Next c

' Set the Frame title to the Icon name
frameTool(1).Caption = listTilesC.List(listTilesC.ListIndex)

Exit Sub
errh:
If Err = 340 Then
  ' Control array element doesnt exsist
  imgCurTileC(d + c).Picture = LoadPicture()
  Resume Next
End If


MsgBox Err & Err.Description
Stop
Resume
End Sub

Private Sub mnuAllTiles_Click(Index As Integer)
IsCharTile = False
CurTile.Picture = Tilefrm.Terrain(TerrainVal(Right(mnuAllTiles(Index).Caption, 4), TagString)).Picture
CurTile.Tag = Tilefrm.Terrain(TerrainVal(Right(mnuAllTiles(Index).Caption, 4), TagString)).Tag
chkShowTiles(0).Value = 1
For c = 0 To listTiles.ListCount - 1
  If listTiles.List(c) = mnuAllTiles(Index).Caption Then
    listTiles.ListIndex = c
  End If
Next c

'Press the Tiles toolbar button
Toolbar1.Buttons(7).Value = tbrPressed
Toolbar1.Buttons(6).Value = tbrUnpressed
Toolbar1.Buttons(8).Value = tbrUnpressed
Toolbar1_ButtonClick Toolbar1.Buttons(7)
End Sub

Private Sub mnuCharacter_Click(Index As Integer)
CurTile.Picture = Tilefrm.Characters(TerrainVal(Right(mnuCharacter(Index).Caption, 4), TagStringC)).Picture
CurTile.Tag = Tilefrm.Characters(TerrainVal(Right(mnuCharacter(Index).Caption, 4), TagStringC)).Tag
IsCharTile = True

'press the character toolbar button
Toolbar1.Buttons(9).Value = tbrPressed
Toolbar1.Buttons(8).Value = tbrUnpressed
Toolbar1.Buttons(7).Value = tbrUnpressed
Toolbar1_ButtonClick Toolbar1.Buttons(9)
End Sub

Private Sub mnuExit_Click()
Form_Unload 0
End Sub




Private Sub mnuNew_Click()
R = MsgBox("Save changes to " & FileName & "?", vbQuestion + vbYesNoCancel, Me.Caption)
If R = vbYes Then
  mnuSave_Click
ElseIf R = vbCancel Then
  Exit Sub
End If

'This Unloads your level
ContainPic.Picture = LoadPicture()

'This Shows thats the map hasn't been saved yet
FirstSave = True

' This resets all variables and deletes charaters
ClearArrays

'This shows the new dialog
NewMapfrm.Show
'TileScreen
End Sub

Private Sub mnuObject_Click(Index As Integer)
IsCharTile = False
CurTile.Picture = Tilefrm.Terrain(TerrainVal(Right(mnuObject(Index).Caption, 4), TagString)).Picture
CurTile.Tag = Tilefrm.Terrain(TerrainVal(Right(mnuObject(Index).Caption, 4), TagString)).Tag
chkShowTiles(4).Value = 1
For c = 0 To listTiles.ListCount - 1
  If listTiles.List(c) = mnuObject(Index).Caption Then
    listTiles.ListIndex = c
  End If
Next c
'Press the Tiles toolbar button
Toolbar1.Buttons(7).Value = tbrPressed
Toolbar1.Buttons(6).Value = tbrUnpressed
Toolbar1.Buttons(8).Value = tbrUnpressed
Toolbar1_ButtonClick Toolbar1.Buttons(7)
End Sub

Private Sub mnuOpen_Click()
R = MsgBox("Save changes to " & FileName & "?", vbQuestion + vbYesNoCancel, Me.Caption)
If R = vbYes Then
  mnuSave_Click
ElseIf R = vbCancel Then
  Exit Sub
End If
NewMapfrm.Hide
'This tiles the screen, like it is a new level.
OpenMap
Exit Sub
End Sub

Private Sub mnuPForest_Click(Index As Integer)
SetTiles 7
Terrain_Click Index
End Sub



Private Sub mnuPaths_Click(Index As Integer)
IsCharTile = False
CurTile.Picture = Tilefrm.Terrain(TerrainVal(Right(mnuPaths(Index).Caption, 4), TagString)).Picture
CurTile.Tag = Tilefrm.Terrain(TerrainVal(Right(mnuPaths(Index).Caption, 4), TagString)).Tag
chkShowTiles(1).Value = 1
For c = 0 To listTiles.ListCount - 1
  If listTiles.List(c) = mnuPaths(Index).Caption Then
    listTiles.ListIndex = c
  End If
Next c
'Press the Tiles toolbar button
Toolbar1.Buttons(7).Value = tbrPressed
Toolbar1.Buttons(6).Value = tbrUnpressed
Toolbar1.Buttons(8).Value = tbrUnpressed
Toolbar1_ButtonClick Toolbar1.Buttons(7)
End Sub

Private Sub mnuProperties_Click(Index As Integer)
JustSaved = False
Properties.Show
'Properties.SSTab1.Tab = Index
End Sub



Private Sub mnuSave_Click()
'If Not TotalPlayers >= 0 Then
'  R = MsgBox("You must have a Player 0.", vbCritical, "No Player 0")
'  Exit Sub
'End If
On Error GoTo errh
If Not mnuSave.Enabled Then Exit Sub
If (FileName = "Untitled" And FirstSave) Or FirstSave Then
  mnuSaveAs_Click
  Exit Sub
End If
FileHandle = FreeFile
If Mid(FileName, Len(FileName) - 3, 1) = "." Then
  FileName = Mid(FileName, 1, Len(FileName) - 4)
End If
Close
Open filepath & FileName & ".rpg" For Output As #FileHandle
  c = 0
  For Y = 1 To MapY
    SavedMap = ""
    For X = 1 To MapX
      c = c + 1
      SavedMap = SavedMap & Tile.XY(X, Y).Tag
    Next X
    Print #FileHandle, SavedMap
  Next Y
c = 0
  Write #FileHandle, "MapName=", MapName
  Write #FileHandle, "FileName=", FileName
  Write #FileHandle, "Size=", MapSize
  Write #FileHandle, "DefaultTerrain=", DefaultTile
  Write #FileHandle, "TotalMonsters=", TotalMonsters
  Write #FileHandle, "TotalPlayers=", TotalPlayers
  Write #FileHandle, "TotalSigns=", TotalSigns
  Write #FileHandle, "TotalHouses=", TotalHouses
  Write #FileHandle, "MapTopX=", Player(0).StartX
  Write #FileHandle, "MapTopY=", Player(0).StartY
  Write #FileHandle, "Max="
  For c = 0 To TotalMonsters
    Write #FileHandle, Monster(c).MaxX, Monster(c).MaxY
  Next c
  If TotalMonsters = -1 Then
    'If there are no monsters, insert placeholders
    Write #FileHandle, -1, -1
  End If
  Write #FileHandle, "Min="
  For c = 0 To TotalMonsters
    Write #FileHandle, Monster(c).MinX, Monster(c).MinY
  Next c
  If TotalMonsters = -1 Then
    'If there are no monsters, insert placeholders
    Write #FileHandle, -1, -1
  End If
  Write #FileHandle, "MonsterText="
  For c = 0 To TotalMonsters
    Write #FileHandle, Monster(c).Text
  Next c
  If TotalMonsters = -1 Then
    'If there are no monsters, insert placeholders
    Write #FileHandle, -1, -1
  End If
  Write #FileHandle, "MonsterWeapon="
  For c = 0 To TotalMonsters
    Write #FileHandle, Monster(c).Weapon
  Next c
  If TotalMonsters = -1 Then
    'If there are no monsters, insert placeholders
    Write #FileHandle, -1, -1
  End If
  Write #FileHandle, "MonsterHealth="
  For c = 0 To TotalMonsters
    Write #FileHandle, Monster(c).Health
  Next c
  If TotalMonsters = -1 Then
    'If there are no monsters, insert placeholders
    Write #FileHandle, -1, -1
  End If
  Write #FileHandle, "MonsterStart="
  For c = 0 To TotalMonsters
    Write #FileHandle, Monster(c).StartX, Monster(c).StartY
  Next c
  If TotalMonsters = -1 Then
    'If there are no monsters, insert placeholders
    Write #FileHandle, -1, -1
  End If
  Write #FileHandle, "MonsterIndex="
  For c = 0 To TotalMonsters
    Write #FileHandle, Monster(c).Index
  Next c
  If TotalMonsters = -1 Then
    'If there are no monsters, insert placeholders
    Write #FileHandle, -1, -1
  End If
  Write #FileHandle, "MonsterImage="
  For c = 0 To TotalMonsters
    Write #FileHandle, Monster(c).Image
  Next c
  If TotalMonsters = -1 Then
    'If there are no monsters, insert placeholders
    Write #FileHandle, -1, -1
  End If

  
  Write #FileHandle, "Sign="
  For c = 0 To TotalSigns
    Write #FileHandle, Sign(c).Caption, Sign(c).Choice1, Sign(c).Choice2, Sign(c).Image, Sign(c).Sound, Sign(c).Text, Sign(c).Visits
  Next c
  If TotalSigns = -1 Then
    'If there are no signs, insert placeholders
    Write #FileHandle, "", "", "", "", "", "", -1
  End If
  Write #FileHandle, "House="
  For c = 0 To TotalHouses
    Write #FileHandle, House(c).Caption, House(c).Choice1, House(c).Choice2, House(c).Image, House(c).Sound, House(c).Text, House(c).Visits
  Next c
  If TotalHouses = -1 Then
    'If there are no signs, insert placeholders
    Write #FileHandle, "", "", "", "", "", "", -1
  End If
  
  Write #FileHandle, "PlayerText="
  For c = 0 To TotalPlayers
    Write #FileHandle, Player(c).Text
  Next c
  Write #FileHandle, "PlayerWeapon="
  For c = 0 To TotalPlayers
    Write #FileHandle, Player(c).Weapon
  Next c
  Write #FileHandle, "PlayerHealth="
  For c = 0 To TotalPlayers
    Write #FileHandle, Player(c).Health
  Next c
  Write #FileHandle, "PlayerStart="
  For c = 0 To TotalPlayers
    Write #FileHandle, Player(c).StartX, Player(c).StartY
  Next c
  Write #FileHandle, "PlayerIndex="
  For c = 0 To TotalPlayers
    Write #FileHandle, Player(c).Index
  Next c
  Write #FileHandle, "PlayerImage="
  For c = 0 To TotalPlayers
    Write #FileHandle, Player(c).Image
  Next c

  
  
  
Close #FileHandle





'This shows that it is opened, and can be saveed w/o saveas
FirstSave = False
JustSaved = True

Me.Caption = "RPG Map Editor - " & FileName & ".rpg"
Exit Sub
errh:
MsgBox Err & " " & Error$
Stop
Resume
End Sub

Private Sub mnuSaveAs_Click()
On Error GoTo errh
Dim cutfile As Integer
Dim cutfile2 As Integer
'ComDialog.FileName = FileName
'ComDialog.Filter = "RPG Maps (.rpg)|*.rpg"
'ComDialog.InitDir = App.Path & "\maps\"
'ComDialog.CancelError = True
'ComDialog.ShowSave



FirstSave = False
'FileName = ComDialog.FileName
FileName = App.Path & "\maps\" & "arthur.rpg"
'This finds and separates the name from path
cutfile = 0
cutfile2 = 1
While cutfile2 > 0
  cutfile2 = InStr(cutfile + 1, FileName, "\")
  If cutfile2 <> 0 Then cutfile = cutfile2 ' +CutFile
Wend
filepath = Mid(FileName, 1, cutfile)
FileName = Mid(FileName, cutfile + 1)
mnuSave_Click
Exit Sub
errh:
If Err = 32755 Then
  Exit Sub
Else
  MsgBox Err & ", " & Error$
  Stop
  Resume
End If
End Sub

Private Sub mnuWater_Click(Index As Integer)
SetTiles 8
Terrain_Click Index
End Sub

Private Sub ClickTile(X1 As Single, Y1 As Single)
   Dim rc
   X = Int(X1 / Tile.PixelWidth)
   Y = Int(Y1 / Tile.PixelHeight)
   If Not SetBounds Then
      mnuSave.Enabled = False
      mnuSaveAs.Enabled = False
      'rc = BitBlt(ContainPic.hDC, X * Tile.PixelWidth, Y * Tile.PixelHeight, Tile.PixelWidth, Tile.PixelHeight, ContainPic.hDC, 0, 0, SRCCOPY)
      MonCurIndex = Int((Y * MapX - MapX) + X)
      
      MonCurX = X
      MonCurY = Y
      MakePlayer.Show
      
    Else
      If SetFirst And SetLast Then
        Outline2(0).Visible = True
        Outline2(0).Left = Tile.PixelWidth * X
        Outline2(0).Top = Tile.PixelHeight * Y
        Outline2(0).ZOrder
        'Shape2(0).Width = Tile(Index).Width
        'Outline2(0).Height = Tile(Index).Height
        'Outline2Pic(0).Height = Tile(Index).Height
        'Shape2(0).Height = Tile(Index).Height
        MonMinX1 = X
        MonMinY1 = Y
        StatusBar1.Panels(3).Text = MonMinX1 & "," & MonMinY1 & "-"
      ElseIf SetLast Then
        If MonMinX1 = X And MonMinY1 = Y Then Exit Sub
        Outline2(1).Visible = True
        Outline2(1).Left = Tile.PixelWidth * X
        Outline2(1).Top = Tile.PixelHeight * Y
        Outline2(1).ZOrder
        MonMaxX2 = X
        MonMaxY2 = Y
        SetFirst = True
        SetLast = False
        MakePlayer.Visible = True
        StatusBar1.Panels(3).Text = MonMinX1 & "," & MonMinY1 & "-" & MonMaxX2 & "," & MonMaxY2
      End If
      SetFirst = False
    End If
End Sub







Private Sub mnuBaseTerrain_Click(Index As Integer)
IsCharTile = False
CurTile.Picture = Tilefrm.BaseTerrain(TerrainVal(Right(mnuBaseTerrain(Index).Caption, 4), TagStringB)).Picture
' Set Tag = the tag of the terrain
CurTile.Tag = Tilefrm.BaseTerrain(TerrainVal(Right(mnuBaseTerrain(Index).Caption, 4), TagStringB)).Tag

'Press the Terrain toolbar button
Toolbar1.Buttons(8).Value = tbrPressed
Toolbar1_ButtonClick Toolbar1.Buttons(8)

End Sub

Private Sub mnuWall_Click(Index As Integer)
IsCharTile = False
CurTile.Picture = Tilefrm.Terrain(TerrainVal(Right(mnuWall(Index).Caption, 4), TagString)).Picture
CurTile.Tag = Tilefrm.Terrain(TerrainVal(Right(mnuWall(Index).Caption, 4), TagString)).Tag
chkShowTiles(2).Value = 1
For c = 0 To listTiles.ListCount - 1
  If listTiles.List(c) = mnuWall(Index).Caption Then
    listTiles.ListIndex = c
  End If
Next c
'Press the Tiles toolbar button
'Toolbar1.Buttons(7).Value = tbrPressed
'Toolbar1.Buttons(8).Value = tbrUnpressed
'Toolbar1.Buttons(9).Value = tbrUnpressed
'Toolbar1_ButtonClick Toolbar1.Buttons(7)
End Sub

Private Sub mnuWeapon_Click(Index As Integer)
IsCharTile = False
CurTile.Picture = Tilefrm.Terrain(TerrainVal(Right(mnuWeapon(Index).Caption, 4), TagString)).Picture
CurTile.Tag = Tilefrm.Terrain(TerrainVal(Right(mnuWeapon(Index).Caption, 4), TagString)).Tag
chkShowTiles(3).Value = 1
For c = 0 To listTiles.ListCount - 1
  If listTiles.List(c) = mnuWeapon(Index).Caption Then
    listTiles.ListIndex = c
  End If
Next c
'Press the Tiles toolbar button
Toolbar1.Buttons(7).Value = tbrPressed
Toolbar1.Buttons(6).Value = tbrUnpressed
Toolbar1.Buttons(8).Value = tbrUnpressed
Toolbar1_ButtonClick Toolbar1.Buttons(7)
End Sub

Private Sub TempPic_Click()
MsgBox TempPic.Tag
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button)
On Error GoTo errh

If Button.ToolTipText = "Terrain" Then
  If frameTool(1).Top <> 0 Then frameTool(1).Top = 0
  frameTool(1).Visible = True
  frameTool(2).Visible = False
  frameTool(3).Visible = False
  chkErase = (chkErase / Abs(chkErase)) * 1 'Takes chkErase, divides it by it's abs, so that way it becomes '1', but keeps it's sign. then, it multiplies by the # to make it erase that thing
  'click on the current image
  imgCurTileb_Click 3
ElseIf Button.ToolTipText = "All Tiles" Then
  If frameTool(2).Top <> 0 Then frameTool(2).Top = 0
  frameTool(1).Visible = False
  frameTool(2).Visible = True
  frameTool(3).Visible = False
  chkErase = (chkErase / Abs(chkErase)) * 2
  'click on the current image
  imgCurTile_Click 3
ElseIf Button.ToolTipText = "Characters" Then
  If frameTool(3).Top <> 0 Then frameTool(3).Top = 0
  frameTool(1).Visible = False
  frameTool(2).Visible = False
  frameTool(3).Visible = True
  chkErase = (chkErase / Abs(chkErase)) * 3
  'click on the current image
  imgCurTileC_Click 3 'should be imgCurTileC when i make it
ElseIf Button.ToolTipText = "Toggle Sidebar" Then
  If Button.Value = 0 Then
    'button checked, bar invisible
    'ToolWindow.Visible = False
    ToolWindow.Width = 1
  Else
    'button unchecked, bar visible
    'ToolWindow.Visible = True
    ToolWindow.Width = 3615
  End If
  Form_Resize
ElseIf Button.ToolTipText = "Eraser" Then
  ' Will erase the Tiles and leave the Terrain
  If Button.Value = tbrPressed Then
    If Toolbar1.Buttons(5).Value = tbrPressed Then
      ' Terrain
      chkErase = 1
    ElseIf Toolbar1.Buttons(6).Value = tbrPressed Then
      ' Tiles
      chkErase = 2
    ElseIf Toolbar1.Buttons(7).Value = tbrPressed Then
      ' Tiles
      chkErase = 3
    End If
  Else
    ' If its being unchecked
    chkErase = 0
  End If
ElseIf Button.ToolTipText = "New" Then
  mnuNew_Click
ElseIf Button.ToolTipText = "Open" Then
  mnuOpen_Click
ElseIf Button.ToolTipText = "Save" Then
  mnuSave_Click
Else
  MsgBox "Button Not Functional"
End If

Exit Sub
errh:
If Err = 6 Then
  'Overflow (on the eraser calculation)
  chkErase = 0
  Resume Next
End If
MsgBox Err & Error$
Stop
Resume

End Sub

Private Sub VScroll1_Change()
VScroll1_Scroll
End Sub

Private Sub VScroll1_Scroll()
ContainPic.Top = (-VScroll1.Value) + Toolbar1.Height
End Sub

Public Sub TileScreen(LoadMap As Boolean, FileHandle As Integer)
Dim T As String
Dim rc
On Error GoTo errh
If Not LoadMap Then
  
  ContainPic.Width = MapX * Tile.TwipWidth
  ContainPic.Height = MapY * Tile.TwipHeight
  ContainPic.ScaleMode = 3
End If
c = 0
'Make sure user can see whats goin on
NewMapfrm.Visible = False
Me.Show
'Close
For Y = 1 To MapY
  If LoadMap Then Line Input #FileHandle, Data
  For X = 1 To MapX
    c = c + 1
    If LoadMap Then
      rc = BitBlt(ContainPic.hDC, X * Tile.PixelWidth - Tile.PixelWidth, Y * Tile.PixelHeight - Tile.PixelHeight, Tile.TwipWidth, Tile.TwipHeight, Tilefrm.BaseTerrain((InStr(TagStringB, Mid(Data, (X * 4) - 3, 4)) - 1) / 4).hDC, 0, 0, SRCCOPY)
      Tile.XY(X, Y).Tag = Mid(Data, (X * 4) - 3, 4)
    Else
      rc = BitBlt(ContainPic.hDC, X * Tile.PixelWidth - Tile.PixelWidth, Y * Tile.PixelHeight - Tile.PixelHeight, Tile.TwipWidth, Tile.TwipHeight, Tilefrm.BaseTerrain((InStr(TagStringB, CurTile.Tag) - 1) / 4).hDC, 0, 0, SRCCOPY)
      Tile.XY(X, Y).Tag = CurTile.Tag
    End If
    'StatusBar1.Panels(4).Text = "Loading" & " [" & FileName & ".rpg" & "] - "
    'StatusBar1.Panels(5).Text = LTrim(RTrim(Str(Int(((c / (MapX * MapY)) * 100))))) & "% Done"
  Next X
Next Y
If LoadMap Then Close #FileHandle
'StatusBar1.Panels(4).Text = "Done"
'StatusBar1.Panels(5).Text = ""
On Error GoTo 0


Exit Sub
errh:
If Err = 360 Then
  Resume Next
End If
MsgBox Err & Error$
Stop
Resume
End Sub

Public Sub OpenMap()
On Error GoTo errh
Close
Dim cutfile As Integer ' This stores where the last \ is
Dim cutfile2 As Integer ' This is a temp storage
Dim rc

'Have to make sure ppl know something is happening!
Me.Show

ComDialog.InitDir = App.Path & "\maps\"
ComDialog.Filter = "RPG Maps (.rpg)|*.rpg"
ComDialog.ShowOpen
If ComDialog.FileName = "" Then
  Exit Sub
Else
  FileName = ComDialog.FileName
End If
If Mid(FileName, Len(FileName) - 3, 1) = "." Then
  FileName = Mid(FileName, 1, Len(FileName) - 4)
End If

'This finds and separates the name from path
cutfile = 0
cutfile2 = 1
While cutfile2 > 0
  cutfile2 = InStr(cutfile + 1, FileName, "\")
  If cutfile2 <> 0 Then cutfile = cutfile2 ' +CutFile
Wend
filepath = Mid(FileName, 1, cutfile)
FileName = Mid(FileName, cutfile + 1)

FileHandle = FreeFile
Open filepath & FileName & ".rpg" For Input As #FileHandle
    While Not EOF(FileHandle)
      Input #FileHandle, Data
      If Data = "size=" Then
        Input #FileHandle, MapSize
      End If
    Wend
Close #FileHandle
Open filepath & FileName & ".rpg" For Input As #FileHandle

'The tile width is always 480 unless dimentions are 135x135
Tile.TwipWidth = 480
Tile.TwipHeight = 480
Tile.PixelWidth = 480 / TwipTOPixel
Tile.PixelHeight = 480 / TwipTOPixel

'This sets the map size acording to loaded data
Select Case MapSize
  Case 0:
    MapX = 25
    MapY = 25
  Case 1:
    MapX = 50
    MapY = 50
  Case 2:
    MapX = 100
    MapY = 100
    Tile.TwipWidth = 240
    Tile.TwipHeight = 240
    Tile.PixelWidth = 240 / TwipTOPixel
    Tile.PixelHeight = 240 / TwipTOPixel
  Case 3:
    MapX = 135
    MapY = 135
    Tile.TwipWidth = 240
    Tile.TwipHeight = 240
    Tile.PixelWidth = 240 / TwipTOPixel
    Tile.PixelHeight = 240 / TwipTOPixel
End Select
ReDim Tile.XY(MapX, MapY)

ContainPic.Width = MapX * Tile.PixelWidth
ContainPic.Height = MapY * Tile.PixelHeight

' This Unloads your level
ContainPic.Picture = Me.Picture

' This loads your level in
TileScreen True, FileHandle

Open filepath & FileName & ".rpg" For Input As #FileHandle

MapName = ""
MapSize = -1
DefaultTile = -1
  
  While Not EOF(FileHandle)
    Input #FileHandle, Data
    Data = LCase$(Data)
    If Data = "mapname=" Then
      Input #FileHandle, MapName
    ElseIf Data = "filename=" Then
      Input #FileHandle, FileName
    ElseIf Data = "size=" Then
      Input #FileHandle, MapSize
    ElseIf Data = "defaultterrain=" Then
      Input #FileHandle, DefaultTile
    ElseIf Data = LCase("TotalMonsters=") Then
      Input #FileHandle, TotalMonsters
    ElseIf Data = LCase("TotalPlayers=") Then
      Input #FileHandle, TotalPlayers
    ElseIf Data = LCase("TotalSigns=") Then
      Input #FileHandle, TotalSigns
    ElseIf Data = LCase("TotalHouses=") Then
      Input #FileHandle, TotalHouses
    ElseIf Data = LCase("MapTopX=") Then
      Input #FileHandle, MapTopX
    ElseIf Data = LCase("MapTopY=") Then
      Input #FileHandle, MapTopY
    ElseIf Data = LCase("Max=") Then
      For c = 0 To TotalMonsters
        Input #FileHandle, Monster(c).MaxX, Monster(c).MaxY
      Next c
    ElseIf Data = LCase("Min=") Then
      For c = 0 To TotalMonsters
        Input #FileHandle, Monster(c).MinX, Monster(c).MinY
      Next c
    ElseIf Data = LCase("MonsterText=") Then
      For c = 0 To TotalMonsters
        Input #FileHandle, Monster(c).Text
      Next c
    ElseIf Data = LCase("MonsterWeapon=") Then
      For c = 0 To TotalMonsters
        Input #FileHandle, Monster(c).Text
      Next c
    ElseIf Data = LCase("monsterHealth=") Then
      For c = 0 To TotalMonsters
        Input #FileHandle, Monster(c).Health
      Next c
    ElseIf Data = LCase("MonsterStart=") Then
      For c = 0 To TotalMonsters
        Input #FileHandle, Monster(c).StartX, Monster(c).StartY
      Next c
    ElseIf Data = LCase("MonsterIndex=") Then
      For c = 0 To TotalMonsters
        Input #FileHandle, Monster(c).Index
      Next c
    ElseIf Data = LCase("Monsterimage=") Then
      For c = 0 To TotalMonsters
        Input #FileHandle, Monster(c).Image
        rc = BitBlt(ContainPic.hDC, Monster(c).Index - Int(MapX - Monster(c).Index), Int(Monster(c).Index / MapY), Tile.PixelWidth, Tile.PixelHeight, Tilefrm.Characters(Monster(c).Image), 0, 0, SRCCOPY)
        Tile.XY(Monster(c).Index - Int(MapX - Monster(c).Index), Int(Monster(c).Index / MapY)).SpecialTile = "Monster #" & c
      Next c
  
    ElseIf Data = LCase("Sign=") Then
      For c = 0 To TotalSigns
        Input #FileHandle, Sign(c).Caption, Sign(c).Choice1, Sign(c).Choice2, Sign(c).Image, Sign(c).Sound, Sign(c).Text, Sign(c).Visits
      Next c
    ElseIf Data = LCase("House=") Then
      For c = 0 To TotalHouses
        Input #FileHandle, House(c).Caption, House(c).Choice1, House(c).Choice2, House(c).Image, House(c).Sound, House(c).Text, House(c).Visits
      Next c
    
    ElseIf Data = LCase("PlayerText=") Then
      For c = 0 To TotalPlayers
        Input #FileHandle, Player(c).Text
      Next c
    ElseIf Data = LCase("playerweapon=") Then
      For c = 0 To TotalPlayers
        Input #FileHandle, Player(c).Weapon
      Next c
    ElseIf Data = LCase("playerHealth=") Then
      For c = 0 To TotalMonsters
        Input #FileHandle, Player(c).Health
      Next c
    ElseIf Data = LCase("PlayerStart=") Then
      For c = 0 To TotalPlayers
        Input #FileHandle, Player(c).StartX, Player(c).StartY
      Next c
    ElseIf Data = LCase("PlayerIndex=") Then
      For c = 0 To TotalPlayers
        Input #FileHandle, Player(c).Index
      Next c
    ElseIf Data = LCase("Playerimage=") Then
      For c = 0 To TotalPlayers
        Input #FileHandle, Player(c).Image
        rc = BitBlt(ContainPic.hDC, Player(c).Index - Int(MapX - Player(c).Index), Int(Player(c).Index / MapY), Tile.PixelWidth, Tile.PixelHeight, Tilefrm.Characters(Player(c).Image), 0, 0, SRCCOPY)
        Tile.XY(Player(c).Index - Int(MapX - Player(c).Index), Int(Player(c).Index / MapY)).SpecialTile = "Player #" & c
      Next c
        
  
  
  
    End If
  Wend
  
  
  
  
  
  
  
  
  
  
Close #FileHandle


EraseTile.Picture = Tilefrm.Terrain(DefaultTile)

Me.Caption = "RPG Map Editor - " & FileName & ".rpg"

Exit Sub
'----------------------------------------
errh:
If Err = 53 Then
  R = MsgBox("Error! Map Information not found! Map might not load correctly. I'll try to load it anyway. Hint: Make sure the .rpg file is in the same directory as the map.", vbCritical, "File Not Found")
  Open FileName & ".rpg" For Output As #FileHandle
    Write #FileHandle, "", ""
  Close #FileHandle
  Resume
ElseIf Err = 340 Then
  Resume Next
ElseIf Err = 35722 Then
  Exit Sub
Else
  MsgBox Err & Error$
  Stop
  Resume
End If
End Sub




'-------------------------Tool WIndow---------------------------------------------
'---------------------------------------------------------------------------------


Public Sub SaveMakePlayerData()
On Error GoTo errh

Dim Temp1
StatusBar1.Panels(2).Text = ""
If MakePlayer.Option1(0).Value = True Then
  TotalPlayers = TotalPlayers + 1
  Player(TotalPlayers).Text = MakePlayer.CharMsgText
  Player(TotalPlayers).Weapon = MakePlayer.CharWeapon.ListIndex
  Player(TotalPlayers).Health = MakePlayer.CharHitPts
  Player(TotalPlayers).StartX = MonCurX - 1
  Player(TotalPlayers).StartY = MonCurY - 1
  Player(TotalPlayers).Image = MonCurImage
  Player(TotalPlayers).Index = MonCurIndex
  Tile.XY(MonCurX, MonCurY).SpecialTile = "Player #" & TotalPlayers
    'player
  Load shapePlayer(TotalPlayers)
  shapePlayer(TotalPlayers).Visible = True
  shapePlayer(TotalPlayers).Left = (MonCurX - 1) * Tile.PixelWidth
  shapePlayer(TotalPlayers).Top = (MonCurY - 1) * Tile.PixelHeight
Else
  TotalMonsters = TotalMonsters + 1
  Monster(TotalMonsters).MinX = MonMinX1
  Monster(TotalMonsters).MinY = MonMinY1
  Monster(TotalMonsters).MaxX = MonMaxX2
  Monster(TotalMonsters).MaxY = MonMaxY2
  Monster(TotalMonsters).Index = MonCurIndex
  MonMinX1 = -1
  Monster(TotalMonsters).Image = MonCurImage
  Monster(TotalMonsters).Health = MakePlayer.CharHitPts
  Monster(TotalMonsters).Text = MakePlayer.CharMsgText
  Monster(TotalMonsters).Weapon = MakePlayer.CharWeapon.ListIndex
  Monster(TotalMonsters).StartX = MonCurX - 1
  Monster(TotalMonsters).StartY = MonCurY - 1
  Tile.XY(MonCurX, MonCurY).SpecialTile = "Monster #" & TotalMonsters
    'monster
  Load shapeMonster(TotalMonsters)
  shapeMonster(TotalMonsters).Visible = True
  shapeMonster(TotalMonsters).Left = MonCurX * Tile.PixelWidth
  shapeMonster(TotalMonsters).Top = MonCurY * Tile.PixelHeight
End If
Outline2(0).Visible = False
Outline2(1).Visible = False
mnuSave.Enabled = True
mnuSaveAs.Enabled = True
SetBounds = False
Unload MakePlayer
Exit Sub
errh:
If Err = 360 Then Resume Next 'already loaded
MsgBox Err & Err.Description
Stop
Resume
End Sub

Private Sub KillCharacter(X1 As Single, Y1 As Single)
'Tile(Index).Picture = Terrain(((InStr(TagString, c)) - 1) / 4).Picture
If Left(Tile(Index).ToolTipText, 6) = "Player" Then
  For c = Val(Mid(Tile(Index).ToolTipText, 9)) To TotalPlayers - 1
    Player(c).Text = Player(c + 1).Text
    Player(c).Weapon = Player(c + 1).Weapon
    Player(c).StartX = Player(c + 1).StartX
    Player(c).StartY = Player(c + 1).StartY
  Next c
  Player(c).Text = ""
  Player(c).Weapon = ""
  Player(c).StartX = -1
  Player(c).StartY = -1
  Tile(Index).ToolTipText = ""
  TotalPlayers = TotalPlayers - 1
ElseIf Left(Tile(Index).ToolTipText, 7) = "Monster" Then
  For c = Val(Mid(Tile(Index).ToolTipText, 9)) To TotalMonsters - 1
    Monster(c).Text = Monster(c + 1).Text
    Monster(c).Weapon = Monster(c + 1).Weapon
    Monster(c).StartX = Monster(c + 1).StartX
    Monster(c).StartY = Monster(c + 1).StartY
    Min(c - 1, 0) = Min(c, 0)
    Min(c - 1, 1) = Min(c, 1)
    Max(c - 1, 0) = Max(c, 0)
    Max(c - 1, 1) = Max(c, 1)
  Next c
  Monster(0).Text = ""
  Monster(0).Weapon = ""
  Monster(0).StartX = -1
  Monster(0).StartY = -1
  Min(c - 1, 0) = -1
  Min(c - 1, 1) = -1
  Max(c - 1, 0) = -1
  Max(c - 1, 1) = -1
  Tile(Index).ToolTipText = ""
  TotalMonsters = TotalMonsters - 1
End If
End Sub

'------------------------------------------------------------------------
'----------------------------Properties----------------------------------

Public Sub btnCreateDelete_Click_(Index As Integer)
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

If LCase(MakePlayer.CharWeapon.Text) <> "club" And LCase(MakePlayer.CharWeapon.Text) <> "dagger" And LCase(MakePlayer.CharWeapon.Text) <> "sword" Then
  R = MsgBox("Invalid choice!", vbCritical, "Invalid Choice")
  MakePlayer.CharWeapon.SetFocus
  MakePlayer.CharWeapon.SelStart = 1
  MakePlayer.CharWeapon.SelLength = Len(MakePlayer.CharWeapon.Text)
  Exit Sub
ElseIf Val(MakePlayer.CharHitPts.Text) < 1 Or Val(MakePlayer.CharHitPts.Text) > 3 Then
  R = MsgBox("Invalid choice!", vbCritical, "Invalid Choice")
  MakePlayer.CharHitPts.SetFocus
  MakePlayer.CharHitPts.SelStart = 1
  MakePlayer.CharHitPts.SelLength = Len(MakePlayer.CharWeapon.Text)
  Exit Sub
ElseIf MakePlayer.Option1(1).Value And (MapEdit.MonMinX1 < 0 Or (MapEdit.MonMinY1 = Empty And Not MapEdit.MonMinY1 = 0)) Then
  R = MsgBox("You must first set the monster's range!", vbCritical, "Set range")
'  MakePlayer.Command3.SetFocus
  Exit Sub
ElseIf MakePlayer.Option1(1).Value And (MapEdit.MonMinX1 > MapEdit.MonCurX Or MapEdit.MonMinY1 > MapEdit.MonCurY Or MapEdit.MonMaxX2 < MapEdit.MonCurX Or MapEdit.MonMaxY2 < MapEdit.MonCurY) Then
    R = MsgBox("The monster's range must include the monster!! Reset it's range.", vbCritical, "Invalid range")
    MakePlayer.Command3.SetFocus
    Exit Sub
End If
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
End Sub

Public Sub btnMonsterLeft_Click_()

End Sub

Public Sub btnMonsterRight_Click_()

End Sub

Public Sub btnPlayerLeft_Click_()

End Sub

Public Sub btnPlayerRight_Click_()

End Sub

Public Sub btnRedefine_Click_()

End Sub

Public Sub btnSave_Click_(Index As Integer)

End Sub

Public Sub Form_Load_()
Properties.MapName = MapEdit.MapName
Properties.MapDesc = MapEdit.MapDesc
Properties.lblMapSize = MapEdit.MapX & "x" & MapEdit.MapY
Properties.lblNumMonsters = MapEdit.TotalMonsters + 1 'Add one so it shows the correct amount
Properties.lblNumPlayers = MapEdit.TotalPlayers + 1 'Add one so it shows the correct amount


' Disable all the non-exsistant players
For c = 0 To 7
  If Player(c).Health = 0 Then Properties.POption(c).Enabled = False
Next c
For c = 1 To TotalMonsters
  Properties.cmboMonster.AddItem c
Next c

' Select the appropriate tab/monster
If Left(CurMon, 6) = "Player" Then
  MapEdit.POption_Click_ Int(Mid(CurMon, 9))
ElseIf Left(CurMon, 7) = "Monster" Then
  Properties.cmboMonster.Text = Mid(CurMon, 10)
  MapEdit.cmboMonster_Click_ Mid(CurMon, 10)
End If
End Sub

Public Sub cmboMonster_Click_(Index As String)
  
If Properties.MOption(Index).Enabled Then
  Properties.MonsterX = Monster(0).StartX
  Properties.MonsterY = Monster(0).StartY
  Properties.MonsterHealth = Monster(0).Health
  Properties.MonsterWeapon = Monster(0).Weapon
  Properties.XMin = Monster(0).MinX
  Properties.YMin = Monster(0).MinY
  Properties.XMax = Monster(0).MaxX
  Properties.YMax = Monster(0).MaxY
  Properties.MonsterText = Monster(0).Text
  Properties.MonsterImage = Tilefrm.Characters(Monster(0).Image).Picture
End If

End Sub

Public Sub POption_Click_(Index As Integer)
If Properties.POption(Index).Enabled Then
  Properties.PlayerX = Player(Index).StartX
  Properties.PlayerY = Player(Index).StartY
  Properties.PlayerHealth = Player(Index).Health
  Properties.PlayerWeapon = Player(Index).Weapon
  Properties.PlayerText = Player(Index).Text
  Properties.PlayerImage.Picture = Tilefrm.Characters(Index).Picture
End If
End Sub
'----------------------------Properties----------------------------------
'------------------------------------------------------------------------

Public Sub ClearArrays()
For c = 0 To MAXPLAYERS
  Player(c).Health = 0
  Player(c).Image = -1
  Player(c).Index = 0
  Player(c).StartX = -1
  Player(c).StartY = -1
  Player(c).Text = ""
  Player(c).Weapon = 0
Next c
For c = 0 To TotalMonsters
  Monster(c).Health = 0
  Monster(c).Image = -1
  Monster(c).Index = 0
  Monster(c).StartX = -1
  Monster(c).StartY = -1
  Monster(c).Text = ""
  Monster(c).Weapon = 0
  Monster(c).MaxX = -1
  Monster(c).MaxY = -1
  Monster(c).MinX = -1
  Monster(c).MinY = -1
Next c
MapEdit.TotalHouses = -1
MapEdit.TotalMonsters = -1
MapEdit.TotalPlayers = -1
MapEdit.TotalSigns = -1
For c = 0 To shapeMonster.Count - 1
  shapeMonster(c).Visible = False
Next c
For c = 0 To shapePlayer.Count - 1
  shapePlayer(c).Visible = False
Next c
End Sub


'---------------------------------------------------------------------------
'-------------------------NewNapfrm-----------------------------------------
Public Sub Command1N_Click()
OpenMap
End Sub

Public Sub Command2N_Click()
FileName = NewMapfrm.MapName.Text
MapName = NewMapfrm.MapName.Text
MapX = NewMapfrm.MapX
MapY = NewMapfrm.MapY
ReDim Tile.XY(MapX, MapY)
TileScreen False, 0
FirstSave = True
ClearArrays
Form_Resize
NewMapfrm.Hide
End Sub


Public Sub Form_UnloadN(Cancel As Integer)
Command2_Click
End Sub

Public Sub Option_ClickN(Index As Integer)
MapSize = Index
' the '/ TwipTOPixel ' converts it to pixels as opposed to twips
Tile.PixelWidth = 480 / TwipTOPixel
Tile.PixelHeight = 480 / TwipTOPixel
Tile.TwipWidth = 480
Tile.TwipHeight = 480
Select Case Index
  Case 0:
    NewMapfrm.MapX = 25
    NewMapfrm.MapY = 25
  Case 1:
    NewMapfrm.MapX = 50
    NewMapfrm.MapY = 50
  Case 2:
    NewMapfrm.MapX = 100
    NewMapfrm.MapY = 100
    'Tile.TwipWidth = 240
    'Tile.TwipHeight = 240
    'Tile.PixelWidth = 240 / TwipTOPixel
    'Tile.PixelHeight = 240 / TwipTOPixel
  Case 3:
    NewMapfrm.MapX = 135
    NewMapfrm.MapY = 135
    'Tile.TwipWidth = 240
    'Tile.TwipHeight = 240
    'Tile.PixelWidth = 240 / TwipTOPixel
    'Tile.PixelHeight = 240 / TwipTOPixel
End Select
End Sub

Public Sub Option1_ClickN(Index As Integer)
CurTile.Picture = Tilefrm.BaseTerrain(Index).Picture
CurTile.Tag = Tilefrm.BaseTerrain(Index).Tag
EraseTile.Picture = Tilefrm.BaseTerrain(Index).Picture
EraseTile.Tag = Tilefrm.BaseTerrain(Index).Tag
DefaultTile = Index
End Sub

'-------------------------NewNapfrm-----------------------------------------
'---------------------------------------------------------------------------

Public Sub TileChange(X As Single, Y As Single)
    Dim X1
    Dim Y1
    Dim rc
    X1 = Int(X / Tile.TwipWidth) * Tile.TwipWidth
    Y1 = Int(Y / Tile.PixelHeight) * Tile.PixelHeight
    X = Int(X / Tile.PixelWidth)
    Y = Int(Y / Tile.PixelHeight)
    'For grass
    If Left(Tile.XY(X + 1, Y + 1).Tag, 1) = "G" Then
      'above
      If Tile.XY(X + 1, Y).Tag = "D00_" Then
        rc = BitBlt(ContainPic.hDC, X1, Y1 - Tile.PixelWidth, Tile.PixelWidth, Tile.PixelHeight, Tilefrm.Terrain(TerrainVal("F00_")), 0, 0, SRCCOPY)
        rc = BitBlt(ContainPic.hDC, X1, Y1 - Tile.TwipWidth, Tile.PixelWidth, Tile.PixelHeight, Tilefrm.Black(TerrainVal(" 02_")), 0, 0, SRCAND)
        rc = BitBlt(ContainPic.hDC, X1, Y1 - Tile.TwipWidth, Tile.PixelWidth, Tile.PixelHeight, Tilefrm.Terrain(TerrainVal(" 02_")), 0, 0, SRCPAINT)
      End If
   End If
End Sub

Public Function TerrainVal(Tag As String, Optional TagStringVal As String = "%NONE%")
    If TagStringVal = "%NONE%" Then TagStringVal = TagString
    TerrainVal = InStr(1, TagStringVal, Tag) - 1
    TerrainVal = TerrainVal / 4
End Function


Public Function GetTileXY(InfoType As String, X1 As Integer, Y1 As Integer, Optional SpecialTile_TileText_Tag_Picture) As String
  If InfoType = "SpecialTile" Then
    GetTileXY = Tile.XY(X1, Y1).SpecialTile
  ElseIf InfoType = "TileText" Then
    GetTileXY = Tile.XY(X1, Y1).TileText
  ElseIf InfoType = "Tag" Then
    GetTileXY = Tile.XY(X1, Y1).Tag
  ElseIf InfoType = "Picture" Then
    GetTileXY = Tile.XY(X1, Y1).Picture
  End If
End Function

Sub CreateMenu()

On Error GoTo errh


' Creates the Base Terrain tiles
Tilefrm.File1 = App.Path & "\bitmaps\base tiles\"

For c = 0 To Tilefrm.File1.ListCount - 1
  Load mnuBaseTerrain(c)
  mnuBaseTerrain(c).Visible = True
  mnuBaseTerrain(c).Caption = Left(Tilefrm.File1.List(c), 4)  'add the 4 chr code e.g. G00_
  Select Case Left(mnuBaseTerrain(c).Caption, 1)
    Case "G"
      mnuBaseTerrain(c).Caption = "Grass " & mnuBaseTerrain(c).Caption
    Case "D"
      mnuBaseTerrain(c).Caption = "Desert " & mnuBaseTerrain(c).Caption
    Case "R"
      mnuBaseTerrain(c).Caption = "Dirt " & mnuBaseTerrain(c).Caption
    Case "W"
      mnuBaseTerrain(c).Caption = "Water " & mnuBaseTerrain(c).Caption
    Case "M"
      mnuBaseTerrain(c).Caption = "Mountain " & mnuBaseTerrain(c).Caption
    Case "F"
      mnuBaseTerrain(c).Caption = "Forest " & mnuBaseTerrain(c).Caption
    Case "0"
      mnuBaseTerrain(c).Caption = "Blank " & mnuBaseTerrain(c).Caption
    Case "S"
      mnuBaseTerrain(c).Caption = "Snow " & mnuBaseTerrain(c).Caption
    Case Else
      mnuBaseTerrain(c).Caption = "Unknown Terrain " & Chr(34) & Tilefrm.File1.List(c) & Chr(34) ' & mnuBaseTerrain(c).Caption
      mnuBaseTerrain(c).Enabled = False
  End Select
Next c
  
  
'This creates the Character menu items

Tilefrm.File1 = App.Path & "\bitmaps\characters\"

For c = 0 To Tilefrm.File1.ListCount - 1 Step 2
  Load mnuCharacter(c / 2)
  mnuCharacter(c / 2).Visible = True
  mnuCharacter(c / 2).Caption = Left(Tilefrm.File1.List(c), 4) 'add the 3 chr code e.g. C00
  Select Case Left(mnuCharacter(c / 2).Caption, 1)
    Case "C"
      mnuCharacter(c / 2).Caption = "Character " & mnuCharacter(c / 2).Caption
    Case Else
      mnuCharacter(c / 2).Caption = "Unknown Character " & Chr(34) & Tilefrm.File1.List(c) & Chr(34) ' & " " & mnuCharacter(c)
      mnuCharacter(c / 2).Enabled = False
  End Select
Next c


' This creates the Terrain menus

Tilefrm.File1 = App.Path & "\bitmaps\tiles\"

' This is the "All Tiles" menu
For c = 0 To Tilefrm.File1.ListCount - 1 Step 2
  Load mnuAllTiles(c / 2)
  mnuAllTiles(c / 2).Visible = True
  mnuAllTiles(c / 2).Caption = Left(Tilefrm.File1.List(c), 3)
  mnuAllTiles(c / 2).Caption = " " & mnuAllTiles(c / 2).Caption
  mnuAllTiles(c / 2).Caption = FindTileType(mnuAllTiles(c / 2).Caption) & " " & mnuAllTiles(c / 2).Caption
Next c

R = 0
' This is the "Paths" menu
For c = 0 To Tilefrm.File1.ListCount - 1 Step 2
  If Mid(Left(Tilefrm.File1.List(c), 3), 1, 1) = "1" Or Mid(Left(Tilefrm.File1.List(c), 3), 1, 1) = "3" Or Mid(Left(Tilefrm.File1.List(c), 3), 1, 1) = "4" Then ' These are the ones that are paths
    Load mnuPaths(R)
    mnuPaths(R).Visible = True
    mnuPaths(R).Caption = Left(Tilefrm.File1.List(c), 3)
    mnuPaths(R).Caption = " " & mnuPaths(R).Caption
    mnuPaths(R).Caption = FindTileType(mnuPaths(R).Caption) & " " & mnuPaths(R).Caption
    R = R + 1
  End If
Next c

R = 0
' This is the "Walls" menu
For c = 0 To Tilefrm.File1.ListCount - 1 Step 2
  If Mid(Left(Tilefrm.File1.List(c), 3), 1, 1) = "5" Or Mid(Left(Tilefrm.File1.List(c), 3), 1, 1) = "5" Or Mid(Left(Tilefrm.File1.List(c), 3), 1, 1) = "5" Then ' These are the ones that are paths
    Load mnuWall(R)
    mnuWall(R).Visible = True
    mnuWall(R).Caption = Left(Tilefrm.File1.List(c), 3)
    mnuWall(R).Caption = " " & mnuWall(R).Caption
    mnuWall(R).Caption = FindTileType(mnuWall(R).Caption) & " " & mnuWall(R).Caption
    R = R + 1
  End If
Next c

R = 0
' This is the "Weapons" menu
For c = 0 To Tilefrm.File1.ListCount - 1 Step 2
  If Mid(Left(Tilefrm.File1.List(c), 3), 1, 1) = "w" Or Mid(Left(Tilefrm.File1.List(c), 3), 1, 1) = "w" Or Mid(Left(Tilefrm.File1.List(c), 3), 1, 1) = "w" Then ' These are the ones that are paths
    Load mnuWeapon(R)
    mnuWeapon(R).Visible = True
    mnuWeapon(R).Caption = Left(Tilefrm.File1.List(c), 3)
    mnuWeapon(R).Caption = " " & mnuWeapon(R).Caption
    mnuWeapon(R).Caption = FindTileType(mnuWeapon(R).Caption) & " " & mnuWeapon(R).Caption
    R = R + 1
  End If
Next c

R = 0
' This is the "Objects" menu
For c = 0 To Tilefrm.File1.ListCount - 1 Step 2
  If Mid(Left(Tilefrm.File1.List(c), 3), 1, 1) = "s" Or Mid(Left(Tilefrm.File1.List(c), 3), 1, 1) = "h" Or _
     Mid(Left(Tilefrm.File1.List(c), 3), 1, 1) = "t" Or Mid(Left(Tilefrm.File1.List(c), 3), 1, 1) = "l" Then
    Load mnuObject(R)
    mnuObject(R).Visible = True
    mnuObject(R).Caption = Left(Tilefrm.File1.List(c), 3)
    mnuObject(R).Caption = " " & mnuObject(R).Caption
    mnuObject(R).Caption = FindTileType(mnuObject(R).Caption) & " " & mnuObject(R).Caption
    R = R + 1
  End If
Next c





On Error GoTo 0
Exit Sub
errh:
If Err = 360 Then Resume Next
MsgBox Err & Err.Description
Stop
Resume

End Sub


Public Sub CreateList()
On Error GoTo errh

Dim c1 As Integer

' first, clear the list
listTiles.Clear

' if "All Tiles" is checked, display them all, otherwise go item by item
If chkShowTiles(0).Value = 1 Then
  For c = 0 To mnuAllTiles.Count - 1
    listTiles.AddItem mnuAllTiles(c).Caption
  Next c
Else
  For c1 = 1 To chkShowTiles.Count - 1
    If chkShowTiles(c1).Caption = "Paths" And chkShowTiles(c1).Value = 1 Then
      For c = 0 To mnuPaths.Count - 1
        listTiles.AddItem mnuPaths(c).Caption
      Next c
    ElseIf chkShowTiles(c1).Caption = "Walls" And chkShowTiles(c1).Value = 1 Then
      For c = 0 To mnuWall.Count - 1
        listTiles.AddItem mnuWall(c).Caption
      Next c
    ElseIf chkShowTiles(c1).Caption = "Weapons" And chkShowTiles(c1).Value = 1 Then
      For c = 0 To mnuWeapon.Count - 1
        listTiles.AddItem mnuWeapon(c).Caption
      Next c
    ElseIf chkShowTiles(c1).Caption = "Objects" And chkShowTiles(c1).Value = 1 Then
      For c = 0 To mnuObject.Count - 1
        listTiles.AddItem mnuObject(c).Caption
      Next c
    End If
  Next c1
End If
  
' Set the Pictures
If listTiles.ListCount > 0 Then
  listTiles.ListIndex = 0
Else
  For c = 0 To 6
    imgCurTile(c).Picture = LoadPicture()
  Next c
End If

' Set the scroll bar
If listTiles.ListCount = 0 Then
  VScroll2.Visible = False
  VScroll2.Max = listTiles.ListCount
  VScroll2.SmallChange = 1
Else
  VScroll2.Visible = True
  VScroll2.Max = listTiles.ListCount - 1
  VScroll2.SmallChange = 1
End If
VScroll2.LargeChange = VScroll2.Max / 10

'---------------------------------------------------------------------------------------
'Create the Base Tiles list

' First, Clear the list
listTilesB.Clear

'Display all the Base Tiles
For c = 0 To mnuBaseTerrain.Count - 1
  listTilesB.AddItem mnuBaseTerrain(c).Caption
Next c

' Set the Pictures
If listTilesB.ListCount > 0 Then
  listTilesB.ListIndex = 0
Else
  For c = 0 To 6
    imgCurTileb(c).Picture = LoadPicture()
  Next c
End If

' Set the scroll bar
If listTilesB.ListCount = 0 Then
  HScroll2.Visible = False
  HScroll2.Max = listTilesB.ListCount
  HScroll2.SmallChange = 1
Else
  HScroll2.Visible = True
  HScroll2.Max = listTilesB.ListCount - 1
  HScroll2.SmallChange = 1
End If
HScroll2.LargeChange = HScroll2.Max / 10

'---------------------------------------------------------------------------------------
'Create the Character Tiles list

' First, Clear the list
listTilesC.Clear

'Display all the Base Tiles
For c = 0 To mnuCharacter.Count - 1
  listTilesC.AddItem mnuCharacter(c).Caption
Next c

' Set the Pictures
If listTilesC.ListCount > 0 Then
  listTilesC.ListIndex = 0
Else
  For c = 0 To 6
    imgCurTileC(c).Picture = LoadPicture()
  Next c
End If

' Set the scroll bar
If listTilesC.ListCount = 0 Then
  HScroll3.Visible = False
  HScroll3.Max = listTilesC.ListCount
  HScroll3.SmallChange = 1
Else
  HScroll3.Visible = True
  HScroll3.Max = listTilesC.ListCount - 1
  HScroll3.SmallChange = 1
End If
HScroll3.LargeChange = HScroll3.Max / 10


Exit Sub
errh:
If Err = 380 Then Resume Next
MsgBox Err & Err.Description
Stop
Resume
End Sub

Public Function FindTileType(Tag As String) As String
If Len(Tag) = 3 Then Tag = " " & Tag    ' if its a tile, it needs a space to be detected right
If Left(Tag, 1) = "C" Then
  ' Its a character
  FindTileType = "Character"
  Exit Function
ElseIf Left(Tag, 1) = " " Then
  Select Case Mid(Tag, 2, 1)
    Case "1"
      FindTileType = "Dirt Road: "
    Case "2"
      FindTileType = "Grass Edge: " 'Took this out of the folder for now
    Case "3"
      FindTileType = "Bridge: "
    Case "4"
      FindTileType = "Windy Dirt Path: "
    Case "5"
      FindTileType = "Wall: "
    Case "6"
      FindTileType = "Tree Edge: " 'Took this out of the folder for now
    Case "h"
      FindTileType = "House: "
    Case "w"
      FindTileType = "Weapon: "
    Case "t"
      FindTileType = "Treasure: "
    Case "s"
      FindTileType = "Sign: "
    Case "l"
      FindTileType = "Landscape Item: "
    Case Else
      FindTileType = "Undefined: "
  End Select
  Dim NF As String
  If Mid(Tag, 2, 1) <> "h" And Mid(Tag, 2, 1) <> "w" And Mid(Tag, 2, 1) <> "t" And Mid(Tag, 2, 1) <> "s" And Mid(Tag, 2, 1) <> "l" Then
    Select Case Mid(Tag, 3, 1)
      Case "0"
        NF = "Crossroad"
      Case "1"
        NF = "Left to Right"
      Case "2"
        NF = "Top to Bottom"
      Case "3"
        NF = "Left to Bottom"
      Case "4"
        NF = "Left to Top"
      Case "5"
        NF = "Right to Bottom"
      Case "6"
        NF = "Right to Top"
      Case "7"
        NF = "End: Exit at Bottom"
      Case "8"
        NF = "End: Exit at Left"
      Case "9"
        NF = "End: Exit at Right"
      Case "a"
        NF = "End: Exit at Top"
      Case "b"
        NF = "T - Intersection: Road to Bottom"
      Case "c"
        NF = "T - Intersection: Road to Left"
      Case "d"
        NF = "T - Intersection: Road to Right"
      Case "e"
        NF = "T - Intersection: Road to Top"
      Case "f"
        NF = "Top Left to Bottom"
      Case "g"
        NF = "Top Right to Bottom"
      Case "h"
        NF = "Left to Bottom Right"
      Case "i"
       NF = "Left to Top Right"
      Case "j"
        NF = "Top Left to Right"
      Case "k"
        NF = "Bottom Left to Right"
      Case "l"
        NF = "Bottom Left to Top"
      Case "m"
        NF = "Top to Bottom Right"
      Case "n"
        NF = "Top to Bottom"
      Case "o"
        NF = "Left to Right"
      Case "p"
        NF = "Top Left to Bottom Right"
      Case "q"
        NF = "Bottom Left to Top Right"
      Case "r"
        NF = "Closed Gate"
      Case "s"
        NF = "Open Gate"
    End Select
    FindTileType = FindTileType & NF
  ElseIf Mid(Tag, 2, 1) = "l" Then
    ' If the second character is an 'l' then it is a landscape item. Replace "Landscape Item: " with the new text
    Select Case Mid(Tag, 3, 1)
      Case "0"
        FindTileType = "Impassible Trees"
      Case "1"
        FindTileType = "Mountains"
      Case "2"
        FindTileType = "Rock"
    End Select
    'FindTileType = FindTileType & NF
  End If
Else
  ' If it is a base terrain tile
  Select Case Left(Tag, 1)
    Case "G"
      FindTileType = "Grass "
    Case "D"
      FindTileType = "Desert "
    Case "R"
      FindTileType = "Dirt "
    Case "W"
      FindTileType = "Water "
    Case "M"
      FindTileType = "Mountain "
    Case "S"
      FindTileType = "Snow "
    Case "F"
      FindTileType = "Forest "
    Case "0"
      FindTileType = "Blank "
    Case Else
      FindTileType = "Unknown "
  End Select
  'FindTileType = FindTileType & Mid(Tag, 2, 2)
End If


End Function

Public Function DrawMask(X As Single, Y As Single, BaseTile As String, Optional TopTile As String = "", Optional TagStringTop As String = "", Optional TagStringBase As String = "%TagStringBase%") As String
On Error GoTo errh

Dim rc

If TagStringBase = "%TagStringBase%" Then
  'This means it is default
  TagStringBase = TagStringB
End If

' If the length of BaseTile is one, then add the appropriate suffix
If Len(BaseTile) = 1 Then
  BaseTile = Tilefrm.BaseTerrain(TerrainVal(BaseTile, TagStringB)).Tag
End If

' Redraw base
rc = BitBlt(ContainPic.hDC, Int(X / Tile.PixelWidth) * Tile.PixelWidth, Int(Y / Tile.PixelHeight) * Tile.PixelHeight, Tile.PixelWidth, Tile.PixelHeight, Tilefrm.BaseTerrain(TerrainVal(Left(BaseTile, 1), TagStringBase)).hDC, 0, 0, SRCCOPY)
If TopTile <> "" Then ' If TopTile = "" then that means there is no mask to draw... just the base tile.
  
  ' Make sure TopTile has a space, not a Terrain Modifier
  TopTile = " " & Mid(TopTile, 2)
  ' Draw mask
  rc = BitBlt(ContainPic.hDC, Int(X / Tile.PixelWidth) * Tile.PixelWidth, Int(Y / Tile.PixelHeight) * Tile.PixelHeight, Tile.PixelWidth, Tile.PixelHeight, Tilefrm.Black(TerrainVal(TopTile, TagStringTop)).hDC, 0, 0, SRCAND)
  ' Draw the top tile
  rc = BitBlt(ContainPic.hDC, Int(X / Tile.PixelWidth) * Tile.PixelWidth, Int(Y / Tile.PixelHeight) * Tile.PixelHeight, Tile.PixelWidth, Tile.PixelHeight, Tilefrm.Terrain(TerrainVal(TopTile, TagStringTop)).hDC, 0, 0, SRCOR)
  
  ' set DrawMask equal to the new tile
  DrawMask = Left(BaseTile, 1) & Mid(TopTile, 2)
Else
  ' set DrawMask equal to the Terrain Tile
  DrawMask = BaseTile
End If



Exit Function
errh:
MsgBox Err & Err.Description
Stop
Resume
End Function

Private Sub VScroll2_Change()
VScroll2_Scroll
End Sub

Private Sub VScroll2_Scroll()
On Error Resume Next
listTiles.ListIndex = VScroll2.Value
End Sub

Private Sub CreateToolbar()
On Error GoTo errh

Dim d As Integer
Dim c1 As Integer


' Create the Toolbar
'Toolbar1.ImageList = ImageList

c = 0
'count thru each button, but only increment an extra c if its a separator
For c1 = 1 To 0 'Toolbar1.Buttons.Count
  If True Then ' Toolbar1.Buttons(c1).Style = 3 Then
    c1 = c1 + 1 'separator
  End If
  c = c + 1
  'Toolbar1.Buttons(c1).ToolTipText = ImageList.ListImages(c).Tag
  'Toolbar1.Buttons(c1).Image = c
Next c1

' Click the first button
frameTool(1).Visible = True

Exit Sub
errh:
MsgBox Err & Err.Description
Stop
Resume
End Sub

Private Sub CreateImageList()
On Error GoTo errh

Tilefrm.File1 = App.Path & "\Bitmaps\Toolbar\"

For c = 1 To Tilefrm.File1.ListCount
  'ImageList.ListImages.Add c, "", LoadPicture(App.Path & "\Bitmaps\Toolbar\" & Tilefrm.File1.List(c - 1))
  'ImageList.ListImages(c).Tag = Mid(Tilefrm.File1.List(c - 1), 3, Len(Tilefrm.File1.List(c - 1)) - 6)
Next c


Exit Sub
errh:
MsgBox Err & Err.Description
Stop
Resume
End Sub

'The following four functions convert the selection of tiles

Public Function pY(Y As Single) As Single
' Converts the selection of tiles to blocks instead of pixels
pY = Int(Y / Tile.PixelHeight) '+ 1
End Function

Public Function pX(X As Single) As Single
' Converts the selection of tiles to blocks instead of pixels
pX = Int(X / Tile.PixelWidth) '+ 1
End Function

