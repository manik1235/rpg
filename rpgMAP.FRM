VERSION 5.00
Begin VB.Form Map 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Role Playing Game"
   ClientHeight    =   11055
   ClientLeft      =   225
   ClientTop       =   720
   ClientWidth     =   9420
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "rpgMAP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   11055
   ScaleWidth      =   9420
   Begin VB.PictureBox ComDialog 
      Height          =   480
      Left            =   600
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   11
      Top             =   5400
      Width           =   1200
   End
   Begin VB.CommandButton ExitButton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Exit RPG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6840
      TabIndex        =   1
      Top             =   4140
      Width           =   1875
   End
   Begin VB.PictureBox picMapCanvas 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4830
      Left            =   4920
      ScaleHeight     =   320
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   8
      Top             =   5760
      Width           =   4830
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4830
      Left            =   180
      ScaleHeight     =   320
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   0
      Top             =   180
      Width           =   4830
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   5040
      Picture         =   "rpgMAP.frx":030A
      ScaleHeight     =   3000
      ScaleWidth      =   1200
      TabIndex        =   6
      Top             =   1200
      Width           =   1200
      Begin VB.Image imgEast 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   780
         MousePointer    =   9  'Size W E
         Top             =   1200
         Width           =   435
      End
      Begin VB.Image imgNE 
         Appearance      =   0  'Flat
         Height          =   855
         Left            =   780
         MousePointer    =   6  'Size NE SW
         Top             =   300
         Width           =   435
      End
      Begin VB.Image imgWest 
         Appearance      =   0  'Flat
         Height          =   675
         Left            =   60
         MousePointer    =   9  'Size W E
         Top             =   1140
         Width           =   435
      End
      Begin VB.Image imgSouth 
         Appearance      =   0  'Flat
         Height          =   855
         Left            =   420
         MousePointer    =   7  'Size N S
         Top             =   1860
         Width           =   435
      End
      Begin VB.Image imgSE 
         Appearance      =   0  'Flat
         Height          =   855
         Left            =   780
         MousePointer    =   8  'Size NW SE
         Top             =   1800
         Width           =   435
      End
      Begin VB.Image imgSW 
         Appearance      =   0  'Flat
         Height          =   855
         Left            =   60
         MousePointer    =   6  'Size NE SW
         Top             =   1800
         Width           =   435
      End
      Begin VB.Image imgNW 
         Appearance      =   0  'Flat
         Height          =   855
         Left            =   60
         MousePointer    =   8  'Size NW SE
         Top             =   300
         Width           =   435
      End
      Begin VB.Image imgNorth 
         Appearance      =   0  'Flat
         Height          =   735
         Left            =   420
         MousePointer    =   7  'Size N S
         Top             =   360
         Width           =   435
      End
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Weapon:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   6600
      TabIndex        =   3
      Top             =   660
      Width           =   1155
   End
   Begin VB.Label lblWeapon 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Club"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   7800
      TabIndex        =   4
      Top             =   660
      Width           =   1215
   End
   Begin VB.Label lblHealth 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Excellent"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   7800
      TabIndex        =   5
      Top             =   300
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Health:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6600
      TabIndex        =   9
      Top             =   300
      Width           =   975
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "lblStatus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2955
      Left            =   6600
      TabIndex        =   10
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "To get help on how to play RPG, press F1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   5100
      Width           =   8895
   End
   Begin VB.Label lblFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   6360
      TabIndex        =   2
      Top             =   180
      Width           =   2715
   End
End
Attribute VB_Name = "Map"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------
' Declare all the functions, constants, and variables used
' in MAP.FRM.
'------------------------------------------------------------

' Declare Windows API functions.
'Private Declare Function BitBlt Lib "GDI" (ByVal hDestDC As Integer, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal xSrc As Integer, ByVal ySrc As Integer, ByVal dwRop As Long) As Integer
'Private Declare Function CreateCompatibleBitmap Lib "GDI" (ByVal hdc As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer) As Integer
'Private Declare Function CreateCompatibleDC Lib "GDI" (ByVal hdc As Integer) As Integer
'Private Declare Function SelectObject Lib "GDI" (ByVal hdc As Integer, ByVal hObject As Integer) As Integer
'Private Declare Function DeleteObject Lib "GDI" (ByVal hObject As Integer) As Integer
'Private Declare Function DeleteDC Lib "GDI" (ByVal hdc As Integer) As Integer

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "user32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

'Private Declare Function sndPlaySound Lib "MMSYSTEM" (ByVal lpSound As String, ByVal Flags As Integer) As Integer
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
' constants for snyPlaySound
Const SND_SYNC = (0)
Const SND_ASYNC = (1)
Const SND_NODEFAULT = (2)
Const SND_MEMORY = (4)
Const SND_LOOP = (8)
Const SND_NOSTOP = (16)

' constants for keystrokes
Const VK_PRIOR = &H21
Const VK_NEXT = &H22
Const VK_END = &H23
Const VK_HOME = &H24
Const VK_LEFT = &H25
Const VK_UP = &H26
Const VK_RIGHT = &H27
Const VK_DOWN = &H28

' size of displayed map, in tiles
Const MapXSize = 10
Const MapYSize = 10

' constant for BitBlt
Const SRCCOPY = &HCC0020
Const SRCAND = &H8800C6
Const SRCPAINT = &HEE0086
Const SRCOR = &HEE0086 'same as paint, but i like this name more

' to store the map name
Public MapName As String
Public FilePath As String 'Stores the path of the file

' weapons
Const WEAPON_CLUB = 0
Const WEAPON_DAGGER = 1
Const WEAPON_SWORD = 2

Public FlagSylph As Integer ' True if sylph has been visited
Public FlagHermit As Integer ' True if the hermit has been visited

Public StepNumber As Integer 'keeps track of wave file to play

Public StatusCountdown As Integer  ' keeps track of when status display
                                ' should be cleared

' holds map
Dim MapValues(1 To 150) As String
Public MapTopX As Integer
Public MapTopY As Integer
Public BigMapXSize As Integer
Public BigMapYSize As Integer

' holds TagStrings
Public TagString As String
Public TagStringB As String
Public TagStringC As String


' just a generic msgbox var
Public R As Integer

' characters in game
Public Numonster As Integer
'Dim player() As Character
'Dim monster() As Character

Const DESC_SYLPH_SICK = "Hello there, my old friend.  It's been far too long since you visited.  You look quite ill.  I can help with that ... but, then I must go.  If you want to cure the poison, you should seek out the Golden Chalice.  Rumor has it you can find it to east."
Const DESC_SYLPH_WELL = "Hello there, my old friend.  It's been far too long since you visited.  I wish I could stick around and talk, but, I must go.  If you want to cure the poison, you should seek out the Golden Chalice.  Rumor has it you can find it to east."
Const DESC_BOATDOCK = "You're on a dock.  There is a rickety old row boat tied up here.  You can see an island off to the west."
Const DESC_FOUNTAIN = "You see a magnificant fountain.  All of the plants around the fountain are thriving.  The water coming from it practically glows.  The water looks so clean, so good.  What do you do?"
Const DESC_BOAT_SINKS = "As you pull up to the dock, the boat springs a leak.  You manage to leap out before it sinks, but, the boat is gone."
Const DESC_CHALICE = "You found it!  Mmmm... you drink from it and feel MUCH better.  Time to wander home and return to the 'hard' life of your pipe and your hammock ... The End"
Const DESC_DEAD = "You're dead.  I guess you won't be needed that chalice after all ... thanks for playing."
Const DESC_SLEEPING = "It's late at night, and you're sleeping soundly.  Ahhh, a nice long night's rest is nice after a long day lounging in your hammock smoking your pipe and drinking homebrew.  But, what's that?  Something's in bed with you."
Const DESC_SNAKEBITE = "Ouch!  It bit you ... you feel a pain shooting up your leg.  It's a huge snake with big dripping fangs.  You grab the club you keep next to your bed that you use for rats and quickly dispatch the snake.  It'll make a nice belt now ..."
Const DESC_BEGIN_PLAY = "Unfortunately, you're feeling very, very weak.  If it wasn't for your strong constitution, you'd probably be dead now.  Maybe it's time to go looking for a cure and get some revenge.  Your friend Sylvia the Sylph can help.  She lives in the forest in the valley to the south."
Const DESC_HERMIT = "Leave me alone!  Can't you see that I'm taking care of Clyde?  He's my only friend in the world.  Some nasty ruffians just ran through here and uprooted him.  Now before you go asking a bunch of questions, I'll just tell you what I know ... there's some treasure to the northwest ... and the big evil guy ran straight north.  That's all!  Go away!"

'These are from file
Public TotalMonsters
Public TotalPlayers
Public TotalSigns
Public TotalHouses
Dim Player() As CharacterType  '(Player#, Intro message)
Dim Monster() As CharacterType  '(Monster#, Encounter message)
Dim Sign() As SignType  'Sign's pos. (Sign#, (Text, Sound, Choice #1, Choice #2, Num. Times to visit (0 for infinite)))
Dim House() As HouseType '(House#, (Text, Sound, Choice #1, Choice #2, Num. Times to visit (0 for infinite)))



Private Sub AdjustDisplay()
'------------------------------------------------------------
' Adjust the Display form's controls based on the value of
' its Tag property.  The Tag property indicates if a bitmap
' will be displayed on the form, or just text.
'------------------------------------------------------------
Dim margin As Integer
Dim BW As Integer

    margin = (9 * Screen.TwipsPerPixelX)
    BW = Display.Width - Display.ScaleWidth

    ' Make room for a bitmap
    If UCase$(Display.Tag) = "SHOW" Then
        Display.Pict.Visible = True

        Display.Width = BW + Display.Pict.Width + (margin * 2) + Display.Description.Width
        Display.Description.Left = margin + Display.Pict.Width
        Display.Choice(0).Left = margin + Display.Pict.Width
        Display.Choice(1).Left = margin + Display.Pict.Width

    ' Just show the text
    Else
        Display.Pict.Visible = False

        Display.Width = BW + (margin * 2) + Display.Description.Width
        Display.Description.Left = margin
        Display.Choice(0).Left = margin
        Display.Choice(1).Left = margin
    End If
    
    ' hide choice button(s) if needed
    Display.Choice(0).Visible = (Len(Display.Choice(0).Caption) <> 0)
    Display.Choice(1).Visible = (Len(Display.Choice(1).Caption) <> 0)

    ' center form on screen
    Display.Move (Screen.Width - Display.Width) / 2, (Screen.Height - Display.Height) / 2
End Sub

Private Sub Attack(Attacker As Integer, Defender As Integer, IsPlayer As Boolean)
  If IsPlayer Then
'------------------------------------------------------------
' Compute the results of an attack.
'------------------------------------------------------------

    ' if the weapon hits, apply damage
    If WeaponHit(Player(Attacker).Weapon) Then
        PlaySound "HIT", True

        ' compute damage and apply to defender
        SetHealth Defender, Player(Defender).Health - WeaponDamage(Player(Attacker).Weapon), False

    Else
        ' play swish (miss) sound
        PlaySound "MISS", True
    End If
  Else
'------------------------------------------------------------
' Compute the results of an attack.
'------------------------------------------------------------

    ' if the weapon hits, apply damage
    If WeaponHit(Monster(Attacker).Weapon) Then
        PlaySound "HIT", True

        ' compute damage and apply to defender
        SetHealth Defender, Monster(Defender).Health - WeaponDamage(Monster(Attacker).Weapon), True

    ' display result
        SetStatus "He hit you!  Ouch!"
        
    Else
        ' play swish (miss) sound
        PlaySound "MISS", True
    End If
  End If
End Sub

Private Sub BoatDock()
'------------------------------------------------------------
' Handle choice at the Boat Dock.
'------------------------------------------------------------
Static IslandCount As Integer

    ' if the boat has sunk ... can't go to island '
    If IslandCount >= 5 Then
        Exit Sub
    End If

    ' play the harbor sounds
    PlaySound "BOATDOCK", False

'    ShowChoice DESC_BOATDOCK, "Row To Island", "Stay Here", -1

    ' ensure proper repaint
    DoEvents

    ' if the user wants to go to the island
    If DisplayChoice = 0 Then
        ' row to island
        BoatDockRow 31, 26, -1, 2

        ' finally at the island
        Player(0).Image = 0
        Player(0).StartY = 12
        Player(0).StartX = 25
        CenterMap
        PaintMap

        ' handle fountain
        PlaySound "FOUNTAIN", False
    
        ShowChoice DESC_FOUNTAIN, "Drink Water and Leave", "Just Leave", -1

        ' handle user choice of drinking
        If DisplayChoice = 0 Then
            ' increment drinks from the fountain
            IslandCount = IslandCount + 1

            ' if the user needs healing, do it
            If Player(0).Health < 3 Then
                ' ensure map repaints
                DoEvents
                PlaySound "CURE", True
                SetHealth 0, Player(0).Health + 1, True

                ' ensure health level repaints
                DoEvents
            End If
        End If

        ' row to dock
        BoatDockRow 26, 31, 1, 3

        ' finally at the dock
        Player(0).Image = 0
        Player(0).StartY = 14
        Player(0).StartX = 32
        CenterMap
        PaintMap

        ' check to see if island is used up
        If IslandCount = 5 Then
            ShowChoice DESC_BOAT_SINKS, "", "Continue ...", -1
        End If
    End If
End Sub

Private Sub BoatDockRow(XStart As Integer, XEnd As Integer, XInc As Integer, Symbol As Integer)
'------------------------------------------------------------
' Handle journey across the lake.
'------------------------------------------------------------
Dim i As Integer
Dim TimeTemp As String

    ' play boat rowing sound
    PlaySound "BOAT", False

    ' row to island
    Player(0).Y = 13

    ' use west headed boat for char
    Player(0).Symbol = Symbol

    ' move across the lake
    For i = XStart To XEnd Step XInc
        Player(0).X = i
        CenterMap
        PaintMap

        ' wait a second before moving
        TimeTemp = Time$
        Do While TimeTemp = Time$
        Loop
    Next i

    ' return symbol to normal
    Player(0).Symbol = 0
End Sub

Private Sub CenterMap()
'------------------------------------------------------------
' Keep player centered in map display.
'------------------------------------------------------------

    MapTopX = Player(0).X - (MapXSize / 2)
    MapTopY = Player(0).Y - (MapYSize / 2)
End Sub

Private Sub Chalice()
'------------------------------------------------------------
' The player has found the Chalice!
'------------------------------------------------------------
    
    PlaySound "CHALICE", False

    ' display the chalice and winning message
    ShowChoice DESC_CHALICE, "", "Exit", 2

    ' exit the program
    End
End Sub

Private Sub CreateMonsters()
' This is now done when the map is loaded
    Numonster = TotalMonsters + 1
End Sub

Private Sub Dead()
'------------------------------------------------------------
' Uh-oh. Your dead!
'------------------------------------------------------------
    
    ' display the chalice and winning message
    ShowChoice DESC_DEAD, "", "Exit", 4

    ' exit the program
    End
End Sub

Private Sub DisplayHealth()
'------------------------------------------------------------
' Display health to user
'------------------------------------------------------------
    
    Select Case Player(0).Health
        Case 0:
            lblHealth.ForeColor = RGB(0, 0, 0)
            lblHealth.Caption = "Dead"
        Case 1:
            lblHealth.ForeColor = RGB(255, 0, 0)
            lblHealth.Caption = "Poor"
        Case 2:
            lblHealth.ForeColor = RGB(255, 255, 0)
            lblHealth.Caption = "Average"
        Case 3:
            lblHealth.ForeColor = RGB(0, 255, 0)
            lblHealth.Caption = "Excellent"
    End Select
End Sub

Private Sub DisplayWeapon()
'------------------------------------------------------------
' Display the player's current weapon.
'------------------------------------------------------------
    
    Select Case Player(0).Weapon
        Case 0:
            lblWeapon = "Club"
        Case 1:
            lblWeapon = "Dagger"
        Case 2:
            lblWeapon = "Sword"
    End Select
End Sub

Private Sub ExitButton_Click()
'------------------------------------------------------------
' Exit the program
'------------------------------------------------------------
    
    'End
    Unload Display
    Unload PictureForm
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'------------------------------------------------------------
' Handle the player's request to move to another Tilefrm.
'------------------------------------------------------------
picMapCanvas.Refresh
Dim OldX As Integer
Dim OldY As Integer
Dim OldMapX As Integer
Dim OldMapY As Integer
Dim NewTerrain As Integer
Dim NewTerrainStr As String

    ' save the old position of the character / screen
    OldX = Player(0).X
    OldY = Player(0).Y
    OldMapX = MapTopX
    OldMapY = MapTopY

    ' get the new position of the character
    Select Case KeyCode
        Case VK_DOWN:
            If Player(0).Y < BigMapYSize Then
                Player(0).Y = Player(0).Y + 1
            End If
        Case VK_UP:
            If Player(0).Y > 1 Then
                Player(0).Y = Player(0).Y - 1
            End If
        Case VK_LEFT:
            If Player(0).X > 1 Then
                Player(0).X = Player(0).X - 1
            End If
        Case VK_RIGHT:
            If Player(0).X < BigMapXSize Then
                Player(0).X = Player(0).X + 1
            End If
        Case VK_HOME:
            If Player(0).X > 1 Then
                Player(0).X = Player(0).X - 1
            End If

            If Player(0).Y > 1 Then
                Player(0).Y = Player(0).Y - 1
            End If
        Case VK_PRIOR:
            If Player(0).X < BigMapXSize Then
                Player(0).X = Player(0).X + 1
            End If

            If Player(0).Y > 1 Then
                Player(0).Y = Player(0).Y - 1
            End If
        Case VK_NEXT:
            If Player(0).X < BigMapXSize Then
                Player(0).X = Player(0).X + 1
            End If

            If Player(0).Y < BigMapYSize Then
                Player(0).Y = Player(0).Y + 1
            End If
        Case VK_END:
            If Player(0).X > 1 Then
                Player(0).X = Player(0).X - 1
            End If

            If Player(0).Y < BigMapYSize Then
                Player(0).Y = Player(0).Y + 1
            End If
    End Select

'-------------------------------------------------------------------------------------
'--------------------------Calls terrain----------------------------------------------
'-------------------------------------------------------------------------------------
    'Just get the base terrain info
    NewTerrainStr = Mid(MapValues(Player(0).Y), (Player(0).X * 4) + 1, 4)

' if this is water, don't move unless you're a boat or can fly
    If Right(NewTerrainStr, 1) = "~" Then
        Player(0).X = OldX
        Player(0).Y = OldY
    ' is this impassable? if so move only if you can fly
    ElseIf Right(NewTerrainStr, 1) = "-" Then
        Player(0).X = OldX
        Player(0).Y = OldY
    ' is there a live monster at this location?
    ElseIf IsMonster(Player(0).X, Player(0).Y) Then
        Dim MonsterIndex As Integer

        ' get the index to the monster
        MonsterIndex = MonsterAt(Player(0).X, Player(0).Y)

        ' we don't need to move
        Player(0).X = OldX
        Player(0).Y = OldY

        ' attempt to hit the monster
        Attack 0, MonsterIndex, True
        ' move monsters
        MoveMonsters
        PaintMap
        Exit Sub
    End If

    If (Player(0).X <> OldX) Or (Player(0).Y <> OldY) Then
        Dim TempResult As Integer

        ' move monsters
        MoveMonsters

        ' check to see if we should clear the status
        If StatusCountdown = 1 Then
            ' clear status label
            lblStatus = ""
            StatusCountdown = 0
        ' if we're still counting down, decrement
        ElseIf StatusCountdown > 0 Then
            ' count down
            StatusCountdown = StatusCountdown - 1
        End If

        ' if they found a dagger
        If Mid(NewTerrainStr, 2, 2) = "w0" Then
            StrOverwrite MapValues(Player(0).Y), Tilefrm.BaseTerrain(TerrainVal(Left(NewTerrainStr, 1), TagStringB)).Tag, Player(0).X
            SetStatus "You found a dagger!"
            SetWeapon 0, WEAPON_DAGGER, True
        ' check for a sword
        ElseIf Mid(NewTerrainStr, 2, 2) = "w1" Then
            StrOverwrite MapValues(Player(0).Y), Tilefrm.BaseTerrain(TerrainVal(Left(NewTerrainStr, 1), TagStringB)).Tag, Player(0).X
            SetStatus "You found a sword!"
            SetWeapon 0, WEAPON_SWORD, True
        ElseIf Mid(NewTerrainStr, 2, 2) = "t0" Then
            Chalice
        End If

        ' play the step sound
        
        If Left$(Mid(MapValues(Player(0).Y), (Player(0).X * 4) + 1, 4), 1) = "D" Then
            ' step sound for sand / desert
            PlaySound "STEPDES" & Format(StepNumber, "0"), False
        ElseIf Left$(Mid(MapValues(Player(0).Y), (Player(0).X * 4) + 1, 4), 1) = "F" Then
            ' step sound for forest
            PlaySound "STEPFOR" & Format(StepNumber, "0"), False
        ElseIf Left$(Mid(MapValues(Player(0).Y), (Player(0).X * 4) + 1, 4), 1) = "G" Then
            ' step sound for plains
            PlaySound "STEPGRE" & Format(StepNumber, "0"), False
        ElseIf Right$(Mid(MapValues(Player(0).Y), (Player(0).X * 4) + 1, 4), 1) = "&" Then
            ' step sound for bridge
            PlaySound "STEPBRI" & Format(StepNumber, "0"), False
        Else
            ' generic step sound
            PlaySound "STEP" & Format(StepNumber, "0"), False
        End If

        ' move to the next step sound
        StepNumber = 3 - StepNumber

        ' recompute x position of screen
        If ((Player(0).X - MapTopX) < 3) And (Player(0).X > 3) Then
            MapTopX = Player(0).X - 3
        ElseIf ((Player(0).X - MapTopX) > 6) And (Player(0).X < 198) Then
            MapTopX = Player(0).X - 6
        End If

        ' recompute y position of screen
        If ((Player(0).Y - MapTopY) < 3) And (Player(0).Y > 3) Then
            MapTopY = Player(0).Y - 3
        ElseIf ((Player(0).Y - MapTopY) > 6) And (Player(0).Y < 198) Then
            MapTopY = Player(0).Y - 6
        End If

        PaintMap
    End If

    ' check for special stuff
    If (Player(0).X = 41) And (Player(0).Y = 18) Then
        LakePath
    ElseIf (Player(0).X = 66) And (Player(0).Y = 52) Then
        SecretCave
    ' check for sign at entrance to player's valley
    ElseIf (Player(0).X = 6) And (Player(0).Y = 9) Then
        SetStatus "You see a sign that says 'Private Valley - Keep Out!'"
    ElseIf ((Player(0).X = 62) And (Player(0).Y = 70)) Or ((Player(0).X = 78) And (Player(0).Y = 63)) Then
        SetStatus "You see a sign that says 'Big Valley - Nothing Happens Here ... And We Like It That Way'"
    ElseIf (Player(0).X = 7) And (Player(0).Y = 32) Then
        SylviaTheSylph
    ' check for boat/island
    ElseIf (Player(0).X = 32) And (Player(0).Y = 14) Then
        BoatDock
    ' check for hermit
    ElseIf (Player(0).X = 126) And (Player(0).Y = 57) Then
        Hermit
    ElseIf (Player(0).X = 127) And (Player(0).Y = 6) Then
        SetStatus "There's the chalice!  Boy, that guy guarding it sure looks big.  I hope you're well armed!"
    ElseIf (Player(0).X = 127) And (Player(0).Y = 5) Then
        SetStatus "The closer you get, the bigger he is.  Good luck!  You'll need it ..."
    End If

End Sub

Private Sub Form_Load()
'------------------------------------------------------------
' Load the map information and run the opening sequence.
'------------------------------------------------------------
Dim FileHandle As Integer
Dim cutfile As Integer
Dim cutfile2 As Integer
Dim MapSize As Integer
Dim Data As Variant

    ' assigns TagString
    AssignTagString
    
    ' assigns tags to the tiles
    'AssignTags
    
    ' loads the map
    LoadTheMap

    ' set up help file
    App.HelpFile = App.Path & "\RPG.HLP"

    ' make sure that things are random
    Randomize

    ' set the height
    Map.Height = 6000
    
    ' load picture form, but don't show
    Load PictureForm
    Load Display

'-------------------------------------------------------------------------
'--------------------Change below to get from file------------------------
'-------------------------------------------------------------------------
    ' load map from disk
    LoadMap

    ' set current coordinates
    MapTopX = Player(0).StartX - 5
    MapTopY = Player(0).StartY - 5
    If MapTopX < 1 Then MapTopX = 1
    If MapTopY < 1 Then MapTopY = 1

    ' set initial character position
    'Player(0).StartX = 5
    'Player(0).StartY = 5
    SetHealth 0, Player(0).Health, True
    SetWeapon 0, Player(0).Weapon, True

    ' start up monsters / NPCs
    FlagSylph = False
    CreateMonsters

    StepNumber = 1

    ' show initial messages
    SetStatus "Your adventure begins ..."
'    PlaySound "SNORE", False

    ShowChoice Player(0).Text, "", "Continue ...", 5

    ' show snake screen
'    PlaySound "BITE", False

'    ShowChoice DESC_SNAKEBITE, "", "Continue ...", 1

'    ShowChoice DESC_BEGIN_PLAY, "", "Begin Play", 4

'----------------------------------------------------------------------
'-------------------Change above to get from file----------------------
'----------------------------------------------------------------------

    ' paint initial map
    PaintMap

    ' Center the main map form.
    Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2
End Sub

Private Sub Form_Paint()
'    Make3D Me, Picture1, BORDER_RAISED, 2
'    Make3D Me, lblFrame, BORDER_RAISED, 2
'    Make3D Me, ExitButton, BORDER_INSET, 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
'------------------------------------------------------------
' End the program.
'------------------------------------------------------------
    
    End
End Sub

Private Sub Hermit()
'------------------------------------------------------------
' The player encounters the Hermit.
'------------------------------------------------------------
    
    ' if we've been here before, don't do it again
    If FlagHermit Then
        Exit Sub
    End If

    ' play the Sylvia's Theme
    PlaySound "HERMIT", False
    
    ' make sure we don't come back here again
    FlagHermit = True
    
    ' give hermit's message to player
    ShowChoice DESC_HERMIT, "", "Continue ...", 3

End Sub

Private Sub imgEast_Click()
'------------------------------------------------------------
' Move east via mouse click.
'------------------------------------------------------------
    
    Form_KeyDown VK_RIGHT, 0
End Sub

Private Sub imgNE_Click()
'------------------------------------------------------------
' Move northeast via mouse click.
'------------------------------------------------------------
    
    Form_KeyDown VK_PRIOR, 0
End Sub

Private Sub imgNorth_Click()
'------------------------------------------------------------
' Move north via mouse click.
'------------------------------------------------------------
    
    Form_KeyDown VK_UP, 0
End Sub

Private Sub imgNW_Click()
'------------------------------------------------------------
' Move northwest via mouse click.
'------------------------------------------------------------
    
    Form_KeyDown VK_HOME, 0
End Sub

Private Sub imgSE_Click()
'------------------------------------------------------------
' Move southeast via mouse click.
'------------------------------------------------------------
    
    Form_KeyDown VK_NEXT, 0
End Sub

Private Sub imgSouth_Click()
'------------------------------------------------------------
' Move south via mouse click.
'------------------------------------------------------------
    
    Form_KeyDown VK_DOWN, 0
End Sub

Private Sub imgSW_Click()
'------------------------------------------------------------
' Move southwest via mouse click.
'------------------------------------------------------------
    
    Form_KeyDown VK_END, 0
End Sub

Private Sub imgWest_Click()
'------------------------------------------------------------
' Move west via mouse click.
'------------------------------------------------------------
    
    Form_KeyDown VK_LEFT, 0
End Sub

Private Function IsMonster(X As Integer, Y As Integer) As Integer
'------------------------------------------------------------
' Is there a monster at this location?
'------------------------------------------------------------
    
    IsMonster = (MonsterAt(X, Y) <> -1)
End Function

Private Sub LakePath()
'------------------------------------------------------------
' The player has encountered a secret lake path.
'------------------------------------------------------------
    
    ' check to see if we've done this before
    If Mid(MapValues(17), 40, 1) = "C" Then
        ' already done, get out
        Exit Sub
    End If

    PlaySound "LAKEPATH", False

    ' tell user they found a path
    SetStatus "You found a secret path!"

    ' show the path to the user - This actually modifies the map.
    StrOverwrite MapValues(17), "C", 40
    PaintMap

    StrOverwrite MapValues(17), "E", 39
    PaintMap

    StrOverwrite MapValues(16), "C", 38
    PaintMap

    StrOverwrite MapValues(16), "F", 37
    PaintMap

    StrOverwrite MapValues(17), "D", 36
    PaintMap

    StrOverwrite MapValues(17), "E", 35
    PaintMap

    StrOverwrite MapValues(16), "C", 34
    PaintMap

    StrOverwrite MapValues(16), "E", 33
    PaintMap
End Sub

'--------------------------------------------------------------------------
'-----------------------------LOAD MAP-------------------------------------
'---------------------------------------------------------------------------

Private Sub LoadMap()
'------------------------------------------------------------
' Read the map file and store in MapValues string array.
'------------------------------------------------------------
Dim FileHandle As Integer ' file handle for read
Dim YIndex                ' index for map read

    ' get handle to use
    FileHandle = FreeFile

    On Error Resume Next
    Open FilePath & MapName & ".rpg" For Input As #FileHandle

    If Err = 0 Then
        For YIndex = 1 To BigMapYSize
            Line Input #FileHandle, MapValues(YIndex)
        Next YIndex
    End If
    
    Close #FileHandle

    On Error GoTo 0
End Sub

'--------------------------------------------------------------------------
'-----------------------------LOAD MAP-------------------------------------
'---------------------------------------------------------------------------



Private Function MonsterAt(X As Integer, Y As Integer) As Integer
'------------------------------------------------------------
' What is the index of the monster at this location?
'------------------------------------------------------------
Dim i As Integer

    ' scan monsters
    For i = 0 To TotalMonsters
        ' found the monster, return the index
        If (Monster(i).Health > 0) And (Monster(i).X = X) And (Monster(i).Y = Y) Then
            MonsterAt = i
            Exit Function
        End If
    Next i

    ' couldn't find monster, return -1
    MonsterAt = -1
End Function

Private Sub MoveMonster(MonsterIndex As Integer)
'------------------------------------------------------------
' Move this monster, try to move towards character.
'------------------------------------------------------------
Dim NewX As Integer
Dim NewY As Integer
Dim NewTerrain As Integer

    ' get attempted new X position
    If Player(0).X < Monster(MonsterIndex).X Then
        NewX = Monster(MonsterIndex).X - 1
    ElseIf Player(0).X > Monster(MonsterIndex).X Then
        NewX = Monster(MonsterIndex).X + 1
    Else
        NewX = Monster(MonsterIndex).X
    End If

    ' get attempted new Y position
    If Player(0).Y < Monster(MonsterIndex).Y Then
        NewY = Monster(MonsterIndex).Y - 1
    ElseIf Player(0).Y > Monster(MonsterIndex).Y Then
        NewY = Monster(MonsterIndex).Y + 1
    Else
        NewY = Monster(MonsterIndex).Y
    End If
    
'-------------------------------------------------------------------------------------
'--------------------------Calls terrain----------------------------------------------
'-------------------------------------------------------------------------------------

    ' get value of current terrain
    NewTerrain = TerrainVal(Mid(MapValues(NewY), (NewX * 4) + 1, 4))
    
    ' if this is where the character is, attack!
    If (NewX = Player(0).X) And (NewY = Player(0).Y) Then
        Attack MonsterIndex, 0, False
        Exit Sub
    ' if there's no monster here, try to move
    ElseIf (MonsterAt(NewX, NewY) = -1) And Right$(Tilefrm.Terrain(NewTerrain), 1) <> "~" And Right$(Tilefrm.Terrain(NewTerrain), 1) <> "&" Then
        ' if this is within bounds for this monster
        If (NewX >= Monster(MonsterIndex).MinX) And (NewX <= Monster(MonsterIndex).MaxX) And (NewY >= Monster(MonsterIndex).MinY) And (NewY <= Monster(MonsterIndex).MaxY) Then
            ' move the monster
            Monster(MonsterIndex).X = NewX
            Monster(MonsterIndex).Y = NewY
            Monster(MonsterIndex).StartX = NewX
            Monster(MonsterIndex).StartY = NewY
        End If
    End If
End Sub

Private Sub MoveMonsters()
'------------------------------------------------------------
' Move all of the monsters
'------------------------------------------------------------
Dim i As Integer ' used for monster index

    ' scan monsters and move the ones than can
    For i = 0 To TotalMonsters
        If Monster(i).Health > 0 Then
            MoveMonster i
        End If
    Next i
End Sub



Private Sub PaintMapSquare(X As Integer, Y As Integer)
'------------------------------------------------------------
' Redraw just a single tile on the map.
'------------------------------------------------------------
Dim Temp

    If (X <> Player(0).X) Or (Y <> Player(0).Y) Then
        Temp = BitBlt(Picture1.hdc, (X - MapTopX) * 32, (Y - MapTopY) * 32, 32, 32, Terrain(TerrainVal(Mid(MapValues(Y), X, 1))).hdc, 0, 0, SRCCOPY)
    Else
        Temp = BitBlt(Picture1.hdc, (X - MapTopX) * 32, (Y - MapTopY) * 32, 32, 32, Characters(0).hdc, 0, 0, SRCCOPY)
    End If
End Sub





Private Sub Picture1_Paint()
'------------------------------------------------------------
' Repaint the map as necessary.
'------------------------------------------------------------
    
    ' ensure we've got a good map
    PaintMap
End Sub

Private Sub PlaySound(SoundName As String, Exclusive As Integer)
'------------------------------------------------------------
' Play a wave audio file.
'------------------------------------------------------------
Dim TempResult As Integer

    If Exclusive Then
        TempResult = sndPlaySound(App.Path & "\sounds\" & SoundName & ".WAV", SND_NODEFAULT)
    Else
        TempResult = sndPlaySound(App.Path & "\sounds\" & SoundName & ".WAV", SND_NODEFAULT + SND_ASYNC)
    End If
End Sub

Private Sub SecretCave()
'------------------------------------------------------------
' Reveal the Secret Cave.
'------------------------------------------------------------
    
    ' check to see if we've done this before
    If Mid(MapValues(51), 66, 1) = "A" Then
        ' already done, get out
        Exit Sub
    End If

    ' play sound
    PlaySound "SECCAVE", False
    
    ' tell user they found a cave
    SetStatus "You found a secret cave!"

    ' show the cave to the user - This actually modifies the map.
    StrOverwrite MapValues(51), "A", 66
    PaintMap

    StrOverwrite MapValues(50), "H", 65
    PaintMap

    StrOverwrite MapValues(49), "0", 65
    StrOverwrite MapValues(48), "0", 64
    StrOverwrite MapValues(48), "X", 65
    StrOverwrite MapValues(48), "0", 66
    StrOverwrite MapValues(47), "0", 65
    PaintMap
End Sub

Private Sub SetHealth(CharNum As Integer, NewHealth As Integer, IsPlayer As Boolean)
  If IsPlayer Then
'------------------------------------------------------------
' Modify a character's health value.
'------------------------------------------------------------

    ' set the new health
    Player(CharNum).Health = NewHealth

    ' this is creature is dead ...
    If Player(CharNum).Health <= 0 Then
        ' make it scream
        PlaySound "DEAD", True
        Player(CharNum).Health = 0

        ' if this is the player, end the game
        If CharNum = 0 Then
            Dead
        End If
    End If

    ' if this is the user
    If CharNum = 0 Then
        ' display the new health
        DisplayHealth
    End If
  Else
'------------------------------------------------------------
' Modify a monster's health value.
'------------------------------------------------------------

    ' set the new health
    Monster(CharNum).Health = NewHealth

    ' this is creature is dead ...
    If Monster(CharNum).Health <= 0 Then
        ' make it scream
        PlaySound "DEAD", True
        Monster(CharNum).Health = 0
    End If

  End If
End Sub



Private Sub SetStatus(NewStatus As String)
'------------------------------------------------------------
' Set status and countdown
'------------------------------------------------------------
    
    lblStatus.Caption = NewStatus
    StatusCountdown = 5
End Sub

Private Sub SetWeapon(CharNum As Integer, Weapon As Integer, IsPlayer As Boolean)
'------------------------------------------------------------
' Assign a new weapon to a character.
'------------------------------------------------------------
    
    ' set the new weapon type
    If IsPlayer Then
      Player(CharNum).Weapon = Weapon
      DisplayWeapon
    Else
      Monster(CharNum).Weapon = Weapon
    End If

End Sub

Private Sub ShowChoice(Description As String, ChoiceA As String, ChoiceB As String, PicNum As Integer)
'------------------------------------------------------------
' When the player encounters something (a monster or object,
' for example), the player may be asked to make a choice.
' This choice is presented in the Display form.  This routine
' sets property values and shows the Display form.
'------------------------------------------------------------

    If PicNum >= 0 Then
        Display.Pict.Picture = PictureForm.Picture1(PicNum)
        
        Display.Tag = "SHOW"
    Else
        Display.Tag = ""
    End If

    Display.Description.Caption = Description
    Display.Choice(0).Caption = ChoiceA
    Display.Choice(1).Caption = ChoiceB
    AdjustDisplay
    Display.Show 1

End Sub

Private Sub StrOverwrite(InputStr As String, OverStr As String, Offset As Integer)
'------------------------------------------------------------
' Used to modify a specified location (or locations) in the
' map array.
'------------------------------------------------------------
    
    Dim strLeft As String
    Dim strMid As String
    Dim Offsetx As Integer
    
    'Offset = Offset * 4
    Offsetx = Offset * 4
    
    'strLeft = Left(InputStr, Offset)
    strLeft = Left(InputStr, Offsetx)
    strMid = Mid(InputStr, Offsetx + Len(OverStr) + 1)
    
    InputStr = strLeft & OverStr & strMid
End Sub

Private Sub SylviaTheSylph()
'------------------------------------------------------------
' Handle visiting Sylvia the Sylph
'------------------------------------------------------------
    
    ' if we've been here before, get out
    If FlagSylph Then
        Exit Sub
    End If

    ' play the Sylvia's Theme
    PlaySound "SYLPH", False
    
    ' make sure we don't come back here again
    FlagSylph = True
    
    ' if the character is wounded
    If Player(0).Health < 3 Then
        ShowChoice DESC_SYLPH_SICK, "Accept Sylvia's Help", "Refuse Her Help", 0
    
        If DisplayChoice = 0 Then
            ' play the healing sound
            PlaySound "CURE", False
    
            ' increase the character's health
            SetHealth 0, Player(0).Health + 1
        End If
    Else
        ' the character is fully healed already
        ShowChoice DESC_SYLPH_WELL, "", "Continue Your Adventure", 0
    End If

End Sub


Private Function WeaponDamage(Weapon As Integer) As Integer
'------------------------------------------------------------
' Calculate damage based on type of weapon.
'------------------------------------------------------------
    
    Select Case Weapon
        Case WEAPON_CLUB:
            WeaponDamage = 1
        Case WEAPON_DAGGER:
            WeaponDamage = 1 + Int(Rnd + 0.5)
        Case WEAPON_SWORD:
            WeaponDamage = 1 + Int(2 * Rnd + 0.5)
    End Select
End Function

Private Function WeaponHit(Weapon As Integer) As Integer
'------------------------------------------------------------
' Did the weapon make a hit?
'------------------------------------------------------------
    
    Select Case Weapon
        Case WEAPON_CLUB:
            WeaponHit = (Rnd < 0.4)
        Case WEAPON_DAGGER:
            WeaponHit = (Rnd < 0.35)
        Case WEAPON_SWORD:
            WeaponHit = (Rnd < 0.65)
    End Select
End Function


Private Sub PaintMapX()
'------------------------------------------------------------
' Build a new map bitmap in memory (using Windows API calls),
' then BitBlt this new bitmap into the on-screen PictureBox.
'------------------------------------------------------------
Dim X
Dim XIndex, YIndex
Dim CharacterIndex As Integer
Dim hdcMem As Long
Dim hbitmap As Long
Dim hbitmapOld As Long

    hdcMem = CreateCompatibleDC(Picture1.hdc)
    hbitmap = CreateCompatibleBitmap(Picture1.hdc, MapXSize * 32, MapYSize * 32)
    hbitmapOld = SelectObject(hdcMem, hbitmap)

    For XIndex = 0 To MapXSize - 1
        For YIndex = 0 To MapYSize - 1
            X = BitBlt(hdcMem, XIndex * 32, YIndex * 32, 32, 32, Terrain(TerrainVal(Mid(MapValues(YIndex + MapTopY), (XIndex + MapTopX) * 4 - 1, 1))).hdc, 0, 0, SRCCOPY)
        Next YIndex
    Next XIndex

    For CharacterIndex = 0 To Numonster
        If (Chars(CharacterIndex).X >= MapTopX) And (Chars(CharacterIndex).X < MapTopX + MapXSize) And (Chars(CharacterIndex).Y >= MapTopY) And (Chars(CharacterIndex).Y < MapTopY + MapYSize) And (Chars(CharacterIndex).Health > 0) Then
            ' paint character
            X = BitBlt(hdcMem, (Chars(CharacterIndex).X - MapTopX) * 32, (Chars(CharacterIndex).Y - MapTopY) * 32, 32, 32, Characters(Chars(CharacterIndex).Symbol).hdc, 0, 0, SRCCOPY)
        End If
    Next CharacterIndex

    X = BitBlt(Picture1.hdc, 0, 0, MapXSize * 32, MapYSize * 32, hdcMem, 0, 0, SRCCOPY)

    hbitmapOld = SelectObject(hdcMem, hbitmapOld)
    X = DeleteObject(hbitmap)
    X = DeleteDC(hdcMem)
End Sub


Public Sub PaintMap()
'------------------------------------------------------------
' Build a new map bitmap in memory (using Windows API calls),
' then BitBlt this new bitmap into the on-screen PictureBox.
'------------------------------------------------------------
Dim X
Dim XIndex, YIndex
Dim CharacterIndex As Integer
Dim tmpString As String
Dim asdf As String

    For XIndex = 0 To MapXSize - 1
        For YIndex = 0 To MapYSize - 1
            'MsgBox MapValues(YIndex + MapTopY)
          
'            X = BitBlt(picMapCanvas.hdc, XIndex * 32, YIndex * 32, 32, 32, Tilefrm.Terrain(TerrainVal(Mid(MapValues(YIndex + MapTopY), (XIndex + MapTopX) * 4 + 1, 4))).hdc, 0, 0, SRCCOPY)
            asdf = Mid(MapValues(YIndex + MapTopY), (XIndex + MapTopX) * 4 + 1, 1)
            tmpString = DrawMask(XIndex * 32, YIndex * 32, asdf, Mid(Mid(MapValues(YIndex + MapTopY), (XIndex + MapTopX) * 4 + 1, 4), 2), TagString, TagStringB)
        Next YIndex
    Next XIndex

    For CharacterIndex = 0 To TotalPlayers
        If (Player(CharacterIndex).X >= MapTopX) And (Player(CharacterIndex).X < MapTopX + MapXSize) And (Player(CharacterIndex).Y >= MapTopY) And (Player(CharacterIndex).Y < MapTopY + MapYSize) And (Player(CharacterIndex).Health > 0) Then
            ' paint character
            X = BitBlt(picMapCanvas.hdc, (Player(CharacterIndex).X - MapTopX) * 32, (Player(CharacterIndex).Y - MapTopY) * 32, 32, 32, Tilefrm.Characters(Player(CharacterIndex).Symbol).hdc, 0, 0, SRCCOPY)
        End If
    Next CharacterIndex
    For CharacterIndex = 0 To TotalMonsters
        If (Monster(CharacterIndex).X >= MapTopX) And (Monster(CharacterIndex).X < MapTopX + MapXSize) And (Monster(CharacterIndex).Y >= MapTopY) And (Monster(CharacterIndex).Y < MapTopY + MapYSize) And (Monster(CharacterIndex).Health > 0) Then
            ' paint character
            X = BitBlt(picMapCanvas.hdc, (Monster(CharacterIndex).X - MapTopX) * 32, (Monster(CharacterIndex).Y - MapTopY) * 32, 32, 32, Tilefrm.Characters(Monster(CharacterIndex).Symbol).hdc, 0, 0, SRCCOPY)
        End If
    Next CharacterIndex

    X = BitBlt(Picture1.hdc, 0, 0, MapXSize * 32, MapYSize * 32, picMapCanvas.hdc, 0, 0, SRCCOPY)

End Sub


Public Function TerrainVal(Tag As String, Optional TagStringVal As String = "%NONE%")
    If TagStringVal = "%NONE%" Then TagStringVal = TagString
    TerrainVal = InStr(1, TagStringVal, Tag) - 1
    TerrainVal = TerrainVal / 4
End Function


Public Sub AssignTagString()
LoadTiles Me
End Sub

Public Sub AssignTags()
Dim ATC As Integer
Load Tilefrm
    'Assigns the tags to the Tiles
    For ATC = 0 To Tilefrm.Terrain.Count - 1
      Tilefrm.Terrain(ATC).Tag = Mid(TagString, ATC * 4 + 1, 4)
    Next ATC
End Sub

Public Sub LoadTheMap()
Dim FileHandle
Dim MapSize
Dim Data
Dim c
On Error GoTo errh
Dim cutfile As Integer ' This stores where the last \ is
Dim cutfile2 As Integer ' This is a temp storage
'ComDialog.Filter = "RPG Level Files (.rpg)|*.rpg|All Files|*.*"
'ComDialog.InitDir = App.Path & "\maps\"
'ComDialog.ShowOpen
'If ComDialog.filename = "" Then
'  Exit Sub
'Else
'  MapName = ComDialog.filename
'End If
MapName = App.Path & "\maps\" & "arthur.rpg"
If Mid(MapName, Len(MapName) - 3, 1) = "." Then
  MapName = Mid(MapName, 1, Len(MapName) - 4)
End If

'This finds and separates the name from path
cutfile = 0
cutfile2 = 1
While cutfile2 > 0
  cutfile2 = InStr(cutfile + 1, MapName, "\")
  If cutfile2 <> 0 Then cutfile = cutfile2 ' +CutFile
Wend
FilePath = Mid(MapName, 1, cutfile)
MapName = Mid(MapName, cutfile + 1)

FileHandle = FreeFile
Open FilePath & MapName & ".rpg" For Input As #FileHandle

MapSize = -1
  
    While Not EOF(FileHandle)
    Input #FileHandle, Data
    Data = LCase$(Data)
    If Data = LCase("filename=") Then
      Input #FileHandle, MapName
    ElseIf Data = LCase("size=") Then
      Input #FileHandle, MapSize
    ElseIf Data = LCase("TotalMonsters=") Then
      Input #FileHandle, TotalMonsters
      ReDim Monster(-1 To TotalMonsters)
    ElseIf Data = LCase("TotalPlayers=") Then
      Input #FileHandle, TotalPlayers
      ReDim Player(-1 To TotalPlayers)
      
    ElseIf Data = LCase("TotalSigns=") Then
      Input #FileHandle, TotalSigns
      ReDim Sign(-1 To TotalSigns)
    ElseIf Data = LCase("TotalHouses=") Then
      Input #FileHandle, TotalHouses
      ReDim House(-1 To TotalHouses)
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
        Input #FileHandle, Monster(c).Weapon
      Next c
    ElseIf Data = LCase("MonsterHealth=") Then
      For c = 0 To TotalMonsters
        Input #FileHandle, Monster(c).Health
      Next c
    ElseIf Data = LCase("MonsterStart=") Then
      For c = 0 To TotalMonsters
        Input #FileHandle, Monster(c).StartX, Monster(c).StartY
        Monster(c).X = Monster(c).StartX
        Monster(c).Y = Monster(c).StartY
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
    ElseIf Data = LCase("PlayerStart=") Then
      For c = 0 To TotalPlayers
        Input #FileHandle, Player(c).StartX, Player(c).StartY
        Player(c).X = Player(c).StartX
        Player(c).Y = Player(c).StartY
      Next c
    ElseIf Data = LCase("PlayerHealth=") Then
      For c = 0 To TotalPlayers
        Input #FileHandle, Player(c).Health
      Next c
      
  
  

  
    End If
  Wend
Close #FileHandle
'This sets the map size acording to loaded data
Select Case MapSize
  Case 0:
    BigMapXSize = 25
    BigMapYSize = 25
  Case 1:
    BigMapXSize = 50
    BigMapYSize = 50
  Case 2:
    BigMapXSize = 100
    BigMapYSize = 100
  Case 3:
    BigMapXSize = 150
    BigMapYSize = 150
End Select



' This loads your level in
'TileScreen True

Exit Sub
'----------------------------------------
errh:
If Err = 53 Then
  R = MsgBox("Error! Map Information not found! Map might not load correctly. I'll try to load it anyway. Hint: Make sure the .rpg file is in the same directory as the map.", vbCritical, "File Not Found")
  Open FilePath & MapName & ".rpg" For Output As #FileHandle
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

' If the length of TopTile is 3 then it needs a space
If Len(TopTile) = 3 Then
  TopTile = " " & TopTile
End If

' Redraw base
rc = BitBlt(picMapCanvas.hdc, Int(X / 32) * 32, Int(Y / 32) * 32, 32, 32, Tilefrm.BaseTerrain(TerrainVal(Left(BaseTile, 1), TagStringBase)).hdc, 0, 0, SRCCOPY)
If TopTile <> "" And Mid(TopTile, 2, 2) <> "00" Then ' If TopTile = "" or "00" then that means there is no mask to draw... just the base tile.
  
  ' Make sure TopTile has a space, not a Terrain Modifier
  TopTile = " " & Mid(TopTile, 2)
  ' Draw mask
  rc = BitBlt(picMapCanvas.hdc, Int(X / 32) * 32, Int(Y / 32) * 32, 32, 32, Tilefrm.Black(TerrainVal(TopTile, TagStringTop)).hdc, 0, 0, SRCAND)
  ' Draw the top tile
  rc = BitBlt(picMapCanvas.hdc, Int(X / 32) * 32, Int(Y / 32) * 32, 32, 32, Tilefrm.Terrain(TerrainVal(TopTile, TagStringTop)).hdc, 0, 0, SRCOR)
  
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

