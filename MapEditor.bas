Attribute VB_Name = "MapEditor"
'These declare types

Public Type XYType
  SpecialTile As String
  TileText As String
  Tag As String
  Picture As PictureBox
End Type
Public Type TileType
  XY() As XYType
  PixelWidth As Long
  PixelHeight As Long
  TwipWidth As Long
  TwipHeight As Long
End Type
Public Type PlayerType
  Text As String
  Weapon As Integer
  Health As Integer
  MaxHealth As Integer
  StartX As Integer 'Player start Pos. (Player#, (x)
  StartY As Integer 'Player start Pos. (Player#, (y)
  Image As Integer
  Index As Integer
  Symbol As Integer
  X As Integer
  Y As Integer
End Type
Public Type MonsterType
  Text As String
  Weapon As Integer
  Health As Integer
  MaxHealth As Integer
  StartX As Integer
  StartY As Integer
  MinX As Integer
  MinY As Integer
  MaxX As Integer
  MaxY As Integer
  Image As Integer
  Index As Integer
  X As Integer
  Y As Integer
End Type
Public Type SignType
  Caption As String
  Text As String
  Sound As String
  Choice1 As String
  Choice2 As String
  Visits As Integer
  Image As String
End Type
Public Type HouseType
  Caption As String
  Text As String
  Sound As String
  Choice1 As String
  Choice2 As String
  Visits As Integer
  Image As String
End Type
 'Sign's pos. (Sign#, (Text, Sound, Choice #1, Choice #2, Num. Times to visit (0 for infinite)))
 '(House#, (Text, Sound, Choice #1, Choice #2, Num. Times to visit (0 for infinite)))











