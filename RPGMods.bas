Attribute VB_Name = "RPGMods"
Option Explicit

Public Sub LoadTiles(frmForm As Object)
On Error GoTo errh

Dim FPlt As String
Dim c As Integer

' Loads the form where the tiles are stored
Load Tilefrm

' load the tiles
FPlt = App.Path & "\Bitmaps\"




' do my thing for index #0 (already loaded) Terrain Tiles

frmForm.TagString = ""
Tilefrm.File1 = FPlt & "Tiles\"

If Len(Tilefrm.File1.List(0)) = 7 Then
  ' if the file name needs to start with a space, add it to the tag
  frmForm.TagString = frmForm.TagString & " " & Left(Tilefrm.File1.List(c), 3)
Else
  ' if it doesnt, use the whole thing
  frmForm.TagString = frmForm.TagString & Left(Tilefrm.File1.List(c), 4)
End If
Set Tilefrm.Terrain(0).Picture = LoadPicture(FPlt & "Tiles\" & Tilefrm.File1.List(0))
Set Tilefrm.Black(0).Picture = LoadPicture(FPlt & "Tiles\" & Tilefrm.File1.List(1))
Tilefrm.Black(0).Tag = Right(frmForm.TagString, 4)
Tilefrm.Terrain(0).Tag = Right(frmForm.TagString, 4)

' do my thing for the rest
For c = 1 To (Tilefrm.File1.ListCount / 2) - 1 ' skip each black one
  Load Tilefrm.Terrain(c)
  Load Tilefrm.Black(c)
  Set Tilefrm.Terrain(c).Picture = LoadPicture(FPlt & "Tiles\" & Tilefrm.File1.List(c * 2))
  Set Tilefrm.Black(c).Picture = LoadPicture(FPlt & "Tiles\" & Tilefrm.File1.List(c * 2 + 1))
  If Len(Tilefrm.File1.List(c * 2)) = 7 Then
    ' if the file name needs to start with a space, add it to the tag
    frmForm.TagString = frmForm.TagString & " " & Left(Tilefrm.File1.List(c * 2), 3)
  Else
    ' if it doesnt, use the whole thing
    frmForm.TagString = frmForm.TagString & Left(Tilefrm.File1.List(c * 2), 4)
  End If
  Tilefrm.Black(c).Tag = Right(frmForm.TagString, 4)
  Tilefrm.Terrain(c).Tag = Right(frmForm.TagString, 4)
Next c




' do my thing for index #0 (already loaded) Character Tiles

frmForm.TagStringC = ""
Tilefrm.File1 = FPlt & "Characters\"

Set Tilefrm.Characters(0).Picture = LoadPicture(FPlt & "Characters\" & Tilefrm.File1.List(0))
Set Tilefrm.BCharacters(0).Picture = LoadPicture(FPlt & "Characters\" & Tilefrm.File1.List(1))
frmForm.TagStringC = frmForm.TagStringC & Left(Tilefrm.File1.List(0), 4)
Tilefrm.BCharacters(0).Tag = Right(frmForm.TagStringC, 4)
Tilefrm.Characters(0).Tag = Right(frmForm.TagStringC, 4)

' do my thing for the rest
For c = 1 To (Tilefrm.File1.ListCount / 2) - 1 ' skip each black one
  Load Tilefrm.Characters(c)
  Load Tilefrm.BCharacters(c)
  Set Tilefrm.Characters(c).Picture = LoadPicture(FPlt & "Characters\" & Tilefrm.File1.List(c * 2))
  Set Tilefrm.BCharacters(c).Picture = LoadPicture(FPlt & "Characters\" & Tilefrm.File1.List(c * 2 + 1))
  frmForm.TagStringC = frmForm.TagStringC & Left(Tilefrm.File1.List(c * 2), 4)
  Tilefrm.BCharacters(c).Tag = Right(frmForm.TagStringC, 4)
  Tilefrm.Characters(c).Tag = Right(frmForm.TagStringC, 4)
Next c




' do my thing for index #0 (already loaded) Terrain Tiles

frmForm.TagStringB = ""
Tilefrm.File1 = FPlt & "Base Tiles\"

Set Tilefrm.BaseTerrain(0).Picture = LoadPicture(FPlt & "Base Tiles\" & Tilefrm.File1.List(0))
frmForm.TagStringB = frmForm.TagStringB & Left(Tilefrm.File1.List(0), 4)
Tilefrm.BaseTerrain(0).Tag = Right(frmForm.TagStringB, 4)

' do my thing for the rest
For c = 1 To (Tilefrm.File1.ListCount - 1)
  Load Tilefrm.BaseTerrain(c)
  Set Tilefrm.BaseTerrain(c).Picture = LoadPicture(FPlt & "Base Tiles\" & Tilefrm.File1.List(c))
  frmForm.TagStringB = frmForm.TagStringB & Left(Tilefrm.File1.List(c), 4)
  Tilefrm.BaseTerrain(c).Tag = Right(frmForm.TagStringB, 4)
Next c



On Error GoTo 0
Exit Sub
errh:
MsgBox Err & Err.Description
Stop
Resume

End Sub

