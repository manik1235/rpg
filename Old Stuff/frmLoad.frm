VERSION 5.00
Begin VB.Form frmLoad 
   Caption         =   "Form1"
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   9795
   StartUpPosition =   3  'Windows Default
   Begin VB.Image MainPic 
      Height          =   3840
      Left            =   0
      Picture         =   "frmLoad.frx":0000
      Top             =   0
      Width           =   1920
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This is the name of the loaded map
Public MapName As String

'These will hold the values that come out of the map file
Public TempXSize As Integer 'BigMapXSize
Public TempYSize As Integer 'BigMapYSize

'This stores msgbox responses
Dim MsgR As Integer

'For the open File
Dim FileHandle As Integer
Dim FileData As Variant
Dim FileDesc As Variant

'For the map sizes
Dim BigMapXSize
Dim BigMapYSize


'crap
Dim DESC_SYLPH_SICK
Dim DESC_SYLPH_WELL
Dim DESC_BOATDOCK
Dim DESC_FOUNTAIN
Dim DESC_BOAT_SINKS
Dim DESC_CHALICE
Dim DESC_DEAD
Dim DESC_SLEEPING
Dim DESC_SNAKEBITE
Dim DESC_BEGIN_PLAY
Dim DESC_HERMIT






Public Sub Form_Load()
    ' set up help file
    App.HelpFile = App.Path & "\RPG.HLP"
    
    ' load picture form, but don't show
    Load PictureForm
    Load Display

    ' sets the Form's Pic
    MainPic.Picture = PictureForm.Picture1(5)

    FileHandle = FreeFile
    
    'Delete this later
    MapName = "Map"
    
    On Error Resume Next
    Open App.Path & "\maps\" & MapName & ".rpg" For Input As #FileHandle

    ' This will open the map's .rpg file to get all the custom stuff
    If Err = 0 Then
        While Not EOF(FileHandle)
          Input #FileHandle, FileDesc, FileData
          If FileDesc = "BigMapXSize" Then BigMapXSize = FileData
          If FileDesc = "BigMapYSize" Then BigMapYSize = FileData
        Wend
    Else
        MsgR = MsgBox(Error$, vbCritical, "Map File Error")
    End If

    Close #FileHandle

    On Error GoTo 0
    
    
    ' size of entire map, in tiles
    Map.BigMapXSize = TempXSize
    Map.BigMapYSize = TempYSize

DESC_SYLPH_SICK = "Hello there, my old friend.  It's been far too long since you visited.  You look quite ill.  I can help with that ... but, then I must go.  If you want to cure the poison, you should seek out the Golden Chalice.  Rumor has it you can find it to east."
DESC_SYLPH_WELL = "Hello there, my old friend.  It's been far too long since you visited.  I wish I could stick around and talk, but, I must go.  If you want to cure the poison, you should seek out the Golden Chalice.  Rumor has it you can find it to east."
DESC_BOATDOCK = "You're on a dock.  There is a rickety old row boat tied up here.  You can see an island off to the west."
DESC_FOUNTAIN = "You see a magnificant fountain.  All of the plants around the fountain are thriving.  The water coming from it practically glows.  The water looks so clean, so good.  What do you do?"
DESC_BOAT_SINKS = "As you pull up to the dock, the boat springs a leak.  You manage to leap out before it sinks, but, the boat is gone."
DESC_CHALICE = "You found it!  Mmmm... you drink from it and feel MUCH better.  Time to wander home and return to the 'hard' life of your pipe and your hammock ... The End"
DESC_DEAD = "You're dead.  I guess you won't be needed that chalice after all ... thanks for playing."
DESC_SLEEPING = "It's late at night, and you're sleeping soundly.  Ahhh, a nice long night's rest is nice after a long day lounging in your hammock smoking your pipe and drinking homebrew.  But, what's that?  Something's in bed with you."
DESC_SNAKEBITE = "Ouch!  It bit you ... you feel a pain shooting up your leg.  It's a huge snake with big dripping fangs.  You grab the club you keep next to your bed that you use for rats and quickly dispatch the snake.  It'll make a nice belt now ..."
DESC_BEGIN_PLAY = "Unfortunately, you're feeling very, very weak.  If it wasn't for your strong constitution, you'd probably be dead now.  Maybe it's time to go looking for a cure and get some revenge.  Your friend Sylvia the Sylph can help.  She lives in the forest in the valley to the south."
DESC_HERMIT = "Leave me alone!  Can't you see that I'm taking care of Clyde?  He's my only friend in the world.  Some nasty ruffians just ran through here and uprooted him.  Now before you go asking a bunch of questions, I'll just tell you what I know ... there's some treasure to the northwest ... and the big evil guy ran straight north.  That's all!  Go away!"



Map.Show
Unload Me

End Sub
