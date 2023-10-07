VERSION 5.00
Begin VB.Form Tilefrm 
   Caption         =   "Form1"
   ClientHeight    =   14460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   ScaleHeight     =   14460
   ScaleWidth      =   11760
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   ">>"
      Height          =   1215
      Left            =   3120
      TabIndex        =   8
      Top             =   6000
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<<"
      Height          =   1215
      Left            =   960
      TabIndex        =   7
      Top             =   6000
      Width           =   2175
   End
   Begin VB.PictureBox BCharacters 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   3000
      Width           =   480
   End
   Begin VB.FileListBox File1 
      Height          =   1650
      Left            =   7200
      TabIndex        =   5
      Top             =   5640
      Width           =   2295
   End
   Begin VB.PictureBox Black 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   3960
      Width           =   480
   End
   Begin VB.PictureBox BaseTerrain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   0
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   2040
      Width           =   480
   End
   Begin VB.Frame Frame1 
      Caption         =   "These just make it so the masks' indexes match w/ the terrain (these arent used)"
      Height          =   855
      Left            =   4080
      TabIndex        =   2
      Top             =   120
      Width           =   6015
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
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   3480
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
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   2520
      Width           =   480
   End
End
Attribute VB_Name = "Tilefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BaseTerrain_Click(Index As Integer)
MsgBox Index & BaseTerrain(Index).Tag
End Sub

Private Sub Characters_Click(Index As Integer)
MsgBox Index & Characters(Index).Tag
End Sub

Private Sub Command1_Click()
On Error Resume Next
For c = 0 To Terrain.Count - 1
  Terrain(c).Left = Terrain(c).Left - Terrain(c).Width * 5
  BaseTerrain(c).Left = BaseTerrain(c).Left - Terrain(c).Width * 5
  Characters(c).Left = Characters(c).Left - Terrain(c).Width * 5
Next c
End Sub

Private Sub Command2_Click()
On Error Resume Next
For c = 0 To Terrain.Count - 1
  Terrain(c).Left = Terrain(c).Left + Terrain(c).Width * 5
  BaseTerrain(c).Left = BaseTerrain(c).Left + Terrain(c).Width * 5
  Characters(c).Left = Characters(c).Left + Terrain(c).Width * 5
Next c
End Sub

Private Sub File1_Click()
MsgBox File1.ListIndex
End Sub

Private Sub Terrain_Click(Index As Integer)
MsgBox Index & Terrain(Index).Tag
End Sub
