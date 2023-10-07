VERSION 5.00
Begin VB.Form NewMapfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Map"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4155
   Icon            =   "meNewMapfrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   4155
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   3360
      TabIndex        =   15
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Load Existing Map"
      Height          =   615
      Left            =   3360
      TabIndex        =   14
      Top             =   720
      Width           =   735
   End
   Begin VB.Frame Frame3 
      Caption         =   "Base Texture"
      Height          =   2055
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   1695
      Begin VB.OptionButton Option1 
         Caption         =   "Blank"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Snow"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Tag             =   "S00_"
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Trees"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Tag             =   "F00_"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Water"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Tag             =   "W00~"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Dirt"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   10
         Tag             =   "R00_"
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Grass"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Tag             =   "G00_"
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Desert"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Tag             =   "D00_"
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Map Name"
      Height          =   615
      Left            =   1800
      TabIndex        =   5
      Top             =   1440
      Width           =   2295
      Begin VB.TextBox MapName 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Map Size"
      Height          =   1335
      Left            =   1800
      TabIndex        =   0
      Top             =   0
      Width           =   1455
      Begin VB.OptionButton Option 
         Caption         =   "135x135"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton Option 
         Caption         =   "100x100"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton Option 
         Caption         =   "50x50"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton Option 
         Caption         =   "25x25"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
   End
End
Attribute VB_Name = "NewMapfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'These will store the vars that will be sent to MapEdit
Public MapX As Integer
Public MapY As Integer
Dim TerrainType As Integer


Private Sub Command1_Click()
MapEdit.Command1N_Click
End Sub

Private Sub Command2_Click()

Dim c As Integer

'click the correct options
For c = 0 To Option1.Count - 1
  If Option1(c).Value Then
    Option1_Click c
    Exit For
  End If
Next c
For c = 0 To Me.Option.Count - 1
  If Me.Option(c).Value = True Then
    Option_Click c
    Exit For
  End If
Next c

' Run rest of sub
MapEdit.Command2N_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 92 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 58 Or KeyAscii = 47 Or KeyAscii = 63 Or KeyAscii = 42 Or KeyAscii = 36 Or KeyAscii = 94 Then
  KeyAscii = 0
End If
End Sub

Private Sub Form_Load()
MapX = 25
MapY = 25
Option_Click 0
Option1_Click 1
MapName.Text = MapEdit.FileName
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
End Sub

Private Sub Option_Click(Index As Integer)
MapEdit.Option_ClickN Index
End Sub

Private Sub Option1_Click(Index As Integer)
MapEdit.Option1_ClickN MapEdit.TerrainVal(Option1(Index).Tag, MapEdit.TagStringB)
End Sub

