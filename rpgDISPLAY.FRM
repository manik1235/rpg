VERSION 5.00
Begin VB.Form Display 
   Appearance      =   0  'Flat
   BackColor       =   &H00404000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3675
   ClientLeft      =   2565
   ClientTop       =   2265
   ClientWidth     =   5955
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3675
   ScaleWidth      =   5955
   Begin VB.CommandButton Choice 
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      Caption         =   "Choice(1)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   3
      Top             =   3120
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton Choice 
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      Caption         =   "Choice(0)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2040
      TabIndex        =   2
      Top             =   2640
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.PictureBox Pict 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3875
      Left            =   -60
      ScaleHeight     =   3840
      ScaleWidth      =   1920
      TabIndex        =   0
      Top             =   -60
      Width           =   1945
   End
   Begin VB.Label Description 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   2040
      TabIndex        =   1
      Top             =   60
      Width           =   3735
   End
End
Attribute VB_Name = "Display"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------
' DISPLAY.FRM - used when the player is presented with a choice.
'------------------------------------------------------------

Private Sub Choice_Click(Index As Integer)
'------------------------------------------------------------
' Return the user's choice via the global DisplayChoice.
'------------------------------------------------------------
    
    ' return choice to calling routine
    DisplayChoice = Index

    ' remove display form
    Display.Hide
End Sub

