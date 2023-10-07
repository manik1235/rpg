VERSION 2.00
Begin Form TextDisplay 
   BackColor       =   &H00404000&
   ClientHeight    =   3675
   ClientLeft      =   3840
   ClientTop       =   3915
   ClientWidth     =   4200
   ControlBox      =   0   'False
   Height          =   4080
   Left            =   3780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   4200
   Top             =   3570
   Width           =   4320
   Begin CommandButton Choice 
      BackColor       =   &H00404000&
      Caption         =   "Choice(0)"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   2640
      Visible         =   0   'False
      Width           =   3735
   End
   Begin CommandButton Choice 
      BackColor       =   &H00404000&
      Caption         =   "Choice(1)"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   12
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   0
      Top             =   3120
      Visible         =   0   'False
      Width           =   3735
   End
   Begin Label Description 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontName        =   "MS Sans Serif"
      FontSize        =   13.5
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   3735
   End
End
Option Explicit
'------------------------------------------------------------
'
'------------------------------------------------------------

Sub Choice_Click (Index As Integer)
'------------------------------------------------------------
'
'------------------------------------------------------------
    ' return choice to calling routine
    DisplayChoice = Index

    ' remove display form
    TextDisplay.Hide
End Sub

Sub Form_Paint ()
'------------------------------------------------------------
'
'------------------------------------------------------------
    ' center form on screen
    Left = (Screen.Width - Width) / 2
    Top = (Screen.Height - Height) / 2

    ' hide choice button(s) if needed
    Choice(0).Visible = (Len(Choice(0).Caption) <> 0)
    Choice(1).Visible = (Len(Choice(1).Caption) <> 0)
End Sub

