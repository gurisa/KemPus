VERSION 5.00
Begin VB.Form FormAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5205
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormAbout.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   5205
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox TextAbout 
      Height          =   2175
      Left            =   2040
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "FormAbout.frx":0CCA
      Top             =   120
      Width           =   3015
   End
   Begin VB.Image ImageLogo 
      Height          =   1935
      Left            =   120
      Picture         =   "FormAbout.frx":0CE7
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "FormAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdExit_Click()
Unload Me
End Sub

