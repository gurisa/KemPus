VERSION 5.00
Begin VB.Form FormLogin 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3765
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormLogin.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameLogin 
      BackColor       =   &H00000000&
      Caption         =   "&Login"
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.Timer TimerExit 
         Left            =   120
         Top             =   1200
      End
      Begin VB.CommandButton CmdLogin 
         Caption         =   "&Login"
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox TextPassword 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1080
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox TextID 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1080
         TabIndex        =   1
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label LabelPassword 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   855
      End
      Begin VB.Label LabelID 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "FormLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Conn As New ADODB.Connection
Dim RS As New ADODB.Recordset

Private Sub CmdLogin_Click()
If TextID.Text = "" Or TextPassword.Text = "" Then
    MsgBox "ID Atau Password Masih Kosong", vbExclamation, "Masukan ID Dan Password"
Else
    If Conn.State = 1 Then Conn.Close
    Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\Database\db_keuangan.mdb"
        If RS.State = 1 Then RS.Close
            RS.Open "SELECT * FROM tb_petugas WHERE id_petugas='" & TextID.Text & "' AND password_petugas='" & TextPassword.Text & "'", Conn, 3, 3
            If Not RS.EOF Then
                FormUtama.Show
                FormUtama.StatusBarUtama.Panels(1).Text = TextID.Text
                FormUtama.StatusBarUtama.Panels(2).Text = RS!otoritas_petugas
                FormUtama.TextUsername.Text = TextID.Text
                FormUtama.TextPassword.Text = TextPassword.Text
                FormUtama.TextOtoritas.Text = RS!otoritas_petugas
                FormUtama.FrameSetPassword.Caption = "&Panel " & TextID.Text & ""
                Unload Me
                If RS!otoritas_petugas = "PETUGAS" Then
                    FormUtama.ImageOperatorStatus.Visible = False
                    FormUtama.ImagePetugasStatus.Visible = True
                ElseIf RS!otoritas_petugas = "OPERATOR" Then
                    FormUtama.ImageOperatorStatus.Visible = True
                    FormUtama.ImagePetugasStatus.Visible = False
                End If
            Else
                MsgBox "ID Atau Password Login Salah", vbCritical, "Periksa ID Atau Password"
            End If
End If
End Sub

Private Sub Form_Load()
If App.PrevInstance Then
    MsgBox "Program Sedang Di Jalankan", vbExclamation, "Periksa Status Program"
Else
    Exit Sub
End If
TimerExit.Enabled = True
TimerExit.Interval = 10000
End Sub

Private Sub TimerExit_Timer()
End
End Sub
