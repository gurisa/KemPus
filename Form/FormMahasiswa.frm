VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormMahasiswa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mahasiswa"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormMahasiswa.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   2040
      TabIndex        =   15
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton CmdShow 
      Caption         =   "Show"
      Height          =   375
      Left            =   4680
      TabIndex        =   14
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3360
      TabIndex        =   12
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   9120
      TabIndex        =   11
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Frame FrameDataMahasiswa 
      Caption         =   "&Data Mahasiswa"
      Height          =   5055
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   5775
      Begin MSDataGridLib.DataGrid DataGridMahasiswa 
         Height          =   4695
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   8281
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame FrameMahasiswa 
      Caption         =   "&Mahasiswa"
      Height          =   2895
      Left            =   6000
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin VB.CommandButton CmdAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   4440
         TabIndex        =   13
         Top             =   2280
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTTahunAkademik 
         Height          =   255
         Left            =   2400
         TabIndex        =   8
         Top             =   1320
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   450
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   91029505
         CurrentDate     =   41925
      End
      Begin VB.TextBox TextNoRekening 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2400
         TabIndex        =   3
         Top             =   1800
         Width           =   3255
      End
      Begin VB.TextBox TextNamaMahasiswa 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2400
         TabIndex        =   2
         Top             =   840
         Width           =   3255
      End
      Begin VB.TextBox TextNIM 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2400
         TabIndex        =   1
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label LabelNoRekening 
         Caption         =   "NO REKENING"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label LabelTahunAkademik 
         Caption         =   "TAHUN AKADEMIK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label LabelNamaMahasiswa 
         Caption         =   "NAMA MAHASISWA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label LabelNIM 
         Caption         =   "NOMOR INDUK MAHASISWA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   2175
      End
   End
End
Attribute VB_Name = "FormMahasiswa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAdd_Click()
BukaDatabase
RS.Open "SELECT nim, nama_mahasiswa FROM tb_mahasiswa WHERE nim = '" & TextNIM.Text & "'", Conn, adOpenDynamic, adLockOptimistic
If RS.EOF Then
    If CmdAdd.Caption = "Add" Then
        CmdAdd.Caption = "Save"
        CmdDelete.Enabled = False
        CmdShow.Enabled = False
        CmdEdit.Enabled = False
        CmdRefresh.Enabled = False
        TextNIM.Enabled = True
        TextNamaMahasiswa.Enabled = True
        DTTahunAkademik.Enabled = True
        TextNoRekening.Enabled = True
    ElseIf CmdAdd.Caption = "Save" Then
        If TextNIM.Text = "" Or TextNamaMahasiswa.Text = "" Or TextNoRekening.Text = "" Then
            MsgBox "Masih Terdapat Data Yang Belum Di Isi", vbExclamation, "Masih Terdapat Data Yang Belum Di Isi"
        Else
            BukaDatabase
            RS.Open "INSERT INTO tb_mahasiswa(nim, nama_mahasiswa, tahun_akademik, nomor_rekening_mahasiswa) VALUES('" & TextNIM.Text & "','" & TextNamaMahasiswa.Text & "','" & DTTahunAkademik.Value & "','" & TextNoRekening.Text & "')", Conn, adOpenDynamic, adLockOptimistic
            Call CmdRefresh_Click
            CmdAdd.Caption = "Add"
            CmdDelete.Enabled = True
            CmdShow.Enabled = True
            CmdEdit.Enabled = True
            CmdRefresh.Enabled = True
            TextNIM.Enabled = False
            TextNamaMahasiswa.Enabled = False
            DTTahunAkademik.Enabled = False
            TextNoRekening.Enabled = False
            MsgBox "Berhasil Menambahkan Data Mahasiswa", vbInformation, "Berhasil Menambahkan Data Mahasiswa"
        End If
    End If
Else
    MsgBox "NIM Sudah Di Gunakan Oleh " & RS!nama_mahasiswa & "", vbExclamation, "NIM Sudah Di Gunakan"
End If
End Sub

Private Sub CmdAdd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TextNIM.SetFocus
End If
End Sub

Private Sub CmdDelete_Click()
If Not RS.EOF Then
    If MsgBox("Hapus Data " & DataGridMahasiswa.Columns(1).Text & " ?", vbExclamation + vbYesNo, "Hapus Data Mahasiswa") = vbYes Then
        BukaDatabase
        RS.Open "DELETE FROM tb_mahasiswa WHERE nim='" & DataGridMahasiswa.Columns(0).Text & "'", Conn, adOpenDynamic, adLockOptimistic
        Call CmdRefresh_Click
        MsgBox "Data Mahasiswa Berhasil Di Hapus", vbInformation, "Berhasil Menghapus Data Mahasiswa"
    Else
        Exit Sub
    End If
Else
    MsgBox "Data Mahasiswa Tidak Tersedia", vbExclamation, "Data Mahasiswa Tidak Tersedia"
End If
End Sub

Private Sub CmdEdit_Click()
If Not RS.EOF Then
    If CmdEdit.Caption = "Edit" Then
        BukaDatabase
        RS.Open "SELECT * FROM tb_mahasiswa WHERE nim = '" & DataGridMahasiswa.Columns(0).Text & "'", Conn, adOpenDynamic, adLockOptimistic
        TextNIM.Text = RS!NIM
        TextNamaMahasiswa.Text = RS!nama_mahasiswa
        DTTahunAkademik.Value = RS!tahun_akademik
        TextNoRekening.Text = RS!nomor_rekening_mahasiswa
        
        CmdEdit.Caption = "Save"
        CmdAdd.Enabled = False
        CmdDelete.Enabled = False
        CmdShow.Enabled = False
        CmdRefresh.Enabled = False
        TextNamaMahasiswa.Enabled = True
        DTTahunAkademik.Enabled = True
        TextNoRekening.Enabled = True
    ElseIf CmdEdit.Caption = "Save" Then
        BukaDatabase
        RS.Open "UPDATE tb_mahasiswa SET nama_mahasiswa = '" & TextNamaMahasiswa.Text & "', tahun_akademik = '" & DTTahunAkademik.Value & "', nomor_rekening_mahasiswa = '" & TextNoRekening.Text & "' WHERE nim = '" & TextNIM.Text & "'", Conn, adOpenDynamic, adLockOptimistic
        MsgBox "Data Mahasiswa Berhasil Di Ubah", vbInformation, "Data Mahasiswa Berhasil Di Ubah"
        CmdEdit.Caption = "Edit"
        CmdAdd.Enabled = True
        CmdDelete.Enabled = True
        CmdShow.Enabled = True
        CmdRefresh.Enabled = True
        TextNamaMahasiswa.Enabled = False
        DTTahunAkademik.Enabled = False
        TextNoRekening.Enabled = False
        Call CmdRefresh_Click
    End If
Else
    MsgBox "Data Mahasiswa Tidak Tersedia", vbExclamation, "Data Mahasiswa Tidak Tersedia"
End If
End Sub

Private Sub CmdRefresh_Click()
BukaDatabase
RS.Open "SELECT * FROM tb_mahasiswa", Conn, adOpenDynamic, adLockOptimistic
Set DataGridMahasiswa.DataSource = RS.DataSource

PengaturanDataGridMahasiswa

TextNIM.Text = ""
TextNamaMahasiswa.Text = ""
TextNoRekening.Text = ""
End Sub

Private Sub CmdShow_Click()
If CmdShow.Caption = "Show" Then
    CmdShow.Caption = "Hide"
With FormMahasiswa
    .Width = 12000
    .ScaleWidth = 11910
    .Height = 6165
    .ScaleHeight = 5730
End With
ElseIf CmdShow.Caption = "Hide" Then
    CmdShow.Caption = "Show"
With FormMahasiswa
    .Width = 6075
    .ScaleWidth = 6000
    .Height = 6165
    .ScaleHeight = 5730
End With
End If
End Sub

Private Sub CmdShow_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
CmdAdd.SetFocus
End If
End Sub

Private Sub Form_Load()
DTTahunAkademik.Value = Format(Now, "dd/mm/yyyy")
With FormMahasiswa
    .Width = 6075
    .ScaleWidth = 6000
    .Height = 6165
    .ScaleHeight = 5730
End With
BukaDatabase
RS.Open "SELECT * FROM tb_mahasiswa", Conn, adOpenDynamic, adLockOptimistic
Set DataGridMahasiswa.DataSource = RS.DataSource

PengaturanDataGridMahasiswa
End Sub

Public Sub PengaturanDataGridMahasiswa()
With DataGridMahasiswa
    .Columns(0).Caption = "NIM"
    .Columns(1).Caption = "NAMA MAHASISWA"
    .Columns(2).Caption = "TAHUN AKADEMIK"
    .Columns(3).Caption = "NO REKENING BANK"
End With
End Sub

Private Sub TextNIM_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TextNamaMahasiswa.SetFocus
End If
End Sub

Private Sub TextNamaMahasiswa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
DTTahunAkademik.SetFocus
End If
End Sub

Private Sub DTTahunAkademik_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TextNoRekening.SetFocus
End If
End Sub

Private Sub TextNoRekening_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
CmdAdd.SetFocus
End If

If Not (KeyAscii >= Asc("0") & Chr(13) _
     And KeyAscii <= Asc("9") & Chr(13) _
     Or KeyAscii = vbKeyBack _
     Or KeyAscii = vbKeyDelete _
     Or KeyAscii = vbKeySpace) Then
    TextNoRekening.Text = ""
    KeyAscii = 0
End If
End Sub
