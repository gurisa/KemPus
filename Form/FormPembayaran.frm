VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormPembayaran 
   Caption         =   "Pembayaran"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9375
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormPembayaran.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   9375
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameTransaksi 
      Caption         =   "&Transaksi"
      Height          =   1695
      Left            =   120
      TabIndex        =   20
      Top             =   0
      Width           =   9135
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Search"
         Height          =   375
         Left            =   8040
         TabIndex        =   24
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton CmdDelete 
         Caption         =   "Delete"
         Height          =   375
         Left            =   8040
         TabIndex        =   23
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton CmdEdit 
         Caption         =   "Edit"
         Height          =   375
         Left            =   8040
         TabIndex        =   22
         Top             =   240
         Width           =   975
      End
      Begin MSDataGridLib.DataGrid DataGridTransaksi 
         Height          =   1335
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   2355
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
   Begin VB.Frame FrameParameter 
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   9135
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   1200
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox TextTotalBayar 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4920
         TabIndex        =   9
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label LabelTotalBayar 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Pembayaran"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4200
         TabIndex        =   17
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame FramePembayaran 
      Caption         =   "&Pembayaran"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   9135
      Begin VB.ComboBox ComboStatus 
         Height          =   330
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1800
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker DTTanggalPembayaran 
         Height          =   375
         Left            =   6360
         TabIndex        =   7
         Top             =   1200
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   393216
         Format          =   91881473
         CurrentDate     =   41925
      End
      Begin VB.ComboBox ComboJenisPembayaran 
         Height          =   330
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1320
         Width           =   2535
      End
      Begin VB.ComboBox ComboSemester 
         Height          =   330
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   840
         Width           =   2535
      End
      Begin VB.ComboBox ComboNIM 
         Height          =   330
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox TextPembayaran 
         Height          =   315
         Left            =   6360
         MaxLength       =   11
         TabIndex        =   3
         Text            =   "0"
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox TextDenda 
         Height          =   315
         Left            =   6360
         MaxLength       =   11
         TabIndex        =   2
         Text            =   "0"
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label LabelNIM 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "NIM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label LabelJenisPembayaran 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Jenis Pembayaran"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label LabelTanggalPembayaran 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Pembayaran"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4560
         TabIndex        =   14
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label LabelDenda 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Denda"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4560
         TabIndex        =   13
         Top             =   720
         Width           =   975
      End
      Begin VB.Label LabelStatus 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label LabelBayar 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Pembayaran"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4560
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label LabelSemester 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Semester"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1575
      End
   End
End
Attribute VB_Name = "FormPembayaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdDelete_Click()
If Not RS.EOF Then
    If MsgBox("Hapus Data Transaksi Nomor " & DataGridTransaksi.Columns(0).Text & " ?", vbExclamation + vbYesNo, "Hapus Data Transaksi") = vbYes Then
        BukaDatabase
        RS.Open "DELETE FROM tb_pembayaran WHERE no_transaksi=" & DataGridTransaksi.Columns(0).Text & "", Conn, adOpenDynamic, adLockOptimistic
        Call CmdRefresh_Click
        MsgBox "Data Transaksi Berhasil Di Hapus", vbInformation, "Data Transaksi Berhasil Di Hapus"
    Else
        Exit Sub
    End If
Else
    MsgBox "Data Transaksi Tidak Tersedia", vbExclamation, "Data Transaksi Tidak Tersedia"
End If
End Sub

Private Sub CmdEdit_Click()
If Not RS.EOF Then
    If CmdEdit.Caption = "Edit" Then
        BukaDatabase
        RS.Open "SELECT * FROM tb_pembayaran WHERE no_transaksi = " & DataGridTransaksi.Columns(0).Text & "", Conn, adOpenDynamic, adLockOptimistic
        ComboNIM.Text = RS!NIM
        ComboSemester.Text = RS!semester
        ComboJenisPembayaran.Text = RS!jenis_pembayaran
        DTTanggalPembayaran.Value = RS!tanggal_pembayaran
        TextDenda.Text = RS!denda
        ComboStatus.Text = RS!Status
        TextPembayaran.Text = RS!Pembayaran
        CmdEdit.Caption = "Save"
        CmdDelete.Enabled = False
        CmdSave.Enabled = False
        CmdRefresh.Enabled = False
        CmdSearch.Enabled = False
    ElseIf CmdEdit.Caption = "Save" Then
        BukaDatabase
        RS.Open "UPDATE tb_pembayaran SET nim = '" & ComboNIM.Text & "', semester = '" & ComboSemester.Text & "', jenis_pembayaran = '" & ComboJenisPembayaran.Text & "', tanggal_pembayaran = '" & DTTanggalPembayaran.Value & "', denda = '" & TextDenda.Text & "', status = '" & ComboStatus.Text & "', pembayaran = '" & TextPembayaran.Text & "', total_pembayaran = '" & TextTotalBayar.Text & "' WHERE no_transaksi = " & DataGridTransaksi.Columns(0).Text & "", Conn, adOpenDynamic, adLockOptimistic
        MsgBox "Data Mahasiswa Berhasil Di Ubah", vbInformation, "Data Mahasiswa Berhasil Di Ubah"
        CmdEdit.Caption = "Edit"
        CmdDelete.Enabled = True
        CmdSave.Enabled = True
        CmdRefresh.Enabled = True
        CmdSearch.Enabled = True
        Call CmdRefresh_Click
    End If
Else
    MsgBox "Data Transaksi Tidak Tersedia", vbExclamation, "Data Transaksi Tidak Tersedia"
End If
End Sub

Private Sub CmdRefresh_Click()
TextDenda.Text = "0"
TextPembayaran.Text = "0"

BukaDatabase
RS.Open "SELECT tb_pembayaran.no_transaksi, tb_pembayaran.nim, tb_mahasiswa.nama_mahasiswa, tb_pembayaran.jenis_pembayaran, status FROM tb_pembayaran, tb_mahasiswa WHERE tb_mahasiswa.nim = tb_pembayaran.nim", Conn, adOpenDynamic, adLockOptimistic
Set DataGridTransaksi.DataSource = RS.DataSource

PengaturanDataGridTransaksi
End Sub

Private Sub CmdSave_Click()
If ComboNIM.Text = "" Or ComboSemester.Text = "" Or ComboJenisPembayaran.Text = "" Or TextDenda.Text = "" Or ComboStatus.Text = "" Or TextPembayaran.Text = "" Then
    MsgBox "Masukan Data Transaksi Dengan Benar", vbExclamation, "Masukan Data Transaksi Dengan Benar"
Else
    BukaDatabase
    RS.Open "INSERT INTO tb_pembayaran(nim, semester, jenis_pembayaran, tanggal_pembayaran, denda, status, pembayaran, total_pembayaran) VALUES('" & ComboNIM.Text & "','" & ComboSemester.Text & "','" & ComboJenisPembayaran.Text & "','" & DTTanggalPembayaran.Value & "','" & TextDenda.Text & "','" & ComboStatus.Text & "','" & TextPembayaran.Text & "','" & TextTotalBayar.Text & "')", Conn, adOpenDynamic, adLockOptimistic
    MsgBox "Pembayaran Berhasil Di Lakukan", vbInformation, "Pembayaran Berhasil Di Lakukan"
    Call CmdRefresh_Click
    If MsgBox("Print Bukti Pembayaran ?", vbInformation + vbYesNo, "Print Bukti Pembayaran") = vbYes Then
        BukaDatabase
        RS.Open "SELECT nim, nama_mahasiswa FROM tb_mahasiswa", Conn, adOpenDynamic, adLockOptimistic
        If Not RS.EOF Then
        Set FormBuktiPembayaran.DataGridPembayaranHead.DataSource = RS.DataSource
        
            PengaturanDataGridPembayaranHead
        
            BukaDatabase
            RS.Open "SELECT * FROM tb_pembayaran WHERE nim='" & FormBuktiPembayaran.DataGridPembayaranHead.Columns(0).Text & "'", Conn, adOpenDynamic, adLockOptimistic
            Set FormBuktiPembayaran.DataGridPembayaranDetail.DataSource = RS.DataSource
            
            PengaturanDataGridPembayaranDetail
        Else
            Exit Sub
        End If
        FormBuktiPembayaran.Show
    Else
        Exit Sub
    End If
End If
End Sub

Private Sub CmdSearch_Click()
FormSearch.Show
End Sub

Private Sub Form_Load()
DTTanggalPembayaran.Value = Format(Now, "dd/mm/yyyy")
Call NIM
Call CmdRefresh_Click
With ComboSemester
    .Clear
    .AddItem "I"
    .AddItem "II"
    .AddItem "III"
    .AddItem "IV"
    .AddItem "V"
    .AddItem "VI"
End With

With ComboJenisPembayaran
    .Clear
    .AddItem "Kuliah"
    .AddItem "Asrama"
    .AddItem "Pembangunan"
    .AddItem "PBL Rumkit"
    .AddItem "PBL Klinik"
    .AddItem "PBL Puskesmas"
    .AddItem "PBL RSU Adam Malik"
    .AddItem "PBL Desa"
    .AddItem "Caping Day"
    .AddItem "Wisuda"
    .AddItem "Ujian"
    .AddItem "Ujian 1"
    .AddItem "Ujian 2"
    .AddItem "Ujian 3"
    .AddItem "Ujian 4"
    .AddItem "Ujian 5"
    .AddItem "Klinik 1"
    .AddItem "Klinik 2"
End With

With ComboStatus
    .Clear
    .AddItem "Lunas"
    .AddItem "Belum Lunas"
End With
End Sub

Private Sub NIM()
BukaDatabase
RS.Open "SELECT nim FROM tb_mahasiswa", Conn, adOpenDynamic, adLockOptimistic
ComboNIM.Clear
Do While Not RS.EOF
    ComboNIM.AddItem RS!NIM
    RS.MoveNext
Loop
End Sub

Private Sub TextDenda_Change()
TextTotalBayar.Text = Val(TextDenda.Text) + Val(TextPembayaran.Text)
End Sub

Private Sub TextPembayaran_Change()
TextTotalBayar.Text = Val(TextDenda.Text) + Val(TextPembayaran.Text)
End Sub

Private Sub PengaturanDataGridTransaksi()
With DataGridTransaksi
    .Columns(0).Caption = "NO"
    .Columns(1).Caption = "NIM"
    .Columns(2).Caption = "NAMA"
    .Columns(3).Caption = "JENIS PEMBAYARAN"
    .Columns(4).Caption = "STATUS"
End With
End Sub

Private Sub ComboNIM_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
ComboSemester.SetFocus
End If
End Sub

Private Sub ComboSemester_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
ComboJenisPembayaran.SetFocus
End If
End Sub

Private Sub ComboJenisPembayaran_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
ComboStatus.SetFocus
End If
End Sub

Private Sub ComboStatus_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TextPembayaran.SetFocus
End If
End Sub

Private Sub TextPembayaran_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TextDenda.SetFocus
End If

If Not (KeyAscii >= Asc("0") & Chr(13) _
     And KeyAscii <= Asc("9") & Chr(13) _
     Or KeyAscii = vbKeyBack _
     Or KeyAscii = vbKeyDelete _
     Or KeyAscii = vbKeySpace) Then
    TextPembayaran.Text = "0"
    KeyAscii = 0
End If
End Sub

Private Sub TextDenda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
CmdSave.SetFocus
End If

If Not (KeyAscii >= Asc("0") & Chr(13) _
     And KeyAscii <= Asc("9") & Chr(13) _
     Or KeyAscii = vbKeyBack _
     Or KeyAscii = vbKeyDelete _
     Or KeyAscii = vbKeySpace) Then
    TextDenda.Text = "0"
    KeyAscii = 0
End If
End Sub

Private Sub PengaturanDataGridPembayaranHead()
With FormBuktiPembayaran.DataGridPembayaranHead
    .Columns(0).Caption = "NIM"
    .Columns(1).Caption = "NAMA"
End With
End Sub

Private Sub PengaturanDataGridPembayaranDetail()
With FormBuktiPembayaran.DataGridPembayaranDetail
    .Columns(0).Caption = "NO"
    .Columns(1).Caption = "NIM"
    .Columns(2).Caption = "SEMESTER"
    .Columns(3).Caption = "JENIS PEMBAYARAN"
    .Columns(4).Caption = "TANGGAL PEMBAYARAN"
    .Columns(5).Caption = "DENDA"
    .Columns(6).Caption = "STATUS"
    .Columns(7).Caption = "PEMBAYARAN"
    .Columns(8).Caption = "TOTAL PEMBAYARAN"
End With
End Sub
