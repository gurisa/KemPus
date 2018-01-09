VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FormBuktiPembayaran 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bukti Pembayaran"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9855
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormBuktiPembayaran.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   9855
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBukti 
      Caption         =   "&Bukti Pembayaran"
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      Begin VB.TextBox TextSimpanJenisPembayaran 
         Height          =   375
         Left            =   5760
         TabIndex        =   6
         Top             =   1200
         Visible         =   0   'False
         Width           =   2655
      End
      Begin Crystal.CrystalReport CrystalReport 
         Left            =   9120
         Top             =   1440
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowState     =   2
         PrintFileLinesPerPage=   60
      End
      Begin VB.ComboBox ComboKategori 
         Height          =   330
         Left            =   5760
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   3735
      End
      Begin VB.TextBox TextSearch 
         Height          =   375
         Left            =   5760
         TabIndex        =   3
         ToolTipText     =   "Cari Data Berdasarkan Kategori"
         Top             =   720
         Width           =   3735
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "Print"
         Height          =   375
         Left            =   8280
         TabIndex        =   1
         Top             =   4920
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid DataGridPembayaranHead 
         Height          =   1575
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   2778
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   17
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Begin MSDataGridLib.DataGrid DataGridPembayaranDetail 
         Height          =   2895
         Left            =   480
         TabIndex        =   4
         Top             =   1920
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   5106
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   17
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
End
Attribute VB_Name = "FormBuktiPembayaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub PengaturanDataGridPembayaranHead()
With DataGridPembayaranHead
    .Columns(0).Caption = "NIM"
    .Columns(1).Caption = "NAMA"
End With
End Sub

Private Sub PengaturanDataGridPembayaranDetail()
With DataGridPembayaranDetail
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

Private Sub CmdPrint_Click()
If MsgBox("Print Bukti Pembayaran ?", vbInformation + vbYesNo, "Print Bukti Pembayaran") = vbYes Then
        If Not RS.EOF Then
        BukaDatabase
        RS.Open "SELECT * FROM tb_pembayaran WHERE no_transaksi=" & DataGridPembayaranDetail.Columns(0).Text & "", Conn, adOpenDynamic, adLockOptimistic
        If RS!jenis_pembayaran = "Kuliah" Then
                TextSimpanJenisPembayaran.Text = "UK-"
            ElseIf RS!jenis_pembayaran = "Asrama" Then
                TextSimpanJenisPembayaran.Text = "UA-"
            ElseIf RS!jenis_pembayaran = "Pembangunan" Then
                TextSimpanJenisPembayaran.Text = "UP-"
            ElseIf RS!jenis_pembayaran = "PBL Rumkit" Then
                TextSimpanJenisPembayaran.Text = "PR-"
            ElseIf RS!jenis_pembayaran = "PBL Klinik" Then
                TextSimpanJenisPembayaran.Text = "PK-"
            ElseIf RS!jenis_pembayaran = "PBL Puskesmas" Then
                TextSimpanJenisPembayaran.Text = "PP-"
            ElseIf RS!jenis_pembayaran = "PBL RSU Adam Malik" Then
                TextSimpanJenisPembayaran.Text = "PRA-"
            ElseIf RS!jenis_pembayaran = "PBL Desa" Then
                TextSimpanJenisPembayaran.Text = "PBD-"
            ElseIf RS!jenis_pembayaran = "Caping Day" Then
                TextSimpanJenisPembayaran.Text = "CD-"
            ElseIf RS!jenis_pembayaran = "Wisuda" Then
                TextSimpanJenisPembayaran.Text = "UW-"
            ElseIf RS!jenis_pembayaran = "Ujian" Then
                TextSimpanJenisPembayaran.Text = "U-"
            ElseIf RS!jenis_pembayaran = "Ujian 1" Then
                TextSimpanJenisPembayaran.Text = "U1-"
            ElseIf RS!jenis_pembayaran = "Ujian 2" Then
                TextSimpanJenisPembayaran.Text = "U2-"
            ElseIf RS!jenis_pembayaran = "Ujian 3" Then
                TextSimpanJenisPembayaran.Text = "U3-"
            ElseIf RS!jenis_pembayaran = "Ujian 4" Then
                TextSimpanJenisPembayaran.Text = "U4-"
            ElseIf RS!jenis_pembayaran = "Ujian 5" Then
                TextSimpanJenisPembayaran.Text = "U5-"
            ElseIf RS!jenis_pembayaran = "Klinik 1" Then
                TextSimpanJenisPembayaran.Text = "K1-"
            ElseIf RS!jenis_pembayaran = "Klinik 2" Then
                TextSimpanJenisPembayaran.Text = "K2-"
            End If
        If Not RS.EOF Then
            BukaDatabase
            RS.Open "SELECT nama_mahasiswa FROM tb_mahasiswa WHERE nim='" & DataGridPembayaranDetail.Columns(1).Text & "'", Conn, adOpenDynamic, adLockOptimistic
            With CrystalReport
                .ReportFileName = App.Path & "\BuktiPembayaran.rpt"
                .Formulas(0) = "NamaMahasiswa = '" & RS!nama_mahasiswa & "'"
                .Formulas(1) = "JenisPembayaran = '" & TextSimpanJenisPembayaran.Text & "'"
                .SelectionFormula = "{tb_pembayaran.no_transaksi} = " & DataGridPembayaranDetail.Columns(0).Text & ""
                .RetrieveDataFiles
                .WindowShowPrintSetupBtn = True
                .WindowShowPrintBtn = True
                .PrintReport
            End With
        Else
            MsgBox "Data Pembayaran Tidak Tersedia", vbExclamation, "Data Pembayaran Tidak Tersedia"
        End If
        Else
            MsgBox "Data Pembayaran Tidak Tersedia", vbExclamation, "Data Pembayaran Tidak Tersedia"
        End If
Else
    Exit Sub
End If
End Sub

Private Sub DataGridPembayaranHead_Click()
BukaDatabase
RS.Open "SELECT nim, nama_mahasiswa FROM tb_mahasiswa", Conn, adOpenDynamic, adLockOptimistic
If Not RS.EOF Then

BukaDatabase
RS.Open "SELECT * FROM tb_pembayaran WHERE nim='" & DataGridPembayaranHead.Columns(0).Text & "'", Conn, adOpenDynamic, adLockOptimistic
If Not RS.EOF Then
    BukaDatabase
    RS.Open "SELECT * FROM tb_pembayaran WHERE nim='" & DataGridPembayaranHead.Columns(0).Text & "'", Conn, adOpenDynamic, adLockOptimistic
    Set DataGridPembayaranDetail.DataSource = RS.DataSource
    
    PengaturanDataGridPembayaranDetail
Else
    MsgBox "Tidak Terdapat Data Pembayaran", vbExclamation, "Tidak Terdapat Data Pembayaran"
End If
Else
    MsgBox "Tidak Terdapat Data Mahasiswa", vbExclamation, "Tidak Terdapat Data Mahasiswa"
End If
End Sub

Private Sub Form_Load()
BukaDatabase
RS.Open "SELECT nim, nama_mahasiswa FROM tb_mahasiswa", Conn, adOpenDynamic, adLockOptimistic
If Not RS.EOF Then
Set DataGridPembayaranHead.DataSource = RS.DataSource

    PengaturanDataGridPembayaranHead

    BukaDatabase
    RS.Open "SELECT * FROM tb_pembayaran WHERE nim='" & DataGridPembayaranHead.Columns(0).Text & "'", Conn, adOpenDynamic, adLockOptimistic
    Set DataGridPembayaranDetail.DataSource = RS.DataSource
    
    PengaturanDataGridPembayaranDetail
Else
    Exit Sub
End If

With ComboKategori
    .Clear
    .AddItem "NIM"
    .AddItem "NAMA"
End With
End Sub

Private Sub TextSearch_Change()
BukaDatabase
If ComboKategori.Text = "" Or TextSearch.Text = "" Then
    MsgBox "Gunakan Kriteria Pencarian", vbExclamation, "Pilih Kriteria Pencarian"
ElseIf ComboKategori.Text = "NIM" Then
    RS.Open "SELECT nim, nama_mahasiswa FROM tb_mahasiswa WHERE nim LIKE '%" & TextSearch.Text & "%'", Conn, adOpenDynamic, adLockOptimistic
    Set DataGridPembayaranHead.DataSource = RS.DataSource
    PengaturanDataGridPembayaranHead
ElseIf ComboKategori.Text = "NAMA" Then
    RS.Open "SELECT nim, nama_mahasiswa FROM tb_mahasiswa WHERE nama_mahasiswa LIKE '%" & TextSearch.Text & "%'", Conn, adOpenDynamic, adLockOptimistic
    Set DataGridPembayaranHead.DataSource = RS.DataSource
    PengaturanDataGridPembayaranHead
End If
End Sub
