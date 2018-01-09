VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FormSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10590
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormSearch.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   10590
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameCari 
      Caption         =   "&Search"
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10335
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Search"
         Height          =   375
         Left            =   9240
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox ComboKategoriCari 
         Height          =   330
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox TextCari 
         Height          =   375
         Left            =   3600
         TabIndex        =   1
         Top             =   360
         Width           =   5535
      End
      Begin MSDataGridLib.DataGrid DataGridCari 
         Height          =   4815
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   8493
         _Version        =   393216
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
      Begin VB.Label LabelNamaMahasiswa 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   6855
      End
   End
End
Attribute VB_Name = "FormSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdSearch_Click()
If ComboKategoriCari.Text = "" Or TextCari.Text = "" Then
    MsgBox "Gunakan Kriteria Pencarian", vbExclamation, "Pilih Kriteria Pencarian"
Else
    BukaDatabase
    If ComboKategoriCari.Text = "NIM" Then
        RS.Open "SELECT tb_mahasiswa.nama_mahasiswa, tb_pembayaran.no_transaksi, tb_pembayaran.nim, tb_pembayaran.semester, tb_pembayaran.jenis_pembayaran, tb_pembayaran.tanggal_pembayaran, tb_pembayaran.denda, tb_pembayaran.status, tb_pembayaran.pembayaran, tb_pembayaran.total_pembayaran FROM tb_mahasiswa, tb_pembayaran WHERE tb_mahasiswa.nim = tb_pembayaran.nim AND tb_mahasiswa.nim = '" & TextCari.Text & "' AND tb_pembayaran.nim = '" & TextCari.Text & "'", Conn, adOpenDynamic, adLockOptimistic
        Set DataGridCari.DataSource = RS.DataSource
        PengaturanDataGridCari
    End If
End If
End Sub

Private Sub Form_Load()
With ComboKategoriCari
    .Clear
    .AddItem "NIM"
End With

    BukaDatabase
    RS.Open "SELECT tb_mahasiswa.nama_mahasiswa, tb_pembayaran.no_transaksi, tb_pembayaran.nim, tb_pembayaran.semester, tb_pembayaran.jenis_pembayaran, tb_pembayaran.tanggal_pembayaran, tb_pembayaran.denda, tb_pembayaran.status, tb_pembayaran.pembayaran, tb_pembayaran.total_pembayaran FROM tb_mahasiswa, tb_pembayaran WHERE tb_mahasiswa.nim = tb_pembayaran.nim", Conn, adOpenDynamic, adLockOptimistic
    Set DataGridCari.DataSource = RS.DataSource
    PengaturanDataGridCari
End Sub

Private Sub PengaturanDataGridCari()
With DataGridCari
    .Columns(0).Caption = "NAMA"
    .Columns(1).Caption = "NO"
    .Columns(2).Caption = "NIM"
    .Columns(3).Caption = "SEMESTER"
    .Columns(4).Caption = "JENIS PEMBAYARAN"
    .Columns(5).Caption = "TANGGAL PEMBAYARAN"
    .Columns(6).Caption = "DENDA"
    .Columns(7).Caption = "STATUS"
    .Columns(8).Caption = "PEMBAYARAN"
    .Columns(9).Caption = "TOTAL PEMBAYARAN"
End With
End Sub

