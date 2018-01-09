VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FormReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3780
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormReport.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   3780
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameReport 
      Caption         =   "&Report"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.TextBox TextJenisPembayaran 
         Height          =   315
         Left            =   2640
         TabIndex        =   10
         Top             =   2160
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox ComboJenisPembayaran 
         Height          =   330
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   840
         Width           =   3255
      End
      Begin VB.TextBox TextTotalSemua 
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox TextTotalPembayaran 
         Height          =   255
         Left            =   1680
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox TextTotalDenda 
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton CmdPreview 
         Caption         =   "Preview"
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "Print"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   1575
      End
      Begin VB.ComboBox ComboKategoriLaporan 
         Height          =   330
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   3255
      End
      Begin MSComCtl2.DTPicker DTAwal 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   90439681
         CurrentDate     =   41925
      End
      Begin MSComCtl2.DTPicker DTAkhir 
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   90439681
         CurrentDate     =   41925
      End
   End
   Begin Crystal.CrystalReport CrystalReport 
      Left            =   0
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "FormReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TanggalAwal As Date
Dim TanggalAkhir As Date

Private Sub CmdPreview_Click()
If ComboKategoriLaporan.Text = "Laporan Keuangan" Then
    If ComboJenisPembayaran.Text = "Keseluruhan" Then
        BukaDatabase
        RS.Open "SELECT * FROM tb_pembayaran", Conn, adOpenDynamic, adLockOptimistic
        If Not RS.EOF Then
            TanggalAwal = Format(DTAwal.Value, "yyyy/MM/dd")
            TanggalAkhir = Format(DTAkhir.Value, "yyyy/MM/dd")
            BukaDatabase
            RS.Open "SELECT SUM(denda) AS TotalDenda FROM tb_pembayaran WHERE (tanggal_pembayaran >= #" & Format(DTAwal.Value, "yyyy/MM/dd") & "# AND tanggal_pembayaran <= #" & Format(DTAkhir.Value, "yyyy/MM/dd") & "#)", Conn, adOpenDynamic, adLockOptimistic
            If RS!TotalDenda <> "" Then
            TextTotalDenda.Text = RS!TotalDenda
            Else
                Exit Sub
            End If
            RS.Close
            RS.Open "SELECT SUM(pembayaran) AS TotalPembayaran FROM tb_pembayaran WHERE (tanggal_pembayaran >= #" & Format(DTAwal.Value, "yyyy/MM/dd") & "# AND tanggal_pembayaran <= #" & Format(DTAkhir.Value, "yyyy/MM/dd") & "#)", Conn, adOpenDynamic, adLockOptimistic
            If RS!TotalPembayaran <> "" Then
            TextTotalPembayaran.Text = RS!TotalPembayaran
            Else
                Exit Sub
            End If
            RS.Close
            RS.Open "SELECT SUM(total_pembayaran) AS TotalSemua FROM tb_pembayaran WHERE (tanggal_pembayaran >= #" & Format(DTAwal.Value, "yyyy/MM/dd") & "# AND tanggal_pembayaran <= #" & Format(DTAkhir.Value, "yyyy/MM/dd") & "#)", Conn, adOpenDynamic, adLockOptimistic
            If RS!TotalSemua <> "" Then
            TextTotalSemua.Text = RS!TotalSemua
            Else
                Exit Sub
            End If
            RS.Close
            With CrystalReport
                .ReportFileName = App.Path & "\LaporanKeuangan.rpt"
                .Formulas(0) = "TglAwal = '" & TanggalAwal & "'"
                .Formulas(1) = "TglAkhir = '" & TanggalAkhir & "'"
                .Formulas(2) = "TotalDenda = '" & TextTotalDenda.Text & "'"
                .Formulas(3) = "TotalPembayaran = '" & TextTotalPembayaran.Text & "'"
                .Formulas(4) = "TotalSemua = '" & TextTotalSemua.Text & "'"
                .SelectionFormula = "{tb_pembayaran.tanggal_pembayaran} >= #" & Format(DTAwal.Value, "yyyy/MM/dd") & "# AND {tb_pembayaran.tanggal_pembayaran} <= #" & Format(DTAkhir.Value, "yyyy/MM/dd") & "#"
                .RetrieveDataFiles
                .WindowShowPrintSetupBtn = True
                .WindowShowPrintBtn = True
                .Action = 1
            End With
        Else
            MsgBox "Data Pembayaran Tidak Tersedia", vbExclamation, "Data Pembayaran Tidak Tersedia"
        End If
    Else
        BukaDatabase
        RS.Open "SELECT * FROM tb_pembayaran", Conn, adOpenDynamic, adLockOptimistic
        If Not RS.EOF Then
            TanggalAwal = Format(DTAwal.Value, "yyyy/MM/dd")
            TanggalAkhir = Format(DTAkhir.Value, "yyyy/MM/dd")
            BukaDatabase
            RS.Open "SELECT SUM(denda) AS TotalDenda FROM tb_pembayaran WHERE (tanggal_pembayaran >= #" & Format(DTAwal.Value, "yyyy/MM/dd") & "# AND tanggal_pembayaran <= #" & Format(DTAkhir.Value, "yyyy/MM/dd") & "# AND jenis_pembayaran='" & ComboJenisPembayaran.Text & "')", Conn, adOpenDynamic, adLockOptimistic
            If RS!TotalDenda <> "" Then
            TextTotalDenda.Text = RS!TotalDenda
            Else
                Exit Sub
            End If
            RS.Close
            RS.Open "SELECT SUM(pembayaran) AS TotalPembayaran FROM tb_pembayaran WHERE (tanggal_pembayaran >= #" & Format(DTAwal.Value, "yyyy/MM/dd") & "# AND tanggal_pembayaran <= #" & Format(DTAkhir.Value, "yyyy/MM/dd") & "# AND jenis_pembayaran='" & ComboJenisPembayaran.Text & "')", Conn, adOpenDynamic, adLockOptimistic
            If RS!TotalPembayaran <> "" Then
            TextTotalPembayaran.Text = RS!TotalPembayaran
            Else
                Exit Sub
            End If
            RS.Close
            RS.Open "SELECT SUM(total_pembayaran) AS TotalSemua FROM tb_pembayaran WHERE (tanggal_pembayaran >= #" & Format(DTAwal.Value, "yyyy/MM/dd") & "# AND tanggal_pembayaran <= #" & Format(DTAkhir.Value, "yyyy/MM/dd") & "# AND jenis_pembayaran='" & ComboJenisPembayaran.Text & "')", Conn, adOpenDynamic, adLockOptimistic
            If RS!TotalSemua <> "" Then
            TextTotalSemua.Text = RS!TotalSemua
            Else
                Exit Sub
            End If
            RS.Close
            With CrystalReport
                .ReportFileName = App.Path & "\LaporanKeuangan.rpt"
                .Formulas(0) = "TglAwal = '" & TanggalAwal & "'"
                .Formulas(1) = "TglAkhir = '" & TanggalAkhir & "'"
                .Formulas(2) = "TotalDenda = '" & TextTotalDenda.Text & "'"
                .Formulas(3) = "TotalPembayaran = '" & TextTotalPembayaran.Text & "'"
                .Formulas(4) = "TotalSemua = '" & TextTotalSemua.Text & "'"
                .SelectionFormula = "{tb_pembayaran.tanggal_pembayaran} >= #" & Format(DTAwal.Value, "yyyy/MM/dd") & "# AND {tb_pembayaran.tanggal_pembayaran} <= #" & Format(DTAkhir.Value, "yyyy/MM/dd") & "# AND {tb_pembayaran.jenis_pembayaran} = '" & ComboJenisPembayaran.Text & "'"
                .RetrieveDataFiles
                .WindowShowPrintSetupBtn = True
                .WindowShowPrintBtn = True
                .Action = 1
            End With
        Else
            MsgBox "Data Pembayaran Tidak Tersedia", vbExclamation, "Data Pembayaran Tidak Tersedia"
        End If
    End If
End If
End Sub

Private Sub CmdPrint_Click()
If ComboKategoriLaporan.Text = "Laporan Keuangan" Then
    If ComboJenisPembayaran.Text = "Keseluruhan" Then
        BukaDatabase
        RS.Open "SELECT * FROM tb_pembayaran", Conn, adOpenDynamic, adLockOptimistic
        If Not RS.EOF Then
            TanggalAwal = Format(DTAwal.Value, "yyyy/MM/dd")
            TanggalAkhir = Format(DTAkhir.Value, "yyyy/MM/dd")
            BukaDatabase
            RS.Open "SELECT SUM(denda) AS TotalDenda FROM tb_pembayaran WHERE (tanggal_pembayaran >= #" & Format(DTAwal.Value, "yyyy/MM/dd") & "# AND tanggal_pembayaran <= #" & Format(DTAkhir.Value, "yyyy/MM/dd") & "#)", Conn, adOpenDynamic, adLockOptimistic
            If RS!TotalDenda <> "" Then
            TextTotalDenda.Text = RS!TotalDenda
            Else
                Exit Sub
            End If
            RS.Close
            RS.Open "SELECT SUM(pembayaran) AS TotalPembayaran FROM tb_pembayaran WHERE (tanggal_pembayaran >= #" & Format(DTAwal.Value, "yyyy/MM/dd") & "# AND tanggal_pembayaran <= #" & Format(DTAkhir.Value, "yyyy/MM/dd") & "#)", Conn, adOpenDynamic, adLockOptimistic
            If RS!TotalPembayaran <> "" Then
            TextTotalPembayaran.Text = RS!TotalPembayaran
            Else
                Exit Sub
            End If
            RS.Close
            RS.Open "SELECT SUM(total_pembayaran) AS TotalSemua FROM tb_pembayaran WHERE (tanggal_pembayaran >= #" & Format(DTAwal.Value, "yyyy/MM/dd") & "# AND tanggal_pembayaran <= #" & Format(DTAkhir.Value, "yyyy/MM/dd") & "#)", Conn, adOpenDynamic, adLockOptimistic
            If RS!TotalSemua <> "" Then
            TextTotalSemua.Text = RS!TotalSemua
            Else
                Exit Sub
            End If
            RS.Close
            With CrystalReport
                .ReportFileName = App.Path & "\LaporanKeuangan.rpt"
                .Formulas(0) = "TglAwal = '" & TanggalAwal & "'"
                .Formulas(1) = "TglAkhir = '" & TanggalAkhir & "'"
                .Formulas(2) = "TotalDenda = '" & TextTotalDenda.Text & "'"
                .Formulas(3) = "TotalPembayaran = '" & TextTotalPembayaran.Text & "'"
                .Formulas(4) = "TotalSemua = '" & TextTotalSemua.Text & "'"
                .SelectionFormula = "{tb_pembayaran.tanggal_pembayaran} >= #" & Format(DTAwal.Value, "yyyy/MM/dd") & "# AND {tb_pembayaran.tanggal_pembayaran} <= #" & Format(DTAkhir.Value, "yyyy/MM/dd") & "#"
                .RetrieveDataFiles
                .WindowShowPrintSetupBtn = True
                .WindowShowPrintBtn = True
                .PrintReport
            End With
        Else
            MsgBox "Data Pembayaran Tidak Tersedia", vbExclamation, "Data Pembayaran Tidak Tersedia"
        End If
    Else
        BukaDatabase
        RS.Open "SELECT * FROM tb_pembayaran", Conn, adOpenDynamic, adLockOptimistic
        If Not RS.EOF Then
            TanggalAwal = Format(DTAwal.Value, "yyyy/MM/dd")
            TanggalAkhir = Format(DTAkhir.Value, "yyyy/MM/dd")
            BukaDatabase
            RS.Open "SELECT SUM(denda) AS TotalDenda FROM tb_pembayaran WHERE (tanggal_pembayaran >= #" & Format(DTAwal.Value, "yyyy/MM/dd") & "# AND tanggal_pembayaran <= #" & Format(DTAkhir.Value, "yyyy/MM/dd") & "# AND jenis_pembayaran='" & ComboJenisPembayaran.Text & "')", Conn, adOpenDynamic, adLockOptimistic
            If RS!TotalDenda <> "" Then
            TextTotalDenda.Text = RS!TotalDenda
            Else
                Exit Sub
            End If
            RS.Close
            RS.Open "SELECT SUM(pembayaran) AS TotalPembayaran FROM tb_pembayaran WHERE (tanggal_pembayaran >= #" & Format(DTAwal.Value, "yyyy/MM/dd") & "# AND tanggal_pembayaran <= #" & Format(DTAkhir.Value, "yyyy/MM/dd") & "# AND jenis_pembayaran='" & ComboJenisPembayaran.Text & "')", Conn, adOpenDynamic, adLockOptimistic
            If RS!TotalPembayaran <> "" Then
            TextTotalPembayaran.Text = RS!TotalPembayaran
            Else
                Exit Sub
            End If
            RS.Close
            RS.Open "SELECT SUM(total_pembayaran) AS TotalSemua FROM tb_pembayaran WHERE (tanggal_pembayaran >= #" & Format(DTAwal.Value, "yyyy/MM/dd") & "# AND tanggal_pembayaran <= #" & Format(DTAkhir.Value, "yyyy/MM/dd") & "# AND jenis_pembayaran='" & ComboJenisPembayaran.Text & "')", Conn, adOpenDynamic, adLockOptimistic
            If RS!TotalSemua <> "" Then
            TextTotalSemua.Text = RS!TotalSemua
            Else
                Exit Sub
            End If
            RS.Close
            With CrystalReport
                .ReportFileName = App.Path & "\LaporanKeuangan.rpt"
                .Formulas(0) = "TglAwal = '" & TanggalAwal & "'"
                .Formulas(1) = "TglAkhir = '" & TanggalAkhir & "'"
                .Formulas(2) = "TotalDenda = '" & TextTotalDenda.Text & "'"
                .Formulas(3) = "TotalPembayaran = '" & TextTotalPembayaran.Text & "'"
                .Formulas(4) = "TotalSemua = '" & TextTotalSemua.Text & "'"
                .SelectionFormula = "{tb_pembayaran.tanggal_pembayaran} >= #" & Format(DTAwal.Value, "yyyy/MM/dd") & "# AND {tb_pembayaran.tanggal_pembayaran} <= #" & Format(DTAkhir.Value, "yyyy/MM/dd") & "# AND {tb_pembayaran.jenis_pembayaran} = '" & ComboJenisPembayaran.Text & "'"
                .RetrieveDataFiles
                .WindowShowPrintSetupBtn = True
                .WindowShowPrintBtn = True
                .PrintReport
            End With
        Else
            MsgBox "Data Pembayaran Tidak Tersedia", vbExclamation, "Data Pembayaran Tidak Tersedia"
        End If
    End If
End If
End Sub

Private Sub ComboJenisPembayaran_Change()
If ComboJenisPembayaran.Text = "" Then

End If
End Sub

Private Sub Form_Load()
With ComboKategoriLaporan
    .Clear
    .AddItem "Laporan Keuangan"
End With

With ComboJenisPembayaran
    .Clear
    .AddItem "Keseluruhan"
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

DTAwal.Value = Format(Now, "dd/mm/yyyy")
DTAkhir.Value = Format(Now, "dd/mm/yyyy")
End Sub
