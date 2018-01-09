VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormUtama 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Main Menu"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   705
   ClientWidth     =   9045
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormUtama.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   9045
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker DTPembanding 
      Height          =   375
      Left            =   480
      TabIndex        =   51
      Top             =   5400
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   91029505
      CurrentDate     =   41927
   End
   Begin VB.Frame FrameLisensi 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Lisensi"
      Height          =   5655
      Left            =   2160
      TabIndex        =   27
      Top             =   120
      Width           =   6735
      Begin VB.Frame FrameLicenseChange 
         BackColor       =   &H00FFFFFF&
         Height          =   3255
         Left            =   120
         TabIndex        =   33
         Top             =   2280
         Width           =   6495
         Begin VB.Frame FrameDonate 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Donate"
            Height          =   2895
            Left            =   3840
            TabIndex        =   45
            Top             =   240
            Width           =   2535
            Begin VB.TextBox TextBitCoin 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Consolas"
                  Size            =   6
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   52
               Text            =   "1QLP8hWB49Pvxg4uhPcKKqWqN2ihwQ5H29"
               Top             =   2640
               Width           =   2295
            End
            Begin VB.Image ImageBitCoin 
               Height          =   855
               Left            =   120
               Picture         =   "FormUtama.frx":0CCA
               Stretch         =   -1  'True
               Top             =   1680
               Width           =   2295
            End
            Begin VB.Image ImageQRCode 
               Height          =   1335
               Left            =   120
               Picture         =   "FormUtama.frx":1F04F
               Stretch         =   -1  'True
               Top             =   240
               Width           =   2295
            End
         End
         Begin VB.Frame FrameLicensed 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Unlicensed Software!"
            ForeColor       =   &H00000000&
            Height          =   2895
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   3615
            Begin VB.CommandButton CmdRefreshLisensi 
               Cancel          =   -1  'True
               Caption         =   "Refresh"
               Height          =   375
               Left            =   120
               TabIndex        =   49
               Top             =   2400
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.TextBox TextSimpanTanggal 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   375
               Left            =   1200
               TabIndex        =   48
               Top             =   2400
               Visible         =   0   'False
               Width           =   2295
            End
            Begin VB.CommandButton CmdGenerate 
               Caption         =   "Generate"
               Height          =   375
               Left            =   1200
               TabIndex        =   47
               Top             =   1920
               Width           =   1095
            End
            Begin VB.CommandButton CmdActivation 
               Caption         =   "Activate"
               Height          =   375
               Left            =   2400
               TabIndex        =   46
               Top             =   1920
               Width           =   1095
            End
            Begin VB.Timer TimerLisensi 
               Left            =   120
               Top             =   1920
            End
            Begin VB.TextBox TextUserLicense 
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   1200
               MaxLength       =   25
               TabIndex        =   41
               Top             =   480
               Width           =   2295
            End
            Begin VB.TextBox TextCodeSend 
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   1200
               TabIndex        =   40
               Top             =   1440
               Width           =   2295
            End
            Begin VB.TextBox TextCodeGet 
               Appearance      =   0  'Flat
               Height          =   375
               Left            =   1200
               Locked          =   -1  'True
               TabIndex        =   39
               Top             =   960
               Width           =   2295
            End
            Begin VB.Label LabelCountDown 
               BackStyle       =   0  'Transparent
               Height          =   375
               Left            =   1200
               TabIndex        =   50
               Top             =   2400
               Width           =   2295
            End
            Begin VB.Label LabelCodeGet 
               BackStyle       =   0  'Transparent
               Caption         =   "Code Get"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   44
               Top             =   960
               Width           =   975
            End
            Begin VB.Label LabelCodeSend 
               BackStyle       =   0  'Transparent
               Caption         =   "Code Send"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   43
               Top             =   1440
               Width           =   975
            End
            Begin VB.Label LabelLicenseName 
               BackStyle       =   0  'Transparent
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   42
               Top             =   480
               Width           =   975
            End
         End
      End
      Begin VB.Frame FrameLicenseStatus 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Status"
         Height          =   1935
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   6495
         Begin VB.TextBox TextDateLicenseStatus 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   37
            Text            =   "-"
            Top             =   720
            Width           =   4575
         End
         Begin VB.TextBox TextCodeGetStatus 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   36
            Text            =   "-"
            Top             =   1080
            Width           =   4575
         End
         Begin VB.TextBox TextCodeSendStatus 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   35
            Text            =   "-"
            Top             =   1440
            Width           =   4575
         End
         Begin VB.TextBox TextUserLicenseStatus 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   315
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   34
            Text            =   "Trial Version"
            Top             =   360
            Width           =   4575
         End
         Begin VB.Label LabelCodeSendStatus 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "CodeSend"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label LabelDateLicenseStatus 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Date Validation"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label LabelCodeGetStatus 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Code Get"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label LabelUserLicenseStatus 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "User License"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   360
            Width           =   1575
         End
      End
   End
   Begin VB.Frame FramePetugas 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Petugas"
      Height          =   5655
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   6735
      Begin VB.Frame FrameSetPassword 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Left            =   240
         TabIndex        =   10
         Top             =   3240
         Width           =   2055
         Begin VB.CommandButton CmdSetPassword 
            Caption         =   "Set Password"
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   1680
            Width           =   1815
         End
         Begin VB.TextBox TextNewPassword 
            Appearance      =   0  'Flat
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   12
            ToolTipText     =   "New Password"
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox TextCurrentPassword 
            Appearance      =   0  'Flat
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   11
            ToolTipText     =   "Current Password"
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label LabelNewPassword 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "New Password"
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
            TabIndex        =   25
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label LabelCurrentPassword 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Current Password"
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
            TabIndex        =   24
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.TextBox TextUsername 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "1234567890"
         Top             =   360
         Width           =   3735
      End
      Begin VB.TextBox TextOtoritas 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "1234567890"
         Top             =   1320
         Width           =   3735
      End
      Begin VB.TextBox TextPassword 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2880
         Locked          =   -1  'True
         PasswordChar    =   "*"
         TabIndex        =   7
         Text            =   "1234567890"
         Top             =   840
         Width           =   3735
      End
      Begin VB.Frame FramePetugasDetail 
         BackColor       =   &H00FFFFFF&
         Height          =   3855
         Left            =   120
         TabIndex        =   2
         Top             =   1680
         Width           =   6495
         Begin VB.Frame FrameSetting 
            BackColor       =   &H00FFFFFF&
            Height          =   2175
            Left            =   2280
            TabIndex        =   14
            Top             =   1560
            Width           =   4095
            Begin VB.CommandButton CmdRefresh 
               Caption         =   "Refresh"
               Height          =   375
               Left            =   2880
               TabIndex        =   26
               Top             =   240
               Width           =   1095
            End
            Begin VB.CommandButton CmdDelete 
               Caption         =   "Delete"
               Height          =   375
               Left            =   2880
               TabIndex        =   23
               Top             =   1680
               Width           =   1095
            End
            Begin VB.CommandButton CmdEdit 
               Caption         =   "Edit"
               Height          =   375
               Left            =   2880
               TabIndex        =   22
               Top             =   1200
               Width           =   1095
            End
            Begin VB.CommandButton CmdAdd 
               Caption         =   "Add"
               Height          =   375
               Left            =   2880
               TabIndex        =   21
               Top             =   720
               Width           =   1095
            End
            Begin VB.ComboBox ComboOtoritas 
               Enabled         =   0   'False
               Height          =   330
               Left            =   1080
               Style           =   2  'Dropdown List
               TabIndex        =   17
               Top             =   1320
               Width           =   1575
            End
            Begin VB.TextBox TextPasswordNew 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   315
               IMEMode         =   3  'DISABLE
               Left            =   1080
               MaxLength       =   10
               PasswordChar    =   "*"
               TabIndex        =   16
               Top             =   840
               Width           =   1575
            End
            Begin VB.TextBox TextUsernameNew 
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               Height          =   315
               Left            =   1080
               MaxLength       =   10
               TabIndex        =   15
               Top             =   360
               Width           =   1575
            End
            Begin VB.Label LabelUsernameNew 
               BackStyle       =   0  'Transparent
               Caption         =   "Username"
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
               TabIndex        =   20
               Top             =   360
               Width           =   975
            End
            Begin VB.Label LabelPasswordNew 
               BackStyle       =   0  'Transparent
               Caption         =   "Password"
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
               TabIndex        =   19
               Top             =   840
               Width           =   975
            End
            Begin VB.Label LabelOtoritasNew 
               BackStyle       =   0  'Transparent
               Caption         =   "Otoritas"
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
               TabIndex        =   18
               Top             =   1320
               Width           =   975
            End
         End
         Begin MSDataGridLib.DataGrid DataGridPetugas 
            Height          =   1215
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   2143
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
      Begin VB.Label LabelJudulOtoritas 
         BackStyle       =   0  'Transparent
         Caption         =   "Otoritas"
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
         Left            =   1800
         TabIndex        =   5
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label LabelJudulPassword 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         Left            =   1800
         TabIndex        =   4
         Top             =   840
         Width           =   975
      End
      Begin VB.Label LabelJudulUsername 
         BackStyle       =   0  'Transparent
         Caption         =   "Username"
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
         Left            =   1800
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.Image ImageOperatorStatus 
         Height          =   1335
         Left            =   120
         Picture         =   "FormUtama.frx":22D33
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1455
      End
      Begin VB.Image ImagePetugasStatus 
         Height          =   1335
         Left            =   120
         Picture         =   "FormUtama.frx":26A4C
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Timer TimerUtama 
      Left            =   0
      Top             =   5400
   End
   Begin MSComctlLib.StatusBar StatusBarUtama 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5880
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7726
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Image ImageLisensiBlack 
      Height          =   1455
      Left            =   240
      Picture         =   "FormUtama.frx":2B419
      Stretch         =   -1  'True
      Top             =   1920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image ImageLisensiColor 
      Height          =   1455
      Left            =   240
      Picture         =   "FormUtama.frx":2E9F1
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Image ImagePetugasBlack 
      Height          =   1455
      Left            =   240
      Picture         =   "FormUtama.frx":31BE6
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image ImagePetugasColor 
      Height          =   1455
      Left            =   240
      Picture         =   "FormUtama.frx":35812
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1575
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Begin VB.Menu Tools 
         Caption         =   "Tools"
         Begin VB.Menu User 
            Caption         =   "User Management"
            Shortcut        =   {F12}
         End
         Begin VB.Menu Mahasiswa 
            Caption         =   "Mahasiswa"
            Shortcut        =   {F9}
         End
         Begin VB.Menu Pembayaran 
            Caption         =   "Pembayaran"
            Shortcut        =   {F8}
         End
         Begin VB.Menu Bukti 
            Caption         =   "Bukti Pembayaran"
            Shortcut        =   {F7}
         End
         Begin VB.Menu Search 
            Caption         =   "Search"
            Shortcut        =   {F6}
         End
         Begin VB.Menu Report 
            Caption         =   "Report"
            Shortcut        =   {F5}
         End
      End
      Begin VB.Menu Logout 
         Caption         =   "Logout"
         Shortcut        =   {F3}
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu Registration 
         Caption         =   "Registration"
         Shortcut        =   {F2}
      End
      Begin VB.Menu About 
         Caption         =   "About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "FormUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Jam As Integer, Menit As Integer, Detik As Integer

Private Sub About_Click()
FormAbout.Show
End Sub

Private Sub Bukti_Click()
FormBuktiPembayaran.Show
End Sub

Private Sub CmdActivation_Click()
If CmdActivation.Caption = "Activate" Then
    If TextUserLicense.Text = "" Or TextCodeGet.Text = "" Or TextCodeSend.Text = "" Then
        MsgBox "Masukan Data Aktivasi Dengan Benar", vbExclamation, "Data Aktivasi Tidak Valid"
    ElseIf Val(TextCodeGet.Text) + 5071997 = Val(TextCodeSend.Text) Then
        TextSimpanTanggal.Text = DateAdd("d", 30, Format(Date, "dd/mm/yyyy"))
        BukaDatabase
        RS.Open "UPDATE tb_lisensi SET code_activation = " & TextCodeSend.Text & ", code_get = " & TextCodeGet.Text & ", user_license = '" & TextUserLicense.Text & "', date_license = #" & TextSimpanTanggal.Text & "#", Conn, adOpenDynamic, adLockOptimistic
        Call CmdRefreshLisensi_Click
        MsgBox "Aktivasi Berhasil Di Lakukan", vbInformation, "Aktivasi Berhasil Di Lakukan"
    ElseIf Val(TextCodeGet.Text) + 12111997 = Val(TextCodeSend.Text) Then
        TextSimpanTanggal.Text = DateAdd("m", 6, Format(Date, "dd/mm/yyyy"))
        BukaDatabase
        RS.Open "UPDATE tb_lisensi SET code_activation = " & TextCodeSend.Text & ", code_get = " & TextCodeGet.Text & ", user_license = '" & TextUserLicense.Text & "', date_license = #" & TextSimpanTanggal.Text & "#", Conn, adOpenDynamic, adLockOptimistic
        Call CmdRefreshLisensi_Click
        MsgBox "Aktivasi Berhasil Di Lakukan", vbInformation, "Aktivasi Berhasil Di Lakukan"
    ElseIf Val(TextCodeGet.Text) + 7102011 = Val(TextCodeSend.Text) Then
        TextSimpanTanggal.Text = DateAdd("yyyy", 1, Format(Date, "dd/mm/yyyy"))
        BukaDatabase
        RS.Open "UPDATE tb_lisensi SET code_activation = " & TextCodeSend.Text & ", code_get = " & TextCodeGet.Text & ", user_license = '" & TextUserLicense.Text & "', date_license = #" & TextSimpanTanggal.Text & "#", Conn, adOpenDynamic, adLockOptimistic
        Call CmdRefreshLisensi_Click
        MsgBox "Aktivasi Berhasil Di Lakukan", vbInformation, "Aktivasi Berhasil Di Lakukan"
    Else
        MsgBox "Kode Aktivasi Salah", vbExclamation, "Kode Aktivasi Salah"
    End If
ElseIf CmdActivation.Caption = "Deactivate" Then
    If MsgBox("Batalkan Status Aktivasi ?", vbExclamation + vbYesNo, "Batalkan Status Aktivasi") = vbYes Then
        BukaDatabase
        RS.Open "UPDATE tb_lisensi SET code_activation = 0, code_get = 0, user_license = 'Trial Version', date_license = #11/12/1997#", Conn, adOpenDynamic, adLockOptimistic
        MsgBox "Berhasil Membatalkan Status Aktivasi", vbInformation, "Berhasil Membatalkan Status Aktivasi"
        Call CmdRefreshLisensi_Click
        Call CmdGenerate_Click
    Else
        Exit Sub
    End If
End If
End Sub

Private Sub CmdAdd_Click()
If CmdAdd.Caption = "Add" Then
    CmdAdd.Caption = "Save"
    TextUsernameNew.Enabled = True
    TextPasswordNew.Enabled = True
    ComboOtoritas.Enabled = True
    CmdEdit.Enabled = False
    CmdDelete.Enabled = False
    CmdSetPassword.Enabled = False
ElseIf CmdAdd.Caption = "Save" Then
    BukaDatabase
    RS.Open "SELECT id_petugas FROM tb_petugas WHERE id_petugas = '" & TextUsernameNew.Text & "'", Conn, adOpenDynamic, adLockOptimistic
    If RS.EOF Then
        RS.Close
        If TextUsernameNew.Text = "" Or TextPassword.Text = "" Or ComboOtoritas.Text = "" Then
            MsgBox "Masukan Data Petugas Dengan Benar", vbExclamation, "Masukan Data Petugas Dengan Benar"
        Else
            RS.Open "INSERT INTO tb_petugas(id_petugas, password_petugas, otoritas_petugas) VALUES('" & TextUsernameNew.Text & "','" & TextPasswordNew.Text & "','" & ComboOtoritas.Text & "')", Conn, adOpenDynamic, adLockOptimistic
            MsgBox "Berhasil Menambahkan " & ComboOtoritas.Text & "", vbInformation, "Berhasil Menambahkan " & ComboOtoritas.Text & ""
            CmdAdd.Caption = "Add"
            CmdEdit.Enabled = True
            CmdDelete.Enabled = True
            CmdSetPassword.Enabled = True
            TextUsernameNew.Enabled = False
            TextPasswordNew.Enabled = False
            ComboOtoritas.Enabled = False
            Call CmdRefresh_Click
        End If
    Else
        MsgBox "Username Sudah Di Gunakan", vbExclamation, "Username Sudah Di Gunakan"
    End If
End If
End Sub

Private Sub CmdDelete_Click()
If Not RS.EOF Then
    If MsgBox("Hapus Data " & DataGridPetugas.Columns(0).Text & " ?", vbExclamation + vbYesNo, "Hapus Data Petugas") = vbYes Then
        If DataGridPetugas.Columns(0).Text = "admin" Or DataGridPetugas.Columns(0).Text = "root" Then
            MsgBox "User Khusus Tidak Dapat Di Hapus", vbExclamation, "User Khusus Tidak Dapat Di Hapus"
        Else
            BukaDatabase
            RS.Open "DELETE FROM tb_petugas WHERE id_petugas='" & DataGridPetugas.Columns(0).Text & "'", Conn, adOpenDynamic, adLockOptimistic
            Call CmdRefresh_Click
            MsgBox "Data Petugas Berhasil Di Hapus", vbInformation, "Data Petugas Berhasil Di Hapus"
        End If
    Else
        Exit Sub
    End If
Else
    MsgBox "Data Petugas Tidak Tersedia", vbExclamation, "Data Petugas Tidak Tersedia"
End If
End Sub

Private Sub CmdEdit_Click()
If Not RS.EOF Then
    If CmdEdit.Caption = "Edit" Then
        BukaDatabase
        RS.Open "SELECT * FROM tb_petugas WHERE id_petugas = '" & DataGridPetugas.Columns(0).Text & "'", Conn, adOpenDynamic, adLockOptimistic
        TextUsernameNew.Text = RS!id_petugas
        TextPasswordNew.Text = RS!password_petugas
        ComboOtoritas.Text = RS!otoritas_petugas
        TextUsernameNew.Enabled = False
        TextPasswordNew.Enabled = True
        ComboOtoritas.Enabled = True
        CmdEdit.Caption = "Save"
        CmdSetPassword.Enabled = False
        CmdDelete.Enabled = False
        CmdAdd.Enabled = False
    ElseIf CmdEdit.Caption = "Save" Then
        BukaDatabase
        RS.Open "UPDATE tb_petugas SET password_petugas = '" & TextPasswordNew.Text & "', otoritas_petugas = '" & ComboOtoritas.Text & "' WHERE id_petugas = '" & TextUsernameNew.Text & "'", Conn, adOpenDynamic, adLockOptimistic
        MsgBox "Berhasil Mengubah Data Login", vbInformation, "Berhasil Mengubah Data Login"
        TextUsernameNew.Enabled = False
        TextPasswordNew.Enabled = False
        ComboOtoritas.Enabled = False
        CmdEdit.Caption = "Edit"
        CmdSetPassword.Enabled = True
        CmdDelete.Enabled = True
        CmdAdd.Enabled = True
        Call CmdRefresh_Click
    End If
Else
    MsgBox "Data Transaksi Tidak Tersedia", vbExclamation, "Data Transaksi Tidak Tersedia"
End If
End Sub

Private Sub CmdGenerate_Click()
TextCodeGet.Text = Format(Date, "dd") & Format(Time, "hh") & Format(Date, "mm") & Format(Time, "mm") & Format(Date, "yyyy") & Format(Time, "ss")
End Sub

Private Sub CmdRefresh_Click()
CmdAdd.Caption = "Add"
CmdEdit.Caption = "Edit"
CmdAdd.Enabled = True
CmdEdit.Enabled = True
CmdDelete.Enabled = True
CmdSetPassword.Enabled = True
TextUsernameNew.Enabled = False
TextPasswordNew.Enabled = False
ComboOtoritas.Enabled = False
            
TextCurrentPassword.Text = ""
TextNewPassword.Text = ""
TextUsernameNew.Text = ""
TextPasswordNew.Text = ""

BukaDatabase
RS.Open "SELECT id_petugas, otoritas_petugas FROM tb_petugas", Conn, adOpenDynamic, adLockOptimistic
Set DataGridPetugas.DataSource = RS.DataSource
PengaturanDataGridPetugas
End Sub

Private Sub CmdRefreshLisensi_Click()
BukaDatabase
RS.Open "SELECT * FROM tb_lisensi WHERE id_activation='1'", Conn, adOpenDynamic, adLockOptimistic
If RS.EOF Then
    Call Registration_Click
    MsgBox "Unlicensed Software", vbCritical, "Unlicensed Software"
Else
    TextUserLicenseStatus.Text = RS!user_license
    TextDateLicenseStatus.Text = RS!date_license
    TextCodeGetStatus.Text = RS!code_get
    TextCodeSendStatus.Text = RS!code_activation
    If Val(TextCodeGetStatus.Text) = 0 Or Val(TextCodeSendStatus.Text) = 0 Then
        TextUserLicense.Enabled = True
        TextCodeGet.Enabled = True
        TextCodeSend.Enabled = True
        CmdGenerate.Enabled = True
        CmdActivation.Caption = "Activate"
        FrameLicensed.Caption = "&Unlicensed Software!"
        FrameLicensed.ForeColor = &HFF&
        TimerLisensi.Enabled = True
        TimerLisensi.Interval = 1000
        Jam = 0
        Menit = 14
        Detik = 59
        StatusBarUtama.Panels(3).Text = Format(Jam, "00") & ":" & Format(Menit, "00") & ":" & Format(Detik, "00")
    ElseIf Val(TextCodeGetStatus.Text) + 5071997 = Val(TextCodeSendStatus.Text) Or Val(TextCodeGetStatus.Text) + 12111997 = Val(TextCodeSendStatus.Text) Or Val(TextCodeGetStatus.Text) + 7102011 = Val(TextCodeSendStatus.Text) Then
        TimerLisensi.Enabled = False
        TextCodeGet.Text = ""
        TextCodeSend.Text = ""
        TextUserLicense.Text = ""
        TextUserLicense.Enabled = False
        TextCodeGet.Enabled = False
        TextCodeSend.Enabled = False
        CmdGenerate.Enabled = False
        CmdActivation.Caption = "Deactivate"
        FrameLicensed.Caption = "&Licensed Software!"
        FrameLicensed.ForeColor = &HC000&
        StatusBarUtama.Panels(3).Text = ""
    End If
End If
End Sub

Private Sub CmdSetPassword_Click()
If TextCurrentPassword.Text = "" Or TextNewPassword.Text = "" Then
    MsgBox "Isi Password Dengan Benar", vbExclamation, "Isi Password Dengan Benar"
Else
    BukaDatabase
    RS.Open "SELECT password_petugas FROM tb_petugas WHERE id_petugas = '" & StatusBarUtama.Panels(1).Text & "'", Conn, adOpenDynamic, adLockOptimistic
    If TextCurrentPassword.Text = RS!password_petugas Then
        RS.Close
        If MsgBox("Ubah Password Sebelumnya ?", vbExclamation + vbYesNo, "Ubah Password") = vbYes Then
            BukaDatabase
            RS.Open "UPDATE tb_petugas SET password_petugas = '" & TextNewPassword.Text & "' WHERE id_petugas = '" & StatusBarUtama.Panels(1).Text & "'", Conn, adOpenDynamic, adLockOptimistic
            MsgBox "Password Berhasil Di Ubah", vbInformation, "Password Berhasil Di Ubah"
            Call CmdRefresh_Click
            FormLogin.Show
            Me.Hide
            Unload FormAbout
            Unload FormMahasiswa
            Unload FormPembayaran
            Unload FormReport
            Unload FormSearch
            Unload FormBuktiPembayaran
            MsgBox "Silahkan Login Untuk Memeriksa Password", vbExclamation, "Login Ulang"
        Else
            Exit Sub
        End If
    Else
        MsgBox "Password Sebelumnya Salah", vbExclamation, "Password Sebelumnya Salah"
    End If
End If
End Sub

Private Sub Exit_Click()
If MsgBox("Keluar Dari Aplikasi?", vbInformation + vbYesNo, "Keluar Dari Aplikasi") = vbYes Then
    End
Else
    Exit Sub
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
Unload FormAbout
Unload FormMahasiswa
Unload FormPembayaran
Unload FormReport
Unload FormSearch
Unload FormBuktiPembayaran
End Sub

Private Sub Form_Load()
DTPembanding.Value = Format(Now, "dd/mm/yyyy")
Call CmdRefresh_Click
Call CmdGenerate_Click
TimerUtama.Enabled = True
TimerUtama.Interval = 100

With ComboOtoritas
    .Clear
    .AddItem "PETUGAS"
    .AddItem "OPERATOR"
End With
Call CmdRefreshLisensi_Click

BukaDatabase
RS.Open "SELECT date_license FROM tb_lisensi", Conn, adOpenDynamic, adLockOptimistic
If DTPembanding.Value > RS!date_license Then
    BukaDatabase
    RS.Open "UPDATE tb_lisensi SET code_activation = 0, code_get = 0, user_license = 'Trial Version', date_license = #11/12/1997#", Conn, adOpenDynamic, adLockOptimistic
    MsgBox "Silahkan Perpanjang Lisensi", vbExclamation, "Perpanjang Lisensi"
    Call CmdRefreshLisensi_Click
    Call CmdGenerate_Click
Else
    Exit Sub
End If
End Sub

Private Sub ImagePetugasColor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImagePetugasColor.Visible = False
ImagePetugasBlack.Visible = True
End Sub

Private Sub ImagePetugasBlack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImagePetugasColor.Visible = True
ImagePetugasBlack.Visible = False
End Sub

Private Sub ImagePetugasColor_click()
FramePetugas.Visible = True
FrameLisensi.Visible = False
End Sub

Private Sub ImagePetugasBlack_click()
FramePetugas.Visible = True
FrameLisensi.Visible = False
End Sub

Private Sub ImageLisensiColor_click()
FramePetugas.Visible = False
FrameLisensi.Visible = True
End Sub

Private Sub ImageLisensiBlack_click()
FramePetugas.Visible = False
FrameLisensi.Visible = True
End Sub

Private Sub ImageLisensiBlack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImageLisensiColor.Visible = True
ImageLisensiBlack.Visible = False
End Sub

Private Sub ImageLisensiColor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImageLisensiColor.Visible = False
ImageLisensiBlack.Visible = True
End Sub

Private Sub Logout_Click()
If MsgBox("Logout Dari Aplikasi?", vbInformation + vbYesNo, "Logout Dari Aplikasi") = vbYes Then
    FormLogin.Show
    Me.Hide
    Unload FormAbout
    Unload FormMahasiswa
    Unload FormPembayaran
    Unload FormReport
    Unload FormSearch
    Unload FormBuktiPembayaran
Else
    Exit Sub
End If
End Sub

Private Sub Mahasiswa_Click()
FormMahasiswa.Show
End Sub

Private Sub Pembayaran_Click()
FormPembayaran.Show
End Sub

Private Sub Registration_Click()
FramePetugas.Visible = False
FrameLisensi.Visible = True
End Sub

Private Sub Report_Click()
FormReport.Show
End Sub

Private Sub Search_Click()
FormSearch.Show
End Sub

Private Sub TimerLisensi_Timer()
Detik = Detik - 1
If Detik < 0 Then
Detik = 59
Menit = Menit - 1
If Menit < 0 Then
Menit = 59
Jam = Jam - 1
End If
End If
StatusBarUtama.Panels(3).Text = Format(Jam, "00") & ":" & Format(Menit, "00") & ":" & Format(Detik, "00")

If Jam = 0 And Menit = 0 And Detik = 0 Then
    TimerLisensi.Enabled = False
    MsgBox "Unlicensed Software!", vbExclamation, "Unlicensed Software!"
    Unload Me
    End
End If
End Sub

Private Sub TimerUtama_Timer()
StatusBarUtama.Panels(4).Text = Format(Now, "dd/MM/yyyy")
End Sub

Private Sub PengaturanDataGridPetugas()
With DataGridPetugas
    .Columns(0).Caption = "USERNAME"
    .Columns(1).Caption = "OTORITAS"
End With
End Sub

Private Sub User_Click()
FramePetugas.Visible = True
FrameLisensi.Visible = False
End Sub
