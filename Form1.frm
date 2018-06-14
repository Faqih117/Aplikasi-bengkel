VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "PENJUALAN DAN SERVICE"
   ClientHeight    =   9045
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9045
   ScaleWidth      =   13440
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Item 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10080
      TabIndex        =   44
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   360
      Top             =   1800
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   7200
      Top             =   8400
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\semester 2\vb\Kelompok 1 (Bengkel)\bengkel.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\semester 2\vb\Kelompok 1 (Bengkel)\bengkel.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Jual"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "Simpan"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   41
      Top             =   8400
      Width           =   1335
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "Tutup"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   40
      Top             =   8400
      Width           =   1335
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "Batal"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   39
      Top             =   8400
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cetak"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   38
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Masukkan No Plat Kendaraan"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9000
      TabIndex        =   36
      Top             =   7200
      Width           =   3975
      Begin VB.TextBox Text14 
         Height          =   375
         Left            =   360
         TabIndex        =   37
         Top             =   480
         Width           =   3135
      End
   End
   Begin MSDataGridLib.DataGrid DTGrid 
      Bindings        =   "Form1.frx":0000
      Height          =   1455
      Left            =   240
      TabIndex        =   35
      Top             =   6720
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   2566
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   20
      FormatLocked    =   -1  'True
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
         Name            =   "Tw Cen MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Detail Barang"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "No_Transaksi"
         Caption         =   "No_Transaksi"
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
         DataField       =   "No_Plat"
         Caption         =   "No_Plat"
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
      BeginProperty Column02 
         DataField       =   "Kode_Barang"
         Caption         =   "Kode_Barang"
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
      BeginProperty Column03 
         DataField       =   "Jumlah"
         Caption         =   "Jumlah"
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
      BeginProperty Column04 
         DataField       =   "Sub_Total"
         Caption         =   "Sub_Total"
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
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1094,74
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Kembali 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10080
      TabIndex        =   34
      Top             =   6360
      Width           =   2175
   End
   Begin VB.TextBox Dibayar 
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10080
      TabIndex        =   32
      Top             =   5640
      Width           =   2175
   End
   Begin VB.TextBox Total 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10080
      TabIndex        =   31
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Frame Frame2 
      Caption         =   "Servis"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   240
      TabIndex        =   24
      Top             =   4560
      Width           =   8295
      Begin VB.TextBox Text10 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   28
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox Text9 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   26
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label10 
         Caption         =   "Biaya Pemasangan"
         BeginProperty Font 
            Name            =   "Tw Cen MT"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   27
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label9 
         Caption         =   "Biaya Servis"
         BeginProperty Font 
            Name            =   "Tw Cen MT"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   25
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10440
      TabIndex        =   22
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Retur"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   21
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   20
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox plat 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   19
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox Nmp 
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   18
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton CmdTambah 
      Caption         =   "Tambahkan"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   17
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Penjualan"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   8295
      Begin VB.CommandButton Command4 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   23
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   12
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         TabIndex        =   11
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   10
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Jumlah"
         BeginProperty Font 
            Name            =   "Tw Cen MT"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6600
         TabIndex        =   16
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Harga"
         BeginProperty Font 
            Name            =   "Tw Cen MT"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         TabIndex        =   15
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Nama Barang"
         BeginProperty Font 
            Name            =   "Tw Cen MT"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   14
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Kode Barang"
         BeginProperty Font 
            Name            =   "Tw Cen MT"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Pemasangan"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Servis"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   960
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Penjualan"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox No_transaksi 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Tanggal 
      Height          =   495
      Left            =   11040
      TabIndex        =   46
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label14 
      Caption         =   "Item"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      TabIndex        =   45
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label jam 
      Height          =   495
      Left            =   11160
      TabIndex        =   43
      Top             =   3360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label namaksr 
      Height          =   495
      Left            =   11160
      TabIndex        =   42
      Top             =   2520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label13 
      Caption         =   "Kembalian"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      TabIndex        =   33
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Label Label12 
      Caption         =   "Bayar"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      TabIndex        =   30
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      TabIndex        =   29
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Mekanik"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   7
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "No Plat"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   6
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Nama Pelanggan"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "No Transaksi"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Nmp_KeyPress(KeyAscii As Integer)
'ubah karakter jadi besar semua
KeyAscii = Asc(UCase(Chr(KeyAscii)))
'jika menekan ESC form ditutup
If KeyAscii = 13 Then
RSCustomer.Open "Select * from Customer where Nama_customer= '" & Nmp.Text & "'", Koneksi
    If Not RSCustomer.EOF Then
        plat = RSCustomer!No_plat
    Else
        MsgBox "Data Pelanggan Tidak Ada", vbInformation, "Informasi"
    End If
End If

End Sub

Private Sub Timer1_Timer()
    jam = Time$
End Sub
 
Private Sub Form_Activate()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\Bengkel.mdb"
Adodc1.RecordSource = "Jual"
Set DTGrid.DataSource = Adodc1
DTGrid.Refresh

Call BukaDB
RSUser.Open "Login", Koneksi
namaksr = RSUser!Kode_user

Call Auto
Tanggal = Date

CmdSimpan.Enabled = False
End Sub

 
Private Sub Auto()
Call BukaDB
RSPenjualan.Open "select * from Penjualan_dua Where No_transaksi(Select Max(No_transaksi)From Penjualan_dua)Order By No_transaksi Desc", Koneksi
RSPenjualan.Requery
    Dim Urutan As String * 10
    Dim Hitung As Long
    With RSPenjualan
        If .EOF Then
            Urutan = Right(Date, 2) + Mid(Date, 4, 2) + Left(Date, 2) + "0001"
            No_transaksi = Urutan
        Else
            If Left(!No_transaksi, 6) <> Right(Date, 2) + Mid(Date, 4, 2) + Left(Date, 2) Then
                Urutan = Right(Date, 2) + Mid(Date, 4, 2) + Left(Date, 2) + "0001"
            Else
                Hitung = (!No_transaksi) + 1
                Urutan = (Right(Date, 2) + Mid(Date, 4, 2) + Left(Date, 2)) + Right("0000" & Hitung, 4)
            End If
        End If
        No_transaksi = Urutan
    End With
End Sub
 


Private Sub Command2_Click()
        Ado.Recordset!No_transaksi = Null
        Ado.Recordset!No_plat = Null
        Adodc1.Recordset!Kode_Barang = Null
        Ado.Recordset!jumlah = Null
        Ado.Recordset!Sub_Total = Null
        Ado.Recordset.Update
        Call TotalItem
        Call TotalHarga
        DTGrid.Refresh
End Sub
 
Private Sub DTGrid_AfterColEdit(ByVal ColIndex As Integer)
 If DTGrid.Col = 1 Then
        If Len(Adodc1.Recordset!Kode) < 5 Then
            MsgBox "Kode Harus 5 digit"
            DTGrid.Col = 1
            Exit Sub
        End If
    
        Call BukaDB
        RSBarang.Open "Select * from Barang where Kode_Barang='" & Adodc1.Recordset!Kode & "'", Koneksi
        If Not RSBarang.EOF Then
        Adodc1.Recordset!Kode = RSBarang!Kode_Barang
        Adodc1.Recordset!Nama_Barang = RSBarang!Nama_Barang
        Adodc1.Recordset!Harga = RSBarang!Harga_Barang
        DTGrid.Col = 4
        DTGrid.Refresh
            Exit Sub
        End If
    End If
    If DTGrid.Col = 4 Then
        If Adodc1.Recordset!jumlah > RSBarang!Stok Then
            MsgBox "STOK BARANG KURANG", , "Informasi"
            Exit Sub
        Else
            Adodc1.Recordset!jumlah = Adodc1.Recordset!jumlah
            Adodc1.Recordset!Total = Adodc1.Recordset!Harga * Adodc1.Recordset!jumlah
            Adodc1.Recordset.Update
            Call Tambah_Baris
            Adodc1.Recordset.MoveNext
            DTGrid.Col = 1
            Adodc1.Recordset.MoveLast
            Call TotalHarga
            Call TotalItem
        End If
    End If
End Sub
 
Function TotalItem()
On Error Resume Next
Adodc1.Recordset.MoveFirst
Item = 0
Do While Not Adodc1.Recordset.EOF And Adodc1.Recordset!jumlah <> 0
    Item = Item + Adodc1.Recordset!jumlah
    Adodc1.Recordset.MoveNext
    Item = Item
Loop
End Function
 
Function TotalHarga()
On Error Resume Next
Adodc1.Recordset.MoveFirst
Total = 0
Do While Not Adodc1.Recordset.EOF And Adodc1.Recordset!Total <> 0
    Total = Total + Adodc1.Recordset!Total
    Adodc1.Recordset.MoveNext
    Total = Format(Total, "#,###,###")
Loop
End Function
 
Private Sub bersihkan()
    Item = ""
    Total = ""
    Dibayar = ""
    Kembali = ""
    namacus = ""
End Sub
 
Private Sub Dibayar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Dibayar = "" Or Val(Dibayar) < (Total) Then
            MsgBox "Jumlah Pembayaran Kurang", , "Informasi"
            Dibayar.SetFocus
        Else
            Dibayar = Format(Dibayar, "Rp ###,###,###")
            If Dibayar = Total Then
                Kembali = Dibayar - Total
            Else
                Kembali = Format(Dibayar - Total, "Rp ###,###,###")
            End If
        CmdSimpan.Enabled = True
        CmdSimpan.SetFocus
        End If
    End If
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub
 
Private Sub CmdSimpan_Keypress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        CmdSimpan.Enabled = False
        Dibayar = ""
        Dibayar.SetFocus
    End If
End Sub
 
Private Sub CmdSimpan_Click()
    Dim SQLTambahJual As String
    SQLTambahJual = "Insert Into Penjualan(Faktur,tanggal,jam,Total,Item,dibayar,kembali,Kode_customer,Kode_user)" & _
    "values('" & Faktur & "','" & Tanggal & "','" & jam & "','" & Total & "','" & Item & "','" & Dibayar & "','" & Kembali & "','" & Combo1.Text & "','" & Kode_user & "')"
    Koneksi.Execute (SQLTambahJual)
        
    Adodc1.Recordset.MoveFirst
    Do While Not Adodc1.Recordset.EOF
        If Adodc1.Recordset!Kode <> vbNullString Then
            Dim SQLTambahDetail As String
            SQLTambahDetail = "Insert Into Detailjual(Faktur,Kode_Barang,Nama_Barang,Harga,JmlJual,Total) " & _
            "values ('" & Faktur & "','" & Adodc1.Recordset!Kode & "','" & Adodc1.Recordset!Nama_Barang & "','" & Adodc1.Recordset!Harga & "','" & Adodc1.Recordset!jumlah & "','" & Adodc1.Recordset!Total & "')"
            Koneksi.Execute (SQLTambahDetail)
        End If
    Adodc1.Recordset.MoveNext
    Loop
       
    Adodc1.Recordset.MoveFirst
    Do While Not Adodc1.Recordset.EOF
        If Adodc1.Recordset!Kode <> vbNullString Then
            Call BukaDB
            RSBarang.Open "Select * from Barang where Kode_Barang='" & Adodc1.Recordset!Kode & "'", Koneksi
            If Not RSBarang.EOF Then
                Dim Kurangi As String
                Kurangi = "update Barang set Stok='" & RSBarang!Stok - Adodc1.Recordset!jumlah & "' where Kode_Barang='" & Adodc1.Recordset!Kode & "'"
                Koneksi.Execute (Kurangi)
            Else
                MsgBox "Stok Barang Sudah Habis", vbInformation, "Info"
            
            End If
        End If
    Adodc1.Recordset.MoveNext
    Loop
    bersihkan
    Form_Activate
    If MsgBox("Apakah Ingin Dicetak layar?", vbYesNo, "CETAK") = vbYes Then
       Call Cetak
    ElseIf vbNo Then
        MsgBox "Hanya Disimpan", vbInformation, "INFO"
    End If
End Sub
 Sub tampilreport(namafile As String, selectionFormula As String, Cetak As Integer)
On Error GoTo 1

    With Me
        .crt1.Reset
        setCRT .crt1
       
        .crt1.ReportFileName = namafile
        .crt1.selectionFormula = selectionFormula
        .crt1.PageZoom 2
        .crt1.WindowState = crptMaximized
        .crt1.WindowShowGroupTree = False
        .crt1.RetrieveDataFiles
        .crt1.Destination = crptToPrinter
        .crt1.PrinterCopies = Cetak
        
        .crt1.Action = 1
        
    End With

1:
If Err.Number = 20515 Then
MsgBox Err.Number & vbNewLine & "Ubah Regional Setting ke English (United States)", vbExclamation, "KESALAHAN CETAK"
Exit Sub
Else
If Not Err.Number = 0 Then MsgBox Err.Number & vbNewLine & Err.Description, vbExclamation

End If
End Sub

Private Sub CmdBatal_Click()
    Dibayar = ""
    Total = ""
    Item = ""
    Kembali = ""
    namacus = ""
    Form_Activate
End Sub
 
Private Sub Cmadodc1utup_Click()
    Unload Me
End Sub
Function cetakrpt()
    'panggil laporan yang tanggalnya = combo1
    CR.selectionFormula = "Totext({Penjualan.Faktur})='" & Faktur & "'"
    CR.ReportFileName = App.Path & "\LapFakturJual1.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Function
Function Cetak()
Call BukaDB
RSPenjualan.Open "select * from penjualan Where Faktur In(Select Max(Faktur)From penjualan)Order By Faktur Desc", Koneksi
Layar.Show
Dim Total, JmlJual, JmlHasil As Double
Dim MGrs As String
Layar.Font = "Courier New"
Layar.Print
Layar.Print
Layar.Print Tab(5); "Faktur     :   "; RSPenjualan!Faktur
Layar.Print Tab(5); "Tanggal    :   "; Format(RSPenjualan!Tanggal, "DD-MMMM-YYYY")
Layar.Print Tab(5); "Jam        :   "; Format(RSPenjualan!jam, "HH:MM:SS")
MGrs = String$(33, "-")
Layar.Print Tab(5); MGrs
RSDetailJual.Open "select * from detailjual Where left(Faktur,10)='" & RSPenjualan!Faktur & "'", Koneksi
RSDetailJual.MoveFirst
No = 0
Do While Not RSDetailJual.EOF
    No = No + 1
    Set RSBarang = New ADODB.Recordset
    RSBarang.Open "select * From Barang where Kode_Barang= '" & RSDetailJual!Kode_Barang & "'", Koneksi
    RSBarang.Requery
    Harga = RSBarang!Harga_Barang
    jumlah = RSDetailJual!JmlJual
    Hasil = Harga * jumlah
    Layar.Print Tab(5); No; Space(2); RSBarang!Nama_Barang
    Layar.Print Tab(10); RKanan(jumlah, "##"); Space(1); "X";
    Layar.Print Tab(15); Format(Harga, "###,###,###");
    Layar.Print Tab(25); RKanan(Hasil, "###,###,###")
    RSDetailJual.MoveNext
Loop
Layar.Print Tab(5); MGrs
Layar.Print Tab(5); "Total      :";
Layar.Print Tab(25); RKanan(RSPenjualan!Total, "###,###,###");
Layar.Print Tab(5); "Dibayar    :";
Layar.Print Tab(25); RKanan(RSPenjualan!Dibayar, "###,###,###");
Layar.Print Tab(5); MGrs
Layar.Print Tab(5); "Kembali    :";
If RSPenjualan!Dibayar = RSPenjualan!Total Then
    Layar.Print Tab(34); RSPenjualan!Dibayar - RSPenjualan!Total
Else
    Layar.Print Tab(25); RKanan(RSPenjualan!Dibayar - RSPenjualan!Total, "###,###,###");
End If
Layar.Print Tab(5); MGrs
Layar.Print Tab(5); "Terima Kasih atas kunjungan Anda"
Layar.Print
Layar.Print
Layar.Print
Koneksi.Close
End Function
 
Private Function RKanan(NData, CFormat) As String
    RKanan = Format(NData, CFormat)
    RKanan = Space(Len(CFormat) - Len(RKanan)) + RKanan
End Function
 
Private Sub List1_keyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If DTGrid.SelText <> Left(List1, 5) Then
            DTGrid.SelText = Left(List1, 5)
            Adodc1.Recordset.Update
            Call BukaDB
            RSBarang.Open "Select * from Barang where Kode_Barang ='" & Left(List1, 5) & "'", Koneksi, adOpenDynamic, adLockOptimistic
            RSBarang.Requery
            If Not RSBarang.EOF Then
                Adodc1.Recordset!Kode = RSBarang!Kode_Barang
                Adodc1.Recordset!Nama_Barang = RSBarang!Nama_Barang
                Adodc1.Recordset!Harga = RSBarang!Harga_Barang
                Adodc1.Recordset.Update
                DTGrid.SetFocus
                DTGrid.Col = 4
            End If
        End If
    End If
End Sub
 
Private Sub CmdTutup_Click()
Unload Me
End Sub




