VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Transaksi 
   BackColor       =   &H8000000E&
   Caption         =   "Transaksi Servis"
   ClientHeight    =   9195
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11535
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   ScaleHeight     =   9195
   ScaleWidth      =   11535
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check3 
      BackColor       =   &H8000000E&
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
      Left            =   7800
      TabIndex        =   36
      Top             =   1320
      Width           =   1815
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
      Left            =   8400
      TabIndex        =   35
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000E&
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
      Left            =   7440
      TabIndex        =   33
      Top             =   7080
      Width           =   3135
      Begin VB.TextBox Text14 
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
         TabIndex        =   34
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000E&
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
      Height          =   1215
      Left            =   480
      TabIndex        =   28
      Top             =   1440
      Width           =   6855
      Begin VB.TextBox Biaya 
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
         Height          =   375
         Left            =   3480
         TabIndex        =   30
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox servis 
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
         Height          =   375
         Left            =   720
         TabIndex        =   29
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000E&
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
         Left            =   3480
         TabIndex        =   32
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
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
         Left            =   720
         TabIndex        =   31
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H8000000E&
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
      Left            =   7800
      TabIndex        =   27
      Top             =   360
      Width           =   1815
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H8000000E&
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
      Left            =   7800
      TabIndex        =   26
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox Montir 
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
      Left            =   2040
      TabIndex        =   25
      Top             =   840
      Width           =   1455
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
      Height          =   375
      Left            =   5400
      TabIndex        =   23
      Top             =   840
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Caption         =   "Proses"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   960
      TabIndex        =   19
      Top             =   7080
      Width           =   2655
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
         Height          =   375
         Left            =   600
         TabIndex        =   22
         Top             =   840
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
         Height          =   375
         Left            =   1440
         TabIndex        =   21
         Top             =   360
         Width           =   975
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
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4140
      Left            =   7560
      TabIndex        =   13
      Top             =   2760
      Width           =   3015
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   840
   End
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
      Height          =   375
      Left            =   5760
      TabIndex        =   8
      Top             =   7200
      Width           =   1455
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
      Height          =   375
      Left            =   5760
      TabIndex        =   7
      Top             =   7680
      Width           =   1455
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
      Height          =   375
      Left            =   5760
      TabIndex        =   6
      Top             =   8160
      Width           =   1455
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
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   8640
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DTGrid 
      Bindings        =   "Transaksi.frx":0000
      Height          =   4215
      Left            =   480
      TabIndex        =   4
      Top             =   2760
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   7435
      _Version        =   393216
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "No"
         Caption         =   "No"
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
         DataField       =   "Kode"
         Caption         =   "Kode"
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
         DataField       =   "Nama_Barang"
         Caption         =   "Nama_Barang"
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
         DataField       =   "Harga"
         Caption         =   "Harga"
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
      BeginProperty Column05 
         DataField       =   "Total"
         Caption         =   "Total"
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
      BeginProperty Column06 
         DataField       =   "Biaya_pemasangan"
         Caption         =   "Biaya_pemasangan"
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
            ColumnWidth     =   599,811
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2085,166
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   599,811
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column06 
            Object.Visible         =   0   'False
            ColumnWidth     =   1814,74
         EndProperty
      EndProperty
   End
   Begin VB.TextBox namacus 
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
      Left            =   5400
      TabIndex        =   3
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox Faktur 
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
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   6600
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      RecordSource    =   "Transaksi"
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
   Begin VB.TextBox Kode_user 
      Height          =   375
      Left            =   4680
      TabIndex        =   18
      Top             =   3840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin Crystal.CrystalReport CR 
      Left            =   600
      Top             =   6600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox Tanggal 
      Enabled         =   0   'False
      Height          =   375
      Left            =   8760
      TabIndex        =   16
      Top             =   3240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No plat"
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
      Left            =   3600
      TabIndex        =   24
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Jam 
      Height          =   375
      Left            =   1440
      TabIndex        =   17
      Top             =   4080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Faktur"
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
      Left            =   480
      TabIndex        =   15
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Pilih Data Barang Dalam ListBox"
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
      Left            =   7560
      TabIndex        =   14
      Top             =   2400
      Width           =   2895
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3960
      TabIndex        =   12
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Harga"
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
      Left            =   3960
      TabIndex        =   11
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3960
      TabIndex        =   10
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3960
      TabIndex        =   9
      Top             =   8640
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nama Customer"
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
      Left            =   3600
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "Transaksi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
            Call BukaDB
            RSBarang.Open "Barang", Koneksi
            List1.Clear
            Do Until RSBarang.EOF
                List1.AddItem RSBarang!Kode_Barang & Space(2) & RSBarang!Nama_Barang
                RSBarang.MoveNext
            Loop
Else
    List1.Clear
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    Form2.Show
Else
    servis.Text = ""
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
    Call Pasang
    Total = Val(Total) + Val(Biaya)
Else
    Total = Val(Total) - Val(Biaya)
    Biaya.Text = ""
End If
End Sub

Private Sub Command5_Click()
    'buka database
    Call BukaDB
    'cari data yang tanggal dan bulannya dipilih di combo4 dan combo5
    RSPenjualan.Open "select * from Penjualan where No_plat='" & Text14 & "'", Koneksi
    'jika tidak cocok, munculkan pesan
    If RSPenjualan.EOF Then
        MsgBox "Data tidak ditemukan"
        Exit Sub
        Command5.SetFocus
    End If
    'jika ditemukan panggil file laporan yang
    CR.selectionFormula = "{Penjualan.No_Plat}='" & Text14 & "'"
    CR.ReportFileName = App.Path & "\LapFakturJual.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub Montir_KeyPress(KeyAscii As Integer)
'ubah karakter jadi besar semua
KeyAscii = Asc(UCase(Chr(KeyAscii)))
'jika menekan ESC form ditutup
If KeyAscii = 13 Then
RSMekanik.Open "Select * from Mekanik where Nm_mekanik= '" & Montir.Text & "'", Koneksi
    If Not RSMekanik.EOF Then
        namacus.SetFocus
    Else
        MsgBox "Data Mekanik Tidak Ada", vbInformation, "Informasi"
        Mekanik.Show
    End If
End If
End Sub

Private Sub namacus_KeyPress(KeyAscii As Integer)
'ubah karakter jadi besar semua
KeyAscii = Asc(UCase(Chr(KeyAscii)))
'jika menekan ESC form ditutup
If KeyAscii = 13 Then
RSCustomer.Open "Select * from Customer where Nama_customer= '" & namacus.Text & "'", Koneksi
    If Not RSCustomer.EOF Then
        plat = RSCustomer!No_plat
    Else
        MsgBox "Data Pelanggan Tidak Ada", vbInformation, "Informasi"
        Pelanggan.Show
    End If
End If
End Sub
Private Sub Timer1_Timer()
    jam = Time$
End Sub
 
Private Sub Form_Activate()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\Bengkel.mdb"
Adodc1.RecordSource = "Transaksi"
Set DTGrid.DataSource = Adodc1
DTGrid.Refresh
Call BukaDB
Kode_user = MenuUtama.StatusBar1.Panels(2)


Call Auto
Call Tabel_Kosong
Tanggal = Date
Adodc1.Recordset.MoveFirst
CmdSimpan.Enabled = False
End Sub

 
Private Sub Auto()
Call BukaDB
RSPenjualan.Open "select * from Penjualan Where Faktur In(Select Max(Faktur)From Penjualan)Order By Faktur Desc", Koneksi
RSPenjualan.Requery
    Dim Urutan As String * 10
    Dim Hitung As Long
    With RSPenjualan
        If .EOF Then
            Urutan = Right(Date, 2) + Mid(Date, 4, 2) + Left(Date, 2) + "0001"
            Faktur = Urutan
        Else
            If Left(!Faktur, 6) <> Right(Date, 2) + Mid(Date, 4, 2) + Left(Date, 2) Then
                Urutan = Right(Date, 2) + Mid(Date, 4, 2) + Left(Date, 2) + "0001"
            Else
                Hitung = (!Faktur) + 1
                Urutan = (Right(Date, 2) + Mid(Date, 4, 2) + Left(Date, 2)) + Right("0000" & Hitung, 4)
            End If
        End If
        Faktur = Urutan
    End With
End Sub
 
Function Tabel_Kosong()
        Adodc1.Recordset.MoveFirst
    Do While Not Adodc1.Recordset.EOF
        Adodc1.Recordset.Delete
        Adodc1.Recordset.MoveNext
    Loop
    For i = 1 To 1
        Adodc1.Recordset.AddNew
        Adodc1.Recordset!No = i
        Adodc1.Recordset.Update
    Next i
    DTGrid.Col = 1
End Function
Function Tambah_Baris()
    For i = Adodc1.Recordset.RecordCount To Adodc1.Recordset.RecordCount
        Adodc1.Recordset.AddNew
        Adodc1.Recordset!No = i + 1
        Adodc1.Recordset.Update
    Next i
End Function
 
Private Sub DTGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyEscape
        Adodc1.Recordset!Kode = Null
        Adodc1.Recordset!Nama_Barang = Null
        Adodc1.Recordset!Harga = Null
        Adodc1.Recordset!jumlah = Null
        Adodc1.Recordset!Total = Null
        Adodc1.Recordset!Biaya_pemasangan = Null
        Adodc1.Recordset.Update
        Call TotalItem
        Call TotalHarga
        DTGrid.Refresh
End Select
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
        Adodc1.Recordset!Biaya_pemasangan = RSBarang!Biaya_pemasangan
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
Loop
    Total = Total + Val(servis)
End Function
Function Pasang()
On Error Resume Next
Adodc1.Recordset.MoveFirst
Biaya = 0
Do While Not Adodc1.Recordset.EOF And Adodc1.Recordset!Biaya_pemasangan <> 0
    Biaya = Biaya + Adodc1.Recordset!Biaya_pemasangan
    Adodc1.Recordset.MoveNext
    Biaya = Biaya
Loop
End Function
Private Sub bersihkan()
    Dibayar = ""
    Total = ""
    Item = ""
    Kembali = ""
    namacus = ""
    plat = ""
    Montir = ""
    Biaya = ""
    servis = ""
    Check1 = 0
    Check2 = 0
    Check3 = 0
    Form_Activate
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
    Dim TambahJual As String
    TambahJual = "Insert Into Penjualan(Faktur,tanggal,jam,Biaya_servis,Biaya_pemasangan,Total,Item,dibayar,kembali,Nama_customer,Kode_user,Nm_mekanik,No_plat)" & _
    "values('" & Faktur & "','" & Tanggal & "','" & jam & "','" & servis & "','" & Biaya & "','" & Total & "','" & Item & "','" & Dibayar & "','" & Kembali & "','" & namacus & "','" & Kode_user & "','" & Montir & "','" & plat & "')"
    Koneksi.Execute (TambahJual)
        
    Adodc1.Recordset.MoveFirst
    Do While Not Adodc1.Recordset.EOF
        If Adodc1.Recordset!Kode <> vbNullString Then
            Dim TambahDetail As String
            TambahDetail = "Insert Into Detailjual(Faktur,Kode_Barang,Nama_Barang,Harga,JmlJual,Total) " & _
            "values ('" & Faktur & "','" & Adodc1.Recordset!Kode & "','" & Adodc1.Recordset!Nama_Barang & "','" & Adodc1.Recordset!Harga & "','" & Adodc1.Recordset!jumlah & "','" & Adodc1.Recordset!Total & "')"
            Koneksi.Execute (TambahDetail)
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
    If MsgBox("Apakah Ingin Dicetak?", vbYesNo, "CETAK") = vbYes Then
        Call cetakrpt
    ElseIf vbNo Then
        MsgBox "Hanya Disimpan", vbInformation, "INFO"
    End If
End Sub
 

Private Sub CmdBatal_Click()
    Dibayar = ""
    Total = ""
    Item = ""
    Kembali = ""
    namacus = ""
    plat = ""
    Montir = ""
    Biaya = ""
    servis = ""
    Check1 = 0
    Check2 = 0
    Check3 = 0
    Form_Activate
End Sub
 
Private Sub Cmadodc1utup_Click()
    Unload Me
End Sub
Function cetakrpt()
    CR.selectionFormula = "{Penjualan.Faktur}='" & RSPenjualan!Faktur & "'"
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
            RSBarang.Open "Select * from Barang where Kode_Barang ='" & Left(List1, 5) & "'", Koneksi
            RSBarang.Requery
            If Not RSBarang.EOF Then
                Adodc1.Recordset!Kode = RSBarang!Kode_Barang
                Adodc1.Recordset!Nama_Barang = RSBarang!Nama_Barang
                Adodc1.Recordset!Harga = RSBarang!Harga_Barang
                Adodc1.Recordset!Biaya_pemasangan = RSBarang!Biaya_pemasangan
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
