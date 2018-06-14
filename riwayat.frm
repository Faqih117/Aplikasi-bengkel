VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form riwayat 
   BackColor       =   &H8000000E&
   Caption         =   "Rincian Penjualan"
   ClientHeight    =   7485
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   11070
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "riwayat.frx":0000
      Height          =   4335
      Left            =   3360
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   7646
      _Version        =   393216
      AllowUpdate     =   0   'False
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "Faktur"
         Caption         =   "Faktur"
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
         Caption         =   "Harga_Barang"
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
         DataField       =   "JmlJual"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   -1  'True
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Caption         =   "Kasir"
      Height          =   1815
      Left            =   240
      TabIndex        =   16
      Top             =   4320
      Width           =   2775
      Begin VB.TextBox namaksr 
         Enabled         =   0   'False
         Height          =   405
         Left            =   1080
         TabIndex        =   19
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Kodeksr 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1080
         TabIndex        =   18
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox tanggal 
         Enabled         =   0   'False
         Height          =   405
         Left            =   1080
         TabIndex        =   17
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nama"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Kode"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tanggal"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   1320
         Width           =   855
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   10920
      Top             =   4200
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
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
      RecordSource    =   "DetailJual"
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
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1320
      TabIndex        =   15
      Text            =   "Combo1"
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "Tutup"
      Height          =   375
      Left            =   4920
      TabIndex        =   12
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hapus"
      Height          =   375
      Left            =   3360
      TabIndex        =   11
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox total 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   10
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox jumlah 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   6360
      Width           =   735
   End
   Begin VB.TextBox telepon 
      Enabled         =   0   'False
      Height          =   405
      Left            =   1320
      TabIndex        =   8
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox alamat 
      Enabled         =   0   'False
      Height          =   405
      Left            =   1320
      TabIndex        =   7
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox namacus 
      Enabled         =   0   'False
      Height          =   405
      Left            =   1320
      TabIndex        =   4
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox kodecst 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   2400
      Picture         =   "riwayat.frx":0015
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label Label10 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total harga"
      Height          =   375
      Left            =   240
      TabIndex        =   24
      Top             =   6720
      Width           =   975
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Junlah beli"
      Height          =   375
      Left            =   240
      TabIndex        =   23
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Faktur"
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Kode_user 
      Height          =   375
      Left            =   8880
      TabIndex        =   13
      Top             =   4800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Telepon"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Alamat"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label label2 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nama"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kode PL"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "BERKAH MOTOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   26
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "KELOMPOK 1 VB 6.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   25
      Top             =   960
      Width           =   3975
   End
End
Attribute VB_Name = "riwayat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdTutup_Click()
Unload Me
End Sub

Private Sub Combo1_Click()
Call pencarian
Call BukaDB
RSPenjualan.Open "select * from Penjualan where Faktur='" & Combo1 & "'", Koneksi
If Not RSPenjualan.EOF Then
    namacus = RSPenjualan!Nama_customer
    Tanggal = RSPenjualan!Tanggal
    Kodeksr = RSPenjualan!Kode_user
    jumlah = RSPenjualan!Item
    Total = RSPenjualan!Total
End If

Call BukaDB
RSCustomer.Open "select * from Customer where Nama_customer='" & namacus & "'", Koneksi
If Not RSCustomer.EOF Then
    kodecst = RSCustomer!Kode_customer
    namacus = RSCustomer!Nama_customer
    alamat = RSCustomer!alamat
    telepon = RSCustomer!telepon
End If

Call BukaDB
RSUser.Open "select * from Login where Kode_user='" & Kodeksr & "'", Koneksi
If Not RSUser.EOF Then
    Kodeksr = RSUser!Kode_user
    namaksr = RSUser!Nama_user
End If


End Sub
Function pencarian()
Call BukaDB
Koneksi.CursorLocation = adUseClient
RSDetailJual.Open "select * from DetailJual where Faktur like '%" & Combo1 & "%'", Koneksi
    With RSDetailJual
        If Not (.BOF And .EOF) Then
            mvBookmark = .Bookmark
        End If
    End With
Set DataGrid1.DataSource = RSDetailJual.DataSource
DataGrid1.Refresh
DataGrid1.Visible = True
End Function
Private Sub Command1_Click()
Call BukaDB
    If Combo1 = "" Then
        MsgBox "Data Harus Dipilih Terlebih Dahulu", vbInformation, "Info"
            ElseIf MsgBox("Anda Yakin Ingin Menghapus", vbYesNo, "Info") = vbYes Then
                Koneksi.Execute "delete from DetailJual where Faktur='" & Combo1 & "'"
                Koneksi.Execute "delete from Penjualan where Faktur='" & Combo1 & "'"
                Adodc1.Recordset.Update
                MsgBox "Data Berhasil di Hapus", vbInformation, "Info"
            ElseIf vbNo Then
           MsgBox "DIBATALKAN !", vbCritical, "Keluar"
    End If
Call bersihkan
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub Form_Load()
Call isi_combo1
End Sub

Sub isi_combo1()
Call BukaDB
     Combo1.Clear
     Combo1.Refresh
       RSPenjualan.Open "Select * from Penjualan where Faktur", Koneksi
    Do While Not RSPenjualan.EOF
        Combo1.AddItem RSPenjualan!Faktur
        RSPenjualan.MoveNext
    Loop
    RSPenjualan.Close
End Sub

Sub bersihkan()
Combo1 = ""
kodecst = ""
namacus = ""
alamat = ""
telepon = ""
Kodeksr = ""
namaksr = ""
Tanggal = ""
jumlah = ""
Total = ""
DataGrid1.Visible = False
End Sub


