VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Mekanik 
   BackColor       =   &H8000000E&
   Caption         =   "Mekanik"
   ClientHeight    =   7155
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6405
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   7155
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Frame Frame2 
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
      Height          =   2415
      Left            =   4560
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
      Begin VB.CommandButton Command1 
         Caption         =   "Simpan"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Hapus"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Ubah"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Tutup"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1680
         Width           =   975
      End
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   3120
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   3600
      Width           =   2175
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Mekanik.frx":0000
      Height          =   2175
      Left            =   360
      TabIndex        =   10
      Top             =   4200
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   3836
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "ID_mekanik"
         Caption         =   "ID_mekanik"
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
         DataField       =   "Nm_mekanik"
         Caption         =   "Nama Mekanik"
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
         DataField       =   "Alamat_mekanik"
         Caption         =   "Alamat mekanik"
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
         DataField       =   "Telpon"
         Caption         =   "Telpon"
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
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3960
      Top             =   5280
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
      RecordSource    =   "Mekanik"
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
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Id Mekanik"
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nama"
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Alamat"
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No Telepon"
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cari Data"
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   960
      Picture         =   "Mekanik.frx":0015
      Top             =   0
      Width           =   1500
   End
   Begin VB.Label Label6 
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
      Left            =   2160
      TabIndex        =   17
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label Label7 
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
      Left            =   2160
      TabIndex        =   16
      Top             =   840
      Width           =   3975
   End
End
Attribute VB_Name = "Mekanik"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call BukaDB
    RSMekanik.Open "Select * from Mekanik", Koneksi
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text5 = "" Then
MsgBox "Data Belum Lengkap", vbInformation, "Info"
Exit Sub
Else
Adodc1.Recordset.AddNew
Adodc1.Recordset!ID_mekanik = Text1.Text
Adodc1.Recordset!Nm_mekanik = Text2.Text
Adodc1.Recordset!Alamat_mekanik = Text3.Text
Adodc1.Recordset!Telpon = Text5.Text
Adodc1.Recordset.Update
MsgBox "Data Berhasil Disimpan", vbInformation, "Info"
Adodc1.Refresh
Call bersih
Call autokode
End If
End Sub
Sub bersih()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text5.Text = ""
End Sub
Private Sub tampilkanketextbox()
Call BukaDB
        RSCustomer.Open "Select * from Customer ", Koneksi
Text1.Text = Adodc1.Recordset!ID_mekanik
Text2.Text = Adodc1.Recordset!Nm_mekanik
Text3.Text = Adodc1.Recordset!Alamat_mekanik
Text5.Text = Adodc1.Recordset!Telpon
Adodc1.Recordset.Update
End Sub
Private Sub Command2_Click()
    If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text5.Text = "" Then
        MsgBox "Data Harus Dipilih Terlebih Dahulu", vbInformation, "Info"
            ElseIf MsgBox("Anda Yakin Ingin Menghapus", vbYesNo, "Info") = vbYes Then
                Adodc1.Recordset.Delete
                Adodc1.Recordset.Update
                MsgBox "Data Berhasil di Hapus", vbInformation, "Info"
            ElseIf vbNo Then
           MsgBox "DIBATALKAN !", vbCritical, "Keluar"
    End If
Call bersih
Command1.Enabled = True
Call autokode
End Sub

Private Sub Command3_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Then
MsgBox "Data Harus Dipilih Terlebih Dahulu", vbInformation, "Info"
Else
With Adodc1.Recordset
!ID_mekanik = Text1.Text
!Nm_mekanik = Text2.Text
!Alamat_mekanik = Text3.Text
!Telpon = Text5.Text
.Update
MsgBox "Data Berhasil Diubah", vbInformation, "Info"
End With
End If
Command1.Enabled = True
Call bersih
Call autokode
End Sub

Private Sub Command4_Click()
Unload Me
End Sub
Private Sub DataGrid1_Click()
tampilkanketextbox
Command1.Enabled = False
End Sub

Sub autokode()
Call BukaDB
RSMekanik.Open ("select * from Mekanik Where ID_mekanik in(select max(ID_mekanik)from Mekanik)order by ID_mekanik Desc"), Koneksi
RSMekanik.Requery
    Dim Urutan As String
    Dim Hitung As String
    With Mekanik
        If RSMekanik.EOF Then
            Urutan = "MK" + "0001"
            Text1.Text = Urutan
        Else
            Hitung = Right(RSMekanik!ID_mekanik, 4) + 1
            Urutan = "MK" + Right("0000" & Hitung, 4)
        End If
        Text1.Text = Urutan
    End With
End Sub

Private Sub Form_Activate()
Call BukaDB
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\Bengkel.mdb"
Adodc1.RecordSource = "Mekanik"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub Form_Load()
Call autokode

End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
'ubah karakter jadi besar semua
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
Text3.SetFocus
End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
'ubah karakter jadi besar semua
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
Text5.SetFocus
End If
End Sub

Private Sub Text4_Change()
Call BukaDB
Koneksi.CursorLocation = adUseClient
RSMekanik.Open "select * from Mekanik where ID_mekanik like '%" & Text4.Text & "%' or Nm_mekanik like '%" & Text4.Text & "%'", Koneksi
    With RSMekanik
        If Not (.BOF And .EOF) Then
            mvBookmark = .Bookmark
        End If
    End With
Set DataGrid1.DataSource = RSMekanik.DataSource
DataGrid1.Refresh
End Sub


Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If IsNumeric(Text5.Text) = False Then
MsgBox "Maaf Data Yang Dimasukan Harus Angka!", vbInformation
Exit Sub
End If
Combo1.SetFocus
End If
End Sub



