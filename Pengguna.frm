VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Pengguna 
   BackColor       =   &H8000000E&
   Caption         =   "Pengguna"
   ClientHeight    =   6600
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6345
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   6600
   ScaleWidth      =   6345
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2160
      TabIndex        =   15
      Top             =   3720
      Width           =   2175
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Pengguna.frx":0000
      Height          =   2055
      Left            =   360
      TabIndex        =   8
      Top             =   4320
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   3625
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
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
         DataField       =   "Kode_user"
         Caption         =   "Kode_user"
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
         DataField       =   "Nama_user"
         Caption         =   "Nama_user"
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
         DataField       =   "Pws_user"
         Caption         =   "Pws_user"
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
         DataField       =   "Status_user"
         Caption         =   "Status_user"
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
            Alignment       =   2
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3600
      Top             =   4920
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
      RecordSource    =   "Login"
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
      Left            =   2160
      TabIndex        =   7
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   2760
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
      TabIndex        =   9
      Top             =   1680
      Width           =   1455
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000E&
         Caption         =   "Simpan"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Hapus"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Ubah"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Tutup"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1680
         Width           =   975
      End
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   480
      Picture         =   "Pengguna.frx":0015
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cari Data"
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kode pengguna"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nama pengguna"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Password"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   3240
      Width           =   1695
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
      Left            =   1680
      TabIndex        =   17
      Top             =   960
      Width           =   3975
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
      Left            =   1680
      TabIndex        =   16
      Top             =   480
      Width           =   3975
   End
End
Attribute VB_Name = "Pengguna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call BukaDB
    RSUser.Open "Select * from Login", Koneksi
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Combo1 = "" Then
MsgBox "Data Belum Lengkap", vbInformation, "Info"
Exit Sub
    Else
        Adodc1.Recordset.AddNew
        Adodc1.Recordset!Kode_user = Text1.Text
        Adodc1.Recordset!Nama_user = Text2.Text
        Adodc1.Recordset!Pws_user = Text3.Text
        Adodc1.Recordset!Status_user = Combo1.Text
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
Combo1.Text = ""
End Sub
Private Sub tampilkanketextbox()
Call BukaDB
        RSUser.Open "Select * from Login ", Koneksi
Text1.Text = Adodc1.Recordset!Kode_user
Text2.Text = Adodc1.Recordset!Nama_user
Text3.Text = Adodc1.Recordset!Pws_user
Combo1.Text = Adodc1.Recordset!Status_user
Adodc1.Recordset.Update
End Sub
Private Sub Command2_Click()
    If Text1 = "" Or Text2 = "" Or Text3 = "" Or Combo1.Text = "" Then
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
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Combo1 = "" Then
MsgBox "Data Harus Dipilih Terlebih Dahulu", vbInformation, "Info"
Else
With Adodc1.Recordset
!Kode_user = Text1.Text
!Nama_user = Text2.Text
!Pws_user = Text3.Text
!Status_user = Combo1.Text
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

Sub autokode()
Call BukaDB
RSUser.Open ("select * from Login Where Kode_user in(select max(Kode_user)from Login)order by Kode_user Desc"), Koneksi
RSUser.Requery
    Dim Urutan As String
    Dim Hitung As String
    With Login
        If RSUser.EOF Then
            Urutan = "USR" + "01"
            Text1.Text = Urutan
        Else
            Hitung = Right(RSUser!Kode_user, 2) + 1
            Urutan = "USR" + Right("00" & Hitung, 2)
        End If
        Text1.Text = Urutan
    End With
End Sub


Private Sub DataGrid1_Click()
tampilkanketextbox
Command1.Enabled = False
End Sub

Private Sub Form_Activate()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\Bengkel.mdb"
Adodc1.RecordSource = "Login"
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
Combo1.AddItem ("admin")
Combo1.AddItem ("kasir")
Text3.PasswordChar = "X"
End Sub



Private Sub Text2_KeyPress(KeyAscii As Integer)
'ubah karakter jadi besar semua
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
Text2.MaxLength = 15
Text2.Enabled = False
Text3.SetFocus
End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
'ubah karakter jadi besar semua
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
Text3.MaxLength = 10
Text3.Enabled = False
Command1.SetFocus
End If
End Sub

Private Sub Text4_Change()
Call BukaDB
Koneksi.CursorLocation = adUseClient
RSUser.Open "select * from Login where Kode_user like '%" & Text4.Text & "%' or Nama_user like '%" & Text4.Text & "%'", Koneksi
    With RSUser
        If Not (.BOF And .EOF) Then
            mvBookmark = .Bookmark
        End If
    End With
Set DataGrid1.DataSource = RSUser.DataSource
DataGrid1.Refresh

End Sub

