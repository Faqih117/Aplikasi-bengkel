VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Lapservice 
   Caption         =   "Laporan servis"
   ClientHeight    =   4815
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   4020
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CR 
      Left            =   120
      Top             =   4200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame3 
      Caption         =   "Bulanan"
      Height          =   1335
      Left            =   480
      TabIndex        =   8
      Top             =   3120
      Width           =   3015
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   1320
         TabIndex        =   10
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   1320
         TabIndex        =   9
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Bulan"
         Height          =   375
         Left            =   600
         TabIndex        =   12
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Tahun"
         Height          =   375
         Left            =   600
         TabIndex        =   11
         Top             =   840
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Mingguan"
      Height          =   1335
      Left            =   480
      TabIndex        =   3
      Top             =   1680
      Width           =   3015
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Top             =   840
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Akhir"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Awal"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Harian"
      Height          =   855
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   3015
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "Lapservice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
If KeyAscii = 27 Then
Unload Me
End If
'buka database
Call BukaDB
'cari tanggal pembelian
RSPenjualan.Open "Select Distinct Tanggal From Penjualan order By 1", Koneksi
RSPenjualan.Requery
'tampilkan tanggal dalam combo
Do Until RSPenjualan.EOF
    Combo1.AddItem RSPenjualan!Tanggal
    Combo2.AddItem Format(RSPenjualan!Tanggal, "DD/MM/YYYY")
    Combo3.AddItem Format(RSPenjualan!Tanggal, "DD/MM/YYYY")
    RSPenjualan.MoveNext
Loop
 
'buatlah looping untuk bulan dari 1-12
'dan tahun mulai tahun 2001 s/d 2020
For i = 1 To 12
    Combo5.AddItem i
Next i
For i = 5 To 25
    Combo4.AddItem 2000 + i
Next i
 
End Sub

Private Sub Combo1_Keypress(KeyAscii As Integer)
If Combo1 = "" Or KeyAscii = 27 Then Unload Me
End Sub
 
'Lap Harian
Private Sub Combo1_Click()
    'panggil laporan yang tanggalnya = combo1
    CR.selectionFormula = "Totext({Penjualan.Tanggal})='" & Combo1 & "'"
    CR.ReportFileName = App.Path & "\LapJualHarian.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub
 
Private Sub Combo2_Keypress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub
 
'Lap Mingguan (Tgl Antara)
Private Sub Combo3_Click()
    'cegah data kosong di combo2 dan combo3
    If Combo2 = "" Then
        MsgBox "Tanggal awal kosong", , "Informasi"
        Combo2.SetFocus
    ElseIf Combo3 = Combo2 Then
        MsgBox "Tanggal Tidak Boleh Sama", , "Informasi"
        Exit Sub
    End If
    'panggil laporan yang tanggal awalnya=combo2 dan tanggal akhirnya = combo3
    CR.selectionFormula = "{Penjualan.Tanggal} in date (" & Combo2 & ") to date (" & Combo3 & ")"
    CR.ReportFileName = App.Path & "\LapJualMingguan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub
  
Private Sub Combo4_Keypress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub
 
 
'Lap Bulanan
Private Sub Combo4_Click()
    'buka database
    Call BukaDB
    'cari data yang tanggal dan bulannya dipilih di combo4 dan combo5
    RSPenjualan.Open "select * from Penjualan where month(tanggal)='" & Val(Combo5) & "' and year(tanggal)='" & (Combo4) & "'", Koneksi
    'jika tidak cocok, munculkan pesan
    If RSPenjualan.EOF Then
        MsgBox "Data tidak ditemukan"
        Exit Sub
        Combo4.SetFocus
    End If
    'jika ditemukan panggil file laporan yang
    'datanya bulannya=combo4 dan tahunnya= combo5
    CR.selectionFormula = "Month({Penjualan.Tanggal})=" & Val(Combo5.Text) & " and Year({Penjualan.Tanggal})=" & Val(Combo4.Text)
    CR.ReportFileName = App.Path & "\LapJualBulanan.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

