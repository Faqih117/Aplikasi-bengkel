VERSION 5.00
Begin VB.Form Login 
   Caption         =   "Login"
   ClientHeight    =   3975
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5475
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   5475
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      BackColor       =   &H8000000E&
      Caption         =   "TERLIHAT"
      Height          =   375
      Left            =   4080
      MaskColor       =   &H8000000F&
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Batalkan"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Oke"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nama pengguna"
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kata sandi pengguna"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   4080
      Left            =   -360
      Picture         =   "Login.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5880
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A As Byte

Private Sub Check1_Click()
If Check1.Value = 1 Then
    Text2.PasswordChar = ""
    Else
    Text2.PasswordChar = "X"
End If
End Sub

Private Sub Command1_Click()
Call BukaDB
        RSUser.Open "Select * from Login where Nama_user ='" & Text1 & "' and Pws_user='" & Text2 & "'", Koneksi
        If RSUser.EOF Then
        A = A + 1
            If 1 - A = 0 Then
                MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
                        "User dan Password tidak dikenal"
                Text2.SetFocus
            ElseIf 2 - A = 0 Then
                MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
                        "User dan Password tidak dikenal"
            ElseIf 3 - A = 0 Then
                MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
                        "User dan Password tidak dikenal" & Chr(13) & _
                        "Kesempatan habis, Ulangi dari awal"
                End
            End If
        Else
                 Unload Me
                 MenuUtama.Show
                 MsgBox "Berhasil Masuk !", vbInformation, "Masuk !"
              MenuUtama.StatusBar1.Panels(2) = RSUser!Kode_user
              MenuUtama.StatusBar1.Panels(6) = RSUser!Status_user
              MenuUtama.StatusBar1.Panels(4) = RSUser!Nama_user
              MenuUtama.StatusBar1.Panels(8) = Date
           If MenuUtama.StatusBar1.Panels(6) = "admin" Then
               MenuUtama.Command1.Enabled = True
               MenuUtama.Command2.Enabled = True
               MenuUtama.Command7.Enabled = True
               MenuUtama.Command4.Enabled = False
           ElseIf MenuUtama.StatusBar1.Panels(6) = "kasir" Then
               MenuUtama.Command1.Enabled = False
               MenuUtama.Command2.Enabled = False
               MenuUtama.Command3.Enabled = False
               MenuUtama.Command7.Enabled = False
           End If
        End If
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
Text1.MaxLength = 15
Text2.MaxLength = 10
Text2.PasswordChar = "X"
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
'ubah karakter jadi besar semua
KeyAscii = Asc(UCase(Chr(KeyAscii)))
'jika menekan ESC form ditutup
If KeyAscii = 27 Then Unload Me
'jika menekan enter setelah mengisi nama, maka..
If KeyAscii = 13 Then
    'buka database
    Call BukaDB
    'cari nama kasir yang diketik
    RSUser.Open "Select * from Login where Nama_user ='" & Text1 & "'", Koneksi
    'jika tidak ditemukan, maka
    If RSUser.EOF Then
        'batasi akses ke nama kasir 3 kali kesempatan
        A = A + 1
        If 1 - A = 0 Then
            MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
                    "Nama '" & Text1 & "' tidak dikenal"
            Text1 = ""
            Text1.SetFocus
        ElseIf 2 - A = 0 Then
            MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
                    "Nama '" & Text1 & "' tidak dikenal"
            Text1 = ""
            Text1.SetFocus
        ElseIf 3 - A = 0 Then
            MsgBox "Kesempatan ke " & A & " Salah" & Chr(13) & _
                    "Nama '" & Text1 & "' tidak dikenal" & Chr(13) & _
                    "Kesempatan habis, Ulangi dari awal"
            'End
            Unload Me
        End If
    Else
        'jika nama kasir benar, maka nama kasir menjadi false
        Text1.Enabled = False
        'password kasir menjadi true dan menjadi fokus kursor
        Text2.Enabled = True
        Text2.SetFocus
    End If
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
'ubah karakter jadi besar semua
KeyAscii = Asc(UCase(Chr(KeyAscii)))
'jika menekan ESC form ditutup
If KeyAscii = 27 Then
    Unload Me
ElseIf KeyAscii = 13 Then
    Command1.SetFocus

End If
End Sub

