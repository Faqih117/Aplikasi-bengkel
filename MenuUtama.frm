VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form MenuUtama 
   Caption         =   "Menu Utama"
   ClientHeight    =   5175
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11955
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   Picture         =   "MenuUtama.frx":0000
   ScaleHeight     =   5175
   ScaleWidth      =   11955
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "Mekanik"
      Height          =   855
      Left            =   2520
      Picture         =   "MenuUtama.frx":5C123
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   480
      Width           =   975
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   4800
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "Kode"
            TextSave        =   "Kode"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "Nama"
            TextSave        =   "Nama"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "Tanggal"
            TextSave        =   "Tanggal"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   8760
      TabIndex        =   8
      Top             =   120
      Width           =   2775
      Begin VB.CommandButton Command9 
         Caption         =   "Keluar"
         Height          =   855
         Left            =   1440
         Picture         =   "MenuUtama.frx":5D56D
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Keluar Akun"
         Height          =   855
         Left            =   240
         Picture         =   "MenuUtama.frx":5E9B7
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Laporan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6240
      TabIndex        =   5
      Top             =   120
      Width           =   2535
      Begin VB.CommandButton Command7 
         Caption         =   "Riwayat"
         Height          =   855
         Left            =   1320
         Picture         =   "MenuUtama.frx":5FE01
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Servis"
         Height          =   855
         Left            =   240
         Picture         =   "MenuUtama.frx":6124B
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Transaksi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3720
      TabIndex        =   3
      Top             =   120
      Width           =   2535
      Begin VB.CommandButton Command10 
         Caption         =   "Pelanggan"
         Height          =   855
         Left            =   240
         Picture         =   "MenuUtama.frx":62695
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Servis"
         Height          =   855
         Left            =   1320
         Picture         =   "MenuUtama.frx":63ADF
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Barang"
      Height          =   855
      Left            =   1440
      Picture         =   "MenuUtama.frx":64F29
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Master"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.CommandButton Command1 
         Caption         =   "Pengguna"
         Height          =   855
         Left            =   240
         Picture         =   "MenuUtama.frx":66373
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Label Label1 
      Height          =   1575
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   30000
   End
End
Attribute VB_Name = "MenuUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Pengguna.Show
End Sub

Private Sub Command10_Click()
Pelanggan.Show
End Sub

Private Sub Command2_Click()
Barang.Show
End Sub

Private Sub Command3_Click()
Mekanik.Show
End Sub

Private Sub Command4_Click()
Transaksi.Show
End Sub

Private Sub Command6_Click()
Lapservice.Show
End Sub

Private Sub Command7_Click()
riwayat.Show
End Sub

Private Sub Command8_Click()
Login.Show
Unload Me
End Sub

Private Sub Command9_Click()
End
End Sub

