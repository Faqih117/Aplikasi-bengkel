VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form progressbar 
   Caption         =   "SELAMAT DATANG DI APLIKASI BENGKEL KELOMPOK 1"
   ClientHeight    =   5055
   ClientLeft      =   2910
   ClientTop       =   3180
   ClientWidth     =   7425
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form8"
   ScaleHeight     =   5055
   ScaleWidth      =   7425
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1080
      Top             =   4200
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   3720
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Min             =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3720
      TabIndex        =   1
      Top             =   4560
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   5400
      Left            =   0
      Picture         =   "progressbar.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7560
   End
End
Attribute VB_Name = "progressbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Timer1.Enabled = True
End Sub


Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 1
Label1.Caption = ProgressBar1.Value + 1 & "%"
    If ProgressBar1.Value = ProgressBar1.Max Then
        ProgressBar1.Value = ProgressBar1.Max
        Unload Me
        Login.Show
    End If
End Sub
