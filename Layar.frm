VERSION 5.00
Begin VB.Form Layar 
   BackColor       =   &H8000000E&
   Caption         =   "cetak layar"
   ClientHeight    =   4530
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   3975
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Layar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

