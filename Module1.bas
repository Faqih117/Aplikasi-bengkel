Attribute VB_Name = "Module1"
Public Koneksi As New ADODB.Connection
Public RSUser As ADODB.Recordset
Public RSJasa As ADODB.Recordset
Public RSCustomer As ADODB.Recordset
Public RSDetailjasa As ADODB.Recordset
Public RSDetailbarang As ADODB.Recordset
Public RSBarang As ADODB.Recordset
Public RSPenjualan As ADODB.Recordset
Public RSDetailJual As ADODB.Recordset
Public RSMekanik As ADODB.Recordset
Public RShapus As ADODB.Recordset
Public RSServis As ADODB.Recordset
Public Sub BukaDB()
 Set Koneksi = New ADODB.Connection
 Set RSUser = New ADODB.Recordset
 Set RSJasa = New ADODB.Recordset
 Set RSCustomer = New ADODB.Recordset
 Set RSDetailjasa = New ADODB.Recordset
 Set RSDetailbarang = New ADODB.Recordset
 Set RSBarang = New ADODB.Recordset
 Set RSPenjualan = New ADODB.Recordset
 Set RSDetailJual = New ADODB.Recordset
 Set RSMekanik = New ADODB.Recordset
 Set RShapus = New ADODB.Recordset
 Set RSServis = New ADODB.Recordset
 Koneksi.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\bengkel.mdb"
End Sub

