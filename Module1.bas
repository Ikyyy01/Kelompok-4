Attribute VB_Name = "Module1"
Public Koneksi As New ADODB.Connection
Public RSAdmin As New ADODB.Recordset
Public RSJabatan As New ADODB.Recordset
Public RSKaryawan As New ADODB.Recordset
Public RSGaji As New ADODB.Recordset

Public Sub BukaDB()
Set Koneksi = New ADODB.Connection
Set RSAdmin = New ADODB.Recordset
Set RSJabatan = New ADODB.Recordset
Set RSKaryawan = New ADODB.Recordset
Set RSGaji = New ADODB.Recordset

Koneksi.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DBPenggajian.mdb"

End Sub
