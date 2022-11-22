Attribute VB_Name = "MdlVCD"

Public Conn As New adodb.Connection
Public RSAnggota As adodb.Recordset
Public RSkasir As adodb.Recordset
Public RSFilm As adodb.Recordset
Public RSPinjam As adodb.Recordset
Public RSDetailPjm As adodb.Recordset
Public RSKembali As adodb.Recordset
Public RSDetailKbl As adodb.Recordset
Public RSTansPjm As adodb.Recordset
Public RSTansKbl As adodb.Recordset

Public Sub BukaDB()
Set Conn = New adodb.Connection
Set RSAnggota = New adodb.Recordset
Set RSkasir = New adodb.Recordset
Set RSFilm = New adodb.Recordset
Set RSPinjam = New adodb.Recordset
Set RSDetailPjm = New adodb.Recordset
Set RSKembali = New adodb.Recordset
Set RSDetailKbl = New adodb.Recordset
Set RSTansPjm = New adodb.Recordset
Set RSTansKbl = New adodb.Recordset
Conn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\ADOVCD.mdb"
End Sub


