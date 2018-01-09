Attribute VB_Name = "ModuleKoneksi"
Public Conn As ADODB.Connection
Public RS As ADODB.Recordset

Public Sub BukaDatabase()
Set Conn = New ADODB.Connection
Set RS = New ADODB.Recordset
Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\Database\db_keuangan.mdb;"
Conn.CursorLocation = adUseClient
End Sub


