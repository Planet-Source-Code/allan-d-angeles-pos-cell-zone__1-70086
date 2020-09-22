Attribute VB_Name = "connection"
Global con As New ADODB.connection
Global rst As New ADODB.Recordset
Global rst1 As New ADODB.Recordset

Public Sub connect()
Set con = Nothing
con.ConnectionString = "Provider = Microsoft.jet.oledb.4.0; Data Source = " & App.Path & "\database.mdb ;"
con.Properties("Jet OLEDB:Database Password") = "09231988"
con.Open
End Sub
