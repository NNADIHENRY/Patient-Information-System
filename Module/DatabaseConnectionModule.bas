Attribute VB_Name = "DatabaseConnectionModule"
Public databaseconnection As ADODB.Connection, recordset As ADODB.recordset, recordset2 As ADODB.recordset, usertype As String
Public recordset3 As ADODB.recordset
Public gListItem As MSComctlLib.ListItem
Public Sub Main()
Set databaseconnection = New ADODB.Connection
databaseconnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source= " & App.Path & "\D'Base\cityhealth.mdb"
'dr.conn.ConnectionString = App.Path & "\cityhealth.mdb"
frmLogin.Show
'frmMain.Show
End Sub
