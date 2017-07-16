Attribute VB_Name = "Module1"
Public con As New ADODB.Connection
Public rs As New ADODB.Recordset

Sub main()
If con.State = 1 Then con.Close
con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\DB\DB.mdb;Persist Security Info=False"
con.Open
frmBill.Show
End Sub
