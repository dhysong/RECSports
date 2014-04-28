<%

'The Access database is in the same directory
dbfile=Server.MapPath("db\TeamSports.mdb")
'Create a Connection Object
Set OBJdbConnection=Server.CreateObject("ADODB.Connection")

'Open the connection (specify the driver tyoe = Access 
'and the file is TeamSports.mdb
OBJdbConnection.Open "driver={Microsoft Access Driver (*.mdb)};dbq="&dbfile



sql="delete * from Roster where player = '" & Request.form("Participant") & "' "

'Recordset object is created
Set rs1 = Server.CreateObject("ADODB.Recordset")

'Sql command is sent
rs1.Open sql, OBJdbConnection, 3, 3

Response.Redirect("Roster.asp")

%>