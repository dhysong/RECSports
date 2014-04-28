<%
'Read data from querystring
'and replace the ' character with ''
StrName   = replace(Request("player"),"'","''")
StrEmail  = replace(Request("email"),"'","''")
StrPhone= replace(Request("phone"),"'","''")
StrSSN= replace(Request("ssn"),"'","''")
Strusername= replace(Request("username"),"'","''")


'The Access database is in the same directory
dbfile=Server.MapPath("db\TeamSports.mdb")
'Create a Connection Object
Set OBJdbConnection=Server.CreateObject("ADODB.Connection")

'Open the connection (specify the driver tyoe = Access 
'and the file is TeamSports.mdb
OBJdbConnection.Open "driver={Microsoft Access Driver (*.mdb)};dbq="&dbfile


'Build sql string
' Syntax: INSERT into TABLE (field1 [, field2, ...])
' VALUES (value1 [,value2, ...] )
sql_ins="INSERT into Roster (player, email, phone, ssn, teamname) VALUES " & _
		"('" & StrName & "', '" & StrEmail & "', '" & _
		StrPhone & "', '" & StrSSN & "', '" & Strusername & "') "

'Recordset object is created
Set rs1 = Server.CreateObject("ADODB.Recordset")

'Sql command is sent
rs1.Open sql_ins, OBJdbConnection, 3, 3

Response.Redirect("Roster.asp")

%>