
<HTML>
<HEAD>
<TITLE>RECsports</TITLE>
<STYLE>
<!--
		

		A:link {text-decoration: none; color: #cc0033;}
		A:visited {text-decoration: none; color: #cccccc;}
		A:active {text-decoration: none; color: #000000;}
		A:hover {text-decoration: none; color:#0000FF;}
-->
</STYLE>
</HEAD>

<BODY>
<CENTER><H1>Schedule</H1></CENTER>
  

<%
dbfile=Server.MapPath("db\TeamSports.mdb") 
Set OBJdbConnection=Server.CreateObject("ADODB.Connection")
OBJdbConnection.Open "driver={Microsoft Access Driver (*.mdb)};dbq="&dbfile

Set rs = Server.CreateObject("ADODB.Recordset")
sql=  "SELECT * FROM Schedule WHERE AwayTeam = '" & Request.form("teamname") & "' " & _
"OR HomeTeam = '" & Request.form("teamname") & "' "
rs.Open sql, OBJdbConnection, 3, 3

%>
<CENTER>
<table border=1>
<tr>
    <th>Date</th>
    <th>Time</th>
    <th>Field</th>
    <th>Home Team</th>
    <th>Away Team</th>
</tr>
<% do while not rs.EOF
  response.write "<tr><td>"
  response.write rs("Date") & "</td><td>"
  response.write rs("Time") & "</td><td>"
  response.write rs("Court") & "</td><td>"
  response.write rs("HomeTeam") & "</td><td>"
  response.write rs("AwayTeam") & "</td>"
  rs.movenext
loop
rs.close
set rs = nothing
OBJdbConnection.Close 
set OBJdbConnection = Nothing
%>
</table>
</CENTER>

</body></html>
