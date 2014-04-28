<%
Response.Expires = -1000 'Makes the browser not cache this page
Response.Buffer = True 'Buffers the content so our Response.Redirect will work

If Session("UserLoggedIn") <> "true" Then
    Response.Redirect("Login.asp")
End If
%>
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

<BODY><Center>
<H1><%= Session("TeamName") %> Team Management Page</H1></CENTER>
 
    <CENTER><H3>Schedule</H3></CENTER>
<%
dbfile=Server.MapPath("db\TeamSports.mdb") 
Set OBJdbConnection=Server.CreateObject("ADODB.Connection")
OBJdbConnection.Open "driver={Microsoft Access Driver (*.mdb)};dbq="&dbfile

Set rs = Server.CreateObject("ADODB.Recordset")
sql=  "SELECT * FROM Schedule WHERE AwayTeam = '" & Session("teamname") & "' " & _
"OR HomeTeam = '" & Session("teamname") & "' "
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
<BR>
<h3><center>Roster</center></h3>
<%
dbfile=Server.MapPath("db\TeamSports.mdb") 
Set OBJdbConnection=Server.CreateObject("ADODB.Connection")
OBJdbConnection.Open "driver={Microsoft Access Driver (*.mdb)};dbq="&dbfile

Set rs = Server.CreateObject("ADODB.Recordset")
sql=  "SELECT * FROM Roster WHERE teamname = '" & Session("TeamName") & "' " 
rs.Open sql, OBJdbConnection, 3, 3

%>
<CENTER>
<table border=1>
<tr>
    <th>Player</th>
    <th>Email</th>
    <th>Phone #</th>
</tr>
<% do while not rs.EOF
  response.write "<tr><td>"
  response.write rs("player") & "</td><td>"
  response.write rs("email") & "</td><td>"
  response.write rs("phone") & "</td>"
  rs.movenext
loop
rs.close
set rs = nothing
OBJdbConnection.Close 
set OBJdbConnection = Nothing
%>
</table>
</CENTER></TD></TR>
    
   
</body>
</html>
