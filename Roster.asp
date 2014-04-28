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
<BODY><CENTER><H1>Adjust Roster</H1></CENTER>
 <%dbfile=Server.MapPath("db\TeamSports.mdb") 
Set OBJdbConnection=Server.CreateObject("ADODB.Connection")
OBJdbConnection.Open "driver={Microsoft Access Driver (*.mdb)};dbq="&dbfile

Set rs = Server.CreateObject("ADODB.Recordset")
sql=  "SELECT * FROM Roster WHERE teamname = '" & Session("TeamName") & "' " 
rs.Open sql, OBJdbConnection, 3, 3%> 
            <CENTER>
<table border=1 cellspacing=0 cellpadding=5>
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
</CENTER>
<BR> 
 
            <FORM ACTION="AddPlayer.asp" METHOD=POST>
              <p style="margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>
              <input type="hidden" name="username" value="<%= Session("TeamName") %>">
              <p style="margin-top: 0; margin-bottom: 0" align="center"><b><font size="4">Roster 
                Change</font></b></p>
              <div align="center"> 
                <center>
                  <table border="1" width="80%" cellpadding=5>
                    <tr> 
                      <td width="50%" align="center"> 
                        <p style="margin-top: 0; margin-bottom: 0">Name:</p>
                        <p style="margin-top: 0; margin-bottom: 0" align="center"> 
                          <input type="text" name="player" size="30" tabindex="3">
                      </td>
                      <td width="50%" align="center"> 
                        <p style="margin-top: 0; margin-bottom: 0">Student ID 
                          Number (without dashes):</p>
                        <p style="margin-top: 0; margin-bottom: 0" align="center"> 
                          <input type="text" name="ssn" size="20" tabindex="4" maxlength="9">
                      </td>
                    </tr>
                    <tr> 
                      <td width="50%" align="center"> 
                        <p style="margin-top: 0; margin-bottom: 0">Email:</p>
                        <p style="margin-top: 0; margin-bottom: 0" align="center"> 
                          <input type="text" name="email" size="40" tabindex="5">
                      </td>
                      <td width="50%" align="center"> 
                        <p style="margin-top: 0; margin-bottom: 0">Phone Number:</p>
                        <p style="margin-top: 0; margin-bottom: 0" align="center"> 
                          <input type="text" name="phone" size="20" tabindex="6">
                      </td>
                    </tr>
                    <tr> 
                  <td width="100%"> 
                    <p align="center"> 
                      <input TYPE="submit" VALUE="Submit" name="Submit" tabindex="85"></a></td> 
                      &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                      <td><center><input TYPE="RESET" VALUE="Reset Form" tabindex="86"></center>
                  </td>
                </tr>
                <tr> 
                  <td width="100%"> 
                    <p align="center">
                  </td>
                </tr>
                  </table>
                </center>
              </div><p 
              <p style="margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>
              
            </FORM>
 <%
dbfile=Server.MapPath("db\TeamSports.mdb") 
Set OBJdbConnection=Server.CreateObject("ADODB.Connection")
OBJdbConnection.Open "driver={Microsoft Access Driver (*.mdb)};dbq="&dbfile

Set rs = Server.CreateObject("ADODB.Recordset")
sql=  "SELECT * FROM Roster WHERE teamname = '" & Session("TeamName") & "' " 
rs.Open sql, OBJdbConnection, 3, 3

%>

<FORM ACTION="DeletePlayer.asp" METHOD=POST name="MyForm">
<Center>
<TABLE BORDER=1 WIDTH=80%>
  <TR ALIGN=CENTER>
    <TD WIDTH=50%><Select name = "Participant">
                 <% While Not RS.EOF %>
                 <option value="<%=RS("player")%>">
                 <%=RS.Fields("player")%>
                 <%
                 RS.MoveNext
                  Wend
                 RS.Close
                 
                 OBJdbConnection.Close
                
                 Set RS = Nothing
                 set OBJdbConnection = Nothing
                 %></TD>
    <TD WIDTH=50%><INPUT TYPE="SUBMIT" VALUE="Delete Player"></TD>
  </TR>
</TABLE> 
</Center>        
</FORM>
</body></html>
