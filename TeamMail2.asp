
<%
Response.Expires = -1000 'Makes the browser not cache this page
Response.Buffer = True 'Buffers the content so our Response.Redirect will work

If Session("UserLoggedIn") <> "true" Then
    Response.Redirect("Login.asp")
End If
%>
<%
dbfile=Server.MapPath("db\TeamSports.mdb") 
Set OBJdbConnection=Server.CreateObject("ADODB.Connection")
OBJdbConnection.Open "driver={Microsoft Access Driver (*.mdb)};dbq="&dbfile

Set rs = Server.CreateObject("ADODB.Recordset")
sql=  "SELECT * FROM Roster WHERE teamname = '" & Session("TeamName") & "' " 
rs.Open sql, OBJdbConnection, 3, 3

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

</head>
    

    <FORM ACTION="TeamMail.asp" METHOD=POST name="form1" onSubmit="return VerifyData()">
   <TABLE BORDER=0>
                <TR>
                 <TD>
                 <Select name = "email">
                 <OPTION VALUE = "all">all</OPTION>
                 <% While Not RS.EOF %>
                 <option value="<%=RS("email")%>">
                 <%=RS.Fields("email")%>
                 <%
                 RS.MoveNext
                  Wend
                 RS.Close
                 
                 OBJdbConnection.Close
                
                 Set RS = Nothing
                 set OBJdbConnection = Nothing
                 %>
                 </TD>
                </TR>
                <TR>
                  <TD><font size="4"><P>Subject:
                  <INPUT TYPE="TEXT" NAME="subject"></P></TD>
                 </TR>
              </TABLE>
   <P>Please enter the message you would like to email to your team:<BR>
   <CENTER><textarea name="message" rows="10" cols="70" wrap="virtual"> 
   </textarea></P>
   <P><INPUT TYPE="RESET" VALUE="Reset">
   <INPUT TYPE="SUBMIT" NAME="button1" VALUE="Submit"></CENTER></P>
</FORM>

</BODY>

</HTML>
