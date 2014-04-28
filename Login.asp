<%
Response.Expires = -1000 'Makes the browser not cache this page
Response.Buffer = True 'Buffers the content so our Response.Redirect will work

Dim Error_Msg

login = Request.Form("login")
If login = "login_again" Then
    Session("UserLoggedIn") = ""
    ShowLogin
Else
    If Session("UserLoggedIn") = "true" Then
        AlreadyLoggedIn
    Else
        If login = "true" Then
            CheckLogin
        Else
            ShowLogin
        End If
    End If
End If

Sub ShowLogin
Response.Write(Error_Msg & "<br>")
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

<BODY>
    <form name=form1 action=Login.asp method=post>
    Team Name : <input type=text name=teamname> <br>
    Password : <input type=password name=pwd><br>
    <input type=hidden name=login value=true>
    <input type=submit value="Login">
    </form>


<%
End Sub

Sub AlreadyLoggedIn
%>

You are already logged in.
Do you want to logout or login as a different user?
<form name=form2 action=MainMenu1.asp method=post>
<input type=submit name=button1 value='Yes'>
<input type=hidden name=login value='login_again'>
</form>

<%
End Sub

Sub CheckLogin
Dim Conn, cStr, sql, RS, teamname, pwd
teamname = Request.Form("teamname")
pwd = Request.Form("pwd")
Set Conn = Server.CreateObject("ADODB.Connection")
cStr = "DRIVER={Microsoft Access Driver (*.mdb)};"
cStr = cStr & "DBQ=" & Server.MapPath("db\TeamSports.mdb") & ";"
Conn.Open(cStr)
sql = "select teamname from Team where teamname = '" & LCase(teamname) & "'"
sql = sql & " and pwd = '" & LCase(pwd) & "'"
Set RS = Conn.Execute(sql)
If RS.BOF And RS.EOF Then
    CheckLoginII
Else 
    Session("TeamName") = Request.Form("teamname")
    Session("UserLoggedIn") = "true"
    Response.Redirect "Manager.asp"
End If
End Sub
%>
<%
Sub CheckLoginII
If LCase(Request.Form("teamname")) = "admin" And LCase(Request.Form("pwd")) = "test" Then 
    Session("UserLoggedIn") = "true"
    Response.Redirect "Staff.asp"
Else
    Response.Write("Login Failed.<br><br>")
    ShowLogin
End If

End Sub
%>

</Body>
</Html>