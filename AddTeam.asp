<%
Dim strTo, strSubject, strBody 'Strings for recipient, subject, body
Dim objCDOMail 'The CDO object


'First we'll read in the values entered
strTo = Request.Form("email")

'These would read the message subject and body if we let you enter it
strSubject = "Registration Confirmation"
strBody = "You have successfully registered your intramural team." 
strBody = strBody & "      Team Name: "
strBody = strBody &  Request.Form("teamname")
strBody = strBody & "  Captain: "
strBody = strBody &  Request.Form("captain")
strBody = strBody & "  Sport: "
strBody = strBody &  Request.Form("keyword")
strBody = strBody & "  Password: "
strBody = strBody &  Request.Form("pwd")

' Some spacing:
strBody = strBody & vbCrLf & vbCrLf



' A final carriage return for good measure!
strBody = strBody & vbCrLf


'Ok we've got the values now on to the point of the script.

'We just check to see if someone has entered anything into the to field.
'If it's equal to nothing we show the form, otherwise we send the message.
'If you were doing this for real you might want to check other fields too
'and do a little entry validation like checking for valid syntax etc.

' Note: I was getting so many bad addresses being entered and bounced
' back to me by mailservers that I've added a quick validation routine.
If strTo = "" Or Not IsValidEmail(strTo) Then
	
Response.Redirect("RegisterTeam.asp")
%>

<%
Else
	' Create an instance of the NewMail object.
	Set objCDOMail = Server.CreateObject("CDONTS.NewMail")
	
	' Set the properties of the object
	
	'***********************************************************
	' PLEASE CHANGE THESE SO WE DON'T APPEAR TO BE SENDING YOUR
	' EMAIL. WE ALSO DON'T WANT THE EMAILS TO GET SENT TO US
	' WHEN SOMETHING GOES WRONG WITH YOUR SCRIPT... THANKS
	'***********************************************************
	
	' This syntax works fine
	'objCDOMail.From = "drew10@arches.uga.edu"
	' But this gets you the appearance of a real name!
	objCDOMail.From = "RECsports <drew10@arches.uga.edu>"
	objCDOMail.To = strTo
	objCDOMail.Subject = strSubject
	objCDOMail.Body = strBody

	' There are lots of other properties you can use.
	' You can send HTML e-mail, attachments, etc...
	' You can also modify most aspects of the message
	' like importance, custom headers, ...
	' Check the documentation for a full list as well
	' as the correct syntax.

	' Some of the more useful ones I've included samples of here:
	'objCDOMail.Cc = "user@domain.com;user@domain.com"
	'objCDOMail.Bcc = "user@domain.com;user@domain.com"
	'objCDOMail.Importance = 1 '(0=Low, 1=Normal, 2=High)
	'objCDOMail.AttachFile "c:\path\filename.txt", "filename.txt"

	' I've had several requests for how to send HTML email.
	' To do so simply set the body format to HTML and then
	' compose your body using standard HTML tags.
	'objCDOMail.BodyFormat = 0 ' CdoBodyFormatHTML
	
	'Outlook gives you grief unless you also set:
	'objCDOMail.MailFormat = 0 ' CdoMailFormatMime

	' Send the message!
	objCDOMail.Send
	
	' Set the object to nothing because it immediately becomes
	' invalid after calling the Send method.
	Set objCDOMail = Nothing

End If
' End page logic
%>

<% ' Only functions and subs follow!

' A quick email syntax checker.  It's not perfect,
' but it's quick and easy and will catch most of
' the bad addresses than people type in.
Function IsValidEmail(strEmail)
	Dim bIsValid
	bIsValid = True
	
	If Len(strEmail) < 5 Then
		bIsValid = False
	Else
		If Instr(1, strEmail, " ") <> 0 Then
			bIsValid = False
		Else
			If InStr(1, strEmail, "@", 1) < 2 Then
				bIsValid = False
			Else
				If InStrRev(strEmail, ".") < InStr(1, strEmail, "@", 1) + 2 Then
					bIsValid = False
				End If
			End If
		End If
	End If

	IsValidEmail = bIsValid
End Function
%>


<%
'Read data from querystring
'and replace the ' character with ''
StrName   = replace(Request("teamname"),"'","''")
StrEmail  = replace(Request("email"),"'","''")
StrCaptain= replace(Request("captain"),"'","''")
StrPwd= replace(Request("pwd"),"'","''")
StrKeyword= replace(Request("keyword"),"'","''")


'The Access database is in the same directory
dbfile=Server.MapPath("db\TeamSports.mdb")
'Create a Connection Object
Set OBJdbConnection=Server.CreateObject("ADODB.Connection")

'Open the connection (specify the driver tyoe = Access 
'and the file is guestbook.mdb
OBJdbConnection.Open "driver={Microsoft Access Driver (*.mdb)};dbq="&dbfile


'Build sql string
' Syntax: INSERT into TABLE (field1 [, field2, ...])
' VALUES (value1 [,value2, ...] )
sql_ins="INSERT into Team (teamname, email, captain, pwd, keyword) VALUES " & _
		"('" & StrName & "', '" & StrEmail & "', '" & _
		StrCaptain & "', '" & StrPwd & "', '" & StrKeyword & "') "

'Recordset object is created
Set rs1 = Server.CreateObject("ADODB.Recordset")

'Sql command is sent
rs1.Open sql_ins, OBJdbConnection, 3, 3

Response.Redirect("Registrationconfirmation.asp")

%>
