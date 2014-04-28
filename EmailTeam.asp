<%
dbfile=Server.MapPath("db\TeamSports.mdb") 
Set OBJdbConnection=Server.CreateObject("ADODB.Connection")
OBJdbConnection.Open "driver={Microsoft Access Driver (*.mdb)};dbq="&dbfile

Set rs = Server.CreateObject("ADODB.Recordset")
sql=  "SELECT * FROM Team WHERE keyword = '" & Request.form("Sport") & "' and paid = No " 
rs.Open sql, OBJdbConnection, 3, 3

%>
<% 

rs.MoveFirst
MyVar= ""
delimiter = ""
do while not rs.eof
If Not IsNull(rs("email")) Then
MyVar= MyVar & delimiter & rs("email")
delimiter = "; "
End If
rs.movenext
loop%> 

<%
Dim strTo, strSubject, strBody 'Strings for recipient, subject, body
Dim objCDOMail 'The CDO object


'First we'll read in the values entered
strTo = Myvar

'These would read the message subject and body if we let you enter it
strSubject = "Registration Deadline Notification"
strBody = Request.Form("message") 


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

	
%>

<%

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


%>


<%Response.Redirect("Staff.asp")%>
