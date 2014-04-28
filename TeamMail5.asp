<%Dim stremail, strmessage, strsubject
stremail = request.form("email")
strsubject = request.form("subject")
strmessage = request.form("message")
If stremail = "all" then 
response.redirect("TeamMail.asp")
else
response.redirect("TeamMail6.asp")
end if
%>