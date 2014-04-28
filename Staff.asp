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

<BODY><CENTER><H1>Staff</H1></CENTER>
 
<FORM ACTION="DeleteTeam.asp" METHOD=POST name="MyForm">
<Center>
<TABLE BORDER=0 WIDTH=100%>
  <TR ALIGN=CENTER>
      <TD WIDTH=55% ALIGN="right"><font size="4">Delete all teams who have not paid for:</TD>
      <TD WIDTH=20%><SELECT NAME = "Sport">
      <OPTION VALUE = "Soccer">Soccer</OPTION>
      <OPTION VALUE = "Flag Football">Flag Football</OPTION>
      <OPTION VALUE = "Basketball">Basketball</OPTION>
      <OPTION VALUE = "Softball">Softball</OPTION>
      </SELECT></TD>
    <TD WIDTH=25%><INPUT TYPE="SUBMIT" VALUE="Delete Team"></TD>
  </TR>
</TABLE> 
</Center>        
</FORM> 
<BR><BR>  
<FORM ACTION="EmailTeam.asp" METHOD=POST name="MyForm2">
<Center>
<TABLE BORDER=0 WIDTH=100%>
  <TR ALIGN=CENTER>
      <TD WIDTH=55% ALIGN="right"><font size="4">Notify all teams who have not paid for:</TD>
      <TD WIDTH=20%><SELECT NAME = "Sport">
      <OPTION VALUE = "Soccer">Soccer</OPTION>
      <OPTION VALUE = "Football">Flag Football</OPTION>
      <OPTION VALUE = "Basketball">Basketball</OPTION>
      <OPTION VALUE = "Softball">Softball</OPTION>
      </SELECT></TD>
    <TD WIDTH=25%><INPUT TYPE="SUBMIT" VALUE="E-Mail Team"></TD>
  </TR>
</TABLE> 
</Center><BR>
<center><textarea name="message" rows="10" cols="65" wrap="virtual"> </textarea></center>
</FORM>    
 
</BODY>
</HTML>
