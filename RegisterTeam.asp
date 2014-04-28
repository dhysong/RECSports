<HTML>
<HEAD>
<TITLE>RECsports - Registration form</TITLE>
<STYLE>
<!--
		body {background: #FFFFFF; font-size: 20pt; color: #000000; font-family: verdana, arial}

		td {font-size: 9pt; color: #000000; font-family: verdana, arial}

		A:link {text-decoration: none; color: #FF0000;}
		A:visited {text-decoration: none; color: #C0C0C0;}
		A:active {text-decoration: none; color: #000000;}
		A:hover {text-decoration: none; color:#336699;}
-->
</STYLE>
<SCRIPT LANGUAGE="JavaScript">
        var fieldnames = new Array (5)
	fieldnames[0] = "email"
	fieldnames[1] = "teamname"
	fieldnames[2] = "captain"
	fieldnames[3] = "pwd"
	fieldnames[4] = "keyword"
		
function validation(form) {

	var onoff=0
	var alertboxnames = ""

	var fields = new Array (5)
	fields[0] = form.email.value.length
	fields[1] = form.teamname.value.length
	fields[2] = form.captain.value.length
	fields[3] = form.pwd.value.length
	fields[4] = form.keyword.value.length
	
//Loop tests each entry (via the fields array) to see if it is empty.
//If it is, add 1 to onoff, and add the name of the field to the 
//alertbox variable.  Loop maximum is the length [property]
//of the fields array.

	for (var i=0; i < fields.length; i++) {
	
        if (fields[i] == 0) {
    	
    	alertboxnames = alertboxnames + fieldnames[i] + ", ";
        
        onoff ++;
		}              
	}

//If everything's ok, onoff would have made it through the for loop
//without anything added.  Return true, submit form.  If not, return
//false and show the correct alert box.

	if (onoff == 0) {
	
		return true
		
	} else {
	
		if (onoff == 1) {

//This first part checks if there is only one missing field.  If there
//is, adjust the grammar of the dialog box to singular form.

//frosty takes only the field name missing,
//not the comma.  Start at the beginning, 0, and go to the 
//position of the comma, since there is only one.  You can also use 
//the below method (alertboxnames.length-2 = two positions minus the
//last)

			var frosty = alertboxnames.substring (0, alertboxnames.indexOf(","));
		
			alert ("The following fieldname is required: \r" + frosty + "\r Please fill in this field before continuing.");

			} else {

//This second part just lists the fieldnames, but the frosty variable 
//parses out the last comma from the list.  Clever, eh?
//Starts at the beginning, 0, and goes to the end minus two.  One because
//the length property reports all the letters while .substring starts at
//zero.  The other one because there is a space added to all the commas.

	        var frosty = alertboxnames.substring (0, alertboxnames.length-2)
				
	        alert ("The following fieldnames are required: \r" + frosty + "\rPlease fill in these fields before continuing."); 

			}
		
		return false
		
	}

}

</SCRIPT>

</HEAD>

<BODY>

<CENTER><H1>Team Registration</H1></CENTER>
<FORM ACTION="AddTeam.asp" METHOD=POST onSubmit="return validation(this)">
                    <center> <TABLE BORDER=0>
                             <TR>
                                     <TD ALIGN="right"><font size="4">Sport:</TD>
                                     <TD><SELECT NAME = "keyword">
                                     <OPTION VALUE = "Soccer">Soccer</OPTION>
                                     <OPTION VALUE = "Football">Flag Football</OPTION>
                                     <OPTION VALUE = "Basketball">Basketball</OPTION>
                                     <OPTION VALUE = "Softball">Softball</OPTION>
                                     </SELECT></TD>
                             </TR>
                    </TABLE></center>
                    <table border="0" width="100%">
                             <tr>
                             <td width="100%" align="center"></td>
                             </tr>
                             <tr>
                                     <td width="100%" align="center"><font size="4">
                                     <b>Team Name </b>(Please keep to 20 characters or less)<b>:</b></font>
                                     &nbsp;<input type="text" name="teamname" size="20" tabindex="1" maxlength="20"></td>
                             </tr>
                             <tr>
                                     <td width="100%" align="center"><font size="4">
                                     <b>Password:</b>(Please keep to 10 characters or less)<b>:</b></font>
                                     &nbsp;<input type="text" name="pwd" size="20" tabindex="1" maxlength="20"></td>
                             </tr>
                             
                    </TABLE></center>
  <p style="margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>
  <p style="margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>
  <p style="margin-top: 0; margin-bottom: 0" align="center"><b><font size="4">Captain's
  Information</font></b></p>
  <div align="center">
    <center>
    <table border="1" width="80%">
      <tr>
        <td width="50%" align="center">
          <p style="margin-top: 0; margin-bottom: 0">Name:</p>
          <p style="margin-top: 0; margin-bottom: 0" align="center"><input type="text" name="captain" size="30" tabindex="3"></td>
        <td width="50%" align="center">
          <p style="margin-top: 0; margin-bottom: 0">Student ID Number (without
          dashes):</p>
          <p style="margin-top: 0; margin-bottom: 0" align="center"><input type="text" name="captssn" size="20" tabindex="4" maxlength="9"></td>
      </tr>
      <tr>
        <td width="50%" align="center">
          <p style="margin-top: 0; margin-bottom: 0">Address:</p>
          <p style="margin-top: 0; margin-bottom: 0" align="center"><input type="text" name="capaddr" size="40" tabindex="5"></td>
        <td width="50%" align="center">
          <p style="margin-top: 0; margin-bottom: 0">Phone Number:</p>
          <p style="margin-top: 0; margin-bottom: 0" align="center"><input type="text" name="capphone" size="20" tabindex="6"></td>
      </tr>
      <tr>
        <td width="50%" align="center">
          <p style="margin-top: 0; margin-bottom: 0">City/State/Zip Code:</p>
          <p style="margin-top: 0; margin-bottom: 0" align="center"><input type="text" name="capstate" size="40" tabindex="7"></td>
        <td width="50%" align="center">
          <p style="margin-top: 0; margin-bottom: 0">E-mail:</p>
          <p style="margin-top: 0; margin-bottom: 0" align="center"><input type="text" name="email" size="20" tabindex="8"></td>
      </tr>
    </table>
    </center>
  </div>
<p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
  <p style="margin-top: 0; margin-bottom: 0">&nbsp;</p>
  <div align="center">
    <center>
  <table border="1" width="80%">
    <tr>
      <td width="50%">

  <p align="left"><b>Gender (Check One)</b>:</p>
  <p align="left"><input type="radio" value="Mens" name="GenderType" style="background-color: rgb(192,192,192); color: rgb(192,192,192)" checked tabindex="15">Men's
  </p>
<p align="left"><input type="radio" value="Womens" name="GenderType" style="background-color: rgb(192,192,192); color: rgb(192,192,192)" tabindex="16">Women's
<p align="left"><input type="radio" value="CoRec" name="GenderType" style="background-color: rgb(192,192,192); color: rgb(192,192,192)" tabindex="17">Co-Rec</td>
      <td width="50%">
<p align="left"><b>League (Check One):</b>
<p align="left"><input type="radio" value="A" name="League" style="background-color: rgb(192,192,192); color: rgb(192,192,192)" checked tabindex="18">A
(Competitive)
<p align="left"><input type="radio" value="B" name="League" style="background-color: rgb(192,192,192); color: rgb(192,192,192)" tabindex="19">B
(Recreational)
        <p>&nbsp;</td>
    </tr>
  </table>
    </center>
  </div>
  <p style="margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>
  <p style="margin-top: 0; margin-bottom: 0" align="center">&nbsp;</p>

  <table border="0" width="100%">
    <tr>
      <td width="100%">
        <p align="center"><input TYPE="submit" VALUE="Submit" name="Submit" tabindex="85"></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <input TYPE="RESET" VALUE="Reset Form" tabindex="86"></td>
    </tr>
    <tr>
      <td width="100%">
        <p align="center"></td>
    </tr>
  </table>
                     </FORM>
<HR>
</BODY>

</HTML>
   
  