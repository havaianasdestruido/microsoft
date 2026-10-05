<% 

Set emp = Session("DBConn").Execute ("SELECT PKId FROM tbEmployee WHERE UserName ='" & request.form("login") & "'")
if emp.EOF then
	response.redirect "intro.asp"
	REM no record for this UserName - that has to be added by the database administrator
end if
%>


<% if ( Session("frames") = 0  or Session("ActiveXControls") = 0) then
	response.redirect "forms.asp?login=" & request.form("login")
end if %>
<html>

<!-- Frameset for the main modelling pages in the Benefits Distribution site. 
     User authentication would happen here, since the next page
	 must get private information from the HR database. 
	 Copyright 1996 by Microsoft Corporation.  -->	
	<% if ( Session("ActiveXControls") = 0 ) then
	response.write "<!-- axc ok -->"
	end if
	
	
	%>

<head><title>Benefits Selection Site</title></head>


	<frameset frameborder=0 cols="60%, 220"> <!-- the forms can get narrower more easily than the applet -->

		<FRAME SRC="forms.asp?login=<%= request.form("login")%>" name="forms" >
		<FRAME SRC="object.asp" name="display">
	</frameset> 

</html>

