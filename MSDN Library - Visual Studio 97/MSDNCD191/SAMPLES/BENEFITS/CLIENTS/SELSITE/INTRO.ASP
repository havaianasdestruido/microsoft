<!-- Introductory page for the benefits sample site. 
     Copyright 1996 by Microsoft Corporation. Written in Developer Studio. -->

<html>
<head><title>Using your Benefits</title>
<link rel=stylesheet href="hr.css" type="text/css">
</head>


<body bgcolor=FFFFFF>    
<!-- the username is being given for simple ID; 
stronger verification should happen in real life - 
a password field, or better yet client/server authentication
(if the client supports it) -->

<CENTER><IMG SRC="images\ben_mast.jpg" WIDTH=600 HEIGHT=200></CENTER>
<BR>
<P><P>
<CENTER>
<TABLEWIDTH=700>
<TR><TD WIDTH=480>
You can model your distribution of pre-tax and post-tax benefits before requesting that the changes be made official.
</TD></TR>
</TABLE>
</CENTER>
<BR>
<P><P>
<BR>
<form name="ID" action="frames.asp" method=POST><center>
<table cellpadding=10 border=1 frame=box bordercolor="0000FF" rules=none>
<TR>
    <TD ALIGN=center><B>Username:</B>
    <input type=text name="login" size = 32><BR><P>

	<img src="images\ico_ani1.gif" height=34 width=32 hspace=10>
	<input type=submit value="OK"></TD>
</TR>
</table>

<BR>
</center>
</form>
<BR>
<P><P>
<CENTER>
<TABLEWIDTH=700>
<TR><TD WIDTH=480>
(If you keep getting bounced back to this page, 
there isn't a record with your username in the database; 
check for typos, or ask the database administrator to 
check for your record.)
</TD></TR>
</TABLE>
</CENTER>
</body>
</html>