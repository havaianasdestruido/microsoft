<!-- Copyright 1996 by Microsoft Corporation. Written in Visual Studio by CPHLewis. -->
<html>
<head><title>The effect on your paycheck</title>
</head>

<body bgcolor="004488" text="FFFFFF" >
<% if (Session("browser") = "IE") then %>
<STYLE> 
  P  {font: 12pt "Arial";
      color: white}
</STYLE>
<% end if%>
<p>The change to your paycheck,
before...</p>

<form name=tF>
<OBJECT 
	CLASSID="clsid:5220cb21-c88d-11cf-b347-00aa00a28331">
	<PARAM NAME="LPKPath" VALUE="BChart3d.lpk">
</OBJECT>
<OBJECT 
	CLASSID="CLSID:36EBB295-7567-11D0-8317-00AA006C3AB9"
	ID=foursquare Width=200 Height=300
	Codebase="objects/bchart3d.cab#version=1,0,0,0">
	<PARAM NAME=MAXBEN VALUE=426.55>
	<PARAM NAME=PRELABEL VALUE="Pre-tax:">
	<PARAM NAME=POSTLABEL VALUE="Post-tax:">
</OBJECT>


</form>


<p><center>...and after taxes.</center></p>
<p>The graph is relative to the benefit allotment.</p>

</body></html>