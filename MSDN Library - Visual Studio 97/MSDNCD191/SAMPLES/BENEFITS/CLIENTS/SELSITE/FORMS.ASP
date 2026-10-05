<%
If Request.querystring("login") = "" then
	response.redirect "intro.asp"
else
	PASS = "SELECT FirstName, LastName, Salary, PKId FROM tbEmployee WHERE UserName = '" & Request.querystring("login") & "'"
	REM We are assuming that HR fills the record when the employee is hired. 
end if

REM All the recordset objects we're going to use;
Set benefits = Server.CreateObject("ADODB.Recordset")
Set description = Server.CreateObject("ADODB.Recordset")
Set sel = Server.CreateObject("ADODB.Recordset")
Set coverage = Server.CreateObject("ADODB.Recordset")
Set thisperk = Server.CreateObject("ADODB.Recordset")
Set emp = Server.CreateObject("ADODB.Recordset")

REM the information we need about the employee - 
emp.Open PASS, Session("DBConn"), 0, 1, 1
REM We use the employee ID in the next page too, so we keep it as a session variable - 
Session("empid") = emp("PKId")

benefits.Open "Select Minimum, Maximum, Multiples from tbBenefit where BnfTypeId = 2", Session("DBConn"), 0, 1, 1
REM This serverside script is analogous to a middle tier in client/server; we have to know what
REM each records' contents mean, not just reformat the data. For instance, there are two possible
REM maximums; the dollar value in Maximum or a proportion of each employees' salary, kept in
REM Multiples as a percentage.
LIFEMAX = CDbl(emp("Salary"))  * benefits("Multiples") / 100
If ( LIFEMAX > benefits("Maximum") ) then 
	LIFEMAX = benefits("Maximum")
end if
Session("LifeMin") = benefits("Minimum")
Session("LifeMax") = LIFEMAX

benefits.Close
benefits.Open "Select Minimum, Maximum, Increment, Multiples, Cost from tbBenefit where BnfTypeId = 3", Session("DBConn"), 0, 1, 1
REM Another piece of business logic; the cost of disability insurance is proportional to the
REM employees' salary, the coverage chosen by the employee, and the constant price stored in Cost.
DISBASE =  CDbl(emp("Salary"))  / ( benefits("Cost") * 100) 
Session("DisMin") = benefits("Minimum")
Session("DisMax") = benefits("Maximum")

benefits.Close
benefits.Open "Select Maximum, Multiples from tbBenefit where BnfTypeId = 4", Session("DBConn"), 0, 1, 1
REM Pension deductions are calculated with two possible upper limits: an absolute annual dollar value,
REM or a percentage of the employee's salary. The legal limit is the smaller of these two. 
REM The percentage is specified by the user, and saved to the database, but we need to 
REM calculate its current effect on the paycheck. 
REM I'm assuming 24 paychecks a year. I am also ignoring the chance of bonuses putting annual
REM deductions over the dollar maximum.
REM Is the annual dollar limit less than the percentage-limit of the user's salary?
REM If so, work out what the dollar limit is as a percentage.

If ( benefits("Maximum") < (CSng(emp("Salary")) * benefits("Multiples") / 100 ) ) then
	Session("RetMax") = CInt( benefits("Maximum") * 100 / emp("Salary") )
else
	Session("RetMax") = benefits("Multiples")
end if

ANNUALRETMAX = benefits("Maximum")

sel.Open "SELECT tbEmplBnf.BenefitId, tbEmplBnf.CoverageId, tbBenefit.Cost FROM tbBenefit INNER JOIN tbEmplBnf ON tbBenefit.PKId = tbEmplBnf.BenefitId WHERE tbBenefit.BnfTypeId=1 AND tbEmplBnf.EmployeeId=" & Session("empid"), Session("DBConn"), 0, 1, 1
REM Get any recorded medical-benefit choices
If (Not sel.EOF) then
	EMPHEL = sel("BenefitId")
	EMPHELCOV = sel("CoverageId") 
	EMPHELCOST = sel("Cost") 
else
	EMPHEL = -1
	EMPHELCOV = 1
	EMPHELCOST = 0
end if 


REM we might need to change our output depending on the browser's capability; 

If ( Session("ActiveXControls") = -1 ) then
	CALLOBJ = "parent.display.document.tF.foursquare.recalc"
else
	if ( Session("javaapplets") = -1 ) then
		REM in Netscape, the object model is slightly different
		CALLOBJ = "document.applets[0].recalc"
	else
		REM putting in a dummy function was shorter than 
		REM adding server-side if..else..end if's to
		REM each function that used the statement.
		CALLOBJ = "donothing"
	end if
end if

%>
<!-- Copyright by Microsoft Corporation. Written by CPHLewis in Developer Studio. -->
<html>
<head><title>Options for Benefit Selection</title>
<link rel=stylesheet href="hr.css" type="text/css">

<% if (Session("javascript") = -1 ) then %>
<script language="JavaScript">

function HealthInsOnChange()
{
	var index = document.testForm.HealthIns.selectedIndex;
	var dpre = 0;
	if ( index == 0 )
		dpre = 0;
	<% 
	benefits.Close
	benefits.Open "SELECT Cost FROM tbBenefit WHERE BnfTypeId = 1", Session("DBConn"), 0, 1, 1
	INDEX = 0
	Do While Not benefits.EOF 
	INDEX = INDEX + 1%>
	if (index == <%= INDEX %>)
			dpre = <%= benefits("Cost")%>;
<% benefits.MoveNext
   Loop %>

	var delta = document.testForm.HealthDed.value - dpre;
	document.testForm.HealthDed.value = dpre;
	<%= CALLOBJ %>( delta, 0);
}

function LifeCovDispOnChange()
{
	var index = document.testForm.LifeCovDisp.value;
	if (index > <% response.write LIFEMAX %>)
	{
		index = <% response.write LIFEMAX %>;
		document.testForm.LifeCovDisp.value = <% response.write LIFEMAX %>;
	}

	if (index != 0 & index < <%= Session("LifeMin")%>)
	{
		index = <%= Session("LifeMin")%>;
		document.testForm.LifeCovDisp.value = <%= Session("LifeMin")%>;
	}

<%	benefits.Close
benefits.Open "SELECT Cost FROM tbBenefit WHERE BnfTypeId = 2", Session("DBConn"), 0, 1, 1 %>
	var dpre = TwoPlaces(index / <%= benefits("Cost")%>);
	document.testForm.LifeDedText.value = dpre;
	delta = document.testForm.LifeDed.value - dpre;
	document.testForm.LifeDed.value = dpre;
	document.testForm.LifeCovText.value = index;

	<%= CALLOBJ %>( delta, 0 );

}


function DisCovDispOnChange()
{
	var index = document.testForm.DisCovDisp.value;
	if (index > <%= Session("DisMax")%>)
	{
		index = <%= Session("DisMax") %>;
		document.testForm.DisCovDisp.value = <%= Session("DisMax")%>;
	}
	if (index != 0 & index < <%= Session("DisMin")%>)
	{
		index = <%= Session("DisMin") %>;
		document.testForm.DisCovDisp.value = <%= Session("DisMin")%>;
	}

	var dpre = TwoPlaces(<%= DISBASE %> * index);
	document.testForm.DisDedText.value = dpre;
	delta = document.testForm.DisDed.value - dpre;
	document.testForm.DisDed.value = dpre;
	if (document.testForm.DisTaxation.selectedIndex == 0) 
		<%= CALLOBJ %>( delta, 0);
	else
		<%= CALLOBJ %> ( 0, delta);
	document.testForm.DisCovText.value = index;
}


function DisTaxationonChange()
{
	var delta = document.testForm.DisDed.value;
	// subtract from the new, add to the old
	if (document.testForm.DisTaxation.selectedIndex == 1) // 1 is after taxes.
	{
		<%= CALLOBJ %>(parseFloat(delta), parseFloat(-delta));
	}
	else
	{
		<%= CALLOBJ %>(parseFloat(-delta), parseFloat(delta));
	}		
}

function RetDedTextonChange()
{
	// check for the allowable limit
	var index = document.testForm.RetDedText.value;
	if (index > <% response.write Session("RetMax") %>) 
	{
		index = <% response.write Session("RetMax") %>;
		document.testForm.RetDedText.value = <% response.write Session("RetMax") %>;
	}


	// calculate the deduction per paycheck
	index = document.testForm.RetDedText.value * <%= emp("Salary") %> / 2400
	var delta = document.testForm.RetDed.value - index;
	document.testForm.RetDed.value = index;
	<%= CALLOBJ %>( delta, 0 );

}

<%
REM We generated the list of perks dynamically, so we generate their handlers likewise.
benefits.Close
benefits.Open "SELECT PKId, Cost FROM tbBenefit WHERE BnfTypeId = 5", Session("DBConn"), 0, 1, 1
Do While Not benefits.EOF %>
function Perk<%= benefits("PKId") %>OnClick()
{
	if (document.testForm.Perk<%= benefits("PKId") %>.checked )
		<%= CALLOBJ %>( 0, -<%= benefits("Cost")%>);
	else
		<%= CALLOBJ %>( 0, <%= benefits("Cost")%>);
}
<% benefits.MoveNext
   Loop
%>


function TwoPlaces(numFloat)
{
	return (Math.floor(numFloat* 100))/100;
}

function donothing(flLeft, flRight)
{
}

function windowonLoad()
{	<% if (CALLOBJ <> "donothing") then 
		if ( Session("ActiveXControls") = -1) then%>
			parent.display.document.tF.foursquare.reset();
		<%else%>
			parent.document.foursquare.reset();
		<%end if%>
	<% End if %>
	// alert("Subtracting " + parseFloat(<%= EMPHELCOST%>));
	<%= CALLOBJ %>( -parseFloat(<%= EMPHELCOST%>), 0 );

	DisCovDispOnChange();
	
	LifeCovDispOnChange();

	RetDedTextonChange();

<%
REM Perk handling is generated dynamically again
sel.Close
sel.Open "SELECT BenefitId FROM tbEmplBnf INNER JOIN tbBenefit ON tbEmplBnf.BenefitId = tbBenefit.PKId WHERE tbBenefit.BnfTypeId=5 AND tbEmplBnf.EmployeeId=" & Session("empid"), Session("DBConn"), 0, 1, 1 
Do While Not sel.EOF %>
	document.testForm.Perk<%= sel("BenefitId")%>.checked = true;
	Perk<%= sel("BenefitId")%>OnClick();
<% sel.MoveNext
	Loop %>
}

</script>

<% end if
REM end of the "if (Session("javascript") = -1 )" clause. 

%>


</head>

<body bgcolor="FFFFFF"  LINK=435BFF VLINK=435BFF ALINK=435BFF TOPMARGIN=0 LEFTMARGIN=0 
<% if ( Session("javascript") = -1) then%>
onLoad="windowonLoad()"
<% end if%>>

<TABLE CELLPADDING=0 CELLSPACING=0>
<TR><TD><IMG SRC="images\header.gif" HEIGHT=75 WIDTH=800></TD></TR>
<TR><TD>
<NOBR><IMG SRC="images\blank.GIF" HEIGHT=10 WIDTH=105>
<SPAN class=jump>
<%REM collecting a handful of names from the database, because we're going to use each of them 
REM twice: this sample is small for it to save much work when the names changed, but you get the idea.
description.Open "SELECT Descr FROM tbBnfType ORDER BY PKId", Session("DBConn"), 0, 1, 1 
%>
<A href="#Medical"><%= description("Descr") %>
<% description.MoveNext%></A> * 
<A href="#Life"><%=description("Descr") %>
<% description.MoveNext%></a> * 
<A href="#Disability"><%= description("Descr") %>
<% description.MoveNext%></A> * 
<A href="#Retirement"><%=description("Descr") %>
<% description.MoveNext%></a> * 
<A href="#Perks"><%= description("Descr") %>
<% description.MoveFirst%></A> 
</SPAN>
</NOBR>
</TD></TR>
</TABLE>
<BR>
<P>

<form name=testForm action="response.asp" method=post target="_top">

<TABLE WIDTH=500 CELLPADDING=12>

<TR><TD><h1> <%= emp("FirstName") & " " & emp("LastName") %> </h1>
<SPAN class=gif><IMG SRC="images\rule.gif" HEIGHT=2 WIDTH=500></SPAN></td></tr>
<tr><td>

Each benefit selection may affect your paycheck either pre- or post-taxes; 
make choices below. When you're satisfied, 
click the "File Changes" button.
</td></tr>


<TR><TD VALIGN="top"><H2><a name="Medical"><%= description("Descr") %></H2>
<SPAN class=gif><IMG SRC="images\rule.gif" HEIGHT=2 WIDTH=500></SPAN>
<P>

Choose a <%= description("Descr") %> plan: 

<SELECT NAME="HealthIns" SIZE="1"  
<% if (Session("javascript") <> 0 ) then %>
onChange=HealthInsOnChange()
<% end if%>>
	<OPTION VALUE="-1">Decline; $0

<%  
benefits.Close
benefits.Open "SELECT * FROM tbBenefit WHERE BnfTypeId = 1", Session("DBConn"), 0, 1, 1
Do While Not benefits.EOF %>
	<OPTION VALUE="<%= benefits("PKId") %>"
<% If (EMPHEL = benefits("PKId")) then %>
	SELECTED
<% End If %>	
	><%= benefits("Description")%>; $<%= benefits("Cost")%>
<% benefits.MoveNext
Loop %>
</SELECT>

<% REM the first of many hidden variables. To store values for
   REM server calculation, Session variables; for client calculation, 
   REM hidden values ( or you could use JScript variables. Matter of taste.) %>

<input type=hidden name=HealthDed value=<%= EMPHELCOST%>>
<br></p>
<p>Who will be covered by this plan?

<SELECT NAME="HealthCov" SIZE="1">
<% coverage.Open "SELECT PKId, Descr FROM tbEmplBnfCoverage", Session("DBConn"), 0, 1, 1
Do While Not coverage.EOF %>
	<OPTION VALUE="<%=coverage("PKId") %>"
	<%If EMPHELCOV = coverage("PKId")  then 
		response.write "SELECTED"
	end if %>> <% response.write coverage("Descr")
	coverage.MoveNext
Loop%>
</SELECT>
<BR>
</TD></TR>

<% description.MoveNext%>
<TR><TD VALIGN="top"><H2><a name="Life"><%= description("Descr")%>
</H2><SPAN class=gif><IMG SRC="images\rule.gif" HEIGHT=2 WIDTH=500></SPAN><P>
<% sel.Close
sel.Open  "SELECT Amount FROM tbBenefit INNER JOIN tbEmplBnf ON tbBenefit.PKId = tbEmplBnf.BenefitId WHERE (tbEmplBnf.EmployeeId =" & Session("empid") & ") AND (tbBenefit.BnfTypeId =2)", Session("DBConn"), 0, 1, 1
if not sel.EOF then
	LIFECOV = sel("Amount")
else
	LIFECOV = 0
end if 
%>
Your minimum coverage is $<%= Session("LifeMin")%>, your maximum $<%= LIFEMAX%>, and the deduction per paycheck
is proportional to your coverage. </p>

<input type=hidden value=0 name="LifeDed"> 
<br><p>Coverage: 
<% if ( Session("ActiveXControls") = -1 ) then 
REM use the custom control%>

<OBJECT ID="LifeCovDisp" WIDTH=100 HEIGHT=20
 CLASSID="CLSID:073F3256-30EB-11D0-B84A-00AA00C0803B"
 CODEBASE="objects/spinmfc.cab#version=1,0,0,0">
    <PARAM NAME="_Version" VALUE="65536">
    <PARAM NAME="_ExtentX" VALUE="2646">
    <PARAM NAME="_ExtentY" VALUE="529">
    <PARAM NAME="_StockProps" VALUE="0">
    <PARAM NAME="Increment" VALUE="1000">
	<PARAM NAME="Minimum" VALUE="<%= Session("LifeMin")%>">
	<PARAM NAME="Maximum" VALUE="<%= Session("LifeMax")%>">
	<PARAM NAME="Value" VALUE="<%= LIFECOV%>">
	<PARAM NAME="Currency" VALUE="1">
</OBJECT>
<% else %>
<input type=text size=8 value=<%= LIFECOV%> name="LifeCovDisp" 
	<% if (Session("javascript") <> 0 ) then %>
	onChange=LifeCovDispOnChange()
	<% end if %>>
<% end if%>
<input type=hidden value=<%= LIFECOV%> name="LifeCovText">

<% if (Session("javascript") = -1 ) then %>
Deduction: $<input type=text size=4 value="0.0" name="LifeDedText" >
<% end if%></TD></TR>

<% description.MoveNext %>
<a name="Disability"><TR><TD VALIGN="top"><H2><%= description("Descr")%>
</H2><SPAN class=gif><IMG SRC="images\rule.gif" HEIGHT=2 WIDTH=500></SPAN><P>


Your coverage can range from <%= Session("DisMin")%>% to <%= Session("DisMax")%>% of your salary at the time of disability. 
The deduction per paycheck is proportional to the percentage and to your current salary.</p>

<% sel.Close
sel.Open "SELECT tbEmplBnf.Amount, tbEmplBnf.TaxId FROM tbBenefit INNER JOIN tbEmplBnf ON tbBenefit.PKId = tbEmplBnf.BenefitID WHERE (tbEmplBnf.EmployeeId =" & Session("empid") & ") AND (tbBenefit.BnfTypeId = 3)", Session("DBConn"), 0, 1, 1 
if not sel.EOF then
	DISAMOUNT = sel("Amount")
	DISTAX = sel("TaxId")
else
	DISAMOUNT = 0
	DISTAX = 2
end if%>

<input type=hidden value=0 name="DisDed"><p>
Amount:
<% if ( Session("ActiveXControls") = -1 ) then %>
<OBJECT ID="DisCovDisp" WIDTH=60 HEIGHT=20
 CLASSID="CLSID:073F3256-30EB-11D0-B84A-00AA00C0803B"
 CODEBASE="objects/spinmfc.cab#version=1,0,0,0">
    <PARAM NAME="_Version" VALUE="65536">
    <PARAM NAME="_ExtentX" VALUE="1482">
    <PARAM NAME="_ExtentY" VALUE="423">
    <PARAM NAME="_StockProps" VALUE="0">
    <PARAM NAME="Increment" VALUE="5">
	<PARAM NAME="Minimum" VALUE="<%= Session("DisMin")%>">
	<PARAM NAME="Maximum" VALUE="<%= Session("DisMax")%>">
	<PARAM NAME="Value" VALUE="<%= DISAMOUNT%>">
	<PARAM NAME="Currency" VALUE="0">
</OBJECT>
<% else %>
<input type=text size=8 value=<%= DISAMOUNT %> name="DisCovDisp" 
<% if (Session("javascript") <> 0 ) then %>
onChange=DisCovDispOnChange()
<%end if%>> 
<% end if %>% 
<% if (Session("javascript") = -1 ) then%>
Deduction: $<input type=text size=4 value="0.0" name="DisDedText">
<% end if %> <input type=hidden value=<%= DISAMOUNT %> name="DisCovText">
</p>
<p>This can be deducted 
<select name="DisTaxation" size=1 
<% if (Session("javascript") <> 0 ) then %>
onChange=DisTaxationonChange()
<%end if%>>
	<option value="2">before
	<option value="3" <% if DISTAX = 3 then%>
								SELECTED
						<%		end if %>
					 >after
</select> taxes.</TD></TR>

<% description.MoveNext%>
<TR><TD VALIGN="top"><H2><a name="Retirement"><%= description("Descr")%>
</H2><SPAN class=gif><IMG SRC="images\rule.gif" HEIGHT=2 WIDTH=500></SPAN><P>

 
<%
sel.Close
sel.Open "SELECT tbEmplBnf.Amount FROM tbBenefit INNER JOIN tbEmplBnf ON tbBenefit.PKId = tbEmplBnf.BenefitID WHERE tbEmplBnf.EmployeeId=" & Session("empid") & " AND tbBenefit.BnfTypeId= 4", Session("DBConn"), 0, 1, 1 
if not sel.EOF then
	RETDED = sel("Amount")
else
	RETDED = 0
end if %>


Pre-tax deductions into a savings fund; the maximum deduction is <%= Session("RetMax")%>% per paycheck. If your 
salary changes, the proportion of deduction to paycheck will remain constant until you reach the legal 
limit of $<%= ANNUALRETMAX %> a year.</p>
<p>Deduction per paycheck: 
<input type=text size=3 value=<%= RETDED%> name="RetDedText" 
<% if (Session("javascript") <> 0 ) then %>
onChange=RetDedTextonChange()
<%end if%>>%
<input type=hidden size=8 value=0 name="RetDed"></p>
</TD></TR>

<% description.MoveNext %>
<TR><TD VALIGN="top"><H2><a name="Perks"><%= description("Descr")%>
</H2><SPAN class=gif><IMG SRC="images\rule.gif" HEIGHT=2 WIDTH=500>
</SPAN><P>
Post-tax deductions for group-purchase benefits. </p><p>
<%	benefits.Close
benefits.Open "SELECT * FROM tbBenefit WHERE BnfTypeId = 5", Session("DBConn"), 0, 1, 1
	REM We don't know how many perks there are, and we don't know how many of them
	REM the employee has already chosen, so we query for all possible perks and for 
	REM any record of each perk having been chosen.

	Do While Not benefits.EOF 
		
		thisperk.Open "SELECT BenefitId FROM tbEmplBnf WHERE BenefitId=" & benefits("PKId") & " AND EmployeeId=" & Session("empid"), Session("DBConn"), 0, 1, 1 %>
		<input type=checkbox name="Perk<%= benefits("PKId") %>" value="<%= benefits("PKId") %>"
		<% If NOT thisperk.EOF  then %>
		checked 
		<%end if %> 
		<% if ( Session("javascript") <> 0 ) then %>
			onClick=Perk<%= benefits("PKId") %>OnClick()
		<% end if%>>
		<%= benefits("Description")%>; $<%= benefits("Cost")%><br>
<%		thisperk.Close 
		benefits.MoveNext
	Loop
%>
</SELECT> <input type=hidden value=0 name="PerksDed"></p>
</TD></TR>
<BR>
<TR><TD><IMG SRC="images\rule.gif" HEIGHT=2 WIDTH=500><P>
You may want to read a <a href="BenGloss.htm">glossary of terms</a>.<BR>
<p><input type=submit name=SubmitForm value="Request changes"><BR>


</TD></TR>
</TABLE>
<% if (Session("ActiveXControls") = 0 and Session("javaapplets") = -1) then
REM I was going to put the applet in the right-hand frame the chart object uses, but 
REM too often the form loaded before the applet was running, causing many
REM script errors. Since IE is the site's primary target, putting the applet next
REM to the Submit button seemed useful enough. %>

<table width=500 cellpadding=0 cellspacing=0><tr><td>
How this will affect your paycheck: <br>
You may be either credited or debited;
some benefits are paid for
before taxes, and some are paid after 
taxes.</td><td>

<applet 
        code=foursquare.class
		codebase="objects"
        name=foursquare

        width=175

        height=200>
    <param name=maxben value=426.55f>
        <param name=bgcolor value=8421631>
        <param name=fgcolor value=16777215>
        <param name=prelabel value = "Pre-tax:">
        <param name=postlabel value= "Post-tax:">
        <param name=midlabel value = "base pay">
        <param name=currency value= "$">
</applet>
</td></tr></table>
<% end if%>


<% if (Session("javascript") = -1  and Session("ActiveXControls") = -1) then %>

<script language="JavaScript" for="LifeCovDisp" event="onChange">
	LifeCovDispOnChange();
</script>
<script language="JavaScript" for="DisCovDisp" event="onChange">
	DisCovDispOnChange();
</script>

<% end if%>

</form>

</body></html>
